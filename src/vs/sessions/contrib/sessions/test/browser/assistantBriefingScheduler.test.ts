/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import assert from 'assert';
import { DeferredPromise } from '../../../../../base/common/async.js';
import { safeIntl } from '../../../../../base/common/date.js';
import { Event } from '../../../../../base/common/event.js';
import { generateUuid } from '../../../../../base/common/uuid.js';
import { upcastPartial } from '../../../../../base/test/common/mock.js';
import { ensureNoDisposablesAreLeakedInTestSuite } from '../../../../../base/test/common/utils.js';
import { NullLogService } from '../../../../../platform/log/common/log.js';
import { InMemoryStorageService } from '../../../../../platform/storage/common/storage.js';
import { IUserDataProfileService } from '../../../../../workbench/services/userDataProfile/common/userDataProfile.js';
import { IReachSubscriptionConnection } from '../../../accountMenu/browser/reachSubscriptionConnection.js';
import { AssistantBriefingOutcome, AssistantBriefingScheduler, parseAssistantBriefingTime } from '../../browser/assistantBriefingScheduler.js';
import { IAssistantRuns } from '../../browser/assistantRuns.js';

suite('AssistantBriefingScheduler', () => {
	const disposables = ensureNoDisposablesAreLeakedInTestSuite();

	test('persists schedules by profile and account and dispatches each occurrence once across schedulers', async () => {
		const storage = disposables.add(new InMemoryStorageService());
		const account = { id: generateUuid(), subscriptionActive: true };
		const reach = upcastPartial<IReachSubscriptionConnection>({
			assistantSupported: true,
			onDidChangeAccount: Event.None,
			getAccount: async () => account,
		});
		const calls: string[] = [];
		const runs = upcastPartial<IAssistantRuns>({
			getPending: () => undefined,
			sendScheduledBriefing: async (accountId, period) => { calls.push(`${accountId}:${period}`); return 'completed'; },
		});
		const profileId = generateUuid();
		const makeScheduler = (id: string) => disposables.add(new AssistantBriefingScheduler(
			reach, runs, upcastPartial<IUserDataProfileService>({ currentProfile: upcastPartial<IUserDataProfileService['currentProfile']>({ id }) }), storage, new NullLogService(),
		));
		const first = makeScheduler(profileId);
		const second = makeScheduler(profileId);
		await first.setSchedule(account.id, 'morning', { hour: 8, minute: 15 });
		const [schedule] = await second.getSchedules(account.id);
		assert.strictEqual(schedule.hour, 8);
		assert.strictEqual(schedule.minute, 15);
		assert.deepStrictEqual(await second.getSchedules('another-account'), []);
		assert.deepStrictEqual(await makeScheduler(generateUuid()).getSchedules(account.id), []);
		await first.checkDue(schedule.nextAt - 1);
		assert.deepStrictEqual(calls, []);
		await Promise.all([first.checkDue(schedule.nextAt), second.checkDue(schedule.nextAt)]);
		await second.checkDue(schedule.nextAt);
		assert.deepStrictEqual(calls, [`${account.id}:morning`]);
		const [next] = await first.getSchedules(account.id);
		assert.ok(next.nextAt > schedule.nextAt);
		assert.strictEqual(next.lastRun?.status, 'completed');
		assert.strictEqual(next.lastRun?.attempt, 1);
		await second.setSchedule(account.id, 'morning', undefined);
		await first.checkDue(next.nextAt);
		assert.deepStrictEqual(calls, [`${account.id}:morning`]);
		assert.deepStrictEqual(await first.getSchedules(account.id), []);
	});

	test('requires a valid 24-hour time', () => {
		assert.deepStrictEqual(parseAssistantBriefingTime(' 18:05 '), { hour: 18, minute: 5 });
		assert.strictEqual(parseAssistantBriefingTime('24:00'), undefined);
		assert.strictEqual(parseAssistantBriefingTime('09:60'), undefined);
	});

	test('persists and manages weekday routines with independent briefings and one dispatch per occurrence', async () => {
		const storage = disposables.add(new InMemoryStorageService());
		const account = { id: generateUuid(), subscriptionActive: true };
		const reach = upcastPartial<IReachSubscriptionConnection>({
			assistantSupported: true, onDidChangeAccount: Event.None, getAccount: async () => account,
		});
		const calls: string[] = [];
		const manualStarted = new DeferredPromise<string>();
		const completeManual = new DeferredPromise<'completed'>();
		const emailDelivery = { connectionId: 'work-mail', recipient: 'recipient@example.test' };
		const runs = upcastPartial<IAssistantRuns>({
			getPending: () => undefined,
			sendScheduledRoutine: async (accountId, routine, runId, canDeliverEmail) => {
				assert.strictEqual(accountId, account.id);
				assert.deepStrictEqual(routine.serverIds, ['work-mail']);
				calls.push(routine.prompt);
				if (calls.length === 1) {
					assert.deepStrictEqual(routine.emailDelivery, emailDelivery);
					assert.strictEqual(await canDeliverEmail?.(), true);
					await manualStarted.complete(runId);
					await completeManual.p;
					assert.strictEqual(await canDeliverEmail?.(), false);
					return 'completed';
				}
				assert.strictEqual(routine.emailDelivery, undefined);
				return 'needsAttention';
			},
		});
		const profile = upcastPartial<IUserDataProfileService>({
			currentProfile: upcastPartial<IUserDataProfileService['currentProfile']>({ id: generateUuid() }),
		});
		const first = disposables.add(new AssistantBriefingScheduler(reach, runs, profile, storage, new NullLogService()));
		const second = disposables.add(new AssistantBriefingScheduler(reach, runs, profile, storage, new NullLogService()));
		const settings = { name: 'Open requests', prompt: 'Find open requests', serverIds: ['work-mail'], cadence: 'weekdays' as const, hour: 8, minute: 30 };
		await first.setSchedule(account.id, 'morning', { hour: 8, minute: 15 });
		assert.strictEqual(await first.saveRoutine(account.id, { ...settings, emailDelivery }), true);
		assert.deepStrictEqual((await second.getSchedules(account.id)).map(item => item.period), ['morning']);
		const [routine] = await second.getRoutines(account.id);
		assert.deepStrictEqual(routine.emailDelivery, emailDelivery);
		assert.strictEqual(routine.lastRun, undefined);
		const formatter = safeIntl.DateTimeFormat('en', { weekday: 'short', timeZone: routine.timeZone }).value;
		const weekday = (epoch: number) => formatter.format(epoch);
		assert.ok(!['Sat', 'Sun'].includes(weekday(routine.nextAt)));
		assert.deepStrictEqual(await second.getRoutines('another-account'), []);
		assert.strictEqual(await second.runRoutineNow('another-account', routine.id), false);
		await first.setSchedule(account.id, 'morning', undefined);
		assert.strictEqual(await second.setRoutineEnabled(account.id, routine.id, false), true);
		await first.checkDue(routine.nextAt);
		assert.deepStrictEqual(calls, []);
		const manualRun = second.runRoutineNow(account.id, routine.id);
		const manualId = await manualStarted.p;
		const [running] = await first.getRoutines(account.id);
		assert.strictEqual(running.lastRun?.id, manualId);
		assert.strictEqual(running.lastRun?.trigger, 'manual');
		assert.strictEqual(running.lastRun?.status, 'running');
		assert.strictEqual(await first.runRoutineNow(account.id, routine.id), false);
		assert.strictEqual(await first.saveRoutine(account.id, { ...settings, prompt: 'Find requests due today' }, routine.id), true);
		assert.strictEqual((await second.getRoutines(account.id))[0].enabled, false);
		await completeManual.complete('completed');
		assert.strictEqual(await manualRun, true);
		const [finished] = await first.getRoutines(account.id);
		assert.strictEqual(finished.lastRun?.id, manualId);
		assert.strictEqual(finished.lastRun?.status, 'completed');
		assert.strictEqual(finished.prompt, 'Find requests due today');
		assert.strictEqual(await first.setRoutineEnabled(account.id, routine.id, true), true);
		const [resumed] = await second.getRoutines(account.id);
		// Claim on a Friday, then assert the next occurrence is Monday in the routine's own time zone.
		let friday = resumed.nextAt;
		while (weekday(friday) !== 'Fri') { friday += 24 * 60 * 60 * 1000; }
		await Promise.all([first.checkDue(friday), second.checkDue(friday)]);
		await second.checkDue(friday);
		assert.deepStrictEqual(calls, ['Find open requests', 'Find requests due today']);
		const [scheduled] = await first.getRoutines(account.id);
		assert.strictEqual(weekday(scheduled.nextAt), 'Mon');
		assert.strictEqual(scheduled.lastRun?.trigger, 'scheduled');
		assert.strictEqual(scheduled.lastRun?.status, 'needsAttention');
		assert.strictEqual(scheduled.lastRun?.startedAt, friday);
		assert.notStrictEqual(scheduled.lastRun?.id, manualId);
		assert.strictEqual(await second.deleteRoutine(account.id, routine.id), true);
		assert.strictEqual(await first.runRoutineNow(account.id, routine.id), false);
		assert.deepStrictEqual(await first.getRoutines(account.id), []);
	});

	test('preserves an interrupted routine attempt across scheduler restart', async () => {
		const storage = disposables.add(new InMemoryStorageService());
		const account = { id: generateUuid(), subscriptionActive: true };
		const reach = upcastPartial<IReachSubscriptionConnection>({
			assistantSupported: false, onDidChangeAccount: Event.None, getAccount: async () => account,
		});
		const started = new DeferredPromise<void>();
		const stopped = new DeferredPromise<'stopped'>();
		const runs = upcastPartial<IAssistantRuns>({
			getPending: () => undefined,
			sendScheduledRoutine: async () => {
				await started.complete();
				return stopped.p;
			},
		});
		const profile = upcastPartial<IUserDataProfileService>({
			currentProfile: upcastPartial<IUserDataProfileService['currentProfile']>({ id: generateUuid() }),
		});
		const first = disposables.add(new AssistantBriefingScheduler(reach, runs, profile, storage, new NullLogService()));
		await first.saveRoutine(account.id, { name: 'Review', prompt: 'Read requests', serverIds: ['work-mail'], cadence: 'daily', hour: 8, minute: 30 });
		const [routine] = await first.getRoutines(account.id);
		const dispatch = first.runRoutineNow(account.id, routine.id);
		await Promise.race([started.p, dispatch.then(() => assert.fail('Routine returned before dispatch'))]);
		const [running] = await first.getRoutines(account.id);
		first.dispose();
		await stopped.complete('stopped');
		await dispatch;
		const second = disposables.add(new AssistantBriefingScheduler(reach, runs, profile, storage, new NullLogService()));
		const [recovered] = await second.getRoutines(account.id);
		assert.deepStrictEqual(recovered.lastRun, { ...running.lastRun, status: 'interrupted' });
		assert.strictEqual(recovered.nextAt, routine.nextAt);
	});

	test('retries a transient briefing failure once and persists each attempt', async () => {
		const storage = disposables.add(new InMemoryStorageService());
		const account = { id: generateUuid(), subscriptionActive: true };
		const reach = upcastPartial<IReachSubscriptionConnection>({
			assistantSupported: false, onDidChangeAccount: Event.None, getAccount: async () => account,
		});
		const calls: string[] = [];
		let outcome: AssistantBriefingOutcome = 'retryable';
		const runs = upcastPartial<IAssistantRuns>({
			getPending: () => undefined,
			sendScheduledBriefing: async (accountId, period) => { calls.push(`${accountId}:${period}`); return outcome; },
		});
		const profile = upcastPartial<IUserDataProfileService>({
			currentProfile: upcastPartial<IUserDataProfileService['currentProfile']>({ id: generateUuid() }),
		});
		const first = disposables.add(new AssistantBriefingScheduler(reach, runs, profile, storage, new NullLogService()));
		const second = disposables.add(new AssistantBriefingScheduler(reach, runs, profile, storage, new NullLogService()));
		await first.setSchedule(account.id, 'morning', { hour: 8, minute: 15 });
		const [schedule] = await second.getSchedules(account.id);
		await first.checkDue(schedule.nextAt);
		assert.deepStrictEqual(calls, [`${account.id}:morning`]);
		const [attemptOne] = await second.getSchedules(account.id);
		assert.strictEqual(attemptOne.lastRun?.status, 'failed');
		assert.strictEqual(attemptOne.lastRun?.attempt, 1);
		assert.strictEqual(attemptOne.lastRun?.startedAt, schedule.nextAt);
		const retryAt = attemptOne.lastRun?.retryAt;
		assert.ok(retryAt !== undefined && retryAt > schedule.nextAt && retryAt < attemptOne.nextAt);
		assert.ok(attemptOne.nextAt > schedule.nextAt);
		await first.checkDue(retryAt - 1);
		assert.strictEqual(calls.length, 1);
		await Promise.all([first.checkDue(retryAt), second.checkDue(retryAt)]);
		assert.strictEqual(calls.length, 2);
		const [attemptTwo] = await first.getSchedules(account.id);
		assert.strictEqual(attemptTwo.lastRun?.status, 'failed');
		assert.strictEqual(attemptTwo.lastRun?.attempt, 2);
		assert.strictEqual(attemptTwo.lastRun?.retryAt, undefined);
		assert.strictEqual(attemptTwo.lastRun?.startedAt, retryAt);
		assert.strictEqual(attemptTwo.nextAt, attemptOne.nextAt);
		await first.checkDue(attemptTwo.nextAt - 1);
		assert.strictEqual(calls.length, 2);
	});

	test('does not replay an ordinary briefing failure', async () => {
		const storage = disposables.add(new InMemoryStorageService());
		const account = { id: generateUuid(), subscriptionActive: true };
		const reach = upcastPartial<IReachSubscriptionConnection>({
			assistantSupported: false, onDidChangeAccount: Event.None, getAccount: async () => account,
		});
		const calls: string[] = [];
		const runs = upcastPartial<IAssistantRuns>({
			getPending: () => undefined,
			sendScheduledBriefing: async (accountId, period) => { calls.push(`${accountId}:${period}`); return 'failed'; },
		});
		const profile = upcastPartial<IUserDataProfileService>({
			currentProfile: upcastPartial<IUserDataProfileService['currentProfile']>({ id: generateUuid() }),
		});
		const first = disposables.add(new AssistantBriefingScheduler(reach, runs, profile, storage, new NullLogService()));
		await first.setSchedule(account.id, 'morning', { hour: 8, minute: 15 });
		const [schedule] = await first.getSchedules(account.id);
		await first.checkDue(schedule.nextAt);
		assert.deepStrictEqual(calls, [`${account.id}:morning`]);
		const [finished] = await first.getSchedules(account.id);
		assert.strictEqual(finished.lastRun?.status, 'failed');
		assert.strictEqual(finished.lastRun?.attempt, 1);
		assert.strictEqual(finished.lastRun?.retryAt, undefined);
		await first.checkDue(schedule.nextAt + 61_000);
		assert.strictEqual(calls.length, 1);
	});

	test('does not write a briefing outcome over a schedule edited while it ran', async () => {
		const storage = disposables.add(new InMemoryStorageService());
		const account = { id: generateUuid(), subscriptionActive: true };
		const reach = upcastPartial<IReachSubscriptionConnection>({
			assistantSupported: false, onDidChangeAccount: Event.None, getAccount: async () => account,
		});
		const started = new DeferredPromise<void>();
		const gate = new DeferredPromise<'completed'>();
		const runs = upcastPartial<IAssistantRuns>({
			getPending: () => undefined,
			sendScheduledBriefing: async () => { await started.complete(); return gate.p; },
		});
		const profile = upcastPartial<IUserDataProfileService>({
			currentProfile: upcastPartial<IUserDataProfileService['currentProfile']>({ id: generateUuid() }),
		});
		const first = disposables.add(new AssistantBriefingScheduler(reach, runs, profile, storage, new NullLogService()));
		await first.setSchedule(account.id, 'morning', { hour: 8, minute: 15 });
		const [schedule] = await first.getSchedules(account.id);
		const dispatch = first.checkDue(schedule.nextAt);
		await started.p;
		await first.setSchedule(account.id, 'morning', { hour: 9, minute: 30 });
		await gate.complete('completed');
		await dispatch;
		const [edited] = await first.getSchedules(account.id);
		assert.strictEqual(edited.hour, 9);
		assert.strictEqual(edited.minute, 30);
		assert.strictEqual(edited.lastRun, undefined);
	});

	test('preserves an interrupted briefing attempt across scheduler restart', async () => {
		const storage = disposables.add(new InMemoryStorageService());
		const account = { id: generateUuid(), subscriptionActive: true };
		const reach = upcastPartial<IReachSubscriptionConnection>({
			assistantSupported: false, onDidChangeAccount: Event.None, getAccount: async () => account,
		});
		const started = new DeferredPromise<void>();
		const gate = new DeferredPromise<'completed'>();
		const runs = upcastPartial<IAssistantRuns>({
			getPending: () => undefined,
			sendScheduledBriefing: async () => { await started.complete(); return gate.p; },
		});
		const profile = upcastPartial<IUserDataProfileService>({
			currentProfile: upcastPartial<IUserDataProfileService['currentProfile']>({ id: generateUuid() }),
		});
		const first = disposables.add(new AssistantBriefingScheduler(reach, runs, profile, storage, new NullLogService()));
		await first.setSchedule(account.id, 'morning', { hour: 8, minute: 15 });
		const [schedule] = await first.getSchedules(account.id);
		const dispatch = first.checkDue(schedule.nextAt);
		await started.p;
		const [running] = await first.getSchedules(account.id);
		assert.strictEqual(running.lastRun?.status, 'running');
		first.dispose();
		await gate.complete('completed');
		await dispatch;
		const second = disposables.add(new AssistantBriefingScheduler(reach, runs, profile, storage, new NullLogService()));
		const [recovered] = await second.getSchedules(account.id);
		assert.strictEqual(recovered.lastRun?.status, 'interrupted');
		assert.strictEqual(recovered.lastRun?.attempt, 1);
		assert.strictEqual(recovered.nextAt, running.nextAt);
	});
});
