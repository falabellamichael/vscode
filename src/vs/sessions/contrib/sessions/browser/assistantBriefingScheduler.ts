/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { mainWindow } from '../../../../base/browser/window.js';
import { IndexedDB } from '../../../../base/browser/indexedDB.js';
import { RunOnceScheduler } from '../../../../base/common/async.js';
import { Disposable, toDisposable } from '../../../../base/common/lifecycle.js';
import { hasKey } from '../../../../base/common/types.js';
import { generateUuid } from '../../../../base/common/uuid.js';
import { nextAutomationCronOccurrence } from '../../../../platform/agentHost/common/automationCron.js';
import { InstantiationType, registerSingleton } from '../../../../platform/instantiation/common/extensions.js';
import { createDecorator } from '../../../../platform/instantiation/common/instantiation.js';
import { ILogService } from '../../../../platform/log/common/log.js';
import { IStorageService, StorageScope, StorageTarget } from '../../../../platform/storage/common/storage.js';
import { IUserDataProfileService } from '../../../../workbench/services/userDataProfile/common/userDataProfile.js';
import { isSignalCMailAddress } from '../../../../workbench/contrib/mcp/browser/signalCDirectMail.js';
import { IReachSubscriptionConnection } from '../../accountMenu/browser/reachSubscriptionConnection.js';
import { AssistantBriefingPeriod } from './assistantBriefings.js';
import { IAssistantRuns } from './assistantRuns.js';

const DATABASE = 'signal-c-assistant-briefings';
const STORE = 'schedules';
const CONFIGURED = 'sessions.assistant.briefingSchedulerConfigured';
const CHECK_INTERVAL_MS = 60_000;
const BRIEFING_RETRY_DELAY_MS = CHECK_INTERVAL_MS;

export interface IAssistantBriefingTime {
	readonly hour: number;
	readonly minute: number;
}

interface IAssistantBriefingSchedule extends IAssistantBriefingTime {
	readonly period: AssistantBriefingPeriod;
	readonly timeZone: string;
	readonly nextAt: number;
	readonly lastRun?: IAssistantBriefingRun;
}

export interface IAssistantRoutineSettings extends IAssistantBriefingTime {
	readonly name: string;
	readonly prompt: string;
	readonly serverIds: readonly string[];
	readonly cadence: 'daily' | 'weekdays';
	readonly emailDelivery?: IAssistantRoutineEmailDelivery;
}

export interface IAssistantRoutineEmailDelivery {
	readonly connectionId: string;
	readonly recipient: string;
}

export type AssistantRoutineDraft = Partial<Pick<IAssistantRoutineSettings, 'cadence' | 'hour' | 'minute'>> & Pick<IAssistantRoutineSettings, 'name' | 'prompt'>;

export type AssistantRoutineOutcome = 'completed' | 'needsAttention' | 'failed' | 'stopped';
export type AssistantBriefingOutcome = AssistantRoutineOutcome | 'retryable';

export interface IAssistantRoutineRun {
	readonly id: string;
	readonly startedAt: number;
	readonly trigger: 'scheduled' | 'manual';
	readonly status: AssistantRoutineOutcome | 'running' | 'interrupted';
}

interface IAssistantBriefingRun {
	readonly id: string;
	readonly startedAt: number;
	readonly attempt: 1 | 2;
	readonly status: AssistantRoutineOutcome | 'running' | 'interrupted';
	readonly retryAt?: number;
}

export interface IAssistantRoutine extends IAssistantRoutineSettings {
	readonly id: string;
	readonly enabled: boolean;
	readonly timeZone: string;
	readonly nextAt: number;
	readonly lastRun?: IAssistantRoutineRun;
}

type AssistantSchedule = IAssistantBriefingSchedule | IAssistantRoutine;

export function parseAssistantBriefingTime(value: string): IAssistantBriefingTime | undefined {
	const match = /^([01]\d|2[0-3]):([0-5]\d)$/.exec(value.trim());
	return match ? { hour: Number(match[1]), minute: Number(match[2]) } : undefined;
}

function nextOccurrence(schedule: IAssistantBriefingTime, timeZone: string, now: number, cadence: IAssistantRoutineSettings['cadence'] = 'daily'): number {
	const days = cadence === 'weekdays' ? '1-5' : '*';
	return nextAutomationCronOccurrence(`${schedule.minute} ${schedule.hour} * * ${days}`, timeZone, new Date(now)).getTime();
}

function isBriefing(schedule: AssistantSchedule): schedule is IAssistantBriefingSchedule {
	return hasKey(schedule, { period: true });
}

function isRoutine(schedule: AssistantSchedule): schedule is IAssistantRoutine {
	return !isBriefing(schedule);
}

function startRoutine(routine: IAssistantRoutine, trigger: IAssistantRoutineRun['trigger'], startedAt: number): IAssistantRoutine & { readonly lastRun: IAssistantRoutineRun } {
	return { ...routine, lastRun: { id: generateUuid(), trigger, startedAt, status: 'running' } };
}

function briefingDueAt(schedule: AssistantSchedule): number {
	return isBriefing(schedule) && schedule.lastRun?.retryAt !== undefined ? Math.min(schedule.nextAt, schedule.lastRun.retryAt) : schedule.nextAt;
}

function isRoutineEmailDelivery(value: unknown): value is IAssistantRoutineEmailDelivery {
	return !!value && typeof value === 'object'
		&& 'connectionId' in value && typeof value.connectionId === 'string' && !!value.connectionId
		&& 'recipient' in value && typeof value.recipient === 'string' && isSignalCMailAddress(value.recipient);
}

function decodeSchedules(raw: string | undefined): readonly AssistantSchedule[] {
	try {
		const schedules = raw && JSON.parse(raw);
		if (Array.isArray(schedules)) {
			return schedules.filter(s => s
				&& Number.isInteger(s.hour) && s.hour >= 0 && s.hour <= 23
				&& Number.isInteger(s.minute) && s.minute >= 0 && s.minute <= 59
				&& typeof s.timeZone === 'string' && Number.isFinite(s.nextAt)
				&& (s.period === 'morning' || s.period === 'evening' || (
					typeof s.id === 'string' && s.period === undefined && typeof s.name === 'string' && typeof s.prompt === 'string'
					&& typeof s.enabled === 'boolean' && (s.cadence === 'daily' || s.cadence === 'weekdays')
					&& (s.emailDelivery === undefined || isRoutineEmailDelivery(s.emailDelivery))
					&& Array.isArray(s.serverIds) && (s.serverIds.length > 0 || s.emailDelivery !== undefined)
					&& s.serverIds.every((id: unknown) => typeof id === 'string')
				))).map((schedule: AssistantSchedule) => {
					if (schedule.lastRun === undefined) { return schedule; }
					if (isBriefing(schedule)) {
						const run = schedule.lastRun;
						const valid = run && typeof run.id === 'string' && Number.isFinite(run.startedAt)
							&& (run.attempt === 1 || run.attempt === 2)
							&& ['completed', 'needsAttention', 'failed', 'stopped', 'running', 'interrupted'].includes(run.status)
							&& (run.retryAt === undefined || (Number.isFinite(run.retryAt) && run.status === 'failed' && run.attempt === 1));
						return valid ? schedule : { ...schedule, lastRun: undefined };
					}
					const run = schedule.lastRun;
					const valid = run && typeof run.id === 'string' && Number.isFinite(run.startedAt)
						&& (run.trigger === 'scheduled' || run.trigger === 'manual')
						&& ['completed', 'needsAttention', 'failed', 'stopped', 'running', 'interrupted'].includes(run.status);
					return valid ? schedule : { ...schedule, lastRun: undefined };
				});
		}
	} catch { /* Previous or invalid local data is ignored. */ }
	return [];
}

export const IAssistantBriefingScheduler = createDecorator<IAssistantBriefingScheduler>('assistantBriefingScheduler');
export interface IAssistantBriefingScheduler {
	readonly _serviceBrand: undefined;
	start(): void;
	getSchedules(accountId: string): Promise<readonly IAssistantBriefingSchedule[]>;
	setSchedule(accountId: string, period: AssistantBriefingPeriod, time: IAssistantBriefingTime | undefined): Promise<void>;
	getRoutines(accountId: string): Promise<readonly IAssistantRoutine[]>;
	saveRoutine(accountId: string, settings: IAssistantRoutineSettings, id?: string): Promise<boolean>;
	setRoutineEnabled(accountId: string, id: string, enabled: boolean): Promise<boolean>;
	deleteRoutine(accountId: string, id: string): Promise<boolean>;
	runRoutineNow(accountId: string, id: string): Promise<boolean>;
}

export class AssistantBriefingScheduler extends Disposable implements IAssistantBriefingScheduler {
	declare readonly _serviceBrand: undefined;
	private database: Promise<IndexedDB> | undefined;
	private readonly check = this._register(new RunOnceScheduler(() => void this.checkDue(), CHECK_INTERVAL_MS));

	constructor(
		@IReachSubscriptionConnection private readonly reach: IReachSubscriptionConnection,
		@IAssistantRuns private readonly runs: IAssistantRuns,
		@IUserDataProfileService private readonly profiles: IUserDataProfileService,
		@IStorageService private readonly storage: IStorageService,
		@ILogService private readonly log: ILogService,
	) {
		super();
		this._register(toDisposable(() => void this.database?.then(db => db.close(), () => undefined)));
		this._register(this.storage.onDidChangeValue(StorageScope.PROFILE, CONFIGURED, this._store)(() => this.start()));
		this._register(this.reach.onDidChangeAccount(() => this.start()));
	}

	start(): void {
		if (this.reach.assistantSupported && this.storage.getBoolean(CONFIGURED, StorageScope.PROFILE)) {
			this.check.schedule(0);
		}
	}

	async getSchedules(accountId: string): Promise<readonly IAssistantBriefingSchedule[]> {
		const key = this.key(accountId);
		const schedules = (await this.read(key)).filter(isBriefing);
		if (!schedules.some(schedule => schedule.lastRun?.status === 'running')) { return schedules; }
		return mainWindow.navigator.locks.request(`${DATABASE}:${key}`, { ifAvailable: true }, async lock => {
			if (!lock) { return schedules; }
			// A live briefing holds this lock until its final status has been saved.
			await this.update(key, schedules => {
				if (!schedules.some(schedule => isBriefing(schedule) && schedule.lastRun?.status === 'running')) { return undefined; }
				const recovered = schedules.map(schedule => isBriefing(schedule) && schedule.lastRun?.status === 'running'
					? { ...schedule, lastRun: { ...schedule.lastRun, status: 'interrupted' as const } } : schedule);
				return { schedules: recovered, result: undefined };
			});
			return (await this.read(key)).filter(isBriefing);
		});
	}

	async setSchedule(accountId: string, period: AssistantBriefingPeriod, time: IAssistantBriefingTime | undefined): Promise<void> {
		const timeZone = Intl.DateTimeFormat().resolvedOptions().timeZone;
		const schedule = time && { ...time, period, timeZone, nextAt: nextOccurrence(time, timeZone, Date.now()) };
		await this.update(this.key(accountId), schedules => {
			const otherSchedules = schedules.filter(item => !isBriefing(item) || item.period !== period);
			return { schedules: schedule ? [...otherSchedules, schedule] : otherSchedules, result: undefined };
		});
		if (schedule) {
			this.activate();
		}
	}

	async getRoutines(accountId: string): Promise<readonly IAssistantRoutine[]> {
		const key = this.key(accountId);
		const routines = (await this.read(key)).filter(isRoutine);
		if (!routines.some(routine => routine.lastRun?.status === 'running')) { return routines; }
		return mainWindow.navigator.locks.request(`${DATABASE}:${key}`, { ifAvailable: true }, async lock => {
			if (!lock) { return routines; }
			// A live routine holds this lock until its final status has been saved.
			await this.update(key, schedules => {
				if (!schedules.some(schedule => isRoutine(schedule) && schedule.lastRun?.status === 'running')) { return undefined; }
				const recovered = schedules.map(schedule => isRoutine(schedule) && schedule.lastRun?.status === 'running'
					? { ...schedule, lastRun: { ...schedule.lastRun, status: 'interrupted' as const } } : schedule);
				return { schedules: recovered, result: undefined };
			});
			return (await this.read(key)).filter(isRoutine);
		});
	}

	async saveRoutine(accountId: string, settings: IAssistantRoutineSettings, id?: string): Promise<boolean> {
		if (!settings.name.trim() || !settings.prompt.trim() || (!settings.serverIds.length && !settings.emailDelivery)
			|| (settings.emailDelivery !== undefined && !isRoutineEmailDelivery(settings.emailDelivery))) { return false; }
		const timeZone = Intl.DateTimeFormat().resolvedOptions().timeZone;
		const nextAt = nextOccurrence(settings, timeZone, Date.now(), settings.cadence);
		const routine = { ...settings, emailDelivery: settings.emailDelivery, timeZone, nextAt };
		const saved = id
			? await this.changeRoutine(this.key(accountId), id, previous => ({ ...previous, ...routine }))
			: await this.update(this.key(accountId), schedules => ({ schedules: [...schedules, { ...routine, id: generateUuid(), enabled: true }], result: true }));
		if (saved) { this.activate(); }
		return !!saved;
	}

	setRoutineEnabled(accountId: string, id: string, enabled: boolean): Promise<boolean> {
		return this.changeRoutine(this.key(accountId), id, routine => ({
			...routine, enabled,
			nextAt: enabled ? nextOccurrence(routine, routine.timeZone, Date.now(), routine.cadence) : routine.nextAt,
		}));
	}

	deleteRoutine(accountId: string, id: string): Promise<boolean> {
		return this.changeRoutine(this.key(accountId), id, () => undefined);
	}

	async runRoutineNow(accountId: string, id: string): Promise<boolean> {
		const key = this.key(accountId);
		return mainWindow.navigator.locks.request(`${DATABASE}:${key}`, { ifAvailable: true }, async lock => {
			if (!lock || this._store.isDisposed || this.runs.getPending(accountId)) { return false; }
			const account = await this.reach.getAccount();
			if (account?.id !== accountId || !account.subscriptionActive || this.runs.getPending(accountId)) { return false; }
			const routine = await this.update(key, schedules => {
				const previous = schedules.find((item): item is IAssistantRoutine => isRoutine(item) && item.id === id);
				if (!previous) { return undefined; }
				const started = startRoutine(previous, 'manual', Date.now());
				return { schedules: schedules.map(item => item === previous ? started : item), result: started };
			});
			if (!routine || this._store.isDisposed) { return false; }
			await this.finishRoutine(key, accountId, routine, routine.lastRun);
			return true;
		});
	}

	private async finishRoutine(key: string, accountId: string, routine: IAssistantRoutine, run: IAssistantRoutineRun): Promise<void> {
		let status: AssistantRoutineOutcome = 'failed';
		try {
			const delivery = routine.emailDelivery;
			status = await this.runs.sendScheduledRoutine(accountId, routine, run.id, delivery && (async () => {
				const current = (await this.read(key)).find(item => isRoutine(item) && item.id === routine.id);
				return !!current && isRoutine(current) && current.lastRun?.id === run.id
					&& (run.trigger === 'manual' || current.enabled)
					&& current.emailDelivery?.connectionId === delivery.connectionId && current.emailDelivery.recipient === delivery.recipient
					&& current.name === routine.name && current.prompt === routine.prompt
					&& current.serverIds.length === routine.serverIds.length && current.serverIds.every((id, index) => id === routine.serverIds[index]);
			}));
		} catch (error) {
			this.log.error('[AssistantBriefingScheduler] Routine failed before returning a result', error);
		}
		if (!this._store.isDisposed) {
			await this.changeRoutine(key, routine.id, current => current.lastRun?.id === run.id ? { ...current, lastRun: { ...run, status } } : current);
		}
	}

	private async finishBriefing(key: string, accountId: string, period: AssistantBriefingPeriod, run: IAssistantBriefingRun, now: number): Promise<void> {
		let outcome: AssistantBriefingOutcome = 'failed';
		try {
			outcome = await this.runs.sendScheduledBriefing(accountId, period);
		} catch (error) {
			this.log.error('[AssistantBriefingScheduler] Briefing failed before returning a result', error);
		}
		const retryAt = Math.max(now, Date.now()) + BRIEFING_RETRY_DELAY_MS;
		await this.update(key, schedules => {
			const current = schedules.find((item): item is IAssistantBriefingSchedule => isBriefing(item) && item.period === period);
			if (!current || current.lastRun?.id !== run.id) { return undefined; }
			const lastRun: IAssistantBriefingRun = { ...run, status: outcome === 'retryable' ? 'failed' : outcome,
				...(outcome === 'retryable' && run.attempt === 1 && retryAt < current.nextAt ? { retryAt } : {}) };
			return { schedules: schedules.map(item => item === current ? { ...current, lastRun } : item), result: undefined };
		});
	}

	private activate(): void {
		this.storage.store(CONFIGURED, true, StorageScope.PROFILE, StorageTarget.MACHINE);
		this.start();
	}

	private async changeRoutine(key: string, id: string, change: (routine: IAssistantRoutine) => IAssistantRoutine | undefined): Promise<boolean> {
		return !!(await this.update(key, schedules => {
			const index = schedules.findIndex(item => isRoutine(item) && item.id === id);
			const existing = schedules[index];
			if (!existing || !isRoutine(existing)) { return undefined; }
			const next = change(existing);
			if (next === existing) { return undefined; }
			const updated = [...schedules];
			updated.splice(index, 1, ...(next ? [next] : []));
			return { schedules: updated, result: true };
		}));
	}

	private key(accountId: string): string {
		return JSON.stringify([this.profiles.currentProfile.id, accountId]);
	}

	private getDatabase(): Promise<IndexedDB> {
		return this.database ??= IndexedDB.create(DATABASE, 1, [STORE]).catch(error => {
			this.database = undefined;
			throw error;
		});
	}

	private async read(key: string): Promise<readonly AssistantSchedule[]> {
		const db = await this.getDatabase();
		return decodeSchedules(await db.runInTransaction(STORE, 'readonly', store => store.get(key)));
	}

	private async update<T>(key: string, change: (schedules: readonly AssistantSchedule[]) => { schedules: readonly AssistantSchedule[]; result: T } | undefined): Promise<T | undefined> {
		const db = await this.getDatabase();
		let raw: unknown = await db.runInTransaction(STORE, 'readonly', store => store.get(key));
		const isString = (value: unknown): value is string => typeof value === 'string';
		while (!this._store.isDisposed) {
			const previous = isString(raw) ? raw : undefined;
			const next = change(decodeSchedules(previous));
			if (!next) { return undefined; }
			const outcome = await db.compareAndSwap(STORE, key, previous, JSON.stringify(next.schedules), isString);
			if (outcome.swapped) { return next.result; }
			raw = outcome.currentValue;
		}
		return undefined;
	}

	async checkDue(now = Date.now()): Promise<void> {
		try {
			const account = await this.reach.getAccount();
			if (!account?.subscriptionActive || this._store.isDisposed || this.runs.getPending(account.id)) { return; }
			const key = this.key(account.id);
			await mainWindow.navigator.locks.request(`${DATABASE}:${key}`, { ifAvailable: true }, async lock => {
				if (!lock || this._store.isDisposed || this.runs.getPending(account.id)) { return; }
				const due = await this.update(key, schedules => {
					const first = schedules.filter(item => isBriefing(item) || item.enabled).sort((a, b) => briefingDueAt(a) - briefingDueAt(b))[0];
					if (!first || briefingDueAt(first) > now) { return undefined; }
					let scheduled: AssistantSchedule;
					if (isBriefing(first)) {
						const retry = first.lastRun?.retryAt !== undefined && now < first.nextAt;
						scheduled = { ...first, nextAt: retry ? first.nextAt : nextOccurrence(first, first.timeZone, now),
							lastRun: { id: generateUuid(), startedAt: now, attempt: retry ? 2 as const : 1 as const, status: 'running' as const } };
					} else {
						scheduled = { ...startRoutine(first, 'scheduled', now), nextAt: nextOccurrence(first, first.timeZone, now, first.cadence) };
					}
					return { schedules: schedules.map(item => item === first ? scheduled : item), result: scheduled };
				});
				if (due && !this._store.isDisposed) {
					if (isBriefing(due)) {
						if (due.lastRun) { await this.finishBriefing(key, account.id, due.period, due.lastRun, now); }
					} else if (due.lastRun) {
						await this.finishRoutine(key, account.id, due, due.lastRun);
					}
				}
			});
		} catch (error) {
			this.log.error('[AssistantBriefingScheduler] Could not dispatch a scheduled Assistant run', error);
		} finally {
			if (!this._store.isDisposed && this.storage.getBoolean(CONFIGURED, StorageScope.PROFILE)) {
				this.check.schedule();
			}
		}
	}
}

export class AssistantBriefingScheduleContribution {
	static readonly ID = 'workbench.contrib.assistantBriefingSchedule';
	constructor(@IAssistantBriefingScheduler scheduler: IAssistantBriefingScheduler) {
		scheduler.start();
	}
}

registerSingleton(IAssistantBriefingScheduler, AssistantBriefingScheduler, InstantiationType.Delayed);
