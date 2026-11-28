/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Signal-C. All rights reserved.
 *--------------------------------------------------------------------------------------------*/

import assert from 'assert';
import { Event } from '../../../../../base/common/event.js';
import { observableValue } from '../../../../../base/common/observable.js';
import { ensureNoDisposablesAreLeakedInTestSuite } from '../../../../../base/test/common/utils.js';
import { mock, upcastPartial } from '../../../../../base/test/common/mock.js';
import { IConfigurationService } from '../../../../../platform/configuration/common/configuration.js';
import { ILogService } from '../../../../../platform/log/common/log.js';
import { InMemoryStorageService } from '../../../../../platform/storage/common/storage.js';
import { directMailGmailQuery, directMailGraphFilter, ISignalCDirectMail, ISignalCDirectMailConnection } from '../../../../../workbench/contrib/mcp/browser/signalCDirectMail.js';
import { IStatusbarService } from '../../../../../workbench/services/statusbar/browser/statusbar.js';
import { IReachSubscriptionConnection } from '../../../accountMenu/browser/reachSubscriptionConnection.js';
import { AssistantConversationStore } from '../../browser/assistantConversation.js';
import { AssistantMailWatcher } from '../../browser/assistantMailWatcher.js';

suite('AssistantMailWatcher', () => {
	const store = ensureNoDisposablesAreLeakedInTestSuite();

	const connected = upcastPartial<ISignalCDirectMailConnection>({ id: 'outlook-conn', provider: 'outlook', label: 'Outlook', needsSignIn: false });
	const signedOut = upcastPartial<ISignalCDirectMailConnection>({ id: 'gmail-conn', provider: 'gmail', label: 'Gmail', needsSignIn: true });

	function fixture(options: { enabled?: boolean; failReads?: boolean } = {}) {
		const storage = store.add(new InMemoryStorageService());
		const reads: { connection: string; args: Record<string, unknown> }[] = [];
		const directMail = new class extends mock<ISignalCDirectMail>() {
			override readonly connections = observableValue<readonly ISignalCDirectMailConnection[]>('fixtureMail', [connected, signedOut]);
			override async read(connection: ISignalCDirectMailConnection, args: Record<string, unknown>): Promise<string> {
				reads.push({ connection: connection.id, args });
				if (options.failReads) { throw new Error('mailbox offline'); }
				return JSON.stringify({ messages: [{ id: 'm1', subject: 'Urgent: review', from: 'boss@corp.com', received: '2026-11-19T12:00:00Z', url: 'https://outlook.office.com/mail/m1' }] });
			}
		};
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override readonly assistantSupported = true;
			override async getAccount() { return upcastPartial({ id: 'account', subscriptionActive: true }); }
		};
		let enabled = options.enabled ?? true;
		const configuration = upcastPartial<IConfigurationService>({
			getValue: () => enabled,
			updateValue: async (_key: string, value: unknown) => { enabled = value as boolean; },
			onDidChangeConfiguration: Event.None,
		});
		let badge: { text: string } | undefined;
		let badgeDisposed = false;
		const statusBar = upcastPartial<IStatusbarService>({
			addEntry: entry => {
				badge = { text: entry.text };
				return { update: updated => { badge = { text: updated.text }; }, dispose: () => { badgeDisposed = true; badge = undefined; } };
			},
		});
		const log = upcastPartial<ILogService>({ debug: () => undefined });
		const watcher = store.add(new AssistantMailWatcher(reach, directMail, configuration, storage, statusBar, log));
		return { storage, reads, watcher, getBadge: () => badge, wasBadgeDisposed: () => badgeDisposed };
	}

	test('builds Gmail and Graph importance filters', () => {
		assert.strictEqual(directMailGmailQuery('', true, 'high'), 'is:unread is:important');
		assert.strictEqual(directMailGmailQuery('from:boss', false, ''), 'from:boss');
		assert.strictEqual(directMailGmailQuery('from:boss', true, 'high'), 'from:boss is:unread is:important');
		assert.strictEqual(directMailGraphFilter(true, 'high'), "isRead eq false and importance eq 'high'");
		assert.strictEqual(directMailGraphFilter(false, ''), '');
		assert.strictEqual(directMailGraphFilter(false, 'bogus'), '');
	});

	test('reports new important mail only from connected mailboxes, silently and once', async () => {
		const { storage, reads, watcher, getBadge, wasBadgeDisposed } = fixture();
		assert.strictEqual(await watcher.checkNow('account'), 1);
		assert.deepStrictEqual(reads.map(read => read.connection), ['outlook-conn']);
		assert.deepStrictEqual(reads[0].args, { folder: 'inbox', unreadOnly: true, importance: 'high', limit: 10 });
		const alerts = new AssistantConversationStore(storage).read('account').mailAlerts ?? [];
		assert.strictEqual(alerts.length, 1);
		assert.strictEqual(alerts[0].subject, 'Urgent: review');
		assert.ok(getBadge() && /1 important/.test(getBadge()!.text));

		assert.strictEqual(await watcher.checkNow('account'), 0);
		assert.strictEqual(reads.length, 2);

		watcher.markSeen('account');
		assert.ok(wasBadgeDisposed());
	});

	test('a failing mailbox read is skipped without reporting', async () => {
		const { reads, watcher } = fixture({ failReads: true });
		assert.strictEqual(await watcher.checkNow('account'), 0);
		assert.strictEqual(reads.length, 1);
	});

	test('status reflects the enabled setting and setEnabled updates it', async () => {
		const { watcher } = fixture();
		assert.strictEqual(watcher.status('account').enabled, true);
		assert.deepStrictEqual(watcher.status('account').connections, [
			{ label: 'Outlook', connected: true },
			{ label: 'Gmail', connected: false },
		]);
		await watcher.setEnabled(false);
		assert.strictEqual(watcher.isEnabled(), false);
		assert.strictEqual(watcher.status('account').enabled, false);
	});
});