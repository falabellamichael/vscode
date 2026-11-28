/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Signal-C. All rights reserved.
 *--------------------------------------------------------------------------------------------*/

import { RunOnceScheduler } from '../../../../base/common/async.js';
import { CancellationTokenSource } from '../../../../base/common/cancellation.js';
import { Emitter, Event } from '../../../../base/common/event.js';
import { Disposable } from '../../../../base/common/lifecycle.js';
import { autorun } from '../../../../base/common/observable.js';
import { localize } from '../../../../nls.js';
import { CommandsRegistry } from '../../../../platform/commands/common/commands.js';
import { IConfigurationService } from '../../../../platform/configuration/common/configuration.js';
import { InstantiationType, registerSingleton } from '../../../../platform/instantiation/common/extensions.js';
import { createDecorator, ServicesAccessor } from '../../../../platform/instantiation/common/instantiation.js';
import { ILogService } from '../../../../platform/log/common/log.js';
import { IOpenerService } from '../../../../platform/opener/common/opener.js';
import { IQuickInputService } from '../../../../platform/quickinput/common/quickInput.js';
import { IStorageService, StorageScope, StorageTarget } from '../../../../platform/storage/common/storage.js';
import { URI } from '../../../../base/common/uri.js';
import { ISignalCDirectMail, ISignalCDirectMailConnection } from '../../../../workbench/contrib/mcp/browser/signalCDirectMail.js';
import { ASSISTANT_MAIL_WATCH_SETTING } from '../../../../workbench/contrib/sessions/common/sessionConfig.js';
import { IStatusbarService, IStatusbarEntryAccessor, StatusbarAlignment } from '../../../../workbench/services/statusbar/browser/statusbar.js';
import { IReachSubscriptionConnection } from '../../accountMenu/browser/reachSubscriptionConnection.js';
import { AssistantConversationStore, IAssistantMailAlert } from './assistantConversation.js';

const CHECK_INTERVAL_MS = 5 * 60_000;
const REPORTED_PREFIX = 'sessions.assistant.mailWatch.reported.';
const SEEN_PREFIX = 'sessions.assistant.mailWatch.seen.';
const REPORTED_CAP = 300;
const SHOW_ALERTS_COMMAND = 'assistant.mailWatch.showAlerts';

export interface IAssistantMailWatchStatus {
	readonly enabled: boolean;
	readonly lastCheck?: number;
	readonly connections: readonly { readonly label: string; readonly connected: boolean }[];
	readonly alerts: readonly IAssistantMailAlert[];
}

export const IAssistantMailWatcher = createDecorator<IAssistantMailWatcher>('assistantMailWatcher');
export interface IAssistantMailWatcher {
	readonly _serviceBrand: undefined;
	readonly onDidChange: Event<void>;
	isEnabled(): boolean;
	setEnabled(enabled: boolean): Promise<void>;
	status(accountId: string): IAssistantMailWatchStatus;
	/** One passive check right now; still only reads already-connected mailboxes. */
	checkNow(accountId: string): Promise<number>;
	markSeen(accountId: string): void;
	start(): void;
}

/**
 * Passive important-mail watcher. It only reads mailboxes that are already
 * connected, never starts servers or sign-in flows, and reports quietly:
 * a status-bar badge plus alerts in the Assistant's context — no toasts or dialogs.
 */
export class AssistantMailWatcher extends Disposable implements IAssistantMailWatcher {
	declare readonly _serviceBrand: undefined;
	private readonly check = this._register(new RunOnceScheduler(() => void this.checkDue(), CHECK_INTERVAL_MS));
	private readonly _onDidChange = this._register(new Emitter<void>());
	readonly onDidChange: Event<void> = this._onDidChange.event;
	private checking = false;
	private lastCheck: number | undefined;
	private currentAccount: string | undefined;
	private statusEntry: IStatusbarEntryAccessor | undefined;

	constructor(
		@IReachSubscriptionConnection private readonly reach: IReachSubscriptionConnection,
		@ISignalCDirectMail private readonly directMail: ISignalCDirectMail,
		@IConfigurationService private readonly configuration: IConfigurationService,
		@IStorageService private readonly storage: IStorageService,
		@IStatusbarService private readonly statusBar: IStatusbarService,
		@ILogService private readonly log: ILogService,
	) {
		super();
		this._register(this.reach.onDidChangeAccount(() => this.start()));
		this._register(this.configuration.onDidChangeConfiguration(e => {
			if (e.affectsConfiguration(ASSISTANT_MAIL_WATCH_SETTING)) { this.start(); }
		}));
		this._register(autorun(reader => {
			this.directMail.connections.read(reader);
			this.start();
		}));
		this._register(this.onDidChange(() => this.updateStatusBar()));
		this._register(CommandsRegistry.registerCommand(SHOW_ALERTS_COMMAND, accessor => this.showAlerts(accessor)));
	}

	override dispose(): void {
		this.statusEntry?.dispose();
		this.statusEntry = undefined;
		super.dispose();
	}

	start(): void {
		if (this.isEnabled() && this.reach.assistantSupported) {
			this.check.schedule(0);
		}
		void this.reach.getAccount().then(account => {
			this.currentAccount = account?.id;
			this.updateStatusBar();
		}, () => undefined);
	}

	isEnabled(): boolean {
		return this.configuration.getValue<boolean>(ASSISTANT_MAIL_WATCH_SETTING) ?? true;
	}

	async setEnabled(enabled: boolean): Promise<void> {
		await this.configuration.updateValue(ASSISTANT_MAIL_WATCH_SETTING, enabled);
		this._onDidChange.fire();
	}

	status(accountId: string): IAssistantMailWatchStatus {
		return {
			enabled: this.isEnabled(),
			lastCheck: this.lastCheck,
			connections: this.directMail.connections.get().map(connection => ({ label: connection.label, connected: !connection.needsSignIn })),
			alerts: new AssistantConversationStore(this.storage).read(accountId).mailAlerts ?? [],
		};
	}

	async checkNow(accountId: string): Promise<number> {
		this.currentAccount = accountId;
		return this.checkAccount(accountId);
	}

	markSeen(accountId: string): void {
		this.storage.store(SEEN_PREFIX + accountId, Date.now(), StorageScope.PROFILE, StorageTarget.MACHINE);
		this.updateStatusBar();
	}

	private async checkDue(): Promise<void> {
		if (this.checking || !this.isEnabled()) { return; }
		this.checking = true;
		try {
			const account = await this.reach.getAccount();
			this.currentAccount = account?.id;
			if (account?.id && account.subscriptionActive) {
				await this.checkAccount(account.id);
			}
		} catch (error) {
			this.log.debug('[AssistantMailWatcher] periodic check failed', error);
		} finally {
			this.checking = false;
			this.lastCheck = Date.now();
			this._onDidChange.fire();
			if (this.isEnabled() && !this._store.isDisposed) { this.check.schedule(CHECK_INTERVAL_MS); }
		}
	}

	private async checkAccount(accountId: string): Promise<number> {
		const store = new AssistantConversationStore(this.storage);
		let reported = 0;
		for (const connection of this.directMail.connections.get()) {
			// Passive by design: a mailbox that needs sign-in is skipped, never prompted.
			if (connection.needsSignIn) { continue; }
			try {
				reported += await this.checkConnection(connection, accountId, store);
			} catch (error) {
				this.log.debug('[AssistantMailWatcher] mailbox check skipped', connection.label, error);
			}
		}
		if (reported) { this._onDidChange.fire(); }
		return reported;
	}

	private async checkConnection(connection: ISignalCDirectMailConnection, accountId: string, store: AssistantConversationStore): Promise<number> {
		const cancellation = new CancellationTokenSource();
		let raw: string;
		try {
			raw = await this.directMail.read(connection, { folder: 'inbox', unreadOnly: true, importance: 'high', limit: 10 }, cancellation.token);
		} finally {
			cancellation.dispose();
		}
		const parsed = JSON.parse(raw) as { messages?: readonly { id?: string; subject?: string; from?: string; received?: string; url?: string }[] };
		const reportedIds = this.readReported(accountId);
		const known = new Set(reportedIds);
		const alerts: IAssistantMailAlert[] = [];
		for (const message of parsed.messages ?? []) {
			if (!message.id || known.has(message.id)) { continue; }
			known.add(message.id);
			alerts.push({
				id: message.id, connectionId: connection.id, label: connection.label,
				subject: (message.subject || '(no subject)').slice(0, 200),
				from: message.from, received: message.received, url: message.url, at: Date.now(),
			});
		}
		if (!alerts.length) { return 0; }
		const added = store.recordMailAlerts(accountId, alerts);
		if (!added) { return 0; }
		this.writeReported(accountId, [...alerts.map(alert => alert.id), ...reportedIds].slice(0, REPORTED_CAP));
		return added;
	}

	private async showAlerts(accessor: ServicesAccessor): Promise<void> {
		const account = await accessor.get(IReachSubscriptionConnection).getAccount();
		if (!account?.id) { return; }
		const alerts = new AssistantConversationStore(this.storage).read(account.id).mailAlerts ?? [];
		this.markSeen(account.id);
		if (!alerts.length) { return; }
		const pick = await accessor.get(IQuickInputService).pick(alerts.map(alert => ({
			label: alert.subject,
			description: alert.from ? `${alert.label} · ${alert.from}` : alert.label,
			detail: alert.received,
			alert,
		})), { title: localize('assistantMailWatch.title', "Important mail"), placeHolder: localize('assistantMailWatch.pick', "Select an email to open it") });
		if (pick?.alert.url) {
			await accessor.get(IOpenerService).open(URI.parse(pick.alert.url));
		}
	}

	private updateStatusBar(): void {
		const unseen = this.currentAccount ? this.unseenAlerts(this.currentAccount) : [];
		if (!unseen.length) {
			this.statusEntry?.dispose();
			this.statusEntry = undefined;
			return;
		}
		const name = localize('assistantMailWatch.entryName', "Assistant important mail");
		const text = localize('assistantMailWatch.badge', "$(mail) {0} important", unseen.length);
		const tooltip = localize('assistantMailWatch.tooltip', "Assistant: {0} unread important email(s). Click to review.", unseen.length);
		if (this.statusEntry) {
			this.statusEntry.update({ name, text, ariaLabel: tooltip, tooltip, command: SHOW_ALERTS_COMMAND });
		} else {
			this.statusEntry = this.statusBar.addEntry({ name, text, ariaLabel: tooltip, tooltip, command: SHOW_ALERTS_COMMAND }, 'assistant.mailWatch', StatusbarAlignment.LEFT, -100);
		}
	}

	private unseenAlerts(accountId: string): readonly IAssistantMailAlert[] {
		const seen = this.storage.getNumber(SEEN_PREFIX + accountId, StorageScope.PROFILE, 0);
		return (new AssistantConversationStore(this.storage).read(accountId).mailAlerts ?? []).filter(alert => alert.at > seen);
	}

	private readReported(accountId: string): readonly string[] {
		const stored = this.storage.getObject<readonly string[]>(REPORTED_PREFIX + accountId, StorageScope.PROFILE);
		return Array.isArray(stored) ? stored.filter((id): id is string => typeof id === 'string') : [];
	}

	private writeReported(accountId: string, ids: readonly string[]): void {
		this.storage.store(REPORTED_PREFIX + accountId, JSON.stringify(ids), StorageScope.PROFILE, StorageTarget.MACHINE);
	}
}

registerSingleton(IAssistantMailWatcher, AssistantMailWatcher, InstantiationType.Delayed);