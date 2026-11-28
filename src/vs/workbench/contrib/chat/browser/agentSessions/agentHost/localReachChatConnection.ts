/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { Sequencer } from '../../../../../../base/common/async.js';
import { Disposable, MutableDisposable } from '../../../../../../base/common/lifecycle.js';
import { localize, localize2 } from '../../../../../../nls.js';
import { Action2, registerAction2 } from '../../../../../../platform/actions/common/actions.js';
import { IAgentHostService } from '../../../../../../platform/agentHost/common/agentService.js';
import { AuthRequiredReason, NotificationType } from '../../../../../../platform/agentHost/common/state/sessionActions.js';
import { CommandsRegistry } from '../../../../../../platform/commands/common/commands.js';
import { InstantiationType, registerSingleton } from '../../../../../../platform/instantiation/common/extensions.js';
import { createDecorator, ServicesAccessor } from '../../../../../../platform/instantiation/common/instantiation.js';
import { ILogService } from '../../../../../../platform/log/common/log.js';
import { REACH_ACCOUNT_ORIGIN, REACH_CHAT_PROVIDER_ID } from '../../../../../../platform/native/common/native.js';
import { INotificationService, Severity } from '../../../../../../platform/notification/common/notification.js';
import { RemoteNameContext } from '../../../../../common/contextkeys.js';
import { registerWorkbenchContribution2, WorkbenchPhase } from '../../../../../common/contributions.js';
import { IReachSubscriptionConnection } from '../../../../../services/authentication/browser/reachSubscriptionConnection.js';
import { IWorkbenchEnvironmentService } from '../../../../../services/environment/common/environmentService.js';

export const ILocalReachChatConnection = createDecorator<ILocalReachChatConnection>('localReachChatConnection');
export interface ILocalReachChatConnection {
	readonly _serviceBrand: undefined;
	connect(): Promise<boolean>;
	prepareHistory(): Promise<void>;
}

/** Only the desktop's local management connection may receive the saved REACH credential. */
export class LocalReachChatConnection extends Disposable implements ILocalReachChatConnection {
	declare readonly _serviceBrand: undefined;
	private readonly queue = new Sequencer();
	private readonly rootListener = this._register(new MutableDisposable());
	private available = false;
	private revision = 0;
	private forwarded: string | undefined;

	constructor(
		@IAgentHostService private readonly host: IAgentHostService,
		@IReachSubscriptionConnection private readonly reach: IReachSubscriptionConnection,
		@IWorkbenchEnvironmentService environment: IWorkbenchEnvironmentService,
		@INotificationService private readonly notifications: INotificationService,
		@ILogService private readonly log: ILogService,
	) {
		super();
		if (environment.remoteAuthority) { return; }
		this._register(this.host.onDidNotification(notification => {
			if (this.available && notification.type === NotificationType.AuthRequired && notification.reason === AuthRequiredReason.Expired && notification.resource.resource === REACH_ACCOUNT_ORIGIN) {
				this.notifications.prompt(Severity.Info, localize('reachChat.expired', "Sign in to REACH to resume the current chat."), [
					{ label: localize('reachChat.reconnect', "Connect REACH Chat"), run: () => this.connect().then(() => {}) },
				]);
			}
		}));
		const bind = () => {
			this.invalidate();
			this.available = false;
			const update = () => {
				const available = !!this.provider();
				if (available === this.available) { return; }
				this.available = available;
				this.invalidate();
				if (available) { this.syncInBackground(); }
			};
			this.rootListener.value = this.host.rootState.onDidChange(update);
			update();
		};
		this._register(this.host.onAgentHostStart(bind));
		this._register(this.reach.onDidChangeAccount(() => {
			this.invalidate();
			if (this.available) { this.syncInBackground(); }
		}));
		bind();
	}

	async connect(): Promise<boolean> {
		if (!this.available) {
			this.notifications.warn(localize('reachChat.hostUnavailable', "REACH is unavailable on this agent host. Open a local Signal-C window and try again."));
			return false;
		}
		try {
			const account = await this.reach.signIn();
			if (!account) { return false; }
			const revision = this.revision;
			if (await this.sync(true)) { return true; }
			// A successful sign-in can notify secret listeners while the explicit connect is pending.
			if (revision !== this.revision && (await this.reach.getAccount())?.id === account.id && await this.sync(true)) { return true; }
		} catch {
			this.log.warn('Could not complete the local REACH chat sign-in.');
		}
		this.notifications.warn(localize('reachChat.connectFailed', "REACH could not connect this chat. Check your REACH subscription and try again."));
		return false;
	}

	private provider() {
		const root = this.host.rootState.value;
		return root && !(root instanceof Error) ? root.agents.find(agent => agent.provider === REACH_CHAT_PROVIDER_ID
			&& agent.protectedResources?.some(resource => resource.resource === REACH_ACCOUNT_ORIGIN)) : undefined;
	}

	private invalidate(): void {
		this.revision++;
		this.forwarded = undefined;
	}

	async prepareHistory(): Promise<void> {
		let revision: number;
		do {
			revision = this.revision;
			if (await this.sync(false)) { return; }
		} while (this.available && !this._store.isDisposed && revision !== this.revision);
		if (!await this.reach.getAccount()) {
			throw new Error(localize('reachChat.historySignIn', "Sign in with Connect REACH Chat, then reopen this conversation."));
		}
		throw new Error(localize('reachChat.historyConnect', "REACH could not verify your saved account. Reconnect with Connect REACH Chat, then reopen this conversation."));
	}

	private syncInBackground(): void {
		void this.sync(false).catch(() => this.log.warn('Could not synchronize the local REACH chat connection.'));
	}

	private sync(refresh: boolean): Promise<boolean> {
		const revision = this.revision;
		const current = () => this.available && !this._store.isDisposed && revision === this.revision;
		return this.queue.queue(async () => {
			if (!current()) { return false; }
			const credential = await this.reach.getLocalChatCredential();
			if (!current()) { return false; }
			const seconds = credential ? Math.ceil((Date.parse(credential.expiresAt) - Date.now()) / 1000) : 0;
			const token = seconds > 0 && credential ? credential.accessToken : '';
			if (this.forwarded === token && !(refresh && token && !this.provider()?.models.length)) { return !!token; }
			const result = await this.host.authenticate({ resource: REACH_ACCOUNT_ORIGIN, token, ...(token ? { expiresIn: seconds } : {}) });
			if (!current()) { return false; }
			if (result.authenticated) { this.forwarded = token; }
			return !!token && result.authenticated;
		});
	}
}

registerSingleton(ILocalReachChatConnection, LocalReachChatConnection, InstantiationType.Delayed);
class ReachChatStartup {
	constructor(@ILocalReachChatConnection _connection: ILocalReachChatConnection) { }
}
registerWorkbenchContribution2('workbench.contrib.localReachChatConnection', ReachChatStartup, WorkbenchPhase.BlockRestore);

CommandsRegistry.registerCommand('_workbench.chat.prepareReachHistory', accessor => accessor.get(ILocalReachChatConnection).prepareHistory());

registerAction2(class ConnectReachChatAction extends Action2 {
	constructor() {
		super({
			id: 'workbench.action.chat.connectReach',
			title: localize2('reachChat.connect', 'Connect REACH Chat'),
			f1: true,
			precondition: RemoteNameContext.isEqualTo(''),
		});
	}
	run(accessor: ServicesAccessor): Promise<boolean> {
		return accessor.get(ILocalReachChatConnection).connect();
	}
});
