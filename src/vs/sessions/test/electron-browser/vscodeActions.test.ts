/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import assert from 'assert';
import { DeferredPromise } from '../../../base/common/async.js';
import { mock, upcastPartial } from '../../../base/test/common/mock.js';
import { ensureNoDisposablesAreLeakedInTestSuite } from '../../../base/test/common/utils.js';
import { INativeHostService } from '../../../platform/native/common/native.js';
import { IOpenedMainWindow, IOpenEmptyWindowOptions, IOpenWindowOptions, IWindowOpenable, isFolderToOpen } from '../../../platform/window/common/window.js';
import { constObservable, observableValue } from '../../../base/common/observable.js';
import { hasKey } from '../../../base/common/types.js';
import { URI } from '../../../base/common/uri.js';
import { OpenSessionInVSCodeAction, openSessionInVSCode, returnToVSCodeEditor, shouldShowReturnToVSCodeEditor } from '../../electron-browser/actions/vscodeActions.js';
import { getChatSessionToOpenInEditor } from '../../../workbench/contrib/sessions/browser/openInVSCodeUtils.js';
import { IChat, ISession, SessionStatus } from '../../../workbench/services/sessions/common/session.js';
import { ISessionsProvidersService } from '../../../workbench/contrib/sessions/browser/sessionsProvidersService.js';
import { getEntryAddress, IRemoteAgentHostEntry, IRemoteAgentHostService, RemoteAgentHostEntryType } from '../../../platform/agentHost/common/remoteAgentHostService.js';
import { Codicon } from '../../../base/common/codicons.js';
import { agentHostAuthority, agentHostLabelFormatter, toAgentHostUri } from '../../../platform/agentHost/common/agentHostUri.js';
import { IAgentHostSessionsProvider } from '../../../workbench/contrib/sessions/common/agentHostSessionsProvider.js';
import { ISessionsProvider } from '../../../workbench/contrib/sessions/common/sessionsProvider.js';
import { LabelService } from '../../../workbench/services/label/common/labelService.js';
import { workbenchInstantiationService } from '../../../workbench/test/browser/workbenchTestServices.js';
import { splitRecentLabel } from '../../../base/common/labels.js';
import { ILabelService } from '../../../platform/label/common/label.js';
import { TestInstantiationService } from '../../../platform/instantiation/test/common/instantiationServiceMock.js';
import { IAgentHostConnectionsService } from '../../../platform/agentHost/common/agentHostConnectionsService.js';
import { remoteAgentHostSessionTypeId } from '../../../platform/agentHost/common/agentHostSessionType.js';
import { INotificationService } from '../../../platform/notification/common/notification.js';
import { IRemoteAgentHostLocationPreferenceService, RemoteAgentHostLocationPreference } from '../../../platform/agentHost/common/remoteAgentHostLocationPreference.js';
import { PathStyle } from '../../../platform/agentHost/common/state/protocol/common/commands.js';
import { ISSHAgentHostConnection, ISSHRemoteAgentHostService } from '../../../platform/agentHost/common/sshRemoteAgentHost.js';
import { ITunnelAgentHostService, TunnelGatewayServerType } from '../../../platform/agentHost/common/tunnelAgentHost.js';
import { IWorkbenchEnvironmentService } from '../../../workbench/services/environment/common/environmentService.js';
import { ICustomViewDescriptor, ICustomViewService } from '../../../workbench/contrib/sessions/browser/customView.js';
import { ASSISTANT_CUSTOM_VIEW_ID } from '../../../workbench/contrib/sessions/common/chatConstants.js';
import { AssistantWorktreeStatus, IAssistantAgentChats } from '../../contrib/sessions/browser/assistantAgentChats.js';

suite('VS Code Actions', () => {
	const disposables = ensureNoDisposablesAreLeakedInTestSuite();

	let instantiationService: TestInstantiationService;
	let labelService: LabelService;
	let notifications: string[];

	setup(() => {
		instantiationService = workbenchInstantiationService(undefined, disposables);
		labelService = disposables.add(instantiationService.createInstance(LabelService));
		instantiationService.stub(IWorkbenchEnvironmentService, upcastPartial<IWorkbenchEnvironmentService>({ remoteAuthority: undefined }));
		instantiationService.stub(ILabelService, labelService);
		instantiationService.stub(ISessionsProvidersService, upcastPartial<ISessionsProvidersService>({ getProvider: () => undefined }));
		instantiationService.stub(IRemoteAgentHostService, new class extends mock<IRemoteAgentHostService>() { }());
		instantiationService.stub(IAgentHostConnectionsService, upcastPartial<IAgentHostConnectionsService>({ resolveSessionResourceIdentity: () => undefined }));
		instantiationService.stub(IRemoteAgentHostLocationPreferenceService, upcastPartial<IRemoteAgentHostLocationPreferenceService>({ getPreference: () => undefined }));
		instantiationService.stub(ISSHRemoteAgentHostService, upcastPartial<ISSHRemoteAgentHostService>({ connections: [] }));
		instantiationService.stub(ITunnelAgentHostService, upcastPartial<ITunnelAgentHostService>({ getSelectedServerType: () => undefined }));
		notifications = [];
		instantiationService.stub(INotificationService, upcastPartial<INotificationService>({ info: message => { notifications.push(String(message)); } }));
	});

	test('repeated mode switches to the IDE share one attempt and can retry after failure', async () => {
		const firstHandoff = new DeferredPromise<void>();
		let openCount = 0;
		const workspace = URI.file('/workspace');
		instantiationService.stub(ICustomViewService, upcastPartial<ICustomViewService>({
			activeCustomView: constObservable(upcastPartial<ICustomViewDescriptor>({ id: ASSISTANT_CUSTOM_VIEW_ID })),
		}));
		instantiationService.stub(IAssistantAgentChats, upcastPartial<IAssistantAgentChats>({ target: { workspace, fileRoots: [workspace] } }));
		instantiationService.stub(INativeHostService, upcastPartial<INativeHostService>({
			openWindow: () => ++openCount === 1 ? firstHandoff.p : Promise.resolve(),
		}));
		const action = new OpenSessionInVSCodeAction();
		const switchToIDE = () => instantiationService.invokeFunction(accessor => action.run(accessor));
		const first = assert.rejects(switchToIDE(), /Could not open editor/);
		const repeated = switchToIDE();
		try {
			assert.strictEqual(openCount, 1);
		} finally {
			await firstHandoff.error(new Error('Could not open editor'));
			await Promise.all([first, repeated]);
		}
		await switchToIDE();
		assert.strictEqual(openCount, 2);
	});
	test('opens Assistant’s exact worktree and blocks fallback when it is missing or untrusted', async () => {
		const project = URI.file('/project');
		const checkout = URI.file('/project-feature');
		const opened: (IOpenEmptyWindowOptions | IWindowOpenable[] | undefined)[] = [];
		let status: AssistantWorktreeStatus = 'ready';
		instantiationService.stub(ICustomViewService, upcastPartial<ICustomViewService>({
			activeCustomView: constObservable(upcastPartial<ICustomViewDescriptor>({ id: ASSISTANT_CUSTOM_VIEW_ID })),
		}));
		instantiationService.stub(IAssistantAgentChats, upcastPartial<IAssistantAgentChats>({
			target: { workspace: project, fileRoots: [project, URI.file('/utilities')], worktree: { uri: checkout, label: 'feature' } },
			getWorktreeStatus: async worktree => { assert.deepStrictEqual(worktree.uri, checkout); return status; },
		}));
		instantiationService.stub(INativeHostService, new class extends mock<INativeHostService>() {
			override async openWindow(toOpen?: IOpenEmptyWindowOptions | IWindowOpenable[]): Promise<void> { opened.push(toOpen); }
		}());
		instantiationService.stub(INotificationService, upcastPartial<INotificationService>({ warn: message => { notifications.push(String(message)); } }));
		const action = new OpenSessionInVSCodeAction();
		const switchToIDE = () => instantiationService.invokeFunction(accessor => action.run(accessor));
		await switchToIDE();
		status = 'missing';
		await switchToIDE();
		status = 'untrusted';
		await switchToIDE();
		status = 'changed';
		await switchToIDE();
		assert.deepStrictEqual(opened, [[{ folderUri: checkout }]]);
		assert.match(notifications[0], /worktree is unavailable/);
		assert.match(notifications[1], /worktree needs trust/);
		assert.match(notifications[2], /Git checkout at the selected path has changed/);
	});

	test('Assistant switches to IDE with every saved root and the same window reuse as regular chat', async () => {
		const folders = [URI.file('/workspace/ui'), URI.file('/workspace/service')];
		const calls: { target: IOpenEmptyWindowOptions | IWindowOpenable[] | undefined; options: IOpenWindowOptions | undefined }[] = [];
		instantiationService.stub(ICustomViewService, upcastPartial<ICustomViewService>({
			activeCustomView: constObservable(upcastPartial<ICustomViewDescriptor>({ id: ASSISTANT_CUSTOM_VIEW_ID })),
		}));
		instantiationService.stub(IAssistantAgentChats, upcastPartial<IAssistantAgentChats>({
			target: { workspace: folders[0], fileRoots: folders },
		}));
		instantiationService.stub(INativeHostService, new class extends mock<INativeHostService>() {
			override async openWindow(target?: IOpenEmptyWindowOptions | IWindowOpenable[], options?: IOpenWindowOptions): Promise<void> { calls.push({ target, options }); }
		}());
		await instantiationService.invokeFunction(accessor => new OpenSessionInVSCodeAction().run(accessor));
		assert.deepStrictEqual(calls.map(({ target, options }) => ({
			folders: Array.isArray(target) ? target.filter(isFolderToOpen).map(folder => folder.folderUri) : [],
			reuse: options?.reuseWorkspaceContainingFolder,
			chat: options?.chatSessionToOpen,
		})), [{ folders, reuse: true, chat: undefined }]);
	});

	test('Assistant refuses a split workspace instead of choosing one of its connections', async () => {
		instantiationService.stub(ICustomViewService, upcastPartial<ICustomViewService>({
			activeCustomView: constObservable(upcastPartial<ICustomViewDescriptor>({ id: ASSISTANT_CUSTOM_VIEW_ID })),
		}));
		const project = URI.file('/workspace/ui');
		instantiationService.stub(IAssistantAgentChats, upcastPartial<IAssistantAgentChats>({
			target: { workspace: project, fileRoots: [project, URI.parse('vscode-remote://ssh-remote+fixture/workspace/service')] },
		}));
		const opened: unknown[] = [];
		instantiationService.stub(INativeHostService, new class extends mock<INativeHostService>() {
			override async openWindow(target?: IOpenEmptyWindowOptions | IWindowOpenable[]): Promise<void> { opened.push(target); }
		}());
		await instantiationService.invokeFunction(accessor => new OpenSessionInVSCodeAction().run(accessor));
		assert.deepStrictEqual(opened, []);
		assert.match(notifications[0], /different remote connections/);
	});

	test('a folderless Assistant stays on the current connection without borrowing a chat workspace', async () => {
		instantiationService.stub(ICustomViewService, upcastPartial<ICustomViewService>({
			activeCustomView: constObservable(upcastPartial<ICustomViewDescriptor>({ id: ASSISTANT_CUSTOM_VIEW_ID })),
		}));
		instantiationService.stub(IAssistantAgentChats, upcastPartial<IAssistantAgentChats>({ target: { fileRoots: [] } }));
		instantiationService.stub(IWorkbenchEnvironmentService, upcastPartial<IWorkbenchEnvironmentService>({ remoteAuthority: 'ssh-remote+assistant-fixture' }));
		const opened: unknown[] = [];
		instantiationService.stub(INativeHostService, new class extends mock<INativeHostService>() {
			override async openWindow(target?: IOpenEmptyWindowOptions | IWindowOpenable[]): Promise<void> { opened.push(target); }
		}());
		await instantiationService.invokeFunction(accessor => new OpenSessionInVSCodeAction().run(accessor));
		assert.deepStrictEqual(opened, [{ remoteAuthority: 'ssh-remote+assistant-fixture', reuseExistingWindow: true }]);
	});

	for (const source of [undefined, 'ssh-remote+source']) {
		for (const count of [0, 1, 2]) {
			for (const sent of [false, true]) {
				test(`retains ambient ownership for local or folderless workspaces (${source ?? 'local'}, ${count}, ${sent ? 'sent' : 'draft'})`, async () => {
					instantiationService.stub(IWorkbenchEnvironmentService, upcastPartial<IWorkbenchEnvironmentService>({ remoteAuthority: source }));
					const folders = Array.from({ length: count }, (_, i) => URI.file(`/local${i}`));
					const calls: { folders?: string[]; remote?: string | null; chat?: string }[] = [];
					instantiationService.stub(INativeHostService, new class extends mock<INativeHostService>() {
						override async openWindow(toOpen?: IOpenEmptyWindowOptions | IWindowOpenable[], options?: IOpenWindowOptions): Promise<void> { calls.push(Array.isArray(toOpen) ? { folders: toOpen.filter(isFolderToOpen).map(folder => folder.folderUri.toString()), chat: options?.chatSessionToOpen?.toString() } : { remote: toOpen?.remoteAuthority, chat: toOpen?.chatSessionToOpen?.toString() }); }
					}());
					await instantiationService.invokeFunction(openSessionInVSCode, createSession('ambient', sent, folders));
					const chat = sent && (!source || count === 0) ? 'test:/ambient' : undefined;
					assert.deepStrictEqual({ calls, notifications }, { calls: [count === 0 ? { remote: source ?? null, chat } : { folders: folders.map(folder => folder.toString()), chat }], notifications: sent && source && count > 0 ? ['This chat stays in Agents because the selected editor workspace uses a different connection.'] : [] });
				});
			}
		}
	}
	test('keeps the ambient remote when opening an editor without a selected session', async () => {
		instantiationService.stub(IWorkbenchEnvironmentService, upcastPartial<IWorkbenchEnvironmentService>({ remoteAuthority: 'ssh-remote+source' }));
		const calls: (IOpenEmptyWindowOptions | IWindowOpenable[] | undefined)[] = [];
		instantiationService.stub(INativeHostService, new class extends mock<INativeHostService>() {
			override async openWindow(toOpen?: IOpenEmptyWindowOptions | IWindowOpenable[]): Promise<void> { calls.push(toOpen); }
		}());
		await instantiationService.invokeFunction(openSessionInVSCode, undefined);
		assert.deepStrictEqual(calls, [{ remoteAuthority: 'ssh-remote+source', reuseExistingWindow: true }]);
	});
	test('keeps the ambient remote for an unsupported single-folder fallback', async () => {
		instantiationService.stub(IWorkbenchEnvironmentService, upcastPartial<IWorkbenchEnvironmentService>({ remoteAuthority: 'ssh-remote+source' }));
		const calls: (IOpenEmptyWindowOptions | IWindowOpenable[] | undefined)[] = [];
		instantiationService.stub(INativeHostService, new class extends mock<INativeHostService>() {
			override async openWindow(toOpen?: IOpenEmptyWindowOptions | IWindowOpenable[]): Promise<void> { calls.push(toOpen); }
		}());
		await instantiationService.invokeFunction(openSessionInVSCode, createSession('unsupported', true, [toAgentHostUri(URI.file('/unsupported'), agentHostAuthority('ws:other'))]));
		assert.deepStrictEqual(calls, [{ remoteAuthority: 'ssh-remote+source' }]);
	});
	test('keeps a registered remote conversation out of an explicit local workspace', async () => {
		const address = 'ssh:source';
		const authority = agentHostAuthority(address);
		const resource = URI.from({ scheme: remoteAgentHostSessionTypeId(authority, 'copilotcli'), path: '/conversation' });
		const provider: ISessionsProvider = upcastPartial<IAgentHostSessionsProvider>({ id: 'agenthost-test', remoteAddress: address });
		instantiationService.stub(ISessionsProvidersService, upcastPartial<ISessionsProvidersService>({ getProvider: <T extends ISessionsProvider>() => provider as T }));
		instantiationService.stub(IRemoteAgentHostService, upcastPartial<IRemoteAgentHostService>({ getEntryByAddress: () => ({ name: 'Source', connection: { type: RemoteAgentHostEntryType.SSH, address, hostName: 'source' } }) }));
		instantiationService.stub(IAgentHostConnectionsService, upcastPartial<IAgentHostConnectionsService>({ resolveSessionResourceIdentity: () => ({ connectionAuthority: authority, backendSession: URI.parse('copilotcli:/conversation') }) }));
		const chats: (string | undefined)[] = [];
		instantiationService.stub(INativeHostService, new class extends mock<INativeHostService>() {
			override async openWindow(_toOpen?: IOpenEmptyWindowOptions | IWindowOpenable[], options?: IOpenWindowOptions): Promise<void> { chats.push(options?.chatSessionToOpen?.toString()); }
		}());
		await instantiationService.invokeFunction(openSessionInVSCode, createSession('remote', true, [URI.file('/local')], resource));
		assert.deepStrictEqual({ chats, notifications }, { chats: [undefined], notifications: ['This chat stays in Agents because the selected editor workspace uses a different connection.'] });
	});
	for (const source of [undefined, 'ssh-remote+target', 'ssh-remote+other']) {
		for (const count of [1, 2]) {
			for (const sent of [false, true]) {
				test(`keeps explicit remote chats on their source connection (${source ?? 'local'}, ${count}, ${sent ? 'sent' : 'draft'})`, async () => {
					const folders = Array.from({ length: count }, (_, i) => URI.from({ scheme: 'vscode-remote', authority: 'ssh-remote+target', path: `/project${i}` }));
					const calls: { folders: string[]; chat: string | undefined }[] = [];
					instantiationService.stub(IWorkbenchEnvironmentService, upcastPartial<IWorkbenchEnvironmentService>({ remoteAuthority: source }));
					instantiationService.stub(INativeHostService, new class extends mock<INativeHostService>() {
						override async openWindow(toOpen?: IOpenEmptyWindowOptions | IWindowOpenable[], options?: IOpenWindowOptions): Promise<void> { calls.push({ folders: Array.isArray(toOpen) ? toOpen.filter(isFolderToOpen).map(folder => folder.folderUri.toString()) : [], chat: options?.chatSessionToOpen?.toString() }); }
					}());
					await instantiationService.invokeFunction(openSessionInVSCode, createSession('explicit-remote', sent, folders));
					assert.deepStrictEqual({ calls, notifications }, { calls: [{ folders: folders.map(folder => folder.toString()), chat: sent && source === 'ssh-remote+target' ? 'test:/explicit-remote' : undefined }], notifications: sent && source !== 'ssh-remote+target' ? ['This chat stays in Agents because the selected editor workspace uses a different connection.'] : [] });
				});
			}
		}
	}

	test('shows return action only when there is no other main window', () => {
		const currentWindow = createWindow(1);
		const otherWindow = createWindow(2);

		assert.deepStrictEqual({
			onlyAgentsWindow: shouldShowReturnToVSCodeEditor([currentWindow], currentWindow.id),
			agentsWindowNotListed: shouldShowReturnToVSCodeEditor([], currentWindow.id),
			otherWindowOpen: shouldShowReturnToVSCodeEditor([currentWindow, otherWindow], currentWindow.id),
			onlyOtherWindowListed: shouldShowReturnToVSCodeEditor([otherWindow], currentWindow.id),
		}, {
			onlyAgentsWindow: true,
			agentsWindowNotListed: true,
			otherWindowOpen: false,
			onlyOtherWindowListed: false,
		});
	});

	test('opens an editor window before closing the Agents window', async () => {
		const calls: string[] = [];
		const nativeHostService = new class extends mock<INativeHostService>() {
			override async openWindow(): Promise<void> {
				calls.push('open');
			}
			override async closeWindow(options?: { targetWindowId?: number }): Promise<void> {
				calls.push(`close:${options?.targetWindowId}`);
			}
		}();

		await returnToVSCodeEditor(nativeHostService, 7);

		assert.deepStrictEqual(calls, ['open', 'close:7']);
	});

	test('only transfers materialized sessions to the editor window', () => {
		const provisional = createSession('provisional', false);
		const materialized = createSession('materialized', true);
		const nestedDraft = createSession('nested-draft', false).mainChat.get();

		// Created per the provider's status: a draft carries no flag that says it is one.
		assert.deepStrictEqual({
			provisional: getChatSessionToOpenInEditor(provisional)?.toString(),
			provisionalChat: getChatSessionToOpenInEditor(provisional, provisional.mainChat.get())?.toString(),
			materialized: getChatSessionToOpenInEditor(materialized)?.toString(),
			materializedChat: getChatSessionToOpenInEditor(materialized, materialized.mainChat.get())?.toString(),
			nestedDraft: getChatSessionToOpenInEditor(materialized, nestedDraft)?.toString(),
			missing: getChatSessionToOpenInEditor(undefined),
		}, {
			provisional: undefined,
			provisionalChat: undefined,
			materialized: 'test:/materialized',
			materializedChat: 'test:/materialized/chat',
			nestedDraft: undefined,
			missing: undefined,
		});
	});

	test('hands off the selected chat after its first send, including active and failed conversations', () => {
		const session = createSession('created-parent', true);
		const status = observableValue<SessionStatus>('chatStatus', SessionStatus.Untitled);
		const resource = URI.parse('test:/nested');
		const chat = upcastPartial<IChat>({ resource, status });
		assert.strictEqual(getChatSessionToOpenInEditor(session, chat), undefined);
		for (const sentStatus of [SessionStatus.InProgress, SessionStatus.NeedsInput, SessionStatus.Completed, SessionStatus.Error]) {
			status.set(sentStatus, undefined);
			assert.strictEqual(getChatSessionToOpenInEditor(session, chat), resource);
		}
	});

	test('opens every active chat workspace folder in one editor window', async () => {
		const calls: { folders?: { uri: string; label: string | undefined }[]; forceNewWindow?: boolean; reuseWorkspaceContainingFolder?: boolean; chatSessionToOpen?: string }[] = [];
		const nativeHostService = new class extends mock<INativeHostService>() {
			override async openWindow(toOpen?: IOpenEmptyWindowOptions | IWindowOpenable[], options?: IOpenWindowOptions): Promise<void> {
				calls.push({
					folders: Array.isArray(toOpen) ? toOpen.filter(isFolderToOpen).map(openable => ({ uri: openable.folderUri.toString(), label: openable.label })) : undefined,
					forceNewWindow: options?.forceNewWindow,
					reuseWorkspaceContainingFolder: options?.reuseWorkspaceContainingFolder,
					chatSessionToOpen: options?.chatSessionToOpen?.toString(),
				});
			}
		}();
		const session = createSession('multi-folder', true, [URI.file('/repo-b'), URI.file('/repo-c')]);

		instantiationService.stub(INativeHostService, nativeHostService);
		await instantiationService.invokeFunction(openSessionInVSCode, session);

		assert.deepStrictEqual(calls, [{
			folders: [{ uri: 'file:///repo-b', label: undefined }, { uri: 'file:///repo-c', label: undefined }],
			forceNewWindow: true,
			reuseWorkspaceContainingFolder: true,
			chatSessionToOpen: 'test:/multi-folder',
		}]);
	});

	test('requests workspace-member reuse for an unsent draft without transferring its provisional chat', async () => {
		const calls: (IOpenWindowOptions | undefined)[] = [];
		instantiationService.stub(INativeHostService, new class extends mock<INativeHostService>() {
			override async openWindow(_toOpen?: IOpenEmptyWindowOptions | IWindowOpenable[], options?: IOpenWindowOptions): Promise<void> {
				calls.push(options);
			}
		}());
		await instantiationService.invokeFunction(openSessionInVSCode, createSession('draft', false, [URI.file('/repo')]));
		assert.deepStrictEqual(calls, [{ forceNewWindow: true, reuseWorkspaceContainingFolder: true, chatSessionToOpen: undefined }]);
	});

	test('opens an unsent nested chat folder without transferring the draft from its created parent', async () => {
		const calls: { folders: string[]; chat: string | undefined }[] = [];
		instantiationService.stub(INativeHostService, new class extends mock<INativeHostService>() {
			override async openWindow(toOpen?: IOpenEmptyWindowOptions | IWindowOpenable[], options?: IOpenWindowOptions): Promise<void> {
				calls.push({ folders: Array.isArray(toOpen) ? toOpen.filter(isFolderToOpen).map(folder => folder.folderUri.toString()) : [], chat: options?.chatSessionToOpen?.toString() });
			}
		}());
		const parent = createSession('created-parent', true, [URI.file('/parent')]);
		const draft = createSession('nested-draft', false, [URI.file('/nested')]).mainChat.get();
		await instantiationService.invokeFunction(openSessionInVSCode, parent, draft);
		assert.deepStrictEqual(calls, [{ folders: ['file:///nested'], chat: undefined }]);
	});

	test('requests reuse of a local editor without a session workspace', async () => {
		const calls: (IOpenEmptyWindowOptions | IWindowOpenable[] | undefined)[] = [];
		const nativeHostService = new class extends mock<INativeHostService>() {
			override async openWindow(toOpen?: IOpenEmptyWindowOptions | IWindowOpenable[]): Promise<void> {
				calls.push(toOpen);
			}
		}();

		instantiationService.stub(INativeHostService, nativeHostService);
		await instantiationService.invokeFunction(openSessionInVSCode, undefined);
		await instantiationService.invokeFunction(openSessionInVSCode, createSession('no-workspace', true));
		await instantiationService.invokeFunction(openSessionInVSCode, createSession('empty-workspace', true, []));

		assert.deepStrictEqual(calls, [
			{ remoteAuthority: null, reuseExistingWindow: true },
			{ remoteAuthority: null, reuseExistingWindow: true, chatSessionToOpen: URI.parse('test:/no-workspace') },
			{ remoteAuthority: null, reuseExistingWindow: true, chatSessionToOpen: URI.parse('test:/empty-workspace') },
		]);
	});

	test('keeps a folderless nested draft in Agents while requesting its editor window', async () => {
		const calls: (IOpenEmptyWindowOptions | IWindowOpenable[] | undefined)[] = [];
		instantiationService.stub(INativeHostService, new class extends mock<INativeHostService>() {
			override async openWindow(toOpen?: IOpenEmptyWindowOptions | IWindowOpenable[]): Promise<void> { calls.push(toOpen); }
		}());
		await instantiationService.invokeFunction(openSessionInVSCode, createSession('parent', true, [URI.file('/parent')]), createSession('draft', false).mainChat.get());
		assert.deepStrictEqual(calls, [{ remoteAuthority: null, reuseExistingWindow: true }]);
	});

	for (const [connection, remoteAuthority] of [
		[{ type: RemoteAgentHostEntryType.SSH, address: 'test:host', hostName: 'host', sshConfigHost: 'my-ssh-host' }, 'ssh-remote+my-ssh-host'],
		[{ type: RemoteAgentHostEntryType.Tunnel, tunnelId: 'host', clusterId: 'region' }, 'tunnel+host.region'],
		[{ type: RemoteAgentHostEntryType.WSL, address: 'test:host', distro: 'Ubuntu' }, 'wsl+Ubuntu'],
		[{ type: RemoteAgentHostEntryType.CloudSandbox, address: 'test:host', environmentId: 'environment' }, undefined],
		[{ type: RemoteAgentHostEntryType.WebSocket, address: 'test:host' }, undefined],
	] satisfies [IRemoteAgentHostEntry['connection'], string | undefined][]) {
		test(`opens an empty ${remoteAuthority ? 'remote' : 'local'} editor for a folderless ${connection.type} chat`, async () => {
			const calls: { toOpen?: IOpenEmptyWindowOptions | IWindowOpenable[]; options?: IOpenWindowOptions }[] = [];
			const provider: ISessionsProvider = upcastPartial<IAgentHostSessionsProvider>({ id: 'agenthost-test', remoteAddress: 'test:host' });
			instantiationService.stub(ISessionsProvidersService, { getProvider: <T extends ISessionsProvider>() => provider as T });
			instantiationService.stub(IRemoteAgentHostService, { getEntryByAddress: () => ({ name: 'Remote host', connection }) });
			instantiationService.stub(INativeHostService, new class extends mock<INativeHostService>() {
				override async openWindow(toOpen?: IOpenEmptyWindowOptions | IWindowOpenable[], options?: IOpenWindowOptions): Promise<void> {
					calls.push({ toOpen, options });
				}
			}());

			await instantiationService.invokeFunction(openSessionInVSCode, createSession('no-workspace', true));
			await instantiationService.invokeFunction(openSessionInVSCode, createSession('empty-workspace', true, []));
			await instantiationService.invokeFunction(
				openSessionInVSCode,
				createSession('parent', true, [toAgentHostUri(URI.file('/remote/project'), agentHostAuthority('test:host'))]),
				createSession('nested', true).mainChat.get(),
			);

			const toOpen = { remoteAuthority: remoteAuthority ?? null, reuseExistingWindow: true };
			assert.deepStrictEqual(calls, [
				{ toOpen: { ...toOpen, chatSessionToOpen: URI.parse('test:/no-workspace') }, options: undefined },
				{ toOpen: { ...toOpen, chatSessionToOpen: URI.parse('test:/empty-workspace') }, options: undefined },
				{ toOpen: { ...toOpen, chatSessionToOpen: URI.parse('test:/nested/chat') }, options: undefined },
			]);
		});
	}

	test('opens the clicked nested chat and its workspace instead of the active chat', async () => {
		const calls: { folders: string[]; chat: string | undefined }[] = [];
		const session = createSession('parent', true, [URI.file('/active')]);
		const folder = URI.file('/nested');
		const chat = upcastPartial<IChat>({
			resource: URI.parse('test:/nested-chat'),
			status: constObservable(SessionStatus.Completed),
			workspace: constObservable({
				...session.mainChat.get().workspace.get()!,
				folders: [{ root: folder, workingDirectory: folder, name: 'nested', description: undefined }],
			}),
		});
		instantiationService.stub(INativeHostService, new class extends mock<INativeHostService>() {
			override async openWindow(toOpen?: IOpenEmptyWindowOptions | IWindowOpenable[], options?: IOpenWindowOptions): Promise<void> {
				calls.push({
					folders: Array.isArray(toOpen) ? toOpen.filter(isFolderToOpen).map(folder => folder.folderUri.toString()) : [],
					chat: options?.chatSessionToOpen?.toString(),
				});
			}
		}());

		await instantiationService.invokeFunction(openSessionInVSCode, session, chat);

		assert.deepStrictEqual(calls, [{ folders: ['file:///nested'], chat: 'test:/nested-chat' }]);
	});

	for (const [connection, serverType, remoteAuthority, reason] of [
		[{ type: RemoteAgentHostEntryType.SSH, address: 'ssh:host', hostName: 'host' }, 'standalone', 'ssh-remote+host', undefined],
		[{ type: RemoteAgentHostEntryType.SSH, address: 'ssh:host', hostName: 'host' }, 'editor', 'ssh-remote+host', 'differentAgentHost'],
		[{ type: RemoteAgentHostEntryType.WSL, address: 'wsl:Ubuntu', distro: 'Ubuntu' }, undefined, 'wsl+Ubuntu', undefined],
		[{ type: RemoteAgentHostEntryType.WebSocket, address: 'ws:host' }, undefined, null, 'unsupported'],
	] satisfies [IRemoteAgentHostEntry['connection'], TunnelGatewayServerType | undefined, string | null, 'differentAgentHost' | 'unsupported' | undefined][]) {
		test(`preserves the host and validates the identity for a folderless ${connection.type} conversation (${serverType ?? 'default'})`, async () => {
			const entry: IRemoteAgentHostEntry = { name: 'Remote host', connection };
			const address = getEntryAddress(entry);
			const authority = agentHostAuthority(address);
			const resource = URI.from({ scheme: remoteAgentHostSessionTypeId(authority, 'copilotcli'), path: '/folderless', fragment: 'nested-chat' });
			const calls: (IOpenEmptyWindowOptions | IWindowOpenable[] | undefined)[] = [];
			instantiationService.stub(INativeHostService, new class extends mock<INativeHostService>() {
				override async openWindow(toOpen?: IOpenEmptyWindowOptions | IWindowOpenable[]): Promise<void> { calls.push(toOpen); }
			}());
			const provider: ISessionsProvider = upcastPartial<IAgentHostSessionsProvider>({ id: 'agenthost-test', remoteAddress: address });
			instantiationService.stub(ISessionsProvidersService, upcastPartial<ISessionsProvidersService>({ getProvider: <T extends ISessionsProvider>() => provider as T }));
			instantiationService.stub(IRemoteAgentHostService, upcastPartial<IRemoteAgentHostService>({ getEntryByAddress: () => entry }));
			instantiationService.stub(IAgentHostConnectionsService, upcastPartial<IAgentHostConnectionsService>({ resolveSessionResourceIdentity: () => ({ connectionAuthority: authority, backendSession: URI.parse('copilotcli:/folderless') }) }));
			instantiationService.stub(ISSHRemoteAgentHostService, upcastPartial<ISSHRemoteAgentHostService>({ connections: [upcastPartial<ISSHAgentHostConnection>({ localAddress: address, serverType })] }));
			await instantiationService.invokeFunction(openSessionInVSCode, createSession('remote', true, undefined, resource));
			assert.deepStrictEqual(calls, [{ remoteAuthority, reuseExistingWindow: true, ...(!reason ? { chatSessionToOpen: resource.with({ scheme: 'agent-host-copilotcli' }) } : {}) }]);
			assert.deepStrictEqual(notifications, reason === 'differentAgentHost' ? ['This chat can\'t be opened in the editor because it uses a different agent host on the remote machine.'] : reason === 'unsupported' ? ['Editor windows can\'t open chats from this kind of remote.'] : []);
		});
	}

	for (const folders of [
		[toAgentHostUri(URI.file('C:\\Users\\test\\project'), agentHostAuthority('cloudsandbox:environment'))],
		[URI.file('/local/project'), toAgentHostUri(URI.file('/remote/project'), agentHostAuthority('ws:host'))],
	]) {
		test(`handles unsupported remote folders without opening a partial workspace (${folders.length} folders)`, async () => {
			const calls: { toOpen?: IOpenEmptyWindowOptions | IWindowOpenable[]; options?: IOpenWindowOptions }[] = [];
			const session = createSession('unsupported', true, folders);
			const provider: ISessionsProvider = upcastPartial<IAgentHostSessionsProvider>({ id: 'agenthost-test', remoteAddress: 'cloudsandbox:environment' });
			instantiationService.stub(INativeHostService, new class extends mock<INativeHostService>() {
				override async openWindow(toOpen?: IOpenEmptyWindowOptions | IWindowOpenable[], options?: IOpenWindowOptions): Promise<void> {
					calls.push({ toOpen, options });
				}
			}());
			instantiationService.stub(ISessionsProvidersService, { getProvider: <T extends ISessionsProvider>() => provider as T });
			instantiationService.stub(IRemoteAgentHostService, {
				getEntryByAddress: () => ({
					name: 'Windows dev box',
					connection: { type: RemoteAgentHostEntryType.CloudSandbox, address: 'cloudsandbox:environment', environmentId: 'environment', environmentKind: 'user-local' },
				}),
			});
			await instantiationService.invokeFunction(openSessionInVSCode, session);
			assert.deepStrictEqual({
				calls,
				notifications,
				session: session.resource.toString(),
				folders: session.mainChat.get().workspace.get()?.folders.map(folder => folder.workingDirectory.toString()),
			}, {
				calls: folders.length === 1 ? [{ toOpen: { remoteAuthority: null }, options: undefined }] : [],
				notifications: folders.length === 1 ? [] : ['This workspace cannot be opened in one editor window because its folders use unsupported or different remote connections.'],
				session: 'test:/unsupported',
				folders: folders.map(folder => folder.toString()),
			});
		});
	}
	for (const [name, otherRoot] of [
		['another wrapped agent host', toAgentHostUri(URI.file('/other'), agentHostAuthority('ssh:other'))],
		['another editor remote', URI.parse('vscode-remote://ssh-remote+other/other')],
		['a local folder', URI.file('/local')],
	] as const) {
		test(`rejects a multi-root remote workspace containing ${name} before opening windows`, async () => {
			const address = 'ssh:host';
			const folders = [toAgentHostUri(URI.file('/project'), agentHostAuthority(address)), otherRoot];
			const session = createSession('mixed-hosts', true, folders);
			let opens = 0;
			instantiationService.stub(INativeHostService, new class extends mock<INativeHostService>() {
				override async openWindow(): Promise<void> { opens++; }
			}());
			const provider: ISessionsProvider = upcastPartial<IAgentHostSessionsProvider>({ id: 'agenthost-test', remoteAddress: address });
			instantiationService.stub(ISessionsProvidersService, upcastPartial<ISessionsProvidersService>({ getProvider: <T extends ISessionsProvider>() => provider as T }));
			instantiationService.stub(IRemoteAgentHostService, upcastPartial<IRemoteAgentHostService>({ getEntryByAddress: () => ({ name: 'Remote host', connection: { type: RemoteAgentHostEntryType.SSH, address, hostName: 'host' } }) }));
			await instantiationService.invokeFunction(openSessionInVSCode, session);
			assert.deepStrictEqual({ opens, notifications, folders: session.mainChat.get().workspace.get()?.folders.map(folder => folder.workingDirectory.toString()) }, { opens: 0, notifications: ['This workspace cannot be opened in one editor window because its folders use unsupported or different remote connections.'], folders: folders.map(folder => folder.toString()) });
		});
	}
	for (const [kind, folders] of [
		['wrapped', ['/one', '/two'].map(path => toAgentHostUri(URI.file(path), agentHostAuthority('ssh:other')))],
		['editor', ['/one', '/two'].map(path => URI.from({ scheme: 'vscode-remote', authority: 'ssh-remote+other', path }))],
	] as const) {
		test(`rejects a complete ${kind} root list owned by a different provider host`, async () => {
			const address = 'ssh:host';
			let opens = 0;
			instantiationService.stub(INativeHostService, new class extends mock<INativeHostService>() {
				override async openWindow(): Promise<void> { opens++; }
			}());
			const provider: ISessionsProvider = upcastPartial<IAgentHostSessionsProvider>({ id: 'agenthost-test', remoteAddress: address });
			instantiationService.stub(ISessionsProvidersService, upcastPartial<ISessionsProvidersService>({ getProvider: <T extends ISessionsProvider>() => provider as T }));
			instantiationService.stub(IRemoteAgentHostService, upcastPartial<IRemoteAgentHostService>({ getEntryByAddress: () => ({ name: 'Remote host', connection: { type: RemoteAgentHostEntryType.SSH, address, hostName: 'host' } }) }));
			await instantiationService.invokeFunction(openSessionInVSCode, createSession('other-provider', true, folders));
			assert.deepStrictEqual({ opens, notifications }, { opens: 0, notifications: ['This workspace cannot be opened in one editor window because its folders use unsupported or different remote connections.'] });
		});
	}
	for (const folders of [
		[URI.file('/local'), URI.parse('file://fileserver/share/project')],
		[URI.parse('vscode-remote://wsl+Ubuntu/one'), URI.parse('vscode-remote://WSL+ubuntu/two')],
	]) {
		test(`retains all compatible ${folders[0].scheme} roots`, async () => {
			instantiationService.stub(IWorkbenchEnvironmentService, upcastPartial<IWorkbenchEnvironmentService>({ remoteAuthority: folders[0].scheme === 'vscode-remote' ? folders[0].authority : undefined }));
			const opened: string[][] = [];
			instantiationService.stub(INativeHostService, new class extends mock<INativeHostService>() {
				override async openWindow(toOpen?: IOpenEmptyWindowOptions | IWindowOpenable[]): Promise<void> { opened.push(Array.isArray(toOpen) ? toOpen.filter(isFolderToOpen).map(folder => folder.folderUri.toString()) : []); }
			}());
			await instantiationService.invokeFunction(openSessionInVSCode, createSession('compatible-hosts', true, folders));
			assert.deepStrictEqual({ opened, notifications }, { opened: [folders.map(folder => folder.toString())], notifications: [] });
		});
	}

	for (const [connection, selectedServerType, preference] of [
		[{ type: RemoteAgentHostEntryType.SSH, address: 'ssh:host', hostName: 'host', sshConfigHost: 'host' }, 'standalone', undefined],
		[{ type: RemoteAgentHostEntryType.SSH, address: 'ssh:host', hostName: 'host', sshConfigHost: 'host' }, 'editor', undefined],
		[{ type: RemoteAgentHostEntryType.SSH, address: 'ssh:host', hostName: 'host', sshConfigHost: 'host' }, 'editor', 'editor'],
		[{ type: RemoteAgentHostEntryType.SSH, address: 'ssh:host', hostName: 'host', sshConfigHost: 'host' }, 'editor', 'dedicated'],
		[{ type: RemoteAgentHostEntryType.SSH, address: 'ssh:host', hostName: 'host', sshConfigHost: 'host' }, undefined, undefined],
		[{ type: RemoteAgentHostEntryType.Tunnel, tunnelId: 'host', clusterId: 'region' }, 'standalone', undefined],
		[{ type: RemoteAgentHostEntryType.Tunnel, tunnelId: 'host', clusterId: 'region' }, 'editor', undefined],
		[{ type: RemoteAgentHostEntryType.WSL, address: 'wsl:Ubuntu', distro: 'Ubuntu' }, undefined, undefined],
		[{ type: RemoteAgentHostEntryType.DevContainer, address: 'devcontainer:host', hostPath: '/source/project' }, undefined, undefined],
		[{ type: RemoteAgentHostEntryType.DevContainer, address: 'devcontainer:environment', environmentId: 'environment', volume: { volumeName: 'ubuntu', folder: 'ubuntu' }, hostAuthority: 'ssh-remote+build' }, undefined, undefined],
	] satisfies [IRemoteAgentHostEntry['connection'], TunnelGatewayServerType | undefined, RemoteAgentHostLocationPreference | undefined][]) {
		test(`preserves supported ${connection.type}${hasKey(connection, { environmentId: true }) ? ' environment' : ''} workspace and hands off chats from shared hosts (${selectedServerType ?? 'unspecified'}, ${preference ?? 'no preference'})`, async () => {
			const calls: { folders?: { uri: URI; label: string | undefined }[]; options?: IOpenWindowOptions }[] = [];
			const entry: IRemoteAgentHostEntry = { name: 'Remote host', connection };
			const address = getEntryAddress(entry);
			const folder = toAgentHostUri(URI.file('/remote/project'), agentHostAuthority(address));
			const provider: ISessionsProvider = upcastPartial<IAgentHostSessionsProvider>({ id: 'agenthost-test', remoteAddress: address });
			const formatter = agentHostLabelFormatter(folder.authority, PathStyle.Posix);
			disposables.add(labelService.registerFormatter({
				...formatter,
				formatting: { ...formatter.formatting, workspaceSuffix: entry.name },
			}));
			instantiationService.stub(INativeHostService, new class extends mock<INativeHostService>() {
				override async openWindow(toOpen?: IOpenEmptyWindowOptions | IWindowOpenable[], options?: IOpenWindowOptions): Promise<void> {
					calls.push({ folders: Array.isArray(toOpen) ? toOpen.filter(isFolderToOpen).map(openable => ({ uri: openable.folderUri, label: openable.label })) : undefined, options });
				}
			}());
			instantiationService.stub(ISessionsProvidersService, { getProvider: <T extends ISessionsProvider>() => provider as T });
			instantiationService.stub(IRemoteAgentHostService, { getEntryByAddress: () => entry });
			instantiationService.stub(IAgentHostConnectionsService, upcastPartial<IAgentHostConnectionsService>({
				resolveSessionResourceIdentity: resource => ({ connectionAuthority: folder.authority, backendSession: resource.with({ scheme: 'copilotcli', fragment: '' }) }),
			}));
			instantiationService.stub(IRemoteAgentHostLocationPreferenceService, upcastPartial<IRemoteAgentHostLocationPreferenceService>({ getPreference: key => key === address ? preference : undefined }));
			instantiationService.stub(ISSHRemoteAgentHostService, upcastPartial<ISSHRemoteAgentHostService>({
				connections: [
					upcastPartial<ISSHAgentHostConnection>({ localAddress: 'ssh:another-host', serverType: 'editor' }),
					upcastPartial<ISSHAgentHostConnection>({ localAddress: address, serverType: selectedServerType }),
				],
			}));
			instantiationService.stub(ITunnelAgentHostService, upcastPartial<ITunnelAgentHostService>({ getSelectedServerType: () => selectedServerType }));
			const resource = URI.from({ scheme: remoteAgentHostSessionTypeId(folder.authority, 'copilotcli'), path: '/remote', fragment: 'chat' });
			await instantiationService.invokeFunction(openSessionInVSCode, createSession('remote', true, [folder], resource));
			const sharesAgentHost = selectedServerType !== 'editor' || preference === 'editor';
			assert.deepStrictEqual({
				folders: calls[0].folders?.map(folder => ({ scheme: folder.uri.scheme, path: folder.uri.path, label: folder.label, title: folder.label && splitRecentLabel(folder.label).name })),
				options: calls[0].options,
				notifications,
			}, {
				folders: [{ scheme: 'vscode-remote', path: '/remote/project', label: '/remote/project [Remote host]', title: 'project [Remote host]' }],
				options: { forceNewWindow: true, reuseWorkspaceContainingFolder: true, chatSessionToOpen: sharesAgentHost ? URI.from({ scheme: 'agent-host-copilotcli', path: '/remote', fragment: 'chat' }) : undefined },
				notifications: sharesAgentHost ? [] : ['The session\'s folder was opened in a new editor window, but the session can\'t be opened there. This window is connected to a different agent host on the remote machine than the one its editor windows use.'],
			});
		});
	}

	test('preserves Windows remote labels for every folder opened through a tunnel', async () => {
		const calls: { uri: URI; label: string | undefined; title: string | undefined }[][] = [];
		const authority = agentHostAuthority('test:host');
		const folders = ['/c:/remote/project', '/d:/other/repository'].map(path => toAgentHostUri(URI.from({ scheme: 'file', path }), authority));
		const formatter = agentHostLabelFormatter(authority, PathStyle.Windows);
		disposables.add(labelService.registerFormatter({
			...formatter,
			formatting: { ...formatter.formatting, workspaceSuffix: 'Windows dev box' },
		}));
		const provider: ISessionsProvider = upcastPartial<IAgentHostSessionsProvider>({ id: 'agenthost-test', remoteAddress: 'test:host' });
		instantiationService.stub(INativeHostService, new class extends mock<INativeHostService>() {
			override async openWindow(toOpen?: IOpenEmptyWindowOptions | IWindowOpenable[]): Promise<void> {
				calls.push(Array.isArray(toOpen) ? toOpen.filter(isFolderToOpen).map(openable => ({
					uri: openable.folderUri,
					label: openable.label,
					title: openable.label && splitRecentLabel(openable.label).name,
				})) : []);
			}
		}());
		instantiationService.stub(ISessionsProvidersService, { getProvider: <T extends ISessionsProvider>() => provider as T });
		instantiationService.stub(IRemoteAgentHostService, {
			getEntryByAddress: () => ({ name: 'Windows dev box', connection: { type: RemoteAgentHostEntryType.Tunnel, tunnelId: 'host', clusterId: 'region', label: 'my-tunnel' } }),
		});
		await instantiationService.invokeFunction(openSessionInVSCode, createSession('remote', true, folders));

		assert.deepStrictEqual(calls, [[{
			uri: URI.from({ scheme: 'vscode-remote', authority: 'tunnel+my-tunnel', path: '/c:/remote/project' }),
			label: 'C:\\remote\\project [Windows dev box]',
			title: 'project [Windows dev box]',
		}, {
			uri: URI.from({ scheme: 'vscode-remote', authority: 'tunnel+my-tunnel', path: '/d:/other/repository' }),
			label: 'D:\\other\\repository [Windows dev box]',
			title: 'repository [Windows dev box]',
		}]]);
		calls.length = 0;
		const servers = ['fileserver', 'other-server'];
		const uncFolders = servers.map(server => toAgentHostUri(URI.from({ scheme: 'file', authority: server, path: '/share/project # 100%' }), authority));
		await instantiationService.invokeFunction(openSessionInVSCode, createSession('remote-unc', true, uncFolders));
		assert.deepStrictEqual(calls, [servers.map(server => ({
			uri: URI.from({ scheme: 'vscode-remote', authority: 'tunnel+my-tunnel', path: '//' + server + '/share/project # 100%' }),
			label: '\\\\' + server + '\\share\\project # 100% [Windows dev box]',
			title: 'project # 100% [Windows dev box]',
		}))]);
	});

	test('propagates a failure to open the fallback Editor window', async () => {
		const error = new Error('Editor window could not be opened');
		instantiationService.stub(INativeHostService, new class extends mock<INativeHostService>() {
			override async openWindow(): Promise<void> { throw error; }
		}());
		instantiationService.stub(ISessionsProvidersService, { getProvider: () => undefined });
		await assert.rejects(instantiationService.invokeFunction(
			openSessionInVSCode,
			createSession('unsupported', true, [toAgentHostUri(URI.file('/remote/project'), agentHostAuthority('ws:host'))]),
		), error);
	});

	test('does not replace an unexpected supported handoff failure with an empty window', async () => {
		const calls: string[] = [];
		const error = new Error('Remote resolver failed');
		instantiationService.stub(INativeHostService, new class extends mock<INativeHostService>() {
			override async openWindow(toOpen?: IOpenEmptyWindowOptions | IWindowOpenable[]): Promise<void> {
				calls.push(Array.isArray(toOpen) ? 'workspace' : 'empty');
				throw error;
			}
		}());
		await assert.rejects(instantiationService.invokeFunction(
			openSessionInVSCode,
			createSession('supported', true, [URI.file('/local/project')]),
		), error);
		assert.deepStrictEqual(calls, ['workspace']);
	});
});

function createWindow(id: number): IOpenedMainWindow {
	return {
		id,
		title: `Window ${id}`,
		dirty: false,
	};
}

/** A provider's session: created per its status, with no pane state; `folders` are its main chat's workspace. */
function createSession(id: string, isCreated: boolean, folders?: URI[], resource = URI.from({ scheme: 'test', path: `/${id}` })): ISession {
	const workspace = constObservable(folders ? {
		uri: URI.from({ scheme: 'test', path: `/${id}/workspace` }),
		label: id,
		icon: Codicon.folder,
		folders: folders.map((folder, index) => ({
			root: folder,
			workingDirectory: folder,
			name: `Repository ${index + 1}`,
			description: undefined,
		})),
		requiresWorkspaceTrust: false,
		isVirtualWorkspace: false,
	} : undefined);
	const mainChat = new class extends mock<IChat>() {
		override readonly resource = URI.from({ scheme: 'test', path: `/${id}/chat` });
		override readonly status = constObservable(isCreated ? SessionStatus.Completed : SessionStatus.Untitled);
		override readonly workspace = workspace;
	}();
	return new class extends mock<ISession>() {
		override readonly resource = resource;
		override readonly providerId = 'test';
		override readonly status = constObservable(isCreated ? SessionStatus.Completed : SessionStatus.Untitled);
		override readonly mainChat = constObservable(mainChat);
	}();
}
