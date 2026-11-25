/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import '../../browser/parts/media/titlebarpart.css';
import '../../../workbench/browser/media/style.css';
import assert from 'assert';
import { $, append } from '../../../base/browser/dom.js';
import { mainWindow } from '../../../base/browser/window.js';
import { Action } from '../../../base/common/actions.js';
import { Codicon } from '../../../base/common/codicons.js';
import { Event } from '../../../base/common/event.js';
import { toDisposable } from '../../../base/common/lifecycle.js';
import { OperatingSystem } from '../../../base/common/platform.js';
import { constObservable } from '../../../base/common/observable.js';
import { URI } from '../../../base/common/uri.js';
import { mock, upcastPartial } from '../../../base/test/common/mock.js';
import { ensureNoDisposablesAreLeakedInTestSuite } from '../../../base/test/common/utils.js';
import { agentHostAuthority, toAgentHostUri } from '../../../platform/agentHost/common/agentHostUri.js';
import { getEntryAddress, IRemoteAgentHostEntry, IRemoteAgentHostService, RemoteAgentHostEntryType } from '../../../platform/agentHost/common/remoteAgentHostService.js';
import { remoteAgentHostSessionTypeId } from '../../../platform/agentHost/common/agentHostSessionType.js';
import { IAgentHostConnectionsService } from '../../../platform/agentHost/common/agentHostConnectionsService.js';
import { IRemoteAgentHostLocationPreferenceService, RemoteAgentHostLocationPreference } from '../../../platform/agentHost/common/remoteAgentHostLocationPreference.js';
import { ISSHAgentHostConnection, ISSHRemoteAgentHostService } from '../../../platform/agentHost/common/sshRemoteAgentHost.js';
import { ITunnelAgentHostService } from '../../../platform/agentHost/common/tunnelAgentHost.js';
import { TestInstantiationService } from '../../../platform/instantiation/test/common/instantiationServiceMock.js';
import { NullHoverService } from '../../../platform/hover/test/browser/nullHoverService.js';
import { IKeybindingService } from '../../../platform/keybinding/common/keybinding.js';
import { INotificationService } from '../../../platform/notification/common/notification.js';
import { IOpenerService } from '../../../platform/opener/common/opener.js';
import { IProductService } from '../../../platform/product/common/productService.js';
import { getIconRegistry } from '../../../platform/theme/common/iconRegistry.js';
import { NullTelemetryService } from '../../../platform/telemetry/common/telemetryUtils.js';
import { ITelemetryService } from '../../../platform/telemetry/common/telemetry.js';
import { TestExperimentTriggerTelemetryService } from '../../../platform/telemetry/test/common/experimentTriggerTestUtils.js';
import { IConfigurationChangeEvent } from '../../../platform/configuration/common/configuration.js';
import { TestConfigurationService } from '../../../platform/configuration/test/common/testConfigurationService.js';
import { AGENTS_WINDOW_TITLE_BAR_ICONS_SETTING } from '../../../platform/chat/common/chatSettings.js';
import { Context, ContextKeyService } from '../../../platform/contextkey/browser/contextKeyService.js';
import { IsAuxiliaryWindowContext, IsSessionsWindowContext, LayoutModeContext } from '../../../workbench/common/contextkeys.js';
import { IsPhoneLayoutContext, SessionsWelcomeVisibleContext } from '../../../workbench/contrib/sessions/common/contextkeys.js';
import { registerAction2 } from '../../../platform/actions/common/actions.js';
import { KeybindingsRegistry } from '../../../platform/keybinding/common/keybindingsRegistry.js';
import { KeybindingResolver, ResultKind } from '../../../platform/keybinding/common/keybindingResolver.js';
import { ResolvedKeybindingItem } from '../../../platform/keybinding/common/resolvedKeybindingItem.js';
import { USLayoutResolvedKeybinding } from '../../../platform/keybinding/common/usLayoutResolvedKeybinding.js';
import { OpenInVSCodeAction } from '../../browser/actions/vscodeActions.js';
import { OpenInVSCodeTitleBarWidget } from '../../browser/widget/openInVSCodeWidget.js';
import { AgentsWindowMenus } from '../../browser/agentsWindowMenus.js';
import { IAgentHostSessionsProvider } from '../../../workbench/contrib/sessions/common/agentHostSessionsProvider.js';
import { ISessionsProvidersService } from '../../../workbench/contrib/sessions/browser/sessionsProvidersService.js';
import { ISessionsService } from '../../../workbench/contrib/sessions/browser/sessionsService.js';
import { ISessionsProvider } from '../../../workbench/contrib/sessions/common/sessionsProvider.js';
import { IChat, ISession, ISessionWorkspace, SessionStatus } from '../../../workbench/services/sessions/common/session.js';
import { parseExternalEditorWindowLinkUri } from '../../../platform/window/common/window.js';
import { IFocusedSessionHostService, ISessionHost, ISessionHostService } from '../../../workbench/services/sessions/browser/sessionHost.js';
import { ITestSessionsService, ITestSessionWithChat, stubSessionHostsFromSessionsService } from '../../../workbench/contrib/sessions/test/browser/sessionHostTestUtils.js';
import { IWorkbenchEnvironmentService } from '../../../workbench/services/environment/common/environmentService.js';

suite('Web Open in Editor', () => {
	const store = ensureNoDisposablesAreLeakedInTestSuite();
	function explicitRemoteFixture(source?: string) {
		const instantiation = store.add(new TestInstantiationService());
		const links: URI[] = [];
		const notifications: string[] = [];
		instantiation.stub(ISessionsProvidersService, upcastPartial<ISessionsProvidersService>({ getProvider: () => undefined }));
		instantiation.stub(IRemoteAgentHostService, upcastPartial<IRemoteAgentHostService>({ getEntryByAddress: () => undefined }));
		instantiation.stub(IAgentHostConnectionsService, upcastPartial<IAgentHostConnectionsService>({ resolveSessionResourceIdentity: () => undefined }));
		instantiation.stub(IRemoteAgentHostLocationPreferenceService, upcastPartial<IRemoteAgentHostLocationPreferenceService>({ getPreference: () => undefined }));
		instantiation.stub(ISSHRemoteAgentHostService, upcastPartial<ISSHRemoteAgentHostService>({ connections: [] }));
		instantiation.stub(ITunnelAgentHostService, upcastPartial<ITunnelAgentHostService>({ getSelectedServerType: () => undefined }));
		instantiation.stub(IWorkbenchEnvironmentService, upcastPartial<IWorkbenchEnvironmentService>({ remoteAuthority: source }));
		instantiation.stub(INotificationService, upcastPartial<INotificationService>({ info: message => { notifications.push(String(message)); } }));
		instantiation.stub(ITelemetryService, NullTelemetryService);
		instantiation.stub(IProductService, upcastPartial<IProductService>({ urlProtocol: 'signal-c' }));
		instantiation.stub(IOpenerService, upcastPartial<IOpenerService>({ open: async uri => { links.push(URI.parse(uri.toString())); return true; } }));
		return { instantiation, links, notifications };
	}
	for (const source of [undefined, 'ssh-remote+source']) {
		for (const count of [0, 1, 2]) {
			for (const sent of [false, true]) {
				test(`keeps ambient connection context for local or folderless handoff (${source ?? 'local'}, ${count}, ${sent ? 'sent' : 'draft'})`, async () => {
					const { instantiation, links, notifications } = explicitRemoteFixture(source);
					const folders = Array.from({ length: count }, (_, i) => URI.file(`/local${i}`));
					const chat = upcastPartial<IChat>({ resource: URI.parse('test:/ambient'), status: constObservable(sent ? SessionStatus.Completed : SessionStatus.Untitled), workspace: constObservable(upcastPartial<ISessionWorkspace>({ folders: folders.map(root => ({ root, workingDirectory: root, name: 'local', description: undefined })) })) });
					const session = upcastPartial<ISession>({ resource: URI.parse('test:/parent'), providerId: 'test', status: constObservable(SessionStatus.Completed) });
					await instantiation.invokeFunction(accessor => new OpenInVSCodeAction().run(accessor, { session, chat }));
					assert.strictEqual(links.length, 1);
					const link = links[0];
					const parsed = parseExternalEditorWindowLinkUri(link, 'signal-c');
					assert.deepStrictEqual({ authority: link.authority, path: link.path, remote: parsed?.remoteAuthority, folders: parsed?.folders?.map(folder => folder.toString()), chat: new URLSearchParams(link.query).get('session'), notifications }, {
						authority: count === 1 ? 'file' : 'editor-window', path: count === 1 ? folders[0].path : '/open', remote: count === 0 ? source : undefined, folders: count === 2 ? folders.map(folder => folder.toString()) : undefined, chat: sent && (!source || count === 0) ? chat.resource.toString() : null, notifications: sent && source && count > 0 ? ['This chat stays in Agents because the selected editor workspace uses a different connection.'] : [],
					});
				});
			}
		}
	}
	test('keeps a remote ambient connection when there is no selected session', async () => {
		const { instantiation, links } = explicitRemoteFixture('ssh-remote+source');
		instantiation.stub(ISessionsService, upcastPartial<ITestSessionsService>({ activeSession: constObservable(undefined) }));
		stubSessionHostsFromSessionsService(instantiation);
		await instantiation.invokeFunction(accessor => new OpenInVSCodeAction().run(accessor));
		assert.strictEqual(parseExternalEditorWindowLinkUri(links[0], 'signal-c')?.remoteAuthority, 'ssh-remote+source');
	});
	test('keeps the ambient remote on an unsupported single-folder fallback without transferring a chat', async () => {
		const { instantiation, links, notifications } = explicitRemoteFixture('ssh-remote+source');
		const folder = toAgentHostUri(URI.file('/unsupported'), agentHostAuthority('ws:other'));
		const chat = upcastPartial<IChat>({ resource: URI.parse('test:/chat'), status: constObservable(SessionStatus.Completed), workspace: constObservable(upcastPartial<ISessionWorkspace>({ folders: [{ root: folder, workingDirectory: folder, name: 'other', description: undefined }] })) });
		const session = upcastPartial<ISession>({ resource: URI.parse('test:/parent'), providerId: 'test', status: constObservable(SessionStatus.Completed) });
		await instantiation.invokeFunction(accessor => new OpenInVSCodeAction().run(accessor, { session, chat }));
		const parsed = parseExternalEditorWindowLinkUri(links[0], 'signal-c');
		assert.deepStrictEqual({ remote: parsed?.remoteAuthority, chat: parsed?.chatSessionToOpen, notifications }, { remote: 'ssh-remote+source', chat: undefined, notifications: [] });
	});
	test('does not reinterpret explicit local folders through a remote provider', async () => {
		const { instantiation, links, notifications } = explicitRemoteFixture();
		const provider: ISessionsProvider = upcastPartial<IAgentHostSessionsProvider>({ id: 'agenthost-test', remoteAddress: 'ssh:source' });
		instantiation.stub(ISessionsProvidersService, upcastPartial<ISessionsProvidersService>({ getProvider: <T extends ISessionsProvider>() => provider as T }));
		instantiation.stub(IRemoteAgentHostService, upcastPartial<IRemoteAgentHostService>({ getEntryByAddress: () => ({ name: 'Source', connection: { type: RemoteAgentHostEntryType.SSH, address: 'ssh:source', hostName: 'source' } }) }));
		const folder = URI.file('/local');
		const chat = upcastPartial<IChat>({ resource: URI.parse('test:/remote'), status: constObservable(SessionStatus.Completed), workspace: constObservable(upcastPartial<ISessionWorkspace>({ folders: [{ root: folder, workingDirectory: folder, name: 'local', description: undefined }] })) });
		const session = upcastPartial<ISession>({ resource: URI.parse('test:/parent'), providerId: 'test', status: constObservable(SessionStatus.Completed) });
		await instantiation.invokeFunction(accessor => new OpenInVSCodeAction().run(accessor, { session, chat }));
		assert.deepStrictEqual({ links: links.map(link => ({ authority: link.authority, path: link.path, chat: new URLSearchParams(link.query).get('session') })), notifications }, { links: [{ authority: 'file', path: '/local', chat: null }], notifications: ['This chat stays in Agents because the selected editor workspace uses a different connection.'] });
	});
	for (const source of [undefined, 'ssh-remote+target', 'ssh-remote+other']) {
		for (const count of [1, 2]) {
			for (const sent of [false, true]) {
				test(`preserves explicit remote roots and chat ownership without a provider resolver (${source ?? 'local'}, ${count}, ${sent ? 'sent' : 'draft'})`, async () => {
					const { instantiation, links, notifications } = explicitRemoteFixture(source);
					const folders = Array.from({ length: count }, (_, i) => URI.from({ scheme: 'vscode-remote', authority: 'ssh-remote+target', path: `/project${i}` }));
					const chat = upcastPartial<IChat>({ resource: URI.parse('test:/chat'), status: constObservable(sent ? SessionStatus.Completed : SessionStatus.Untitled), workspace: constObservable(upcastPartial<ISessionWorkspace>({ folders: folders.map(root => ({ root, workingDirectory: root, name: 'project', description: undefined })) })) });
					const session = upcastPartial<ISession>({ resource: URI.parse('test:/parent'), providerId: 'test', status: constObservable(SessionStatus.Completed) });
					await instantiation.invokeFunction(accessor => new OpenInVSCodeAction().run(accessor, { session, chat }));
					assert.strictEqual(links.length, 1);
					const link = links[0];
					const parsed = parseExternalEditorWindowLinkUri(link, 'signal-c');
					assert.deepStrictEqual({ authority: link.authority, path: link.path, folders: parsed?.folders?.map(folder => folder.toString()), remote: parsed?.remoteAuthority, chat: new URLSearchParams(link.query).get('session'), notifications }, {
						authority: count === 1 ? 'vscode-remote' : 'editor-window', path: count === 1 ? '/ssh-remote+target/project0' : '/open', folders: count === 1 ? undefined : folders.map(folder => folder.toString()), remote: count === 1 ? undefined : 'ssh-remote+target', chat: sent && source === 'ssh-remote+target' ? chat.resource.toString() : null, notifications: sent && source !== 'ssh-remote+target' ? ['This chat stays in Agents because the selected editor workspace uses a different connection.'] : [],
					});
				});
			}
		}
	}
	for (const folders of [
		[URI.parse('vscode-remote://ssh-remote+target/project'), URI.file('/local')],
		[URI.parse('vscode-remote://ssh-remote+target/project'), URI.parse('vscode-remote://ssh-remote+other/other')],
		[URI.from({ scheme: 'vscode-remote', path: '/missing-host' })],
	]) {
		test(`does not infer one remote destination from incompatible roots: ${folders.map(folder => folder.toString()).join(', ')}`, async () => {
			const { instantiation, links, notifications } = explicitRemoteFixture();
			const chat = upcastPartial<IChat>({ resource: URI.parse('test:/chat'), status: constObservable(SessionStatus.Completed), workspace: constObservable(upcastPartial<ISessionWorkspace>({ folders: folders.map(root => ({ root, workingDirectory: root, name: 'project', description: undefined })) })) });
			const session = upcastPartial<ISession>({ resource: URI.parse('test:/parent'), providerId: 'test', status: constObservable(SessionStatus.Completed) });
			await instantiation.invokeFunction(accessor => new OpenInVSCodeAction().run(accessor, { session, chat }));
			assert.deepStrictEqual({ links, notifications }, { links: [], notifications: ['This workspace cannot be opened in the editor because its folder paths are unsupported or its link is too large.'] });
		});
	}

	for (const [status, parentStatus] of [[SessionStatus.Completed, SessionStatus.Completed], [SessionStatus.Untitled, SessionStatus.Untitled], [SessionStatus.Untitled, SessionStatus.Completed]]) {
		for (const { explicit, network, folderless, multiRoot } of [
			...[{ explicit: false, network: false, folderless: false }, { explicit: true, network: false, folderless: false }, { explicit: false, network: true, folderless: false }, { explicit: true, network: true, folderless: false }, { explicit: false, network: false, folderless: true }, { explicit: true, network: false, folderless: true }].map(options => ({ ...options, multiRoot: false })),
			{ explicit: false, network: false, folderless: false, multiRoot: true }, { explicit: true, network: true, folderless: false, multiRoot: true },
		]) {
			test(`opens the ${explicit ? 'clicked' : 'focused'} chat ${multiRoot ? 'multi-root' : folderless ? 'folderless' : network ? 'network' : 'local'} workspace and ${status === SessionStatus.Untitled ? 'keeps the draft in Agents' : 'hands off that chat'} (parent status: ${parentStatus})`, async () => {
				const instantiation = store.add(new TestInstantiationService());
				const mainFolder = URI.file('/main project');
				instantiation.stub(IWorkbenchEnvironmentService, upcastPartial<IWorkbenchEnvironmentService>({ remoteAuthority: undefined }));
				const folder = network
					? URI.from({ scheme: 'file', authority: 'fileserver', path: explicit ? '/shared projects/clicked # & 100%/项目' : '/shared projects/nested # & 100%/项目' })
					: URI.file(explicit ? '/clicked project' : '/nested project');
				const mainChat = upcastPartial<IChat>({ resource: URI.parse('test:/main-chat'), workspace: constObservable(upcastPartial<ISessionWorkspace>({ folders: [{ root: mainFolder, workingDirectory: mainFolder, name: 'main', description: undefined }] })) });
				const folders = folderless ? [] : [folder, ...(multiRoot ? [URI.file('/second project # 100%/项目')] : [])];
				const chat = upcastPartial<IChat>({ resource: URI.from({ scheme: 'test', path: explicit ? '/clicked chat' : '/nested chat', fragment: 'chat&part=2' }), status: constObservable(status), workspace: constObservable(upcastPartial<ISessionWorkspace>({ folders: folders.map(root => ({ root, workingDirectory: root, name: 'nested', description: undefined })) })) });
				const session = upcastPartial<ISession>({ resource: URI.parse('test:/parent'), providerId: 'test', status: constObservable(parentStatus), mainChat: constObservable(mainChat) });
				const focused = upcastPartial<ISessionHost>({ session: constObservable(session), chat: constObservable(explicit ? mainChat : chat) });
				const opened: { scheme: string; authority: string; path: string; folder: string | undefined; folders: string[] | undefined; session: string | null }[] = [];
				instantiation.stub(ISessionHostService, upcastPartial<ISessionHostService>({}));
				instantiation.stub(IFocusedSessionHostService, upcastPartial<IFocusedSessionHostService>({ getFocusedHost: () => focused }));
				instantiation.stub(ISessionsProvidersService, upcastPartial<ISessionsProvidersService>({ getProvider: () => undefined }));
				instantiation.stub(IRemoteAgentHostService, upcastPartial<IRemoteAgentHostService>({ getEntryByAddress: () => undefined }));
				instantiation.stub(IAgentHostConnectionsService, upcastPartial<IAgentHostConnectionsService>({ resolveSessionResourceIdentity: () => undefined }));
				instantiation.stub(IRemoteAgentHostLocationPreferenceService, upcastPartial<IRemoteAgentHostLocationPreferenceService>({ getPreference: () => undefined }));
				instantiation.stub(ISSHRemoteAgentHostService, upcastPartial<ISSHRemoteAgentHostService>({ connections: [] }));
				instantiation.stub(ITunnelAgentHostService, upcastPartial<ITunnelAgentHostService>({ getSelectedServerType: () => undefined }));
				instantiation.stub(ITelemetryService, NullTelemetryService);
				instantiation.stub(IOpenerService, upcastPartial<IOpenerService>({ open: async resource => {
					const uri = URI.parse(resource.toString());
					// The native protocol handler reconstructs the file URI from the link's fsPath.
					opened.push({ scheme: uri.scheme, authority: uri.authority, path: uri.path, folder: uri.authority === 'file' ? URI.file(uri.fsPath).toString() : undefined, folders: parseExternalEditorWindowLinkUri(uri.toString(), 'signal-c')?.folders?.map(folder => folder.toString()), session: new URLSearchParams(uri.query).get('session') });
					return true;
				} }));
				for (const quality of [undefined, 'stable', 'insider', 'exploration'] as const) {
					instantiation.stub(IProductService, upcastPartial<IProductService>({ quality, urlProtocol: 'signal-c' }));
					await instantiation.invokeFunction(accessor => new OpenInVSCodeAction().run(accessor, explicit ? { session, chat } : undefined));
				}
				assert.deepStrictEqual(opened, Array.from({ length: 4 }, () => ({ scheme: 'signal-c', authority: folderless || multiRoot ? 'editor-window' : 'file', path: folderless || multiRoot ? '/open' : network ? `//fileserver${folder.path}` : folder.path, folder: folderless || multiRoot ? undefined : folder.toString(), folders: multiRoot ? folders.map(folder => folder.toString()) : undefined, session: status === SessionStatus.Untitled ? null : chat.resource.toString() })));
				if (multiRoot) {
					opened.length = 0;
					const notifications: string[] = [];
					instantiation.stub(INotificationService, upcastPartial<INotificationService>({ info: message => { notifications.push(String(message)); } }));
					const unsupported = URI.parse('unsupported-root:/second');
					const unsupportedChat = { ...chat, workspace: constObservable(upcastPartial<ISessionWorkspace>({ folders: [folder, unsupported].map(root => ({ root, workingDirectory: root, name: 'root', description: undefined })) })) };
					await instantiation.invokeFunction(accessor => new OpenInVSCodeAction().run(accessor, { session, chat: unsupportedChat }));
					assert.deepStrictEqual({ opened, notifications }, { opened: [], notifications: ['This workspace cannot be opened in the editor because its folder paths are unsupported or its link is too large.'] });
				}
			});
		}
	}

	for (const [connection, expected] of [
		[
			{ type: RemoteAgentHostEntryType.CloudSandbox, address: 'cloudsandbox:environment', environmentId: 'environment' },
			{ authority: '', path: '', query: 'windowId=_blank' },
		],
		[
			{ type: RemoteAgentHostEntryType.SSH, address: 'ssh:host', hostName: 'host' },
			{ authority: 'vscode-remote', path: '/ssh-remote+host/c:/Users/test/project', query: 'windowId=_blank&session=test%3A%2Fsession' },
		],
	] satisfies [IRemoteAgentHostEntry['connection'], { authority: string; path: string; query: string }][]) {
		test(`routes ${connection.type} without opening remote paths locally`, async () => {
			const instantiationService = store.add(new TestInstantiationService());
			const calls: { authority: string; path: string; query: string; openExternal: boolean | undefined }[] = [];
			instantiationService.stub(IWorkbenchEnvironmentService, upcastPartial<IWorkbenchEnvironmentService>({ remoteAuthority: undefined }));
			const folder = toAgentHostUri(URI.from({ scheme: 'file', path: '/c:/Users/test/project' }), agentHostAuthority(connection.address));
			const workspace = upcastPartial<ISessionWorkspace>({ folders: [{ workingDirectory: folder, root: folder, name: 'project', description: undefined }] });
			const session = upcastPartial<ITestSessionWithChat>({
				resource: URI.from({ scheme: 'test', path: '/session' }),
				providerId: 'test',
				status: constObservable(SessionStatus.Completed),
				activeChat: constObservable(upcastPartial<IChat>({ status: constObservable(SessionStatus.Completed), workspace: constObservable(workspace) })),
			});
			const provider: ISessionsProvider = upcastPartial<IAgentHostSessionsProvider>({ id: 'agenthost-test', remoteAddress: connection.address });
			instantiationService.stub(ISessionsService, upcastPartial<ITestSessionsService>({ activeSession: constObservable(session) }));
			stubSessionHostsFromSessionsService(instantiationService);
			instantiationService.stub(ISessionsProvidersService, upcastPartial<ISessionsProvidersService>({
				getProvider: <T extends ISessionsProvider>() => provider as T,
			}));
			instantiationService.stub(IRemoteAgentHostService, upcastPartial<IRemoteAgentHostService>({
				getEntryByAddress: () => ({ name: 'Remote host', connection }),
			}));
			instantiationService.stub(IAgentHostConnectionsService, upcastPartial<IAgentHostConnectionsService>({ resolveSessionResourceIdentity: () => undefined }));
			instantiationService.stub(IRemoteAgentHostLocationPreferenceService, upcastPartial<IRemoteAgentHostLocationPreferenceService>({ getPreference: () => undefined }));
			instantiationService.stub(ISSHRemoteAgentHostService, upcastPartial<ISSHRemoteAgentHostService>({ connections: [] }));
			instantiationService.stub(ITunnelAgentHostService, upcastPartial<ITunnelAgentHostService>({ getSelectedServerType: () => undefined }));
			instantiationService.stub(ITelemetryService, NullTelemetryService);
			instantiationService.stub(IProductService, upcastPartial<IProductService>({ urlProtocol: 'vscode-test' }));
			instantiationService.stub(IOpenerService, new class extends mock<IOpenerService>() {
				override async open(resource: URI | string, options?: { openExternal?: boolean }): Promise<boolean> {
					const uri = typeof resource === 'string' ? URI.parse(resource) : resource;
					calls.push({ authority: uri.authority, path: uri.path, query: uri.query, openExternal: options?.openExternal });
					return true;
				}
			}());

			await instantiationService.invokeFunction(accessor => new OpenInVSCodeAction().run(accessor));

			assert.deepStrictEqual(calls, [{ ...expected, openExternal: true }]);

			calls.length = 0;
			const clickedChat = upcastPartial<IChat>({
				resource: URI.parse('test:/nested-chat'),
				status: constObservable(SessionStatus.Completed),
				workspace: constObservable(workspace),
			});
			await instantiationService.invokeFunction(accessor => new OpenInVSCodeAction().run(accessor, { session, chat: clickedChat }));

			assert.deepStrictEqual(calls, [{
				...expected,
				query: connection.type === RemoteAgentHostEntryType.SSH ? 'windowId=_blank&session=test%3A%2Fnested-chat' : expected.query,
				openExternal: true,
			}]);
		});
	}

	for (const [connection, serverType, preference] of [
		[{ type: RemoteAgentHostEntryType.SSH, address: 'ssh:host', hostName: 'host', sshConfigHost: 'host' }, 'standalone', undefined],
		[{ type: RemoteAgentHostEntryType.SSH, address: 'ssh:host', hostName: 'host', sshConfigHost: 'host' }, 'editor', undefined],
		[{ type: RemoteAgentHostEntryType.SSH, address: 'ssh:host', hostName: 'host', sshConfigHost: 'host' }, 'editor', 'editor'],
		[{ type: RemoteAgentHostEntryType.SSH, address: 'ssh:host', hostName: 'host', sshConfigHost: 'host' }, 'editor', 'dedicated'],
		[{ type: RemoteAgentHostEntryType.WSL, address: 'wsl:Ubuntu', distro: 'Ubuntu' }, undefined, undefined],
		[{ type: RemoteAgentHostEntryType.DevContainer, address: 'devcontainer:host', hostPath: '/source/project' }, undefined, undefined],
	] satisfies [IRemoteAgentHostEntry['connection'], ISSHAgentHostConnection['serverType'], RemoteAgentHostLocationPreference | undefined][]) {
		test(`hands off ${connection.type} sessions when the editor follows the selected host (${serverType ?? 'standalone'}, ${preference ?? 'no preference'})`, async () => {
			const instantiationService = store.add(new TestInstantiationService());
			const entry = { name: 'Remote host', connection };
			instantiationService.stub(IWorkbenchEnvironmentService, upcastPartial<IWorkbenchEnvironmentService>({ remoteAuthority: undefined }));
			const address = getEntryAddress(entry);
			const authority = agentHostAuthority(address);
			const folder = toAgentHostUri(URI.file('/project'), authority);
			const resource = URI.from({ scheme: remoteAgentHostSessionTypeId(authority, 'copilotcli'), path: '/session', fragment: 'chat' });
			const chat = upcastPartial<IChat>({
				resource,
				status: constObservable(SessionStatus.Completed),
				workspace: constObservable(upcastPartial<ISessionWorkspace>({ folders: [{ workingDirectory: folder, root: folder, name: 'project', description: undefined }] })),
			});
			const session = upcastPartial<ITestSessionWithChat>({ resource, providerId: 'test', status: constObservable(SessionStatus.Completed), activeChat: constObservable(chat) });
			const provider: ISessionsProvider = upcastPartial<IAgentHostSessionsProvider>({ id: 'agenthost-test', remoteAddress: address });
			const sessions: (string | null)[] = [];
			const links: URI[] = [];
			const notifications: string[] = [];
			instantiationService.stub(ISessionsProvidersService, upcastPartial<ISessionsProvidersService>({ getProvider: <T extends ISessionsProvider>() => provider as T }));
			instantiationService.stub(IRemoteAgentHostService, upcastPartial<IRemoteAgentHostService>({ getEntryByAddress: () => entry }));
			instantiationService.stub(IAgentHostConnectionsService, upcastPartial<IAgentHostConnectionsService>({
				resolveSessionResourceIdentity: () => ({ connectionAuthority: authority, backendSession: URI.parse('copilotcli:/session') }),
			}));
			instantiationService.stub(IRemoteAgentHostLocationPreferenceService, upcastPartial<IRemoteAgentHostLocationPreferenceService>({ getPreference: key => key === address ? preference : undefined }));
			instantiationService.stub(ISSHRemoteAgentHostService, upcastPartial<ISSHRemoteAgentHostService>({
				connections: [upcastPartial<ISSHAgentHostConnection>({ localAddress: address, serverType })],
			}));
			instantiationService.stub(ITunnelAgentHostService, upcastPartial<ITunnelAgentHostService>({ getSelectedServerType: () => undefined }));
			instantiationService.stub(INotificationService, upcastPartial<INotificationService>({ info: message => { notifications.push(String(message)); } }));
			instantiationService.stub(ITelemetryService, NullTelemetryService);
			instantiationService.stub(IProductService, upcastPartial<IProductService>({ urlProtocol: 'vscode-test' }));
			instantiationService.stub(IOpenerService, upcastPartial<IOpenerService>({
				open: async resource => {
					const link = typeof resource === 'string' ? URI.parse(resource) : resource;
					links.push(link);
					sessions.push(new URLSearchParams(link.query).get('session'));
					return true;
				},
			}));

			await instantiationService.invokeFunction(accessor => new OpenInVSCodeAction().run(accessor, { session, chat }));

			const sharesAgentHost = serverType !== 'editor' || preference === 'editor';
			assert.deepStrictEqual({ sessions, notifications }, {
				sessions: [sharesAgentHost ? 'agent-host-copilotcli:/session#chat' : null],
				notifications: !sharesAgentHost ? ['The session\'s folder was opened in a new editor window, but the session can\'t be opened there. This window is connected to a different agent host on the remote machine than the one its editor windows use.'] : [],
			});
			sessions.length = 0; notifications.length = 0; links.length = 0;
			await instantiationService.invokeFunction(accessor => new OpenInVSCodeAction().run(accessor, { session, chat: { ...chat, workspace: constObservable(undefined) } }));
			const parsed = parseExternalEditorWindowLinkUri(links[0], 'vscode-test');
			assert.ok(parsed);
			assert.strictEqual(parsed.remoteAuthority, connection.type === RemoteAgentHostEntryType.SSH ? 'ssh-remote+host' : connection.type === RemoteAgentHostEntryType.WSL ? 'wsl+Ubuntu' : 'dev-container+2f736f757263652f70726f6a656374');
			assert.deepStrictEqual({ sessions, notifications }, {
				sessions: [sharesAgentHost ? 'agent-host-copilotcli:/session#chat' : null],
				notifications: sharesAgentHost ? [] : ['This chat can\'t be opened in the editor because it uses a different agent host on the remote machine.'],
			});
			const uncFolder = toAgentHostUri(URI.from({ scheme: 'file', authority: 'fileserver', path: '/share/project # 100%' }), authority);
			sessions.length = 0; notifications.length = 0; links.length = 0;
			await instantiationService.invokeFunction(accessor => new OpenInVSCodeAction().run(accessor, { session, chat: { ...chat, workspace: constObservable(upcastPartial<ISessionWorkspace>({ folders: [{ root: uncFolder, workingDirectory: uncFolder, name: 'share', description: undefined }] })) } }));
			assert.deepStrictEqual({ path: links[0].path, sessions }, { path: `/${parsed.remoteAuthority}//fileserver/share/project # 100%`, sessions: [sharesAgentHost ? 'agent-host-copilotcli:/session#chat' : null] });
			sessions.length = 0; notifications.length = 0; links.length = 0;
			const roots = [folder, toAgentHostUri(URI.file('/second project # 100%/项目'), authority), uncFolder];
			await instantiationService.invokeFunction(accessor => new OpenInVSCodeAction().run(accessor, { session, chat: { ...chat, workspace: constObservable(upcastPartial<ISessionWorkspace>({ folders: roots.map(root => ({ root, workingDirectory: root, name: 'root', description: undefined })) })) } }));
			const multiRoot = parseExternalEditorWindowLinkUri(links[0], 'vscode-test');
			assert.deepStrictEqual({ host: multiRoot?.remoteAuthority, roots: multiRoot?.folders?.map(folder => folder.toString()), sessions }, { host: parsed.remoteAuthority, roots: ['/project', '/second project # 100%/项目', '//fileserver/share/project # 100%'].map(path => URI.from({ scheme: 'vscode-remote', authority: parsed.remoteAuthority, path }).toString()), sessions: [sharesAgentHost ? 'agent-host-copilotcli:/session#chat' : null] });
			for (const otherHostRoot of [URI.parse('vscode-remote://ssh-remote+other/second'), toAgentHostUri(URI.file('/second'), agentHostAuthority('ssh:other'))]) {
				for (const roots of [[folder, otherHostRoot], [otherHostRoot, otherHostRoot.with({ path: '/third' })]]) {
					sessions.length = 0; notifications.length = 0; links.length = 0;
					await instantiationService.invokeFunction(accessor => new OpenInVSCodeAction().run(accessor, { session, chat: { ...chat, workspace: constObservable(upcastPartial<ISessionWorkspace>({ folders: roots.map(root => ({ root, workingDirectory: root, name: 'root', description: undefined })) })) } }));
					assert.deepStrictEqual({ links, notifications }, { links: [], notifications: ['This workspace cannot be opened in the editor because its folder paths are unsupported or its link is too large.'] });
				}
			}
			sessions.length = 0; notifications.length = 0; links.length = 0;
			const explicitOther = URI.parse('vscode-remote://ssh-remote+other/second');
			await instantiationService.invokeFunction(accessor => new OpenInVSCodeAction().run(accessor, { session, chat: { ...chat, workspace: constObservable(upcastPartial<ISessionWorkspace>({ folders: [{ root: explicitOther, workingDirectory: explicitOther, name: 'other', description: undefined }] })) } }));
			assert.deepStrictEqual({ links: links.map(link => ({ authority: link.authority, path: link.path, query: link.query })), sessions, notifications }, { links: [{ authority: '', path: '', query: 'windowId=_blank' }], sessions: [null], notifications: [] });
		});
	}
});

suite('Open in Editor in Agents windows', () => {
	const store = ensureNoDisposablesAreLeakedInTestSuite();
	for (const os of [OperatingSystem.Windows, OperatingSystem.Macintosh, OperatingSystem.Linux]) {
		test(`resolves the workspace mode chord only in a ready desktop Agents window (${os})`, () => {
			store.add(registerAction2(OpenInVSCodeAction));
			const rules = KeybindingsRegistry.getDefaultKeybindingsForOS(os).filter(rule => rule.command === OpenInVSCodeAction.ID);
			const resolver = new KeybindingResolver(rules.map(rule => new ResolvedKeybindingItem(
				rule.keybinding ? USLayoutResolvedKeybinding.resolveKeybinding(rule.keybinding, os)[0] : undefined, rule.command, rule.commandArgs, rule.when ?? undefined, true, null, false,
			)), [], () => { });
			const context = new Context(0, null);
			context.setValue(IsSessionsWindowContext.key, true);
			const first = os === OperatingSystem.Macintosh ? 'meta+K' : 'ctrl+K';
			const second = os === OperatingSystem.Macintosh ? 'shift+meta+A' : 'ctrl+shift+A';
			assert.strictEqual(resolver.resolve(context, [], first).kind, ResultKind.MoreChordsNeeded);
			assert.deepStrictEqual(resolver.resolve(context, [first], second), {
				kind: ResultKind.KbFound, commandId: OpenInVSCodeAction.ID, commandArgs: undefined, isBubble: false,
			});
			for (const key of [IsPhoneLayoutContext.key, SessionsWelcomeVisibleContext.key, IsAuxiliaryWindowContext.key]) {
				context.setValue(key, true);
				assert.strictEqual(resolver.resolve(context, [], first).kind, ResultKind.NoMatchingKb);
				context.setValue(key, false);
			}
			context.setValue(IsSessionsWindowContext.key, false);
			assert.strictEqual(resolver.resolve(context, [], first).kind, ResultKind.NoMatchingKb);
		});
	}

	function attach(container: HTMLElement): HTMLElement {
		const root = append(mainWindow.document.body, $('.monaco-workbench'));
		root.style.cssText = 'position: fixed; top: 0; left: 0';
		store.add(toDisposable(() => root.remove()));
		append(root, container);
		return root;
	}

	async function waitForObservation(element: HTMLElement): Promise<boolean> {
		await Promise.resolve();
		return new Promise(resolve => {
			const observer = new mainWindow.IntersectionObserver(entries => {
				observer.disconnect();
				resolve(entries[0].isIntersecting);
			});
			observer.observe(element);
		});
	}

	test('uses the editor icon for the product quality', () => {
		const keybindings = upcastPartial<IKeybindingService>({ appendKeybinding: label => `${label} (shortcut)` });
		const configuration = new TestConfigurationService({ [AGENTS_WINDOW_TITLE_BAR_ICONS_SETTING]: true });
		store.add(configuration.onDidChangeConfigurationEmitter);
		const results = ['stable', 'insider', 'exploration', undefined].map(quality => {
			const action = store.add(new Action('test', 'Open in Editor'));
			const widget = store.add(new OpenInVSCodeTitleBarWidget(action, undefined, 'test', upcastPartial<IProductService>({ quality }), NullHoverService, keybindings, configuration, NullTelemetryService));
			const container = document.createElement('div');
			widget.render(container);
			const icon = container.querySelector('.open-in-vscode-titlebar-widget-icon');
			return {
				icon: icon?.className,
				hidden: icon?.getAttribute('aria-hidden'),
				label: container.getAttribute('aria-label'),
			};
		});
		assert.deepStrictEqual({
			results,
			stableDefault: getIconRegistry().getIcon('open-in-vscode')?.defaults,
			insidersDefault: getIconRegistry().getIcon('open-in-vscode-insiders')?.defaults,
		}, {
			results: [
				{ icon: 'open-in-vscode-titlebar-widget-icon codicon codicon-open-in-vscode', hidden: 'true', label: 'Open in VS Code Editor Window (shortcut)' },
				{ icon: 'open-in-vscode-titlebar-widget-icon codicon codicon-open-in-vscode-insiders', hidden: 'true', label: 'Open in VS Code Editor Window (shortcut)' },
				{ icon: 'open-in-vscode-titlebar-widget-icon codicon codicon-open-in-vscode', hidden: 'true', label: 'Open in VS Code Editor Window (shortcut)' },
				{ icon: 'open-in-vscode-titlebar-widget-icon codicon codicon-open-in-vscode', hidden: 'true', label: 'Open in VS Code Editor Window (shortcut)' },
			],
			stableDefault: Codicon.vscode,
			insidersDefault: Codicon.vscodeInsiders,
		});
	});

	for (const useCodicons of [false, true]) {
		test(`uses the ${useCodicons ? 'codicon' : 'branded'} editor icon in the ${useCodicons ? 'treatment' : 'control'} arm and logs its trigger`, async () => {
			const setting = AGENTS_WINDOW_TITLE_BAR_ICONS_SETTING;
			const configuration = new class extends TestConfigurationService {
				override inspect<T>(key: string) {
					const inspected = super.inspect<T>(key);
					return key === setting ? { ...inspected, userValue: undefined, userLocalValue: undefined } : inspected;
				}
			}({ [setting]: useCodicons });
			store.add(configuration.onDidChangeConfigurationEmitter);
			const telemetry = new TestExperimentTriggerTelemetryService();
			const keybindings = upcastPartial<IKeybindingService>({ appendKeybinding: label => `${label} (shortcut)` });
			const action = store.add(new Action('test', 'Open in Editor'));
			const widget = store.add(new OpenInVSCodeTitleBarWidget(action, undefined, 'test', upcastPartial<IProductService>({ quality: 'stable' }), NullHoverService, keybindings, configuration, telemetry));
			const container = document.createElement('div');
			widget.render(container);
			const icon = container.querySelector<HTMLElement>('.open-in-vscode-titlebar-widget-icon')!;
			icon.style.cssText = 'width: 16px; height: 16px';
			const triggered = Event.toPromise(telemetry.onDidTrigger);
			const root = attach(container);
			root.style.display = 'none';
			assert.strictEqual(await waitForObservation(container), false);
			assert.strictEqual(telemetry.triggers.length, 0);
			root.style.display = '';
			// Revealing an unpainted element does not schedule an intersection update in headless WebKit.
			await Promise.all([triggered, waitForObservation(container)]);
			assert.deepStrictEqual({
				hasCodicon: icon?.classList.contains('codicon'),
				triggerLogged: telemetry.triggers.includes(`config.${setting}`),
			}, { hasCodicon: useCodicons, triggerLogged: true });

			await configuration.setUserConfiguration(setting, !useCodicons);
			configuration.onDidChangeConfigurationEmitter.fire(new class extends mock<IConfigurationChangeEvent>() {
				override affectsConfiguration(section: string): boolean {
					return section === setting;
				}
			}());
			assert.strictEqual(icon?.classList.contains('codicon'), !useCodicons);
		});

		test(`logs the editor icon trigger when a user resets to the ${useCodicons ? 'codicon' : 'branded'} default`, async () => {
			const setting = AGENTS_WINDOW_TITLE_BAR_ICONS_SETTING;
			const configuration = new class extends TestConfigurationService {
				usingDefault = false;
				override inspect<T>(key: string) {
					const inspected = super.inspect<T>(key);
					return key === setting && this.usingDefault ? { ...inspected, userValue: undefined, userLocalValue: undefined } : inspected;
				}
			}({ [setting]: !useCodicons });
			store.add(configuration.onDidChangeConfigurationEmitter);
			const telemetry = new TestExperimentTriggerTelemetryService();
			const keybindings = upcastPartial<IKeybindingService>({ appendKeybinding: label => label });
			const action = store.add(new Action('test', 'Open in Editor'));
			const widget = store.add(new OpenInVSCodeTitleBarWidget(action, undefined, 'test', upcastPartial<IProductService>({ quality: 'stable' }), NullHoverService, keybindings, configuration, telemetry));
			const change = new class extends mock<IConfigurationChangeEvent>() {
				override affectsConfiguration(section: string): boolean {
					return section === setting;
				}
			}();
			configuration.usingDefault = true;
			configuration.onDidChangeConfigurationEmitter.fire(change);
			configuration.usingDefault = false;
			const container = document.createElement('div');
			widget.render(container);
			container.querySelector<HTMLElement>('.open-in-vscode-titlebar-widget-icon')!.style.cssText = 'width: 16px; height: 16px';
			attach(container);
			assert.strictEqual(await waitForObservation(container), true);
			const beforeReset = [...telemetry.triggers];

			await configuration.setUserConfiguration(setting, useCodicons);
			configuration.usingDefault = true;
			const triggered = Event.toPromise(telemetry.onDidTrigger);
			configuration.onDidChangeConfigurationEmitter.fire(change);
			configuration.onDidChangeConfigurationEmitter.fire(change);
			await triggered;
			assert.deepStrictEqual({
				beforeReset,
				hasCodicon: container.querySelector('.open-in-vscode-titlebar-widget-icon')?.classList.contains('codicon'),
				triggers: telemetry.triggers,
			}, { beforeReset: [], hasCodicon: useCodicons, triggers: [`config.${setting}`] });
		});
	}

	test('editor icon uses the title bar foreground in active and inactive windows', () => {
		const root = append(mainWindow.document.body, $('.monaco-workbench.agent-sessions-workbench'));
		root.style.cssText = '--vscode-titleBar-activeForeground: #ffffff; --vscode-titleBar-inactiveForeground: #aabbcc; --vscode-icon-foreground: #000000';
		store.add(toDisposable(() => root.remove()));
		const titlebar = append(root, $('.part.titlebar'));
		const titlebarContainer = append(titlebar, $('.sessions-titlebar-container'));
		const center = append(titlebarContainer, $('.titlebar-center'));
		const actions = append(center, $('.titlebar-actions-container'));
		const actionBar = append(actions, $('.monaco-action-bar'));
		const container = append(actionBar, $('li.action-item'));
		const configuration = new TestConfigurationService({ [AGENTS_WINDOW_TITLE_BAR_ICONS_SETTING]: true });
		store.add(configuration.onDidChangeConfigurationEmitter);
		const action = store.add(new Action('test', 'Open in Editor'));
		const keybindings = upcastPartial<IKeybindingService>({ appendKeybinding: label => label });
		const widget = store.add(new OpenInVSCodeTitleBarWidget(action, undefined, 'test', upcastPartial<IProductService>({ quality: 'stable' }), NullHoverService, keybindings, configuration, NullTelemetryService));
		widget.render(container);
		const icon = container.querySelector<HTMLElement>('.open-in-vscode-titlebar-widget-icon')!;
		const colors = ['vs', 'vs-dark', 'hc-black', 'hc-light'].map(theme => {
			root.classList.add(theme);
			const active = mainWindow.getComputedStyle(icon).color;
			container.classList.add('disabled');
			const disabled = mainWindow.getComputedStyle(icon).color;
			container.classList.remove('disabled');
			titlebar.classList.add('inactive');
			const inactive = mainWindow.getComputedStyle(icon).color;
			titlebar.classList.remove('inactive');
			root.classList.remove(theme);
			return { active, disabled, inactive };
		});
		const expected = { active: 'rgb(255, 255, 255)', disabled: 'rgb(255, 255, 255)', inactive: 'rgb(170, 187, 204)' };
		assert.deepStrictEqual(colors, [expected, expected, expected, expected]);
	});

	test('shows in the main window and in chat windows, as the main window always showed it, and not in editor windows', () => {
		const menus = new OpenInVSCodeAction().desc.menu;
		const when = (Array.isArray(menus) ? menus : [menus]).find(menu => menu?.id === AgentsWindowMenus.TitleBarLeftLayout)?.when;
		const contextKeyService = store.add(new ContextKeyService(new TestConfigurationService()));
		const shows = (auxiliary: boolean, layoutMode: string | undefined) => {
			const scoped = store.add(contextKeyService.createScoped(document.createElement('div')));
			IsAuxiliaryWindowContext.bindTo(scoped).set(auxiliary);
			LayoutModeContext.bindTo(scoped).set(layoutMode);
			return scoped.contextMatchesRules(when);
		};

		assert.deepStrictEqual({ main: shows(false, undefined), chatWindow: shows(true, 'sideBySide'), editorWindow: shows(true, undefined) }, { main: true, chatWindow: true, editorWindow: false });
	});
});
