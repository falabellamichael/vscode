/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import assert from 'assert';
import { VSBuffer } from '../../../../../base/common/buffer.js';
import { CancellationToken } from '../../../../../base/common/cancellation.js';
import { CancellationError } from '../../../../../base/common/errors.js';
import { Emitter, Event } from '../../../../../base/common/event.js';
import { MarkdownString } from '../../../../../base/common/htmlContent.js';
import { isDisposable } from '../../../../../base/common/lifecycle.js';
import { constObservable, observableValue } from '../../../../../base/common/observable.js';
import { isNative } from '../../../../../base/common/platform.js';
import { extUri } from '../../../../../base/common/resources.js';
import { URI } from '../../../../../base/common/uri.js';
import { mock, upcastDeepPartial, upcastPartial } from '../../../../../base/test/common/mock.js';
import { ensureNoDisposablesAreLeakedInTestSuite } from '../../../../../base/test/common/utils.js';
import { IRange, Range } from '../../../../../editor/common/core/range.js';
import { ICommandService } from '../../../../../platform/commands/common/commands.js';
import { IDialogService, IFileDialogService, IPrompt } from '../../../../../platform/dialogs/common/dialogs.js';
import { ConfigurationTarget, IConfigurationService } from '../../../../../platform/configuration/common/configuration.js';
import { TestConfigurationService } from '../../../../../platform/configuration/test/common/testConfigurationService.js';
import { ContextKeyValue } from '../../../../../platform/contextkey/common/contextkey.js';
import { FileOperationError, FileOperationResult, IFileService, IFileStatWithMetadata } from '../../../../../platform/files/common/files.js';
import { TestInstantiationService } from '../../../../../platform/instantiation/test/common/instantiationServiceMock.js';
import { ILabelService } from '../../../../../platform/label/common/label.js';
import { IAllowedMcpServersService, IGalleryMcpServer, IInstallableMcpServer } from '../../../../../platform/mcp/common/mcpManagement.js';
import { IMcpResourceScannerService } from '../../../../../platform/mcp/common/mcpResourceScannerService.js';
import { McpResourceFormat } from '../../../../../platform/mcp/common/mcpWorkspaceConfiguration.js';
import { IMcpServerConfiguration, McpServerType, McpServerVariableType } from '../../../../../platform/mcp/common/mcpPlatformTypes.js';
import { INativeHostService } from '../../../../../platform/native/common/native.js';
import { INotificationService } from '../../../../../platform/notification/common/notification.js';
import { IOpenerService } from '../../../../../platform/opener/common/opener.js';
import { IProgressService } from '../../../../../platform/progress/common/progress.js';
import { IProductService } from '../../../../../platform/product/common/productService.js';
import { IInputOptions, IPickOptions, IQuickInputService, IQuickPick, IQuickPickDidAcceptEvent, IQuickPickItem, QuickPickInput } from '../../../../../platform/quickinput/common/quickInput.js';
import { ISecretStorageService } from '../../../../../platform/secrets/common/secrets.js';
import { StorageScope } from '../../../../../platform/storage/common/storage.js';
import { ITelemetryService } from '../../../../../platform/telemetry/common/telemetry.js';
import { NullTelemetryService } from '../../../../../platform/telemetry/common/telemetryUtils.js';
import { IUriIdentityService } from '../../../../../platform/uriIdentity/common/uriIdentity.js';
import { isWorkspaceFolder, IWorkspaceContextService, toWorkspaceFolder, WorkbenchState } from '../../../../../platform/workspace/common/workspace.js';
import { Workspace } from '../../../../../platform/workspace/test/common/testWorkspace.js';
import { ActiveEditorContext, ResourceContextKey } from '../../../../common/contextkeys.js';
import { IAuthenticationService } from '../../../../services/authentication/common/authentication.js';
import { IAuthenticationQueryService } from '../../../../services/authentication/common/authenticationQuery.js';
import { IEditorService } from '../../../../services/editor/common/editorService.js';
import { IWorkbenchEnvironmentService } from '../../../../services/environment/common/environmentService.js';
import { IRemoteUserDataProfilesService } from '../../../../services/userDataProfile/common/remoteUserDataProfiles.js';
import { IUserDataProfileService } from '../../../../services/userDataProfile/common/userDataProfile.js';
import { IWorkbenchLocalMcpServer, IWorkbenchMcpManagementService, IWorkbencMcpServerInstallOptions, LocalMcpServerScope, WorkspaceMcpConfigKind } from '../../../../services/mcp/common/mcpWorkbenchManagementService.js';
import { IMcpWorkspaceInstallTargetService, McpWorkspaceInstallTargetService } from '../../../../services/mcp/common/mcpWorkspaceInstallTargetService.js';
import { IAgentHostCustomizationService } from '../../../chat/browser/agentSessions/agentHost/agentHostCustomizationService.js';
import { IChatWidget, IChatWidgetService } from '../../../chat/browser/chat.js';
import { ICustomizationHarnessService, type IHarnessDescriptor } from '../../../chat/common/customizationHarnessService.js';
import { ChatContextKeys } from '../../../chat/common/actions/chatContextKeys.js';
import { AICustomizationManagementCommands, AICustomizationManagementSection } from '../../../chat/common/aiCustomizationWorkspaceService.js';
import { SessionType } from '../../../chat/common/chatSessionsService.js';
import { ContributionEnablementState } from '../../../chat/common/enablement.js';
import { TEXT_FILE_EDITOR_ID } from '../../../files/common/files.js';
import { AddConfigurationAction, ConnectAppAction, McpServerOptionsCommand, OpenRemoteUserMcpResourceCommand, OpenUserMcpResourceCommand, OpenWorkspaceFolderMcpResourceCommand } from '../../browser/mcpCommands.js';
import { findSignalCAppConnection, matchesSignalCAppConnection, microsoft365WorkIq, signalCAppConnections } from '../../common/signalCAppConnections.js';
import { McpConfigurationDestination } from '../../browser/mcpConfigurationDestination.js';
import { getContextMenuActions, InstallAction, InstallInRemoteAction, InstallInWorkspaceAction, ShowServerJsonConfigurationAction } from '../../browser/mcpServerActions.js';
import { mcpWorkspaceRootConfig } from '../../common/mcpConfiguration.js';
import { McpCommandIds } from '../../common/mcpCommandIds.js';
import { getMcpGenerationSchema, McpGeneratedConfiguration } from '../../common/mcpConfigurationGeneration.js';
import { IMcpCopilotGlobalConfigurationService } from '../../common/mcpCopilotGlobalConfigurationService.js';
import { IMcpRegistry } from '../../common/mcpRegistryTypes.js';
import { IMcpSamplingService, IMcpServer, IMcpService, IMcpWorkbenchService, IWorkbenchMcpServer, McpCollectionDefinition, McpConnectionState, mcpOAuthClientSecretStorageKey, McpServerDefinition, McpServerInstallState, McpServerLaunch, McpServerTransportType } from '../../common/mcpTypes.js';
import { startServerByFilter } from '../../common/mcpTypesUtils.js';

class TestQuickInputService extends mock<IQuickInputService>() {
	readonly selections: (string | undefined)[] = [];
	readonly inputs: (string | undefined)[] = [];
	readonly pickLabels: string[][] = [];
	readonly questionOrder: string[] = [];
	readonly prompts: (string | undefined)[] = [];
	readonly inputOptions: (IInputOptions | undefined)[] = [];
	readonly pickOptions: { placeholder?: string; descriptions: (string | undefined)[]; details: (string | undefined)[] }[] = [];

	override pick<T extends IQuickPickItem>(picks: Promise<QuickPickInput<T>[]> | QuickPickInput<T>[], options?: IPickOptions<T> & { canPickMany: true }, token?: CancellationToken): Promise<T[] | undefined>;
	override pick<T extends IQuickPickItem>(picks: Promise<QuickPickInput<T>[]> | QuickPickInput<T>[], options?: IPickOptions<T> & { canPickMany: false }, token?: CancellationToken): Promise<T | undefined>;
	override pick<T extends IQuickPickItem>(picks: Promise<QuickPickInput<T>[]> | QuickPickInput<T>[], options?: Omit<IPickOptions<T>, 'canPickMany'>, token?: CancellationToken): Promise<T | undefined>;
	override async pick<T extends IQuickPickItem>(picks: Promise<QuickPickInput<T>[]> | QuickPickInput<T>[], options?: IPickOptions<T>): Promise<T | T[] | undefined> {
		this.questionOrder.push('pick');
		const items = (await picks).filter((item): item is T => item.type !== 'separator');
		this.prompts.push(options?.placeHolder);
		this.pickLabels.push(items.map(item => item.label));
		this.pickOptions.push({ placeholder: options?.placeHolder, descriptions: items.map(item => item.description), details: items.map(item => item.detail) });
		assert.ok(this.selections.length, `Unexpected picker: ${items.map(item => item.label).join(', ')}`);
		const label = this.selections.shift();
		if (label === undefined) {
			return undefined;
		}
		const selected = items.find(item => item.label === label);
		assert.ok(selected, `Missing choice: ${label}`);
		return options?.canPickMany ? [selected] : selected;
	}

	override async input(options?: IInputOptions): Promise<string | undefined> {
		this.inputOptions.push(options);
		this.prompts.push(options?.title);
		this.questionOrder.push('input');
		assert.ok(this.inputs.length, 'Unexpected input');
		const value = this.inputs.shift();
		return value === '$suggested' ? options?.value : value;
	}
}

suite('MCP configuration entry points', () => {
	const store = ensureNoDisposablesAreLeakedInTestSuite();
	const rootFile = '.mcp.json';
	const legacyFile = '.vscode/mcp.json';
	const installable: IInstallableMcpServer = { name: 'same-name', config: { type: McpServerType.LOCAL, command: 'node', args: ['server.js'] } };
	const agentHostSession = URI.from({ scheme: SessionType.AgentHostCopilot, path: '/test-session' });

	function chatWidget(sessionResource?: URI): IChatWidget {
		return upcastDeepPartial<IChatWidget>({ viewModel: sessionResource ? { sessionResource } : undefined });
	}

	function setup(enabled: boolean, existing: string[] = [], multiRoot = false, withAgentHostSession = false) {
		const instantiation = store.add(new TestInstantiationService());
		const quickInput = new TestQuickInputService();
		const folder = toWorkspaceFolder(URI.file('/project'));
		const secondFolder = toWorkspaceFolder(URI.file('/second'));
		const workspace = new Workspace('mcp-test', multiRoot ? [folder, secondFolder] : [folder], multiRoot ? URI.file('/project.code-workspace') : null);
		const opened: URI[] = [];
		const selections: (IRange | undefined)[] = [];
		const started: string[] = [];
		const secrets = new Map<string, string>();
		const keysAtStart: string[][] = [];
		const appDialogs: IPrompt<unknown>[] = [];
		const chosenFiles: (URI[] | undefined)[] = [];
		const errors: Parameters<INotificationService['error']>[0][] = [];
		const existenceChecks: string[] = [];
		const installs: { server: IInstallableMcpServer; options: IWorkbencMcpServerInstallOptions | undefined }[] = [];
		const agentHostAdds: { session: URI; name: string; config: IMcpServerConfiguration }[] = [];
		const agentHostGlobalAdds: { name: string; config: IMcpServerConfiguration }[] = [];
		const servers = observableValue<readonly IMcpServer[]>('servers', []);
		const collections = observableValue<readonly McpCollectionDefinition[]>('collections', []);

		instantiation.stub(IConfigurationService, new TestConfigurationService({ [mcpWorkspaceRootConfig]: enabled }));
		instantiation.stub(IProductService, { urlProtocol: 'signal-c' });
		instantiation.stub(IQuickInputService, quickInput);
		instantiation.stub(ISecretStorageService, { get: async key => secrets.get(key), set: async (key, value) => { secrets.set(key, value); } });
		instantiation.stub(IFileDialogService, { showOpenDialog: async () => { assert.ok(chosenFiles.length, 'Unexpected file picker'); return chosenFiles.shift(); } });
		instantiation.stub(IDialogService, {
			prompt: async prompt => {
				appDialogs.push(prompt);
				const choice = await quickInput.pick((prompt.buttons ?? []).map(button => ({ label: button.label, button })), { canPickMany: false });
				return { result: choice ? await choice.button.run({}) : undefined };
			},
		});
		instantiation.stub(IWorkspaceContextService, {
			getWorkspace: () => workspace,
			getWorkspaceFolder: resource => workspace.getFolder(resource),
			getWorkbenchState: () => multiRoot ? WorkbenchState.WORKSPACE : WorkbenchState.FOLDER,
		});
		instantiation.stub(IMcpWorkspaceInstallTargetService, instantiation.createInstance(McpWorkspaceInstallTargetService));
		instantiation.stub(IUriIdentityService, { extUri });
		instantiation.stub(IFileService, {
			resolve: async resource => {
				existenceChecks.push(resource.path);
				if (existing.some(file => resource.path === folder.toResource(file).path)) {
					return upcastPartial<IFileStatWithMetadata>({ resource });
				}
				throw new FileOperationError('Not found', FileOperationResult.FILE_NOT_FOUND);
			},
			createFile: async () => { throw new Error('The destination UI must not create files'); },
			writeFile: async () => { throw new Error('The destination UI must not write files'); },
		});
		instantiation.stub(IEditorService, {});
		instantiation.stub(IEditorService, 'openEditor', async (editor: { resource: URI; options?: { selection?: IRange } }) => {
			opened.push(editor.resource);
			selections.push(editor.options?.selection);
			return undefined;
		});
		instantiation.stub(ICommandService, {});
		instantiation.stub(ICommandService, 'executeCommand', async () => false);
		instantiation.stub(IWorkbenchEnvironmentService, {});
		instantiation.stub(IOpenerService, {});
		instantiation.stub(INotificationService, { error: error => errors.push(error) });
		instantiation.stub(ILabelService, {});
		instantiation.stub(IAgentHostCustomizationService, { addMcpServer: (session, name, config) => agentHostAdds.push({ session, name, config }) });
		instantiation.stub(ICustomizationHarnessService, {
			getActiveDescriptor: () => upcastDeepPartial<IHarnessDescriptor>({
				mcpConfigurationProvider: {
					addGlobal: async (name: string, config: IMcpServerConfiguration) => { agentHostGlobalAdds.push({ name, config }); },
				},
			}),
		});
		const lastFocusedWidget = withAgentHostSession ? chatWidget(agentHostSession) : undefined;
		instantiation.stub(IChatWidgetService, { lastFocusedWidget, getAllWidgets: () => lastFocusedWidget ? [lastFocusedWidget] : [] });
		instantiation.stub(ITelemetryService, NullTelemetryService);
		instantiation.stub(IMcpService, { servers });
		instantiation.stub(IMcpRegistry, { collections });
		instantiation.stub(IMcpCopilotGlobalConfigurationService, { getConfigurationResource: async () => undefined });
		instantiation.stub(IMcpResourceScannerService, {});
		instantiation.stub(IAllowedMcpServersService, { isAllowed: () => true });
		instantiation.stub(IUserDataProfileService, upcastDeepPartial<IUserDataProfileService>({ currentProfile: { mcpResource: URI.file('/user/mcp.json') } }));

		const runtimeServer = (id: string, resource: URI, options?: { label: string; launch: McpServerLaunch; state?: McpConnectionState; collectionId?: string }) => {
			const label = options?.label ?? installable.name;
			const definition = upcastPartial<McpServerDefinition>({ id, label, launch: options?.launch, presentation: { origin: { uri: resource, range: new Range(3, 1, 3, 5) } } });
			const collection = upcastPartial<McpCollectionDefinition>({ id: options?.collectionId ?? id, label: 'User MCP', remoteAuthority: null, scope: StorageScope.PROFILE, serverDefinitions: constObservable([definition]) });
			collections.set([...collections.get().filter(existing => existing.id !== collection.id), collection], undefined);
			return upcastPartial<IMcpServer>({
				definition: { id, label },
				collection: { id: collection.id, label: collection.label, order: 0 },
				connectionState: constObservable<McpConnectionState>(options?.state ?? { state: McpConnectionState.Kind.Stopped }),
				enablement: constObservable(ContributionEnablementState.EnabledProfile),
				readDefinitions: () => constObservable({ server: definition, collection }),
				capabilities: constObservable(undefined),
				start: async () => {
					started.push(id);
					keysAtStart.push([...secrets.keys()]);
					return { state: McpConnectionState.Kind.Stopped };
				},
			});
		};

		const install = async (server: IInstallableMcpServer | URI, options?: IWorkbencMcpServerInstallOptions) => {
			assert.ok(!URI.isUri(server));
			installs.push({ server, options });
			const target = options?.target;
			const resource = isWorkspaceFolder(target) ? target.toResource(options?.workspaceConfig === WorkspaceMcpConfigKind.Root ? rootFile : legacyFile)
				: target === ConfigurationTarget.WORKSPACE ? workspace.configuration!
					: URI.file(target === ConfigurationTarget.USER_REMOTE ? '/remote/mcp.json' : '/user/mcp.json');
			const id = `installed:${resource.path}`;
			const local = upcastPartial<IWorkbenchLocalMcpServer>({ id, name: server.name, config: server.config, mcpResource: resource, scope: LocalMcpServerScope.Workspace });
			servers.set([runtimeServer('other-same-name', URI.file('/other/.mcp.json')), runtimeServer(id, resource)], undefined);
			return local;
		};
		instantiation.stub(IWorkbenchMcpManagementService, { install });
		instantiation.stub(IMcpWorkbenchService, {
			canInstall: () => true,
			install: async (server, options) => {
				const local = await install(server.installable ?? installable, options);
				return upcastPartial<IWorkbenchMcpServer>({ id: local.id, name: local.name, local });
			},
		});
		return { instantiation, quickInput, appDialogs, chosenFiles, folder, secondFolder, workspace, opened, selections, started, secrets, keysAtStart, errors, existenceChecks, installs, agentHostAdds, agentHostGlobalAdds, servers, runtimeServer, destination: instantiation.createInstance(McpConfigurationDestination) };
	}

	test('app connections install through the existing MCP management and start flow', async () => {
		const fixture = setup(true);
		fixture.quickInput.selections.push('Notion', 'Connect with OAuth', 'Global');
		fixture.quickInput.inputs.push('notion');
		await new ConnectAppAction().run(fixture.instantiation);
		assert.deepStrictEqual(fixture.errors, []);
		assert.strictEqual(fixture.installs.length, 1);
		assert.deepStrictEqual(fixture.installs[0], {
			server: { name: 'notion', config: { type: McpServerType.REMOTE, transport: 'http', url: 'https://mcp.notion.com/mcp' }, inputs: undefined },
			options: { target: ConfigurationTarget.USER_LOCAL, workspaceConfig: undefined },
		});
		assert.deepStrictEqual(fixture.started, ['installed:/user/mcp.json']);
		assert.ok(fixture.quickInput.pickOptions[0].placeholder?.includes('review setup'));
	});

	test('Outlook and Microsoft 365 names lead to one native Work IQ connection with the published public client', async () => {
		for (const name of ['Outlook', 'Outlook Calendar', 'Microsoft 365', 'Teams', 'OneDrive', 'SharePoint']) {
			assert.strictEqual(findSignalCAppConnection(name), microsoft365WorkIq);
			assert.ok(matchesSignalCAppConnection(microsoft365WorkIq, name));
		}
		const fixture = setup(true, [], false, true);
		fixture.quickInput.selections.push('Connect with OAuth');
		fixture.quickInput.inputs.push('Work Outlook');
		await new ConnectAppAction().run(fixture.instantiation, undefined, microsoft365WorkIq.id, 'assistant');
		assert.deepStrictEqual(fixture.errors, []);
		assert.deepStrictEqual(fixture.installs, [{
			server: {
				name: 'Work Outlook',
				config: {
					type: McpServerType.REMOTE, transport: 'http', url: 'https://workiq.svc.cloud.microsoft/mcp',
					oauth: { clientId: 'ba081686-5d24-4bc6-a0d6-d034ecffed87' },
				},
				inputs: undefined,
			},
			options: { target: ConfigurationTarget.USER_LOCAL, workspaceConfig: undefined },
		}]);
		assert.deepStrictEqual(fixture.started, ['installed:/user/mcp.json']);
		assert.match(String(fixture.appDialogs[0].detail), /work or school account.*Copilot Credits/);
		assert.strictEqual(fixture.quickInput.questionOrder.filter(item => item === 'input').length, 1, 'The account setup needs only a connection name');
		assert.deepStrictEqual(fixture.keysAtStart, [[]]);
	});

	test('Assistant Slack connects in its active profile and stores its masked credential before start', async () => {
		const fixture = setup(true, [], false, true);
		fixture.instantiation.stub(IMcpCopilotGlobalConfigurationService, { getConfigurationResource: async () => URI.file('/copilot/mcp-config.json') });
		fixture.quickInput.selections.push('Connect with OAuth');
		fixture.quickInput.inputs.push('registered-slack-id', 'support-slack', 'synthetic-slack-secret');
		await new ConnectAppAction().run(fixture.instantiation, undefined, 'slack', 'assistant');
		assert.deepStrictEqual(fixture.errors, []);
		assert.deepStrictEqual(fixture.installs, [{
			server: { name: 'support-slack', config: { type: McpServerType.REMOTE, transport: 'http', url: 'https://mcp.slack.com/mcp', oauth: { clientId: 'registered-slack-id' } }, inputs: undefined },
			options: { target: ConfigurationTarget.USER_LOCAL, workspaceConfig: undefined },
		}]);
		assert.deepStrictEqual(fixture.agentHostAdds, []);
		assert.deepStrictEqual(fixture.agentHostGlobalAdds, []);
		assert.deepStrictEqual(fixture.opened, []);
		assert.deepStrictEqual(fixture.started, ['installed:/user/mcp.json']);
		assert.strictEqual(fixture.quickInput.inputOptions[2]?.password, true);
		const key = mcpOAuthClientSecretStorageKey('https://mcp.slack.com/mcp', 'registered-slack-id');
		assert.strictEqual(fixture.secrets.get(key), 'synthetic-slack-secret');
		assert.deepStrictEqual(fixture.keysAtStart, [[key]]);
	});

	for (const [id, url] of [
		['slack', 'https://mcp.slack.com/mcp'],
		['github', 'https://api.githubcopilot.com/mcp/x/all'],
	]) {
		(isNative ? test : test.skip)(`Assistant ${id} can use a registered client with the HTTPS callback`, async () => {
			const fixture = setup(true, [], false, true);
			const callback = 'https://callback.example.com/oauth/return';
			fixture.instantiation.stub(IProductService, { urlProtocol: 'signal-c', signalCOAuthRedirectUri: callback });
			fixture.quickInput.selections.push('Connect with HTTPS Callback');
			fixture.quickInput.inputs.push('registered-app-id', 'work-app', 'synthetic-app-secret');
			await new ConnectAppAction().run(fixture.instantiation, undefined, id, 'assistant');
			assert.deepStrictEqual(fixture.errors, []);
			assert.ok(String(fixture.appDialogs[0].detail).includes(callback));
			assert.deepStrictEqual(fixture.installs[0].server.config, {
				type: McpServerType.REMOTE, transport: 'http', url,
				oauth: { clientId: 'registered-app-id', useHttpsCallback: true },
			});
			assert.strictEqual(fixture.quickInput.inputOptions[2]?.password, true);
			const key = mcpOAuthClientSecretStorageKey(url, 'registered-app-id');
			assert.strictEqual(fixture.secrets.get(key), 'synthetic-app-secret');
			assert.deepStrictEqual(fixture.keysAtStart, [[key]]);
			assert.deepStrictEqual(fixture.started, ['installed:/user/mcp.json']);
		});
	}

	for (const continueWithPublicClient of [false, true]) {
		test(`Assistant Slack ${continueWithPublicClient ? 'accepts a public PKCE client' : 'cancels before writing setup'}`, async () => {
			const fixture = setup(true);
			fixture.quickInput.selections.push('Connect with OAuth');
			fixture.quickInput.inputs.push('registered-slack-id', 'support-slack', continueWithPublicClient ? '' : undefined);
			await new ConnectAppAction().run(fixture.instantiation, undefined, 'slack', 'assistant');
			assert.deepStrictEqual(fixture.errors, []);
			assert.strictEqual(fixture.installs.length, continueWithPublicClient ? 1 : 0);
			assert.strictEqual(fixture.started.length, continueWithPublicClient ? 1 : 0);
			assert.strictEqual(fixture.secrets.size, 0);
		});
	}

	test('Assistant opens existing Slack options by exact endpoint and keeps setup unchanged', async () => {
		const fixture = setup(true);
		const commands: unknown[][] = [];
		fixture.instantiation.stub(ICommandService, 'executeCommand', async (...args: unknown[]) => { commands.push(args); });
		const config = URI.file('/user/mcp.json');
		const official = URI.parse('https://mcp.slack.com/mcp');
		fixture.servers.set([
			fixture.runtimeServer('support', config, {
				label: 'Support workspace', launch: { type: McpServerTransportType.HTTP, uri: official, headers: [] },
				state: { state: McpConnectionState.Kind.Error, message: 'Credential failure details' },
			}),
			fixture.runtimeServer('engineering', config, {
				label: 'Engineering workspace', launch: { type: McpServerTransportType.HTTP, uri: URI.parse('https://MCP.SLACK.COM:443/mcp'), headers: [] },
				state: { state: McpConnectionState.Kind.Running },
			}),
			fixture.runtimeServer('lookalike', config, {
				label: 'Slack lookalike', launch: { type: McpServerTransportType.HTTP, uri: URI.parse('https://mcp.slack.com/other'), headers: [] },
			}),
		], undefined);
		fixture.quickInput.selections.push('Support workspace', undefined);
		await new ConnectAppAction().run(fixture.instantiation, undefined, 'slack', 'assistant');
		await new ConnectAppAction().run(fixture.instantiation, undefined, 'slack', 'assistant');
		assert.deepStrictEqual(fixture.errors, []);
		assert.deepStrictEqual(fixture.quickInput.pickLabels[0], ['Support workspace', 'Engineering workspace', 'Add another Slack connection']);
		assert.deepStrictEqual(fixture.quickInput.pickOptions[0].descriptions, ['Needs attention', 'Running', undefined]);
		assert.deepStrictEqual(commands, [[McpCommandIds.ServerOptions, 'support', { collectionId: 'support' }]]);
		assert.deepStrictEqual(fixture.installs, []);
		assert.deepStrictEqual(fixture.started, []);
		assert.deepStrictEqual(fixture.appDialogs, []);
		assert.strictEqual(fixture.secrets.size, 0);
	});

	test('Assistant can explicitly add a second Slack connection', async () => {
		const fixture = setup(true);
		const existing = fixture.runtimeServer('first', URI.file('/user/mcp.json'), {
			label: 'First Slack', launch: { type: McpServerTransportType.HTTP, uri: URI.parse('https://mcp.slack.com/mcp'), headers: [] },
		});
		fixture.servers.set([existing], undefined);
		fixture.quickInput.selections.push('Add another Slack connection', 'Connect with OAuth');
		fixture.quickInput.inputs.push('registered-slack-id', 'second-slack', '');
		await new ConnectAppAction().run(fixture.instantiation, undefined, 'slack', 'assistant');
		assert.deepStrictEqual(fixture.errors, []);
		assert.strictEqual(fixture.installs.length, 1);
		assert.strictEqual(fixture.installs[0].server.name, 'second-slack');
		assert.strictEqual(existing.definition.label, 'First Slack');
	});

	test('Assistant reopens a configured local QuickBooks server before asking for file paths', async () => {
		const fixture = setup(true);
		const commands: unknown[][] = [];
		fixture.instantiation.stub(ICommandService, 'executeCommand', async (...args: unknown[]) => { commands.push(args); });
		fixture.servers.set([fixture.runtimeServer('books', URI.file('/user/mcp.json'), {
			label: 'Company Books', launch: {
				type: McpServerTransportType.Stdio, command: 'node', args: ['/private/quickbooks/dist/index.js'],
				env: { QUICKBOOKS_TOKEN_STORE_PATH: '/private/tokens.json' }, cwd: undefined, envFile: undefined, sandbox: undefined,
			},
		})], undefined);
		fixture.quickInput.selections.push('Company Books');
		await new ConnectAppAction().run(fixture.instantiation, undefined, 'quickbooks-local', 'assistant');
		assert.deepStrictEqual(fixture.errors, []);
		assert.deepStrictEqual(commands, [[McpCommandIds.ServerOptions, 'books', { collectionId: 'books' }]]);
		assert.deepStrictEqual(fixture.installs, []);
		assert.deepStrictEqual(fixture.quickInput.pickOptions[0].details, ['User MCP', undefined]);
	});

	test('server options opens the native OAuth credential editor with the configured endpoint and client ID', async () => {
		const fixture = setup(true);
		fixture.instantiation.stub(IMcpSamplingService, { hasLogs: () => false });
		fixture.instantiation.stub(IAuthenticationService, {});
		fixture.instantiation.stub(IAuthenticationQueryService, { mcpServer: () => upcastPartial<ReturnType<IAuthenticationQueryService['mcpServer']>>({ getAllAccountPreferences: () => new Map() }) });
		const commands: unknown[][] = [];
		fixture.instantiation.stub(ICommandService, 'executeCommand', async (...args: unknown[]) => { commands.push(args); });
		const url = 'https://mcp.slack.com/mcp';
		fixture.servers.set([fixture.runtimeServer('support', URI.file('/user/mcp.json'), {
			label: 'Support workspace', launch: { type: McpServerTransportType.HTTP, uri: URI.parse(url), headers: [], oauth: { clientId: 'approved-id' } },
		})], undefined);
		fixture.quickInput.selections.push('Set OAuth Client Secret…');
		await new McpServerOptionsCommand().run(fixture.instantiation, 'support');
		assert.deepStrictEqual(commands, [[McpCommandIds.SetOAuthClientSecret, 'approved-id', url, 'Support workspace']]);
		assert.deepStrictEqual(fixture.started, []);
	});

	test('Assistant can sign in to an existing Microsoft 365 connection after OAuth was dismissed', async () => {
		const fixture = setup(true);
		fixture.servers.set([fixture.runtimeServer('workiq', URI.file('/user/mcp.json'), {
			label: 'Work Outlook',
			launch: {
				type: McpServerTransportType.HTTP,
				uri: URI.parse(microsoft365WorkIq.url),
				headers: [],
				oauth: { clientId: microsoft365WorkIq.publicClientId },
			},
			state: { state: McpConnectionState.Kind.Stopped, reason: 'needs-user-interaction' },
		})], undefined);
		fixture.quickInput.selections.push('Work Outlook');
		await new ConnectAppAction().run(fixture.instantiation, undefined, microsoft365WorkIq.id, 'assistant');
		assert.strictEqual(fixture.quickInput.pickOptions[0].descriptions[0], 'Sign-in required');
		assert.strictEqual(fixture.quickInput.pickLabels.length, 1);
		assert.deepStrictEqual(fixture.started, ['workiq']);
		assert.deepStrictEqual(fixture.installs, []);
		assert.deepStrictEqual(fixture.errors, []);
	});

	test('report recovery finds only the selected connection when server IDs match across collections', async () => {
		const fixture = setup(true);
		fixture.instantiation.stub(IMcpSamplingService, { hasLogs: () => false });
		fixture.instantiation.stub(IAuthenticationService, {});
		fixture.instantiation.stub(IAuthenticationQueryService, { mcpServer: () => upcastPartial<ReturnType<IAuthenticationQueryService['mcpServer']>>({ getAllAccountPreferences: () => new Map() }) });
		const uri = URI.parse(microsoft365WorkIq.url);
		const launch: McpServerLaunch = { type: McpServerTransportType.HTTP, uri, headers: [] };
		fixture.servers.set([
			fixture.runtimeServer('microsoft365', URI.file('/other/mcp.json'), { label: 'Other outlook', launch, collectionId: 'other' }),
			fixture.runtimeServer('microsoft365', URI.file('/user/mcp.json'), {
				label: 'Work Outlook', launch, collectionId: 'work',
			}),
		], undefined);
		const command = new McpServerOptionsCommand();
		await command.run(fixture.instantiation, 'microsoft365', { collectionId: 'removed', needsSignIn: true });
		assert.deepStrictEqual(fixture.quickInput.pickLabels, []);
		fixture.quickInput.selections.push('Sign In…');
		await command.run(fixture.instantiation, 'microsoft365', { collectionId: 'work', needsSignIn: true });
		assert.match(fixture.quickInput.pickOptions[0].placeholder ?? '', /Work Outlook/);
		assert.deepStrictEqual(fixture.started, ['microsoft365']);
	});

	for (const [name, id, url] of [
		['GitHub', 'github', 'https://api.githubcopilot.com/mcp/x/all'],
		['PayPal (Production)', 'paypal', 'https://mcp.paypal.com/sse'],
		['PayPal (Sandbox)', 'paypal-sandbox', 'https://mcp.sandbox.paypal.com/sse'],
		['PostHog', 'posthog', 'https://mcp.posthog.com/mcp'],
		['Supabase', 'supabase', 'https://mcp.supabase.com/mcp'],
		['Supabase (All Tools)', 'supabase-all-tools', 'https://mcp.supabase.com/mcp?features=docs,account,database,debugging,development,functions,branching,feedback,notebooks,storage'],
		['Supabase (Read Only)', 'supabase-read-only', 'https://mcp.supabase.com/mcp?read_only=true'],
		['Hugging Face', 'huggingface', 'https://huggingface.co/mcp?login'],
		['GitLab.com', 'gitlab-com', 'https://gitlab.com/api/v4/mcp'],
		['Netlify', 'netlify', 'https://mcp.netlify.com/mcp'],
		['Cal.com', 'cal-com', 'https://mcp.cal.com/mcp'],
		['Attio', 'attio', 'https://mcp.attio.com/mcp'],
		['Fireflies', 'fireflies', 'https://api.fireflies.ai/mcp'],
		['Apollo', 'apollo', 'https://mcp.apollo.io/mcp'],
		['Amplitude (US)', 'amplitude-us', 'https://mcp.amplitude.com/mcp'],
		['Amplitude (EU)', 'amplitude-eu', 'https://mcp.eu.amplitude.com/mcp'],
		['Datadog (US1)', 'datadog-us1', 'https://mcp.datadoghq.com/v1/mcp'],
		['Datadog (US3)', 'datadog-us3', 'https://mcp.us3.datadoghq.com/v1/mcp'],
		['Datadog (US5)', 'datadog-us5', 'https://mcp.us5.datadoghq.com/v1/mcp'],
		['Datadog (EU1)', 'datadog-eu1', 'https://mcp.datadoghq.eu/v1/mcp'],
		['Datadog (AP1)', 'datadog-ap1', 'https://mcp.ap1.datadoghq.com/v1/mcp'],
		['Datadog (AP2)', 'datadog-ap2', 'https://mcp.ap2.datadoghq.com/v1/mcp'],
		['Datadog (UK1)', 'datadog-uk1', 'https://mcp.uk1.datadoghq.com/v1/mcp'],
	]) {
		test(`${name} installs its own endpoint using OAuth discovery`, async () => {
			const fixture = setup(true);
			fixture.quickInput.selections.push(name, 'Connect with OAuth', 'Global');
			fixture.quickInput.inputs.push('$suggested');
			await new ConnectAppAction().run(fixture.instantiation);
			assert.deepStrictEqual(fixture.errors, []);
			assert.strictEqual(fixture.installs.length, 1);
			assert.strictEqual(fixture.installs[0].server.name.replace(/-[0-9a-f]{8}$/, ''), id);
			assert.deepStrictEqual(fixture.installs[0].server.config, {
				type: McpServerType.REMOTE, transport: id.startsWith('paypal') ? 'sse' : 'http', url,
			});
			assert.deepStrictEqual(fixture.started, ['installed:/user/mcp.json']);
		});
	}

	test('registered OAuth apps store only a public registered client ID', async () => {
		const fixture = setup(true);
		fixture.quickInput.selections.push('Asana', 'Connect with OAuth', 'Global');
		fixture.quickInput.inputs.push('  public-client-id  ', 'asana', 'fixture-asana-secret');
		await new ConnectAppAction().run(fixture.instantiation);
		assert.deepStrictEqual(fixture.errors, []);
		assert.strictEqual(fixture.installs.length, 1);
		assert.deepStrictEqual(fixture.installs[0].server.config, {
			type: McpServerType.REMOTE, transport: 'http', url: 'https://mcp.asana.com/v2/mcp', oauth: { clientId: 'public-client-id' },
		});
		assert.ok(String(fixture.appDialogs[0].detail).includes('Asana MCP OAuth app'));
		assert.strictEqual(fixture.appDialogs[0].custom, true);
		assert.ok(fixture.quickInput.prompts.includes('Asana OAuth App Client ID'));
		assert.strictEqual(fixture.secrets.get(mcpOAuthClientSecretStorageKey('https://mcp.asana.com/v2/mcp', 'public-client-id')), 'fixture-asana-secret');
		assert.strictEqual(fixture.quickInput.questionOrder.filter(item => item === 'input').length, 3);
		const validateClientId = fixture.quickInput.inputOptions[0]?.validateInput;
		assert.ok(validateClientId);
		assert.strictEqual(await validateClientId('public-client-id'), undefined);
		assert.ok(await validateClientId(''));
		assert.ok(await validateClientId('two words'));
		assert.ok(await validateClientId('client\u0000id'));
		assert.ok(await validateClientId('x'.repeat(2049)));
	});

	for (const [name, id, url] of [
		['Docusign (Production)', 'docusign', 'https://mcp.docusign.com/mcp'],
		['Docusign (Demo)', 'docusign-demo', 'https://mcp-d.docusign.com/mcp'],
	]) {
		test(`${name} uses its registered client and keeps its secret out of MCP configuration`, async () => {
			const fixture = setup(true);
			fixture.quickInput.selections.push(name, 'Connect with OAuth', 'Global');
			fixture.quickInput.inputs.push('fixture-integration-key', '$suggested', 'fixture-docusign-secret');
			await new ConnectAppAction().run(fixture.instantiation);
			assert.deepStrictEqual(fixture.errors, []);
			assert.strictEqual(fixture.installs[0]?.server.name.replace(/-[0-9a-f]{8}$/, ''), id);
			assert.deepStrictEqual(fixture.installs[0].server.config, {
				type: McpServerType.REMOTE, transport: 'http', url, oauth: { clientId: 'fixture-integration-key' },
			});
			assert.strictEqual(fixture.secrets.get(mcpOAuthClientSecretStorageKey(url, 'fixture-integration-key')), 'fixture-docusign-secret');
			assert.deepStrictEqual(fixture.started, ['installed:/user/mcp.json']);
		});
	}

	test('a gallery card opens its selected app review directly', async () => {
		const fixture = setup(true);
		fixture.quickInput.selections.push('Connect with OAuth', 'Global');
		fixture.quickInput.inputs.push('$suggested');
		await new ConnectAppAction().run(fixture.instantiation, undefined, 'datadog-eu1');
		assert.deepStrictEqual(fixture.errors, []);
		assert.strictEqual(fixture.appDialogs.length, 1);
		assert.strictEqual(fixture.appDialogs[0].message, 'Connect Datadog (EU1)');
		assert.deepStrictEqual(fixture.installs[0].server.config, {
			type: McpServerType.REMOTE, transport: 'http', url: 'https://mcp.datadoghq.eu/v1/mcp',
		});
		assert.ok(fixture.quickInput.pickLabels.every(labels => !labels.includes('Datadog (EU1)')));
	});

	test('an unknown gallery app ID leaves configuration untouched', async () => {
		const fixture = setup(true);
		await new ConnectAppAction().run(fixture.instantiation, undefined, 'unknown-app');
		assert.deepStrictEqual(fixture.appDialogs, []);
		assert.deepStrictEqual(fixture.installs, []);
		assert.deepStrictEqual(fixture.quickInput.pickLabels, []);
	});

	test('repeated app setup suggests distinct names to preserve existing connections', async () => {
		const fixture = setup(true);
		for (let index = 0; index < 2; index++) {
			fixture.quickInput.selections.push('Notion', 'Connect with OAuth', 'Global');
			fixture.quickInput.inputs.push('$suggested');
			await new ConnectAppAction().run(fixture.instantiation);
		}
		assert.deepStrictEqual(fixture.errors, []);
		assert.strictEqual(fixture.installs.length, 2);
		assert.ok(fixture.installs.every(install => install.server.name.startsWith('notion-')));
		assert.notStrictEqual(fixture.installs[0].server.name, fixture.installs[1].server.name);
	});

	for (const name of ['Gmail', 'Google Drive', 'Box', 'Canva', 'HubSpot', 'Dropbox', 'Xero']) {
		test(`${name} uses registered OAuth setup without storing a client secret in configuration`, async () => {
			const fixture = setup(true);
			const googleWorkspace = name === 'Gmail' || name === 'Google Drive';
			const confidential = googleWorkspace || ['Box', 'Canva', 'HubSpot'].includes(name);
			fixture.quickInput.selections.push(name, 'Connect with OAuth', 'Global');
			fixture.quickInput.inputs.push('test-public-client-id', '$suggested');
			if (confidential) {
				fixture.quickInput.inputs.push('fixture-client-secret');
			}
			await new ConnectAppAction().run(fixture.instantiation);
			assert.deepStrictEqual(fixture.errors, []);
			assert.strictEqual(fixture.installs.length, 1);
			const config = fixture.installs[0].server.config;
			assert.strictEqual(config.type, McpServerType.REMOTE);
			if (config.type === McpServerType.REMOTE) {
				assert.deepStrictEqual(config.oauth, { clientId: 'test-public-client-id' });
				assert.strictEqual(config.headers, undefined);
				if (googleWorkspace) {
					assert.ok(String(fixture.appDialogs[0].detail).includes('http://127.0.0.1:33418/'));
				}
				if (confidential) {
					const key = mcpOAuthClientSecretStorageKey(URI.parse(config.url).toString(true), 'test-public-client-id');
					assert.strictEqual(fixture.secrets.get(key), 'fixture-client-secret');
					const passwordInput = fixture.quickInput.inputOptions.find(input => input?.password);
					assert.ok(passwordInput?.validateInput);
					assert.ok(await passwordInput.validateInput('  '));
					assert.strictEqual(await passwordInput.validateInput('fixture-client-secret'), undefined);
				}
			}
			assert.ok(fixture.quickInput.prompts.includes(`${name} OAuth App Client ID`));
		});
	}

	test('QuickBooks hosted pilot cannot install while its metadata is incompatible', async () => {
		const fixture = setup(true, [], false, true);
		fixture.instantiation.stub(IProductService, { urlProtocol: 'signal-c', signalCOAuthRedirectUri: 'https://callback.example.com/oauth/return' });
		fixture.quickInput.selections.push('QuickBooks Online (Hosted Pilot)', undefined);
		await new ConnectAppAction().run(fixture.instantiation, undefined, undefined, 'assistant');
		assert.deepStrictEqual(fixture.quickInput.pickLabels[1], ['Open Setup Guide']);
		assert.ok(String(fixture.appDialogs[0].detail).includes('invitation-only pilot'));
		assert.ok(String(fixture.appDialogs[0].detail).includes('currently unavailable in Signal-C'));
		assert.deepStrictEqual(fixture.installs, []);
		assert.deepStrictEqual(fixture.quickInput.inputs, []);
	});

	test('Figma directs Assistant to the client approval guide when selected', async () => {
		const fixture = setup(true, [], false, true);
		const guides: string[] = [];
		fixture.instantiation.stub(IOpenerService, { open: async resource => { guides.push(resource.toString()); return true; } });
		fixture.quickInput.selections.push('Open Setup Guide');
		await new ConnectAppAction().run(fixture.instantiation, undefined, 'figma', 'assistant');
		assert.deepStrictEqual(fixture.quickInput.pickLabels[0], ['Open Setup Guide']);
		assert.ok(String(fixture.appDialogs[0].detail).includes('requires Figma to approve this MCP client'));
		assert.deepStrictEqual(guides, ['https://www.figma.com/mcp-catalog/']);
		assert.deepStrictEqual({ installs: fixture.installs, started: fixture.started, errors: fixture.errors }, { installs: [], started: [], errors: [] });
	});

	test('QuickBooks local configuration retains only paths and stays on the local machine', async () => {
		const fixture = setup(true, [], false, true);
		const entrypoint = URI.file('/official-quickbooks/dist/index.js');
		const tokenStore = URI.file('/private/quickbooks/company.env');
		fixture.chosenFiles.push([entrypoint], [tokenStore]);
		fixture.quickInput.selections.push('QuickBooks Online (Local Server)', 'Connect Local Server');
		fixture.quickInput.inputs.push('quickbooks-local');
		await new ConnectAppAction().run(fixture.instantiation);
		assert.deepStrictEqual(fixture.errors, []);
		assert.deepStrictEqual(fixture.installs.map(install => ({ config: install.server.config, options: install.options })), [{
			config: { type: McpServerType.LOCAL, command: 'node', args: [entrypoint.fsPath], env: { QUICKBOOKS_TOKEN_STORE_PATH: tokenStore.fsPath } },
			options: { target: ConfigurationTarget.USER_LOCAL, workspaceConfig: undefined },
		}]);
		assert.deepStrictEqual(fixture.agentHostAdds, []);
		assert.deepStrictEqual(fixture.agentHostGlobalAdds, []);
		assert.deepStrictEqual(fixture.existenceChecks, []);
	});

	for (const stage of ['entrypoint', 'tokenStore']) {
		test(`cancelling QuickBooks local ${stage} selection never installs a server`, async () => {
			const fixture = setup(true);
			fixture.quickInput.selections.push('QuickBooks Online (Local Server)', 'Connect Local Server');
			if (stage === 'tokenStore') { fixture.chosenFiles.push([URI.file('/official/index.js')]); }
			fixture.chosenFiles.push(undefined);
			await new ConnectAppAction().run(fixture.instantiation);
			assert.deepStrictEqual({ installs: fixture.installs, adds: fixture.agentHostAdds, errors: fixture.errors }, { installs: [], adds: [], errors: [] });
		});
	}

	(isNative ? test : test.skip)('QuickBooks sandbox opens desktop authorization before installing only the returned credential-file path', async () => {
		const fixture = setup(true);
		const entrypoint = URI.file('/official-quickbooks/dist/index.js');
		const tokenStore = URI.file('/private/profile/quickbooks-mcp/fixture.env');
		fixture.chosenFiles.push([entrypoint]);
		fixture.quickInput.selections.push('QuickBooks Online (Local Server)', 'Sign In to Sandbox');
		fixture.quickInput.inputs.push('quickbooks-sandbox', 'fixture-client', 'fixture-secret');
		fixture.instantiation.stub(IProgressService, { withProgress: async (_options, task) => task({ report() { } }) });
		let authorizations = 0;
		fixture.instantiation.stub(INativeHostService, {
			authorizeQuickBooksSandbox: async (clientId, secret) => {
				authorizations++;
				assert.deepStrictEqual([clientId, secret, fixture.installs.length], ['fixture-client', 'fixture-secret', 0]);
				return tokenStore.fsPath;
			},
		});
		await new ConnectAppAction().run(fixture.instantiation);
		assert.deepStrictEqual(fixture.errors, []);
		assert.strictEqual(authorizations, 1);
		assert.strictEqual(fixture.quickInput.inputOptions.at(-1)?.password, true);
		assert.deepStrictEqual(fixture.installs.map(install => ({ config: install.server.config, options: install.options })), [{
			config: { type: McpServerType.LOCAL, command: 'node', args: [entrypoint.fsPath], env: { QUICKBOOKS_TOKEN_STORE_PATH: tokenStore.fsPath } },
			options: { target: ConfigurationTarget.USER_LOCAL, workspaceConfig: undefined },
		}]);
	});

	for (const declined of [false, true]) {
		(isNative ? test : test.skip)(`QuickBooks sandbox sign-in ${declined ? 'resumes the running connection after cancellation' : 'restarts the same connection after approval'}`, async () => {
			const fixture = setup(true);
			fixture.instantiation.stub(IMcpSamplingService, { hasLogs: () => false });
			fixture.instantiation.stub(IAuthenticationService, {});
			fixture.instantiation.stub(IAuthenticationQueryService, { mcpServer: () => upcastPartial<ReturnType<IAuthenticationQueryService['mcpServer']>>({ getAllAccountPreferences: () => new Map() }) });
			fixture.instantiation.stub(IProgressService, { withProgress: async (_options, task) => task({ report() { } }) });
			const tokenStore = '/profile/quickbooks-mcp/10000000-0000-4000-8000-000000000000.env';
			const server = fixture.runtimeServer('quickbooks', URI.file('/user/mcp.json'), {
				label: 'QuickBooks', launch: { type: McpServerTransportType.Stdio, command: 'node', args: ['index.js'], env: { QUICKBOOKS_TOKEN_STORE_PATH: tokenStore }, envFile: undefined, cwd: undefined, sandbox: undefined },
				state: { state: declined ? McpConnectionState.Kind.Running : McpConnectionState.Kind.Stopped },
			});
			const calls: string[] = [];
			server.withServerStopped = async task => {
				calls.push('stop');
				return task();
			};
			fixture.instantiation.stub(INativeHostService, { reauthorizeQuickBooksSandbox: async path => {
				assert.strictEqual(path, tokenStore);
				assert.deepStrictEqual(fixture.started, []);
				calls.push('sign in');
				if (declined) { throw new CancellationError(); }
			} });
			fixture.servers.set([server], undefined);
			fixture.quickInput.selections.push('Sign In to QuickBooks Again…');
			await new McpServerOptionsCommand().run(fixture.instantiation, 'quickbooks');
			assert.deepStrictEqual(calls, ['stop', 'sign in']);
			assert.deepStrictEqual(fixture.started, ['quickbooks']);
			assert.deepStrictEqual({ installs: fixture.installs, errors: fixture.errors }, { installs: [], errors: [] });
		});
	}

	test('QuickBooks local setup refuses an explicit remote owner before collecting paths', async () => {
		const fixture = setup(true);
		fixture.quickInput.selections.push('QuickBooks Online (Local Server)', 'Connect Local Server');
		await new ConnectAppAction().run(fixture.instantiation, agentHostSession);
		assert.strictEqual(fixture.errors.length, 1);
		assert.ok(String(fixture.errors[0]).includes('local user configuration'));
		assert.deepStrictEqual(fixture.installs, []);
		assert.deepStrictEqual(fixture.agentHostAdds, []);
	});

	test('opening an app setup guide does not install or start a server', async () => {
		const fixture = setup(true);
		const guides: string[] = [];
		fixture.instantiation.stub(IOpenerService, { open: async resource => { guides.push(resource.toString()); return true; } });
		fixture.quickInput.selections.push('QuickBooks Online (Hosted Pilot)', 'Open Setup Guide');
		await new ConnectAppAction().run(fixture.instantiation);
		assert.deepStrictEqual(guides, ['https://intuitdeveloper.github.io/intuit-3p-ai-pilot/']);
		assert.deepStrictEqual({ installs: fixture.installs, adds: fixture.agentHostAdds, started: fixture.started, errors: fixture.errors }, { installs: [], adds: [], started: [], errors: [] });
	});

	for (const stage of ['app', 'review', 'clientId', 'name', 'destination']) {
		test(`cancelling app connection at ${stage} leaves configuration untouched`, async () => {
			const fixture = setup(true);
			if (stage === 'app') { fixture.quickInput.selections.push(undefined); }
			else {
				fixture.quickInput.selections.push(stage === 'clientId' ? 'Asana' : 'Notion', stage === 'review' ? undefined : 'Connect with OAuth');
				if (stage === 'clientId' || stage === 'name') { fixture.quickInput.inputs.push(undefined); }
				if (stage === 'destination') { fixture.quickInput.inputs.push('notion'); fixture.quickInput.selections.push(undefined); }
			}
			await new ConnectAppAction().run(fixture.instantiation);
			assert.deepStrictEqual({ installs: fixture.installs, adds: fixture.agentHostAdds, global: fixture.agentHostGlobalAdds, started: fixture.started, errors: fixture.errors }, { installs: [], adds: [], global: [], started: [], errors: [] });
		});
	}

	test('manual Add Server also exposes the app connection flow', async () => {
		const fixture = setup(true);
		fixture.quickInput.selections.push('Connect an App (OAuth)...', 'Linear', 'Connect with OAuth', 'Global');
		fixture.quickInput.inputs.push('linear');
		await new AddConfigurationAction().run(fixture.instantiation);
		assert.deepStrictEqual(fixture.errors, []);
		assert.strictEqual(fixture.installs[0].server.name, 'linear');
		assert.strictEqual(fixture.installs[0].server.config.type, McpServerType.REMOTE);
	});

	test('app connections preserve explicit agent-host ownership and policy checks', async () => {
		for (const blocked of [false, true]) {
			const fixture = setup(true);
			const selected = URI.parse('remote-example:/chosen-host?revision=4');
			if (blocked) { fixture.instantiation.stub(IAllowedMcpServersService, { isAllowed: () => new MarkdownString('App blocked by policy') }); }
			fixture.quickInput.selections.push('Notion', 'Connect with OAuth', 'Agent Host');
			fixture.quickInput.inputs.push('notion');
			await new ConnectAppAction().run(fixture.instantiation, selected);
			assert.deepStrictEqual(fixture.installs, []);
			assert.strictEqual(fixture.agentHostAdds.length, blocked ? 0 : 1);
			assert.deepStrictEqual(fixture.errors.map(String), blocked ? ['Error: App blocked by policy'] : []);
			if (!blocked) { assert.strictEqual(fixture.agentHostAdds[0].session.toString(), selected.toString()); }
		}
	});

	test('app presets contain public HTTPS endpoints and no account credentials', () => {
		const ids = new Set<string>();
		const publicUrlOptions = new Set([
			'https://huggingface.co/mcp?login',
			'https://mcp.supabase.com/mcp?read_only=true',
			'https://mcp.supabase.com/mcp?features=docs,account,database,debugging,development,functions,branching,feedback,notebooks,storage',
		]);
		for (const app of signalCAppConnections) {
			assert.ok(!ids.has(app.id), `Duplicate app ID: ${app.id}`);
			ids.add(app.id);
			assert.notStrictEqual(!!app.url, !!app.local, 'Each app has exactly one remote or local transport');
			for (const value of [app.url, app.documentation].filter((value): value is string => typeof value === 'string')) {
				const url = new URL(value);
				assert.strictEqual(url.protocol, 'https:');
				assert.strictEqual(url.username + url.password + url.hash, '');
				assert.ok(!url.search || (value === app.url && publicUrlOptions.has(value)), `Unexpected query parameters for ${app.id}`);
			}
			assert.ok(!Object.keys(app.headers ?? {}).some(header => /authorization|token|secret/i.test(header)));
		}
	});

	for (const enabled of [false, true]) {
		for (const hasRoot of [false, true]) {
			for (const hasLegacy of [false, true]) {
				const existing = [...(hasRoot ? [rootFile] : []), ...(hasLegacy ? [legacyFile] : [])];
				test(`open: enabled=${enabled}, root=${hasRoot}, legacy=${hasLegacy}`, async () => {
					const fixture = setup(enabled, existing);
					if (hasRoot && hasLegacy) {
						fixture.quickInput.selections.push(rootFile);
					}
					await new OpenWorkspaceFolderMcpResourceCommand().run(fixture.instantiation);
					assert.deepStrictEqual({
						opened: fixture.opened.map(uri => uri.path),
						pickers: fixture.quickInput.pickLabels,
						installs: fixture.installs,
					}, {
						opened: [fixture.folder.toResource(hasRoot || (!hasLegacy && enabled) ? rootFile : legacyFile).path],
						pickers: hasRoot && hasLegacy ? [[rootFile, legacyFile]] : [],
						installs: [],
					});
				});

				test(`add destination: enabled=${enabled}, root=${hasRoot}, legacy=${hasLegacy}`, async () => {
					const fixture = setup(enabled, existing);
					if (enabled && hasLegacy) {
						fixture.quickInput.selections.push(rootFile);
					}
					const result = await fixture.destination.selectForAdd(fixture.folder, installable);
					assert.deepStrictEqual({
						result,
						pickers: fixture.quickInput.pickLabels,
						descriptions: fixture.quickInput.pickOptions.map(options => options.descriptions),
					}, {
						result: enabled ? WorkspaceMcpConfigKind.Root : WorkspaceMcpConfigKind.LegacyVscode,
						pickers: enabled && hasLegacy ? [[rootFile, legacyFile]] : [],
						descriptions: enabled && hasLegacy ? [['Workspace root', 'Deprecated']] : [],
					});
				});
			}
		}
	}

	test('opening either file can be cancelled without opening or installing', async () => {
		const fixture = setup(false, [rootFile, legacyFile]);
		fixture.quickInput.selections.push(undefined);
		await new OpenWorkspaceFolderMcpResourceCommand().run(fixture.instantiation);
		assert.deepStrictEqual({ opened: fixture.opened, installs: fixture.installs }, { opened: [], installs: [] });
	});

	for (const operation of ['open', 'add'] as const) {
		test(`${operation} surfaces file provider failures instead of treating them as absence`, async () => {
			const fixture = setup(true);
			const error = new FileOperationError('Provider unavailable', FileOperationResult.FILE_OTHER_ERROR);
			fixture.instantiation.stub(IFileService, 'resolve', async () => { throw error; });
			if (operation === 'open') {
				await assert.rejects(fixture.destination.selectForOpen(fixture.folder), error);
			} else {
				await assert.rejects(fixture.destination.selectForAdd(fixture.folder, installable), error);
			}
			assert.deepStrictEqual({ pickers: fixture.quickInput.pickLabels, installs: fixture.installs }, { pickers: [], installs: [] });
		});
	}

	test('multi-root open preserves folder selection and inspects only the selected folder', async () => {
		const fixture = setup(true, [], true);
		fixture.instantiation.stub(ICommandService, 'executeCommand', async () => fixture.secondFolder);
		await new OpenWorkspaceFolderMcpResourceCommand().run(fixture.instantiation);
		assert.deepStrictEqual({
			opened: fixture.opened.map(uri => uri.path),
			checked: fixture.existenceChecks,
			pickers: fixture.quickInput.pickLabels,
		}, {
			opened: [fixture.secondFolder.toResource(rootFile).path],
			checked: [fixture.secondFolder.toResource(rootFile).path, fixture.secondFolder.toResource(legacyFile).path],
			pickers: [],
		});
	});

	for (const selection of [legacyFile, undefined]) {
		test(`unsupported root configuration offers explicit legacy selection: ${selection}`, async () => {
			const fixture = setup(true);
			fixture.quickInput.selections.push(selection);
			const server = { ...installable, inputs: [{ id: 'token', type: McpServerVariableType.PROMPT, description: 'Token', password: true }] };
			const result = await fixture.destination.selectForAdd(fixture.folder, server);
			assert.deepStrictEqual({
				result,
				pickers: fixture.quickInput.pickLabels,
				placeholder: fixture.quickInput.pickOptions[0].placeholder,
				details: fixture.quickInput.pickOptions[0].details,
				checked: fixture.existenceChecks,
			}, {
				result: selection ? WorkspaceMcpConfigKind.LegacyVscode : undefined,
				pickers: [[legacyFile]],
				placeholder: 'This server requires .vscode/mcp.json',
				details: ['\'inputs\' is not supported in .mcp.json. Use .vscode/mcp.json.'],
				checked: [],
			});
		});
	}

	test('explicit unsupported root destination fails rather than rerouting', async () => {
		const fixture = setup(true);
		await assert.rejects(fixture.destination.selectForAdd(fixture.folder, { ...installable, inputs: [{ id: 'token', type: McpServerVariableType.PROMPT, description: 'Token', password: true }] }, WorkspaceMcpConfigKind.Root));
		assert.deepStrictEqual({ pickers: fixture.quickInput.pickLabels, installs: fixture.installs }, { pickers: [], installs: [] });
	});

	for (const enabled of [false, true]) {
		test(`manual multi-root Add Server folder choices follow the flag: ${enabled}`, async () => {
			const fixture = setup(enabled, [], true);
			fixture.quickInput.selections.push('Command (stdio)', enabled ? `Workspace (${fixture.secondFolder.name})` : 'Workspace');
			fixture.quickInput.inputs.push('node server.js', installable.name);
			await new AddConfigurationAction().run(fixture.instantiation);
			assert.deepStrictEqual({
				scopes: fixture.quickInput.pickLabels[1],
				target: fixture.installs[0].options?.target,
				kind: fixture.installs[0].options?.workspaceConfig,
				opened: fixture.opened.map(uri => uri.path),
				selections: fixture.selections,
				started: fixture.started,
			}, {
				scopes: enabled ? ['Global', 'Workspace', `Workspace (${fixture.folder.name})`, `Workspace (${fixture.secondFolder.name})`] : ['Global', 'Workspace'],
				target: enabled ? fixture.secondFolder : ConfigurationTarget.WORKSPACE,
				kind: enabled ? WorkspaceMcpConfigKind.Root : undefined,
				opened: [enabled ? fixture.secondFolder.toResource(rootFile).path : fixture.workspace.configuration!.path],
				selections: [new Range(3, 1, 3, 5)],
				started: [`installed:${enabled ? fixture.secondFolder.toResource(rootFile).path : fixture.workspace.configuration!.path}`],
			});
		});
	}

	for (const enabled of [false, true]) {
		for (const selection of ['Agent Host', 'Global', 'Remote', 'Workspace', ...(enabled ? ['Workspace (second)'] : [])]) {
			test(`agent-host target is shown alongside persistent targets: ${selection}, root=${enabled}`, async () => {
				const fixture = setup(enabled, [], true, true);
				fixture.instantiation.stub(IWorkbenchEnvironmentService, { remoteAuthority: 'ssh-remote+test' });
				fixture.instantiation.stub(ILabelService, { getHostLabel: () => 'Test Remote' });
				fixture.quickInput.selections.push('Command (stdio)', selection);
				fixture.quickInput.inputs.push('node server.js', installable.name);
				await new AddConfigurationAction().run(fixture.instantiation);
				assert.deepStrictEqual({
					targets: fixture.quickInput.pickLabels[1],
					descriptions: fixture.quickInput.pickOptions[1].descriptions,
					prompts: fixture.quickInput.prompts,
					agentHostAdds: fixture.agentHostAdds,
					installs: fixture.installs.map(install => install.options),
					errors: fixture.errors,
				}, {
					targets: ['Agent Host', 'Global', 'Remote', 'Workspace', ...(enabled ? [`Workspace (${fixture.folder.name})`, `Workspace (${fixture.secondFolder.name})`] : [])],
					descriptions: ['Available to sessions on this agent host', 'Available in all workspaces, runs locally', 'Available on this remote machine, runs on Test Remote', 'Available in this workspace, runs on Test Remote', ...(enabled ? ['Workspace Folder', 'Workspace Folder'] : [])],
					prompts: ['Choose the type of MCP server to add', 'Enter Command', 'Enter Server ID', 'Select the configuration target'],
					agentHostAdds: selection === 'Agent Host' ? [{ session: agentHostSession, name: installable.name, config: installable.config }] : [],
					installs: selection === 'Agent Host' ? [] : [{
						target: selection === 'Global' ? ConfigurationTarget.USER_LOCAL
							: selection === 'Remote' ? ConfigurationTarget.USER_REMOTE
								: selection === 'Workspace' ? ConfigurationTarget.WORKSPACE : fixture.secondFolder,
						workspaceConfig: selection === 'Workspace (second)' ? WorkspaceMcpConfigKind.Root : undefined,
					}],
					errors: [],
				});
			});
		}
	}

	const secondAgentHostSession = agentHostSession.with({ path: '/second-session' });
	const localSession = URI.from({ scheme: SessionType.Local, path: '/local-session' });
	for (const asString of [false, true]) {
		test(`explicit native management target wins over another focused chat (string=${asString})`, async () => {
			const fixture = setup(true);
			const selected = URI.parse('agent-host-devin-subscription:/opaque-selection');
			fixture.instantiation.stub(IChatWidgetService, {
				lastFocusedWidget: chatWidget(agentHostSession),
				getAllWidgets: () => [chatWidget(agentHostSession), chatWidget(secondAgentHostSession)],
			});
			fixture.quickInput.selections.push('Command (stdio)', 'Agent Host');
			fixture.quickInput.inputs.push('node server.js', installable.name);
			await new AddConfigurationAction().run(fixture.instantiation, asString ? selected.toString() : selected);
			assert.deepStrictEqual({ adds: fixture.agentHostAdds.map(add => ({ ...add, session: add.session.toString() })), installs: fixture.installs, errors: fixture.errors }, {
				adds: [{ session: selected.toString(), name: installable.name, config: installable.config }], installs: [], errors: [],
			});
			assert.deepStrictEqual(fixture.quickInput.pickLabels[1], ['Agent Host', 'Global', 'Workspace']);
		});
	}

	test('an explicit remote management target still permits a local destination', async () => {
		const fixture = setup(true);
		const selected = URI.parse('remote-example:/opaque-selection?revision=2');
		fixture.quickInput.selections.push('Command (stdio)', 'Global');
		fixture.quickInput.inputs.push('node server.js', installable.name);
		await new AddConfigurationAction().run(fixture.instantiation, selected);
		assert.deepStrictEqual({ adds: fixture.agentHostAdds, installs: fixture.installs.map(install => install.options), errors: fixture.errors }, {
			adds: [], installs: [{ target: ConfigurationTarget.USER_LOCAL, workspaceConfig: undefined }], errors: [],
		});
	});

	test('an explicit remote management target preserves the selected opaque URI', async () => {
		const fixture = setup(true, [], false, true);
		const selected = URI.parse('remote-example:/opaque-selection?revision=2');
		fixture.quickInput.selections.push('Command (stdio)', 'Agent Host');
		fixture.quickInput.inputs.push('node server.js', installable.name);
		await new AddConfigurationAction().run(fixture.instantiation, selected);
		assert.deepStrictEqual({ adds: fixture.agentHostAdds.map(add => ({ ...add, session: add.session.toString() })), installs: fixture.installs, errors: fixture.errors }, {
			adds: [{ session: selected.toString(), name: installable.name, config: installable.config }], installs: [], errors: [],
		});
	});

	test('agent-host policy rejection occurs before mutating the selected host configuration', async () => {
		const fixture = setup(true);
		fixture.instantiation.stub(IAllowedMcpServersService, { isAllowed: () => new MarkdownString('Blocked by policy') });
		fixture.quickInput.selections.push('Command (stdio)', 'Agent Host');
		fixture.quickInput.inputs.push('node server.js', installable.name);
		await new AddConfigurationAction().run(fixture.instantiation, agentHostSession);
		assert.deepStrictEqual({ adds: fixture.agentHostAdds, installs: fixture.installs, errors: fixture.errors.map(String) }, {
			adds: [], installs: [], errors: ['Error: Blocked by policy'],
		});
	});

	for (const { name, focused, open, expected } of [
		{ name: 'focused agent-host session wins over other open sessions', focused: agentHostSession, open: [secondAgentHostSession, agentHostSession], expected: agentHostSession },
		{ name: 'sole open agent-host session is offered when no chat widget remains focused', focused: undefined, open: [agentHostSession], expected: agentHostSession },
		{ name: 'sole agent-host session is offered when a non-agent-host chat is focused', focused: localSession, open: [localSession, agentHostSession], expected: agentHostSession },
		{ name: 'widgets without view models are ignored', focused: undefined, open: [undefined, agentHostSession], expected: agentHostSession },
		{ name: 'duplicate widgets for the same session are not ambiguous', focused: undefined, open: [agentHostSession, URI.parse(agentHostSession.toString())], expected: agentHostSession },
		{ name: 'distinct agent-host sessions are ambiguous without focus', focused: undefined, open: [agentHostSession, secondAgentHostSession], expected: undefined },
		{ name: 'non-agent-host focus does not resolve ambiguity', focused: localSession, open: [agentHostSession, localSession, secondAgentHostSession], expected: undefined },
		{ name: 'non-agent-host sessions are not offered', focused: localSession, open: [localSession, undefined], expected: undefined },
		{ name: 'no open sessions leaves only persistent targets', focused: undefined, open: [], expected: undefined },
	]) {
		test(name, async () => {
			const fixture = setup(true);
			fixture.instantiation.stub(IChatWidgetService, {
				lastFocusedWidget: focused ? chatWidget(focused) : undefined,
				getAllWidgets: () => open.map(session => chatWidget(session)),
			});
			fixture.quickInput.selections.push('Command (stdio)', expected ? 'Agent Host' : 'Global');
			fixture.quickInput.inputs.push('node server.js', installable.name);
			await new AddConfigurationAction().run(fixture.instantiation);
			assert.deepStrictEqual({
				targets: fixture.quickInput.pickLabels[1],
				agentHostAdds: fixture.agentHostAdds,
				installs: fixture.installs.map(install => install.options),
				errors: fixture.errors,
			}, {
				targets: [...(expected ? ['Agent Host'] : []), 'Global', 'Workspace'],
				agentHostAdds: expected ? [{ session: expected, name: installable.name, config: installable.config }] : [],
				installs: expected ? [] : [{ target: ConfigurationTarget.USER_LOCAL, workspaceConfig: undefined }],
				errors: [],
			});
		});
	}

	for (const withAgentHostSession of [false, true]) {
		test(`empty window preserves destination selection: agent host=${withAgentHostSession}`, async () => {
			const fixture = setup(true, [], false, withAgentHostSession);
			fixture.instantiation.stub(IWorkspaceContextService, 'getWorkbenchState', () => WorkbenchState.EMPTY);
			fixture.quickInput.selections.push('Command (stdio)', ...(withAgentHostSession ? ['Global'] : []));
			fixture.quickInput.inputs.push('node server.js', installable.name);
			await new AddConfigurationAction().run(fixture.instantiation);
			assert.deepStrictEqual({
				targets: fixture.quickInput.pickLabels.slice(1),
				installs: fixture.installs.map(install => install.options),
				agentHostAdds: fixture.agentHostAdds,
			}, {
				targets: withAgentHostSession ? [['Agent Host', 'Global']] : [],
				installs: [{ target: ConfigurationTarget.USER_LOCAL, workspaceConfig: undefined }],
				agentHostAdds: [],
			});
		});
	}

	for (const stage of ['type', 'command', 'name', 'destination', 'file']) {
		test(`agent-host add cancellation at ${stage} has no side effects`, async () => {
			const fixture = setup(true, [legacyFile], false, true);
			fixture.quickInput.selections.push(...(stage === 'type' ? [undefined] : ['Command (stdio)']));
			if (stage === 'destination') {
				fixture.quickInput.selections.push(undefined);
			} else if (stage === 'file') {
				fixture.quickInput.selections.push('Workspace', undefined);
			}
			if (stage !== 'type') {
				fixture.quickInput.inputs.push(stage === 'command' ? undefined : 'node server.js');
				if (stage !== 'command') {
					fixture.quickInput.inputs.push(stage === 'name' ? undefined : installable.name);
				}
			}
			await new AddConfigurationAction().run(fixture.instantiation);
			assert.deepStrictEqual({
				agentHostAdds: fixture.agentHostAdds, installs: fixture.installs, opened: fixture.opened, started: fixture.started, errors: fixture.errors,
				remainingPicks: fixture.quickInput.selections, remainingInputs: fixture.quickInput.inputs,
			}, {
				agentHostAdds: [], installs: [], opened: [], started: [], errors: [],
				remainingPicks: [], remainingInputs: [],
			});
		});
	}

	for (const folderCount of [0, 1, 2]) {
		for (const enabled of [false, true]) {
			test(`manual add excludes unsupported aggregate workspace: folders=${folderCount}, root=${enabled}`, async () => {
				const fixture = setup(enabled, [], true, true);
				fixture.workspace.folders = fixture.workspace.folders.slice(0, folderCount);
				fixture.instantiation.stub(IMcpWorkspaceInstallTargetService, { getTargets: () => fixture.workspace.folders });
				const selectFolder = enabled && folderCount > 0;
				fixture.quickInput.selections.push('Command (stdio)', selectFolder ? `Workspace (${fixture.folder.name})` : 'Global');
				fixture.quickInput.inputs.push('node server.js', installable.name);
				await new AddConfigurationAction().run(fixture.instantiation);
				assert.deepStrictEqual({
					targets: fixture.quickInput.pickLabels[1],
					options: fixture.installs.map(install => install.options),
					errors: fixture.errors,
				}, {
					targets: ['Agent Host', 'Global', ...(enabled ? fixture.workspace.folders.map(folder => `Workspace (${folder.name})`) : [])],
					options: [{ target: selectFolder ? fixture.folder : ConfigurationTarget.USER_LOCAL, workspaceConfig: selectFolder ? WorkspaceMcpConfigKind.Root : undefined }],
					errors: [],
				});
			});
		}
	}

	test('folder-only destination picker can be cancelled without side effects', async () => {
		const fixture = setup(true, [], true, true);
		fixture.instantiation.stub(IMcpWorkspaceInstallTargetService, { getTargets: () => fixture.workspace.folders });
		fixture.quickInput.selections.push('Command (stdio)', undefined);
		fixture.quickInput.inputs.push('node server.js', installable.name);
		await new AddConfigurationAction().run(fixture.instantiation);
		assert.deepStrictEqual({
			targets: fixture.quickInput.pickLabels[1],
			installs: fixture.installs, agentHostAdds: fixture.agentHostAdds, opened: fixture.opened, started: fixture.started, errors: fixture.errors,
		}, {
			targets: ['Agent Host', 'Global', `Workspace (${fixture.folder.name})`, `Workspace (${fixture.secondFolder.name})`],
			installs: [], agentHostAdds: [], opened: [], started: [], errors: [],
		});
	});

	test('manual folder add cancellation does not install, reveal, or start', async () => {
		const fixture = setup(true, [legacyFile]);
		fixture.quickInput.selections.push('Command (stdio)', 'Workspace', undefined);
		fixture.quickInput.inputs.push('node server.js', installable.name);
		await new AddConfigurationAction().run(fixture.instantiation);
		assert.deepStrictEqual({ installed: fixture.installs, opened: fixture.opened, started: fixture.started, errors: fixture.errors }, { installed: [], opened: [], started: [], errors: [] });
	});

	test('manual add with interpolation requires explicit legacy confirmation', async () => {
		const fixture = setup(true);
		fixture.quickInput.selections.push('Command (stdio)', 'Workspace', legacyFile);
		fixture.quickInput.inputs.push('node ${input:token}', installable.name);
		await new AddConfigurationAction().run(fixture.instantiation);
		assert.deepStrictEqual({
			lastPicker: fixture.quickInput.pickLabels.at(-1),
			kind: fixture.installs[0].options?.workspaceConfig,
			opened: fixture.opened.map(uri => uri.path),
		}, {
			lastPicker: [legacyFile],
			kind: WorkspaceMcpConfigKind.LegacyVscode,
			opened: [fixture.folder.toResource(legacyFile).path],
		});
	});

	test('manual add still offers the aggregate workspace when enabled', async () => {
		const fixture = setup(true, [legacyFile], true);
		fixture.quickInput.selections.push('Command (stdio)', 'Workspace');
		fixture.quickInput.inputs.push('node server.js', installable.name);
		await new AddConfigurationAction().run(fixture.instantiation);
		assert.deepStrictEqual({
			pickerCount: fixture.quickInput.pickLabels.length,
			options: fixture.installs[0].options,
			opened: fixture.opened.map(uri => uri.path),
		}, {
			pickerCount: 2,
			options: { target: ConfigurationTarget.WORKSPACE, workspaceConfig: undefined },
			opened: [fixture.workspace.configuration!.path],
		});
	});

	for (const file of [rootFile, legacyFile]) {
		test(`folder-only workspace preserves explicit ${file} destination`, async () => {
			const fixture = setup(true, [rootFile, legacyFile], true, true);
			fixture.instantiation.stub(IMcpWorkspaceInstallTargetService, { getTargets: () => fixture.workspace.folders });
			fixture.quickInput.selections.push('Command (stdio)');
			fixture.quickInput.inputs.push('node server.js', installable.name);
			await new AddConfigurationAction().run(fixture.instantiation, fixture.folder.toResource(file));
			assert.deepStrictEqual({
				pickerCount: fixture.quickInput.pickLabels.length,
				options: fixture.installs.map(install => install.options),
				agentHostAdds: fixture.agentHostAdds,
				errors: fixture.errors,
			}, {
				pickerCount: 1,
				options: [{ target: fixture.folder, workspaceConfig: file === rootFile ? WorkspaceMcpConfigKind.Root : WorkspaceMcpConfigKind.LegacyVscode }],
				agentHostAdds: [],
				errors: [],
			});
		});

		test(`editor Add Server preserves ${file} and skips scope/file selection`, async () => {
			const fixture = setup(true, [rootFile, legacyFile], false, true);
			fixture.quickInput.selections.push('Command (stdio)');
			fixture.quickInput.inputs.push('node server.js', installable.name);
			const resource = fixture.folder.toResource(file);
			await new AddConfigurationAction().run(fixture.instantiation, resource);
			assert.deepStrictEqual({
				pickerCount: fixture.quickInput.pickLabels.length,
				kind: fixture.installs[0].options?.workspaceConfig,
				opened: fixture.opened.map(uri => uri.path),
				selections: fixture.selections,
				started: fixture.started,
				errors: fixture.errors,
				agentHostAdds: fixture.agentHostAdds,
			}, {
				pickerCount: 1,
				kind: file === rootFile ? WorkspaceMcpConfigKind.Root : WorkspaceMcpConfigKind.LegacyVscode,
				opened: [resource.path],
				selections: [new Range(3, 1, 3, 5)],
				started: [`installed:${resource.path}`],
				errors: [],
				agentHostAdds: [],
			});
		});
	}

	test('legacy editor Add Server still works with the flag off and a string URI', async () => {
		const fixture = setup(false, [rootFile, legacyFile]);
		fixture.quickInput.selections.push('Command (stdio)');
		fixture.quickInput.inputs.push('node server.js', installable.name);
		await new AddConfigurationAction().run(fixture.instantiation, fixture.folder.toResource(legacyFile).toString());
		assert.deepStrictEqual({
			pickerCount: fixture.quickInput.pickLabels.length,
			kind: fixture.installs[0].options?.workspaceConfig,
			opened: fixture.opened.map(uri => uri.path),
		}, {
			pickerCount: 1,
			kind: WorkspaceMcpConfigKind.LegacyVscode,
			opened: [fixture.folder.toResource(legacyFile).path],
		});
	});

	test('folder-scoped Add Server uses the selected workspace folder', async () => {
		const fixture = setup(true, [], true);
		fixture.quickInput.selections.push('Command (stdio)');
		fixture.quickInput.inputs.push('node server.js', installable.name);

		await new AddConfigurationAction().run(fixture.instantiation, fixture.secondFolder.uri);

		assert.deepStrictEqual({
			pickerCount: fixture.quickInput.pickLabels.length,
			target: fixture.installs[0].options?.target,
			kind: fixture.installs[0].options?.workspaceConfig,
			opened: fixture.opened.map(uri => uri.path),
		}, {
			pickerCount: 1,
			target: fixture.secondFolder,
			kind: WorkspaceMcpConfigKind.Root,
			opened: [fixture.secondFolder.toResource(rootFile).path],
		});
	});

	for (const path of ['/project/nested/.mcp.json', '/outside/.mcp.json', '/project/settings.json']) {
		test(`invalid editor destination fails before prompting: ${path}`, async () => {
			const fixture = setup(true);
			await new AddConfigurationAction().run(fixture.instantiation, URI.file(path));
			assert.deepStrictEqual({ pickers: fixture.quickInput.pickLabels, installs: fixture.installs, errors: fixture.errors }, {
				pickers: [],
				installs: [],
				errors: [new Error('Select a .mcp.json or .vscode/mcp.json file at the root of an open workspace folder.')],
			});
		});
	}

	test('explicit root editor add is rejected when the flag is off', async () => {
		const fixture = setup(false, [rootFile]);
		await new AddConfigurationAction().run(fixture.instantiation, fixture.folder.toResource(rootFile));
		assert.deepStrictEqual({ pickers: fixture.quickInput.pickLabels, installs: fixture.installs, errors: fixture.errors }, {
			pickers: [],
			installs: [],
			errors: [new Error('Enable chat.mcp.workspaceRootConfig.enabled to add servers to .mcp.json, or use .vscode/mcp.json.')],
		});
	});

	test('explicit root editor add notifies once for unsupported configuration without side effects', async () => {
		const fixture = setup(true, [rootFile, legacyFile]);
		fixture.quickInput.selections.push('Command (stdio)');
		fixture.quickInput.inputs.push('node ${input:token}', installable.name);
		await new AddConfigurationAction().run(fixture.instantiation, fixture.folder.toResource(rootFile));
		assert.deepStrictEqual({
			pickerCount: fixture.quickInput.pickLabels.length,
			installs: fixture.installs,
			opened: fixture.opened,
			started: fixture.started,
			errors: fixture.errors,
		}, {
			pickerCount: 1,
			installs: [],
			opened: [],
			started: [],
			errors: [new Error('\'${...}\' is not supported in .mcp.json. Use .vscode/mcp.json.')],
		});
	});

	for (const cancelled of [false, true]) {
		test(`editor add ${cancelled ? 'silently handles cancellation' : 'notifies once for an installation failure'}`, async () => {
			const fixture = setup(true, [rootFile, legacyFile]);
			const error = cancelled ? new CancellationError() : new Error('Unable to write MCP configuration');
			fixture.instantiation.stub(IWorkbenchMcpManagementService, 'install', async () => { throw error; });
			fixture.quickInput.selections.push('Command (stdio)');
			fixture.quickInput.inputs.push('node server.js', installable.name);
			await new AddConfigurationAction().run(fixture.instantiation, fixture.folder.toResource(rootFile));
			assert.deepStrictEqual({ errors: fixture.errors, opened: fixture.opened, started: fixture.started }, {
				errors: cancelled ? [] : [error],
				opened: [],
				started: [],
			});
		});
	}

	suite('manual Copilot Global', () => {
		function setupGlobal(resource: URI | undefined = URI.file('/home/me/.copilot/mcp-config.json')) {
			const fixture = setup(true);
			const writes: { servers: IInstallableMcpServer[]; resource: URI; format: McpResourceFormat | undefined }[] = [];
			const notifications: Parameters<INotificationService['info']>[0][] = [];
			let lookups = 0;
			fixture.instantiation.stub(IMcpCopilotGlobalConfigurationService, {
				getConfigurationResource: async () => { lookups++; return resource; },
			});
			fixture.instantiation.stub(IMcpResourceScannerService, {
				addMcpServers: async (servers, resource, _target, format) => { writes.push({ servers, resource, format }); },
			});
			fixture.instantiation.stub(ILabelService, { getUriLabel: uri => uri.toString(true), getHostLabel: () => 'test-host' });
			fixture.instantiation.stub(INotificationService, 'info', (message: Parameters<INotificationService['info']>[0]) => { notifications.push(message); });
			return { ...fixture, writes, notifications, get lookups() { return lookups; } };
		}

		test('adds Gmail to Copilot Global and saves the client secret separately', async () => {
			const fixture = setupGlobal();
			const url = 'https://gmailmcp.googleapis.com/mcp/v1';
			const key = mcpOAuthClientSecretStorageKey(url, 'gmail-client');
			fixture.quickInput.selections.push('Connect with OAuth', 'Global', 'Copilot Global');
			fixture.quickInput.inputs.push('gmail-client', 'work-gmail', '  fixture-secret  ');
			await new ConnectAppAction().run(fixture.instantiation, undefined, 'gmail');
			assert.deepStrictEqual(fixture.errors, []);
			assert.deepStrictEqual(fixture.agentHostGlobalAdds, [{
				name: 'work-gmail',
				config: { type: McpServerType.REMOTE, transport: 'http', url, oauth: { clientId: 'gmail-client' } },
			}]);
			assert.strictEqual(fixture.quickInput.inputOptions[2]?.password, true);
			assert.strictEqual(fixture.secrets.get(key), 'fixture-secret');
			assert.strictEqual(fixture.notifications.length, 1);
		});

		for (const credential of ['existing', 'missing', 'failed'] as const) {
			test(`handles a ${credential} Gmail credential when adding to Copilot Global`, async () => {
				const fixture = setupGlobal();
				const key = mcpOAuthClientSecretStorageKey('https://gmailmcp.googleapis.com/mcp/v1', 'gmail-client');
				fixture.quickInput.selections.push('Connect with OAuth', 'Global', 'Copilot Global');
				fixture.quickInput.inputs.push('gmail-client', 'work-gmail');
				if (credential === 'existing') {
					fixture.secrets.set(key, 'existing-secret');
				} else {
					fixture.quickInput.inputs.push(credential === 'failed' ? 'fixture-secret' : undefined);
				}
				if (credential === 'failed') {
					fixture.instantiation.stub(ICustomizationHarnessService, {
						getActiveDescriptor: () => upcastDeepPartial<IHarnessDescriptor>({
							mcpConfigurationProvider: { addGlobal: async () => { throw new Error('Configuration write failed'); } },
						}),
					});
				}
				await new ConnectAppAction().run(fixture.instantiation, undefined, 'gmail');
				assert.strictEqual(fixture.agentHostGlobalAdds.length, credential === 'existing' ? 1 : 0);
				assert.strictEqual(fixture.secrets.get(key), credential === 'existing' ? 'existing-secret' : undefined);
				assert.strictEqual(fixture.opened.length, credential === 'existing' ? 1 : 0);
				assert.strictEqual(fixture.errors.length, credential === 'failed' ? 1 : 0);
			});
		}

		for (const remote of [false, true]) {
			for (const http of [false, true]) {
				test(`writes manual ${http ? 'HTTP' : 'stdio'} to ${remote ? 'remote' : 'local'} Copilot Global`, async () => {
					const resource = remote ? URI.parse('vscode-remote://ssh-remote+test/home/me/custom/mcp-config.json') : URI.file('/home/me/.copilot/mcp-config.json');
					const fixture = setupGlobal(resource);
					if (remote) {
						fixture.instantiation.stub(IWorkbenchEnvironmentService, { remoteAuthority: 'ssh-remote+test' });
					}
					fixture.quickInput.selections.push(http ? 'HTTP (HTTP or Server-Sent Events)' : 'Command (stdio)', 'Global', 'Copilot Global');
					fixture.quickInput.inputs.push(http ? 'https://example.com/mcp' : 'node server.js', installable.name);
					await new AddConfigurationAction().run(fixture.instantiation);
					assert.deepStrictEqual({
						writes: fixture.writes,
						agentHostGlobalAdds: fixture.agentHostGlobalAdds,
						opened: fixture.opened,
						installs: fixture.installs,
						started: fixture.started,
						pickers: fixture.quickInput.pickLabels.slice(1),
						globalDescriptions: fixture.quickInput.pickOptions.at(-1)?.descriptions,
						globalDetails: fixture.quickInput.pickOptions.at(-1)?.details,
						questions: fixture.quickInput.questionOrder,
						notified: fixture.notifications.length === 1 && String(fixture.notifications[0]).includes('environment variables on the agent-host machine'),
					}, {
						writes: [],
						agentHostGlobalAdds: [{
							name: installable.name,
							config: http ? { type: McpServerType.REMOTE, url: 'https://example.com/mcp' } : installable.config,
						}],
						opened: [resource],
						installs: [], started: [],
						pickers: [remote ? ['Global', 'Remote', 'Workspace'] : ['Global', 'Workspace'], ['Copilot Global', 'User Configuration (deprecated)']],
						globalDescriptions: [resource.toString(true), URI.file('/user/mcp.json').toString(true)],
						globalDetails: [undefined, undefined],
						questions: ['pick', 'input', 'input', 'pick', 'pick'],
						notified: true,
					});
				});
			}
		}

		for (const choice of ['User Configuration (deprecated)', undefined]) {
			test(`global choice ${choice ?? 'cancellation'} preserves legacy behavior`, async () => {
				const fixture = setupGlobal();
				fixture.quickInput.selections.push('Command (stdio)', 'Global', choice);
				fixture.quickInput.inputs.push('node server.js', installable.name);
				await new AddConfigurationAction().run(fixture.instantiation);
				assert.deepStrictEqual({ writes: fixture.writes, targets: fixture.installs.map(install => install.options?.target), notifications: fixture.notifications },
					{ writes: [], targets: choice ? [ConfigurationTarget.USER_LOCAL] : [], notifications: [] });
			});
		}

		test('shows the active VS Code profile destination rather than the default profile', async () => {
			const fixture = setupGlobal();
			const profileResource = URI.file('/user/profiles/custom/mcp.json');
			fixture.instantiation.stub(IUserDataProfileService, upcastDeepPartial<IUserDataProfileService>({ currentProfile: { mcpResource: profileResource } }));
			fixture.quickInput.selections.push('Command (stdio)', 'Global', undefined);
			fixture.quickInput.inputs.push('node server.js', installable.name);
			await new AddConfigurationAction().run(fixture.instantiation);
			assert.deepStrictEqual(fixture.quickInput.pickOptions.at(-1)?.descriptions, [
				URI.file('/home/me/.copilot/mcp-config.json').toString(true),
				profileResource.toString(true),
			]);
		});

		test('unavailable host falls back to VS Code user configuration', async () => {
			const fixture = setupGlobal();
			fixture.instantiation.stub(IMcpCopilotGlobalConfigurationService, { getConfigurationResource: async () => undefined });
			fixture.quickInput.selections.push('Command (stdio)', 'Global');
			fixture.quickInput.inputs.push('node server.js', installable.name);
			await new AddConfigurationAction().run(fixture.instantiation);
			assert.deepStrictEqual({ targets: fixture.installs.map(install => install.options?.target), writes: fixture.writes, pickers: fixture.quickInput.pickLabels.length },
				{ targets: [ConfigurationTarget.USER_LOCAL], writes: [], pickers: 2 });
		});

		test('older Agent Host falls back to writing Copilot Global configuration', async () => {
			const fixture = setupGlobal();
			fixture.instantiation.stub(ICustomizationHarnessService, { getActiveDescriptor: () => upcastDeepPartial<IHarnessDescriptor>({}) });
			fixture.quickInput.selections.push('Command (stdio)', 'Global', 'Copilot Global');
			fixture.quickInput.inputs.push('node server.js', installable.name);

			await new AddConfigurationAction().run(fixture.instantiation);

			assert.deepStrictEqual({
				writes: fixture.writes,
				agentHostGlobalAdds: fixture.agentHostGlobalAdds,
			}, {
				writes: [{
					servers: [{ ...installable, inputs: undefined }],
					resource: URI.file('/home/me/.copilot/mcp-config.json'),
					format: McpResourceFormat.CopilotGlobal,
				}],
				agentHostGlobalAdds: [],
			});
		});

		for (const failure of ['policy', 'provider']) {
			test(`${failure} failure is surfaced without opening a file or reporting success`, async () => {
				const fixture = setupGlobal();
				if (failure === 'policy') {
					fixture.instantiation.stub(IAllowedMcpServersService, { isAllowed: () => new MarkdownString('Blocked by policy') });
				} else {
					fixture.instantiation.stub(ICustomizationHarnessService, {
						getActiveDescriptor: () => upcastDeepPartial<IHarnessDescriptor>({
							mcpConfigurationProvider: { addGlobal: async () => { throw new Error('Provider failed'); } },
						}),
					});
				}
				fixture.quickInput.selections.push('Command (stdio)', 'Global', 'Copilot Global');
				fixture.quickInput.inputs.push('node server.js', installable.name);
				await new AddConfigurationAction().run(fixture.instantiation);
				assert.deepStrictEqual({ errors: fixture.errors.map(String), writes: fixture.writes, opened: fixture.opened, notifications: fixture.notifications },
					{ errors: [failure === 'policy' ? 'Error: Blocked by policy' : 'Error: Provider failed'], writes: [], opened: [], notifications: [] });
			});
		}

		test('explicit file additions never resolve or offer Copilot Global', async () => {
			const fixture = setupGlobal();
			fixture.quickInput.selections.push('Command (stdio)');
			fixture.quickInput.inputs.push('node server.js', installable.name);
			await new AddConfigurationAction().run(fixture.instantiation, fixture.folder.toResource(legacyFile));
			assert.deepStrictEqual({ lookups: fixture.lookups, writes: fixture.writes, pickers: fixture.quickInput.pickLabels.length, installs: fixture.installs.length },
				{ lookups: 0, writes: [], pickers: 1, installs: 1 });
		});

		function setupPackage(fixture: ReturnType<typeof setup>, result: McpGeneratedConfiguration) {
			const requests: { targetFormat: McpResourceFormat; targetConfig: object }[] = [];
			const savedInputs: { key: string; value: string }[] = [];
			fixture.instantiation.stub(IMcpRegistry, 'setSavedInput', async (key: string, _target: ConfigurationTarget, value: string) => { savedInputs.push({ key, value }); });
			fixture.instantiation.stub(ICommandService, 'executeCommand', async (command: string, args: { targetFormat: McpResourceFormat; targetConfig: object }) => {
				switch (command) {
					case 'github.copilot.chat.mcp.setup.check': return true;
					case 'github.copilot.chat.mcp.setup.validatePackage':
						fixture.quickInput.questionOrder.push('validate');
						requests.push(args);
						return { state: 'ok', publisher: 'test' };
					case 'github.copilot.chat.mcp.setup.flow':
						fixture.quickInput.questionOrder.push('generate');
						return result;
				}
				throw new Error(`Unexpected command: ${command}`);
			});
			fixture.instantiation.stub(IQuickInputService, 'createQuickPick', <T extends IQuickPickItem>() => {
				const accepted = store.add(new Emitter<IQuickPickDidAcceptEvent>());
				let items: readonly T[] = [];
				return upcastPartial<IQuickPick<T>>({
					set items(value: readonly T[]) { items = value; queueMicrotask(() => accepted.fire({ inBackground: false })); },
					get selectedItems() { return items.slice(0, 1); },
					onDidAccept: accepted.event,
					onDidHide: Event.None,
					show: () => { },
					dispose: () => accepted.dispose(),
				});
			});
			return { requests, savedInputs };
		}

		for (const packageLabel of ['NPM Package', 'Pip Package', 'NuGet Package', 'Docker Image']) {
			test(`${packageLabel} selects Copilot Global before generation and preserves remote transport`, async () => {
				const fixture = setupGlobal();
				fixture.instantiation.stub(IConfigurationService, new TestConfigurationService({ [mcpWorkspaceRootConfig]: true, 'chat.mcp.assisted.nuget.enabled': true }));
				const { requests } = setupPackage(fixture, { type: 'assisted', format: McpResourceFormat.CopilotGlobal, name: 'test', server: { type: 'sse', url: 'https://example.com', oauthClientId: 'client', headers: { Authorization: 'Bearer ${TOKEN}' } } });
				fixture.quickInput.selections.push(packageLabel, 'Global', 'Copilot Global');
				fixture.quickInput.inputs.push('test-package', installable.name);
				await new AddConfigurationAction().run(fixture.instantiation);
				assert.deepStrictEqual({
					format: requests[0].targetFormat,
					schema: requests[0].targetConfig,
					config: fixture.agentHostGlobalAdds[0]?.config,
					questions: fixture.quickInput.questionOrder,
					errors: fixture.errors,
				}, {
					format: McpResourceFormat.CopilotGlobal,
					schema: getMcpGenerationSchema(McpResourceFormat.CopilotGlobal),
					config: { type: 'http', transport: 'sse', url: 'https://example.com', headers: { Authorization: 'Bearer ${TOKEN}' }, oauth: { clientId: 'client' }, dev: undefined },
					questions: ['pick', 'pick', 'pick', 'input', 'validate', 'generate', 'input'],
					errors: [],
				});
			});
		}

		test('package-assisted additions can still select VS Code user configuration', async () => {
			const fixture = setupGlobal();
			setupPackage(fixture, { type: 'mapped', name: 'test', server: installable.config });
			fixture.quickInput.selections.push('NPM Package', 'Global', 'User Configuration (deprecated)');
			fixture.quickInput.inputs.push('test-package', installable.name);
			await new AddConfigurationAction().run(fixture.instantiation);
			assert.deepStrictEqual({ lookups: fixture.lookups, writes: fixture.writes, pickers: fixture.quickInput.pickLabels.slice(1), targets: fixture.installs.map(install => install.options?.target), errors: fixture.errors },
				{ lookups: 1, writes: [], pickers: [['Global', 'Workspace'], ['Copilot Global', 'User Configuration (deprecated)']], targets: [ConfigurationTarget.USER_LOCAL], errors: [] });
		});

		for (const choice of [legacyFile, undefined]) {
			test(`generated workspace inputs require explicit fallback: ${choice}`, async () => {
				const fixture = setupGlobal();
				const result: McpGeneratedConfiguration = { type: 'assisted', format: McpResourceFormat.WorkspaceRoot, name: 'test', server: { command: 'node', args: ['--token=${input:token}'] }, inputs: [{ id: 'token', type: McpServerVariableType.PROMPT, description: 'Token', password: true }], inputValues: { '${input:token}': 'secret' } };
				const { requests, savedInputs } = setupPackage(fixture, result);
				fixture.quickInput.selections.push('NPM Package', 'Workspace', choice);
				fixture.quickInput.inputs.push('test-package', 'test');
				await new AddConfigurationAction().run(fixture.instantiation);
				assert.deepStrictEqual({
					format: requests[0].targetFormat,
					configs: fixture.installs.map(i => i.options?.workspaceConfig),
					savedInputs, writes: fixture.writes, errors: fixture.errors,
				}, {
					format: McpResourceFormat.WorkspaceRoot,
					configs: choice ? [WorkspaceMcpConfigKind.LegacyVscode] : [],
					savedInputs: choice ? [{ key: '${input:token}', value: 'secret' }] : [], writes: [], errors: [],
				});
			});
		}

		test('explicit root file does not silently redirect generated inputs', async () => {
			const fixture = setupGlobal();
			setupPackage(fixture, { type: 'mapped', name: 'test', server: { type: McpServerType.LOCAL, command: 'node', args: ['${input:token}'] } });
			fixture.quickInput.selections.push('NPM Package');
			fixture.quickInput.inputs.push('test-package', 'test');
			await new AddConfigurationAction().run(fixture.instantiation, fixture.folder.toResource(rootFile));
			assert.deepStrictEqual({ installs: fixture.installs, writes: fixture.writes, lookups: fixture.lookups, errors: fixture.errors.map(String) },
				{ installs: [], writes: [], lookups: 0, errors: ['Error: \'${...}\' is not supported in .mcp.json. Use .vscode/mcp.json.'] });
		});

		test('package destination cancellation prevents validation and generation', async () => {
			const fixture = setupGlobal();
			const { requests } = setupPackage(fixture, { type: 'mapped', name: 'test', server: installable.config });
			fixture.quickInput.selections.push('NPM Package', 'Global', undefined);
			await new AddConfigurationAction().run(fixture.instantiation);
			assert.deepStrictEqual({ requests, writes: fixture.writes, installs: fixture.installs, errors: fixture.errors },
				{ requests: [], writes: [], installs: [], errors: [] });
		});

		test('package Copilot Global policy rejection occurs before any write', async () => {
			const fixture = setupGlobal();
			setupPackage(fixture, { type: 'mapped', name: 'test', server: installable.config });
			fixture.instantiation.stub(IAllowedMcpServersService, { isAllowed: () => new MarkdownString('Blocked by policy') });
			fixture.quickInput.selections.push('NPM Package', 'Global', 'Copilot Global');
			fixture.quickInput.inputs.push('test-package', 'test');
			await new AddConfigurationAction().run(fixture.instantiation);
			assert.deepStrictEqual({ writes: fixture.writes, opened: fixture.opened, errors: fixture.errors.map(String) },
				{ writes: [], opened: [], errors: ['Error: Blocked by policy'] });
		});

		test('workspace generation preserves SSE and offers a legacy file instead of changing transport', async () => {
			const fixture = setupGlobal();
			setupPackage(fixture, { type: 'assisted', format: McpResourceFormat.WorkspaceRoot, name: 'test', server: { type: 'sse', url: 'https://example.com' } });
			fixture.quickInput.selections.push('Pip Package', 'Workspace', legacyFile);
			fixture.quickInput.inputs.push('test-package', 'test');
			await new AddConfigurationAction().run(fixture.instantiation);
			assert.deepStrictEqual({
				config: fixture.installs[0]?.server.config,
				destination: fixture.installs[0]?.options?.workspaceConfig,
				errors: fixture.errors,
			}, { config: { type: McpServerType.REMOTE, transport: 'sse', url: 'https://example.com', headers: undefined, dev: undefined }, destination: WorkspaceMcpConfigKind.LegacyVscode, errors: [] });
		});

		for (const choice of ['User Configuration (deprecated)', undefined]) {
			test(`Copilot manifest inputs require explicit VS Code fallback: ${choice}`, async () => {
				const fixture = setupGlobal();
				const inputs = [{ id: 'token', type: McpServerVariableType.PROMPT, description: 'Token', password: true }];
				const config: IMcpServerConfiguration = { type: McpServerType.LOCAL, command: 'node', args: ['${input:token}'] };
				setupPackage(fixture, { type: 'mapped', name: 'test', server: config, inputs });
				fixture.quickInput.selections.push('NPM Package', 'Global', 'Copilot Global', choice);
				fixture.quickInput.inputs.push('test-package', 'test');
				await new AddConfigurationAction().run(fixture.instantiation);
				assert.deepStrictEqual({ installed: fixture.installs.map(i => i.server), writes: fixture.writes, errors: fixture.errors },
					{ installed: choice ? [{ name: 'test', config, inputs }] : [], writes: [], errors: [] });
			});
		}

		test('explicit legacy file is selected before generation without scope pickers', async () => {
			const fixture = setupGlobal();
			const { requests } = setupPackage(fixture, { type: 'mapped', name: 'test', server: installable.config });
			fixture.quickInput.selections.push('NPM Package');
			fixture.quickInput.inputs.push('test-package', 'test');
			await new AddConfigurationAction().run(fixture.instantiation, fixture.folder.toResource(legacyFile));
			assert.deepStrictEqual({ format: requests[0].targetFormat, lookups: fixture.lookups, pickers: fixture.quickInput.pickLabels.length, kind: fixture.installs[0]?.options?.workspaceConfig, errors: fixture.errors },
				{ format: McpResourceFormat.Vscode, lookups: 0, pickers: 1, kind: WorkspaceMcpConfigKind.LegacyVscode, errors: [] });
		});

		for (const file of [rootFile, legacyFile, undefined]) {
			test(`workspace file is chosen before generation: ${file}`, async () => {
				const fixture = setup(true, [legacyFile]);
				const { requests } = setupPackage(fixture, { type: 'mapped', name: 'test', server: installable.config });
				fixture.quickInput.selections.push('NPM Package', 'Workspace', file);
				if (file) {
					fixture.quickInput.inputs.push('test-package', 'test');
				}
				await new AddConfigurationAction().run(fixture.instantiation);
				assert.deepStrictEqual({
					formats: requests.map(r => r.targetFormat),
					questions: fixture.quickInput.questionOrder,
					kinds: fixture.installs.map(i => i.options?.workspaceConfig),
					errors: fixture.errors,
				}, {
					formats: file ? [file === rootFile ? McpResourceFormat.WorkspaceRoot : McpResourceFormat.Vscode] : [],
					questions: file ? ['pick', 'pick', 'pick', 'input', 'validate', 'generate', 'input'] : ['pick', 'pick', 'pick'],
					kinds: file ? [file === rootFile ? WorkspaceMcpConfigKind.Root : WorkspaceMcpConfigKind.LegacyVscode] : [],
					errors: [],
				});
			});
		}

		test('current agent session rejects unresolved inputs rather than losing collected values', async () => {
			const fixture = setup(true, [], false, true);
			setupPackage(fixture, { type: 'mapped', name: 'test', server: { type: McpServerType.LOCAL, command: 'node', args: ['${input:token}'] } });
			fixture.quickInput.selections.push('NPM Package', 'Agent Host');
			fixture.quickInput.inputs.push('test-package', 'test');
			await new AddConfigurationAction().run(fixture.instantiation);
			assert.deepStrictEqual({ adds: fixture.agentHostAdds, installs: fixture.installs, errors: fixture.errors.map(String) },
				{ adds: [], installs: [], errors: ['Error: This server requires VS Code input variables. Add it to a VS Code configuration file instead of the current agent session.'] });
		});

		test('server-name cancellation has no writes or saved inputs', async () => {
			const fixture = setupGlobal();
			const { savedInputs } = setupPackage(fixture, { type: 'mapped', name: 'test', server: installable.config });
			fixture.quickInput.selections.push('NPM Package', 'Global', 'Copilot Global');
			fixture.quickInput.inputs.push('test-package', undefined);
			await new AddConfigurationAction().run(fixture.instantiation);
			assert.deepStrictEqual({ savedInputs, installs: fixture.installs, writes: fixture.writes, errors: fixture.errors },
				{ savedInputs: [], installs: [], writes: [], errors: [] });
		});
	});

	suite('Open User Configuration', () => {
		const copilotGlobal = URI.file('/home/me/.copilot/mcp-config.json');
		const vscodeUser = URI.file('/user/mcp.json');

		function setupOpenUser(resource: URI | undefined, existing: URI[]) {
			const fixture = setup(true);
			const created: { resource: string; content: string }[] = [];
			const commands: unknown[][] = [];
			fixture.instantiation.stub(IMcpCopilotGlobalConfigurationService, { getConfigurationResource: async () => resource });
			fixture.instantiation.stub(ILabelService, { getUriLabel: uri => uri.path });
			fixture.instantiation.stub(IFileService, {
				exists: async resource => existing.some(uri => extUri.isEqual(uri, resource)),
				createFile: async (resource, content) => {
					created.push({ resource: resource.path, content: content instanceof VSBuffer ? content.toString() : '' });
					return upcastPartial<IFileStatWithMetadata>({ resource });
				},
			});
			fixture.instantiation.stub(ICommandService, 'executeCommand', async (...args: unknown[]) => { commands.push(args); });
			return { ...fixture, created, commands };
		}

		for (const [selection, existing] of [['Copilot Global', []], ['Copilot Global', [copilotGlobal]], ['Copilot Global', [vscodeUser]], ['VS Code User Configuration', [copilotGlobal, vscodeUser]]] as const) {
			test(`opens ${selection} and offers the legacy file only when it exists (existing: ${existing.map(uri => uri.path).join(', ') || 'none'})`, async () => {
				const fixture = setupOpenUser(copilotGlobal, [...existing]);
				fixture.quickInput.selections.push(selection);
				await new OpenUserMcpResourceCommand().run(fixture.instantiation);
				const selected = selection === 'Copilot Global' ? copilotGlobal : vscodeUser;
				const hasCopilotGlobal = existing.some(uri => uri === copilotGlobal);
				const hasVscodeUser = existing.some(uri => uri === vscodeUser);
				assert.deepStrictEqual({
					pickers: fixture.quickInput.pickLabels,
					descriptions: fixture.quickInput.pickOptions.map(options => options.descriptions),
					details: fixture.quickInput.pickOptions.map(options => options.details),
					created: fixture.created,
					opened: fixture.opened.map(uri => uri.path),
					commands: fixture.commands,
				}, {
					pickers: [['Copilot Global', ...(hasVscodeUser ? ['VS Code User Configuration'] : []), 'Open MCP Customizations']],
					descriptions: [[hasCopilotGlobal ? 'Recommended' : 'Recommended, not existing', ...(hasVscodeUser ? ['Legacy'] : []), undefined]],
					details: [['mcp-config.json', ...(hasVscodeUser ? ['mcp.json'] : []), undefined]],
					created: hasCopilotGlobal || selected !== copilotGlobal ? [] : [{
						resource: copilotGlobal.path,
						content: JSON.stringify({ mcpServers: {} }, null, '\t'),
					}],
					opened: [selected.path],
					commands: [],
				});
			});
		}

		test('can open MCP customizations instead of a file', async () => {
			const fixture = setupOpenUser(copilotGlobal, []);
			fixture.quickInput.selections.push('Open MCP Customizations');
			await new OpenUserMcpResourceCommand().run(fixture.instantiation);
			assert.deepStrictEqual({ created: fixture.created, opened: fixture.opened, commands: fixture.commands },
				{ created: [], opened: [], commands: [[AICustomizationManagementCommands.OpenEditor, AICustomizationManagementSection.McpServers]] });
		});

		test('cancellation neither creates nor opens a file', async () => {
			const fixture = setupOpenUser(copilotGlobal, []);
			fixture.quickInput.selections.push(undefined);
			await new OpenUserMcpResourceCommand().run(fixture.instantiation);
			assert.deepStrictEqual({ created: fixture.created, opened: fixture.opened, commands: fixture.commands }, { created: [], opened: [], commands: [] });
		});

		for (const [name, resource, options] of [['unavailable Copilot Global', undefined, undefined], ['explicit legacy request', copilotGlobal, { legacy: true }]] as const) {
			test(`${name} opens VS Code user configuration without prompting`, async () => {
				const fixture = setupOpenUser(resource, []);
				await new OpenUserMcpResourceCommand().run(fixture.instantiation, options);
				assert.deepStrictEqual({ pickers: fixture.quickInput.pickLabels, created: fixture.created.map(file => file.resource), opened: fixture.opened.map(uri => uri.path) },
					{ pickers: [], created: [vscodeUser.path], opened: [vscodeUser.path] });
			});
		}
	});

	suite('Open Remote User Configuration', () => {
		const remoteCopilot = URI.parse('vscode-remote://ssh-remote+test/home/me/.copilot/mcp-config.json');
		const remoteVscode = URI.parse('vscode-remote://ssh-remote+test/home/me/.vscode-server/data/User/mcp.json');

		function setupOpenRemote(resource: URI | undefined, existing: URI[]) {
			const fixture = setup(true);
			const created: { resource: string; content: string }[] = [];
			fixture.instantiation.stub(IMcpCopilotGlobalConfigurationService, { getConfigurationResource: async () => resource });
			fixture.instantiation.stub(IRemoteUserDataProfilesService, upcastDeepPartial<IRemoteUserDataProfilesService>({ getRemoteProfile: async () => ({ mcpResource: remoteVscode }) }));
			fixture.instantiation.stub(ILabelService, { getUriLabel: uri => uri.path });
			fixture.instantiation.stub(IFileService, {
				exists: async resource => existing.some(uri => extUri.isEqual(uri, resource)),
				createFile: async (resource, content) => {
					created.push({ resource: resource.toString(), content: content instanceof VSBuffer ? content.toString() : '' });
					return upcastPartial<IFileStatWithMetadata>({ resource });
				},
			});
			return { ...fixture, created };
		}

		test('opens Copilot directly and creates it when the VS Code file is absent', async () => {
			const fixture = setupOpenRemote(remoteCopilot, []);
			await new OpenRemoteUserMcpResourceCommand().run(fixture.instantiation);
			assert.deepStrictEqual({ pickers: fixture.quickInput.pickLabels, created: fixture.created, opened: fixture.opened.map(uri => uri.toString()) }, {
				pickers: [],
				created: [{ resource: remoteCopilot.toString(), content: JSON.stringify({ mcpServers: {} }, null, '\t') }],
				opened: [remoteCopilot.toString()],
			});
		});

		for (const selection of ['Copilot Remote Configuration', 'VS Code Remote User Configuration']) {
			test(`offers both files when the VS Code file exists and opens ${selection} on the remote`, async () => {
				const fixture = setupOpenRemote(remoteCopilot, [remoteVscode]);
				fixture.quickInput.selections.push(selection);
				await new OpenRemoteUserMcpResourceCommand().run(fixture.instantiation);
				const isCopilot = selection === 'Copilot Remote Configuration';
				const target = isCopilot ? remoteCopilot : remoteVscode;
				assert.deepStrictEqual({
					pickers: fixture.quickInput.pickLabels,
					descriptions: fixture.quickInput.pickOptions.map(options => options.descriptions),
					created: fixture.created,
					opened: fixture.opened.map(uri => uri.toString()),
				}, {
					pickers: [['Copilot Remote Configuration', 'VS Code Remote User Configuration']],
					descriptions: [['Recommended, not existing', 'VS Code profile']],
					created: isCopilot ? [{ resource: remoteCopilot.toString(), content: JSON.stringify({ mcpServers: {} }, null, '\t') }] : [],
					opened: [target.toString()],
				});
			});
		}

		test('existing files are described as existing and not recreated', async () => {
			const fixture = setupOpenRemote(remoteCopilot, [remoteCopilot, remoteVscode]);
			fixture.quickInput.selections.push('VS Code Remote User Configuration');
			await new OpenRemoteUserMcpResourceCommand().run(fixture.instantiation);
			assert.deepStrictEqual({ descriptions: fixture.quickInput.pickOptions.map(options => options.descriptions), created: fixture.created, opened: fixture.opened.map(uri => uri.toString()) },
				{ descriptions: [['Recommended', 'VS Code profile']], created: [], opened: [remoteVscode.toString()] });
		});

		test('cancellation neither creates nor opens a file', async () => {
			const fixture = setupOpenRemote(remoteCopilot, [remoteVscode]);
			fixture.quickInput.selections.push(undefined);
			await new OpenRemoteUserMcpResourceCommand().run(fixture.instantiation);
			assert.deepStrictEqual({ created: fixture.created, opened: fixture.opened }, { created: [], opened: [] });
		});

		test('unavailable Copilot configuration opens the VS Code remote file without prompting', async () => {
			const fixture = setupOpenRemote(undefined, []);
			await new OpenRemoteUserMcpResourceCommand().run(fixture.instantiation);
			assert.deepStrictEqual({ pickers: fixture.quickInput.pickLabels, opened: fixture.opened.map(uri => uri.toString()) },
				{ pickers: [], opened: [remoteVscode.toString()] });
		});
	});

	test('editor menu retains AI gates and only offers root additions when enabled', () => {
		const menu = new AddConfigurationAction().desc.menu;
		assert.ok(menu && !Array.isArray(menu) && menu.when);
		const visible = (file: string, enabled: boolean, hidden = false) => {
			const values: Record<string, ContextKeyValue> = {
				[ResourceContextKey.Path.key]: `/project/${file}`,
				[ActiveEditorContext.key]: TEXT_FILE_EDITOR_ID,
				[`config.${mcpWorkspaceRootConfig}`]: enabled,
				[ChatContextKeys.Setup.hidden.key]: hidden,
				[ChatContextKeys.Setup.disabledInWorkspace.key]: false,
			};
			return menu.when!.evaluate({ getValue: <T extends ContextKeyValue>(key: string) => values[key] as T });
		};
		assert.deepStrictEqual({
			rootOff: visible(rootFile, false),
			rootOn: visible(rootFile, true),
			legacyOff: visible(legacyFile, false),
			legacyOn: visible(legacyFile, true),
			hiddenRoot: visible(rootFile, true, true),
			hiddenLegacy: visible(legacyFile, true, true),
			otherFile: visible('settings.json', true),
		}, {
			rootOff: false,
			rootOn: true,
			legacyOff: true,
			legacyOn: true,
			hiddenRoot: false,
			hiddenLegacy: false,
			otherFile: false,
		});
	});

	for (const gallery of [false, true]) {
		for (const enabled of [false, true]) {
			test(`Install in Workspace: gallery=${gallery}, enabled=${enabled}`, async () => {
				const fixture = setup(enabled, [legacyFile]);
				if (!gallery && enabled) {
					fixture.quickInput.selections.push(rootFile);
				}
				const action = store.add(fixture.instantiation.createInstance(InstallInWorkspaceAction, false));
				action.mcpServer = upcastPartial<IWorkbenchMcpServer>({
					name: installable.name,
					gallery: gallery ? upcastPartial<IGalleryMcpServer>({ name: installable.name }) : undefined,
					installable: gallery ? undefined : installable,
					installState: McpServerInstallState.Uninstalled,
				});
				await action.run();
				const file = !gallery && enabled ? rootFile : legacyFile;
				assert.deepStrictEqual({
					kind: fixture.installs[0].options?.workspaceConfig,
					pickers: fixture.quickInput.pickLabels,
					started: fixture.started,
				}, {
					kind: !gallery && enabled ? WorkspaceMcpConfigKind.Root : WorkspaceMcpConfigKind.LegacyVscode,
					pickers: !gallery && enabled ? [[rootFile, legacyFile]] : [],
					started: [`installed:${fixture.folder.toResource(file).path}`],
				});
			});
		}
	}

	test('raw URL workspace install can cancel the configuration-file picker', async () => {
		const fixture = setup(true, [legacyFile]);
		fixture.quickInput.selections.push(undefined);
		const action = store.add(fixture.instantiation.createInstance(InstallInWorkspaceAction, false));
		action.mcpServer = upcastPartial<IWorkbenchMcpServer>({ name: installable.name, installable, installState: McpServerInstallState.Uninstalled });
		await action.run();
		assert.deepStrictEqual({ installs: fixture.installs, started: fixture.started }, { installs: [], started: [] });
	});

	test('gallery workspace install keeps folder and aggregate scope selection', async () => {
		const fixture = setup(true, [legacyFile], true);
		fixture.quickInput.selections.push('Workspace');
		const action = store.add(fixture.instantiation.createInstance(InstallInWorkspaceAction, false));
		action.mcpServer = upcastPartial<IWorkbenchMcpServer>({ name: installable.name, gallery: upcastPartial<IGalleryMcpServer>({ name: installable.name }), installState: McpServerInstallState.Uninstalled });
		await action.run();
		assert.deepStrictEqual({
			pickers: fixture.quickInput.pickLabels,
			options: fixture.installs[0].options,
		}, {
			pickers: [[`Workspace (${fixture.folder.name})`, `Workspace (${fixture.secondFolder.name})`, 'Workspace']],
			options: { target: ConfigurationTarget.WORKSPACE, workspaceConfig: undefined },
		});
	});

	for (const gallery of [false, true]) {
		for (const folderCount of [0, 1, 2]) {
			test(`Install in Workspace respects folder-only targets: gallery=${gallery}, folders=${folderCount}`, async () => {
				const fixture = setup(true, [], true);
				fixture.workspace.folders = fixture.workspace.folders.slice(0, folderCount);
				fixture.instantiation.stub(IMcpWorkspaceInstallTargetService, { getTargets: () => fixture.workspace.folders });
				if (folderCount > 1) {
					fixture.quickInput.selections.push(`Workspace (${fixture.folder.name})`);
				}
				const action = store.add(fixture.instantiation.createInstance(InstallInWorkspaceAction, false));
				const mcpServer = upcastPartial<IWorkbenchMcpServer>({
					name: installable.name,
					gallery: gallery ? upcastPartial<IGalleryMcpServer>({ name: installable.name }) : undefined,
					installable: gallery ? undefined : installable,
					installState: McpServerInstallState.Uninstalled,
				});
				action.mcpServer = mcpServer;
				const menuActions = getContextMenuActions(mcpServer, false, fixture.instantiation).flat();
				for (const action of menuActions) {
					if (isDisposable(action)) {
						store.add(action);
					}
				}
				await action.run();
				assert.deepStrictEqual({
					enabled: action.enabled,
					menuActionIds: menuActions.map(action => action.id),
					pickers: fixture.quickInput.pickLabels,
					options: fixture.installs.map(install => install.options),
					started: fixture.started,
				}, {
					enabled: folderCount > 0,
					menuActionIds: folderCount > 0 ? [action.id] : [],
					pickers: folderCount > 1 ? [[`Workspace (${fixture.folder.name})`, `Workspace (${fixture.secondFolder.name})`]] : [],
					options: folderCount > 0 ? [{ target: fixture.folder, workspaceConfig: gallery ? WorkspaceMcpConfigKind.LegacyVscode : WorkspaceMcpConfigKind.Root }] : [],
					started: folderCount > 0 ? [`installed:${fixture.folder.toResource(gallery ? legacyFile : rootFile).path}`] : [],
				});
			});
		}
	}

	for (const remote of [false, true]) {
		test(`${remote ? 'remote' : 'global'} install starts the exact installed ID`, async () => {
			const fixture = setup(true);
			if (remote) {
				fixture.instantiation.stub(IWorkbenchEnvironmentService, { remoteAuthority: 'test-remote' });
				fixture.instantiation.stub(ILabelService, { getHostLabel: () => 'Remote' });
			}
			const action = store.add(remote ? fixture.instantiation.createInstance(InstallInRemoteAction, false) : fixture.instantiation.createInstance(InstallAction, false));
			action.mcpServer = upcastPartial<IWorkbenchMcpServer>({ name: installable.name, installable, installState: McpServerInstallState.Uninstalled });
			await action.run();
			assert.deepStrictEqual(fixture.started, [`installed:/${remote ? 'remote' : 'user'}/mcp.json`]);
		});
	}

	test('Show Configuration (JSON) selects the exact ID among duplicate names', async () => {
		const fixture = setup(true);
		const legacy = fixture.runtimeServer('legacy-id', fixture.folder.toResource(legacyFile));
		const root = fixture.runtimeServer('root-id', fixture.folder.toResource(rootFile));
		fixture.servers.set([legacy, root], undefined);
		const action = store.add(fixture.instantiation.createInstance(ShowServerJsonConfigurationAction));
		action.mcpServer = upcastPartial<IWorkbenchMcpServer>({ id: 'root-id', name: installable.name, local: upcastPartial<IWorkbenchLocalMcpServer>({ id: 'root-id' }) });
		await action.run();
		assert.deepStrictEqual(fixture.opened.map(uri => uri.path), [fixture.folder.toResource(rootFile).path]);
	});

	test('startup waits for the installed ID instead of starting an existing same-name server', async () => {
		const fixture = setup(true);
		const legacy = fixture.runtimeServer('legacy-id', fixture.folder.toResource(legacyFile));
		const root = fixture.runtimeServer('root-id', fixture.folder.toResource(rootFile));
		fixture.servers.set([legacy], undefined);
		const waiting = startServerByFilter(fixture.instantiation.get(IMcpService), server => server.definition.id === 'root-id');
		const beforeDiscovery = [...fixture.started];
		fixture.servers.set([legacy, root], undefined);
		await waiting;
		assert.deepStrictEqual({ beforeDiscovery, afterDiscovery: fixture.started }, { beforeDiscovery: [], afterDiscovery: ['root-id'] });
	});

	test('startup timeout disposes the discovery observer', async () => {
		const fixture = setup(true);
		await assert.rejects(startServerByFilter(fixture.instantiation.get(IMcpService), () => true, 0), CancellationError);
		fixture.servers.set([fixture.runtimeServer('late-id', fixture.folder.toResource(rootFile))], undefined);
		assert.deepStrictEqual(fixture.started, []);
	});
});
