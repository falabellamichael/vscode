/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { assertNever } from '../../../../base/common/assert.js';
import { disposableTimeout } from '../../../../base/common/async.js';
import { CancellationError } from '../../../../base/common/errors.js';
import { escapeIcons } from '../../../../base/common/iconLabels.js';
import { parse as parseJsonc } from '../../../../base/common/jsonc.js';
import { mnemonicButtonLabel } from '../../../../base/common/labels.js';
import { DisposableStore } from '../../../../base/common/lifecycle.js';
import { Schemas } from '../../../../base/common/network.js';
import { autorunSelfDisposable } from '../../../../base/common/observable.js';
import { getSignalCOAuthRedirectUri } from '../../../../base/common/oauth.js';
import { isNative } from '../../../../base/common/platform.js';
import { basename, isEqual } from '../../../../base/common/resources.js';
import { URI } from '../../../../base/common/uri.js';
import { generateUuid } from '../../../../base/common/uuid.js';
import { hasKey } from '../../../../base/common/types.js';
import { localize } from '../../../../nls.js';
import { ICommandService } from '../../../../platform/commands/common/commands.js';
import { ConfigurationTarget, IConfigurationService } from '../../../../platform/configuration/common/configuration.js';
import { IFileService } from '../../../../platform/files/common/files.js';
import { IInstantiationService } from '../../../../platform/instantiation/common/instantiation.js';
import { ILabelService } from '../../../../platform/label/common/label.js';
import { IMcpRemoteServerConfiguration, IMcpServerConfiguration, IMcpServerVariable, IMcpStdioServerConfiguration, McpServerType } from '../../../../platform/mcp/common/mcpPlatformTypes.js';
import { IAllowedMcpServersService, IGalleryMcpServerConfiguration, RegistryType } from '../../../../platform/mcp/common/mcpManagement.js';
import { IMcpResourceScannerService } from '../../../../platform/mcp/common/mcpResourceScannerService.js';
import { getCopilotGlobalMcpConfigurationError } from '../../../../platform/mcp/common/mcpCopilotGlobalConfiguration.js';
import { McpResourceFormat } from '../../../../platform/mcp/common/mcpWorkspaceConfiguration.js';
import { INativeHostService, isQuickBooksOAuthCredential } from '../../../../platform/native/common/native.js';
import { INotificationService } from '../../../../platform/notification/common/notification.js';
import { IOpenerService } from '../../../../platform/opener/common/opener.js';
import { IProgressService, ProgressLocation } from '../../../../platform/progress/common/progress.js';
import { IProductService } from '../../../../platform/product/common/productService.js';
import { IQuickInputService, IQuickPickItem, QuickPickInput } from '../../../../platform/quickinput/common/quickInput.js';
import { ISecretStorageService } from '../../../../platform/secrets/common/secrets.js';
import { ITelemetryService } from '../../../../platform/telemetry/common/telemetry.js';
import { isWorkspaceFolder, IWorkspaceContextService, IWorkspaceFolder, WorkbenchState } from '../../../../platform/workspace/common/workspace.js';
import { IDialogService, IFileDialogService } from '../../../../platform/dialogs/common/dialogs.js';
import { IEditorService } from '../../../services/editor/common/editorService.js';
import { IWorkbenchEnvironmentService } from '../../../services/environment/common/environmentService.js';
import { IWorkbenchLocalMcpServer, IWorkbenchMcpManagementService, WorkspaceMcpConfigKind } from '../../../services/mcp/common/mcpWorkbenchManagementService.js';
import { IUserDataProfileService } from '../../../services/userDataProfile/common/userDataProfile.js';
import { IMcpWorkspaceInstallTargetService } from '../../../services/mcp/common/mcpWorkspaceInstallTargetService.js';
import { IAgentHostCustomizationService } from '../../chat/browser/agentSessions/agentHost/agentHostCustomizationService.js';
import { IChatWidgetService } from '../../chat/browser/chat.js';
import { ICustomizationHarnessService } from '../../chat/common/customizationHarnessService.js';
import { isContributionDisabled } from '../../chat/common/enablement.js';
import { isAgentHostTarget } from '../../chat/common/chatSessionsService.js';
import { getChatSessionType } from '../../chat/common/model/chatUri.js';
import { McpCommandIds } from '../common/mcpCommandIds.js';
import { allDiscoverySources, ExternalDiscoverySource, mcpDiscoverySection, mcpWorkspaceRootConfig } from '../common/mcpConfiguration.js';
import { getMcpGenerationSchema, McpGeneratedConfiguration, normalizeMcpGeneratedConfiguration } from '../common/mcpConfigurationGeneration.js';
import { IMcpRegistry } from '../common/mcpRegistryTypes.js';
import { IMcpServer, IMcpService, McpConnectionState, mcpOAuthClientSecretStorageKey, McpServerTransportType } from '../common/mcpTypes.js';
import { ILogService } from '../../../../platform/log/common/log.js';
import { getSignalCDirectMailProvider, ISignalCAppConnection, signalCAppConnections, signalCDirectMailConnections } from '../common/signalCAppConnections.js';
import { ISignalCDirectMail } from './signalCDirectMail.js';
import { IWorkspaceMcpConfigurationTarget, McpConfigurationDestination } from './mcpConfigurationDestination.js';
import { IMcpCopilotGlobalConfigurationService } from '../common/mcpCopilotGlobalConfigurationService.js';

export const enum AddConfigurationType {
	Stdio,
	HTTP,

	NpmPackage,
	PipPackage,
	NuGetPackage,
	DockerImage,
	App,
}

type AssistedConfigurationType = AddConfigurationType.NpmPackage | AddConfigurationType.PipPackage | AddConfigurationType.NuGetPackage | AddConfigurationType.DockerImage;

type McpInstallTarget = { kind: 'local'; target: ConfigurationTarget | IWorkspaceFolder } | { kind: 'agentHost'; session: URI };

/** A bound host session supplies the host candidate while setup still offers storage choices. */
export type McpAddConfigurationTarget = IWorkspaceMcpConfigurationTarget | { readonly sessionResource: URI };

export const AssistedTypes = {
	[AddConfigurationType.NpmPackage]: {
		title: localize('mcp.npm.title', "Enter NPM Package Name"),
		placeholder: localize('mcp.npm.placeholder', "Package name (e.g., @org/package)"),
		pickLabel: localize('mcp.serverType.npm', "NPM Package"),
		pickDescription: localize('mcp.serverType.npm.description', "Install from an NPM package name"),
		enabledConfigKey: null, // always enabled
	},
	[AddConfigurationType.PipPackage]: {
		title: localize('mcp.pip.title', "Enter Pip Package Name"),
		placeholder: localize('mcp.pip.placeholder', "Package name (e.g., package-name)"),
		pickLabel: localize('mcp.serverType.pip', "Pip Package"),
		pickDescription: localize('mcp.serverType.pip.description', "Install from a Pip package name"),
		enabledConfigKey: null, // always enabled
	},
	[AddConfigurationType.NuGetPackage]: {
		title: localize('mcp.nuget.title', "Enter NuGet Package Name"),
		placeholder: localize('mcp.nuget.placeholder', "Package name (e.g., Package.Name)"),
		pickLabel: localize('mcp.serverType.nuget', "NuGet Package"),
		pickDescription: localize('mcp.serverType.nuget.description', "Install from a NuGet package name"),
		enabledConfigKey: 'chat.mcp.assisted.nuget.enabled',
	},
	[AddConfigurationType.DockerImage]: {
		title: localize('mcp.docker.title', "Enter Docker Image Name"),
		placeholder: localize('mcp.docker.placeholder', "Image name (e.g., mcp/imagename)"),
		pickLabel: localize('mcp.serverType.docker', "Docker Image"),
		pickDescription: localize('mcp.serverType.docker.description', "Install from a Docker image"),
		enabledConfigKey: null, // always enabled
	},
};

const enum AddConfigurationCopilotCommand {
	/** Returns whether MCP enhanced setup is enabled. */
	IsSupported = 'github.copilot.chat.mcp.setup.check',

	/** Takes an npm/pip package name, validates its owner. */
	ValidatePackage = 'github.copilot.chat.mcp.setup.validatePackage',

	/** Returns the resolved MCP configuration. */
	StartFlow = 'github.copilot.chat.mcp.setup.flow',
}

type ValidatePackageResult =
	{ state: 'ok'; publisher: string; name?: string; version?: string }
	| { state: 'error'; error: string; helpUri?: string; helpUriLabel?: string };

type AddServerData = {
	packageType: string;
};
type AddServerClassification = {
	owner: 'digitarald';
	comment: 'Generic details for adding a new MCP server';
	packageType: { classification: 'SystemMetaData'; purpose: 'FeatureInsight'; comment: 'The type of MCP server package' };
};
type AddServerCompletedData = {
	packageType: string;
	serverType: string | undefined;
	target: string;
};
type AddServerCompletedClassification = {
	owner: 'digitarald';
	comment: 'Generic details for successfully adding model-assisted MCP server';
	packageType: { classification: 'SystemMetaData'; purpose: 'FeatureInsight'; comment: 'The type of MCP server package' };
	serverType: { classification: 'SystemMetaData'; purpose: 'FeatureInsight'; comment: 'The type of MCP server' };
	target: { classification: 'SystemMetaData'; purpose: 'FeatureInsight'; comment: 'The target of the MCP server configuration' };
};

type McpSetupDestination = {
	installTarget: McpInstallTarget;
	format: McpResourceFormat;
	workspaceConfig?: WorkspaceMcpConfigKind;
	resource?: URI;
};

export function authorizeQuickBooksInBrowser<T>(instantiationService: IInstantiationService, authorize: (host: INativeHostService, operationId: string) => Promise<T>): Promise<T> {
	const [nativeHost, progress] = instantiationService.invokeFunction(accessor => [accessor.get(INativeHostService), accessor.get(IProgressService)] as const);
	const operationId = generateUuid();
	return progress.withProgress({ location: ProgressLocation.Notification, title: localize('mcp.app.quickBooksSigningIn', "Complete QuickBooks sandbox sign-in in your browser"), cancellable: true },
		() => authorize(nativeHost, operationId),
		() => { void nativeHost.cancelQuickBooksSandboxAuthorization(operationId).catch(() => { /* the window or operation already closed */ }); });
}

export class McpAddConfigurationCommand {
	constructor(
		private readonly configurationTarget: McpAddConfigurationTarget | undefined,
		@IQuickInputService private readonly _quickInputService: IQuickInputService,
		@IWorkbenchMcpManagementService private readonly _mcpManagementService: IWorkbenchMcpManagementService,
		@IWorkspaceContextService private readonly _workspaceService: IWorkspaceContextService,
		@IWorkbenchEnvironmentService private readonly _environmentService: IWorkbenchEnvironmentService,
		@ICommandService private readonly _commandService: ICommandService,
		@IMcpRegistry private readonly _mcpRegistry: IMcpRegistry,
		@IOpenerService private readonly _openerService: IOpenerService,
		@IEditorService private readonly _editorService: IEditorService,
		@IFileService private readonly _fileService: IFileService,
		@INotificationService private readonly _notificationService: INotificationService,
		@ITelemetryService private readonly _telemetryService: ITelemetryService,
		@IMcpService private readonly _mcpService: IMcpService,
		@ILabelService private readonly _label: ILabelService,
		@IUserDataProfileService private readonly _userDataProfileService: IUserDataProfileService,
		@IMcpCopilotGlobalConfigurationService private readonly _copilotGlobalConfigurationService: IMcpCopilotGlobalConfigurationService,
		@IMcpResourceScannerService private readonly _mcpResourceScannerService: IMcpResourceScannerService,
		@IAllowedMcpServersService private readonly _allowedMcpServersService: IAllowedMcpServersService,
		@IConfigurationService private readonly _configurationService: IConfigurationService,
		@IAgentHostCustomizationService private readonly _agentHostCustomizations: IAgentHostCustomizationService,
		@ICustomizationHarnessService private readonly _customizationHarnessService: ICustomizationHarnessService,
		@IChatWidgetService private readonly _chatWidgetService: IChatWidgetService,
		@IInstantiationService private readonly _instantiationService: IInstantiationService,
		@IMcpWorkspaceInstallTargetService private readonly _workspaceInstallTargetService: IMcpWorkspaceInstallTargetService,
		@IDialogService private readonly _dialogService: IDialogService,
		@IFileDialogService private readonly _fileDialogService: IFileDialogService,
		@ISecretStorageService private readonly _secretStorageService: ISecretStorageService,
		@IProductService private readonly _productService: IProductService,
	) { }

	private async getServerType(): Promise<AddConfigurationType | undefined> {
		type TItem = { kind: AddConfigurationType | 'browse' | 'discovery' } & IQuickPickItem;
		const items: QuickPickInput<TItem>[] = [
			{ kind: AddConfigurationType.App, label: localize('mcp.serverType.app', "Connect an App (OAuth)..."), description: localize('mcp.serverType.app.description', "Choose a published app connection, including QuickBooks, Notion, Slack and GitHub") },
			{ kind: AddConfigurationType.Stdio, label: localize('mcp.serverType.command', "Command (stdio)"), description: localize('mcp.serverType.command.description', "Run a local command that implements the MCP protocol") },
			{ kind: AddConfigurationType.HTTP, label: localize('mcp.serverType.http', "HTTP (HTTP or Server-Sent Events)"), description: localize('mcp.serverType.http.description', "Connect to a remote HTTP server that implements the MCP protocol") }
		];

		let aiSupported: boolean | undefined;
		try {
			aiSupported = await this._commandService.executeCommand<boolean>(AddConfigurationCopilotCommand.IsSupported);
		} catch {
			// ignored
		}

		if (aiSupported) {
			items.unshift({ type: 'separator', label: localize('mcp.serverType.manual', "Manual Install") });

			const elligableTypes = Object.entries(AssistedTypes).map(([type, { pickLabel, pickDescription, enabledConfigKey }]) => {
				if (enabledConfigKey) {
					const enabled = this._configurationService.getValue<boolean>(enabledConfigKey) ?? false;
					if (!enabled) {
						return;
					}
				}
				return {
					kind: Number(type) as AddConfigurationType,
					label: pickLabel,
					description: pickDescription,
				};
			}).filter(x => !!x);

			items.push(
				{ type: 'separator', label: localize('mcp.serverType.copilot', "Model-Assisted") },
				...elligableTypes
			);
		}

		items.push({ type: 'separator' });

		const discovery = this._configurationService.getValue<{ [K in ExternalDiscoverySource]: boolean }>(mcpDiscoverySection);
		if (discovery && typeof discovery === 'object' && allDiscoverySources.some(d => !discovery[d])) {
			items.push({
				kind: 'discovery',
				label: localize('mcp.servers.discovery', "Add from another application..."),
			});
		}

		items.push({
			kind: 'browse',
			label: localize('mcp.servers.browse', "Browse MCP Servers..."),
		});

		const result = await this._quickInputService.pick<TItem>(items, {
			placeHolder: localize('mcp.serverType.placeholder', "Choose the type of MCP server to add"),
		});

		if (result?.kind === 'browse') {
			this._commandService.executeCommand(McpCommandIds.Browse);
			return undefined;
		}

		if (result?.kind === 'discovery') {
			this._commandService.executeCommand('workbench.action.openSettings', mcpDiscoverySection);
			return undefined;
		}

		return result?.kind;
	}

	private async pickAssistantAppConnection(app: ISignalCAppConnection): Promise<IMcpServer | 'new' | undefined> {
		const endpoint = app.url && new URL(app.url).href;
		const existing = this._mcpService.servers.get().filter(server => {
			const launch = server.readDefinitions().get().server?.launch;
			if (app.local) {
				return launch?.type === McpServerTransportType.Stdio && launch.command === app.local.command && !!launch.env[app.local.tokenStoreVariable];
			}
			if (!endpoint || launch?.type !== McpServerTransportType.HTTP) {
				return false;
			}
			const url = launch.uri.toString(true);
			return URL.canParse(url) && new URL(url).href === endpoint;
		});
		if (!existing.length) { return 'new'; }

		const items: (IQuickPickItem & { connection: IMcpServer | 'new' })[] = existing.map(server => {
			const state = server.connectionState.get();
			const description = isContributionDisabled(server.enablement.get()) ? localize('mcp.app.connectionDisabled', "Disabled")
				: state.state === McpConnectionState.Kind.Error ? localize('mcp.app.connectionNeedsAttention', "Needs attention") : McpConnectionState.toString(state);
			return { label: escapeIcons(server.definition.label), description, detail: escapeIcons(server.collection.label), connection: server };
		});
		if (!app.blocked) {
			items.push({ label: localize('mcp.app.addAnotherConnection', "Add another {0} connection", app.name), connection: 'new' });
		}
		return (await this._quickInputService.pick(items, {
			title: localize('mcp.app.existingConnections', "{0} connections", app.name),
			placeHolder: localize('mcp.app.pickExistingConnection', "Choose a connection to manage or select a new connection"),
			matchOnDescription: true,
			matchOnDetail: true,
			ignoreFocusLost: true,
		}))?.connection;
	}

	private async getAppConfig(appId?: string, connectInAssistant = false): Promise<{ app: ISignalCAppConnection; config: IMcpServerConfiguration; quickBooksSandbox?: true } | undefined> {
		const catalog: readonly ISignalCAppConnection[] = connectInAssistant ? [...Object.values(signalCDirectMailConnections), ...signalCAppConnections] : signalCAppConnections;
		const connection = appId !== undefined ? catalog.find(app => app.id === appId) : (await this._quickInputService.pick(catalog.map(app => ({
			label: app.name,
			description: app.blocked ? localize('mcp.app.setupOnly', "Setup guide · connection unavailable") : app.local ? localize('mcp.app.localOAuth', "Local server · existing OAuth setup") : app.clientIdRequired ? localize('mcp.app.registeredOAuth', "OAuth · app registration required") : localize('mcp.app.oauth', "OAuth"),
			detail: app.description,
			app,
		})), {
			title: localize('mcp.app.title', "Signal-C App Connections"),
			placeHolder: localize('mcp.app.placeholder', "Search apps. Select to review setup before adding a connection."),
			matchOnDescription: true, matchOnDetail: true, ignoreFocusLost: true,
		}))?.app;
		if (!connection) { return undefined; }
		const directMail = getSignalCDirectMailProvider(connection.id);
		if (directMail) {
			await this._instantiationService.invokeFunction(accessor => accessor.get(ISignalCDirectMail)).manage(directMail);
			return undefined;
		}
		if (connectInAssistant) {
			const selection = await this.pickAssistantAppConnection(connection);
			if (!selection) { return undefined; }
			if (selection !== 'new') {
				if (!isContributionDisabled(selection.enablement.get()) && McpConnectionState.needsSignIn(selection.connectionState.get())) {
					await selection.start({ promptType: 'all-untrusted' });
				} else {
					await this._commandService.executeCommand(McpCommandIds.ServerOptions, selection.definition.id, { collectionId: selection.collection.id });
				}
				return undefined;
			}
		}
		const httpsCallback = isNative && connectInAssistant && !connection.blocked && (connection.clientIdRequired || connection.supportsClientSecret) && !connection.requiresDesktopCallback
			? getSignalCOAuthRedirectUri(this._productService) : undefined;
		const { result } = await this._dialogService.prompt({
			type: 'info', custom: true,
			message: localize('mcp.app.review', "Connect {0}", connection.name),
			detail: [connection.setup ?? localize('mcp.app.defaultSetup', "Add the official remote MCP server. Your browser opens for authorization when the server starts; permissions and availability are controlled by the app."), connection.blocked, connectInAssistant && !connection.blocked && localize('mcp.app.assistantProfile', "This connection will be available to Assistant in your current Signal-C profile."), httpsCallback && localize('mcp.app.httpsCallbackSetup', "If your OAuth client requires a public HTTPS callback, register {0} with the provider and choose Connect with HTTPS Callback.", httpsCallback), connection.url].filter(Boolean).join('\n\n'),
			buttons: [
				...(isNative && connection.id === 'quickbooks-local' ? [{ label: localize('mcp.app.quickBooksSandbox', "Sign In to Sandbox"), run: () => 'quickBooksSandbox' as const }] : []),
				...(!connection.blocked ? [{ label: connection.local ? localize('mcp.app.localConnect', "Connect Local Server") : localize('mcp.app.connect', "Connect with OAuth"), run: () => 'connect' as const }] : []),
				...(httpsCallback ? [{ label: localize('mcp.app.connectHttpsCallback', "Connect with HTTPS Callback"), run: () => 'https' as const }] : []),
				{ label: localize('mcp.app.guide', "Open Setup Guide"), run: () => 'guide' as const },
			],
			cancelButton: true,
		});
		if (!result) { return undefined; }
		if (result === 'guide') {
			await this._openerService.open(URI.parse(connection.documentation));
			return undefined;
		}
		if (connection.local) {
			if (this.configurationTarget) { throw new Error(localize('mcp.app.localTarget', "This app runs on the local machine and must use local user configuration. Run Signal-C: Connect App without an explicit workspace or agent-host target.")); }
			const entrypoint = await this._fileDialogService.showOpenDialog({ title: localize('mcp.app.localEntrypoint', "Select the built MCP server JavaScript entrypoint"), canSelectFiles: true, canSelectFolders: false, canSelectMany: false, availableFileSystems: [Schemas.file], filters: [{ name: 'JavaScript', extensions: ['js', 'cjs', 'mjs'] }] });
			if (entrypoint?.length !== 1 || entrypoint[0].scheme !== Schemas.file) { return undefined; }
			if (result === 'quickBooksSandbox') {
				return { app: connection, config: { type: McpServerType.LOCAL, command: connection.local.command, args: [entrypoint[0].fsPath] }, quickBooksSandbox: true };
			}
			const tokenStore = await this._fileDialogService.showOpenDialog({ title: localize('mcp.app.localTokenStore', "Select the private OAuth token-store file (its contents stay with the MCP server)"), canSelectFiles: true, canSelectFolders: false, canSelectMany: false, availableFileSystems: [Schemas.file] });
			if (tokenStore?.length !== 1 || tokenStore[0].scheme !== Schemas.file) { return undefined; }
			return { app: connection, config: { type: McpServerType.LOCAL, command: connection.local.command, args: [entrypoint[0].fsPath], env: { [connection.local.tokenStoreVariable]: tokenStore[0].fsPath } } };
		}
		if (!connection.url) { throw new Error(localize('mcp.app.missingUrl', "This app has no published MCP server endpoint.")); }
		let clientId = connection.publicClientId;
		if (connection.clientIdRequired || result === 'https') {
			clientId = await this._quickInputService.input({
				title: localize('mcp.app.clientId.title', "{0} OAuth App Client ID", connection.name),
				prompt: localize('mcp.app.clientId.prompt', "Enter your registered app's public client ID. Client secrets and account tokens belong in the MCP credential store, not this field."),
				ignoreFocusLost: true,
				validateInput: async value => !value.trim() || value.length > 2048 || /\s|[\u0000-\u001f\u007f]/.test(value.trim()) ? localize('mcp.app.clientId.invalid', "Enter a public OAuth client ID without spaces or control characters (up to 2048 characters).") : undefined,
			});
			if (!clientId?.trim()) { return undefined; }
		}
		return {
			app: connection,
			config: { type: McpServerType.REMOTE, transport: connection.transport ?? 'http', url: connection.url, ...(connection.headers ? { headers: { ...connection.headers } } : {}), ...(clientId ? { oauth: { clientId: clientId.trim(), ...(result === 'https' ? { useHttpsCallback: true } : {}) } } : {}) },
		};
	}

	private async signInToQuickBooksSandbox(): Promise<string | undefined> {
		const validateInput = async (value: string) => isQuickBooksOAuthCredential(value.trim()) ? undefined : localize('mcp.app.quickBooksCredentialInvalid', "Enter the value from the Development keys for your Intuit app.");
		const clientId = (await this._quickInputService.input({ title: localize('mcp.app.quickBooksClient', "QuickBooks Development Client ID"), ignoreFocusLost: true, validateInput }))?.trim();
		if (!clientId) { return undefined; }
		const secret = (await this._quickInputService.input({ title: localize('mcp.app.quickBooksSecret', "QuickBooks Development Client Secret"), password: true, ignoreFocusLost: true, validateInput }))?.trim();
		if (!secret) { return undefined; }
		return authorizeQuickBooksInBrowser(this._instantiationService, (host, operationId) => host.authorizeQuickBooksSandbox(clientId, secret, operationId));
	}

	private async getAppClientSecret(connection: ISignalCAppConnection | undefined, config: IMcpServerConfiguration): Promise<{ key: string; value: string } | undefined> {
		if (!connection?.supportsClientSecret || config.type !== McpServerType.REMOTE || !config.oauth?.clientId) {
			return undefined;
		}
		const key = mcpOAuthClientSecretStorageKey(URI.parse(config.url).toString(true), config.oauth.clientId);
		if (await this._secretStorageService.get(key)) {
			return undefined;
		}
		const value = (await this._quickInputService.input({
			title: localize('mcp.app.clientSecret.title', "{0} OAuth Client Secret", connection.name),
			prompt: connection.clientSecretRequired
				? localize('mcp.app.clientSecret.required', "Enter the client secret from the same OAuth app as your client ID.")
				: localize('mcp.app.clientSecret.prompt', "For a confidential app, enter its client secret. Leave blank for a public app with PKCE enabled."),
			password: true,
			ignoreFocusLost: true,
			validateInput: async value => connection.clientSecretRequired && !value.trim() ? localize('mcp.app.clientSecret.empty', "Enter the client secret to continue.") : undefined,
		}))?.trim();
		if (value === undefined || (connection.clientSecretRequired && !value)) {
			throw new CancellationError();
		}
		return value ? { key, value } : undefined;
	}

	private async getStdioConfig(): Promise<IMcpStdioServerConfiguration | undefined> {
		const command = await this._quickInputService.input({
			title: localize('mcp.command.title', "Enter Command"),
			placeHolder: localize('mcp.command.placeholder', "Command to run (with optional arguments)"),
			ignoreFocusLost: true,
		});

		if (!command) {
			return undefined;
		}

		this._telemetryService.publicLog2<AddServerData, AddServerClassification>('mcp.addserver', {
			packageType: 'stdio'
		});

		// Split command into command and args, handling quotes
		const parts = command.match(/(?:[^\s"]+|"[^"]*")+/g)!;
		return {
			type: McpServerType.LOCAL,
			command: parts[0].replace(/"/g, ''),

			args: parts.slice(1).map(arg => arg.replace(/"/g, ''))
		};
	}

	private async getSSEConfig(): Promise<IMcpRemoteServerConfiguration | undefined> {
		const url = await this._quickInputService.input({
			title: localize('mcp.url.title', "Enter Server URL"),
			placeHolder: localize('mcp.url.placeholder', "URL of the MCP server (e.g., http://localhost:3000)"),
			ignoreFocusLost: true,
		});

		if (!url) {
			return undefined;
		}

		this._telemetryService.publicLog2<AddServerData, AddServerClassification>('mcp.addserver', {
			packageType: 'sse'
		});

		return { url, type: McpServerType.REMOTE };
	}

	private async getServerId(suggestion = `my-mcp-server-${generateUuid().split('-')[0]}`): Promise<string | undefined> {
		const id = await this._quickInputService.input({
			title: localize('mcp.serverId.title', "Enter Server ID"),
			placeHolder: localize('mcp.serverId.placeholder', "Unique identifier for this server"),
			value: suggestion,
			ignoreFocusLost: true,
		});

		return id;
	}

	private async getInstallTarget(): Promise<McpInstallTarget | undefined> {
		if (this.workspaceConfigurationTarget) {
			return { kind: 'local', target: this.workspaceConfigurationTarget.folder };
		}

		const options: (IQuickPickItem & { target: McpInstallTarget })[] = [];
		const session = this.getCurrentAgentHostSession();
		if (session) {
			options.push({
				target: { kind: 'agentHost', session },
				label: localize('mcp.target.agentHost', "Agent Host"),
				description: localize('mcp.target.agentHost.description', "Available to sessions on this agent host"),
			});
		}

		options.push({
			target: { kind: 'local', target: ConfigurationTarget.USER_LOCAL },
			label: localize('mcp.target.user', "Global"),
			description: localize('mcp.target.user.description', "Available in all workspaces, runs locally"),
		});

		const raLabel = this._environmentService.remoteAuthority && this._label.getHostLabel(Schemas.vscodeRemote, this._environmentService.remoteAuthority);
		if (raLabel) {
			options.push({ target: { kind: 'local', target: ConfigurationTarget.USER_REMOTE }, label: localize('mcp.target.remote', "Remote"), description: localize('mcp.target..remote.description', "Available on this remote machine, runs on {0}", raLabel) });
		}

		const workbenchState = this._workspaceService.getWorkbenchState();
		const workspaceTargets = this._workspaceInstallTargetService.getTargets();
		const target = workbenchState === WorkbenchState.FOLDER
			? workspaceTargets.find(isWorkspaceFolder)
			: workspaceTargets.find(target => target === ConfigurationTarget.WORKSPACE);
		if (target !== undefined) {
			if (this._environmentService.remoteAuthority) {
				options.push({ target: { kind: 'local', target }, label: localize('mcp.target.workspace', "Workspace"), description: localize('mcp.target.workspace.description.remote', "Available in this workspace, runs on {0}", raLabel) });
			} else {
				options.push({ target: { kind: 'local', target }, label: localize('mcp.target.workspace', "Workspace"), description: localize('mcp.target.workspace.description', "Available in this workspace, runs locally") });
			}
		}
		if (workbenchState === WorkbenchState.WORKSPACE && this._configurationService.getValue<boolean>(mcpWorkspaceRootConfig)) {
			for (const folder of workspaceTargets.filter(isWorkspaceFolder)) {
				options.push({ target: { kind: 'local', target: folder }, label: localize('mcp.target.workspaceFolderLabel', "Workspace ({0})", folder.name), description: localize('mcp.target.workspaceFolder', "Workspace Folder") });
			}
		}

		if (options.length === 1) {
			return options[0].target;
		}

		const targetPick = await this._quickInputService.pick(options, {
			title: localize('mcp.target.title', "Add MCP Server"),
			placeHolder: localize('mcp.target.placeholder', "Select the configuration target")
		});

		return targetPick?.target;
	}

	private getCurrentAgentHostSession(): URI | undefined {
		if (this.configurationTarget && hasKey(this.configurationTarget, { sessionResource: true })) {
			return this.configurationTarget.sessionResource;
		}
		const focusedSession = this._chatWidgetService.lastFocusedWidget?.viewModel?.sessionResource;
		if (focusedSession && isAgentHostTarget(getChatSessionType(focusedSession))) {
			return focusedSession;
		}

		let session: URI | undefined;
		for (const widget of this._chatWidgetService.getAllWidgets()) {
			const candidate = widget.viewModel?.sessionResource;
			if (!candidate || !isAgentHostTarget(getChatSessionType(candidate))) {
				continue;
			}
			if (session && !isEqual(session, candidate)) {
				return undefined;
			}
			session = candidate;
		}
		return session;
	}

	private get workspaceConfigurationTarget(): IWorkspaceMcpConfigurationTarget | undefined {
		return this.configurationTarget && hasKey(this.configurationTarget, { folder: true }) ? this.configurationTarget : undefined;
	}

	private async getSetupDestination(): Promise<McpSetupDestination | undefined> {
		const installTarget = await this.getInstallTarget();
		if (!installTarget) {
			return undefined;
		}
		if (installTarget.kind === 'local') {
			if (installTarget.target === ConfigurationTarget.USER_LOCAL) {
				const resource = await this._copilotGlobalConfigurationService.getConfigurationResource();
				if (resource) {
					const selected = await this._quickInputService.pick([
						{ id: 'copilot', label: localize('mcp.target.copilotGlobal', "Copilot Global"), description: this._label.getUriLabel(resource) },
						{ id: 'vscode', label: localize('mcp.target.userConfigurationDeprecated', "User Configuration (deprecated)"), description: this._label.getUriLabel(this._userDataProfileService.currentProfile.mcpResource) },
					], {
						title: localize('mcp.target.title', "Add MCP Server"),
						placeHolder: localize('mcp.target.global.placeholder', "Select the global configuration"),
					});
					if (!selected) {
						return undefined;
					}
					if (selected.id === 'copilot') {
						return { installTarget, format: McpResourceFormat.CopilotGlobal, resource };
					}
				}
			}
		}
		return { installTarget, format: McpResourceFormat.Vscode };
	}

	private async getAssistedConfig(type: AssistedConfigurationType, format: McpResourceFormat): Promise<McpGeneratedConfiguration | undefined> {
		const packageName = await this._quickInputService.input({
			ignoreFocusLost: true,
			title: AssistedTypes[type].title,
			placeHolder: AssistedTypes[type].placeholder,
		});

		if (!packageName) {
			return undefined;
		}

		const enum LoadAction {
			Retry = 'retry',
			Cancel = 'cancel',
			Allow = 'allow',
			OpenUri = 'openUri',
		}

		const loadingQuickPickStore = new DisposableStore();
		const loadingQuickPick = loadingQuickPickStore.add(this._quickInputService.createQuickPick<IQuickPickItem & { id: LoadAction; helpUri?: URI }>());
		loadingQuickPick.title = localize('mcp.loading.title', "Loading package details...");
		loadingQuickPick.busy = true;
		loadingQuickPick.ignoreFocusOut = true;

		const packageType = this.getPackageType(type);

		this._telemetryService.publicLog2<AddServerData, AddServerClassification>('mcp.addserver', {
			packageType: packageType!
		});

		this._commandService.executeCommand<ValidatePackageResult>(
			AddConfigurationCopilotCommand.ValidatePackage,
			{
				type: packageType,
				name: packageName,
				targetFormat: format,
				targetConfig: getMcpGenerationSchema(format),
			}
		).then(result => {
			if (!result || result.state === 'error') {
				loadingQuickPick.title = result?.error || 'Unknown error loading package';

				const items: Array<IQuickPickItem & { id: LoadAction; helpUri?: URI }> = [];

				if (result?.helpUri) {
					items.push({
						id: LoadAction.OpenUri,
						label: result.helpUriLabel ?? localize('mcp.error.openHelpUri', 'Open help URL'),
						helpUri: URI.parse(result.helpUri),
					});
				}

				items.push(
					{ id: LoadAction.Retry, label: localize('mcp.error.retry', 'Try a different package') },
					{ id: LoadAction.Cancel, label: localize('cancel', 'Cancel') },
				);

				loadingQuickPick.items = items;
			} else {
				loadingQuickPick.title = localize(
					'mcp.confirmPublish', 'Install {0}{1} from {2}?',
					result.name ?? packageName,
					result.version ? `@${result.version}` : '',
					result.publisher);
				loadingQuickPick.items = [
					{ id: LoadAction.Allow, label: localize('allow', "Allow") },
					{ id: LoadAction.Cancel, label: localize('cancel', 'Cancel') }
				];
			}
			loadingQuickPick.busy = false;
		}, () => {
			loadingQuickPick.title = localize('mcp.package.loadFailed', "Unable to load MCP package details. Try again.");
			loadingQuickPick.items = [{ id: LoadAction.Retry, label: localize('mcp.error.retry', 'Try a different package') }, { id: LoadAction.Cancel, label: localize('cancel', 'Cancel') }];
			loadingQuickPick.busy = false;
		});

		const loadingAction = await new Promise<{ id: LoadAction; helpUri?: URI } | undefined>(resolve => {
			loadingQuickPickStore.add(loadingQuickPick.onDidAccept(() => resolve(loadingQuickPick.selectedItems[0])));
			loadingQuickPickStore.add(loadingQuickPick.onDidHide(() => resolve(undefined)));
			loadingQuickPick.show();
		}).finally(() => loadingQuickPickStore.dispose());

		switch (loadingAction?.id) {
			case LoadAction.Retry:
				return this.getAssistedConfig(type, format);
			case LoadAction.OpenUri:
				if (loadingAction.helpUri) { this._openerService.open(loadingAction.helpUri); }
				return undefined;
			case LoadAction.Allow:
				break;
			case LoadAction.Cancel:
			default:
				return undefined;
		}

		const config = await this._commandService.executeCommand<McpGeneratedConfiguration | undefined>(
			AddConfigurationCopilotCommand.StartFlow,
			{
				name: packageName,
				type: packageType
			}
		);

		return config;
	}

	/** Starts an installed server once it's discovered and optionally reveals its configuration. */
	private showOnceDiscovered(installed: IWorkbenchLocalMcpServer, revealConfiguration = true) {
		const store = new DisposableStore();
		store.add(disposableTimeout(() => store.dispose(), 5000));
		const observer = autorunSelfDisposable(reader => {
			const servers = this._mcpService.servers.read(reader);
			const server = servers.find(s => s.definition.id === installed.id);

			if (server) {
				if (revealConfiguration) {
					const collection = this._mcpRegistry.collections.read(reader).find(collection => collection.id === server.collection.id);
					const definition = collection?.serverDefinitions.read(reader).find(definition => definition.id === installed.id);
					this._editorService.openEditor({
						resource: installed.mcpResource,
						options: {
							selection: definition?.presentation?.origin?.range,
							preserveFocus: true,
						}
					});
				}

				server.start({ promptType: 'all-untrusted' }).then(state => {
					if (state.state === McpConnectionState.Kind.Error) {
						server.showOutput();
					}
				});

				reader.dispose();
				store.dispose();
			}
		});
		if (!store.isDisposed) {
			store.add(observer);
		}
	}

	public async run(connectApp = false, appId?: string, connectInAssistant = false): Promise<void> {
		// Step 1: Choose server type
		const serverType = connectApp ? AddConfigurationType.App : await this.getServerType();
		if (serverType === undefined) {
			return;
		}

		// Step 2: Get server details based on type
		let config: IMcpServerConfiguration | undefined;
		let connection: ISignalCAppConnection | undefined;
		let suggestedName: string | undefined;
		let inputs: IMcpServerVariable[] | undefined;
		let inputValues: Record<string, string> | undefined;
		let destination: McpSetupDestination | undefined;
		let quickBooksSandbox = false;
		switch (serverType) {
			case AddConfigurationType.App: {
				const details = await this.getAppConfig(connectApp ? appId : undefined, connectInAssistant);
				config = details?.config;
				connection = details?.app;
				quickBooksSandbox = details?.quickBooksSandbox === true;
				suggestedName = connection?.id;
				if (connection?.local || connectInAssistant) { destination = { installTarget: { kind: 'local', target: ConfigurationTarget.USER_LOCAL }, format: McpResourceFormat.Vscode }; }
				break;
			}
			case AddConfigurationType.Stdio:
				config = await this.getStdioConfig();
				break;
			case AddConfigurationType.HTTP:
				config = await this.getSSEConfig();
				break;
			case AddConfigurationType.NpmPackage:
			case AddConfigurationType.PipPackage:
			case AddConfigurationType.NuGetPackage:
			case AddConfigurationType.DockerImage: {
				destination = await this.getSetupDestination();
				if (!destination) {
					return;
				}
				const installTarget = destination.installTarget;
				if (installTarget.kind === 'local' && isWorkspaceFolder(installTarget.target)) {
					const workspaceConfig = await this._instantiationService.createInstance(McpConfigurationDestination).selectForAdd(installTarget.target, undefined, this.workspaceConfigurationTarget?.kind);
					if (workspaceConfig === undefined) {
						return;
					}
					destination.workspaceConfig = workspaceConfig;
					destination.format = workspaceConfig === WorkspaceMcpConfigKind.Root ? McpResourceFormat.WorkspaceRoot : McpResourceFormat.Vscode;
				}
				const r = await this.getAssistedConfig(serverType, destination.format);
				config = r ? normalizeMcpGeneratedConfiguration(r, destination.format) : undefined;
				suggestedName = r?.name;
				inputs = r?.inputs;
				inputValues = r?.inputValues;
				break;
			}
			default:
				assertNever(serverType);
		}

		if (!config) {
			return;
		}

		// Step 3: Get server ID
		// A second app connection should not silently replace a same-named manual or account connection.
		const suggestion = serverType === AddConfigurationType.App && suggestedName ? `${suggestedName}-${generateUuid().slice(0, 8)}` : suggestedName;
		const name = await this.getServerId(suggestion);
		if (!name) {
			return;
		}

		// Step 4: Choose configuration target
		destination ??= await this.getSetupDestination();
		if (!destination) {
			return;
		}
		const { installTarget } = destination;

		if (installTarget.kind === 'agentHost') {
			if (Object.hasOwn(AssistedTypes, serverType) && (inputs?.length || JSON.stringify(config).includes('${input:'))) {
				throw new Error(localize('mcp.agentHost.inputsUnsupported', "This server requires VS Code input variables. Add it to a VS Code configuration file instead of the current agent session."));
			}
			const allowed = this._allowedMcpServersService.isAllowed({ name, config, inputs });
			if (allowed !== true) {
				throw new Error(allowed.value);
			}
			this._agentHostCustomizations.addMcpServer(installTarget.session, name, config);
			return;
		}

		const { target } = installTarget;
		if (quickBooksSandbox && config.type === McpServerType.LOCAL && connection?.local) {
			const allowed = this._allowedMcpServersService.isAllowed({ name, config, inputs });
			if (allowed !== true) { throw new Error(allowed.value); }
			const tokenStore = await this.signInToQuickBooksSandbox();
			if (!tokenStore) { return; }
			config = { ...config, env: { ...config.env, [connection.local.tokenStoreVariable]: tokenStore } };
		}
		const installable = { name, config, inputs };
		if (destination.format === McpResourceFormat.CopilotGlobal && destination.resource) {
			const resource = destination.resource;
			const error = getCopilotGlobalMcpConfigurationError(installable);
			if (error) {
				if (!Object.hasOwn(AssistedTypes, serverType)) {
					throw new Error(error);
				}
				const fallback = await this._quickInputService.pick([{
					label: localize('mcp.target.userConfigurationDeprecated', "User Configuration (deprecated)"),
					detail: error,
				}], { placeHolder: localize('mcp.target.vscodeRequired', "This server requires VS Code configuration"), ignoreFocusLost: true });
				if (!fallback) {
					return;
				}
			} else {
				const allowed = this._allowedMcpServersService.isAllowed(installable);
				if (allowed !== true) {
					throw new Error(allowed.value);
				}
				const credential = await this.getAppClientSecret(connection, config);
				const configurationProvider = this._customizationHarnessService.getActiveDescriptor().mcpConfigurationProvider;
				if (configurationProvider) {
					await configurationProvider.addGlobal(name, config);
				} else {
					await this._mcpResourceScannerService.addMcpServers([installable], resource, undefined, McpResourceFormat.CopilotGlobal);
				}
				if (credential) {
					await this._secretStorageService.set(credential.key, credential.value);
				}
				await this._editorService.openEditor({ resource });
				this._notificationService.info(localize('mcp.copilotGlobal.added', "Added MCP server '{0}' to {1}. Set any referenced environment variables on the agent-host machine before starting a new Copilot session. VS Code input variables are not supported in this file.", name, this._label.getUriLabel(resource)));
				return;
			}
		}
		const workspaceConfig = isWorkspaceFolder(target)
			? destination.workspaceConfig === undefined
				? await this._instantiationService.createInstance(McpConfigurationDestination).selectForAdd(target, installable, this.workspaceConfigurationTarget?.kind)
				: await this._instantiationService.createInstance(McpConfigurationDestination).checkGenerated(target, installable, destination.workspaceConfig, !!this.workspaceConfigurationTarget)
			: undefined;
		if (isWorkspaceFolder(target) && workspaceConfig === undefined) {
			return;
		}
		const credential = await this.getAppClientSecret(connection, config);
		const installed = await this._mcpManagementService.install(installable, { target, workspaceConfig });
		if (credential) {
			await this._secretStorageService.set(credential.key, credential.value);
		}

		if (inputValues) {
			for (const [key, value] of Object.entries(inputValues)) {
				await this._mcpRegistry.setSavedInput(key, (isWorkspaceFolder(target) ? ConfigurationTarget.WORKSPACE_FOLDER : target) ?? ConfigurationTarget.WORKSPACE, value);
			}
		}

		const packageType = this.getPackageType(serverType);
		if (packageType) {
			this._telemetryService.publicLog2<AddServerCompletedData, AddServerCompletedClassification>('mcp.addserver.completed', {
				packageType,
				serverType: config.type,
				target: target === ConfigurationTarget.WORKSPACE ? 'workspace' : 'user'
			});
		}

		this.showOnceDiscovered(installed, !connectInAssistant);
	}

	public async pickForUrlHandler(resource: URI, showIsPrimary = false): Promise<void> {
		const name = decodeURIComponent(basename(resource)).replace(/\.json$/, '');
		const placeHolder = localize('install.title', 'Install MCP server {0}', name);

		const items: IQuickPickItem[] = [
			{ id: 'install', label: localize('install.start', 'Install Server') },
			{ id: 'show', label: localize('install.show', 'Show Configuration', name) },
			{ id: 'rename', label: localize('install.rename', 'Rename "{0}"', name) },
			{ id: 'cancel', label: localize('cancel', 'Cancel') },
		];
		if (showIsPrimary) {
			[items[0], items[1]] = [items[1], items[0]];
		}

		const pick = await this._quickInputService.pick(items, { placeHolder, ignoreFocusLost: true });
		const getEditors = () => this._editorService.findEditors(resource);

		switch (pick?.id) {
			case 'show':
				await this._editorService.openEditor({ resource });
				break;
			case 'install':
				await this._editorService.save(getEditors());
				try {
					const contents = await this._fileService.readFile(resource);
					const { inputs, ...config }: IMcpServerConfiguration & { inputs?: IMcpServerVariable[] } = parseJsonc(contents.value.toString());
					const installed = await this._mcpManagementService.install({ name, config, inputs });
					this._editorService.closeEditors(getEditors());
					this.showOnceDiscovered(installed);
				} catch (e) {
					this._notificationService.error(localize('install.error', 'Error installing MCP server {0}: {1}', name, e.message));
					await this._editorService.openEditor({ resource });
				}
				break;
			case 'rename': {
				const newName = await this._quickInputService.input({ placeHolder: localize('install.newName', 'Enter new name'), value: name });
				if (newName) {
					const newURI = resource.with({ path: `/${encodeURIComponent(newName)}.json` });
					await this._editorService.save(getEditors());
					await this._fileService.move(resource, newURI);
					return this.pickForUrlHandler(newURI, showIsPrimary);
				}
				break;
			}
		}
	}

	private getPackageType(serverType: AddConfigurationType): string | undefined {
		switch (serverType) {
			case AddConfigurationType.NpmPackage:
				return 'npm';
			case AddConfigurationType.PipPackage:
				return 'pip';
			case AddConfigurationType.NuGetPackage:
				return 'nuget';
			case AddConfigurationType.DockerImage:
				return 'docker';
			case AddConfigurationType.Stdio:
				return 'stdio';
			case AddConfigurationType.HTTP:
				return 'sse';
			default:
				return undefined;
		}
	}
}

export class McpInstallFromManifestCommand {
	constructor(
		@IFileDialogService private readonly _fileDialogService: IFileDialogService,
		@IFileService private readonly _fileService: IFileService,
		@IQuickInputService private readonly _quickInputService: IQuickInputService,
		@INotificationService private readonly _notificationService: INotificationService,
		@IWorkbenchMcpManagementService private readonly _mcpManagementService: IWorkbenchMcpManagementService,
		@ILogService private readonly _logService: ILogService,
	) { }

	async run(): Promise<void> {
		// Step 1: Open file dialog to select the manifest file
		const result = await this._fileDialogService.showOpenDialog({
			title: localize('mcp.installFromManifest.title', "Select MCP Server Manifest"),
			filters: [{ name: localize('mcp.installFromManifest.filter', "MCP Manifest"), extensions: ['json'] }],
			canSelectFiles: true,
			canSelectMany: false,
			openLabel: mnemonicButtonLabel(localize({ key: 'mcp.installFromManifest.openLabel', comment: ['&& denotes a mnemonic'] }, "&&Install"))
		});

		if (!result?.[0]) {
			return;
		}

		const manifestUri = result[0];

		// Step 2: Read and parse the manifest file
		let manifest: unknown;
		try {
			const contents = await this._fileService.readFile(manifestUri);
			manifest = parseJsonc(contents.value.toString());
		} catch (e) {
			this._notificationService.error(localize('mcp.installFromManifest.readError', "Failed to read manifest file: {0}", e.message));
			return;
		}

		if (!manifest || typeof manifest !== 'object') {
			this._notificationService.error(localize('mcp.installFromManifest.invalidJson', "Invalid manifest file: expected a JSON object"));
			return;
		}

		// Step 3: Validate and extract configuration from gallery manifest
		const galleryManifest = manifest as IGalleryMcpServerConfiguration & { name?: string };

		// Determine package type from manifest
		let packageType: RegistryType;
		if (Array.isArray(galleryManifest.packages) && galleryManifest.packages.length > 0) {
			packageType = galleryManifest.packages[0].registryType;
		} else if (Array.isArray(galleryManifest.remotes) && galleryManifest.remotes.length > 0) {
			packageType = RegistryType.REMOTE;
		} else {
			this._notificationService.error(localize('mcp.installFromManifest.invalidManifest', "Invalid manifest: expected 'packages' or 'remotes' with at least one entry"));
			return;
		}

		let config: IMcpServerConfiguration;
		let inputs: IMcpServerVariable[] | undefined;
		try {
			const { mcpServerConfiguration, notices } = this._mcpManagementService.getMcpServerConfigurationFromManifest(galleryManifest, packageType);
			config = mcpServerConfiguration.config;
			inputs = mcpServerConfiguration.inputs;

			if (notices.length > 0) {
				this._logService.warn(`MCP Management Service: Warnings while installing the MCP server from ${manifestUri.path}`, notices);
			}
		} catch (e) {
			this._notificationService.error(localize('mcp.installFromManifest.parseError', "Failed to parse manifest: {0}", e.message));
			return;
		}

		// Step 4: Get server name from manifest or prompt user
		let name = galleryManifest.name;
		if (!name) {
			name = await this._quickInputService.input({
				title: localize('mcp.installFromManifest.serverId.title', "Enter Server ID"),
				placeHolder: localize('mcp.installFromManifest.serverId.placeholder', "Unique identifier for this server"),
				value: basename(manifestUri).replace(/\.json$/i, ''),
				ignoreFocusLost: true,
			});

			if (!name) {
				return;
			}
		}

		// Step 5: Install to user settings
		try {
			await this._mcpManagementService.install({ name, config, inputs });
			this._notificationService.info(localize('mcp.installFromManifest.success', "MCP server '{0}' installed successfully", name));
		} catch (e) {
			this._notificationService.error(localize('mcp.installFromManifest.installError', "Failed to install MCP server: {0}", e.message));
		}
	}
}
