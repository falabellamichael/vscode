/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import * as assert from 'assert';
import { DeferredPromise } from '../../../../base/common/async.js';
import { isCancellationError } from '../../../../base/common/errors.js';
import { Emitter, Event } from '../../../../base/common/event.js';
import { DisposableStore } from '../../../../base/common/lifecycle.js';
import { observableValue } from '../../../../base/common/observable.js';
import { URI } from '../../../../base/common/uri.js';
import { mock, upcastPartial } from '../../../../base/test/common/mock.js';
import { ensureNoDisposablesAreLeakedInTestSuite } from '../../../../base/test/common/utils.js';
import { ConfigurationTarget, IConfigurationService } from '../../../../platform/configuration/common/configuration.js';
import { TestConfigurationService } from '../../../../platform/configuration/test/common/testConfigurationService.js';
import { IContextKeyService } from '../../../../platform/contextkey/common/contextkey.js';
import { IDialogService } from '../../../../platform/dialogs/common/dialogs.js';
import { TestInstantiationService } from '../../../../platform/instantiation/test/common/instantiationServiceMock.js';
import { NullLogService } from '../../../../platform/log/common/log.js';
import { IMcpServerMatcher } from '../../../../platform/mcp/common/allowedMcpServers.js';
import { AllowedMcpServersService } from '../../../../platform/mcp/common/allowedMcpServersService.js';
import { IAllowedMcpServersService, mcpAccessConfig, mcpAllowedServersConfig, mcpDeniedServersConfig, McpAccessValue } from '../../../../platform/mcp/common/mcpManagement.js';
import { IQuickInputService } from '../../../../platform/quickinput/common/quickInput.js';
import { ISecretStorageService } from '../../../../platform/secrets/common/secrets.js';
import { StorageScope } from '../../../../platform/storage/common/storage.js';
import { ITelemetryService } from '../../../../platform/telemetry/common/telemetry.js';
import { IWorkbenchMcpGatewayService } from '../../../contrib/mcp/common/mcpGatewayService.js';
import { IMcpHostDelegate, IMcpRegistry } from '../../../contrib/mcp/common/mcpRegistryTypes.js';
import { McpCollectionDefinition, McpCollectionSortOrder, McpConnectionState, McpServerDefinition, McpServerLaunch, McpServerTransportType, McpServerTrust, mcpOAuthClientSecretStorageKey } from '../../../contrib/mcp/common/mcpTypes.js';
import { mcpEnterpriseManagedAuthIdpSection } from '../../../contrib/mcp/common/mcpConfiguration.js';
import { MCP } from '../../../contrib/mcp/common/modelContextProtocol.js';
import { IAuthenticationMcpAccessService } from '../../../services/authentication/browser/authenticationMcpAccessService.js';
import { IAuthenticationMcpService } from '../../../services/authentication/browser/authenticationMcpService.js';
import { IAuthenticationMcpUsageService } from '../../../services/authentication/browser/authenticationMcpUsageService.js';
import { AuthenticationSession, AuthenticationSessionsChangeEvent, getDynamicAuthenticationProviderId, getOAuthTokenFingerprint, IAuthenticationGetSessionsOptions, IAuthenticationProvider, IAuthenticationService, IAuthenticationWwwAuthenticateRequest, MCP_REFRESH_REJECTED_OAUTH_TOKEN, MCP_USE_HTTPS_OAUTH_CALLBACK } from '../../../services/authentication/common/authentication.js';
import { IDynamicAuthenticationProviderStorageService } from '../../../services/authentication/common/dynamicAuthenticationProviderStorage.js';
import { IExtensionService } from '../../../services/extensions/common/extensions.js';
import { TestExtensionService } from '../../../test/common/workbenchTestServices.js';
import { IMcpServerAuthContext, MainThreadMcp, McpServerAuthTracker } from '../../browser/mainThreadMcp.js';
import { ExtHostMcpShape, IMcpAuthenticationDetails } from '../../common/extHost.protocol.js';
import { CommonRequestInit, CommonResponse, McpHTTPHandle } from '../../common/extHostMcp.js';
import { createEnvVarResolver } from '../../../../platform/mcp/common/envVarResolution.js';
import { SingleProxyRPCProtocol } from '../common/testRPCProtocol.js';

function createMainThreadMcp(disposables: Pick<DisposableStore, 'add'>, proxy: Partial<ExtHostMcpShape>, registry: IMcpRegistry, configure?: (services: TestInstantiationService) => void): MainThreadMcp {
	const services = disposables.add(new TestInstantiationService());
	services.stub(IMcpRegistry, registry);
	services.stub(IDialogService, {});
	services.stub(IAuthenticationService, { onDidChangeSessions: Event.None });
	services.stub(IAuthenticationMcpService, {});
	services.stub(IAuthenticationMcpAccessService, {});
	services.stub(IAuthenticationMcpUsageService, {});
	services.stub(IDynamicAuthenticationProviderStorageService, {});
	services.stub(IExtensionService, new TestExtensionService());
	services.stub(IContextKeyService, {});
	services.stub(ITelemetryService, {});
	services.stub(IWorkbenchMcpGatewayService, {});
	services.stub(IConfigurationService, {});
	services.stub(ISecretStorageService, {});
	services.stub(IQuickInputService, {});
	services.stub(IAllowedMcpServersService, {
		onDidChangeAllowedMcpServers: Event.None,
		isServerAllowed: () => true,
		getServerPolicyVariableValues: () => [],
	});
	configure?.(services);
	return disposables.add(services.createInstance(MainThreadMcp, SingleProxyRPCProtocol(proxy)));
}

suite('MainThreadMcp - McpServerAuthTracker', () => {

	ensureNoDisposablesAreLeakedInTestSuite();

	/**
	 * Builds a representative auth context. Defaults model a tenant-specific Entra sign-in so tests
	 * assert that the authorization server / resource survive tracking (the values that were dropped
	 * on re-validation in #324925).
	 */
	function ctx(overrides: Partial<IMcpServerAuthContext> = {}): IMcpServerAuthContext {
		return {
			authorizationServer: URI.parse('https://login.microsoftonline.com/11111111-1111-1111-1111-111111111111'),
			clientId: 'client-a',
			resource: 'api://33333333-3333-3333-3333-333333333333',
			audience: undefined,
			...overrides,
		};
	}

	test('retains the full auth context per tracked server and groups by provider', () => {
		const tracker = new McpServerAuthTracker();
		const first = ctx();
		const second = ctx({
			clientId: 'client-b',
			authorizationServer: URI.parse('https://login.microsoftonline.com/22222222-2222-2222-2222-222222222222'),
		});

		tracker.track('microsoft', 1, ['scope.a'], first);
		tracker.track('microsoft', 2, ['scope.b'], second);

		assert.deepStrictEqual(tracker.get('microsoft'), [
			{ serverId: 1, scopes: ['scope.a'], context: first },
			{ serverId: 2, scopes: ['scope.b'], context: second },
		]);
	});

	test('re-tracking the same server replaces its context without duplicating', () => {
		const tracker = new McpServerAuthTracker();
		tracker.track('microsoft', 1, ['scope.a'], ctx());
		const updated = ctx({ authorizationServer: URI.parse('https://login.microsoftonline.com/44444444-4444-4444-4444-444444444444') });
		tracker.track('microsoft', 1, ['scope.a'], updated);

		assert.deepStrictEqual(tracker.get('microsoft'), [
			{ serverId: 1, scopes: ['scope.a'], context: updated },
		]);
	});

	test('untrack removes a server across every provider and drops empty provider buckets', () => {
		const tracker = new McpServerAuthTracker();
		tracker.track('microsoft', 1, ['scope.a'], ctx());
		tracker.track('github', 1, ['repo'], ctx({ authorizationServer: undefined, resource: undefined }));
		tracker.track('microsoft', 2, ['scope.b'], ctx());

		tracker.untrack(1);

		assert.strictEqual(tracker.get('github'), undefined, 'empty provider bucket is removed');
		assert.deepStrictEqual(tracker.get('microsoft'), [
			{ serverId: 2, scopes: ['scope.b'], context: ctx() },
		]);
	});

	test('clear removes all tracking', () => {
		const tracker = new McpServerAuthTracker();
		tracker.track('microsoft', 1, ['scope.a'], ctx());
		tracker.clear();

		assert.strictEqual(tracker.get('microsoft'), undefined);
	});
});

suite('MainThreadMcp - launch', () => {

	const disposables = ensureNoDisposablesAreLeakedInTestSuite();

	function createLaunchFixture(proxy: Partial<ExtHostMcpShape>) {
		let capturedDelegate: IMcpHostDelegate | undefined;
		const registry = new class extends mock<IMcpRegistry>() {
			override readonly collections = observableValue<readonly McpCollectionDefinition[]>('collections', []);
			override registerDelegate(delegate: IMcpHostDelegate) {
				capturedDelegate = delegate;
				return { dispose() { } };
			}
		};
		const mainThread = createMainThreadMcp(disposables, {
			$onDidChangeMcpServerDefinitions() { },
			...proxy,
		}, registry);
		const launch: McpServerLaunch = { type: McpServerTransportType.HTTP, uri: URI.parse('https://server.example/mcp'), headers: [] };
		const definition: McpServerDefinition = { id: 'server', label: 'Server', launch, cacheNonce: 'nonce' };
		const collection: McpCollectionDefinition = {
			id: 'collection',
			label: 'Collection',
			remoteAuthority: null,
			serverDefinitions: observableValue<readonly McpServerDefinition[]>('definitions', [definition]),
			trustBehavior: McpServerTrust.Kind.Trusted,
			scope: StorageScope.WORKSPACE,
			configTarget: ConfigurationTarget.USER,
			order: McpCollectionSortOrder.WorkspaceFolder,
		};
		assert.ok(capturedDelegate);
		const delegate = capturedDelegate;
		return { mainThread, start: () => delegate.start(collection, definition, launch) };
	}

	test('routes outgoing messages and stops to the correct extension-host server', () => {
		const calls: { method: 'start' | 'send' | 'stop'; id: number; message?: string }[] = [];
		const fixture = createLaunchFixture({
			$startMcp(id) {
				calls.push({ method: 'start', id });
			},
			$sendMessage(id, message) {
				calls.push({ method: 'send', id, message });
			},
			$stopMcp(id) {
				calls.push({ method: 'stop', id });
			},
		});
		const first = fixture.start();
		const second = fixture.start();
		const request: MCP.JSONRPCMessage = { jsonrpc: '2.0', id: 7, method: 'ping' };
		const notification: MCP.JSONRPCMessage = { jsonrpc: '2.0', method: 'notifications/initialized' };

		first.send(request);
		second.send(notification);
		second.stop();
		first.stop();

		assert.deepStrictEqual(calls, [
			{ method: 'start', id: 1 },
			{ method: 'start', id: 2 },
			{ method: 'send', id: 1, message: '{"jsonrpc":"2.0","id":7,"method":"ping"}' },
			{ method: 'send', id: 2, message: '{"jsonrpc":"2.0","method":"notifications/initialized"}' },
			{ method: 'stop', id: 2 },
			{ method: 'stop', id: 1 },
		]);
	});

	test('only stops active transports on disposal', () => {
		const states: McpConnectionState[] = [
			{ state: McpConnectionState.Kind.Starting },
			{ state: McpConnectionState.Kind.Running },
			{ state: McpConnectionState.Kind.Stopped },
			{ state: McpConnectionState.Kind.Error, message: 'Connection failed' },
		];
		const results = states.map(state => {
			const stopped: number[] = [];
			const fixture = createLaunchFixture({
				$startMcp() { },
				$stopMcp(id) {
					stopped.push(id);
				},
			});
			const transport = fixture.start();
			fixture.mainThread.$onDidChangeState(1, state);
			transport.dispose();
			fixture.mainThread.dispose();
			return { state: state.state, stopped };
		});

		assert.deepStrictEqual(results, [
			{ state: McpConnectionState.Kind.Starting, stopped: [1] },
			{ state: McpConnectionState.Kind.Running, stopped: [1] },
			{ state: McpConnectionState.Kind.Stopped, stopped: [] },
			{ state: McpConnectionState.Kind.Error, stopped: [] },
		]);
	});

	test('extension-host shutdown stops local state without sending another stop command', () => {
		const stopped: number[] = [];
		const fixture = createLaunchFixture({
			$startMcp() { },
			$stopMcp(id) {
				stopped.push(id);
			},
		});
		const transport = fixture.start();
		fixture.mainThread.$onDidChangeState(1, { state: McpConnectionState.Kind.Running });

		fixture.mainThread.dispose();

		assert.deepStrictEqual({ stopped, state: transport.state.get() }, {
			stopped: [],
			state: { state: McpConnectionState.Kind.Stopped },
		});
	});

	test('forwards the definition default cwd to the extension host', () => {
		let startOptions: Parameters<ExtHostMcpShape['$startMcp']>[1] | undefined;
		const proxy: Partial<ExtHostMcpShape> = {
			$startMcp(_id, options) {
				startOptions = options;
			},
			$stopMcp() { },
			$sendMessage() { },
			$onDidChangeMcpServerDefinitions() { },
		};

		let capturedDelegate: IMcpHostDelegate | undefined;
		const mcpRegistry = new class extends mock<IMcpRegistry>() {
			override readonly collections = observableValue<readonly McpCollectionDefinition[]>('collections', []);
			override registerDelegate(delegate: IMcpHostDelegate) {
				capturedDelegate = delegate;
				return { dispose() { } };
			}
		};

		createMainThreadMcp(disposables, proxy, mcpRegistry);

		const launch: McpServerLaunch = {
			type: McpServerTransportType.HTTP,
			uri: URI.parse('https://myserver.example/mcp'),
			headers: [],
		};
		const defaultCwd = URI.parse('vscode-remote://ssh-remote+linux/home/test/workspace');
		const serverDefinition: McpServerDefinition = {
			id: 'my-server',
			label: 'My Server',
			launch,
			defaultCwd,
			cacheNonce: 'nonce-1',
			variableReplacement: {
				folder: { uri: URI.file('/fallback'), name: 'fallback', index: 0 },
				target: ConfigurationTarget.WORKSPACE_FOLDER,
			},
		};
		const collection: McpCollectionDefinition = {
			remoteAuthority: 'ssh-remote+linux',
			id: 'collection-1',
			label: 'Collection',
			serverDefinitions: observableValue<readonly McpServerDefinition[]>('serverDefinitions', [serverDefinition]),
			trustBehavior: McpServerTrust.Kind.Trusted,
			scope: StorageScope.WORKSPACE,
			configTarget: ConfigurationTarget.WORKSPACE_FOLDER,
			order: McpCollectionSortOrder.WorkspaceFolder,
		};

		assert.ok(capturedDelegate);
		capturedDelegate.start(collection, serverDefinition, launch, {});
		assert.ok(startOptions?.defaultCwd);
		assert.strictEqual(URI.revive(startOptions.defaultCwd).toString(), defaultCwd.toString());
	});

	test('expands environment variables in the extension host only for flagged definitions', async () => {
		const requests: Parameters<ExtHostMcpShape['$expandEnvironmentVariables']>[] = [];
		const proxy: Partial<ExtHostMcpShape> = {
			async $expandEnvironmentVariables(launch, expansion) {
				requests.push([launch, expansion]);
				return { ...launch, uri: URI.parse('https://expanded.example/mcp').toJSON(), headers: [['Authorization', 'Bearer expanded']] } as McpServerLaunch.Serialized;
			},
			$onDidChangeMcpServerDefinitions() { },
		};

		let capturedDelegate: IMcpHostDelegate | undefined;
		const mcpRegistry = new class extends mock<IMcpRegistry>() {
			override readonly collections = observableValue<readonly McpCollectionDefinition[]>('collections', []);
			override registerDelegate(delegate: IMcpHostDelegate) {
				capturedDelegate = delegate;
				return { dispose() { } };
			}
		};
		createMainThreadMcp(disposables, proxy, mcpRegistry);
		assert.ok(capturedDelegate);

		const url = 'https://${HOST}/mcp';
		const launch: McpServerLaunch = { type: McpServerTransportType.HTTP, uri: URI.parse(url), headers: [['Authorization', 'Bearer ${TOKEN}']] };
		const definition: McpServerDefinition = { id: 'server', label: 'Server', launch, cacheNonce: 'nonce' };
		const unflagged = await capturedDelegate.expandEnvironmentVariables(definition, launch);
		const expanded = await capturedDelegate.expandEnvironmentVariables({ ...definition, environmentVariableExpansion: { url } }, launch);

		assert.deepStrictEqual({
			unflaggedIsSame: unflagged === launch,
			requests: requests.map(([, expansion]) => expansion),
			expanded: expanded.type === McpServerTransportType.HTTP ? { isUri: URI.isUri(expanded.uri), url: expanded.uri.toString(true), headers: expanded.headers } : undefined,
		}, {
			unflaggedIsSame: true,
			requests: [{ url }],
			expanded: { isUri: true, url: 'https://expanded.example/mcp', headers: [['Authorization', 'Bearer expanded']] },
		});
	});
});

suite('MainThreadMcp - re-validation', () => {

	const disposables = ensureNoDisposablesAreLeakedInTestSuite();

	async function assertRevalidationContext(enterpriseManaged: boolean, clientId: string | undefined): Promise<void> {
		const authorizationServer = URI.parse('https://login.microsoftonline.com/72f988bf-86f1-41af-91ab-2d7cd011db47');
		const configuredIssuer = URI.parse('https://sso.example/issuer');
		const resource = 'https://resource.example/mcp';
		const serverUrl = 'https://myserver.example/mcp';
		const audience = enterpriseManaged ? 'https://resource-as.example' : undefined;
		const secretStorageKey = clientId ? mcpOAuthClientSecretStorageKey(enterpriseManaged ? resource : serverUrl, clientId) : undefined;
		const secrets = new Map<string, string>();
		if (secretStorageKey) {
			secrets.set(secretStorageKey, 'original-secret');
		}
		const secretStorageReads: string[] = [];
		const session: AuthenticationSession = {
			id: 'session-1',
			accessToken: 'access-token',
			account: { id: 'account-1', label: 'user@contoso.com' },
			scopes: ['scope.read'],
		};

		const getSessionsOptions: Array<IAuthenticationGetSessionsOptions | undefined> = [];
		let revalidated = new DeferredPromise<void>();

		const onDidChangeSessions = disposables.add(new Emitter<{ providerId: string; label: string; event: AuthenticationSessionsChangeEvent }>());

		const authenticationService = new class extends mock<IAuthenticationService>() {
			override readonly onDidChangeSessions = onDidChangeSessions.event;
			override async getOrActivateProviderIdForServer(): Promise<string | undefined> {
				return 'test-provider';
			}
			override async createOrGetXaaProvider(): Promise<string | undefined> {
				return 'test-provider';
			}
			override isDynamicAuthenticationProvider(): boolean {
				return false;
			}
			override isAuthenticationProviderRegistered(id: string): boolean {
				return id === 'test-provider';
			}
			override getProvider(id: string): IAuthenticationProvider {
				return new class extends mock<IAuthenticationProvider>() {
					override readonly id = id;
					override readonly label = 'Test Provider';
					override readonly supportsMultipleAccounts = false;
				};
			}
			override async getSessions(_id: string, _scopes?: ReadonlyArray<string> | IAuthenticationWwwAuthenticateRequest, options?: IAuthenticationGetSessionsOptions): Promise<ReadonlyArray<AuthenticationSession>> {
				getSessionsOptions.push(options);
				if (getSessionsOptions.length > 1) {
					revalidated.complete();
				}
				return [session];
			}
		};

		const proxy: Partial<ExtHostMcpShape> = {
			$startMcp() { },
			$stopMcp() { },
			$sendMessage() { },
			$onDidChangeMcpServerDefinitions() { },
		};

		let capturedDelegate: IMcpHostDelegate | undefined;
		const mcpRegistry = new class extends mock<IMcpRegistry>() {
			override readonly collections = observableValue<readonly McpCollectionDefinition[]>('collections', []);
			override registerDelegate(delegate: IMcpHostDelegate) {
				capturedDelegate = delegate;
				return { dispose() { } };
			}
		};
		const configurationService = new TestConfigurationService({
			[mcpEnterpriseManagedAuthIdpSection]: { issuer: configuredIssuer.toString(true) }
		});
		disposables.add(configurationService.onDidChangeConfigurationEmitter);

		const mainThreadMcp = createMainThreadMcp(disposables, proxy, mcpRegistry, services => {
			services.stub(IAuthenticationService, authenticationService);
			services.stub(IAuthenticationMcpService, { getAccountPreference: () => undefined });
			services.stub(IAuthenticationMcpAccessService, { isAccessAllowedForUrl: () => true });
			services.stub(IAuthenticationMcpUsageService, { addAccountUsage() { } });
			services.stub(IConfigurationService, configurationService);
			services.stub(ISecretStorageService, {
				async get(key: string) {
					secretStorageReads.push(key);
					return secrets.get(key);
				}
			});
		});

		// Register a running HTTP server via the host delegate (the only path into the private maps).
		const launch: McpServerLaunch = { type: McpServerTransportType.HTTP, uri: URI.parse(serverUrl), headers: [] };
		const serverDefinition: McpServerDefinition = { id: 'my-server', label: 'My Server', launch, cacheNonce: 'nonce-1' };
		const collection: McpCollectionDefinition = {
			remoteAuthority: null,
			id: 'collection-1',
			label: 'Collection',
			serverDefinitions: observableValue<readonly McpServerDefinition[]>('serverDefinitions', [serverDefinition]),
			trustBehavior: McpServerTrust.Kind.Trusted,
			scope: StorageScope.WORKSPACE,
			configTarget: ConfigurationTarget.USER,
			order: McpCollectionSortOrder.Extension,
		};
		assert.ok(capturedDelegate, 'the MCP host delegate is registered');
		capturedDelegate.start(collection, serverDefinition, launch, {});
		mainThreadMcp.$onDidChangeState(1, { state: McpConnectionState.Kind.Running });

		const authDetails: IMcpAuthenticationDetails = {
			authorizationServer,
			authorizationServerMetadata: { issuer: authorizationServer.toString(), response_types_supported: ['code'], scopes_supported: ['scope.read'] },
			resourceMetadata: { resource, scopes_supported: ['scope.read'], authorization_servers: ['https://resource-as.example'] },
			scopes: ['scope.read'],
			clientId,
			enterpriseManaged,
		};
		await mainThreadMcp.$getTokenFromServerMetadata(1, authDetails, {});

		for (const secret of ['rotated-secret', undefined]) {
			if (secretStorageKey) {
				if (secret === undefined) {
					secrets.delete(secretStorageKey);
				} else {
					secrets.set(secretStorageKey, secret);
				}
			}
			revalidated = new DeferredPromise<void>();
			onDidChangeSessions.fire({ providerId: 'test-provider', label: 'Test Provider', event: { added: undefined, removed: undefined, changed: undefined } });
			await revalidated.p;
		}

		const expectedSecrets = clientId ? ['original-secret', 'rotated-secret', undefined] : [undefined, undefined, undefined];
		assert.deepStrictEqual({ getSessionsOptions, secretStorageReads }, {
			getSessionsOptions: expectedSecrets.map((clientSecret, index) => ({
				authorizationServer: enterpriseManaged ? configuredIssuer : authorizationServer,
				clientId,
				clientSecret,
				resource,
				audience,
				silent: index > 0,
			})),
			secretStorageReads: secretStorageKey ? [secretStorageKey, secretStorageKey, secretStorageKey] : [],
		});
	}

	test('replays the tracked auth context with a rotated or removed client secret (#324925)', async () => {
		await assertRevalidationContext(false, 'client-abc');
	});

	test('replays XAA resource auth context with a rotated or removed client secret', async () => {
		await assertRevalidationContext(true, 'resource-client');
	});

	test('revalidates clients without a secret-storage key without reading secret storage', async () => {
		await assertRevalidationContext(false, undefined);
	});
});

suite('MainThreadMcp - explicit OAuth clients', () => {
	const disposables = ensureNoDisposablesAreLeakedInTestSuite();

	test('recovers and upgrades the correct OAuth account while keeping other grants', async () => {
		const issuer = URI.parse('https://auth.example');
		const uri = URI.parse('https://mcp.example/mcp');
		const resource = { resource: uri.toString(true) };
		const providerId = getDynamicAuthenticationProviderId(issuer, resource);
		const account = { id: 'support', label: 'Support workspace' };
		const rejected: AuthenticationSession = { id: 'old-session', accessToken: 'old-token', account, scopes: ['read'] };
		const other: AuthenticationSession = { id: 'other-session', accessToken: 'other-token', account: { id: 'engineering', label: 'Engineering workspace' }, scopes: ['read'] };
		const renewed: AuthenticationSession = { ...rejected, id: 'renewed-session', accessToken: 'renewed-token' };
		const upgraded: AuthenticationSession = { ...renewed, id: 'upgraded-session', accessToken: 'upgraded-token', scopes: ['read', 'write'] };
		const otherUpgraded: AuthenticationSession = { ...other, id: 'other-upgraded', accessToken: 'other-write-token', scopes: ['read', 'write'] };
		const sessions = [rejected, other, otherUpgraded];
		const removed: string[] = [];
		const refreshed: AuthenticationSession = { ...renewed, id: 'refreshed-session', accessToken: 'refreshed-token' };
		let allowRefresh = false;
		let allowSupportAccess = true;
		let successfulRefreshes = 0;
		let signIns = 0;
		let prompts = 0;
		let allow = false;
		let switchAccounts = false;
		let delegate: IMcpHostDelegate | undefined;
		const registry = new class extends mock<IMcpRegistry>() {
			override readonly collections = observableValue<readonly McpCollectionDefinition[]>('collections', []);
			override registerDelegate(value: IMcpHostDelegate) { delegate = value; return { dispose() { } }; }
		};
		const provider = new class extends mock<IAuthenticationProvider>() {
			override readonly id = providerId;
			override readonly label = 'Fixture provider';
			override readonly supportsMultipleAccounts = false;
		};
		const mainThread = createMainThreadMcp(disposables, { $onDidChangeMcpServerDefinitions() { }, $startMcp() { } }, registry, services => {
			services.stub(IAuthenticationService, {
				onDidChangeSessions: Event.None,
				getOrActivateProviderIdForServer: async () => providerId,
				isDynamicAuthenticationProvider: id => id === providerId,
				getProvider: () => provider,
				getSessions: async (id, scopes, options) => {
					if (options?.[MCP_REFRESH_REJECTED_OAUTH_TOKEN] && allowRefresh) {
						assert.strictEqual(id, providerId);
						assert.strictEqual(options.silent, true);
						assert.strictEqual(options[MCP_REFRESH_REJECTED_OAUTH_TOKEN], await getOAuthTokenFingerprint(renewed.accessToken));
						successfulRefreshes++;
						sessions.splice(sessions.indexOf(renewed), 1, refreshed);
						return [refreshed];
					}
					return sessions.filter(session => JSON.stringify(session.scopes) === JSON.stringify(scopes));
				},
				unregisterAuthenticationProvider: () => assert.fail('Recovery must keep the registration'),
				createDynamicAuthenticationProvider: async () => assert.fail('Recovery must reuse the existing provider'),
				createSession: async (id, scopes, options) => {
					assert.deepStrictEqual([id, options?.account], [providerId, account]);
					signIns++;
					if (JSON.stringify(scopes) === JSON.stringify(upgraded.scopes)) {
						if (switchAccounts) { return otherUpgraded; }
						sessions.push(upgraded);
						return upgraded;
					}
					assert.deepStrictEqual(scopes, ['read']);
					sessions.push(renewed);
					return renewed;
				},
				removeSession: async (id, sessionId) => {
					assert.strictEqual(id, providerId);
					assert.strictEqual(sessionId, rejected.id);
					removed.push(sessionId);
					sessions.splice(sessions.indexOf(rejected), 1);
				},
			});
			services.stub(IDialogService, { prompt: async () => { prompts++; return { result: (allow || undefined) as never }; } });
			services.stub(IAuthenticationMcpService, {
				getAccountPreference: definitionId => definitionId === 'support' ? account.label : other.account.label,
				updateAccountPreference() { },
			});
			services.stub(IAuthenticationMcpAccessService, { isAccessAllowedForUrl: (_provider, _account, serverId) => serverId !== 'support' || allowSupportAccess, updateAllowedMcpServers() { } });
			services.stub(IAuthenticationMcpUsageService, { addAccountUsage() { } });
		});
		const launch: McpServerLaunch = { type: McpServerTransportType.HTTP, uri, headers: [] };
		const definitions: McpServerDefinition[] = ['support', 'engineering'].map(name => ({ id: name, label: name, launch, cacheNonce: 'fixture' }));
		const collection: McpCollectionDefinition = {
			id: 'fixtures', label: 'Fixtures', remoteAuthority: null, serverDefinitions: observableValue<readonly McpServerDefinition[]>('definitions', definitions),
			trustBehavior: McpServerTrust.Kind.Trusted, scope: StorageScope.PROFILE,
			configTarget: ConfigurationTarget.USER_LOCAL, order: McpCollectionSortOrder.Extension,
		};
		assert.ok(delegate);
		for (const definition of definitions) { delegate.start(collection, definition, launch, {}); }
		const details: IMcpAuthenticationDetails = {
			authorizationServer: issuer, authorizationServerMetadata: { issuer: issuer.toString(true), response_types_supported: ['code'] },
			resourceMetadata: resource, scopes: ['read'],
		};
		const recovery = { rejectedTokenFingerprint: await getOAuthTokenFingerprint(rejected.accessToken) };
		assert.deepStrictEqual(await mainThread.$getTokenFromServerMetadata(1, details), { accessToken: rejected.accessToken, scopes: ['read'] });
		await assert.rejects(mainThread.$getTokenFromServerMetadata(1, details, { ...recovery, errorOnUserInteraction: true }), /User interaction required: authentication/);
		assert.strictEqual(prompts, 0);
		await assert.rejects(mainThread.$getTokenFromServerMetadata(1, details, recovery), isCancellationError);
		assert.deepStrictEqual(sessions, [rejected, other, otherUpgraded]);
		allow = true;
		assert.deepStrictEqual(await Promise.all([
			mainThread.$getTokenFromServerMetadata(1, details, recovery),
			mainThread.$getTokenFromServerMetadata(1, details, recovery),
		]), [renewed, renewed].map(({ accessToken, scopes }) => ({ accessToken, scopes })));
		assert.deepStrictEqual({ prompts, signIns, removed, other: await mainThread.$getTokenFromServerMetadata(2, details) }, {
			prompts: 2, signIns: 1, removed: [rejected.id], other: { accessToken: other.accessToken, scopes: ['read'] },
		});
		const expandedDetails = { ...details, scopes: [...upgraded.scopes] };
		const upgrade = { scopeUpgradeFrom: { scopes: [...renewed.scopes], tokenFingerprint: await getOAuthTokenFingerprint(renewed.accessToken) } };
		await assert.rejects(mainThread.$getTokenFromServerMetadata(1, expandedDetails, { ...upgrade, errorOnUserInteraction: true }), /User interaction required: authentication/);
		assert.strictEqual(prompts, 2);
		switchAccounts = true;
		await assert.rejects(mainThread.$getTokenFromServerMetadata(1, expandedDetails, upgrade), /User interaction required: authentication/);
		assert.deepStrictEqual(await mainThread.$getTokenFromServerMetadata(1, details), { accessToken: renewed.accessToken, scopes: ['read'] });
		switchAccounts = false;
		assert.deepStrictEqual(await Promise.all([
			mainThread.$getTokenFromServerMetadata(1, expandedDetails, upgrade),
			mainThread.$getTokenFromServerMetadata(1, expandedDetails, upgrade),
		]), [upgraded, upgraded].map(({ accessToken, scopes }) => ({ accessToken, scopes })));
		assert.deepStrictEqual({ prompts, signIns, removed, other: await mainThread.$getTokenFromServerMetadata(2, details) }, {
			prompts: 4, signIns: 3, removed: [rejected.id], other: { accessToken: other.accessToken, scopes: ['read'] },
		});

		allowRefresh = true;
		const earlyRejection = { rejectedTokenFingerprint: await getOAuthTokenFingerprint(renewed.accessToken), errorOnUserInteraction: true };
		allowSupportAccess = false;
		await assert.rejects(mainThread.$getTokenFromServerMetadata(1, details, earlyRejection), /User interaction required: authentication/);
		assert.strictEqual(successfulRefreshes, 0);
		allowSupportAccess = true;
		assert.deepStrictEqual(await Promise.all([
			mainThread.$getTokenFromServerMetadata(1, details, earlyRejection),
			mainThread.$getTokenFromServerMetadata(1, details, earlyRejection),
		]), [refreshed, refreshed].map(({ accessToken, scopes }) => ({ accessToken, scopes })));
		assert.deepStrictEqual({ successfulRefreshes, prompts, signIns, removed }, { successfulRefreshes: 1, prompts: 4, signIns: 3, removed: [rejected.id] });
	});

	test('routes an explicit HTTPS callback to the registered MCP client and its browser flow', async () => {
		const issuer = URI.parse('https://issuer.example.com');
		const uri = URI.parse('https://service.example.com/mcp');
		const resource = { resource: uri.toString(true) };
		const providerId = getDynamicAuthenticationProviderId(issuer, resource, 'registered-client');
		const launch: McpServerLaunch = { type: McpServerTransportType.HTTP, uri, headers: [], oauth: { useHttpsCallback: true } };
		let delegate: IMcpHostDelegate | undefined;
		const registry = new class extends mock<IMcpRegistry>() {
			override readonly collections = observableValue<readonly McpCollectionDefinition[]>('collections', []);
			override registerDelegate(value: IMcpHostDelegate) { delegate = value; return { dispose() { } }; }
		};
		const mainThread = createMainThreadMcp(disposables, { $onDidChangeMcpServerDefinitions() { }, $startMcp() { } }, registry, services => {
			services.stub(IAuthenticationService, {
				onDidChangeSessions: Event.None,
				getOrActivateProviderIdForServer: async () => assert.fail('A shared provider cannot choose the registered callback'),
				isDynamicAuthenticationProvider: () => false,
				isAuthenticationProviderRegistered: () => false,
				createDynamicAuthenticationProvider: async (_server, _metadata, _resource, clientId, _secret, id) => {
					assert.deepStrictEqual([clientId, id], ['registered-client', providerId]);
					return upcastPartial<IAuthenticationProvider>({ id: providerId, label: 'Registered client' });
				},
				getProvider: () => upcastPartial<IAuthenticationProvider>({ id: providerId, label: 'Registered client', supportsMultipleAccounts: true }),
				getSessions: async () => [],
				createSession: async (id, scopes, options) => {
					assert.strictEqual(id, providerId);
					assert.strictEqual(options?.[MCP_USE_HTTPS_OAUTH_CALLBACK], true);
					assert.deepStrictEqual(scopes, ['read']);
					return { id: 'fixture-session', accessToken: 'fixture-token', account: { id: 'account', label: 'Account' }, scopes: ['read'] };
				},
			});
			services.stub(IDynamicAuthenticationProviderStorageService, { getClientRegistration: async () => undefined });
			services.stub(IDialogService, { prompt: async () => ({ result: true as never }) });
			services.stub(IAuthenticationMcpService, { getAccountPreference: () => undefined, updateAccountPreference() { } });
			services.stub(IAuthenticationMcpAccessService, { updateAllowedMcpServers() { } });
			services.stub(IAuthenticationMcpUsageService, { addAccountUsage() { } });
		});
		const definition: McpServerDefinition = { id: 'https', label: 'HTTPS fixture', launch, cacheNonce: 'fixture' };
		const collection: McpCollectionDefinition = {
			id: 'fixtures', label: 'Fixtures', remoteAuthority: null,
			serverDefinitions: observableValue<readonly McpServerDefinition[]>('definitions', [definition]),
			trustBehavior: McpServerTrust.Kind.Trusted, scope: StorageScope.PROFILE,
			configTarget: ConfigurationTarget.USER_LOCAL, order: McpCollectionSortOrder.Extension,
		};
		assert.ok(delegate);
		delegate.start(collection, definition, launch, {});
		const details: IMcpAuthenticationDetails = {
			authorizationServer: issuer, resourceMetadata: resource, scopes: ['read'],
			authorizationServerMetadata: { issuer: issuer.toString(true), response_types_supported: ['code'] },
		};
		await assert.rejects(mainThread.$getTokenFromServerMetadata(1, details), /requires a registered client ID/);
		assert.deepStrictEqual(await mainThread.$getTokenFromServerMetadata(1, { ...details, clientId: 'registered-client' }), { accessToken: 'fixture-token', scopes: ['read'] });
	});

	test('prompts once for a missing Google Workspace client secret only on an interactive connection', async () => {
		const authorizationServer = URI.parse('https://accounts.google.com');
		const uri = URI.parse('https://gmailmcp.googleapis.com/mcp/v1');
		const key = mcpOAuthClientSecretStorageKey(uri.toString(true), 'gmail-client');
		let value: string | undefined;
		let promptCount = 0;
		const registeredSecrets: (string | undefined)[] = [];
		let delegate: IMcpHostDelegate | undefined;
		const registry = new class extends mock<IMcpRegistry>() {
			override readonly collections = observableValue<readonly McpCollectionDefinition[]>('collections', []);
			override registerDelegate(value: IMcpHostDelegate) { delegate = value; return { dispose() { } }; }
		};
		const mainThread = createMainThreadMcp(disposables, { $onDidChangeMcpServerDefinitions() { }, $startMcp() { } }, registry, services => {
			services.stub(IAuthenticationService, {
				onDidChangeSessions: Event.None,
				getOrActivateProviderIdForServer: async () => undefined,
				isDynamicAuthenticationProvider: () => false,
				isAuthenticationProviderRegistered: () => false,
				createDynamicAuthenticationProvider: async (_server, _metadata, _resource, _client, secret) => {
					registeredSecrets.push(secret);
					return undefined;
				},
			});
			services.stub(IDynamicAuthenticationProviderStorageService, { getClientRegistration: async () => undefined });
			services.stub(ISecretStorageService, {
				get: async requestKey => { assert.strictEqual(requestKey, key); return value; },
				set: async (requestKey, secret) => { assert.strictEqual(requestKey, key); value = secret; },
			});
			services.stub(IQuickInputService, {
				input: async options => {
					promptCount++;
					assert.strictEqual(options?.password, true);
					return 'fixture-google-secret';
				},
			});
		});
		const launch: McpServerLaunch = { type: McpServerTransportType.HTTP, uri, headers: [] };
		const definition: McpServerDefinition = { id: 'gmail', label: 'Gmail', launch, cacheNonce: 'fixture' };
		const collection: McpCollectionDefinition = {
			id: 'fixtures', label: 'Fixtures', remoteAuthority: null,
			serverDefinitions: observableValue<readonly McpServerDefinition[]>('definitions', [definition]),
			trustBehavior: McpServerTrust.Kind.Trusted, scope: StorageScope.PROFILE,
			configTarget: ConfigurationTarget.USER_LOCAL, order: McpCollectionSortOrder.Extension,
		};
		assert.ok(delegate);
		delegate.start(collection, definition, launch, {});
		const details: IMcpAuthenticationDetails = {
			authorizationServer, clientId: 'gmail-client', resourceMetadata: { resource: uri.toString(true) }, scopes: ['read'],
			authorizationServerMetadata: {
				issuer: authorizationServer.toString(true), response_types_supported: ['code'],
				authorization_endpoint: 'https://accounts.google.com/o/oauth2/v2/auth',
				token_endpoint: 'https://oauth2.googleapis.com/token',
			},
		};
		await assert.rejects(mainThread.$getTokenFromServerMetadata(1, details, { errorOnUserInteraction: true }), /User interaction required: authentication/);
		assert.deepStrictEqual({ promptCount, value, registeredSecrets }, { promptCount: 0, value: undefined, registeredSecrets: [] });
		await Promise.all([mainThread.$getTokenFromServerMetadata(1, details), mainThread.$getTokenFromServerMetadata(1, details)]);
		assert.deepStrictEqual({ promptCount, value, registeredSecrets }, {
			promptCount: 1, value: 'fixture-google-secret', registeredSecrets: ['fixture-google-secret', 'fixture-google-secret'],
		});
	});

	test('retains a matching legacy sign-in, separates concurrent client registrations and rotates only the selected client', async () => {
		const authority = URI.parse('https://auth.example');
		const uri = URI.parse('https://mcp.example/mcp');
		const resource = { resource: uri.toString(true) };
		const legacyId = getDynamicAuthenticationProviderId(authority, resource);
		const secondId = getDynamicAuthenticationProviderId(authority, resource, 'engineering');
		const registrations = new Map<string, { clientId: string; clientSecret?: string }>([[legacyId, { clientId: 'support' }]]);
		const tokens = new Map<string, string>([[legacyId, 'support-session']]);
		const removed: string[] = [];
		const created: { id: string; clientId: string; clientSecret?: string }[] = [];
		const clientSecrets = new Map<string, string>();
		let delegate: IMcpHostDelegate | undefined;
		const mcpRegistry = new class extends mock<IMcpRegistry>() {
			override readonly collections = observableValue<readonly McpCollectionDefinition[]>('collections', []);
			override registerDelegate(value: IMcpHostDelegate) { delegate = value; return { dispose() { } }; }
		};
		function provider(id: string): IAuthenticationProvider {
			return new class extends mock<IAuthenticationProvider>() {
				override readonly id = id;
				override readonly label = 'Fixture provider';
				override readonly supportsMultipleAccounts = false;
			};
		}
		const mainThread = createMainThreadMcp(disposables, { $onDidChangeMcpServerDefinitions() { }, $startMcp() { } }, mcpRegistry, services => {
			services.stub(IAuthenticationService, {
				onDidChangeSessions: Event.None,
				getOrActivateProviderIdForServer: async () => legacyId,
				isDynamicAuthenticationProvider: id => tokens.has(id),
				unregisterAuthenticationProvider: id => { tokens.delete(id); },
				createDynamicAuthenticationProvider: async (_authority, _metadata, _resource, clientId, clientSecret, id) => {
					assert.ok(id && clientId);
					assert.strictEqual(tokens.has(id), false);
					await Promise.resolve();
					registrations.set(id, { clientId, clientSecret });
					created.push({ id, clientId, clientSecret });
					tokens.set(id, `${clientId}-session-${created.length}`);
					return provider(id);
				},
				getProvider: provider,
				getSessions: async id => {
					const accessToken = tokens.get(id);
					assert.ok(accessToken);
					return [{ id: accessToken, accessToken, account: { id: 'workspace', label: 'Same workspace label' }, scopes: ['read'] }];
				},
			});
			services.stub(IDynamicAuthenticationProviderStorageService, {
				getClientRegistration: async id => registrations.get(id),
				removeDynamicProvider: async id => { removed.push(id); registrations.delete(id); },
			});
			services.stub(ISecretStorageService, { get: async key => clientSecrets.get(key) });
			services.stub(IAuthenticationMcpService, { getAccountPreference: () => undefined });
			services.stub(IAuthenticationMcpAccessService, { isAccessAllowedForUrl: () => true });
			services.stub(IAuthenticationMcpUsageService, { addAccountUsage() { } });
		});
		const launch: McpServerLaunch = { type: McpServerTransportType.HTTP, uri, headers: [] };
		const definitions = ['support', 'engineering', 'engineering-copy'].map(name => ({ id: name, label: name, launch, cacheNonce: 'fixture' }));
		const collection: McpCollectionDefinition = {
			id: 'fixtures', label: 'Fixtures', remoteAuthority: null, serverDefinitions: observableValue<readonly McpServerDefinition[]>('definitions', definitions),
			trustBehavior: McpServerTrust.Kind.Trusted, scope: StorageScope.PROFILE, configTarget: ConfigurationTarget.USER_LOCAL, order: McpCollectionSortOrder.Extension,
		};
		assert.ok(delegate);
		for (const definition of definitions) { delegate.start(collection, definition, launch, {}); }
		const getToken = async (id: number, clientId: string) => (await mainThread.$getTokenFromServerMetadata(id, {
			authorizationServer: authority,
			authorizationServerMetadata: { issuer: authority.toString(true), response_types_supported: ['code'] },
			resourceMetadata: resource, scopes: ['read'], clientId,
		}))?.accessToken;
		assert.strictEqual(await getToken(1, 'support'), 'support-session');
		assert.deepStrictEqual(await Promise.all([getToken(2, 'engineering'), getToken(3, 'engineering')]), ['engineering-session-1', 'engineering-session-1']);
		assert.deepStrictEqual(created, [{ id: secondId, clientId: 'engineering', clientSecret: undefined }]);
		clientSecrets.set(mcpOAuthClientSecretStorageKey(uri.toString(true), 'engineering'), 'rotated-engineering-secret');
		assert.strictEqual(await getToken(2, 'engineering'), 'engineering-session-2');
		assert.deepStrictEqual(removed, [secondId]);
		assert.strictEqual(await getToken(1, 'support'), 'support-session');
		assert.deepStrictEqual(registrations.get(secondId), { clientId: 'engineering', clientSecret: 'rotated-engineering-secret' });
	});
});

suite('MainThreadMcp - request policy', () => {
	const disposables = ensureNoDisposablesAreLeakedInTestSuite();
	const initialUrl = 'https://allowed.example/mcp';
	const redirectedUrl = 'https://redirected.example/mcp';
	const name = 'Configured MCP Server';

	function createConnection(options: {
		allowed?: readonly IMcpServerMatcher[];
		denied?: readonly IMcpServerMatcher[];
		access?: McpAccessValue;
		destination?: string;
		env?: Record<string, string>;
	} = {}) {
		const configuration = new TestConfigurationService({
			[mcpAllowedServersConfig]: options.allowed,
			[mcpDeniedServersConfig]: options.denied,
			[mcpAccessConfig]: options.access,
		});
		disposables.add(configuration.onDidChangeConfigurationEmitter);
		const policy = disposables.add(new AllowedMcpServersService(configuration));
		const stopped: number[] = [];
		const requests: string[] = [];
		let delegate: IMcpHostDelegate | undefined;
		const registry = new class extends mock<IMcpRegistry>() {
			override readonly collections = observableValue<readonly McpCollectionDefinition[]>('collections', []);
			override registerDelegate(value: IMcpHostDelegate) {
				delegate = value;
				return { dispose() { } };
			}
		};
		const mainThread = createMainThreadMcp(disposables, {
			$startMcp() { },
			$stopMcp(id) {
				stopped.push(id);
				httpHandle.dispose();
			},
			$sendMessage() { },
			$onDidChangeMcpServerDefinitions() { },
			async $resolveMcpPolicyVariables(values) {
				const resolve = createEnvVarResolver(options.env, false);
				return Object.fromEntries(values.map(value => [value, resolve(value)]));
			},
		}, registry, services => {
			services.stub(IConfigurationService, configuration);
			services.stub(IAllowedMcpServersService, policy);
		});
		const launch = { type: McpServerTransportType.HTTP, uri: URI.parse(initialUrl), headers: [] } satisfies McpServerLaunch;
		const definition: McpServerDefinition = { id: 'server', label: name, cacheNonce: 'nonce', launch };
		const collection: McpCollectionDefinition = {
			id: 'collection',
			label: 'Collection',
			remoteAuthority: null,
			serverDefinitions: observableValue<readonly McpServerDefinition[]>('definitions', [definition]),
			trustBehavior: McpServerTrust.Kind.Trusted,
			scope: StorageScope.WORKSPACE,
			configTarget: ConfigurationTarget.USER,
			order: McpCollectionSortOrder.WorkspaceFolder,
		};
		assert.ok(delegate);
		const transport = delegate.start(collection, definition, launch);
		const destination = options.destination ?? redirectedUrl;
		const httpHandle = disposables.add(new class extends McpHTTPHandle {
			protected override async _fetchInternal(url: string, _init?: CommonRequestInit): Promise<CommonResponse> {
				requests.push(url);
				return url === initialUrl
					? new Response(null, { status: 307, headers: { location: destination } })
					: new Response(null, { status: 202 });
			}
		}(1, launch, mainThread, new NullLogService()));
		return {
			mainThread, policy, transport, stopped, requests,
			// eslint-disable-next-line local/code-no-bracket-notation-for-identifiers -- Exercise the private request boundary without widening the transport API.
			fetch: () => httpHandle['_fetch'](initialUrl, { method: 'POST', headers: {} }),
			async denyDestination() {
				await configuration.setUserConfiguration(mcpDeniedServersConfig, [{ serverUrl: destination }]);
				configuration.onDidChangeConfigurationEmitter.fire({
					source: ConfigurationTarget.USER,
					affectedKeys: new Set([mcpDeniedServersConfig]),
					change: { keys: [mcpDeniedServersConfig], overrides: [] },
					affectsConfiguration: section => section === mcpDeniedServersConfig,
				});
			},
		};
	}

	for (const destination of [redirectedUrl, 'https://allowed.example/private']) {
		test(`enforces URL denials before dispatch to ${destination}`, async () => {
			const { fetch, policy, requests, transport } = createConnection({ denied: [{ serverUrl: destination }], destination });
			assert.strictEqual(policy.isServerAllowed({ name, url: initialUrl }), true);
			const denial = policy.isServerAllowed({ name, url: destination });
			assert.notStrictEqual(denial, true);
			assert.ok(denial !== true);
			await assert.rejects(fetch(), { message: denial.value });
			assert.deepStrictEqual({ requests, state: transport.state.get() }, {
				requests: [initialUrl],
				state: { state: McpConnectionState.Kind.Error, message: denial.value },
			});
		});
	}

	test('requires a URL-allowlisted redirect destination to be separately allowed', async () => {
		const { fetch, requests } = createConnection({ allowed: [{ serverUrl: initialUrl }] });
		await assert.rejects(fetch(), /not in the list of servers allowed/);
		assert.deepStrictEqual(requests, [initialUrl]);
	});

	test('a name entry cannot bypass a URL entry at the request boundary', async () => {
		const { fetch, requests } = createConnection({ allowed: [{ serverName: name }, { serverUrl: initialUrl }] });
		await assert.rejects(fetch(), /not in the list of servers allowed/);
		assert.deepStrictEqual(requests, [initialUrl]);
	});

	test('expands URL policies on the connecting host for initial and redirected requests', async () => {
		const { fetch, requests } = createConnection({
			allowed: [{ serverUrl: 'https://${INITIAL}/mcp' }, { serverUrl: 'https://${REDIRECTED}/mcp' }],
			env: { INITIAL: 'allowed.example', REDIRECTED: 'redirected.example' },
		});
		await fetch();
		assert.deepStrictEqual(requests, [initialUrl, redirectedUrl]);
	});

	test('an expanded URL denial wins over an expanded allow entry', async () => {
		const { fetch, requests } = createConnection({
			allowed: [{ serverUrl: 'https://${INITIAL}/mcp' }, { serverUrl: 'https://${REDIRECTED}/mcp' }],
			denied: [{ serverUrl: 'https://${REDIRECTED}/mcp' }],
			env: { INITIAL: 'allowed.example', REDIRECTED: 'redirected.example' },
		});
		await assert.rejects(fetch(), /blocked by your organization\./);
		assert.deepStrictEqual(requests, [initialUrl]);
	});

	test('preserves the configured server name for destination policy decisions', async () => {
		const { fetch, requests } = createConnection({ allowed: [{ serverName: name }] });
		await fetch();
		assert.deepStrictEqual(requests, [initialUrl, redirectedUrl]);
	});

	test('a destination denial overrides a matching server name allowlist', async () => {
		const { fetch, requests } = createConnection({ allowed: [{ serverName: name }], denied: [{ serverUrl: redirectedUrl }] });
		await assert.rejects(fetch(), /blocked by your organization\./);
		assert.deepStrictEqual(requests, [initialUrl]);
	});

	test('allows explicitly permitted redirect destinations', async () => {
		const { fetch, requests } = createConnection({ allowed: [{ serverUrl: initialUrl }, { serverUrl: redirectedUrl }] });
		await fetch();
		assert.deepStrictEqual(requests, [initialUrl, redirectedUrl]);
	});

	test('enforces global MCP disablement before dispatch', async () => {
		const { fetch, requests } = createConnection({ access: McpAccessValue.None });
		await assert.rejects(fetch(), /Model Context Protocol servers are disabled/);
		assert.deepStrictEqual(requests, []);
	});

	test('enforces configured server name denials before dispatch', async () => {
		const { fetch, requests } = createConnection({ denied: [{ serverName: name }] });
		await assert.rejects(fetch(), /blocked by your organization\./);
		assert.deepStrictEqual(requests, []);
	});

	test('stops an active redirected transport when destination policy is revoked', async () => {
		const { fetch, denyDestination, stopped, requests, transport } = createConnection();
		await fetch();
		await denyDestination();
		assert.deepStrictEqual({ stopped, state: transport.state.get().state, requests }, {
			stopped: [1],
			state: McpConnectionState.Kind.Error,
			requests: [initialUrl, redirectedUrl],
		});
	});

	test('does not dispatch using a removed connection', async () => {
		const { mainThread, fetch, requests } = createConnection();
		mainThread.$onDidChangeState(1, { state: McpConnectionState.Kind.Stopped });
		await assert.rejects(fetch(), isCancellationError);
		assert.deepStrictEqual(requests, []);
	});
});
