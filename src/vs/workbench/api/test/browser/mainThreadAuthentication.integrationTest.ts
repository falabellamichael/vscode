/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import assert from 'assert';
import { CancellationToken, CancellationTokenSource } from '../../../../base/common/cancellation.js';
import { isCancellationError } from '../../../../base/common/errors.js';
import { URI } from '../../../../base/common/uri.js';
import { IDialogService } from '../../../../platform/dialogs/common/dialogs.js';
import { TestDialogService } from '../../../../platform/dialogs/test/common/testDialogService.js';
import { TestInstantiationService } from '../../../../platform/instantiation/test/common/instantiationServiceMock.js';
import { INotificationService } from '../../../../platform/notification/common/notification.js';
import { TestNotificationService } from '../../../../platform/notification/test/common/testNotificationService.js';
import { IQuickInputService } from '../../../../platform/quickinput/common/quickInput.js';
import { IStorageService } from '../../../../platform/storage/common/storage.js';
import { ITelemetryService } from '../../../../platform/telemetry/common/telemetry.js';
import { IURLService } from '../../../../platform/url/common/url.js';
import { NativeURLService } from '../../../../platform/url/common/urlService.js';
import { NullTelemetryService } from '../../../../platform/telemetry/common/telemetryUtils.js';
import { MainThreadAuthentication } from '../../browser/mainThreadAuthentication.js';
import { ExtHostAuthenticationShape, ExtHostContext, MainContext } from '../../common/extHost.protocol.js';
import { upcastPartial } from '../../../../base/test/common/mock.js';
import { IActivityService } from '../../../services/activity/common/activity.js';
import { AuthenticationService } from '../../../services/authentication/browser/authenticationService.js';
import { IAuthenticationExtensionsService, IAuthenticationService } from '../../../services/authentication/common/authentication.js';
import { IExtensionService } from '../../../services/extensions/common/extensions.js';
import { IRemoteAgentService } from '../../../services/remote/common/remoteAgentService.js';
import { TestRPCProtocol } from '../common/testRPCProtocol.js';
import { TestEnvironmentService, TestHostService, TestQuickInputService, TestRemoteAgentService } from '../../../test/browser/workbenchTestServices.js';
import { TestActivityService, TestExtensionService, TestProductService, TestStorageService } from '../../../test/common/workbenchTestServices.js';
import { IBrowserWorkbenchEnvironmentService } from '../../../services/environment/browser/environmentService.js';
import { IProductService } from '../../../../platform/product/common/productService.js';
import { AuthenticationAccessService, IAuthenticationAccessService } from '../../../services/authentication/browser/authenticationAccessService.js';
import { IAccountUsage, IAuthenticationUsageService } from '../../../services/authentication/browser/authenticationUsageService.js';
import { AuthenticationExtensionsService } from '../../../services/authentication/browser/authenticationExtensionsService.js';
import { ILogService, NullLogService } from '../../../../platform/log/common/log.js';
import { IHostService } from '../../../services/host/browser/host.js';
import { IOpenerService } from '../../../../platform/opener/common/opener.js';
import { IUserActivityService, UserActivityService } from '../../../services/userActivity/common/userActivityService.js';
import { ISecretStorageService } from '../../../../platform/secrets/common/secrets.js';
import { TestSecretStorageService } from '../../../../platform/secrets/test/common/testSecretStorageService.js';
import { IDynamicAuthenticationProviderStorageService } from '../../../services/authentication/common/dynamicAuthenticationProviderStorage.js';
import { DynamicAuthenticationProviderStorageService } from '../../../services/authentication/browser/dynamicAuthenticationProviderStorageService.js';
import { ensureNoDisposablesAreLeakedInTestSuite } from '../../../../base/test/common/utils.js';
import { ServiceCollection } from '../../../../platform/instantiation/common/serviceCollection.js';
import { SyncDescriptor } from '../../../../platform/instantiation/common/descriptors.js';

class TestAuthUsageService implements IAuthenticationUsageService {
	_serviceBrand: undefined;
	initializeExtensionUsageCache(): Promise<void> { return Promise.resolve(); }
	extensionUsesAuth(extensionId: string): Promise<boolean> { return Promise.resolve(false); }
	readAccountUsages(providerId: string, accountName: string): IAccountUsage[] { return []; }
	removeAccountUsage(providerId: string, accountName: string): void { }
	addAccountUsage(providerId: string, accountName: string, scopes: ReadonlyArray<string>, extensionId: string, extensionName: string): void { }
}

suite('MainThreadAuthentication', () => {
	const disposables = ensureNoDisposablesAreLeakedInTestSuite();

	let mainThreadAuthentication: MainThreadAuthentication;
	let instantiationService: TestInstantiationService;
	let rpcProtocol: TestRPCProtocol;

	setup(async () => {
		// services
		const services = new ServiceCollection();
		services.set(ILogService, new SyncDescriptor(NullLogService));
		services.set(IDialogService, new SyncDescriptor(TestDialogService, [{ confirmed: true }]));
		services.set(IStorageService, new SyncDescriptor(TestStorageService));
		services.set(ISecretStorageService, new SyncDescriptor(TestSecretStorageService));
		services.set(IDynamicAuthenticationProviderStorageService, new SyncDescriptor(DynamicAuthenticationProviderStorageService));
		services.set(IQuickInputService, new SyncDescriptor(TestQuickInputService));
		services.set(IExtensionService, new SyncDescriptor(TestExtensionService));
		services.set(IActivityService, new SyncDescriptor(TestActivityService));
		services.set(IRemoteAgentService, new SyncDescriptor(TestRemoteAgentService));
		services.set(INotificationService, new SyncDescriptor(TestNotificationService));
		services.set(IHostService, new SyncDescriptor(TestHostService));
		services.set(IUserActivityService, new SyncDescriptor(UserActivityService));
		services.set(IAuthenticationAccessService, new SyncDescriptor(AuthenticationAccessService));
		services.set(IAuthenticationService, new SyncDescriptor(AuthenticationService));
		services.set(IAuthenticationUsageService, new SyncDescriptor(TestAuthUsageService));
		services.set(IAuthenticationExtensionsService, new SyncDescriptor(AuthenticationExtensionsService));
		instantiationService = disposables.add(new TestInstantiationService(services, undefined, undefined, true));

		// stubs
		// eslint-disable-next-line local/code-no-dangerous-type-assertions
		instantiationService.stub(IOpenerService, {} as Partial<IOpenerService>);
		instantiationService.stub(ITelemetryService, NullTelemetryService);
		instantiationService.stub(IBrowserWorkbenchEnvironmentService, TestEnvironmentService);
		instantiationService.stub(IProductService, TestProductService);
		instantiationService.stub(IURLService, disposables.add(new NativeURLService(TestProductService)));

		rpcProtocol = disposables.add(new TestRPCProtocol());
		mainThreadAuthentication = disposables.add(instantiationService.createInstance(MainThreadAuthentication, rpcProtocol));
		rpcProtocol.set(MainContext.MainThreadAuthentication, mainThreadAuthentication);
	});

	test('routes concurrent browser returns from the same provider using their own nonce', async () => {
		const urlService = instantiationService.get(IURLService);
		const callback = URI.parse('signal-c://dynamicauthprovider/issuer.example/authorize');
		const firstUri = callback.with({ query: 'nonce=first' });
		const secondUri = callback.with({ query: 'nonce=second' });
		const first = mainThreadAuthentication.$waitForUriHandler(firstUri, CancellationToken.None);
		const second = mainThreadAuthentication.$waitForUriHandler(secondUri, CancellationToken.None);
		const firstReturn = firstUri.with({ query: 'nonce=first&code=first-grant' });
		const secondReturn = secondUri.with({ query: 'nonce=second&code=second-grant' });

		assert.strictEqual(await urlService.open(callback.with({ query: 'code=missing-nonce' })), false);
		assert.strictEqual(await urlService.open(callback.with({ query: 'nonce=first&nonce=second&code=ambiguous' })), false);
		assert.strictEqual(await urlService.open(secondReturn), true);
		assert.strictEqual(await second, secondReturn);
		assert.strictEqual(await urlService.open(secondReturn), false);
		assert.strictEqual(await urlService.open(firstReturn), true);
		assert.strictEqual(await first, firstReturn);
	});

	test('cancellation and extension host disposal remove waiting browser callbacks', async () => {
		const urlService = instantiationService.get(IURLService);
		const token = disposables.add(new CancellationTokenSource());
		const callback = URI.parse('signal-c://dynamicauthprovider/issuer.example/authorize?nonce=canceled');
		const canceled = assert.rejects(mainThreadAuthentication.$waitForUriHandler(callback, token.token), isCancellationError);
		token.cancel();
		await canceled;
		assert.strictEqual(await urlService.open(callback), false);

		const otherCallback = callback.with({ query: 'nonce=disposed' });
		const disposed = assert.rejects(mainThreadAuthentication.$waitForUriHandler(otherCallback, CancellationToken.None), isCancellationError);
		mainThreadAuthentication.dispose();
		await disposed;
		assert.strictEqual(await urlService.open(otherCallback), false);
	});

	test('requires the original window on a pending HTTPS OAuth return', async () => {
		const urlService = instantiationService.get(IURLService);
		const callback = URI.parse('signal-c://dynamicauthprovider/https/authorize?nonce=fixture&windowId=41');
		const pending = mainThreadAuthentication.$waitForUriHandler(callback, CancellationToken.None);
		assert.strictEqual(await urlService.open(callback.with({ query: 'nonce=fixture&windowId=4&code64=YQ' })), false);
		assert.strictEqual(await urlService.open(callback.with({ query: 'nonce=fixture&windowId=41&windowId=4&code64=YQ' })), false);
		const returned = callback.with({ query: `${callback.query}&code64=YQ` });
		assert.strictEqual(await urlService.open(returned), true);
		assert.strictEqual(await pending, returned);
	});

	test('provider registration completes without errors', async () => {
		// Test basic registration - this should complete without throwing
		await mainThreadAuthentication.$registerAuthenticationProvider({
			id: 'test-provider',
			label: 'Test Provider',
			supportsMultipleAccounts: false
		});

		// Test unregistration - this should also complete without throwing
		await mainThreadAuthentication.$unregisterAuthenticationProvider('test-provider');

		// Success if we reach here without timeout
		assert.ok(true, 'Registration and unregistration completed successfully');
	});

	test('uses an explicit dynamic provider ID for registration and persisted tokens', async () => {
		const authorizationServer = URI.parse('https://auth.example.com');
		const resource = { resource: 'https://resource.example.com/mcp' };
		const providerId = 'https://auth.example.com/ https://resource.example.com/mcp client:configured-client';
		const tokens = [{ access_token: 'stored-token', token_type: 'Bearer', created_at: 0 }];
		const registrations: Parameters<ExtHostAuthenticationShape['$registerDynamicAuthProvider']>[] = [];
		rpcProtocol.set(ExtHostContext.ExtHostAuthentication, upcastPartial<ExtHostAuthenticationShape>({
			$registerDynamicAuthProvider: async (...args) => {
				registrations.push(args);
				const id = args[6] ?? 'unexpected-legacy-id';
				await mainThreadAuthentication.$registerAuthenticationProvider({ id, label: 'Test', supportsMultipleAccounts: false });
				return id;
			},
			$onDidChangeDynamicAuthProviderTokens: async () => { },
		}));
		const storage = instantiationService.get(IDynamicAuthenticationProviderStorageService);
		await storage.storeClientRegistration(providerId, authorizationServer.toString(true), 'configured-client', 'configured-secret', undefined, 'client_secret_basic');
		await storage.setSessionsForDynamicAuthProvider(providerId, 'configured-client', tokens);
		const metadata = { issuer: authorizationServer.toString(true), response_types_supported: ['code'] };

		const authenticationService = instantiationService.get(IAuthenticationService);
		const provider = await authenticationService.createDynamicAuthenticationProvider(authorizationServer, metadata, resource, undefined, undefined, providerId);
		await mainThreadAuthentication.$unregisterAuthenticationProvider(providerId);
		await authenticationService.createDynamicAuthenticationProvider(authorizationServer, metadata, resource, 'other-client', 'other-secret', providerId);
		await mainThreadAuthentication.$unregisterAuthenticationProvider(providerId);
		await rpcProtocol.sync();

		assert.deepStrictEqual({ providerId: provider?.id, registrations }, {
			providerId,
			registrations: [
				[authorizationServer.toJSON(), metadata, resource, 'configured-client', 'configured-secret', tokens, providerId, 'client_secret_basic', undefined],
				[authorizationServer.toJSON(), metadata, resource, 'other-client', 'other-secret', undefined, providerId, undefined, undefined],
			],
		});
	});

	test('event suppression during explicit unregistration', async () => {
		let unregisterEventFired = false;
		let eventProviderId: string | undefined;

		// Mock the ext host to capture unregister events
		const mockExtHost = {
			$onDidUnregisterAuthenticationProvider: (id: string) => {
				unregisterEventFired = true;
				eventProviderId = id;
				return Promise.resolve();
			},
			$getSessions: () => Promise.resolve([]),
			// eslint-disable-next-line local/code-no-any-casts
			$createSession: () => Promise.resolve({} as any),
			$removeSession: () => Promise.resolve(),
			$onDidChangeAuthenticationSessions: () => Promise.resolve(),
			$registerDynamicAuthProvider: () => Promise.resolve('test'),
			$registerXaaAuthProvider: () => Promise.resolve('test'),
			$onDidChangeDynamicAuthProviderTokens: () => Promise.resolve(),
			$getSessionsFromChallenges: () => Promise.resolve([]),
			// eslint-disable-next-line local/code-no-any-casts
			$createSessionFromChallenges: () => Promise.resolve({} as any),
		};
		rpcProtocol.set(ExtHostContext.ExtHostAuthentication, mockExtHost);

		// Register a provider
		await mainThreadAuthentication.$registerAuthenticationProvider({
			id: 'test-suppress',
			label: 'Test Suppress',
			supportsMultipleAccounts: false
		});

		// Reset the flag
		unregisterEventFired = false;
		eventProviderId = undefined;

		// Unregister the provider - this should NOT fire the event due to suppression
		await mainThreadAuthentication.$unregisterAuthenticationProvider('test-suppress');

		// Verify the event was suppressed
		assert.strictEqual(unregisterEventFired, false, 'Unregister event should be suppressed during explicit unregistration');
		assert.strictEqual(eventProviderId, undefined, 'No provider ID should be captured from suppressed event');
	});

	test('concurrent provider registrations complete without errors', async () => {
		// Register multiple providers simultaneously
		const registrationPromises = [
			mainThreadAuthentication.$registerAuthenticationProvider({
				id: 'concurrent-1',
				label: 'Concurrent 1',
				supportsMultipleAccounts: false
			}),
			mainThreadAuthentication.$registerAuthenticationProvider({
				id: 'concurrent-2',
				label: 'Concurrent 2',
				supportsMultipleAccounts: false
			}),
			mainThreadAuthentication.$registerAuthenticationProvider({
				id: 'concurrent-3',
				label: 'Concurrent 3',
				supportsMultipleAccounts: false
			})
		];

		await Promise.all(registrationPromises);

		// Unregister all providers
		const unregistrationPromises = [
			mainThreadAuthentication.$unregisterAuthenticationProvider('concurrent-1'),
			mainThreadAuthentication.$unregisterAuthenticationProvider('concurrent-2'),
			mainThreadAuthentication.$unregisterAuthenticationProvider('concurrent-3')
		];

		await Promise.all(unregistrationPromises);

		// Success if we reach here without timeout
		assert.ok(true, 'Concurrent registrations and unregistrations completed successfully');
	});

	test('never lets an extension send workbench-only session options to a provider', async () => {
		const received: unknown[] = [];
		rpcProtocol.set(ExtHostContext.ExtHostAuthentication, upcastPartial<ExtHostAuthenticationShape>({
			$getSessions: (_id, _scopes, options) => {
				received.push(options);
				return Promise.resolve([]);
			},
		}));
		await mainThreadAuthentication.$registerAuthenticationProvider({ id: 'test-reserved', label: 'Test Reserved', supportsMultipleAccounts: false });

		await mainThreadAuthentication.$getSession('test-reserved', ['scope'], 'publisher.extension', 'Extension', {
			silent: true,
			clientId: 'client',
			_workbenchIncludeUnapprovedAccounts: true,
			_workbenchEntraExchangeProbe: { subjectTokens: ['token'] }
		});
		await mainThreadAuthentication.$unregisterAuthenticationProvider('test-reserved');

		// Undefined options do not survive the trip to the extension host.
		assert.deepStrictEqual(received, [{ clientId: 'client', silent: true }]);
	});
});
