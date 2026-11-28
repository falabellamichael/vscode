/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import assert from 'assert';
import type { ProgressOptions } from 'vscode';
import { encodeBase64, VSBuffer } from '../../../../base/common/buffer.js';
import { CancellationToken } from '../../../../base/common/cancellation.js';
import { CancellationError, isCancellationError } from '../../../../base/common/errors.js';
import { Emitter } from '../../../../base/common/event.js';
import { IAuthorizationServerMetadata } from '../../../../base/common/oauth.js';
import { URI, UriComponents } from '../../../../base/common/uri.js';
import { mock } from '../../../../base/test/common/mock.js';
import { ensureNoDisposablesAreLeakedInTestSuite } from '../../../../base/test/common/utils.js';
import { ILogger, ILoggerService, NullLogger, NullLogService } from '../../../../platform/log/common/log.js';
import { INotificationSource } from '../../../../platform/notification/common/notification.js';
import { Progress } from '../../../../platform/progress/common/progress.js';
import { getOAuthTokenFingerprint, IAuthenticationProviderSessionOptions, MCP_ACCEPT_PARTIAL_OAUTH_SCOPES, MCP_REFRESH_REJECTED_OAUTH_TOKEN, MCP_USE_HTTPS_OAUTH_CALLBACK, signalCMailOAuth, signalCMailOAuthPrefix } from '../../../services/authentication/common/authentication.js';
import { DynamicAuthProvider, ExtHostAuthentication, IAuthorizationToken, reviveAccountIcon, TokenStore } from '../../common/extHostAuthentication.js';
import { IProgressStepDto, MainThreadAuthenticationShape } from '../../common/extHost.protocol.js';
import { IExtHostInitDataService } from '../../common/extHostInitDataService.js';
import { IExtHostProgress } from '../../common/extHostProgress.js';
import { IExtHostRpcService } from '../../common/extHostRpcService.js';
import { IExtHostUrlsService } from '../../common/extHostUrls.js';
import { IExtHostWindow } from '../../common/extHostWindow.js';
import { Proxied } from '../../../services/extensions/common/proxyIdentifier.js';

/** Builds a structurally-valid JWT carrying the given claims. */
function jwt(claims: object): string {
	const segment = (value: object) => encodeBase64(VSBuffer.fromString(JSON.stringify(value)));
	return `${segment({ alg: 'none', typ: 'JWT' })}.${segment(claims)}.signature`;
}

suite('TokenStore', () => {

	const disposables = ensureNoDisposablesAreLeakedInTestSuite();

	function createStore(initialTokens: IAuthorizationToken[]): TokenStore {
		const persistence = {
			onDidChange: disposables.add(new Emitter<IAuthorizationToken[]>()).event,
			set: () => { }
		};
		return disposables.add(new TokenStore(persistence, initialTokens, new NullLogger()));
	}

	// Regression for the MCP sign-in loop: an explicit empty `token.scope` must derive empty session scopes, not the granted scopes from the JWT claims, else empty-scope lookups never match their own session.
	test('derives session scopes from the stored token.scope, falling back to JWT claims only when scope is absent', () => {
		const store = createStore([
			// Explicit empty scope must win over the scopes embedded in the JWT claims.
			{ access_token: jwt({ sub: 'a', scope: 'menu:read orders:create orders:cancel' }), token_type: 'Bearer', scope: '', created_at: 0 },
			// Absent scope (undefined) falls back to the JWT claims.
			{ access_token: jwt({ sub: 'b', scope: 'menu:read orders:create' }), token_type: 'Bearer', created_at: 0 },
			// A non-empty scope is authoritative over the JWT claims.
			{ access_token: jwt({ sub: 'c', scope: 'ignored:claim' }), token_type: 'Bearer', scope: 'read write', created_at: 0 },
		]);

		assert.deepStrictEqual(
			store.sessions.map(session => session.scopes),
			[[], ['menu:read', 'orders:create'], ['read', 'write']]
		);
	});

	test('an omitted token list clears a revoked dynamic session and accepts its replacement', async () => {
		const host = new ExtHostAuthentication(
			new class extends mock<IExtHostRpcService>() {
				override getProxy<T>(): T { return {} as T; }
			}(),
			new class extends mock<IExtHostInitDataService>() { }(),
			new class extends mock<IExtHostWindow>() { }(),
			new class extends mock<IExtHostUrlsService>() { }(),
			new class extends mock<IExtHostProgress>() { }(),
			new class extends mock<ILoggerService>() { }(),
			new NullLogService(),
		);
		/* eslint-disable local/code-no-bracket-notation-for-identifiers -- Exercise the notification without exposing test-only APIs. */
		const emitter = disposables.add(host['_onDidDynamicAuthProviderTokensChange']);
		disposables.add(host['_onDidChangeSessions']);
		/* eslint-enable local/code-no-bracket-notation-for-identifiers */
		const changes = disposables.add(new Emitter<IAuthorizationToken[]>());
		disposables.add(emitter.event(event => changes.fire(event.tokens)));
		const initialToken: IAuthorizationToken = { access_token: 'fixture-original', token_type: 'Bearer', created_at: 0 };
		const store = disposables.add(new TokenStore({ onDidChange: changes.event, set: () => { } }, [initialToken], new NullLogger()));

		await host.$onDidChangeDynamicAuthProviderTokens('fixture', 'client');
		assert.strictEqual(store.sessions.length, 0);
		await host.$onDidChangeDynamicAuthProviderTokens('fixture', 'client', [{ ...initialToken, access_token: 'fixture-replacement' }]);
		assert.deepStrictEqual(store.sessions.map(session => session.accessToken), ['fixture-replacement']);
	});
});

suite('DynamicAuthProvider', () => {

	const disposables = ensureNoDisposablesAreLeakedInTestSuite();

	class TestDynamicAuthProvider extends DynamicAuthProvider {
		generateNewClientIdCalls = 0;

		exchangeTestCode(code: string, verifier: string, redirectUri: string): ReturnType<DynamicAuthProvider['exchangeCodeForToken']> {
			return this.exchangeCodeForToken(code, verifier, redirectUri);
		}

		protected override async _generateNewClientId(): Promise<void> {
			this.generateNewClientIdCalls++;
		}
	}

	interface ITestServer {
		readonly authorizationServer: string;
		readonly serverMetadata: IAuthorizationServerMetadata;
	}

	const resource = 'https://workiq.example.com/mcp';

	const servers: Record<string, ITestServer> = {
		entra: {
			authorizationServer: 'https://login.microsoftonline.com/organizations/v2.0',
			serverMetadata: {
				issuer: 'https://login.microsoftonline.com/organizations/v2.0',
				authorization_endpoint: 'https://login.microsoftonline.com/organizations/oauth2/v2.0/authorize',
				token_endpoint: 'https://login.microsoftonline.com/organizations/oauth2/v2.0/token',
				response_types_supported: ['code'],
				scopes_supported: ['openid', 'profile', 'offline_access'],
			},
		},
		nonMicrosoft: {
			authorizationServer: 'https://auth.example.com',
			serverMetadata: {
				issuer: 'https://auth.example.com',
				authorization_endpoint: 'https://auth.example.com/authorize',
				token_endpoint: 'https://auth.example.com/token',
				response_types_supported: ['code'],
				scopes_supported: ['openid', 'profile', 'offline_access'],
			},
		},
		// A Microsoft authorization endpoint must not opt a non-Microsoft token endpoint into Microsoft behavior.
		entraWithForeignTokenEndpoint: {
			authorizationServer: 'https://login.microsoftonline.com/organizations/v2.0',
			serverMetadata: {
				issuer: 'https://login.microsoftonline.com/organizations/v2.0',
				authorization_endpoint: 'https://login.microsoftonline.com/organizations/oauth2/v2.0/authorize',
				token_endpoint: 'https://auth.example.com/token',
				response_types_supported: ['code'],
				scopes_supported: ['openid', 'profile', 'offline_access'],
			},
		},
	};

	function createProvider(options: ITestServer & {
		readonly resource: string;
		readonly initialTokens?: IAuthorizationToken[];
		readonly providerId?: string;
		readonly clientId?: string;
		readonly clientSecret?: string;
		readonly tokenEndpointAuthMethod?: string;
		readonly signalCOAuthRedirectUri?: string;
		readonly fetcher?: typeof fetch;
		readonly trace?: string[];
		readonly onDidOpenUri?: (uri: string) => void;
		readonly signInCode?: string;
		readonly onBrowserCallback?: (expected: URI) => URI;
	}): TestDynamicAuthProvider {
		const loggerService = new class extends mock<ILoggerService>() {
			override createLogger(): ILogger {
				return new class extends NullLogger {
					override trace(message: string): void { options.trace?.push(message); }
				}();
			}
		}();
		const proxy = new class extends mock<Proxied<MainThreadAuthenticationShape>>() {
			override $setSessionsForDynamicAuthProvider = (): Promise<void> => Promise.resolve();
			// Never settles, so an abandoned authorization flow leaves no rejected promise behind.
			override $waitForUriHandler = (expected: UriComponents): Promise<UriComponents> => options.onBrowserCallback ? Promise.resolve(options.onBrowserCallback(URI.revive(expected))) : options.signInCode
				? Promise.resolve(URI.parse(`vscode://dynamicauthprovider/authorize?code=${options.signInCode}`))
				: new Promise<UriComponents>(() => { });
		}();
		const extHostWindow = new class extends mock<IExtHostWindow>() {
			override async openUri(uri: string | URI): Promise<boolean> {
				options.onDidOpenUri?.(uri.toString());
				return !!(options.signInCode || options.onBrowserCallback);
			}
		}();
		const extHostUrls = new class extends mock<IExtHostUrlsService>() {
			override async createAppUri(uri: URI): Promise<URI> {
				return uri;
			}
		}();
		const initData = new class extends mock<IExtHostInitDataService>() {
			override environment = { appUriScheme: options.signalCOAuthRedirectUri ? 'signal-c' : 'vscode', appName: 'Test' } as IExtHostInitDataService['environment'];
		}();
		const extHostProgress = new class extends mock<IExtHostProgress>() {
			override async withProgressFromSource<R>(_source: string | INotificationSource, _options: ProgressOptions, task: (progress: Progress<IProgressStepDto>, token: CancellationToken) => Thenable<R>): Promise<R> {
				return task(new Progress<IProgressStepDto>(() => { }), CancellationToken.None);
			}
		}();
		return disposables.add(new TestDynamicAuthProvider(
			extHostWindow,
			extHostUrls,
			initData,
			extHostProgress,
			loggerService,
			proxy,
			URI.parse(options.authorizationServer),
			options.serverMetadata,
			{ resource: options.resource },
			options.clientId ?? 'client-id',
			options.clientSecret,
			disposables.add(new Emitter()),
			options.initialTokens ?? [],
			options.providerId,
			options.tokenEndpointAuthMethod,
			options.signalCOAuthRedirectUri ? { redirectUri: options.signalCOAuthRedirectUri, windowId: 42 } : undefined,
			options.fetcher,
		));
	}

	test('exchanges an HTTPS browser return with the exact registered redirect and original code', async () => {
		const redirectUri = 'https://callback.example.com/oauth/return';
		const code = 'fixture+&%=?/unicode-☃';
		const scopes = ['read'];
		let nonce: string | null = null;
		let challenge: string | null = null;
		const accessToken = jwt({ sub: 'https-fixture-account' });
		const provider = createProvider({
			...servers.nonMicrosoft,
			resource,
			signalCOAuthRedirectUri: redirectUri,
			onBrowserCallback: expected => {
				assert.strictEqual(expected.toString(true).split('?')[0], 'signal-c://dynamicauthprovider/https/authorize');
				const query = new URLSearchParams(expected.query);
				nonce = query.get('nonce');
				assert.strictEqual(query.get('windowId'), '42');
				return expected.with({ query: `${expected.query}&code64=${encodeBase64(VSBuffer.fromString(code), false, true)}` });
			},
			onDidOpenUri: value => {
				const authorize = new URL(value);
				assert.match(nonce ?? '', /^[a-f0-9]{32}$/);
				assert.strictEqual(authorize.searchParams.get('state'), `${nonce}.42`);
				assert.strictEqual(authorize.searchParams.get('redirect_uri'), redirectUri);
				assert.strictEqual(authorize.searchParams.get('code_challenge_method'), 'S256');
				challenge = authorize.searchParams.get('code_challenge');
			},
			fetcher: async (url, init) => {
				assert.strictEqual(url, servers.nonMicrosoft.serverMetadata.token_endpoint);
				assert.strictEqual(init?.redirect, 'error');
				const params = new URLSearchParams(String(init?.body));
				assert.strictEqual(params.get('code'), code);
				assert.strictEqual(params.get('redirect_uri'), redirectUri);
				const hash = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(params.get('code_verifier') ?? ''));
				assert.strictEqual(encodeBase64(VSBuffer.wrap(new Uint8Array(hash)), false, true), challenge);
				return Response.json({ access_token: accessToken, token_type: 'Bearer', scope: 'read' });
			},
		});
		const session = await provider.createSession(scopes, { [MCP_USE_HTTPS_OAUTH_CALLBACK]: true });
		assert.strictEqual(session.accessToken, accessToken);
	});

	test('fails before opening a browser when the selected HTTPS callback is unconfigured', async () => {
		const provider = createProvider({ ...servers.nonMicrosoft, resource, onDidOpenUri: () => assert.fail('No browser should open') });
		await assert.rejects(provider.createSession(['read'], { [MCP_USE_HTTPS_OAUTH_CALLBACK]: true }), /cannot use the Signal-C HTTPS OAuth callback/);
	});

	test('honors an explicit provider ID without changing the default identity', () => {
		const providerId = 'https://auth.example.com/ https://workiq.example.com/mcp client:client-id';
		const scoped = createProvider({ ...servers.nonMicrosoft, resource, providerId });
		const legacy = createProvider({ ...servers.nonMicrosoft, resource });

		assert.deepStrictEqual([scoped.id, legacy.id], [providerId, 'https://auth.example.com/ https://workiq.example.com/mcp']);
	});

	test('keeps a browser grant with fewer scopes for MCP and prefers the full grant once consent is completed', async () => {
		const scopes = ['read', 'write'];
		const options = { [MCP_ACCEPT_PARTIAL_OAUTH_SCOPES]: true };
		let responseScopes = 'read';
		let requests = 0;
		const provider = createProvider({
			...servers.nonMicrosoft, resource, signInCode: 'test-code',
			onDidOpenUri: uri => assert.strictEqual(new URL(uri).searchParams.get('scope'), 'read write offline_access'),
			fetcher: async () => Response.json({
				access_token: jwt({ sub: 'account', jti: String(++requests) }), token_type: 'Bearer', scope: responseScopes,
			}),
		});

		const partial = await provider.createSession(scopes, options);
		assert.deepStrictEqual(partial.scopes, ['read']);
		assert.deepStrictEqual(await provider.getSessions(scopes, { silent: true, ...options }), [partial]);
		assert.deepStrictEqual(await provider.getSessions(['read'], { silent: true }), [partial]);
		assert.deepStrictEqual(await provider.getSessions(scopes, { silent: true }), []);
		await assert.rejects(provider.createSession(scopes, {}), /did not grant the required permissions \(write\)/);
		assert.deepStrictEqual(await provider.getSessions(scopes, options), [partial]);

		responseScopes = 'read write offline_access';
		const complete = await provider.createSession(scopes, {});
		assert.deepStrictEqual(complete.scopes, ['read', 'write', 'offline_access']);
		assert.deepStrictEqual(await provider.getSessions(scopes, options), [complete, partial]);
		assert.deepStrictEqual(await provider.getSessions(scopes, {}), [complete]);
	});

	test('keeps a reduced grant on refresh and inherits it when the next response omits scope', async () => {
		const scopes = ['read', 'write'];
		const options = { silent: true, [MCP_ACCEPT_PARTIAL_OAUTH_SCOPES]: true };
		let refreshes = 0;
		const provider = createProvider({
			...servers.nonMicrosoft, resource,
			initialTokens: [{ access_token: 'expired', token_type: 'Bearer', expires_in: 1, created_at: 0, refresh_token: 'fixture-refresh', scope: 'read write' }],
			fetcher: async () => Response.json({
				access_token: `refreshed-${++refreshes}`, token_type: 'Bearer', expires_in: 1,
				...(refreshes === 1 ? { scope: 'read' } : {}),
			}),
		});
		for (let attempt = 1; attempt <= 2; attempt++) {
			const [session] = await provider.getSessions(scopes, options);
			assert.deepStrictEqual({ token: session?.accessToken, scopes: session?.scopes }, { token: `refreshed-${attempt}`, scopes: ['read'] });
			assert.deepStrictEqual(await provider.getSessions(scopes, { silent: true }), []);
		}
	});

	test('refreshes only the rejected access token and shares the result with ordinary lookups', async () => {
		const primary = { access_token: 'rejected-before-expiry', token_type: 'Bearer', scope: 'read', expires_in: 3600, created_at: Date.now(), refresh_token: 'support-refresh', id_token: jwt({ sub: 'support', preferred_username: 'Support' }) };
		const sibling = { ...primary, access_token: 'other-account', refresh_token: 'other-refresh', id_token: jwt({ sub: 'other', preferred_username: 'Other' }) };
		const refreshes: (string | null)[] = [];
		const provider = createProvider({
			...servers.nonMicrosoft, resource, initialTokens: [primary, sibling],
			fetcher: async (url, init) => {
				assert.strictEqual(url, servers.nonMicrosoft.serverMetadata.token_endpoint);
				const params = new URLSearchParams(String(init?.body));
				assert.strictEqual(params.get('grant_type'), 'refresh_token');
				assert.strictEqual(params.get('resource'), resource);
				refreshes.push(params.get('refresh_token'));
				return Response.json({ access_token: 'replacement-token', token_type: 'Bearer', expires_in: 3600, refresh_token: 'rotated-refresh' });
			},
		});
		const options = { [MCP_REFRESH_REJECTED_OAUTH_TOKEN]: await getOAuthTokenFingerprint(primary.access_token) };
		const [first, duplicate] = await Promise.all([provider.getSessions(['read'], options), provider.getSessions(['read'], options)]);
		assert.deepStrictEqual(first.map(session => [session.account.id, session.accessToken, session.scopes]), [['support', 'replacement-token', ['read']]]);
		assert.deepStrictEqual(duplicate, []);
		assert.deepStrictEqual(refreshes, ['support-refresh']);
		assert.deepStrictEqual((await provider.getSessions(['read'], {})).map(session => session.accessToken).sort(), ['other-account', 'replacement-token']);
	});

	for (const { name, response, status } of [
		{ name: 'refuses to switch accounts', response: { access_token: 'replacement', token_type: 'Bearer', id_token: jwt({ sub: 'other' }) }, status: 200 },
		{ name: 'refuses to expand the grant', response: { access_token: 'replacement', token_type: 'Bearer', scope: 'read write' }, status: 200 },
		{ name: 'leaves the grant available for browser recovery when the client is invalid', response: { error: 'invalid_client' }, status: 400 },
		{ name: 'does not accept a token on a failed refresh response', response: { access_token: 'replacement', token_type: 'Bearer', scope: 'read' }, status: 503 },
	]) {
		test(`rejected access token refresh ${name}`, async () => {
			const token = { access_token: 'rejected-access', token_type: 'Bearer', scope: 'read', expires_in: 3600, created_at: Date.now(), refresh_token: 'original-refresh', id_token: jwt({ sub: 'support' }) };
			const provider = createProvider({
				...servers.nonMicrosoft, resource, initialTokens: [token], fetcher: async () => Response.json(response, { status }),
			});
			assert.deepStrictEqual(await provider.getSessions(['read'], { [MCP_REFRESH_REJECTED_OAUTH_TOKEN]: await getOAuthTokenFingerprint(token.access_token) }), []);
			assert.deepStrictEqual((await provider.getSessions(['read'], {})).map(session => session.accessToken), [token.access_token]);
			assert.strictEqual(provider.generateNewClientIdCalls, 0);
		});
	}

	test('reuses provider default scopes only for the matching empty request', async () => {
		const provider = createProvider({
			...servers.nonMicrosoft, resource, signInCode: 'test-code',
			onDidOpenUri: uri => assert.strictEqual(new URL(uri).searchParams.get('scope'), null),
			fetcher: async () => Response.json({ access_token: 'default-token', token_type: 'Bearer', scope: 'default:read default:write' }),
		});
		const grant = await provider.createSession([], {});
		assert.deepStrictEqual(grant.scopes, ['default:read', 'default:write']);
		assert.deepStrictEqual(await provider.getSessions([], { silent: true }), [grant]);
	});

	test('does not rotate the client while silently refreshing a token', async () => {
		let fetchCalls = 0;
		const fetcher: typeof fetch = async () => {
			fetchCalls++;
			return new Response(JSON.stringify({ error: 'invalid_client' }), {
				status: 400,
				headers: { 'Content-Type': 'application/json' },
			});
		};
		const provider = createProvider({
			authorizationServer: 'https://mcp.example.com',
			serverMetadata: {
				issuer: 'https://mcp.example.com',
				response_types_supported: ['code'],
				token_endpoint: 'https://mcp.example.com/token',
			},
			resource: 'https://mcp.example.com/resource',
			initialTokens: [{
				access_token: jwt({ sub: 'account' }),
				token_type: 'Bearer',
				scope: '',
				expires_in: 1,
				refresh_token: 'refresh-token',
				created_at: 0,
			}],
			fetcher,
		});

		const sessions = await provider.getSessions([], { silent: true } satisfies IAuthenticationProviderSessionOptions);

		assert.deepStrictEqual({
			sessions,
			fetchCalls,
			generateNewClientIdCalls: provider.generateNewClientIdCalls,
			clientId: provider.clientId,
		}, {
			sessions: [],
			fetchCalls: 1,
			generateNewClientIdCalls: 0,
			clientId: 'client-id',
		});
	});

	for (const rotate of [false, true]) {
		test(`serializes OAuth refresh with ${rotate ? 'a replacement' : 'the original'} refresh token`, async () => {
			const refreshes: (string | null)[] = [];
			const provider = createProvider({
				...servers.nonMicrosoft,
				resource,
				initialTokens: [{
					access_token: 'expired-token', token_type: 'Bearer', scope: 'read', expires_in: 1,
					created_at: 0, refresh_token: 'original-refresh', id_token: jwt({ sub: 'fixture-account' }),
				}],
				fetcher: async (_input, init) => {
					const params = new URLSearchParams(String(init?.body));
					refreshes.push(params.get('refresh_token'));
					const first = refreshes.length === 1;
					return Response.json({
						access_token: first ? 'first-access' : 'second-access', token_type: 'Bearer',
						expires_in: first ? 1 : 3600, ...(first && rotate ? { refresh_token: 'rotated-refresh' } : {}),
					});
				},
			});
			const results = await Promise.all([provider.getSessions(['read'], { silent: true }), provider.getSessions(['read'], { silent: true }), provider.getSessions(['read'], { silent: true })]);
			assert.deepStrictEqual({ refreshes, sessions: results.map(([session]) => [session?.accessToken, session?.account.id]) }, {
				refreshes: ['original-refresh', rotate ? 'rotated-refresh' : 'original-refresh'],
				sessions: [['first-access', 'fixture-account'], ['second-access', 'fixture-account'], ['second-access', 'fixture-account']],
			});
		});
	}

	test('refresh sends scope without resource for Microsoft auth, and resource without scope otherwise', async () => {
		const results: Record<string, { body: Record<string, string>; scopes: (readonly string[])[] }> = {};
		for (const [name, server] of Object.entries(servers)) {
			const bodies: Record<string, string>[] = [];
			const fetcher: typeof fetch = async (_input, init) => {
				bodies.push(Object.fromEntries(new URLSearchParams(String(init?.body))));
				return new Response(JSON.stringify({
					access_token: jwt({ sub: 'account' }),
					token_type: 'Bearer',
					expires_in: 3600,
					refresh_token: 'new-refresh',
				}), {
					status: 200,
					headers: { 'Content-Type': 'application/json' },
				});
			};
			const provider = createProvider({
				...server,
				resource,
				initialTokens: [{
					access_token: jwt({ sub: 'account', jti: 'expired' }),
					token_type: 'Bearer',
					scope: 'api://res/Read',
					expires_in: 1,
					refresh_token: 'refresh-token',
					created_at: 0,
				}],
				fetcher,
			});

			const sessions = await provider.getSessions(['api://res/Read'], { silent: true } satisfies IAuthenticationProviderSessionOptions);
			results[name] = { body: bodies[0], scopes: sessions.map(session => session.scopes) };
		}

		const refreshBase = { client_id: 'client-id', grant_type: 'refresh_token', refresh_token: 'refresh-token' };
		assert.deepStrictEqual(results, {
			entra: {
				body: { ...refreshBase, scope: 'api://res/Read offline_access' },
				scopes: [['api://res/Read']],
			},
			nonMicrosoft: {
				body: { ...refreshBase, resource },
				scopes: [['api://res/Read']],
			},
			entraWithForeignTokenEndpoint: {
				body: { ...refreshBase, resource },
				scopes: [['api://res/Read']],
			},
		});
	});

	for (const { name, methods, clientSecret, useBasic, registeredMethod } of [
		{ name: 'HTTP Basic', methods: ['client_secret_basic'], clientSecret: 'secret+%:雪', useBasic: true },
		{ name: 'form credentials', methods: ['client_secret_basic', 'client_secret_post'], clientSecret: 'legacy-secret', useBasic: false },
		{ name: 'legacy discovery', methods: undefined, clientSecret: 'legacy-secret', useBasic: false },
		{ name: 'public client', methods: ['client_secret_basic', 'none'], clientSecret: undefined, useBasic: false },
		{ name: 'registered HTTP Basic when both methods are advertised', methods: ['client_secret_basic', 'client_secret_post'], clientSecret: 'secret+%:雪', useBasic: true, registeredMethod: 'client_secret_basic' },
		{ name: 'registered form credentials when discovery lists only HTTP Basic', methods: ['client_secret_basic'], clientSecret: 'registered-secret', useBasic: false, registeredMethod: 'client_secret_post' },
		{ name: 'registered public client even with a returned secret', methods: ['client_secret_basic', 'none'], clientSecret: 'unused-secret', useBasic: false, registeredMethod: 'none' },
	]) {
		test(`uses ${name} for both code exchange and refresh without tracing credentials`, async () => {
			const clientId = 'example:id +';
			const trace: string[] = [];
			const requests: { url: string; method: string | undefined; authorization: string | null; body: Record<string, string>; redirect: RequestRedirect | undefined }[] = [];
			const tokenEndpoint = 'https://auth.example.com/token';
			const accessToken = jwt({ sub: 'account' });
			const provider = createProvider({
				...servers.nonMicrosoft,
				serverMetadata: { ...servers.nonMicrosoft.serverMetadata, token_endpoint_auth_methods_supported: methods },
				resource,
				clientId,
				clientSecret,
				tokenEndpointAuthMethod: registeredMethod,
				trace,
				initialTokens: [{ access_token: jwt({ sub: 'account', jti: 'expired' }), token_type: 'Bearer', scope: 'read', expires_in: 1, refresh_token: 'test-refresh-token', created_at: 0 }],
				fetcher: async (url, init) => {
					requests.push({ url: String(url), method: init?.method, authorization: new Headers(init?.headers).get('Authorization'), body: Object.fromEntries(new URLSearchParams(String(init?.body))), redirect: init?.redirect });
					return Response.json({ access_token: accessToken, token_type: 'Bearer', scope: 'read', expires_in: 3600 });
				},
			});

			const token = await provider.exchangeTestCode('test-code', 'test-verifier', 'http://127.0.0.1:40000/callback');
			const refreshed = await provider.getSessions(['read'], { silent: true });
			assert.strictEqual(token.access_token, accessToken);
			assert.deepStrictEqual(refreshed.map(session => session.accessToken), [accessToken]);

			const authorization = useBasic ? `Basic ${encodeBase64(VSBuffer.fromString('example%3Aid+%2B:secret%2B%25%3A%E9%9B%AA'))}` : null;
			const sentSecret = registeredMethod === 'none' ? undefined : clientSecret;
			const credentials = useBasic ? {} : { client_id: clientId, ...(sentSecret ? { client_secret: sentSecret } : {}) };
			const request = { url: tokenEndpoint, method: 'POST', authorization, redirect: sentSecret ? 'error' : undefined };
			assert.deepStrictEqual(requests, [
				{ ...request, body: { ...credentials, grant_type: 'authorization_code', code: 'test-code', code_verifier: 'test-verifier', redirect_uri: 'http://127.0.0.1:40000/callback', resource } },
				{ ...request, body: { ...credentials, grant_type: 'refresh_token', refresh_token: 'test-refresh-token', resource } },
			]);
			assert.ok(!trace.some(message => /test-code|test-verifier|test-refresh-token|client_secret|Basic /.test(message)));
		});
	}

	for (const { method, clientSecret, expected } of [
		{ method: 'private_key_jwt', clientSecret: 'unused-secret', expected: /unsupported token authentication method/ },
		{ method: 'client_secret_basic', clientSecret: undefined, expected: /requires a client secret/ },
		{ method: 'client_secret_post', clientSecret: undefined, expected: /requires a client secret/ },
	]) {
		test(`rejects unusable registration for ${method} before sending the browser grant`, async () => {
			const provider = createProvider({
				...servers.nonMicrosoft,
				resource,
				clientSecret,
				tokenEndpointAuthMethod: method,
				fetcher: async () => assert.fail('No token request should be sent'),
			});
			await assert.rejects(provider.exchangeTestCode('test-code', 'test-verifier', 'http://127.0.0.1/callback'), expected);
		});
	}

	for (const method of ['client_secret_basic', 'client_secret_post']) {
		test(`refuses to send ${method} credentials to a plaintext token endpoint`, async () => {
			const provider = createProvider({
				...servers.nonMicrosoft,
				serverMetadata: { ...servers.nonMicrosoft.serverMetadata, token_endpoint: 'http://auth.example.com/token', token_endpoint_auth_methods_supported: [method] },
				resource,
				clientSecret: 'test-secret',
				fetcher: async () => assert.fail('No token request should be sent'),
			});

			await assert.rejects(provider.exchangeTestCode('test-code', 'test-verifier', 'http://127.0.0.1/callback'), /requires an HTTPS token endpoint/);
		});
	}

	test('authorization URL requests offline_access when advertised and omits resource for Microsoft auth', async () => {
		const results: Record<string, { scope: string | null; resource: string | null; prompt: string | null }> = {};
		for (const [name, server] of Object.entries(servers)) {
			const openedUris: string[] = [];
			const provider = createProvider({
				...server,
				resource,
				onDidOpenUri: uri => openedUris.push(uri),
			});

			await assert.rejects(provider.createSession(['api://res/Read'], {}), isCancellationError);

			const params = new URL(openedUris[0]).searchParams;
			results[name] = {
				scope: params.get('scope'),
				resource: params.get('resource'),
				prompt: params.get('prompt'),
			};
		}

		assert.deepStrictEqual(results, {
			entra: { scope: 'api://res/Read offline_access', resource: null, prompt: null },
			nonMicrosoft: { scope: 'api://res/Read offline_access', resource, prompt: null },
			entraWithForeignTokenEndpoint: { scope: 'api://res/Read offline_access', resource, prompt: null },
		});
	});

	test('Google Workspace requires the client secret before opening sign-in and requests refresh access', async () => {
		const options = {
			authorizationServer: 'https://accounts.google.com',
			serverMetadata: {
				issuer: 'https://accounts.google.com',
				authorization_endpoint: 'https://accounts.google.com/o/oauth2/v2/auth',
				token_endpoint: 'https://oauth2.googleapis.com/token',
				response_types_supported: ['code'],
			},
			resource: 'https://gmailmcp.googleapis.com/mcp/v1',
		};
		const provider = createProvider({ ...options, onDidOpenUri: () => assert.fail('The client secret is required before opening the browser') });
		await assert.rejects(provider.createSession(['read'], {}), /Set OAuth Client Secret/);

		const opened: URL[] = [];
		const requests: RequestInit[] = [];
		const registered = createProvider({
			...options,
			clientSecret: 'fixture-secret',
			onDidOpenUri: uri => opened.push(new URL(uri)),
			fetcher: async (_input, init) => {
				assert.ok(init);
				requests.push(init);
				return Response.json({ access_token: 'fixture-token', token_type: 'Bearer', refresh_token: 'fixture-refresh-token' });
			},
		});
		await assert.rejects(registered.createSession(['read'], {}), isCancellationError);
		assert.strictEqual(opened.length, 1);
		assert.strictEqual(opened[0].searchParams.get('access_type'), 'offline');
		assert.strictEqual(opened[0].searchParams.get('prompt'), 'consent');
		assert.strictEqual(opened[0].searchParams.get('scope'), 'read openid');
		assert.ok(!opened[0].searchParams.has('client_secret'));
		await registered.exchangeTestCode('fixture-code', 'fixture-verifier', 'http://127.0.0.1:33418/');
		assert.strictEqual(requests[0].redirect, 'error');
		assert.strictEqual(new URLSearchParams(String(requests[0].body)).get('client_secret'), 'fixture-secret');

		const refreshes: string[] = [];
		const expiring = createProvider({
			...options,
			clientSecret: 'fixture-secret',
			initialTokens: [{
				access_token: 'expired-token', token_type: 'Bearer', scope: 'read',
				expires_in: 1, created_at: 0, refresh_token: 'original-refresh-token',
			}],
			fetcher: async (_input, init) => {
				const request = new URLSearchParams(String(init?.body));
				assert.strictEqual(request.get('grant_type'), 'refresh_token');
				refreshes.push(request.get('refresh_token') ?? '');
				return Response.json({ access_token: `access-token-${refreshes.length}`, token_type: 'Bearer', scope: 'read', expires_in: 1 });
			},
		});
		for (let attempt = 1; attempt <= 2; attempt++) {
			const sessions = await expiring.getSessions(['read'], { silent: true });
			assert.strictEqual(sessions[0]?.accessToken, `access-token-${attempt}`);
		}
		assert.deepStrictEqual(refreshes, ['original-refresh-token', 'original-refresh-token']);
	});

	for (const mail of ['gmail', 'outlook'] as const) {
		test(`direct ${mail} keeps mailbox identity and renewable authorization when a token is refreshed`, async () => {
			const { issuer, authorize, token, resource, scope } = signalCMailOAuth[mail];
			const scopes = [scope, 'openid'];
			const refreshes: string[] = [];
			const provider = createProvider({
				authorizationServer: issuer,
				serverMetadata: { issuer, authorization_endpoint: authorize, token_endpoint: token, response_types_supported: ['code'] },
				resource,
				providerId: `${signalCMailOAuthPrefix}${mail}:isolated`,
				clientSecret: mail === 'gmail' ? 'fixture-secret' : undefined,
				onDidOpenUri: () => assert.fail('Direct browser sign-in needs a local desktop host'),
				initialTokens: [{
					access_token: 'expired-token', token_type: 'Bearer', scope: scopes.join(' '), id_token: jwt({ sub: 'mail-account' }),
					expires_in: 1, created_at: 0, refresh_token: 'fixture-refresh',
				}],
				fetcher: async (input, init) => {
					assert.strictEqual(input, token);
					assert.strictEqual(init?.redirect, 'error');
					const body = new URLSearchParams(String(init?.body));
					assert.ok(!body.has('resource'));
					assert.strictEqual(body.get('client_secret'), mail === 'gmail' ? 'fixture-secret' : null);
					refreshes.push(body.get('refresh_token') ?? '');
					return Response.json({ access_token: `fixture-access-${refreshes.length}`, token_type: 'Bearer', expires_in: 1 });
				},
			});
			await assert.rejects(provider.createSession(scopes, {}), /local Signal-C desktop extension host/);
			for (let attempt = 1; attempt <= 2; attempt++) {
				const [session] = await provider.getSessions(scopes, { silent: true });
				assert.strictEqual(session?.account.id, 'mail-account');
				assert.strictEqual(session?.accessToken, `fixture-access-${attempt}`);
			}
			assert.deepStrictEqual(refreshes, ['fixture-refresh', 'fixture-refresh']);
		});
	}

	test('keeps the refreshed Outlook grant read-only when Graph returns Mail.Read without Mail.Send', async () => {
		const { issuer, authorize, token, resource, scope, sendScope } = signalCMailOAuth.outlook;
		const identity = ['openid', 'profile', 'offline_access'];
		const requested = [scope, sendScope, ...identity];
		const provider = createProvider({
			authorizationServer: issuer,
			serverMetadata: { issuer, authorization_endpoint: authorize, token_endpoint: token, response_types_supported: ['code'] },
			resource,
			providerId: `${signalCMailOAuthPrefix}outlook:isolated`,
			initialTokens: [{
				access_token: 'expired-token', token_type: 'Bearer', scope: requested.join(' '),
				id_token: jwt({ sub: 'mail-account' }), expires_in: 1, created_at: 0, refresh_token: 'fixture-refresh',
			}],
			fetcher: async () => Response.json({ access_token: 'read-only-token', token_type: 'Bearer', scope: 'Mail.Read', expires_in: 3600 }),
		});
		assert.deepStrictEqual(await provider.getSessions(requested, { silent: true }), []);
		const [readGrant] = await provider.getSessions([scope, ...identity], { silent: true });
		assert.strictEqual(readGrant?.account.id, 'mail-account');
		assert.deepStrictEqual(readGrant.scopes, [scope, ...identity]);
	});

	for (const failure of [new CancellationError(), new Error('Browser authorization failed')]) {
		test(`preserves the final authentication flow error: ${failure.message}`, async () => {
			const provider = createProvider({
				...servers.nonMicrosoft,
				resource,
				onDidOpenUri: () => { throw failure; },
			});
			await assert.rejects(provider.createSession(['read'], {}), error => error === failure);
		});
	}
});

suite('Account Icon Revival', () => {

	ensureNoDisposablesAreLeakedInTestSuite();

	const iconComponents: UriComponents = { scheme: 'https', authority: 'example.com', path: '/avatar.png', query: '', fragment: '' };

	test('reviveAccountIcon revives a present icon into a URI and leaves a missing icon undefined', () => {
		const withIcon: { id: string; label: string; icon?: UriComponents } = { id: 'account-with-icon', label: 'Has Icon', icon: iconComponents };
		const withoutIcon: { id: string; label: string; icon?: UriComponents } = { id: 'account-without-icon', label: 'No Icon' };

		assert.deepStrictEqual(
			[reviveAccountIcon(withIcon), reviveAccountIcon(withoutIcon)],
			[
				{ ...withIcon, icon: URI.from(iconComponents) },
				{ ...withoutIcon, icon: undefined }
			]
		);
	});
});
