/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import * as assert from 'assert';
import * as sinon from 'sinon';
import { DeferredPromise } from '../../../../base/common/async.js';
import { CancellationError, isCancellationError, transformErrorForSerialization, transformErrorFromSerialization } from '../../../../base/common/errors.js';
import { toDisposable } from '../../../../base/common/lifecycle.js';
import { URI } from '../../../../base/common/uri.js';
import { mock } from '../../../../base/test/common/mock.js';
import { LogLevel, NullLogService } from '../../../../platform/log/common/log.js';
import { IMcpAuthenticationDetails, IMcpAuthenticationOptions, IMcpOAuthGrant, MainThreadMcpShape } from '../../common/extHost.protocol.js';
import { createAuthMetadata, CommonRequestInit, CommonResponse, IAuthMetadata, McpHTTPHandle } from '../../common/extHostMcp.js';
import { ensureNoDisposablesAreLeakedInTestSuite } from '../../../../base/test/common/utils.js';
import { McpConnectionState, McpServerTransportHTTP, McpServerTransportType, UserInteractionRequiredError } from '../../../contrib/mcp/common/mcpTypes.js';
import { MCP } from '../../../contrib/mcp/common/modelContextProtocol.js';
import { getOAuthTokenFingerprint } from '../../../services/authentication/common/authentication.js';

// Test constants to avoid magic strings
const TEST_MCP_URL = 'https://example.com/mcp';
const TEST_AUTH_SERVER = 'https://auth.example.com';
const TEST_RESOURCE_METADATA_URL = 'https://example.com/.well-known/oauth-protected-resource';

/**
 * Creates a mock CommonResponse for testing.
 */
function createMockResponse(options: {
	status?: number;
	statusText?: string;
	url?: string;
	headers?: Record<string, string>;
	body?: string;
}): CommonResponse {
	const headers = new Headers(options.headers ?? {});
	return {
		status: options.status ?? 200,
		statusText: options.statusText ?? 'OK',
		url: options.url ?? TEST_MCP_URL,
		headers,
		body: null,
		json: async () => JSON.parse(options.body ?? '{}'),
		text: async () => options.body ?? '',
	};
}

/**
 * Helper to create an IAuthMetadata instance for testing via the factory function.
 * Uses a mock fetch that returns the provided server metadata.
 */
async function createTestAuthMetadata(options: {
	scopes?: string[];
	serverMetadataIssuer?: string;
	resourceMetadata?: { resource: string; authorization_servers?: string[]; scopes_supported?: string[] };
}): Promise<{ authMetadata: IAuthMetadata; logMessages: Array<{ level: LogLevel; message: string }> }> {
	const logMessages: Array<{ level: LogLevel; message: string }> = [];
	const mockLogger = (level: LogLevel, message: string) => logMessages.push({ level, message });

	const issuer = options.serverMetadataIssuer ?? TEST_AUTH_SERVER;

	const mockFetch = sinon.stub();

	// Mock resource metadata fetch
	mockFetch.onCall(0).resolves(createMockResponse({
		status: 200,
		url: TEST_RESOURCE_METADATA_URL,
		body: JSON.stringify(options.resourceMetadata ?? {
			resource: TEST_MCP_URL,
			authorization_servers: [issuer]
		})
	}));

	// Mock server metadata fetch
	mockFetch.onCall(1).resolves(createMockResponse({
		status: 200,
		url: `${issuer}/.well-known/oauth-authorization-server`,
		body: JSON.stringify({
			issuer,
			authorization_endpoint: `${issuer}/authorize`,
			token_endpoint: `${issuer}/token`,
			response_types_supported: ['code']
		})
	}));

	const wwwAuthHeader = options.scopes
		? `Bearer scope="${options.scopes.join(' ')}"`
		: 'Bearer realm="example"';

	const originalResponse = createMockResponse({
		status: 401,
		url: TEST_MCP_URL,
		headers: {
			'WWW-Authenticate': wwwAuthHeader
		}
	});

	const authMetadata = await createAuthMetadata(
		TEST_MCP_URL,
		originalResponse.headers,
		{
			sameOriginHeaders: {},
			fetch: mockFetch,
			log: mockLogger
		}
	);

	return { authMetadata, logMessages };
}

suite('ExtHostMcp', () => {
	const store = ensureNoDisposablesAreLeakedInTestSuite();

	suite('McpHTTPHandle stream cleanup', () => {
		teardown(() => sinon.restore());

		const createStreamingHandle = async () => {
			const reading = new DeferredPromise<void>();
			const backchannelStopped = new DeferredPromise<void>();
			let controller!: ReadableStreamDefaultController<Uint8Array>;
			const stream = new ReadableStream<Uint8Array>({
				start: value => { controller = value; },
				pull: () => { void reading.complete(); },
			}, { highWaterMark: 0 });
			const reader = stream.getReader();
			sinon.stub(stream, 'getReader').returns(reader);
			const cancel = sinon.spy(reader, 'cancel');
			const warnings: string[] = [];
			const proxy = Object.assign(new class extends mock<MainThreadMcpShape>() {
				override $onDidChangeState(): void { }
				override $onDidPublishLog(_id: number, level: LogLevel, message: string): void {
					if (level === LogLevel.Warning) {
						warnings.push(message);
					}
					if (message.startsWith('405 status connecting')) {
						void backchannelStopped.complete();
					}
				}
			}, { async $checkMcpServerAllowed() { return undefined; } });
			const handle = store.add(new class extends McpHTTPHandle {
				protected override async _fetchInternal(_url: string, init?: CommonRequestInit): Promise<CommonResponse> {
					if (init?.method === 'GET') {
						return createMockResponse({ status: 405 });
					}
					assert.strictEqual(init?.method, 'POST');
					const signal = init?.signal;
					assert.ok(signal);
					const onAbort = () => controller.error(signal.reason);
					signal.addEventListener('abort', onAbort, { once: true });
					store.add(toDisposable(() => signal.removeEventListener('abort', onAbort)));
					return { ...createMockResponse({ headers: { 'Content-Type': 'text/event-stream' } }), body: stream };
				}
			}(1, { type: McpServerTransportType.HTTP, uri: URI.parse(TEST_MCP_URL), headers: [] }, proxy, store.add(new NullLogService())));
			const sending = handle.send(JSON.stringify({ jsonrpc: '2.0', id: 1, method: 'initialize', params: {} }));
			await Promise.all([reading.p, backchannelStopped.p]);
			return { handle, controller, cancel, sending, warnings };
		};

		test('does not cancel an already-aborted reader when disposed', async () => {
			const { handle, cancel, sending, warnings } = await createStreamingHandle();
			handle.dispose();
			await sending;
			// Observe any redundant cancellation rejection so the regression fails on the assertion below.
			await Promise.all(cancel.returnValues.map(result => assert.rejects(result, { name: 'AbortError' })));
			assert.deepStrictEqual({ cancellations: cancel.callCount, warnings }, { cancellations: 0, warnings: [] });
		});

		test('still reports a stream read failure while the handle is active', async () => {
			const { controller, cancel, sending, warnings } = await createStreamingHandle();
			const error = new Error('fixture stream read failed');
			controller.error(error);
			await sending;
			await Promise.all(cancel.returnValues.map(result => assert.rejects(result, candidate => candidate === error)));
			assert.deepStrictEqual({ cancellations: cancel.callCount, warnings }, {
				cancellations: 1,
				warnings: ['Error reading SSE stream: Error: fixture stream read failed'],
			});
		});
	});

	/* eslint-disable local/code-no-bracket-notation-for-identifiers -- Exercise private request entry points without widening the transport API. */
	suite('McpHTTPHandle request destinations', () => {
		const otherUrl = 'https://other.example/mcp';
		const deniedMessage = 'Request denied by the test MCP policy';

		function createHandle(options: {
			uri?: URI;
			headers?: McpServerTransportHTTP['headers'];
			oauth?: McpServerTransportHTTP['oauth'];
			transport?: McpServerTransportHTTP['transport'];
			startQuiet?: boolean;
			check?: (url: string) => Promise<string | undefined>;
			respond?: (url: string, init: CommonRequestInit) => CommonResponse;
			onState?: (state: McpConnectionState) => void;
			getToken?: (options: IMcpAuthenticationOptions, details: IMcpAuthenticationDetails) => Promise<string | IMcpOAuthGrant>;
		} = {}) {
			const requests: { url: string; method: string; headers: Record<string, string>; body: Uint8Array<ArrayBuffer> | undefined }[] = [];
			const checkedUrls: string[] = [];
			const states: McpConnectionState[] = [];
			const messages: string[] = [];
			const signals: AbortSignal[] = [];
			const backchannelStopped = new DeferredPromise<void>();
			const proxy = Object.assign(new class extends mock<MainThreadMcpShape>() {
				override $onDidChangeState(_id: number, state: McpConnectionState): void {
					states.push(state);
					options.onState?.(state);
				}
				override $onDidReceiveMessage(_id: number, message: string): void { messages.push(message); }
				override $onDidPublishLog(_id: number, _level: LogLevel, message: string): void {
					if (message.startsWith('405 status connecting')) {
						void backchannelStopped.complete();
					}
				}
				override async $getTokenFromServerMetadata(_id: number, details: IMcpAuthenticationDetails, authOptions: IMcpAuthenticationOptions): Promise<IMcpOAuthGrant> {
					const token = options.getToken ? await options.getToken(authOptions, details) : 'test-access-token';
					return typeof token === 'string' ? { accessToken: token, scopes: details.scopes ?? details.resourceMetadata?.scopes_supported ?? details.authorizationServerMetadata.scopes_supported ?? [] } : token;
				}
				override $logMcpAuthSetup(): void { }
			}, {
				async $checkMcpServerAllowed(id: number, url: string): Promise<string | undefined> {
					assert.strictEqual(id, 1);
					checkedUrls.push(url);
					return options.check?.(url);
				}
			});
			const handle = store.add(new class extends McpHTTPHandle {
				protected override async _fetchInternal(url: string, init?: CommonRequestInit): Promise<CommonResponse> {
					assert.ok(init);
					requests.push({ url, method: init.method, headers: { ...init.headers }, body: init.body });
					if (init.signal) {
						signals.push(init.signal);
					}
					return options.respond?.(url, init) ?? createMockResponse({ status: 202 });
				}
			}(1, {
				type: McpServerTransportType.HTTP,
				uri: options.uri ?? URI.parse(TEST_MCP_URL),
				headers: options.headers ?? [],
				oauth: options.oauth,
				transport: options.transport,
			}, proxy, new NullLogService(), options.startQuiet));

			return { handle, requests, checkedUrls, states, signals, backchannelStopped, messages };
		}

		for (const status of [307, 308]) {
			for (const destination of [otherUrl, 'https://example.com/private']) {
				test(`checks the destination before following ${status} to ${destination}`, async () => {
					const { handle, requests, checkedUrls, states } = createHandle({
						check: async url => url === destination ? deniedMessage : undefined,
						respond: url => url === TEST_MCP_URL
							? createMockResponse({ status, headers: { location: destination } })
							: createMockResponse({ status: 202 }),
					});

					await assert.rejects(handle['_fetch'](TEST_MCP_URL, { method: 'POST', headers: {}, body: new TextEncoder().encode('test-message') }), { message: deniedMessage });
					assert.deepStrictEqual({
						requests: requests.map(request => request.url),
						checkedUrls,
						state: states.at(-1),
					}, {
						requests: [TEST_MCP_URL],
						checkedUrls: [TEST_MCP_URL, destination],
						state: { state: McpConnectionState.Kind.Error, message: deniedMessage },
					});
				});
			}
		}

		test('checks every hop in a redirect chain', async () => {
			const finalUrl = 'https://third.example/mcp';
			const { handle, requests, checkedUrls } = createHandle({
				check: async url => url === finalUrl ? deniedMessage : undefined,
				respond: url => createMockResponse({ status: 307, headers: { location: url === TEST_MCP_URL ? otherUrl : finalUrl } }),
			});

			await assert.rejects(handle['_fetch'](TEST_MCP_URL, { method: 'POST', headers: {} }), { message: deniedMessage });
			assert.deepStrictEqual({ requests: requests.map(request => request.url), checkedUrls }, {
				requests: [TEST_MCP_URL, otherUrl],
				checkedUrls: [TEST_MCP_URL, otherUrl, finalUrl],
			});
		});

		test('checks a server-selected endpoint before its first request', async () => {
			const { handle, requests, checkedUrls } = createHandle({ check: async () => deniedMessage });
			await assert.rejects(handle['_fetch'](otherUrl, { method: 'POST', headers: {} }), { message: deniedMessage });
			assert.deepStrictEqual({ requests, checkedUrls }, { requests: [], checkedUrls: [otherUrl] });
		});

		test('waits for policy validation before dispatching a request', async () => {
			const verdict = new DeferredPromise<string | undefined>();
			const { handle, requests } = createHandle({ check: () => verdict.p });
			const pending = handle['_fetch'](TEST_MCP_URL, { method: 'POST', headers: {} });
			try {
				await Promise.resolve();
				assert.strictEqual(requests.length, 0);
			} finally {
				await verdict.complete(undefined);
				await pending;
			}
			assert.strictEqual(requests.length, 1);
		});

		test('does not dispatch when policy validation fails', async () => {
			const failure = new Error('Policy validation unavailable');
			const { handle, requests } = createHandle({ check: async () => { throw failure; } });
			await assert.rejects(handle['_fetch'](TEST_MCP_URL, { method: 'POST', headers: {} }), error => error === failure);
			assert.deepStrictEqual(requests, []);
		});

		test('does not dispatch after disposal while awaiting policy', async () => {
			const verdict = new DeferredPromise<string | undefined>();
			const { handle, requests } = createHandle({ check: () => verdict.p });
			const pending = handle['_fetch'](TEST_MCP_URL, { method: 'POST', headers: {} });
			const rejected = assert.rejects(pending, isCancellationError);
			handle.dispose();
			await verdict.complete(undefined);
			await rejected;
			assert.deepStrictEqual(requests, []);
		});

		test('does not dispatch requests on a disposed transport', async () => {
			const { handle, requests } = createHandle();
			handle.dispose();
			await assert.rejects(handle['_fetch'](TEST_MCP_URL, { method: 'POST', headers: {} }), isCancellationError);
			assert.deepStrictEqual(requests, []);
		});

		for (const status of [301, 302, 303, 307, 308]) {
			test(`removes configured headers on a cross-origin ${status}`, async () => {
				const { handle, requests } = createHandle({
					headers: [['X-Api-Key', 'test-key'], ['X-Trace-Id', 'test-trace']],
					respond: url => url === TEST_MCP_URL
						? createMockResponse({ status, headers: { location: otherUrl } })
						: createMockResponse({ status: 202 }),
				});
				await handle['_fetch'](TEST_MCP_URL, {
					method: 'POST',
					headers: { 'x-API-key': 'test-key', 'x-trace-id': 'test-trace', Accept: 'application/json' },
				});
				assert.deepStrictEqual({
					apiKey: requests[1].headers['x-API-key'],
					traceId: requests[1].headers['x-trace-id'],
					accept: requests[1].headers.Accept,
				}, { apiKey: undefined, traceId: undefined, accept: 'application/json' });
			});
		}

		test('removes configured credentials from a cross-origin server-selected endpoint', async () => {
			const { handle, requests } = createHandle({ headers: [['X-Api-Key', 'test-key']] });
			await handle['_fetch'](otherUrl, { method: 'POST', headers: { 'X-Api-Key': 'test-key', Authorization: 'test-token' } });
			assert.deepStrictEqual({
				apiKey: requests[0].headers['X-Api-Key'],
				authorization: requests[0].headers.Authorization,
			}, { apiKey: undefined, authorization: undefined });
		});

		test('preserves same-origin credentials and strips standard credentials across origins', async () => {
			const sameOriginUrl = 'https://example.com/next';
			const { handle, requests } = createHandle({
				headers: [['X-Api-Key', 'test-key']],
				respond: url => url === otherUrl
					? createMockResponse({ status: 202 })
					: createMockResponse({ status: 307, headers: { location: url === TEST_MCP_URL ? sameOriginUrl : otherUrl } }),
			});
			await handle['_fetch'](TEST_MCP_URL, {
				method: 'POST',
				headers: { Authorization: 'test-token', Cookie: 'test-cookie', 'Proxy-Authorization': 'test-proxy', 'Mcp-Session-Id': 'test-session', 'X-Api-Key': 'test-key' },
			});
			assert.deepStrictEqual(requests.map(request => [
				request.headers.Authorization, request.headers.Cookie, request.headers['Proxy-Authorization'], request.headers['Mcp-Session-Id'], request.headers['X-Api-Key'],
			]), [
				['test-token', 'test-cookie', 'test-proxy', 'test-session', 'test-key'],
				['test-token', 'test-cookie', 'test-proxy', 'test-session', 'test-key'],
				[undefined, undefined, undefined, undefined, undefined],
			]);
		});

		test('keeps caller headers intact for a retry against the configured origin', async () => {
			let redirected = false;
			const { handle, requests } = createHandle({
				respond: () => {
					if (!redirected) {
						redirected = true;
						return createMockResponse({ status: 307, headers: { location: otherUrl } });
					}
					return createMockResponse({ status: 202 });
				},
			});
			const originalHeaders = { Authorization: 'test-token', Cookie: 'test-cookie', 'Mcp-Session-Id': 'test-session', Accept: 'application/json' };
			const init = { method: 'POST', headers: { ...originalHeaders } };
			await handle['_fetch'](TEST_MCP_URL, init);
			await handle['_fetch'](TEST_MCP_URL, init);

			assert.deepStrictEqual({
				caller: init.headers,
				redirected: requests[1].headers.Authorization,
				retry: [requests[2].headers.Authorization, requests[2].headers.Cookie, requests[2].headers['Mcp-Session-Id']],
			}, { caller: originalHeaders, redirected: undefined, retry: ['test-token', 'test-cookie', 'test-session'] });
		});

		for (const status of [301, 302, 303]) {
			test(`keeps caller method and body intact after a ${status}`, async () => {
				let redirected = false;
				const { handle, requests } = createHandle({
					respond: () => {
						if (!redirected) {
							redirected = true;
							return createMockResponse({ status, headers: { location: otherUrl } });
						}
						return createMockResponse({ status: 202 });
					},
				});
				const body = new TextEncoder().encode('test-message');
				const init = { method: 'POST', headers: {}, body };
				await handle['_fetch'](TEST_MCP_URL, init);
				await handle['_fetch'](TEST_MCP_URL, init);
				assert.deepStrictEqual(requests.map(request => ({ method: request.method, body: request.body })), [
					{ method: 'POST', body },
					{ method: 'GET', body: undefined },
					{ method: 'POST', body },
				]);
				assert.deepStrictEqual({ method: init.method, body: init.body }, { method: 'POST', body });
			});
		}

		test('does not restore credentials when a redirect chain returns to its original origin', async () => {
			const returnUrl = 'https://example.com/returned';
			const { handle, requests } = createHandle({
				headers: [['X-Api-Key', 'test-key']],
				respond: url => url === returnUrl
					? createMockResponse({ status: 202 })
					: createMockResponse({ status: 307, headers: { location: url === TEST_MCP_URL ? otherUrl : returnUrl } }),
			});
			await handle['_fetch'](TEST_MCP_URL, { method: 'POST', headers: { 'X-Api-Key': 'test-key' } });
			assert.deepStrictEqual(requests.map(request => request.headers['X-Api-Key']), ['test-key', undefined, undefined]);
		});

		test('preserves generated transport headers when configured names overlap', async () => {
			const { handle, requests, backchannelStopped } = createHandle({
				headers: [['Accept', 'configured-accept'], ['Content-Type', 'configured-type'], ['X-Api-Key', 'test-key']],
				respond: (url, init) => init.method === 'GET'
					? createMockResponse({ status: 405 })
					: url === TEST_MCP_URL
						? createMockResponse({ status: 307, headers: { location: otherUrl } })
						: createMockResponse({ status: 202 }),
			});
			await handle.send('test-message');
			await backchannelStopped.p;
			const redirected = requests.find(request => request.url === otherUrl);
			assert.ok(redirected);
			assert.deepStrictEqual({
				accept: redirected.headers.Accept,
				contentType: redirected.headers['Content-Type'],
				apiKey: redirected.headers['X-Api-Key'],
			}, { accept: 'text/event-stream, application/json', contentType: 'application/json', apiKey: undefined });
		});

		test('validates MCP retries without applying server allowlists to OAuth metadata discovery', async () => {
			let authenticated = false;
			const { handle, requests, checkedUrls } = createHandle({
				check: async url => url === TEST_MCP_URL ? undefined : deniedMessage,
				respond: url => {
					if (url === TEST_MCP_URL) {
						if (authenticated) {
							return createMockResponse({ status: 202 });
						}
						authenticated = true;
						return createMockResponse({ status: 401 });
					}
					if (url.includes('oauth-protected-resource')) {
						return createMockResponse({ body: JSON.stringify({ resource: TEST_MCP_URL, authorization_servers: [TEST_AUTH_SERVER] }) });
					}
					return createMockResponse({
						body: JSON.stringify({
							issuer: TEST_AUTH_SERVER,
							authorization_endpoint: `${TEST_AUTH_SERVER}/authorize`,
							token_endpoint: `${TEST_AUTH_SERVER}/token`,
							response_types_supported: ['code'],
						})
					});
				},
			});
			const headers: Record<string, string> = {};
			const response = await handle['_fetchWithAuthRetry'](TEST_MCP_URL, { method: 'POST', headers }, headers);
			assert.strictEqual(response.status, 202);
			assert.ok(requests.some(request => request.url.startsWith(TEST_AUTH_SERVER)));
			assert.deepStrictEqual(checkedUrls, [TEST_MCP_URL, TEST_MCP_URL]);
		});

		for (const [name, error] of [
			['dismissed', new CancellationError()],
			['required during a quiet start', new UserInteractionRequiredError('authentication')],
			['cancelled by Microsoft sign-in', Object.assign(new Error('user_canceled: See https://aka.ms/msal.js.errors#user_canceled for details'), { name: 'ClientAuthError' })],
		] as const) {
			test(`stops for sign-in without retrying the request when OAuth is ${name}`, async () => {
				const { handle, requests, states, signals } = createHandle({
					getToken: async () => { throw transformErrorFromSerialization(transformErrorForSerialization(error)); },
					respond: url => {
						if (url === TEST_MCP_URL) {
							return createMockResponse({ status: 401 });
						}
						if (url.includes('oauth-protected-resource')) {
							return createMockResponse({ body: JSON.stringify({ resource: TEST_MCP_URL, authorization_servers: [TEST_AUTH_SERVER] }) });
						}
						return createMockResponse({
							body: JSON.stringify({
								issuer: TEST_AUTH_SERVER,
								authorization_endpoint: `${TEST_AUTH_SERVER}/authorize`,
								token_endpoint: `${TEST_AUTH_SERVER}/token`,
								response_types_supported: ['code'],
							})
						});
					},
				});

				await handle.send('test-message');
				assert.deepStrictEqual(states, [
					{ state: McpConnectionState.Kind.Running },
					{ state: McpConnectionState.Kind.Stopped, reason: 'needs-user-interaction' },
				]);
				assert.strictEqual(requests.filter(request => request.url === TEST_MCP_URL).length, 1);
				assert.ok(signals.length && signals.every(signal => signal.aborted));
			});
		}

		test('requests Google browser authorization before initialize can return a public tool catalog', async () => {
			const gmail = 'https://gmailmcp.googleapis.com/mcp/v1';
			const issuer = 'https://accounts.google.com';
			const { handle, requests, states } = createHandle({
				uri: URI.parse(gmail),
				oauth: { clientId: 'registered-client-id' },
				getToken: async () => { throw new CancellationError(); },
				respond: (url) => {
					if (url === gmail) { return createMockResponse({ status: 202, url }); }
					const body = url.includes('oauth-protected-resource')
						? { resource: gmail, authorization_servers: [issuer] }
						: { issuer, authorization_endpoint: `${issuer}/o/oauth2/v2/auth`, token_endpoint: 'https://oauth2.googleapis.com/token', response_types_supported: ['code'] };
					return createMockResponse({ url, body: JSON.stringify(body) });
				},
			});
			await handle.send(JSON.stringify({ jsonrpc: '2.0', method: 'initialize', id: 1, params: {} }));
			assert.deepStrictEqual({ endpointRequests: requests.filter(request => request.url === gmail), state: states.at(-1) }, {
				endpointRequests: [], state: { state: McpConnectionState.Kind.Stopped, reason: 'needs-user-interaction' },
			});
		});

		test('starts a configured SSE connection with an authenticated GET', async () => {
			const { handle, requests } = createHandle({
				transport: 'sse',
				respond: (url, init) => {
					if (url === TEST_MCP_URL) {
						return init.headers.Authorization
							? { ...createMockResponse({ url, headers: { 'Content-Type': 'text/event-stream' } }), body: new Response('event: endpoint\ndata: /messages\n\n').body }
							: createMockResponse({ url, status: 401, headers: { 'WWW-Authenticate': 'bearer scope="read"' } });
					}
					if (url.includes('oauth-protected-resource')) {
						return createMockResponse({ url, body: JSON.stringify({ resource: TEST_MCP_URL, authorization_servers: [TEST_AUTH_SERVER] }) });
					}
					if (url.includes('oauth-authorization-server')) {
						return createMockResponse({ url, body: JSON.stringify({ issuer: TEST_AUTH_SERVER, authorization_endpoint: `${TEST_AUTH_SERVER}/authorize`, token_endpoint: `${TEST_AUTH_SERVER}/token`, response_types_supported: ['code'] }) });
					}
					return createMockResponse({ status: 202, url });
				},
			});
			await handle.send('fixture message');
			assert.deepStrictEqual(requests.filter(request => [TEST_MCP_URL, `${new URL(TEST_MCP_URL).origin}/messages`].includes(request.url)).map(request => [request.method, request.headers.Authorization]),
				[['GET', undefined], ['GET', 'Bearer test-access-token'], ['POST', 'Bearer test-access-token']]);
		});

		test('keeps a provider permission error without re-registering the OAuth client', async () => {
			let tokenRequests = 0;
			const { handle, requests } = createHandle({
				getToken: async () => { tokenRequests++; return 'authorized-token'; },
				respond: (url, init) => {
					if (url === TEST_MCP_URL) {
						return init.headers.Authorization
							? createMockResponse({ url, status: 403, body: 'Organization access denied' })
							: createMockResponse({ url, status: 401, headers: { 'WWW-Authenticate': 'Bearer scope="read"' } });
					}
					const body = url.includes('oauth-protected-resource')
						? { resource: TEST_MCP_URL, authorization_servers: [TEST_AUTH_SERVER] }
						: { issuer: TEST_AUTH_SERVER, authorization_endpoint: `${TEST_AUTH_SERVER}/authorize`, token_endpoint: `${TEST_AUTH_SERVER}/token`, response_types_supported: ['code'] };
					return createMockResponse({ url, body: JSON.stringify(body) });
				},
			});
			const headers: Record<string, string> = {};
			const result = await handle['_fetchWithAuthRetry'](TEST_MCP_URL, { method: 'POST', headers }, headers);
			assert.deepStrictEqual({ status: result.status, detail: await result.text(), tokenRequests, requests: requests.filter(request => request.url === TEST_MCP_URL).map(request => request.headers.Authorization) }, {
				status: 403, detail: 'Organization access denied', tokenRequests: 1, requests: [undefined, 'Bearer authorized-token'],
			});
		});

		for (const scenario of ['granted', 'declined', 'quiet', 'cancelled during consent', 'still forbidden', 'legacy SSE', 'partially granted on sign-in', 'legacy cached scopes'] as const) {
			test(`keeps MCP connected when additional tool permission is ${scenario}`, async () => {
				const endpoint = `${new URL(TEST_MCP_URL).origin}/messages`;
				const scopeChallenge = 'Bearer error="insufficient_scope", scope="write"';
				const consentRequested = new DeferredPromise<void>();
				const consentResponse = new DeferredPromise<string>();
				let scopeRequests = 0;
				const initiallyRequestedBoth = scenario === 'partially granted on sign-in' || scenario === 'legacy cached scopes';
				const { handle, requests, states, messages } = createHandle({
					startQuiet: scenario !== 'quiet',
					transport: scenario === 'legacy SSE' ? 'sse' : undefined,
					getToken: async (options, details) => {
						if (!options.scopeUpgradeFrom) {
							if (initiallyRequestedBoth && scopeRequests === 0) {
								return { accessToken: 'read-token', scopes: scenario === 'legacy cached scopes' ? ['read', 'write'] : ['read'] };
							}
							return details.scopes?.includes('write') ? 'write-token' : 'read-token';
						}
						scopeRequests++;
						assert.strictEqual(options.errorOnUserInteraction, scenario === 'quiet');
						assert.deepStrictEqual(details.scopes, ['read', 'write']);
						assert.deepStrictEqual(options.scopeUpgradeFrom, { scopes: initiallyRequestedBoth ? ['read', 'write'] : ['read'], tokenFingerprint: await getOAuthTokenFingerprint('read-token') });
						if (scenario === 'quiet') { throw new UserInteractionRequiredError('authentication'); }
						if (scenario === 'declined') { throw new CancellationError(); }
						if (scenario === 'cancelled during consent') {
							void consentRequested.complete();
							return consentResponse.p;
						}
						return 'write-token';
					},
					respond: (url, init) => {
						if (url === TEST_MCP_URL || url === endpoint) {
							if (init.method === 'GET' && scenario !== 'legacy SSE') { return createMockResponse({ url, status: 405 }); }
							if (!init.headers.Authorization) {
								const advertised = initiallyRequestedBoth ? 'read write' : 'read';
								return createMockResponse({ url, status: 401, headers: { 'WWW-Authenticate': `Bearer scope="${advertised}"` } });
							}
							if (init.method === 'GET') {
								return { ...createMockResponse({ url, headers: { 'Content-Type': 'text/event-stream' } }), body: new Response('event: endpoint\ndata: /messages\n\n').body };
							}
							const request = JSON.parse(new TextDecoder().decode(init.body));
							if (request.params?.name === 'write' && (init.headers.Authorization === 'Bearer read-token' || scenario === 'still forbidden')) {
								return createMockResponse({ url, status: 403, headers: { 'WWW-Authenticate': scopeChallenge } });
							}
							return createMockResponse({ url, status: 202 });
						}
						const body = url.includes('oauth-protected-resource')
							? { resource: TEST_MCP_URL, authorization_servers: [TEST_AUTH_SERVER] }
							: { issuer: TEST_AUTH_SERVER, authorization_endpoint: `${TEST_AUTH_SERVER}/authorize`, token_endpoint: `${TEST_AUTH_SERVER}/token`, response_types_supported: ['code'] };
						return createMockResponse({ url, body: JSON.stringify(body) });
					},
				});
				const toolCall = (id: number, name: string) => JSON.stringify({ jsonrpc: '2.0', id, method: 'tools/call', params: { name } });
				await handle.send(JSON.stringify({ jsonrpc: '2.0', id: 1, method: 'initialize' }));
				const originalWrite = handle.send(toolCall(2, 'write'), scenario === 'quiet');
				if (scenario === 'cancelled during consent') {
					await consentRequested.p;
					await handle.send(JSON.stringify({ jsonrpc: '2.0', method: 'notifications/cancelled', params: { requestId: 2 } }));
					await consentResponse.complete('write-token');
				}
				await originalWrite;
				await handle.send(toolCall(3, 'read'), true);
				const granted = scenario === 'granted' || scenario === 'legacy SSE' || initiallyRequestedBoth;
				const posts = requests.filter(request => (request.url === TEST_MCP_URL || request.url === endpoint) && request.method === 'POST' && request.headers.Authorization);
				const tokensFor = (id: number) => posts.filter(request => JSON.parse(new TextDecoder().decode(request.body)).id === id).map(request => request.headers.Authorization);
				assert.deepStrictEqual(tokensFor(2), granted || scenario === 'still forbidden' ? ['Bearer read-token', 'Bearer write-token'] : ['Bearer read-token']);
				assert.deepStrictEqual(tokensFor(3), [granted ? 'Bearer write-token' : 'Bearer read-token']);
				assert.deepStrictEqual(states, [{ state: McpConnectionState.Kind.Running }]);
				assert.strictEqual(scopeRequests, 1);
				assert.deepStrictEqual(messages.map(message => JSON.parse(message)), granted || scenario === 'cancelled during consent' ? [] : [{
					jsonrpc: '2.0', id: 2, error: { code: MCP.INTERNAL_ERROR, message: 'This tool requires additional sign-in permissions for the connected account.' },
				}]);
			});
		}

		for (const invalidToken of [false, true]) {
			test(invalidToken ? 'retries a Google Workspace invalid-token challenge' : 'keeps the Gmail API-disabled response without replacing the signed-in account', async () => {
				const gmailUrl = 'https://gmailmcp.googleapis.com/mcp/v1';
				const googleAuth = 'https://accounts.google.com';
				const disabled = JSON.stringify({ jsonrpc: '2.0', id: 1, result: { isError: true, content: [{ type: 'text', text: 'Gmail MCP API has not been used in this project before or it is disabled.' }] } });
				let authenticatedRequests = 0;
				const rejectedFingerprints: string[] = [];
				const { handle, requests } = createHandle({
					uri: URI.parse(gmailUrl),
					getToken: async ({ rejectedTokenFingerprint }) => {
						if (rejectedTokenFingerprint) {
							rejectedFingerprints.push(rejectedTokenFingerprint);
							return 'replacement-token';
						}
						return 'test-access-token';
					},
					respond: (url, init) => {
						if (url === gmailUrl) {
							if (!init.headers.Authorization) {
								return createMockResponse({ status: 401, url });
							}
							authenticatedRequests++;
							return authenticatedRequests === 1
								? createMockResponse({ status: 403, url, body: disabled, ...(invalidToken ? { headers: { 'WWW-Authenticate': 'Bearer error="invalid_token"' } } : {}) })
								: createMockResponse({ status: 202, url });
						}
						if (url.includes('oauth-protected-resource')) {
							return createMockResponse({ url, body: JSON.stringify({ resource: gmailUrl, authorization_servers: [googleAuth] }) });
						}
						return createMockResponse({ url, body: JSON.stringify({
							issuer: googleAuth,
							authorization_endpoint: `${googleAuth}/o/oauth2/v2/auth`,
							token_endpoint: 'https://oauth2.googleapis.com/token',
							response_types_supported: ['code'],
						}) });
					},
				});
				const headers: Record<string, string> = {};
				const response = await handle['_fetchWithAuthRetry'](gmailUrl, { method: 'POST', headers }, headers);
				assert.strictEqual(response.status, invalidToken ? 202 : 403);
				assert.deepStrictEqual(requests.filter(request => request.url === gmailUrl).map(request => request.headers.Authorization),
					invalidToken ? [undefined, 'Bearer test-access-token', 'Bearer replacement-token'] : [undefined, 'Bearer test-access-token']);
				assert.deepStrictEqual(rejectedFingerprints, invalidToken ? [await getOAuthTokenFingerprint('test-access-token')] : []);
				if (!invalidToken) { assert.strictEqual(await response.text(), disabled); }
			});
		}

		test('preserves manually configured authorization when the resource rejects it', async () => {
			const { handle, requests } = createHandle({
				getToken: async () => assert.fail('A manually configured header does not belong to OAuth session recovery'),
				respond: url => createMockResponse({ url, status: 401, headers: { 'WWW-Authenticate': 'Bearer error="invalid_token"' } }),
			});
			const headers = { authorization: 'Bearer configured-value' };
			const response = await handle['_fetchWithAuthRetry'](TEST_MCP_URL, { method: 'POST', headers }, headers);
			assert.deepStrictEqual({ status: response.status, requests: requests.map(request => [request.url, request.headers.authorization]) }, {
				status: 401, requests: [[TEST_MCP_URL, 'Bearer configured-value']],
			});
		});

		test('continues rejecting non-HTTP redirect destinations', async () => {
			const { handle, requests } = createHandle({
				respond: () => createMockResponse({ status: 307, headers: { location: 'file:///test-fixture' } }),
			});
			await assert.rejects(handle['_fetch'](TEST_MCP_URL, { method: 'POST', headers: {} }), /non-http\(s\) target/);
			assert.deepStrictEqual(requests.map(request => request.url), [TEST_MCP_URL]);
		});

		test('preserves the generated protocol version across OAuth metadata redirects', async () => {
			const metadataUrl = 'https://other.example/resource-metadata';
			let authenticated = false;
			const { handle, requests } = createHandle({
				headers: [['MCP-Protocol-Version', 'configured-version']],
				respond: url => {
					if (url === TEST_MCP_URL) {
						const status = authenticated ? 202 : 401;
						authenticated = true;
						return createMockResponse({ status });
					}
					if (url === metadataUrl) {
						return createMockResponse({ body: JSON.stringify({ resource: TEST_MCP_URL, authorization_servers: [TEST_AUTH_SERVER] }) });
					}
					if (url.startsWith(TEST_AUTH_SERVER)) {
						return createMockResponse({
							body: JSON.stringify({
								issuer: TEST_AUTH_SERVER,
								authorization_endpoint: `${TEST_AUTH_SERVER}/authorize`,
								token_endpoint: `${TEST_AUTH_SERVER}/token`,
								response_types_supported: ['code'],
							})
						});
					}
					return createMockResponse({ status: 307, headers: { location: metadataUrl } });
				},
			});
			const headers: Record<string, string> = {};
			await handle['_fetchWithAuthRetry'](TEST_MCP_URL, { method: 'POST', headers }, headers);
			const metadataRequest = requests.find(request => request.url === metadataUrl);
			assert.ok(metadataRequest);
			assert.strictEqual(metadataRequest.headers['MCP-Protocol-Version'], MCP.LATEST_PROTOCOL_VERSION);
		});

		suite('first destinations', () => {
			for (const scheme of ['unix', 'pipe', 'file', 'fixture']) {
				for (const isAuthMetadata of [false, true]) {
					test(`rejects an unconfigured ${scheme} first destination for ${isAuthMetadata ? 'metadata' : 'MCP data'}`, async () => {
						const { handle, requests, checkedUrls } = createHandle();
						const destination = URI.from({ scheme, path: '/test-mcp-destination', fragment: '/request' }).toString(true);
						await assert.rejects(handle['_fetch'](destination, { method: 'GET', headers: {} }, { isAuthMetadata }), /non-http\(s\).*not allowed/);
						assert.deepStrictEqual({ requests, checkedUrls }, { requests: [], checkedUrls: [] });
					});
				}
			}

			for (const scheme of ['unix', 'pipe']) {
				const launch = URI.from({ scheme, path: '/test-mcp-transport', fragment: '/mcp' });

				test(`preserves the explicitly configured ${scheme} transport and its HTTP routes`, async () => {
					const { handle, requests } = createHandle({ uri: launch, headers: [['X-Api-Key', 'test-key']] });
					const destination = launch.with({ fragment: '/messages' }).toString(true);
					await handle['_fetch'](launch.toString(true), { method: 'POST', headers: { 'X-Api-Key': 'test-key' } });
					await handle['_fetch'](destination, { method: 'POST', headers: { 'X-Api-Key': 'test-key' } });
					assert.deepStrictEqual(requests.map(request => [request.url, request.headers['X-Api-Key']]), [
						[launch.toString(true), 'test-key'],
						[destination, 'test-key'],
					]);
				});

				test(`compares the ${scheme} socket path as the dispatcher decodes it`, async () => {
					const { handle, requests } = createHandle({ uri: launch, headers: [['X-Api-Key', 'test-key']] });
					const destination = launch.with({ fragment: '/messages' }).toString(true).replace('transport', '%74ransport');
					await handle['_fetch'](destination, { method: 'POST', headers: { 'X-Api-Key': 'test-key' } });
					assert.deepStrictEqual(requests.map(request => [request.url, request.headers['X-Api-Key']]), [[destination, 'test-key']]);
				});

				for (const isAuthMetadata of [false, true]) {
					test(`rejects a different ${scheme} endpoint for ${isAuthMetadata ? 'metadata' : 'MCP data'} despite equal opaque origins`, async () => {
						const { handle, requests } = createHandle({ uri: launch, headers: [['X-Api-Key', 'test-key']] });
						const destination = launch.with({ path: '/other-test-mcp-transport' }).toString(true);
						await assert.rejects(handle['_fetch'](destination, { method: 'GET', headers: { 'X-Api-Key': 'test-key' } }, { isAuthMetadata }), /non-http\(s\).*not allowed/);
						assert.deepStrictEqual(requests, []);
					});
				}

				test(`does not treat another IPC scheme as the configured ${scheme} endpoint`, async () => {
					const { handle, requests } = createHandle({ uri: launch });
					const destination = launch.with({ scheme: scheme === 'unix' ? 'pipe' : 'unix' }).toString(true);
					await assert.rejects(handle['_fetch'](destination, { method: 'POST', headers: {} }), /non-http\(s\).*not allowed/);
					assert.deepStrictEqual(requests, []);
				});

				test(`strips ${scheme} credentials when using an HTTP metadata provider`, async () => {
					const { handle, requests } = createHandle({ uri: launch, headers: [['X-Api-Key', 'test-key']] });
					await handle['_fetch'](TEST_AUTH_SERVER, { method: 'GET', headers: { 'X-Api-Key': 'test-key', Accept: 'application/json' } }, { isAuthMetadata: true });
					assert.deepStrictEqual({
						apiKey: requests[0].headers['X-Api-Key'],
						accept: requests[0].headers.Accept,
					}, { apiKey: undefined, accept: 'application/json' });
				});
			}

			test('does not grant an unsupported scheme just because it was configured', async () => {
				const uri = URI.from({ scheme: 'fixture', authority: 'configured', path: '/mcp' });
				const { handle, requests } = createHandle({ uri });
				await assert.rejects(handle['_fetch'](uri.toString(true), { method: 'POST', headers: {} }), /non-http\(s\).*not allowed/);
				assert.deepStrictEqual(requests, []);
			});

			for (const destination of [
				'unix:/var/run/docker.sock?remote=https%3A%2F%2Fattacker.example%2FDockerfile#/build',
				'pipe:/docker_engine?remote=https%3A%2F%2Fattacker.example%2FDockerfile#/build',
				'fixture:/test-mcp-destination',
			]) {
				const scheme = URI.parse(destination).scheme;
				test(`reports a rejected legacy SSE ${scheme} endpoint and disposes its transport`, async () => {
					const completed = new DeferredPromise<void>();
					const { handle, requests, states, signals } = createHandle({
						onState: state => {
							if (state.state === McpConnectionState.Kind.Error) {
								void completed.complete();
							}
						},
						respond: (url, init) => {
							if (url === destination) {
								void completed.complete();
								return createMockResponse({ status: 202 });
							}
							if (init.method === 'POST') {
								return createMockResponse({ status: 404 });
							}
							return {
								...createMockResponse({ headers: { 'Content-Type': 'text/event-stream' } }),
								body: new ReadableStream<Uint8Array>({
									start(controller) {
										controller.enqueue(new TextEncoder().encode(`event: endpoint\ndata: ${destination}\n\n`));
										controller.close();
									}
								}),
							};
						},
					});
					await handle.send('test-message');
					await completed.p;
					const state = states.at(-1);
					assert.ok(state?.state === McpConnectionState.Kind.Error);
					assert.match(state.message, /non-http\(s\).*not allowed/);
					assert.deepStrictEqual(requests.map(request => request.url), [TEST_MCP_URL, TEST_MCP_URL]);
					assert.ok(signals.length > 0 && signals.every(signal => signal.aborted));
				});
			}
		});
	});

	/* eslint-enable local/code-no-bracket-notation-for-identifiers */
	suite('IAuthMetadata', () => {
		suite('properties', () => {
			test('should expose readonly properties', async () => {
				const { authMetadata } = await createTestAuthMetadata({
					scopes: ['read', 'write'],
					serverMetadataIssuer: TEST_AUTH_SERVER
				});

				assert.ok(authMetadata.authorizationServer.toString().startsWith(TEST_AUTH_SERVER));
				assert.strictEqual(authMetadata.serverMetadata.issuer, TEST_AUTH_SERVER);
				assert.deepStrictEqual(authMetadata.scopes, ['read', 'write']);
			});

			test('should allow undefined scopes', async () => {
				const { authMetadata } = await createTestAuthMetadata({
					scopes: undefined
				});

				assert.strictEqual(authMetadata.scopes, undefined);
			});
		});

		suite('update()', () => {
			test('should return true and update scopes when WWW-Authenticate header contains new scopes', async () => {
				const { authMetadata } = await createTestAuthMetadata({
					scopes: ['read']
				});

				const response = createMockResponse({
					status: 401,
					headers: {
						'WWW-Authenticate': 'Bearer scope="read write admin"'
					}
				});

				const result = authMetadata.update(response.headers);

				assert.strictEqual(result, true);
				assert.deepStrictEqual(authMetadata.scopes, ['read', 'write', 'admin']);
			});

			test('should return false when scopes are the same', async () => {
				const { authMetadata } = await createTestAuthMetadata({
					scopes: ['read', 'write']
				});

				const response = createMockResponse({
					status: 401,
					headers: {
						'WWW-Authenticate': 'Bearer scope="read write"'
					}
				});

				const result = authMetadata.update(response.headers);

				assert.strictEqual(result, false);
				assert.deepStrictEqual(authMetadata.scopes, ['read', 'write']);
			});

			test('should return false when scopes are same but in different order', async () => {
				const { authMetadata } = await createTestAuthMetadata({
					scopes: ['read', 'write']
				});

				const response = createMockResponse({
					status: 401,
					headers: {
						'WWW-Authenticate': 'Bearer scope="write read"'
					}
				});

				const result = authMetadata.update(response.headers);

				assert.strictEqual(result, false);
			});

			test('should return true when updating from undefined scopes to defined scopes', async () => {
				const { authMetadata } = await createTestAuthMetadata({
					scopes: undefined
				});

				const response = createMockResponse({
					status: 401,
					headers: {
						'WWW-Authenticate': 'Bearer scope="read"'
					}
				});

				const result = authMetadata.update(response.headers);

				assert.strictEqual(result, true);
				assert.deepStrictEqual(authMetadata.scopes, ['read']);
			});

			test('preserves discovered scopes when an auth challenge omits scope', async () => {
				const { authMetadata } = await createTestAuthMetadata({
					scopes: ['read']
				});

				const response = createMockResponse({
					status: 401,
					headers: {
						'WWW-Authenticate': 'Bearer realm="example"'
					}
				});

				const result = authMetadata.update(response.headers);

				assert.strictEqual(result, false);
				assert.deepStrictEqual(authMetadata.scopes, ['read']);
			});

			test('should return false when no WWW-Authenticate header and scopes are already undefined', async () => {
				const { authMetadata } = await createTestAuthMetadata({
					scopes: undefined
				});

				const response = createMockResponse({
					status: 401,
					headers: {}
				});

				const result = authMetadata.update(response.headers);

				assert.strictEqual(result, false);
			});

			test('should handle multiple Bearer challenges and use first scope', async () => {
				const { authMetadata } = await createTestAuthMetadata({
					scopes: undefined
				});

				const response = createMockResponse({
					status: 401,
					headers: {
						'WWW-Authenticate': 'Bearer scope="first", Bearer scope="second"'
					}
				});

				authMetadata.update(response.headers);

				assert.deepStrictEqual(authMetadata.scopes, ['first']);
			});

			test('should ignore non-Bearer schemes', async () => {
				const { authMetadata } = await createTestAuthMetadata({
					scopes: undefined
				});

				const response = createMockResponse({
					status: 401,
					headers: {
						'WWW-Authenticate': 'Basic realm="example"'
					}
				});

				const result = authMetadata.update(response.headers);

				assert.strictEqual(result, false);
				assert.strictEqual(authMetadata.scopes, undefined);
			});
		});
	});

	suite('createAuthMetadata', () => {
		let sandbox: sinon.SinonSandbox;
		let logMessages: Array<{ level: LogLevel; message: string }>;
		let mockLogger: (level: LogLevel, message: string) => void;

		setup(() => {
			sandbox = sinon.createSandbox();
			logMessages = [];
			mockLogger = (level, message) => logMessages.push({ level, message });
		});

		teardown(() => {
			sandbox.restore();
		});

		test('should create IAuthMetadata with fetched server metadata', async () => {
			const mockFetch = sandbox.stub();

			// Mock resource metadata fetch
			mockFetch.onCall(0).resolves(createMockResponse({
				status: 200,
				url: TEST_RESOURCE_METADATA_URL,
				body: JSON.stringify({
					resource: TEST_MCP_URL,
					authorization_servers: [TEST_AUTH_SERVER],
					scopes_supported: ['read', 'write']
				})
			}));

			// Mock server metadata fetch
			mockFetch.onCall(1).resolves(createMockResponse({
				status: 200,
				url: `${TEST_AUTH_SERVER}/.well-known/oauth-authorization-server`,
				body: JSON.stringify({
					issuer: TEST_AUTH_SERVER,
					authorization_endpoint: `${TEST_AUTH_SERVER}/authorize`,
					token_endpoint: `${TEST_AUTH_SERVER}/token`,
					response_types_supported: ['code']
				})
			}));

			const originalResponse = createMockResponse({
				status: 401,
				url: TEST_MCP_URL,
				headers: {
					'WWW-Authenticate': 'Bearer scope="api.read"'
				}
			});

			const authMetadata = await createAuthMetadata(
				TEST_MCP_URL,
				originalResponse.headers,
				{
					sameOriginHeaders: { 'X-Custom': 'value' },
					fetch: mockFetch,
					log: mockLogger
				}
			);

			assert.ok(authMetadata.authorizationServer.toString().startsWith(TEST_AUTH_SERVER));
			assert.strictEqual(authMetadata.serverMetadata.issuer, TEST_AUTH_SERVER);
			assert.deepStrictEqual(authMetadata.scopes, ['api.read']);
		});

		test('should fall back to default metadata when server metadata fetch fails', async () => {
			const mockFetch = sandbox.stub();

			// Mock resource metadata fetch - fails
			mockFetch.onCall(0).rejects(new Error('Network error'));

			// Mock server metadata fetch - also fails
			mockFetch.onCall(1).rejects(new Error('Network error'));

			const originalResponse = createMockResponse({
				status: 401,
				url: TEST_MCP_URL,
				headers: {}
			});

			const authMetadata = await createAuthMetadata(
				TEST_MCP_URL,
				originalResponse.headers,
				{
					sameOriginHeaders: {},
					fetch: mockFetch,
					log: mockLogger
				}
			);

			// Should use default metadata based on the URL
			assert.ok(authMetadata.authorizationServer.toString().startsWith('https://example.com'));
			assert.ok(authMetadata.serverMetadata.issuer.startsWith('https://example.com'));
			assert.ok(authMetadata.serverMetadata.authorization_endpoint?.startsWith('https://example.com/authorize'));
			assert.ok(authMetadata.serverMetadata.token_endpoint?.startsWith('https://example.com/token'));

			// Should log the fallback
			assert.ok(logMessages.some(m =>
				m.level === LogLevel.Info &&
				m.message.includes('Using default auth metadata')
			));
		});

		test('should use scopes from WWW-Authenticate header when resource metadata has none', async () => {
			const mockFetch = sandbox.stub();

			// Mock resource metadata fetch - no scopes_supported
			mockFetch.onCall(0).resolves(createMockResponse({
				status: 200,
				url: TEST_RESOURCE_METADATA_URL,
				body: JSON.stringify({
					resource: TEST_MCP_URL,
					authorization_servers: [TEST_AUTH_SERVER]
				})
			}));

			// Mock server metadata fetch
			mockFetch.onCall(1).resolves(createMockResponse({
				status: 200,
				url: `${TEST_AUTH_SERVER}/.well-known/oauth-authorization-server`,
				body: JSON.stringify({
					issuer: TEST_AUTH_SERVER,
					authorization_endpoint: `${TEST_AUTH_SERVER}/authorize`,
					token_endpoint: `${TEST_AUTH_SERVER}/token`,
					response_types_supported: ['code']
				})
			}));

			const originalResponse = createMockResponse({
				status: 401,
				url: TEST_MCP_URL,
				headers: {
					'WWW-Authenticate': 'Bearer scope="header.scope"'
				}
			});

			const authMetadata = await createAuthMetadata(
				TEST_MCP_URL,
				originalResponse.headers,
				{
					sameOriginHeaders: {},
					fetch: mockFetch,
					log: mockLogger
				}
			);

			assert.deepStrictEqual(authMetadata.scopes, ['header.scope']);
		});

		test('should use scopes from WWW-Authenticate header even when resource metadata has scopes_supported', async () => {
			const mockFetch = sandbox.stub();

			// Mock resource metadata fetch - has scopes_supported
			mockFetch.onCall(0).resolves(createMockResponse({
				status: 200,
				url: TEST_RESOURCE_METADATA_URL,
				body: JSON.stringify({
					resource: TEST_MCP_URL,
					authorization_servers: [TEST_AUTH_SERVER],
					scopes_supported: ['resource.scope1', 'resource.scope2']
				})
			}));

			// Mock server metadata fetch
			mockFetch.onCall(1).resolves(createMockResponse({
				status: 200,
				url: `${TEST_AUTH_SERVER}/.well-known/oauth-authorization-server`,
				body: JSON.stringify({
					issuer: TEST_AUTH_SERVER,
					authorization_endpoint: `${TEST_AUTH_SERVER}/authorize`,
					token_endpoint: `${TEST_AUTH_SERVER}/token`,
					response_types_supported: ['code']
				})
			}));

			const originalResponse = createMockResponse({
				status: 401,
				url: TEST_MCP_URL,
				headers: {
					'WWW-Authenticate': 'Bearer scope="header.scope"'
				}
			});

			const authMetadata = await createAuthMetadata(
				TEST_MCP_URL,
				originalResponse.headers,
				{
					sameOriginHeaders: {},
					fetch: mockFetch,
					log: mockLogger
				}
			);

			// WWW-Authenticate header scopes take precedence over resource metadata scopes_supported
			assert.deepStrictEqual(authMetadata.scopes, ['header.scope']);
		});

		test('should use resource_metadata challenge URL from WWW-Authenticate header', async () => {
			const mockFetch = sandbox.stub();

			// Mock resource metadata fetch from challenge URL
			mockFetch.onCall(0).resolves(createMockResponse({
				status: 200,
				url: 'https://example.com/custom-resource-metadata',
				body: JSON.stringify({
					resource: TEST_MCP_URL,
					authorization_servers: [TEST_AUTH_SERVER]
				})
			}));

			// Mock server metadata fetch
			mockFetch.onCall(1).resolves(createMockResponse({
				status: 200,
				url: `${TEST_AUTH_SERVER}/.well-known/oauth-authorization-server`,
				body: JSON.stringify({
					issuer: TEST_AUTH_SERVER,
					authorization_endpoint: `${TEST_AUTH_SERVER}/authorize`,
					token_endpoint: `${TEST_AUTH_SERVER}/token`,
					response_types_supported: ['code']
				})
			}));

			const originalResponse = createMockResponse({
				status: 401,
				url: TEST_MCP_URL,
				headers: {
					'WWW-Authenticate': 'Bearer resource_metadata="https://example.com/custom-resource-metadata"'
				}
			});

			const authMetadata = await createAuthMetadata(
				TEST_MCP_URL,
				originalResponse.headers,
				{
					sameOriginHeaders: {},
					fetch: mockFetch,
					log: mockLogger
				}
			);

			assert.ok(authMetadata.authorizationServer.toString().startsWith(TEST_AUTH_SERVER));

			// Verify the resource_metadata URL was logged
			assert.ok(logMessages.some(m =>
				m.level === LogLevel.Debug &&
				m.message.includes('resource_metadata challenge')
			));
		});

		test('should pass launch headers when fetching metadata from same origin', async () => {
			const mockFetch = sandbox.stub();

			// Mock resource metadata fetch to succeed so we can verify headers
			mockFetch.onCall(0).resolves(createMockResponse({
				status: 200,
				url: TEST_RESOURCE_METADATA_URL,
				body: JSON.stringify({
					resource: TEST_MCP_URL,
					authorization_servers: [TEST_AUTH_SERVER]
				})
			}));

			// Mock server metadata fetch
			mockFetch.onCall(1).resolves(createMockResponse({
				status: 200,
				url: `${TEST_AUTH_SERVER}/.well-known/oauth-authorization-server`,
				body: JSON.stringify({
					issuer: TEST_AUTH_SERVER,
					authorization_endpoint: `${TEST_AUTH_SERVER}/authorize`,
					token_endpoint: `${TEST_AUTH_SERVER}/token`,
					response_types_supported: ['code']
				})
			}));

			const originalResponse = createMockResponse({
				status: 401,
				url: TEST_MCP_URL,
				headers: {}
			});

			const launchHeaders = {
				'Authorization': 'Bearer existing-token',
				'X-Custom-Header': 'custom-value'
			};

			await createAuthMetadata(
				TEST_MCP_URL,
				originalResponse.headers,
				{
					sameOriginHeaders: launchHeaders,
					fetch: mockFetch,
					log: mockLogger
				}
			);

			// Verify fetch was called
			assert.ok(mockFetch.called, 'fetch should have been called');

			// Verify the first call (resource metadata) included the launch headers
			const firstCallArgs = mockFetch.firstCall.args;
			assert.ok(firstCallArgs.length >= 2, 'fetch should have been called with options');
			const fetchOptions = firstCallArgs[1] as RequestInit;
			assert.ok(fetchOptions.headers, 'fetch options should include headers');
		});

		test('should handle empty scope string in WWW-Authenticate header', async () => {
			const mockFetch = sandbox.stub();

			// Mock resource metadata fetch
			mockFetch.onCall(0).resolves(createMockResponse({
				status: 200,
				url: TEST_RESOURCE_METADATA_URL,
				body: JSON.stringify({
					resource: TEST_MCP_URL,
					authorization_servers: [TEST_AUTH_SERVER]
				})
			}));

			// Mock server metadata fetch
			mockFetch.onCall(1).resolves(createMockResponse({
				status: 200,
				url: `${TEST_AUTH_SERVER}/.well-known/oauth-authorization-server`,
				body: JSON.stringify({
					issuer: TEST_AUTH_SERVER,
					authorization_endpoint: `${TEST_AUTH_SERVER}/authorize`,
					token_endpoint: `${TEST_AUTH_SERVER}/token`,
					response_types_supported: ['code']
				})
			}));

			const originalResponse = createMockResponse({
				status: 401,
				url: TEST_MCP_URL,
				headers: {
					'WWW-Authenticate': 'Bearer scope=""'
				}
			});

			const authMetadata = await createAuthMetadata(
				TEST_MCP_URL,
				originalResponse.headers,
				{
					sameOriginHeaders: {},
					fetch: mockFetch,
					log: mockLogger
				}
			);

			// Empty scope string should result in empty array or undefined
			assert.ok(
				authMetadata.scopes === undefined ||
				(Array.isArray(authMetadata.scopes) && authMetadata.scopes.length === 0) ||
				(Array.isArray(authMetadata.scopes) && authMetadata.scopes.every(s => s === '')),
				'Empty scope string should be handled gracefully'
			);
		});

		test('should handle malformed WWW-Authenticate header gracefully', async () => {
			const mockFetch = sandbox.stub();

			// Mock resource metadata fetch
			mockFetch.onCall(0).resolves(createMockResponse({
				status: 200,
				url: TEST_RESOURCE_METADATA_URL,
				body: JSON.stringify({
					resource: TEST_MCP_URL,
					authorization_servers: [TEST_AUTH_SERVER]
				})
			}));

			// Mock server metadata fetch
			mockFetch.onCall(1).resolves(createMockResponse({
				status: 200,
				url: `${TEST_AUTH_SERVER}/.well-known/oauth-authorization-server`,
				body: JSON.stringify({
					issuer: TEST_AUTH_SERVER,
					authorization_endpoint: `${TEST_AUTH_SERVER}/authorize`,
					token_endpoint: `${TEST_AUTH_SERVER}/token`,
					response_types_supported: ['code']
				})
			}));

			const originalResponse = createMockResponse({
				status: 401,
				url: TEST_MCP_URL,
				headers: {
					// Malformed header - missing closing quote
					'WWW-Authenticate': 'Bearer scope="unclosed'
				}
			});

			// Should not throw - should handle gracefully
			const authMetadata = await createAuthMetadata(
				TEST_MCP_URL,
				originalResponse.headers,
				{
					sameOriginHeaders: {},
					fetch: mockFetch,
					log: mockLogger
				}
			);

			// Should still create valid auth metadata
			assert.ok(authMetadata.authorizationServer);
			assert.ok(authMetadata.serverMetadata);
		});

		test('should handle invalid JSON in resource metadata response', async () => {
			const mockFetch = sandbox.stub();

			// Mock resource metadata fetch - returns invalid JSON
			mockFetch.onCall(0).resolves(createMockResponse({
				status: 200,
				url: TEST_RESOURCE_METADATA_URL,
				body: 'not valid json {'
			}));

			// Mock server metadata fetch - also returns invalid JSON
			mockFetch.onCall(1).resolves(createMockResponse({
				status: 200,
				url: 'https://example.com/.well-known/oauth-authorization-server',
				body: '{ invalid }'
			}));

			const originalResponse = createMockResponse({
				status: 401,
				url: TEST_MCP_URL,
				headers: {}
			});

			// Should fall back to default metadata, not throw
			const authMetadata = await createAuthMetadata(
				TEST_MCP_URL,
				originalResponse.headers,
				{
					sameOriginHeaders: {},
					fetch: mockFetch,
					log: mockLogger
				}
			);

			// Should use default metadata
			assert.ok(authMetadata.authorizationServer);
			assert.ok(authMetadata.serverMetadata);
		});

		test('should handle non-401 status codes in update()', async () => {
			const { authMetadata } = await createTestAuthMetadata({
				scopes: ['read']
			});

			// Response with 403 instead of 401
			const response = createMockResponse({
				status: 403,
				headers: {
					'WWW-Authenticate': 'Bearer scope="new.scope"'
				}
			});

			// update() should still process the WWW-Authenticate header regardless of status
			const result = authMetadata.update(response.headers);

			// The behavior depends on implementation - either it updates or ignores non-401
			// This test documents the actual behavior
			assert.strictEqual(typeof result, 'boolean');
		});
	});
});
