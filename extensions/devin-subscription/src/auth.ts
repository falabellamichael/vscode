/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { createHash, randomBytes } from 'node:crypto';
import { createServer } from 'node:http';
import type { AddressInfo } from 'node:net';

// These endpoints match SimpleRAG and Devin's current public browser flow.
// The exchange mirrors SimpleRAG's working sign-in: an authorization-code
// grant with the Devin CLI client id, the loopback redirect when present,
// and the PKCE verifier.
export const DEVIN_AUTHORIZE_URL = 'https://app.devin.ai/auth/cli/continue';
export const DEVIN_TOKEN_URL = 'https://api.devin.ai/auth/cli/token';
export const DEVIN_CLIENT_ID = '3GUryQ7ldAeKEuD2obYnppsnmj58eP5u';

export interface SignInOptions {
	readonly openExternal: (url: string) => Promise<boolean>;
	readonly signal: AbortSignal;
	readonly readAuthorizationCode?: (signal: AbortSignal) => Promise<string | undefined>;
	readonly onProgress?: (stage: 'waitingForBrowser' | 'waitingForCode' | 'exchangingCode') => void;
	readonly timeoutMs?: number;
	readonly request?: typeof fetch;
}

export async function signInToDevin(options: SignInOptions): Promise<string> {
	const readAuthorizationCode = options.readAuthorizationCode;
	const verifier = randomBytes(32).toString('base64url');
	const state = randomBytes(32).toString('base64url');
	const challenge = createHash('sha256').update(verifier).digest('base64url');
	const controller = new AbortController();
	const signal = AbortSignal.any([options.signal, controller.signal]);
	const timer = setTimeout(() => controller.abort(new Error('Devin sign-in timed out.')), options.timeoutMs ?? 600_000);
	const server = createServer();
	let rejectCallback: ((reason: unknown) => void) | undefined;
	let rejectBrowserLaunch: ((reason: unknown) => void) | undefined;
	const onAbort = () => {
		const reason = signal.reason ?? new Error('Devin sign-in cancelled.');
		rejectCallback?.(reason);
		rejectBrowserLaunch?.(reason);
	};
	signal.addEventListener('abort', onAbort, { once: true });
	try {
		signal.throwIfAborted();
		let redirectUri: string | undefined;
		let callback: Promise<string> | undefined;
		if (!readAuthorizationCode) {
			await new Promise<void>((resolve, reject) => {
				server.once('error', reject);
				server.listen(0, '127.0.0.1', () => { server.removeListener('error', reject); resolve(); });
			});
			const port = (server.address() as AddressInfo).port;
			redirectUri = `http://127.0.0.1:${port}/callback`;
			const callbackBase = redirectUri;
			callback = new Promise<string>((resolve, reject) => {
				rejectCallback = reject;
				server.once('error', reject);
				server.on('request', (request, response) => {
					response.setHeader('Content-Type', 'text/plain; charset=utf-8');
					response.setHeader('Cache-Control', 'no-store');
					response.setHeader('X-Content-Type-Options', 'nosniff');
					response.setHeader('Content-Security-Policy', 'default-src \'none\'; frame-ancestors \'none\'');
					let url: URL;
					try { url = new URL(request.url ?? '/', callbackBase); } catch {
						response.writeHead(400).end('Invalid callback URL.');
						return;
					}
					if (request.method !== 'GET' || request.headers.host !== `127.0.0.1:${port}` || url.pathname !== '/callback') {
						response.writeHead(404).end('Not found.');
						return;
					}
					if (url.searchParams.getAll('state').length !== 1 || url.searchParams.get('state') !== state) {
						response.writeHead(400).end('Invalid sign-in state. Return to the original sign-in page.');
						return;
					}
					if (url.searchParams.has('error')) {
						response.writeHead(400).end('Sign-in was declined. Return to Agent Workspace.');
						reject(new Error('Devin sign-in was declined.'));
						return;
					}
					const code = url.searchParams.get('code');
					if (!code || code.length > 8192 || url.searchParams.getAll('code').length !== 1) {
						response.writeHead(400).end('No valid authorization code was received.');
						return;
					}
					response.writeHead(200).end('Authorization received. Return to Agent Workspace to check the sign-in result.');
					resolve(code);
				});
			});
			// Browser failures/cancellation can win before we start awaiting the code.
			void callback.catch(() => { });
		}
		signal.throwIfAborted();
		const authorize = new URL(DEVIN_AUTHORIZE_URL);
		for (const [key, value] of Object.entries({ response_type: 'code', state, prompt: 'select_account', code_challenge: challenge, code_challenge_method: 'S256', cli_pkce_marker: '1' })) {
			authorize.searchParams.set(key, value);
		}
		// Devin displays a copyable code when redirect_uri is omitted.
		if (redirectUri) { authorize.searchParams.set('redirect_uri', redirectUri); }
		// A stalled opener must not hold the listener or block a fresh sign-in after cancellation.
		const browserOpened = await new Promise<boolean>((resolve, reject) => {
			rejectBrowserLaunch = reject;
			options.openExternal(authorize.toString()).then(resolve, reject);
		});
		rejectBrowserLaunch = undefined;
		if (!browserOpened) {
			throw new Error('The Devin sign-in page could not be opened.');
		}
		signal.throwIfAborted();
		options.onProgress?.(readAuthorizationCode ? 'waitingForCode' : 'waitingForBrowser');
		const code = readAuthorizationCode ? await new Promise<string | undefined>((resolve, reject) => {
			rejectCallback = reject;
			readAuthorizationCode(signal).then(resolve, reject);
		}) : await callback;
		signal.throwIfAborted();
		if (code === undefined) { throw new Error('Devin sign-in cancelled.'); }
		if (!code || code.length > 8192 || (readAuthorizationCode && /\s/.test(code))) { throw new Error('Enter the authorization code shown by the Devin sign-in page.'); }
		options.onProgress?.('exchangingCode');
		const response = await (options.request ?? fetch)(DEVIN_TOKEN_URL, {
			method: 'POST', headers: { 'Content-Type': 'application/json' }, redirect: 'error',
			signal: AbortSignal.any([signal, AbortSignal.timeout(30_000)]),
			body: JSON.stringify({
				grant_type: 'authorization_code',
				client_id: DEVIN_CLIENT_ID,
				code,
				// The manual browser-code flow has no loopback redirect to echo back.
				...(redirectUri ? { redirect_uri: redirectUri } : {}),
				code_verifier: verifier,
				cli_pkce_marker: '1',
			}),
		});
		if (!response.ok) {
			throw new Error(`Devin token exchange failed (HTTP ${response.status}).`);
		}
		const data: unknown = await response.json();
		if (!data || typeof data !== 'object' || Object.hasOwn(data, 'error')) {
			throw new Error('Devin did not return a valid sign-in response.');
		}
		const record = data as Record<string, unknown>;
		// The live exchange returns { token, webapp_host }; devin_session_token and
		// session_token are accepted as older field names.
		const token = [record.devin_session_token, record.session_token, record.token].find((value): value is string => typeof value === 'string' && !!value.trim());
		if (!token) {
			// Key names only: response values may carry credentials and stay private.
			const keys = Object.keys(record).sort().slice(0, 12).join(', ');
			throw new Error(`Devin did not return a session token${keys ? ` (response keys: ${keys})` : ''}.`);
		}
		return token;
	} finally {
		rejectBrowserLaunch = undefined;
		clearTimeout(timer);
		signal.removeEventListener('abort', onAbort);
		rejectCallback?.(new Error('Devin sign-in ended.'));
		server.closeAllConnections();
		await new Promise<void>(resolve => server.close(() => resolve()));
	}
}
