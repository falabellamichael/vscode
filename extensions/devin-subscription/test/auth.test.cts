/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

const assert: typeof import('node:assert/strict') = require('node:assert/strict');
const { createHash }: typeof import('node:crypto') = require('node:crypto');
const { test }: typeof import('node:test') = require('node:test');
const { signInToDevin, DEVIN_AUTHORIZE_URL, DEVIN_TOKEN_URL, DEVIN_CLIENT_ID }: typeof import('../src/auth.ts') = require('../out/auth.js');

function callbackUrl(authorize: URL, changes: Record<string, string | null> = {}) {
	const redirectUri = authorize.searchParams.get('redirect_uri');
	const state = authorize.searchParams.get('state');
	assert.ok(redirectUri && state);
	const url = new URL(redirectUri);
	url.searchParams.set('state', state);
	url.searchParams.set('code', 'test-authorization-code');
	for (const [key, value] of Object.entries(changes)) {
		if (value === null) { url.searchParams.delete(key); } else { url.searchParams.set(key, value); }
	}
	return url;
}

test('manual browser code omits the loopback redirect and exchanges its PKCE-bound code', async () => {
	let authorize: URL | undefined;
	const stages: string[] = [];
	const token = await signInToDevin({
		signal: new AbortController().signal,
		onProgress: stage => stages.push(stage),
		openExternal: async value => {
			authorize = new URL(value);
			assert.strictEqual(authorize.origin + authorize.pathname, DEVIN_AUTHORIZE_URL);
			assert.strictEqual(authorize.searchParams.has('redirect_uri'), false);
			assert.strictEqual(authorize.searchParams.get('code_challenge_method'), 'S256');
			return true;
		},
		readAuthorizationCode: async signal => { assert.strictEqual(signal.aborted, false); return 'manual-test-code'; },
		request: async (url, options) => {
			assert.strictEqual(url, DEVIN_TOKEN_URL);
			assert.ok(options && typeof options.body === 'string' && authorize);
			const body = JSON.parse(options.body);
			assert.deepStrictEqual(Object.keys(body).sort(), ['cli_pkce_marker', 'client_id', 'code', 'code_verifier', 'grant_type']);
			assert.strictEqual(body.grant_type, 'authorization_code');
			assert.strictEqual(body.client_id, DEVIN_CLIENT_ID);
			assert.strictEqual(body.cli_pkce_marker, '1');
			assert.strictEqual(body.code, 'manual-test-code');
			assert.strictEqual(createHash('sha256').update(body.code_verifier).digest('base64url'), authorize.searchParams.get('code_challenge'));
			return Response.json({ session_token: 'manual-test-token' });
		},
	});
	assert.strictEqual(token, 'manual-test-token');
	assert.deepStrictEqual(stages, ['waitingForCode', 'exchangingCode']);
});

for (const code of [undefined, '', 'a b', 'x'.repeat(8193)]) {
	test(`rejects ${code === undefined ? 'cancelled' : 'invalid length ' + code.length} manual input without an exchange`, async () => {
		await assert.rejects(signInToDevin({
			signal: new AbortController().signal,
			openExternal: async () => true,
			readAuthorizationCode: async () => code,
			request: async () => assert.fail('must not exchange cancelled or malformed input'),
		}), code === undefined ? /cancelled/ : /authorization code/);
	});
}

for (const mode of ['cancel', 'timeout']) {
	test(`${mode} releases a stalled manual code input and ignores its late result`, async () => {
		const controller = new AbortController();
		const started = Promise.withResolvers<void>();
		const input = Promise.withResolvers<string>();
		const result = signInToDevin({
			signal: controller.signal, timeoutMs: mode === 'timeout' ? 25 : 5000,
			openExternal: async () => true,
			readAuthorizationCode: async () => { started.resolve(); return input.promise; },
			request: async () => assert.fail('must not exchange a late code'),
		});
		const rejected = assert.rejects(result, mode === 'timeout' ? /timed out/ : /user cancelled/);
		try {
			await started.promise;
			if (mode === 'cancel') { controller.abort(new Error('user cancelled')); }
			await rejected;
			input.resolve('late-test-code');
			await new Promise(resolve => setImmediate(resolve));
		} finally { controller.abort(); input.resolve('late-test-code'); }
	});
}

test('uses the browser PKCE flow and fixed HTTPS token endpoint without starting a CLI', async () => {
	let authorize: URL;
	let redirect: URL | undefined;
	const stages: string[] = [];
	const token = await signInToDevin({
		signal: new AbortController().signal,
		onProgress: stage => stages.push(stage),
		openExternal: async value => {
			authorize = new URL(value);
			assert.strictEqual(authorize.origin + authorize.pathname, DEVIN_AUTHORIZE_URL);
			assert.strictEqual(authorize.searchParams.get('code_challenge_method'), 'S256');
			assert.strictEqual(authorize.searchParams.has('token'), false);
			assert.strictEqual(authorize.searchParams.has('code_verifier'), false);
			redirect = callbackUrl(authorize);
			const response = await fetch(redirect);
			assert.strictEqual(response.status, 200);
			assert.strictEqual(response.headers.get('cache-control'), 'no-store');
			return true;
		},
		request: async (url, options) => {
			assert.ok(options && typeof options.body === 'string');
			assert.deepStrictEqual(stages, ['waitingForBrowser', 'exchangingCode']);
			assert.strictEqual(url, DEVIN_TOKEN_URL);
			assert.strictEqual(options.method, 'POST');
			assert.strictEqual(options.redirect, 'error');
			const body = JSON.parse(options.body);
			assert.strictEqual(body.code, 'test-authorization-code');
			assert.deepStrictEqual(Object.keys(body).sort(), ['cli_pkce_marker', 'client_id', 'code', 'code_verifier', 'grant_type', 'redirect_uri']);
			assert.strictEqual(body.grant_type, 'authorization_code');
			assert.strictEqual(body.client_id, DEVIN_CLIENT_ID);
			assert.strictEqual(body.cli_pkce_marker, '1');
			assert.strictEqual(body.redirect_uri, authorize.searchParams.get('redirect_uri'));
			assert.strictEqual(createHash('sha256').update(body.code_verifier).digest('base64url'), authorize.searchParams.get('code_challenge'));
			return Response.json({ session_token: 'test-session-token' });
		},
	});
	assert.strictEqual(token, 'test-session-token');
	assert.deepStrictEqual(stages, ['waitingForBrowser', 'exchangingCode']);
	assert.ok(redirect);
	await assert.rejects(fetch(redirect), /fetch failed/);
});

test('accepts the devin_session_token exchange response field', async () => {
	const token = await signInToDevin({
		signal: new AbortController().signal,
		openExternal: async value => { await fetch(callbackUrl(new URL(value))); return true; },
		request: async () => Response.json({ devin_session_token: 'test-devin-session-token' }),
	});
	assert.strictEqual(token, 'test-devin-session-token');
});

test('accepts the live exchange response shape of token and webapp_host', async () => {
	const token = await signInToDevin({
		signal: new AbortController().signal,
		openExternal: async value => { await fetch(callbackUrl(new URL(value))); return true; },
		request: async () => Response.json({ token: 'test-live-session-token', webapp_host: 'https://app.devin.ai' }),
	});
	assert.strictEqual(token, 'test-live-session-token');
});

test('ignores missing or wrong state and unrelated paths until the valid callback arrives', async () => {
	let exchanges = 0;
	await signInToDevin({
		signal: new AbortController().signal,
		openExternal: async value => {
			const authorize = new URL(value);
			for (const state of [null, 'wrong-state']) {
				assert.strictEqual((await fetch(callbackUrl(authorize, { state }))).status, 400);
			}
			const unrelated = callbackUrl(authorize);
			unrelated.pathname = '/favicon.ico';
			assert.strictEqual((await fetch(unrelated)).status, 404);
			const duplicate = callbackUrl(authorize);
			const state = authorize.searchParams.get('state');
			assert.ok(state);
			duplicate.searchParams.append('state', state);
			assert.strictEqual((await fetch(duplicate)).status, 400);
			assert.strictEqual(exchanges, 0);
			assert.strictEqual((await fetch(callbackUrl(authorize))).status, 200);
			return true;
		},
		request: async () => { exchanges++; return Response.json({ session_token: 'test-session-token' }); },
	});
	assert.strictEqual(exchanges, 1);
});

test('cancellation closes the loopback listener without exchanging credentials', async () => {
	const controller = new AbortController();
	let redirect: URL | undefined;
	const result = signInToDevin({
		signal: controller.signal,
		openExternal: async value => { redirect = callbackUrl(new URL(value)); queueMicrotask(() => controller.abort(new Error('user cancelled'))); return true; },
		request: async () => { assert.fail('must not exchange after cancellation'); },
	});
	await assert.rejects(result, /user cancelled/);
	assert.ok(redirect);
	await assert.rejects(fetch(redirect), /fetch failed/);
});

test('timeout closes the listener and does not report successful sign-in', async () => {
	let redirect: URL | undefined;
	await assert.rejects(signInToDevin({
		signal: new AbortController().signal, timeoutMs: 50,
		openExternal: async value => { redirect = callbackUrl(new URL(value)); return true; },
	}), /timed out/);
	assert.ok(redirect);
	await assert.rejects(fetch(redirect), /fetch failed/);
});

for (const mode of ['cancel', 'timeout']) {
	for (const lateFailure of [false, true]) {
		test(`${mode} releases a pending browser launch and handles late ${lateFailure ? 'failure' : 'success'}`, async () => {
			const controller = new AbortController();
			let redirect: URL | undefined;
			let exchanges = 0;
			const started = Promise.withResolvers<void>();
			const browser = Promise.withResolvers<boolean>();
			const result = signInToDevin({
				signal: controller.signal, timeoutMs: mode === 'timeout' ? 25 : 5000,
				openExternal: value => { redirect = callbackUrl(new URL(value)); started.resolve(); return browser.promise; },
				request: async () => { exchanges++; assert.fail('must not exchange a cancelled browser launch'); },
			});
			const outcome = result.then(() => ({ kind: 'success' }), error => ({ kind: 'error', message: error.message }));
			try {
				await started.promise;
				if (mode === 'cancel') { controller.abort(new Error('user cancelled')); }
				assert.deepStrictEqual(await Promise.race([outcome, new Promise(resolve => setTimeout(() => resolve({ kind: 'pending' }), 200))]), { kind: 'error', message: mode === 'cancel' ? 'user cancelled' : 'Devin sign-in timed out.' });
				assert.ok(redirect);
				await assert.rejects(fetch(redirect), /fetch failed/);
				if (lateFailure) { browser.reject(new Error('late browser failure')); } else { browser.resolve(true); }
				await new Promise(resolve => setImmediate(resolve));
				assert.strictEqual(exchanges, 0);
			} finally {
				browser.resolve(true);
				controller.abort();
				await outcome;
			}
		});
	}
}

test('a browser launch failure releases the listener', async () => {
	let redirect: URL | undefined;
	await assert.rejects(signInToDevin({
		signal: new AbortController().signal,
		openExternal: async value => { redirect = callbackUrl(new URL(value)); return false; },
	}), /could not be opened/);
	assert.ok(redirect);
	await assert.rejects(fetch(redirect), /fetch failed/);
});

test('provider denial does not expose its description or exchange a code', async () => {
	await assert.rejects(signInToDevin({
		signal: new AbortController().signal,
		openExternal: async value => { await fetch(callbackUrl(new URL(value), { error: 'access_denied', error_description: 'private provider details' })); return true; },
		request: async () => { assert.fail('must not exchange a denied callback'); },
	}), error => error instanceof Error && error.message === 'Devin sign-in was declined.');
});

for (const [name, response, expected] of [
	['HTTP failure', new Response('private response contents', { status: 401 }), /HTTP 401/],
	['error response', Response.json({ error: 'private error details' }), /valid sign-in response/],
	['missing token', Response.json({ nested: { arbitrary_session: 'must-not-be-used' } }), /session token \(response keys: nested\)/],
	['unrelated token field', Response.json({ access_token: 'must-not-be-used' }), /session token/],
] as const) {
	test(`rejects ${name} without treating it as a connected account`, async () => {
		await assert.rejects(signInToDevin({
			signal: new AbortController().signal,
			openExternal: async value => { await fetch(callbackUrl(new URL(value))); return true; },
			request: async () => response,
		}), expected);
	});
}
