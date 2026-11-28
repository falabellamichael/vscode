/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import assert from 'assert';
import { mkdir, mkdtemp, readFile, readdir, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { DeferredPromise } from '../../../../base/common/async.js';
import { CancellationToken, CancellationTokenSource } from '../../../../base/common/cancellation.js';
import { isCancellationError } from '../../../../base/common/errors.js';
import { authorizeQuickBooksSandbox, reauthorizeQuickBooksSandbox } from '../../electron-main/quickBooksOAuth.js';

suite('QuickBooks sandbox OAuth', () => {
	let userData: string;
	setup(async () => { userData = await mkdtemp(join(tmpdir(), 'signal-c-quickbooks-oauth-test-')); });
	teardown(async () => { await rm(userData, { recursive: true, force: true }); });

	function callback(authorize: URL, params: Record<string, string>): URL {
		const url = new URL(authorize.searchParams.get('redirect_uri')!);
		url.search = new URLSearchParams({ state: authorize.searchParams.get('state')!, ...params }).toString();
		return url;
	}

	async function managedStore() {
		const directory = join(userData, 'quickbooks-mcp');
		await mkdir(directory);
		const path = join(directory, '10000000-0000-4000-8000-000000000000.env');
		const contents = [
			'QUICKBOOKS_CLIENT_ID=fixture-client',
			'QUICKBOOKS_CLIENT_SECRET=fixture-secret',
			'QUICKBOOKS_ENVIRONMENT=sandbox',
			'QUICKBOOKS_REALM_ID=5148',
			'QUICKBOOKS_REFRESH_TOKEN=previous-refresh',
			'QUICKBOOKS_DISABLE_WRITE=true',
			'',
		].join('\r\n');
		await writeFile(path, contents);
		return { directory, path, contents };
	}

	test('uses a state-bound browser return, exchanges the code once, and writes a separate private server file', async () => {
		const observedAuthorization = new DeferredPromise<string>();
		const exchangeStarted = new DeferredPromise<void>();
		const allowExchange = new DeferredPromise<void>();
		const requests: { url: string; body: string; authorization: string | null; redirect: RequestRedirect | undefined }[] = [];
		const authorize = authorizeQuickBooksSandbox('fixture-client', 'fixture-secret', userData, url => observedAuthorization.complete(url), CancellationToken.None, async (input, init) => {
			requests.push({ url: String(input), body: String(init?.body), authorization: new Headers(init?.headers).get('authorization'), redirect: init?.redirect });
			await exchangeStarted.complete();
			await allowExchange.p;
			return Response.json({ refresh_token: 'fixture-refresh', access_token: 'fixture-access' });
		}, 0);
		const external = new URL(await observedAuthorization.p);
		assert.strictEqual(external.origin + external.pathname, 'https://appcenter.intuit.com/connect/oauth2');
		assert.strictEqual(external.searchParams.get('scope'), 'com.intuit.quickbooks.accounting');
		assert.strictEqual(external.searchParams.get('client_id'), 'fixture-client');
		assert.ok(!external.href.includes('fixture-secret'));
		const redirectUri = new URL(external.searchParams.get('redirect_uri')!);
		assert.strictEqual(redirectUri.hostname, 'localhost');
		assert.strictEqual(redirectUri.pathname, '/callback');
		const mismatched = await fetch(callback(external, { state: 'incorrect', code: 'unexpected-code', realmId: '11' }));
		assert.strictEqual(mismatched.status, 400);
		await mismatched.body?.cancel();
		assert.strictEqual(requests.length, 0);

		const valid = callback(external, { code: 'fixture/code+more', realmId: '5148' });
		const browser = fetch(valid);
		await exchangeStarted.p;
		const duplicate = await fetch(valid);
		assert.strictEqual(duplicate.status, 409);
		await duplicate.body?.cancel();
		await allowExchange.complete();
		const result = await browser;
		assert.strictEqual(result.status, 200);
		assert.strictEqual(result.headers.get('cache-control'), 'no-store');
		assert.ok(!(await result.text()).includes('fixture-'));
		const path = await authorize;
		assert.deepStrictEqual(requests, [{
			url: 'https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer',
			body: String(new URLSearchParams({ grant_type: 'authorization_code', code: 'fixture/code+more', redirect_uri: redirectUri.href })),
			authorization: `Basic ${Buffer.from('fixture-client:fixture-secret').toString('base64')}`,
			redirect: 'error',
		}]);
		const contents = await readFile(path, 'utf8');
		assert.ok(path.startsWith(join(userData, 'quickbooks-mcp')));
		assert.match(contents, /^QUICKBOOKS_ENVIRONMENT=sandbox$/m);
		assert.match(contents, /^QUICKBOOKS_REFRESH_TOKEN=fixture-refresh$/m);
		assert.match(contents, /^QUICKBOOKS_REALM_ID=5148$/m);
		assert.ok(!contents.includes('fixture-access'));
	});

	test('declining browser consent performs no token exchange and creates no token store', async () => {
		const observedAuthorization = new DeferredPromise<string>();
		const authorize = authorizeQuickBooksSandbox('fixture-client', 'fixture-secret', userData, url => observedAuthorization.complete(url), CancellationToken.None, async () => assert.fail('Declined consent must not exchange tokens'), 0);
		const cancelled = assert.rejects(authorize, isCancellationError);
		const response = await fetch(callback(new URL(await observedAuthorization.p), { error: 'access_denied' }));
		assert.strictEqual(response.status, 400);
		await response.body?.cancel();
		await cancelled;
		assert.deepStrictEqual(await readdir(userData), []);
	});

	test('cancelling while the browser is open closes the listener without creating a token store', async () => {
		const source = new CancellationTokenSource();
		try {
			const observedAuthorization = new DeferredPromise<string>();
			const authorize = authorizeQuickBooksSandbox('fixture-client', 'fixture-secret', userData, url => observedAuthorization.complete(url), source.token, async () => assert.fail('Cancelled sign-in must not exchange tokens'), 0);
			const cancelled = assert.rejects(authorize, isCancellationError);
			const url = callback(new URL(await observedAuthorization.p), { code: 'late', realmId: '11' });
			source.cancel();
			await cancelled;
			await assert.rejects(fetch(url));
			assert.deepStrictEqual(await readdir(userData), []);
		} finally {
			source.dispose();
		}
	});

	test('renews the same company in place and preserves the other connection settings', async () => {
		const { path, directory, contents } = await managedStore();
		const observedAuthorization = new DeferredPromise<string>();
		const renew = reauthorizeQuickBooksSandbox(path, userData, url => observedAuthorization.complete(url), CancellationToken.None, async (_url, init) => {
			assert.strictEqual(new Headers(init.headers).get('authorization'), `Basic ${Buffer.from('fixture-client:fixture-secret').toString('base64')}`);
			return Response.json({ refresh_token: 'renewed-refresh' });
		}, 0);
		const request = new URL(await observedAuthorization.p);
		assert.strictEqual(request.searchParams.get('client_id'), 'fixture-client');
		const response = await fetch(callback(request, { code: 'fixture-code', realmId: '5148' }));
		assert.strictEqual(response.status, 200);
		await response.body?.cancel();
		await renew;
		assert.strictEqual(await readFile(path, 'utf8'), contents.replace('previous-refresh', 'renewed-refresh'));
		assert.strictEqual((await readdir(directory)).length, 1);
	});

	for (const scenario of ['different company', 'declined', 'concurrent token rotation'] as const) {
		test(`preserves the saved grant when QuickBooks renewal is ${scenario}`, async () => {
			const { path, directory, contents } = await managedStore();
			const observedAuthorization = new DeferredPromise<string>();
			const rotated = contents.replace('previous-refresh', 'rotated-by-server');
			let exchanged = false;
			const renew = reauthorizeQuickBooksSandbox(path, userData, url => observedAuthorization.complete(url), CancellationToken.None, async () => {
				exchanged = true;
				await writeFile(path, rotated);
				return Response.json({ refresh_token: 'renewed-refresh' });
			}, 0);
			const failure = assert.rejects(renew, scenario === 'declined' ? isCancellationError : scenario === 'different company' ? /same QuickBooks sandbox company/ : /file changed during sign-in/);
			const request = new URL(await observedAuthorization.p);
			const response = await fetch(callback(request, scenario === 'declined' ? { error: 'access_denied' } : { code: 'fixture-code', realmId: scenario === 'different company' ? '9999' : '5148' }));
			assert.strictEqual(response.status, 400);
			await response.body?.cancel();
			await failure;
			assert.strictEqual(exchanged, scenario === 'concurrent token rotation');
			assert.strictEqual(await readFile(path, 'utf8'), exchanged ? rotated : contents);
			assert.strictEqual((await readdir(directory)).length, 1);
		});
	}

	test('rejects a similar credential-file path outside the Signal-C credential directory', async () => {
		const path = join(userData, 'outside', 'quickbooks-mcp', '10000000-0000-4000-8000-000000000000.env');
		await assert.rejects(reauthorizeQuickBooksSandbox(path, userData, async () => assert.fail('An unmanaged file must not open consent'), CancellationToken.None, undefined, 0), /credential file created by Signal-C/);
	});

	test('rejects a managed file with an ambiguous company before opening the browser', async () => {
		const { path, contents } = await managedStore();
		await writeFile(path, contents + 'export quickbooks_realm_id = 9999\r\n');
		await assert.rejects(reauthorizeQuickBooksSandbox(path, userData, async () => assert.fail('Ambiguous company must not open consent'), CancellationToken.None, undefined, 0), /credential file created by Signal-C/);
	});
});
