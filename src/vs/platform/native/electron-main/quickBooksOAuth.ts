/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { randomBytes, randomUUID } from 'node:crypto';
import { lstat, mkdir, open, readFile, realpath, rename, unlink } from 'node:fs/promises';
import { createServer, ServerResponse } from 'node:http';
import { dirname, isAbsolute, join, relative, resolve } from 'node:path';
import { DeferredPromise } from '../../../base/common/async.js';
import { CancellationToken } from '../../../base/common/cancellation.js';
import { CancellationError } from '../../../base/common/errors.js';
import { localize } from '../../../nls.js';
import { isQuickBooksOAuthCredential, isSignalCQuickBooksSandboxStore, quickBooksSandboxCallbackPath, quickBooksSandboxCallbackPort } from '../common/native.js';

interface QuickBooksOAuthCallback {
	readonly code: string;
	readonly realmId: string;
	readonly response: ServerResponse;
}

interface ExistingSandboxGrant {
	readonly file: string;
	readonly contents: string;
	readonly realmId: string;
}

const refreshTokenLine = /^QUICKBOOKS_REFRESH_TOKEN=[^\r\n]*$/m;

export async function reauthorizeQuickBooksSandbox(
	tokenStorePath: string,
	userDataPath: string,
	openBrowser: (url: string) => Promise<void>,
	token: CancellationToken,
	fetchToken: (url: string, init: RequestInit) => Promise<Response> = fetch,
	port = quickBooksSandboxCallbackPort,
): Promise<void> {
	const invalidStore = () => new Error(localize('quickBooks.invalidManagedStore', "This connection must use a QuickBooks sandbox credential file created by Signal-C in this installation."));
	if (!isSignalCQuickBooksSandboxStore(tokenStorePath) || !isAbsolute(tokenStorePath)) {
		throw invalidStore();
	}
	const privateDirectory = join(userDataPath, 'quickbooks-mcp');
	const file = resolve(tokenStorePath);
	if (relative(privateDirectory, dirname(file)) !== '') {
		throw invalidStore();
	}
	let contents: string;
	try {
		const actualProfilePath = await realpath(userDataPath);
		const info = await lstat(file);
		if (!info.isFile() || info.nlink !== 1 || info.size > 64 * 1024
			|| relative(join(actualProfilePath, 'quickbooks-mcp'), dirname(await realpath(file))) !== '') {
			throw invalidStore();
		}
		contents = await readFile(file, 'utf8');
	} catch {
		throw invalidStore();
	}
	const getValue = (name: string) => {
		const assignment = new RegExp(`^\\s*(?:export\\s+)?${name}\\s*(?:=|:\\s)`, 'i');
		const lines = contents.split(/\r?\n/).filter(line => assignment.test(line));
		return lines.length === 1 && lines[0].startsWith(`${name}=`) ? lines[0].slice(name.length + 1) : undefined;
	};
	const clientId = getValue('QUICKBOOKS_CLIENT_ID');
	const secret = getValue('QUICKBOOKS_CLIENT_SECRET');
	const realmId = getValue('QUICKBOOKS_REALM_ID');
	if (getValue('QUICKBOOKS_ENVIRONMENT') !== 'sandbox' || !isQuickBooksOAuthCredential(clientId) || !isQuickBooksOAuthCredential(secret)
		|| !realmId || !/^\d{1,40}$/.test(realmId) || getValue('QUICKBOOKS_REFRESH_TOKEN') === undefined) {
		throw invalidStore();
	}
	await authorizeQuickBooksSandbox(clientId, secret, userDataPath, openBrowser, token, fetchToken, port, { file, contents, realmId });
}

export async function authorizeQuickBooksSandbox(
	clientId: string,
	clientSecret: string,
	userDataPath: string,
	openBrowser: (url: string) => Promise<void>,
	token: CancellationToken,
	fetchToken: (url: string, init: RequestInit) => Promise<Response> = fetch,
	port = quickBooksSandboxCallbackPort,
	existing?: ExistingSandboxGrant,
): Promise<string> {
	if (!isQuickBooksOAuthCredential(clientId) || !isQuickBooksOAuthCredential(clientSecret)) {
		throw new Error(localize('quickBooks.invalidCredentials', "Enter a valid Intuit development client ID and secret."));
	}
	if (token.isCancellationRequested) {
		throw new CancellationError();
	}

	const state = randomBytes(32).toString('base64url');
	const callback = new DeferredPromise<QuickBooksOAuthCallback>();
	const controller = new AbortController();
	let callbackUrl: URL | undefined;
	const server = createServer((request, response) => {
		if (!callbackUrl || request.method !== 'GET' || request.headers.host !== callbackUrl.host || !request.url || !URL.canParse(request.url, callbackUrl)) {
			return respond(response, 404, 'Not found.');
		}
		const url = new URL(request.url, callbackUrl);
		if (url.origin !== callbackUrl.origin || url.pathname !== callbackUrl.pathname) {
			return respond(response, 404, 'Not found.');
		}
		if (url.searchParams.get('state') !== state) {
			return respond(response, 400, 'This callback did not match the active QuickBooks sign-in.');
		}
		if (callback.isSettled) {
			return respond(response, 409, 'This QuickBooks sign-in has already been handled.');
		}
		if (url.searchParams.has('error')) {
			respond(response, 400, 'QuickBooks authorization did not complete. Return to Signal-C.');
			void callback.error(url.searchParams.get('error') === 'access_denied' ? new CancellationError() : new Error(localize('quickBooks.denied', "Intuit did not authorize the QuickBooks connection.")));
			return;
		}
		const code = url.searchParams.get('code');
		const realmId = url.searchParams.get('realmId');
		if (!code || !realmId || !/^\d{1,40}$/.test(realmId)) {
			respond(response, 400, 'QuickBooks did not return a valid authorization response.');
			void callback.error(new Error(localize('quickBooks.invalidCallback', "QuickBooks did not return a valid company and authorization code.")));
			return;
		}
		void callback.complete({ code, realmId, response });
	});
	const cancellation = token.onCancellationRequested(() => controller.abort(new CancellationError()));
	const timer = setTimeout(() => controller.abort(new Error(localize('quickBooks.signInTimeout', "The QuickBooks browser sign-in timed out."))), 5 * 60 * 1000);
	let browserResponse: ServerResponse | undefined;
	let createdFile: string | undefined;
	try {
		await new Promise<void>((resolve, reject) => {
			server.once('error', () => reject(new Error(localize('quickBooks.listen', "Could not start the QuickBooks callback on localhost port {0}. Close any other QuickBooks sign-in and try again.", port))));
			server.listen(port, '127.0.0.1', resolve);
		});
		controller.signal.throwIfAborted();
		const address = server.address();
		if (!address || typeof address === 'string') {
			throw new Error(localize('quickBooks.noCallback', "Could not start the QuickBooks callback."));
		}
		callbackUrl = new URL(`http://localhost:${address.port}${quickBooksSandboxCallbackPath}`);
		const authorizationUrl = new URL('https://appcenter.intuit.com/connect/oauth2');
		authorizationUrl.search = new URLSearchParams({
			client_id: clientId, response_type: 'code', scope: 'com.intuit.quickbooks.accounting', redirect_uri: callbackUrl.href, state,
		}).toString();
		controller.signal.addEventListener('abort', () => { void callback.error(controller.signal.reason); }, { once: true });
		void openBrowser(authorizationUrl.href).catch(() => callback.error(new Error(localize('quickBooks.openBrowser', "Could not open the browser for QuickBooks sign-in."))));
		const { code, realmId, response } = await callback.p;
		browserResponse = response;
		if (existing && existing.realmId !== realmId) {
			throw new Error(localize('quickBooks.companyChanged', "Choose the same QuickBooks sandbox company when you sign in again. Add another connection to use a different company."));
		}

		let refreshToken: string | undefined;
		try {
			const reply = await fetchToken('https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer', {
				method: 'POST', redirect: 'error', signal: controller.signal,
				headers: {
					Authorization: `Basic ${Buffer.from(`${clientId}:${clientSecret}`).toString('base64')}`,
					'Content-Type': 'application/x-www-form-urlencoded', Accept: 'application/json',
				},
				body: new URLSearchParams({ grant_type: 'authorization_code', code, redirect_uri: callbackUrl.href }),
			});
			if (reply.ok) {
				const grant = await reply.json() as { refresh_token?: unknown } | null;
				if (isQuickBooksOAuthCredential(grant?.refresh_token)) {
					refreshToken = grant.refresh_token;
				}
			} else {
				await reply.body?.cancel();
			}
		} catch {
			// OAuth error bodies can contain credentials or details about the account.
		}
		controller.signal.throwIfAborted();
		if (!refreshToken) {
			throw new Error(localize('quickBooks.exchangeFailed', "Intuit did not issue a QuickBooks refresh token. Check the development credentials and redirect URI."));
		}

		const directory = join(userDataPath, 'quickbooks-mcp');
		await mkdir(directory, { mode: 0o700, recursive: true });
		controller.signal.throwIfAborted();
		const storePath = existing?.file ?? join(directory, `${randomUUID()}.env`);
		const writePath = existing ? join(directory, `${randomUUID()}.tmp`) : storePath;
		const contents = existing ? existing.contents.replace(refreshTokenLine, `QUICKBOOKS_REFRESH_TOKEN=${refreshToken}`) : [
			`QUICKBOOKS_CLIENT_ID=${clientId}`,
			`QUICKBOOKS_CLIENT_SECRET=${clientSecret}`,
			'QUICKBOOKS_ENVIRONMENT=sandbox',
			`QUICKBOOKS_REDIRECT_URI=${callbackUrl.href}`,
			`QUICKBOOKS_REFRESH_TOKEN=${refreshToken}`,
			`QUICKBOOKS_REALM_ID=${realmId}`,
			'',
		].join('\n');
		const file = await open(writePath, 'wx', 0o600);
		createdFile = writePath;
		try {
			await file.writeFile(contents);
		} finally {
			await file.close();
		}
		controller.signal.throwIfAborted();
		if (existing) {
			if (await readFile(storePath, 'utf8') !== existing.contents) {
				throw new Error(localize('quickBooks.storeChanged', "The QuickBooks authorization file changed during sign-in. Try again after the server has stopped."));
			}
			controller.signal.throwIfAborted();
			await rename(writePath, storePath);
			createdFile = undefined;
		}
		respond(response, 200, 'QuickBooks sandbox authorized. Return to Signal-C to finish setup.');
		return storePath;
	} catch (error) {
		if (browserResponse && !browserResponse.writableEnded) {
			respond(browserResponse, 400, 'QuickBooks authorization did not complete. Return to Signal-C.');
		}
		if (createdFile) {
			await unlink(createdFile).catch(() => { /* cleanup of this attempt's file */ });
		}
		throw error;
	} finally {
		clearTimeout(timer);
		cancellation.dispose();
		server.close();
	}
}

function respond(response: ServerResponse, status: number, message: string): void {
	response.writeHead(status, { 'Content-Type': 'text/plain; charset=utf-8', 'Cache-Control': 'no-store', 'Referrer-Policy': 'no-referrer', 'X-Content-Type-Options': 'nosniff', Connection: 'close' });
	response.end(message);
}
