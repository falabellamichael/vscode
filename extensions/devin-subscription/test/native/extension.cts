/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import type { AuthenticationProvider, AuthenticationProviderOptions, Disposable, ExtensionContext } from 'vscode';
import type { SignInOptions } from '../../src/auth.ts';

const assert: typeof import('node:assert/strict') = require('node:assert/strict');
const fs: typeof import('node:fs/promises') = require('node:fs/promises');
const path: typeof import('node:path') = require('node:path');
const Module: { _load(this: unknown, name: string, parent?: NodeModule, isMain?: boolean): unknown } = require('node:module');
const vscode: typeof import('vscode') = require('vscode');

/** Run the compiled provider against native VS Code storage and UI with synthetic browser approval. */
export function activate(context: ExtensionContext): void {
	const directory = path.dirname(context.extensionPath);
	const output = path.join(directory, '..', 'reports', `${process.pid}.json`);
	const production = path.join(context.extensionPath, 'production', 'extension.js');
	const auth: typeof import('../../src/auth.ts') = require(path.join(context.extensionPath, 'production', 'auth.js'));
	let provider: AuthenticationProvider | undefined;
	let role = 'unclaimed';
	let exchanges = 0;
	let phase = 'idle';
	let result = 'idle';
	let pending = false;
	let writes = Promise.resolve();
	const events: { added: string[]; removed: string[]; changed: string[] }[] = [];
	const report = async () => {
		assert.ok(provider);
		const sessions = await provider.getSessions(undefined, {});
		const value = JSON.stringify({ role, pid: process.pid, result, pending, phase, exchanges, events: [...events], sessions: sessions.map(session => ({ id: session.id, accountId: session.account.id })) });
		writes = writes.then(() => fs.writeFile(output, value));
		await writes;
	};
	const originalLoad = Module._load;
	Module._load = function (this: unknown, name: string, ...args: [NodeModule?, boolean?]) {
		if (args[0]?.filename === production) {
			if (name === 'vscode') {
				return {
					...vscode,
					authentication: {
						...vscode.authentication,
						registerAuthenticationProvider: (id: string, label: string, instance: AuthenticationProvider, options: AuthenticationProviderOptions): Disposable => {
							assert.strictEqual(id, 'devin-subscription');
							provider = instance;
							context.subscriptions.push(instance.onDidChangeSessions(event => {
								events.push({ added: event.added?.map(session => session.id) ?? [], removed: event.removed?.map(session => session.id) ?? [], changed: event.changed?.map(session => session.id) ?? [] });
								void report();
							}));
							return vscode.authentication.registerAuthenticationProvider(id, label, instance, options);
						},
					},
				};
			}
			if (name === './auth') {
				return {
					signInToDevin: (options: SignInOptions) => auth.signInToDevin({
						...options,
						onProgress: stage => { phase = stage; options.onProgress?.(stage); void report(); },
						request: async (url, request) => {
							assert.strictEqual(url, auth.DEVIN_TOKEN_URL);
							assert.ok(request && typeof request.body === 'string');
							const body = JSON.parse(request.body);
							assert.strictEqual(body.code, 'native-fixture-code');
							assert.strictEqual(body.grant_type, 'authorization_code');
							assert.strictEqual(body.client_id, auth.DEVIN_CLIENT_ID);
							assert.strictEqual(body.cli_pkce_marker, '1');
							const keys = ['cli_pkce_marker', 'client_id', 'code', 'code_verifier', 'grant_type'];
							if (body.redirect_uri !== undefined) {
								assert.ok(/^http:\/\/127\.0\.0\.1:\d+\/callback$/.test(body.redirect_uri));
								keys.push('redirect_uri');
							}
							assert.deepStrictEqual(Object.keys(body).sort(), keys.sort());
							exchanges++;
							return Response.json({ token: `native-fixture-token-${process.pid}-${exchanges}`, webapp_host: 'https://app.devin.ai' });
						},
					}),
				};
			}
		}
		return originalLoad.call(this, name, ...args);
	};
	try { require(production).activate(context); } finally { Module._load = originalLoad; }
	assert.ok(provider);
	const authentication = provider;
	for (const name of ['Agents', 'IDE']) {
		context.subscriptions.push(vscode.commands.registerCommand(`signalCFixture.claim${name}`, async () => { role = name.toLowerCase(); await report(); }));
	}
	context.subscriptions.push(vscode.commands.registerCommand('signalCFixture.signIn', async () => {
		pending = true;
		result = 'pending';
		await report();
		try {
			await authentication.createSession([], {});
			result = 'signedIn';
		} catch (error) {
			result = error instanceof vscode.CancellationError ? 'cancelled' : 'failed';
			if (result === 'failed') { throw error; }
		} finally { pending = false; await report(); }
	}));
	context.subscriptions.push(vscode.commands.registerCommand('signalCFixture.signOut', async () => {
		for (const session of await authentication.getSessions(undefined, {})) { await authentication.removeSession(session.id); }
		await report();
	}));
	void report();
}
