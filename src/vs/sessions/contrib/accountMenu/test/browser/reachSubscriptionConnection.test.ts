/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import assert from 'assert';
import { encodeBase64, VSBuffer } from '../../../../../base/common/buffer.js';
import { CancellationTokenSource } from '../../../../../base/common/cancellation.js';
import { Event } from '../../../../../base/common/event.js';
import { upcastPartial } from '../../../../../base/test/common/mock.js';
import { ensureNoDisposablesAreLeakedInTestSuite } from '../../../../../base/test/common/utils.js';
import { INativeHostService, REACH_ACCOUNT_ORIGIN, ReachAuthenticationRequest } from '../../../../../platform/native/common/native.js';
import { IProgressService } from '../../../../../platform/progress/common/progress.js';
import { ISecretStorageService } from '../../../../../platform/secrets/common/secrets.js';
import { ReachSubscriptionConnection } from '../../browser/reachSubscriptionConnection.js';

suite('REACH Subscription Connection', () => {
	ensureNoDisposablesAreLeakedInTestSuite();
	const origin = REACH_ACCOUNT_ORIGIN;
	const flowId = 'a'.repeat(64);

	function fixture(browserOrigin = origin, cancelAtExchange = false, rejectFirstFlow = false) {
		const requests: ReachAuthenticationRequest[] = [];
		const opened: string[] = [];
		const secrets = new Map<string, string>();
		const accessToken = 'mock-reach-subscription-token';
		let cancel: (() => void) | undefined;
		const prompts: string[][] = [];
		const turnIds: (string | undefined)[] = [];
		const finishes: (string | undefined)[] = [];
		let started = 0;
		const service = new ReachSubscriptionConnection(upcastPartial<INativeHostService>({ requestReachAuthentication: async request => {
			requests.push(request);
			if (request.action === 'start') { started++; }
			if (rejectFirstFlow && started === 1 && request.action === 'exchange') { return { status: 410 }; }
			const currentFlow = started > 1 ? 'b'.repeat(64) : flowId;
			const expiresAt = new Date(Date.now() + 60_000).toISOString();
			const data = request.action === 'start' ? { flowId: currentFlow, expiresAt, loginUrl: `${browserOrigin}/wallet/connect#flow=${currentFlow}` }
				: { accessToken, expiresAt, account: { id: 'reach-account', access: { activeSubscription: true } } };
			if (cancelAtExchange && request.action === 'exchange') { cancel?.(); }
			return { status: 200, data };
		}, openExternal: async url => { opened.push(url); return true; },
		sendReachAssistantPrompt: async (id, token, request, _requestId, _conversationId, turnId) => {
			prompts.push([id, token, request.messages.at(-1)?.content ?? '']);
			turnIds.push(turnId);
			return { kind: 'reply', text: 'Hello' };
		},
		finishReachAssistantTurn: async turnId => { finishes.push(turnId); } }), upcastPartial<ISecretStorageService>({
			onDidChangeSecret: Event.None,
			get: async key => secrets.get(key),
			set: async (key, value) => { secrets.set(key, value); },
			delete: async key => { secrets.delete(key); },
		}), upcastPartial<IProgressService>({
			withProgress: async (_options, task, onCancel) => { cancel = onCancel; return task({ report() { } }); },
		}));
		return { service, requests, opened, secrets, accessToken, prompts, turnIds, finishes };
	}

	test('uses a browser proof, saves only the authenticated account, and signs out', async () => {
		const { service, requests, opened, secrets, accessToken } = fixture();
		const account = await service.signIn();
		assert.deepStrictEqual(account, { id: 'reach-account', subscriptionActive: true });
		assert.deepStrictEqual(await service.getAccount(), account);
		assert.deepStrictEqual(opened, [`https://signalreach.ca/login.html#client=studio&flow=${flowId}`]);
		assert.ok(!opened[0].includes(accessToken));
		const [start, exchange] = requests;
		assert.ok(start.action === 'start' && exchange.action === 'exchange');
		assert.strictEqual(start.state, exchange.state);
		assert.notStrictEqual(start.state, exchange.codeVerifier);
		const digest = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(exchange.codeVerifier));
		assert.strictEqual(start.codeChallenge, encodeBase64(VSBuffer.wrap(new Uint8Array(digest)), false, true));
		assert.strictEqual(secrets.size, 1);
		await service.signOut();
		assert.strictEqual(secrets.size, 0);
		assert.strictEqual(await service.getAccount(), undefined);
		assert.deepStrictEqual(requests[2], { action: 'logout', accessToken });
	});

	test('rejects a login URL outside the account service without opening it', async () => {
		const { service, opened, secrets } = fixture('https://example.invalid');
		await assert.rejects(service.signIn(), /invalid sign-in response/i);
		assert.deepStrictEqual(opened, []);
		assert.strictEqual(secrets.size, 0);
	});

	test('starts a fresh browser proof when the service rejects the initial flow before polling', async () => {
		const { service, opened, requests } = fixture(origin, false, true);
		assert.deepStrictEqual(await service.signIn(), { id: 'reach-account', subscriptionActive: true });
		assert.deepStrictEqual(requests.map(request => request.action), ['start', 'exchange', 'start', 'exchange']);
		assert.notDeepStrictEqual(requests[0], requests[2]);
		assert.strictEqual(new Set(opened).size, 2);
	});

	test('only sends Assistant prompts under the current signed-in subscription', async () => {
		const { service, prompts, accessToken } = fixture();
		const request = { messages: [{ role: 'user' as const, content: 'Hello' }] };
		assert.deepStrictEqual(await service.sendAssistantPrompt('reach-account', request), { kind: 'signedOut' });
		await service.signIn();
		assert.deepStrictEqual(await service.sendAssistantPrompt('another-account', request), { kind: 'signedOut' });
		assert.deepStrictEqual(await service.sendAssistantPrompt('reach-account', request), { kind: 'reply', text: 'Hello' });
		assert.deepStrictEqual(prompts, [['reach-account', accessToken, 'Hello']]);
	});

	test('revokes a session received after cancellation instead of saving it', async () => {
		const { service, requests, secrets } = fixture(origin, true);
		assert.strictEqual(await service.signIn(), undefined);
		assert.strictEqual(secrets.size, 0);
		assert.strictEqual(requests.at(-1)?.action, 'logout');
	});

	test('scopes native assistant turns to the renderer cancellation token', async () => {
		const { service, turnIds, finishes } = fixture();
		await service.signIn();
		const request = { messages: [{ role: 'user' as const, content: 'Hello' }] };
		const first = new CancellationTokenSource();
		await service.sendAssistantPrompt('reach-account', request, undefined, first.token);
		await service.sendAssistantPrompt('reach-account', request, undefined, first.token);
		const turnId = turnIds[0];
		assert.ok(turnId);
		assert.strictEqual(turnIds[1], turnId);
		assert.deepStrictEqual(finishes, []);
		first.dispose(true);
		assert.deepStrictEqual(finishes, [turnId]);
		const second = new CancellationTokenSource();
		await service.sendAssistantPrompt('reach-account', request, undefined, second.token);
		assert.ok(turnIds[2]);
		assert.notStrictEqual(turnIds[2], turnId);
		second.dispose(true);
		assert.deepStrictEqual(finishes, [turnId, turnIds[2]]);
	});
});
