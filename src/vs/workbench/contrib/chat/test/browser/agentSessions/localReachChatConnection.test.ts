/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import assert from 'assert';
import { DeferredPromise } from '../../../../../../base/common/async.js';
import { Emitter, Event } from '../../../../../../base/common/event.js';
import { upcastPartial } from '../../../../../../base/test/common/mock.js';
import { ensureNoDisposablesAreLeakedInTestSuite } from '../../../../../../base/test/common/utils.js';
import { IAgentHostService } from '../../../../../../platform/agentHost/common/agentService.js';
import { RootState } from '../../../../../../platform/agentHost/common/state/sessionState.js';
import { NullLogService } from '../../../../../../platform/log/common/log.js';
import { REACH_ACCOUNT_ORIGIN, REACH_CHAT_PROVIDER_ID } from '../../../../../../platform/native/common/native.js';
import { TestNotificationService } from '../../../../../../platform/notification/test/common/testNotificationService.js';
import { IReachSubscriptionConnection } from '../../../../../services/authentication/browser/reachSubscriptionConnection.js';
import { IWorkbenchEnvironmentService } from '../../../../../services/environment/common/environmentService.js';
import { LocalReachChatConnection } from '../../../browser/agentSessions/agentHost/localReachChatConnection.js';

suite('Local REACH chat connection', () => {
	const disposables = ensureNoDisposablesAreLeakedInTestSuite();
	function fixture(remoteAuthority?: string, advertisedResource = REACH_ACCOUNT_ORIGIN) {
		const accounts = disposables.add(new Emitter<void>());
		const starts = disposables.add(new Emitter<void>());
		const attempts = disposables.add(new Emitter<string>());
		const tokens: string[] = [];
		let token = 'saved-fixture';
		let reads = 0;
		let signInCalls = 0;
		const root: RootState = { agents: [{ provider: REACH_CHAT_PROVIDER_ID, displayName: 'REACH', description: '',
			models: [{ id: 'chat-fixture', name: 'Fixture', provider: REACH_CHAT_PROVIDER_ID }], protectedResources: [{ resource: advertisedResource }] }] };
		const account = { id: 'fixture', subscriptionActive: true };
		const reach = upcastPartial<IReachSubscriptionConnection>({
			onDidChangeAccount: accounts.event,
			signIn: async () => { signInCalls++; return token ? account : undefined; },
			getAccount: async () => token ? account : undefined,
			getLocalChatCredential: async () => {
				reads++;
				return token ? { accessToken: token, expiresAt: new Date(Date.now() + 60000).toISOString() } : undefined;
			},
		});
		let acknowledge = async () => ({ authenticated: true });
		const host = upcastPartial<IAgentHostService>({
			rootState: upcastPartial<IAgentHostService['rootState']>({ value: root, onDidChange: Event.None }),
			onAgentHostStart: starts.event,
			onDidNotification: Event.None,
			authenticate: async request => {
				assert.strictEqual(request.resource, REACH_ACCOUNT_ORIGIN);
				assert.strictEqual(request.target, undefined);
				assert.ok(request.token ? request.expiresIn && request.expiresIn > 0 : request.expiresIn === undefined);
				tokens.push(request.token);
				attempts.fire(request.token);
				return acknowledge();
			},
		});
		const connection = disposables.add(new LocalReachChatConnection(host, reach, upcastPartial<IWorkbenchEnvironmentService>({ remoteAuthority }), new TestNotificationService(), new NullLogService()));
		return { connection, tokens, starts, attempts, root, get reads() { return reads; }, get signInCalls() { return signInCalls; },
			setAccount: (value: string) => { token = value; accounts.fire(); },
			setAcknowledge: (handler: typeof acknowledge) => { acknowledge = handler; },
		};
	}

	test('opening a saved session waits for the local account handshake without opening sign-in', async () => {
		const f = fixture();
		const accepted = new DeferredPromise<{ authenticated: boolean }>();
		f.setAcknowledge(() => accepted.p);
		const attempted = Event.toPromise(f.attempts.event);
		let restored = false;
		const restore = f.connection.prepareHistory().then(() => { restored = true; });
		await attempted;
		assert.strictEqual(restored, false);
		await accepted.complete({ authenticated: true });
		await restore;
		f.setAccount('');
		await assert.rejects(f.connection.prepareHistory(), /Sign in with Connect REACH Chat/);
		assert.deepStrictEqual([restored, f.signInCalls, f.tokens], [true, 0, ['saved-fixture', '']]);
	});

	test('shares the saved credential, reconnects on host restart and revokes on sign-out', async () => {
		const f = fixture();
		assert.ok(await f.connection.connect());
		assert.ok(await f.connection.connect());
		f.starts.fire();
		assert.ok(await f.connection.connect());
		const revoked = Event.toPromise(f.attempts.event);
		f.setAccount('');
		await revoked;
		assert.strictEqual(await f.connection.connect(), false);
		assert.deepStrictEqual(f.tokens, ['saved-fixture', 'saved-fixture', '']);
	});

	test('leaves remote windows and other advertised resource URLs outside the account relay', async () => {
		for (const f of [fixture('ssh-remote+fixture'), fixture(undefined, 'https://example.test')]) {
			assert.strictEqual(await f.connection.connect(), false);
			assert.deepStrictEqual([f.reads, f.tokens], [0, []]);
		}
	});

	test('a sign-out during an in-flight handshake cannot complete the chat connection', async () => {
		const f = fixture();
		const first = new DeferredPromise<{ authenticated: boolean }>();
		f.setAcknowledge(() => first.p);
		const sent = Event.toPromise(f.attempts.event);
		const connected = f.connection.connect();
		await sent;
		const revoked = Event.toPromise(f.attempts.event);
		f.setAccount('');
		await first.complete({ authenticated: true });
		await revoked;
		assert.strictEqual(await connected, false);
		assert.deepStrictEqual(f.tokens, ['saved-fixture', '']);
	});
});
