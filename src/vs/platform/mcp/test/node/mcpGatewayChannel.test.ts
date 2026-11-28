/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import assert from 'assert';
import { Emitter, Event } from '../../../../base/common/event.js';
import { IPCServer, IChannel } from '../../../../base/parts/ipc/common/ipc.js';
import { upcastPartial } from '../../../../base/test/common/mock.js';
import { ensureNoDisposablesAreLeakedInTestSuite } from '../../../../base/test/common/utils.js';
import { ILoggerService } from '../../../log/common/log.js';
import { IMcpGatewayServerDescriptor, IMcpGatewayService, IMcpGatewayToolInvoker } from '../../common/mcpGateway.js';
import { McpGatewayChannel } from '../../node/mcpGatewayChannel.js';

suite('MCP gateway channel server selection', () => {
	const store = ensureNoDisposablesAreLeakedInTestSuite();
	const all = [{ id: 'workspace-a', label: 'A' }, { id: 'workspace-b', label: 'B' }];

	function fixture() {
		const changed = store.add(new Emitter<readonly IMcpGatewayServerDescriptor[]>());
		let invoker: IMcpGatewayToolInvoker | undefined;
		const broker = upcastPartial<IChannel>({ call: async <T>() => all as T, listen: event => event === 'onDidChangeServers' ? changed.event : Event.None });
		const ipc = upcastPartial<IPCServer<string>>({ onDidRemoveConnection: Event.None, getChannel: () => broker });
		const service = upcastPartial<IMcpGatewayService>({ createGateway: async (_client, value) => {
			invoker = value;
			return { gatewayId: 'test-gateway', servers: [], onDidChangeServers: Event.None };
		} });
		const channel = store.add(new McpGatewayChannel(ipc, service, upcastPartial<ILoggerService>({ getLogger: () => undefined })));
		return { channel, changed, get invoker() { assert.ok(invoker); return invoker; } };
	}

	test('applies the exact selection to initial and later descriptors', async () => {
		const f = fixture();
		await f.channel.call('editor', 'createGateway', { serverIds: ['workspace-a'], chatSessionResource: 'agent-host-devin-subscription:/chat' });
		assert.deepStrictEqual(f.invoker.listServers(), [all[0]]);
		const updates: (readonly IMcpGatewayServerDescriptor[])[] = [];
		store.add(f.invoker.onDidChangeServers(servers => updates.push(servers)));
		f.changed.fire([{ id: 'workspace-b', label: 'B' }, { id: 'workspace-a', label: 'Renamed A' }]);
		assert.deepStrictEqual(updates, [[{ id: 'workspace-a', label: 'Renamed A' }]]);
		assert.deepStrictEqual(f.invoker.listServers(), [{ id: 'workspace-a', label: 'Renamed A' }]);
	});

	test('an empty selection stays empty while omission preserves existing behavior', async () => {
		for (const serverIds of [[], undefined]) {
			const f = fixture();
			await f.channel.call('editor', 'createGateway', { serverIds });
			assert.deepStrictEqual(f.invoker.listServers(), serverIds ? [] : all);
		}
	});

	test('keeps explicitly shared connections off gateways created by other clients without a selection', async () => {
		const f = fixture();
		await f.channel.call('editor', 'createGateway');
		f.changed.fire([...all, { id: 'mail', label: 'Mail', requiresExplicitSelection: true }]);
		assert.deepStrictEqual(f.invoker.listServers(), all);
	});

	test('rejects malformed selections instead of exposing all servers', async () => {
		const f = fixture();
		for (const serverIds of [null, 'workspace-a', [42]]) {
			await assert.rejects(f.channel.call('editor', 'createGateway', { serverIds }), /Invalid MCP gateway server selection/);
		}
	});
});
