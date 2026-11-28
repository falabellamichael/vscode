/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import assert from 'assert';
import { isCancellationError } from '../../../../base/common/errors.js';
import { ensureNoDisposablesAreLeakedInTestSuite } from '../../../../base/test/common/utils.js';
import { ActionType, type ChatAction } from '../../common/state/sessionActions.js';
import { ToolCallContributorKind, ToolResultContentType, type ToolCallResult } from '../../common/state/protocol/state.js';
import { ReachClientTools } from '../../node/reach/reachClientTools.js';

suite('REACH client tools', () => {
	ensureNoDisposablesAreLeakedInTestSuite();
	const signal = new AbortController().signal;

	test('uses only the sending window, preserves its schema and accepts synchronous results', async () => {
		const result: ToolCallResult = { success: true, pastTenseMessage: 'Read the document', content: [{ type: ToolResultContentType.Text, text: 'contents' }] };
		const actions: ChatAction[] = [];
		const tools = new ReachClientTools(action => {
			actions.push(action);
			if (action.type === ActionType.ChatToolCallReady) { tools.complete(action.toolCallId, result); }
		});
		tools.getOrCreateClient('other', undefined).tools = [{ name: 'read' }, { name: 'search' }];
		const owner = tools.getOrCreateClient('editor', 'Editor');
		const definition = { name: 'read', title: 'Read document', inputSchema: { type: 'object' as const, properties: { filePath: { type: 'string' } }, required: ['filePath'] } };
		owner.tools = [definition];
		assert.deepStrictEqual(tools.definitions('editor'), [definition]);
		assert.deepStrictEqual(tools.definitions(), []);
		assert.strictEqual((await tools.invoke('turn', 'search', {}, signal, 'editor')).error?.code, 'toolUnavailable');
		assert.strictEqual(await tools.invoke('turn', 'read', { filePath: 'document.csv' }, signal, 'editor'), result);
		const [start, ready] = actions;
		assert(start.type === ActionType.ChatToolCallStart && ready.type === ActionType.ChatToolCallReady);
		assert.deepStrictEqual(start.contributor, { kind: ToolCallContributorKind.Client, clientId: 'editor' });
		assert.strictEqual(start.displayName, 'Read document');
		assert.strictEqual(ready.toolCallId, start.toolCallId);
		assert.strictEqual(ready.confirmed, undefined);
		assert.deepStrictEqual(JSON.parse(String(ready.toolInput)), { filePath: 'document.csv' });
	});

	test('observes enablement changes and ignores a disconnected handle when its client reconnects', async () => {
		const tools = new ReachClientTools(() => assert.fail('disabled tool ran'));
		const old = tools.getOrCreateClient('editor', undefined);
		old.tools = [{ name: 'fetch' }];
		old.tools = [];
		assert.strictEqual((await tools.invoke('turn', 'fetch', {}, signal)).success, false);
		tools.removeClient('editor');
		const current = tools.getOrCreateClient('editor', undefined);
		current.tools = [{ name: 'read' }];
		old.tools = [{ name: 'fetch' }];
		assert.deepStrictEqual(tools.definitions(), [{ name: 'read' }]);
	});

	test('a denied tool request settles even when the workbench sends no completion', async () => {
		const tools = new ReachClientTools(action => {
			if (action.type === ActionType.ChatToolCallReady) { tools.deny(action.toolCallId); }
		});
		tools.getOrCreateClient('editor', undefined).tools = [{ name: 'terminal' }];
		assert.strictEqual((await tools.invoke('turn', 'terminal', {}, signal)).error?.code, 'denied');
	});

	test('client loss and cancellation release pending work without replaying it', async () => {
		const actions: ChatAction[] = [];
		const tools = new ReachClientTools(action => actions.push(action));
		tools.getOrCreateClient('editor', undefined).tools = [{ name: 'write' }];
		const first = tools.invoke('first', 'write', {}, signal);
		tools.removeClient('editor');
		assert.strictEqual((await first).error?.code, 'clientDisconnected');
		tools.getOrCreateClient('editor', undefined).tools = [{ name: 'write' }];
		const controller = new AbortController();
		const interrupted = tools.invoke('second', 'write', {}, controller.signal);
		controller.abort();
		await assert.rejects(interrupted, isCancellationError);
		assert.strictEqual(actions.filter(action => action.type === ActionType.ChatToolCallComplete).length, 2);
		const finished = actions.at(-1)!;
		assert(finished.type === ActionType.ChatToolCallComplete);
		tools.complete(finished.toolCallId, { success: true, pastTenseMessage: 'late result' });
		assert.strictEqual(actions.length, 6);
	});
});
