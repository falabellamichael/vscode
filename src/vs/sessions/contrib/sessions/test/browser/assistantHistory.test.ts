/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import assert from 'assert';
import sinon from 'sinon';
import { CancellationToken } from '../../../../../base/common/cancellation.js';
import { constObservable } from '../../../../../base/common/observable.js';
import { upcastPartial } from '../../../../../base/test/common/mock.js';
import { ensureNoDisposablesAreLeakedInTestSuite } from '../../../../../base/test/common/utils.js';
import { InMemoryStorageService } from '../../../../../platform/storage/common/storage.js';
import { ISignalCDirectMail } from '../../../../../workbench/contrib/mcp/browser/signalCDirectMail.js';
import { IMcpService } from '../../../../../workbench/contrib/mcp/common/mcpTypes.js';
import { AssistantConversationStore } from '../../browser/assistantConversation.js';
import { collectAssistantTools, IAssistantToolContext } from '../../browser/assistantTools.js';

suite('AssistantHistory', () => {
	const disposables = ensureNoDisposablesAreLeakedInTestSuite();
	let store: AssistantConversationStore;

	setup(() => {
		store = new AssistantConversationStore(disposables.add(new InMemoryStorageService()));
	});

	function catalog(overrides: Partial<IAssistantToolContext> = {}) {
		return collectAssistantTools(upcastPartial<IAssistantToolContext>({
			accountId: 'account', store, agentChatTarget: { fileRoots: [] },
			mcpService: upcastPartial<IMcpService>({ servers: constObservable([]) }), ...overrides,
		}));
	}

	async function search(query: string | Record<string, unknown>, taskId?: string) {
		const entry = catalog({ taskId }).entries.get('assistant.search_conversation');
		assert.ok(entry?.kind === 'builtin' && entry.safe);
		return entry.run(typeof query === 'string' ? { query } : query, CancellationToken.None);
	}

	test('recalls older exchanges from the current thread and account', async () => {
		store.appendTurn('account', { prompt: 'Set the launch date.', reply: 'We chose April 12.' });
		for (let i = 0; i < 13; i++) {
			store.appendTurn('account', { prompt: 'Next item.', reply: 'Reviewed.' });
		}
		store.appendTurn('account', { prompt: 'Launch date?', reply: 'Task-only reply.' }, 'task');
		store.appendTurn('other', { prompt: 'Launch date?', reply: 'Other account reply.' });
		assert.deepStrictEqual([await search(' LAUNCH   Date '), await search('launch date', 'task')], [
			'Past exchanges from this thread, newest first:\n\nTurn 1\nUser: Set the launch date.\nAssistant: We chose April 12.',
			'Past exchanges from this thread, newest first:\n\nTurn 1\nUser: Launch date?\nAssistant: Task-only reply.',
		]);
	});

	test('omits pasted text, file bodies, resource bodies and incomplete attachment blocks from matches and excerpts', async () => {
		store.appendTurn('account', {
			prompt: '<file path="notes.md">quoted </file> private file</file>\n<attachment name="paste">more quoted </file> private paste</attachment>\nReview the itinerary.',
			reply: 'Itinerary approved. <resource name="remote">private resource</resource>',
		});
		store.appendTurn('account', { prompt: '<file path="incomplete.md">private draft', reply: 'Acknowledged.' });
		assert.deepStrictEqual([await search('private'), await search('itinerary')], [
			'No matching exchange in this thread.',
			'Past exchanges from this thread, newest first:\n\nTurn 1\nUser: Review the itinerary.\nAssistant: Itinerary approved.',
		]);
	});

	test('keeps the five newest matches, with excerpts around late literal matches', async () => {
		for (let i = 1; i <= 7; i++) {
			store.appendTurn('account', { prompt: 'Status?', reply: 'x'.repeat(700) + ` milestone [v2] #${i} ` + 'y'.repeat(700) });
		}
		const result = await search('[v2]');
		assert.deepStrictEqual(result.match(/milestone \[v2\] #\d/g), [7, 6, 5, 4, 3].map(i => `milestone [v2] #${i}`));
		assert.ok(result.length < 4_000);
		const next = JSON.parse(result.split('Next page: ')[1]);
		assert.deepStrictEqual(next, { query: '[v2]', before_turn: 3 });
		const rest = await search(next);
		assert.deepStrictEqual(rest.match(/milestone \[v2\] #\d/g), ['milestone [v2] #2', 'milestone [v2] #1']);
		assert.doesNotMatch(rest, /Next page:/);
		assert.strictEqual(await search(' '), 'Enter a search phrase of 1 to 120 characters, or provide a date range.');
	});

	test('recalls typed questions about closing tags without including the attached instructions', async () => {
		const attachment = '<attachment name="Skill · example">\nInternal fixture guidance\n</attachment>\n\n';
		const question = 'Why does the transcript contain </attachment>?';
		store.appendTurn('account', { prompt: attachment + question, requestStart: attachment.length, reply: 'It marks the end.' });
		assert.match(await search('transcript'), /Why does the transcript contain <\/attachment>\?/);
		assert.strictEqual(await search('Internal fixture'), 'No matching exchange in this thread.');
	});

	test('recalls yesterday across local midnight with inclusive bounds and skips missing or invalid timestamps', async () => {
		const clock = sinon.useFakeTimers({ now: new Date(2024, 2, 11, 0, 15), toFake: ['Date'] });
		try {
			const add = (label: string, at?: number) => store.appendTurn('account', { prompt: 'Decision', reply: label, at });
			add('Previous day', new Date(2024, 2, 9, 23, 59, 59).getTime());
			add('Morning review', new Date(2024, 2, 10).getTime());
			add('Evening review', new Date(2024, 2, 10, 23, 59, 59, 999).getTime());
			add('New day', new Date(2024, 2, 11).getTime());
			add('Undated legacy exchange');
			add('Invalid timestamp', 1e20);
			assert.deepStrictEqual([
				await search({ from_date: 'yesterday', to_date: 'yesterday' }),
				await search({ query: 'morning', to_date: '2024-03-10' }),
				await search({ from_date: 'today' }),
			], [
				'Past exchanges from this thread, newest first:\n\nTurn 3 (2024-03-10)\nUser: Decision\nAssistant: Evening review\n\nTurn 2 (2024-03-10)\nUser: Decision\nAssistant: Morning review',
				'Past exchanges from this thread, newest first:\n\nTurn 2 (2024-03-10)\nUser: Decision\nAssistant: Morning review',
				'Past exchanges from this thread, newest first:\n\nTurn 4 (2024-03-11)\nUser: Decision\nAssistant: New day',
			]);
			assert.match(await search('Undated legacy'), /Undated legacy exchange/);
		} finally {
			clock.restore();
		}
	});

	test('rejects incorrect dates without widening the search and identifies partial date coverage', async () => {
		for (let i = 1; i <= 7; i++) {
			store.appendTurn('account', { prompt: `Decision ${i}`, reply: 'Accepted', at: new Date(2024, 1, 29, 12).getTime() });
		}
		const result = await search({ from_date: '2024-02-29', to_date: '2024-02-29' });
		assert.deepStrictEqual(result.match(/Decision \d/g), [7, 6, 5, 4, 3].map(i => `Decision ${i}`));
		assert.match(result, /Earlier matches are omitted/);
		for (const bound of ['2024-02-30', 'next week', '2024-02-29T00:00Z', '', 0]) {
			assert.match(await search({ query: 'Decision', from_date: bound }), /Use valid calendar dates/);
		}
		assert.match(await search({ from_date: '2024-03-01', to_date: '2024-02-29' }), /from_date must be on or before to_date/);
	});

	test('continues a filtered date search without repeating turns or shifting the day at midnight', async () => {
		const clock = sinon.useFakeTimers({ now: new Date(2024, 2, 11, 23, 59), toFake: ['Date'] });
		try {
			for (let i = 1; i <= 10; i++) {
				store.appendTurn('account', { prompt: i % 3 === 0 ? 'Other topic' : 'Decision', reply: String(i), at: new Date(2024, 2, i === 1 ? 9 : 10, 12).getTime() });
			}
			const first = await search({ query: 'Decision', from_date: 'yesterday', to_date: 'yesterday' });
			const next = JSON.parse(first.split('Next page: ')[1]);
			assert.deepStrictEqual(next, { query: 'decision', from_date: '2024-03-10', to_date: '2024-03-10', before_turn: 4 });
			clock.setSystemTime(new Date(2024, 2, 12));
			store.appendTurn('account', { prompt: 'Decision', reply: 'newer exchange' });
			const rest = await search(next);
			assert.deepStrictEqual([first.match(/^Turn \d+/gm), rest.match(/^Turn \d+/gm)], [['Turn 10', 'Turn 8', 'Turn 7', 'Turn 5', 'Turn 4'], ['Turn 2']]);
			assert.doesNotMatch(rest, /Next page:/);
			assert.strictEqual(await search({ ...next, before_turn: 1 }), 'No matching exchange in this thread.');
			for (const invalid of [0, -1, 1.5, '4', null]) {
				assert.strictEqual(await search({ ...next, before_turn: invalid }), 'before_turn must be a positive integer.');
			}
		} finally {
			clock.restore();
		}
	});

	test('remains outside connected-history and repository-only runs and leaves room for connected tools', () => {
		const readOnlyModes: Partial<IAssistantToolContext>[] = [{ connectedHistory: true }, { repositoryScope: 'workspace' }, { repositoryScope: 'github' }];
		for (const mode of readOnlyModes) {
			assert.strictEqual(catalog(mode).entries.has('assistant.search_conversation'), false);
		}
		const tools = catalog({
			agentChats: upcastPartial({}), mailWatcher: upcastPartial({}), configuration: upcastPartial({}),
			readNewsFeeds: async () => [], readWeatherForecast: async () => { throw new Error('unused'); },
			connectApp: () => true, proposeRoutine: () => '',
			directMail: upcastPartial<ISignalCDirectMail>({ connections: constObservable([
				{ id: 'gmail', provider: 'gmail', label: 'Gmail', needsSignIn: false },
				{ id: 'outlook', provider: 'outlook', label: 'Outlook', needsSignIn: false },
			]) }),
		}).specs;
		assert.ok(tools.length <= 32);
		assert.deepStrictEqual(tools.filter(tool => tool.name.endsWith('_read')).map(tool => tool.name), ['assistant.gmail_read', 'assistant.outlook_read']);
	});
});
