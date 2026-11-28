/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import assert from 'assert';
import { DisposableStore } from '../../../../../base/common/lifecycle.js';
import { URI } from '../../../../../base/common/uri.js';
import { ensureNoDisposablesAreLeakedInTestSuite } from '../../../../../base/test/common/utils.js';
import { InMemoryStorageService, IStorageService, StorageScope, StorageTarget } from '../../../../../platform/storage/common/storage.js';
import { toPasteVariableEntry } from '../../../../../workbench/contrib/chat/common/attachments/chatVariableEntries.js';
import { ASSISTANT_SAVED_IDEA_TEXT_LIMIT, ASSISTANT_SAVED_IDEAS_LIMIT, AssistantConversationStore, searchAssistantSavedIdeas } from '../../browser/assistantConversation.js';
import { readSavedIdeaPage, readSavedIdeaQuery } from '../../browser/assistantSavedIdeaRecall.js';

suite('AssistantSavedIdeas', () => {
	const disposables = ensureNoDisposablesAreLeakedInTestSuite();
	let storage: IStorageService;
	let store: AssistantConversationStore;

	setup(() => {
		storage = disposables.add(new InMemoryStorageService());
		store = new AssistantConversationStore(storage);
	});

	test('merges independently opened continuation tasks after stale window writes and preserves ordinary tasks', () => {
		const otherStorage = disposables.add(new InMemoryStorageService());
		const second = new AssistantConversationStore(otherStorage);
		const one = store.saveIdea('a', 'First idea', 'Assistant');
		const two = second.saveIdea('a', 'Second idea', 'Assistant');
		assert.ok(one.status === 'saved' && two.status === 'saved');
		const firstTask = store.getOrCreateSavedIdeaTask('a', one.idea.id, 'First', 'Idea', { fileRoots: [] });
		const secondTask = second.getOrCreateSavedIdeaTask('a', two.idea.id, 'Second', 'Idea', { fileRoots: [] });
		assert.ok(firstTask && firstTask !== 'tooLong' && secondTask && secondTask !== 'tooLong');
		const ordinary = store.createTask('a');
		let updates = 0;
		const watch = disposables.add(new DisposableStore());
		store.onDidChange('a', watch)(() => updates++, undefined, watch);
		for (const key of otherStorage.keys(StorageScope.PROFILE, StorageTarget.MACHINE)) {
			storage.store(key, otherStorage.get(key, StorageScope.PROFILE), StorageScope.PROFILE, StorageTarget.MACHINE);
		}
		assert.ok(updates > 0);
		assert.deepStrictEqual(store.read('a').tasks?.map(task => task.id).sort(), [ordinary, firstTask.taskId, secondTask.taskId].sort());
		store.updateTask('a', secondTask.taskId, { title: 'Renamed' });
		assert.strictEqual(store.read('a').tasks?.find(task => task.id === secondTask.taskId)?.title, 'Renamed');
		assert.deepStrictEqual(store.read('foreign').tasks, undefined);
		assert.deepStrictEqual(storage.getObject<{ tasks: { id: string }[] }>('sessions.assistant.conversation.a', StorageScope.PROFILE)?.tasks.map(task => task.id), [ordinary]);
	});

	test('late initialization preserves ongoing edits and stale writes cannot restore a deleted continuation', () => {
		const otherStorage = disposables.add(new InMemoryStorageService());
		const stale = new AssistantConversationStore(otherStorage);
		const idea = store.saveIdea('a', 'A saved passage', 'Assistant');
		assert.ok(idea.status === 'saved');
		for (const key of storage.keys(StorageScope.PROFILE, StorageTarget.MACHINE)) {
			otherStorage.store(key, storage.get(key, StorageScope.PROFILE), StorageScope.PROFILE, StorageTarget.MACHINE);
		}
		const open = (owner: AssistantConversationStore) => {
			const result = owner.getOrCreateSavedIdeaTask('a', idea.idea.id, 'Idea', 'Saved idea', { fileRoots: [] });
			assert.ok(result && result !== 'tooLong');
			return result.taskId;
		};
		const task = open(store);
		const copyStaleWrites = () => {
			for (const key of otherStorage.keys(StorageScope.PROFILE, StorageTarget.MACHINE)) {
				if (key.startsWith('sessions.assistant.ideaTask') || key.startsWith('sessions.assistant.taskConversation')) {
					storage.store(key, otherStorage.get(key, StorageScope.PROFILE), StorageScope.PROFILE, StorageTarget.MACHINE);
				}
			}
		};
		store.setDraft('a', '', task);
		store.updateTask('a', task, { title: 'Revised' });
		assert.strictEqual(open(stale), task);
		copyStaleWrites();
		assert.deepStrictEqual(store.readComposerDraft('a', task).attachments, []);
		assert.strictEqual(store.read('a').tasks?.find(item => item.id === task)?.title, 'Revised');
		store.deleteTask('a', task);
		const fresh = open(store);
		assert.notStrictEqual(task, fresh);
		assert.strictEqual(stale.appendTurn('a', { prompt: 'Old request', reply: 'Late reply' }, task), true);
		stale.updateTask('a', task, { status: 'completed' });
		copyStaleWrites();
		assert.deepStrictEqual(store.read('a').tasks?.map(item => item.id), [fresh]);
		assert.deepStrictEqual(store.read('a', task).turns, []);
		assert.strictEqual(storage.get('sessions.assistant.taskConversation.a.' + task, StorageScope.PROFILE), undefined);
		assert.strictEqual(store.appendTurn('a', { prompt: 'Old request', reply: 'Late again' }, task), false);
		assert.deepStrictEqual(store.read('a', fresh).turns, []);
		assert.strictEqual(store.readComposerDraft('a', fresh).attachments.length, 1);
	});

	test('recognizes direct collection questions and separates a creative request from its search topic', () => {
		assert.strictEqual(readSavedIdeaQuery('Show me my saved ideas'), '');
		assert.strictEqual(readSavedIdeaQuery('What was my saved idea about the paper lantern?'), 'paper lantern');
		assert.strictEqual(readSavedIdeaQuery('What ideas have I bookmarked for the garden?'), 'garden');
		assert.strictEqual(readSavedIdeaQuery('Using my saved ideas about copper, suggest a small prototype.'), 'copper');
		assert.strictEqual(readSavedIdeaQuery('Can you use my saved ideas about paper to draft a workshop plan?'), 'paper');
		assert.strictEqual(readSavedIdeaQuery('Based on my bookmarked ideas for the trip, compare the walking tours.'), 'trip');
		assert.strictEqual(readSavedIdeaQuery('Please show me the ideas I saved about brass'), 'brass');
		assert.strictEqual(readSavedIdeaQuery('Rewrite these notes: What are my saved ideas?'), undefined);
		assert.strictEqual(readSavedIdeaQuery('How do I organize my saved ideas?'), undefined);
	});

	test('pages long matching passages losslessly and keeps late matches visible in the list', () => {
		const text = 'a'.repeat(1_999) + '💡\r\n**late needle**';
		const ideas = Array.from({ length: 6 }, (_, i) => ({ id: String(i), text, createdAt: i, sourceTitle: 'Workshop' }));
		const list = readSavedIdeaPage(ideas, 'needle', {});
		assert.match(list, /6 matching saved ideas, page 1 of 2/);
		assert.strictEqual(list.match(/late needle/g)?.length, 5);
		assert.strictEqual(readSavedIdeaPage(ideas, 'needle', { page: 2 }).match(/late needle/g)?.length, 1);
		const parts = [1, 2].map(page => {
			const result = readSavedIdeaPage(ideas, 'needle', { id: '0', page });
			return result.slice(result.indexOf('\n') + 1);
		});
		assert.ok(parts[1].startsWith('💡'));
		assert.strictEqual(parts.join(''), text);
		assert.match(readSavedIdeaPage(ideas, 'needle', { id: '0', page: 3 }), /exceeds/);
		assert.match(readSavedIdeaPage(ideas, 'needle', { page: 0 }), /positive integer/);
		assert.match(readSavedIdeaPage(ideas, 'needle', { id: 'absent' }), /No matching saved idea/);
	});

	test('preserves exact user-selected text after reopening, clearing a thread and deleting its source task', () => {
		const task = store.createTask('owner');
		const text = '  First line\r\n\tSecond line **as written**\n';
		store.appendTurn('owner', { prompt: 'Source prompt', reply: text }, task);
		const result = store.saveIdea('owner', text, 'Planning', task);
		assert.ok(result.status === 'saved');
		assert.strictEqual(result.idea.text, text);
		store.clearTurns('owner', task);
		store.deleteTask('owner', task);
		const reopened = new AssistantConversationStore(storage);
		assert.deepStrictEqual(reopened.listSavedIdeas('owner'), [result.idea]);
		assert.strictEqual(result.idea.taskId, task);
		assert.strictEqual(result.idea.sourceTitle, 'Planning');
		assert.deepStrictEqual(store.readMemoryBank('owner'), { user: '', memory: '' });
		assert.strictEqual(reopened.removeSavedIdea('owner', result.idea.id), true);
		assert.strictEqual(store.removeSavedIdea('owner', result.idea.id), false);
	});

	test('keeps account namespaces isolated and stores each idea under a separate machine-local profile key', () => {
		const first = store.saveIdea('a', 'First text', 'Assistant');
		const second = new AssistantConversationStore(storage).saveIdea('a', 'Second text', 'Assistant');
		const other = store.saveIdea('a.b/c', 'Private note', 'Assistant');
		assert.ok(first.status === 'saved' && second.status === 'saved' && other.status === 'saved');
		assert.strictEqual(store.listSavedIdeas('a').length, 2);
		assert.deepStrictEqual(store.listSavedIdeas('a.b/c'), [other.idea]);
		assert.deepStrictEqual(store.listSavedIdeas('a.b'), []);
		assert.strictEqual(store.removeSavedIdea('a', other.idea.id), false);
		assert.strictEqual(storage.keys(StorageScope.PROFILE, StorageTarget.MACHINE).length, 3);
		assert.deepStrictEqual(storage.keys(StorageScope.PROFILE, StorageTarget.USER), []);
		assert.throws(() => store.listSavedIdeas(''), /verified/);
		assert.throws(() => store.saveIdea('', 'note', 'Assistant'), /verified/);
		assert.throws(() => store.removeSavedIdea('', first.idea.id), /verified/);
	});

	test('stores identical text from distinct threads and recognizes an exact repeat in its original thread', () => {
		const first = store.saveIdea('owner', 'Keep this', 'Planning', 'one');
		const second = store.saveIdea('owner', 'Keep this', 'Planning', 'two');
		assert.ok(first.status === 'saved' && second.status === 'saved');
		assert.notStrictEqual(first.idea.id, second.idea.id);
		assert.deepStrictEqual(store.saveIdea('owner', 'Keep this', 'Planning', 'one'), { status: 'duplicate', idea: first.idea });
		assert.deepStrictEqual(store.saveIdea('owner', 'Keep this', 'Renamed task', 'one'), { status: 'duplicate', idea: first.idea });
		assert.strictEqual(store.listSavedIdeas('owner').length, 2);
	});

	test('simultaneous saves from two renderer caches address the same item for duplicates and independent keys for different ideas', () => {
		const secondWindow = new AssistantConversationStore(disposables.add(new InMemoryStorageService()));
		const first = store.saveIdea('owner', 'Meeting idea', 'Assistant');
		const duplicate = secondWindow.saveIdea('owner', 'Meeting idea', 'Assistant', '');
		const distinct = secondWindow.saveIdea('owner', 'Another idea', 'Assistant');
		assert.ok(first.status === 'saved' && duplicate.status === 'saved' && distinct.status === 'saved');
		assert.strictEqual(duplicate.idea.id, first.idea.id);
		assert.notStrictEqual(distinct.idea.id, first.idea.id);
	});

	test('rejects blank, oversized and over-capacity saves without evicting prior ideas or truncating text', () => {
		const longest = 'x'.repeat(ASSISTANT_SAVED_IDEA_TEXT_LIMIT);
		assert.deepStrictEqual(store.saveIdea('owner', ' \t\n', 'Assistant'), { status: 'empty' });
		assert.deepStrictEqual(store.saveIdea('owner', longest + '!', 'Assistant'), { status: 'tooLong' });
		const first = store.saveIdea('owner', longest, 'Assistant');
		assert.ok(first.status === 'saved');
		for (let i = 1; i < ASSISTANT_SAVED_IDEAS_LIMIT; i++) {
			assert.strictEqual(store.saveIdea('owner', 'Idea ' + i, 'Assistant').status, 'saved');
		}
		const before = store.listSavedIdeas('owner');
		assert.deepStrictEqual(store.saveIdea('owner', 'Extra idea', 'Assistant'), { status: 'full' });
		assert.deepStrictEqual(store.saveIdea('owner', longest, 'Assistant'), { status: 'duplicate', idea: first.idea });
		assert.deepStrictEqual(store.listSavedIdeas('owner'), before);
		assert.strictEqual(store.removeSavedIdea('owner', first.idea.id), true);
		assert.strictEqual(store.saveIdea('owner', 'Extra idea', 'Assistant').status, 'saved');
	});

	test('searches every term literally across text and source and sorts a copy newest first', () => {
		const ideas = [
			{ id: 'old', text: 'Bring [v2] map', sourceTitle: 'Trip notes', createdAt: 1 },
			{ id: 'new', text: 'TRIP [v2] map', sourceTitle: 'Today', createdAt: 3 },
			{ id: 'middle', text: 'map', sourceTitle: 'Other trip', createdAt: 2 },
		];
		assert.deepStrictEqual(searchAssistantSavedIdeas(ideas, '  [V2]\tTRIP  ').map(i => i.id), ['new', 'old']);
		assert.deepStrictEqual(searchAssistantSavedIdeas(ideas, ' ').map(i => i.id), ['new', 'middle', 'old']);
		assert.deepStrictEqual(ideas.map(i => i.id), ['old', 'new', 'middle']);
	});

	test('skips malformed records while preserving adjacent ideas', () => {
		const saved = store.saveIdea('account', 'Readable', 'Main');
		assert.ok(saved.status === 'saved');
		const prefix = 'sessions.assistant.savedIdeas.account/';
		storage.store(prefix + 'broken', '{', StorageScope.PROFILE, StorageTarget.MACHINE);
		storage.store(prefix + 'mismatch', { ...saved.idea, id: 'wrong' }, StorageScope.PROFILE, StorageTarget.MACHINE);
		for (const invalid of [
			{ ...saved.idea, id: 'impossible-date', createdAt: 1e20 },
			{ ...saved.idea, id: 'oversized-text', text: 'x'.repeat(ASSISTANT_SAVED_IDEA_TEXT_LIMIT + 1) },
			{ ...saved.idea, id: 'x'.repeat(1000) },
			{ ...saved.idea, id: 'contains/slash' },
		]) {
			storage.store(prefix + invalid.id, invalid, StorageScope.PROFILE, StorageTarget.MACHINE);
		}
		assert.deepStrictEqual(store.listSavedIdeas('account'), [saved.idea]);
	});

	test('continues an exact passage in a draft and restores its task, independently of the original chat', () => {
		const source = store.createTask('a');
		const text = '  First\r\n\r\n**second** 💡\t';
		const saved = store.saveIdea('a', text, 'Workshop notes', source);
		assert.ok(saved.status === 'saved');
		store.setModel('a', 'my-model');
		store.setComposerDraft('a', { inputText: 'Another draft', attachments: [toPasteVariableEntry('Other', 'Other material')] }, source);
		const originalDraft = store.readComposerDraft('a', source);
		const workspace = URI.file('/selected/project');
		const first = store.getOrCreateSavedIdeaTask('a', saved.idea.id, 'First second', 'Saved idea', { workspace });
		assert.ok(first && first !== 'tooLong');
		const task = store.read('a').tasks!.find(task => task.id === first.taskId)!;
		assert.deepStrictEqual([task.title, task.target?.workspace, task.status], ['First second', workspace.toString(), 'draft']);
		assert.deepStrictEqual(store.read('a', first.taskId).turns, []);
		assert.strictEqual(store.read('a', first.taskId).modelId, 'my-model');
		const draft = store.readComposerDraft('a', first.taskId);
		assert.strictEqual(draft.inputText, '');
		assert.deepStrictEqual(draft.attachments.map(item => [item.id, item.kind, item.value]), [[`assistant-saved:${saved.idea.id}`, 'paste', text]]);
		assert.deepStrictEqual(store.readComposerDraft('a', source), originalDraft);
		assert.deepStrictEqual(store.listSavedIdeas('a'), [saved.idea]);

		store.setComposerDraft('a', { inputText: 'Try this next', attachments: [] }, first.taskId);
		store.appendTurn('a', { prompt: 'First experiment', reply: 'Try one card' }, first.taskId);
		const before = store.read('a', first.taskId);
		const restarted = new AssistantConversationStore(storage);
		assert.deepStrictEqual(restarted.getOrCreateSavedIdeaTask('a', saved.idea.id, 'Replaced title', 'Other', {}), first);
		assert.deepStrictEqual(restarted.read('a', first.taskId), before);
		assert.deepStrictEqual(restarted.read('a').tasks!.find(item => item.id === first.taskId), task);
		assert.strictEqual(restarted.getOrCreateSavedIdeaTask('another-account', saved.idea.id, 'Other', 'Other', {}), undefined);
	});

	test('rejects an oversized initial context without creating a task and reopens existing work regardless of label length', () => {
		const saved = store.saveIdea('a', 'x'.repeat(ASSISTANT_SAVED_IDEA_TEXT_LIMIT - 70), 'Journal');
		assert.ok(saved.status === 'saved');
		const longLabel = 'label'.repeat(30);
		assert.strictEqual(store.getOrCreateSavedIdeaTask('a', saved.idea.id, 'Journal', longLabel, {}), 'tooLong');
		assert.deepStrictEqual(store.read('a').tasks ?? [], []);
		const short = store.getOrCreateSavedIdeaTask('a', saved.idea.id, 'Journal', 'Note', {});
		assert.ok(short && short !== 'tooLong');
		assert.deepStrictEqual(store.getOrCreateSavedIdeaTask('a', saved.idea.id, 'Journal', longLabel, {}), short);
	});

	test('deleting the task or removing and re-saving the idea begins fresh only on a later explicit continue', () => {
		const saved = store.saveIdea('a', 'Try a pocket map', 'Maps');
		assert.ok(saved.status === 'saved');
		const open = () => store.getOrCreateSavedIdeaTask('a', saved.idea.id, 'Try a pocket map', 'Map', {});
		const first = open();
		assert.ok(first && first !== 'tooLong');
		store.appendTurn('a', { prompt: 'First', reply: 'One' }, first.taskId);
		store.deleteTask('a', first.taskId);
		assert.deepStrictEqual(store.read('a').tasks ?? [], []);
		assert.deepStrictEqual(store.listSavedIdeas('a'), [saved.idea]);
		const second = open();
		assert.ok(second && second !== 'tooLong');
		assert.notStrictEqual(second.taskId, first.taskId);
		assert.deepStrictEqual(store.read('a', second.taskId).turns, []);
		assert.strictEqual(store.readComposerDraft('a', second.taskId).attachments[0].value, saved.idea.text);
		store.setDraft('a', 'Independent work', second.taskId);
		assert.ok(store.removeSavedIdea('a', saved.idea.id));
		assert.strictEqual(open(), undefined);
		const replacement = store.saveIdea('a', saved.idea.text, 'Maps');
		assert.ok(replacement.status === 'saved');
		assert.strictEqual(replacement.idea.id, saved.idea.id);
		assert.ok(!store.isCurrentSavedIdea('a', saved.idea));
		assert.ok(!store.isCurrentSavedIdea('a', { ...saved.idea, createdAt: replacement.idea.createdAt }));
		assert.ok(store.isCurrentSavedIdea('a', replacement.idea));
		const third = open();
		assert.ok(third && third !== 'tooLong');
		assert.notStrictEqual(third.taskId, second.taskId);
		assert.strictEqual(store.readComposerDraft('a', second.taskId).inputText, 'Independent work');
	});

	test('windows starting the same idea independently agree on one task ID, including legacy ideas', () => {
		const otherStorage = disposables.add(new InMemoryStorageService());
		const other = new AssistantConversationStore(otherStorage);
		const ids = [store, other].map(window => {
			const saved = window.saveIdea('a', 'Study the tide chart', 'Tides');
			assert.ok(saved.status === 'saved');
			return window.getOrCreateSavedIdeaTask('a', saved.idea.id, 'Tides', 'Tides', {});
		});
		assert.ok(ids[0] && ids[0] !== 'tooLong');
		assert.deepStrictEqual(ids[1], ids[0]);
		const otherOwner = store.saveIdea('b', 'Study the tide chart', 'Tides');
		assert.ok(otherOwner.status === 'saved');
		const otherTask = store.getOrCreateSavedIdeaTask('b', otherOwner.idea.id, 'Tides', 'Tides', {});
		assert.ok(otherTask && otherTask !== 'tooLong');
		assert.notStrictEqual(otherTask.taskId, ids[0].taskId);
		const key = 'sessions.assistant.savedIdeas.a/legacy';
		const legacy = { id: 'legacy', text: 'A handwritten note', sourceTitle: 'Journal', createdAt: 1 };
		storage.store(key, legacy, StorageScope.PROFILE, StorageTarget.MACHINE);
		otherStorage.store(key, { ...legacy, continuationId: '../other' }, StorageScope.PROFILE, StorageTarget.MACHINE);
		const a = store.getOrCreateSavedIdeaTask('a', 'legacy', 'Note', 'Note', {});
		assert.ok(a && a !== 'tooLong');
		assert.deepStrictEqual(other.getOrCreateSavedIdeaTask('a', 'legacy', 'Note', 'Note', {}), a);
	});

	test('deleting an old task in a stale window never rewrites a removed or re-saved passage', () => {
		const saved = store.saveIdea('a', 'Build a folding map', 'Journal');
		assert.ok(saved.status === 'saved');
		const task = store.getOrCreateSavedIdeaTask('a', saved.idea.id, 'Map', 'Idea', {});
		assert.ok(task && task !== 'tooLong');
		const staleStorage = disposables.add(new InMemoryStorageService());
		for (const key of storage.keys(StorageScope.PROFILE)) {
			staleStorage.store(key, storage.get(key, StorageScope.PROFILE), StorageScope.PROFILE, StorageTarget.MACHINE);
		}
		const key = 'sessions.assistant.savedIdeas.a/' + saved.idea.id;
		const changes = disposables.add(new DisposableStore());
		let staleSavedIdeaWrites = 0;
		staleStorage.onDidChangeValue(StorageScope.PROFILE, key, changes)(() => staleSavedIdeaWrites++, undefined, changes);
		store.removeSavedIdea('a', saved.idea.id);
		const replacement = store.saveIdea('a', saved.idea.text, 'Journal');
		assert.ok(replacement.status === 'saved');
		new AssistantConversationStore(staleStorage).deleteTask('a', task.taskId);
		assert.strictEqual(staleSavedIdeaWrites, 0);
		assert.deepStrictEqual(store.listSavedIdeas('a'), [replacement.idea]);
	});
});
