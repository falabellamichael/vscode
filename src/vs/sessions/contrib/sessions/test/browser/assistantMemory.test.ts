/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Signal-C. All rights reserved.
 *--------------------------------------------------------------------------------------------*/

import assert from 'assert';
import { ensureNoDisposablesAreLeakedInTestSuite } from '../../../../../base/test/common/utils.js';
import { InMemoryStorageService } from '../../../../../platform/storage/common/storage.js';
import { AssistantConversationStore } from '../../browser/assistantConversation.js';
import { applyMemoryOperations, ASSISTANT_MEMORY_LIMITS, memoryEntries, memoryUsage, renderMemoryContext } from '../../browser/assistantMemory.js';

suite('AssistantMemory', () => {
	const store = ensureNoDisposablesAreLeakedInTestSuite();

	test('applies add, patch, and remove as one atomic batch', () => {
		const bank = { user: 'Name: Michael\nPrefers short replies', memory: '' };
		const result = applyMemoryOperations(bank, 'user', [
			{ action: 'patch', old_text: 'Name: Michael', content: 'Name: Michael Anthony Falabella' },
			{ action: 'add', content: 'Builds AI models on AMD hardware' },
			{ action: 'remove', old_text: 'Prefers short replies' },
		]);
		assert.ok(result.ok);
		if (result.ok) {
			assert.deepStrictEqual(memoryEntries(result.bank.user), ['Name: Michael Anthony Falabella', 'Builds AI models on AMD hardware']);
			assert.strictEqual(result.bank.memory, '');
			assert.match(result.usage, /user: \d+%/);
		}
		assert.strictEqual(bank.user, 'Name: Michael\nPrefers short replies');
	});

	test('rejects the whole batch when any operation is invalid', () => {
		const bank = { user: 'Fact one', memory: '' };
		const result = applyMemoryOperations(bank, 'user', [
			{ action: 'add', content: 'Fact two' },
			{ action: 'patch', old_text: 'missing text', content: 'x' },
		]);
		assert.ok(!result.ok);
		assert.strictEqual(bank.user, 'Fact one');
	});

	test('enforces the budget on the final result so one batch can free room and add', () => {
		const filler = 'x'.repeat(ASSISTANT_MEMORY_LIMITS.user - 20);
		const bank = { user: filler, memory: '' };
		const overflow = applyMemoryOperations(bank, 'user', [{ action: 'add', content: 'y'.repeat(100) }]);
		assert.ok(!overflow.ok);
		if (!overflow.ok) {
			assert.match(overflow.error, /SAME batch/);
			assert.strictEqual(overflow.error.split('Current entries:\n')[1], filler);
		}
		const consolidated = applyMemoryOperations(bank, 'user', [
			{ action: 'remove', old_text: filler },
			{ action: 'add', content: 'Compact fact' },
		]);
		assert.ok(consolidated.ok);
	});

	test('requires old_text to identify exactly one entry', () => {
		const bank = { user: 'alpha\nbeta\nalpha gamma', memory: '' };
		assert.ok(!applyMemoryOperations(bank, 'user', [{ action: 'remove', old_text: 'alpha' }]).ok);
		assert.ok(!applyMemoryOperations(bank, 'user', [{ action: 'remove', old_text: 'absent' }]).ok);
		assert.ok(applyMemoryOperations(bank, 'user', [{ action: 'remove', old_text: 'beta' }]).ok);
	});

	test('patch with empty content removes the entry', () => {
		const result = applyMemoryOperations({ user: 'Keep\nDrop me', memory: '' }, 'user', [{ action: 'patch', old_text: 'Drop me', content: '' }]);
		assert.ok(result.ok);
		if (result.ok) { assert.strictEqual(result.bank.user, 'Keep'); }
	});

	test('uses literal replacement text and makes each patched line independently editable', () => {
		const result = applyMemoryOperations({ user: 'Draft template\nOther note', memory: '' }, 'user', [
			{ action: 'patch', old_text: 'Draft template', content: '- Use $& and $$ literally\n- Review later' },
			{ action: 'remove', old_text: 'Review later' },
			{ action: 'add', content: '- Other note\n-10°C is comfortable\n--offline' },
		]);
		assert.ok(result.ok);
		assert.deepStrictEqual(memoryEntries(result.bank.user), ['Use $& and $$ literally', 'Other note', '-10°C is comfortable', '--offline']);
	});

	test('distinguishes ambiguous replacements from a unique entry to forget', () => {
		const bank = { user: 'Echo ooo\nKeep', memory: '' };
		assert.strictEqual(applyMemoryOperations(bank, 'user', [{ action: 'patch', old_text: 'oo', content: 'x' }]).ok, false);
		const result = applyMemoryOperations(bank, 'user', [{ action: 'remove', old_text: 'o' }]);
		assert.ok(result.ok);
		assert.strictEqual(result.bank.user, 'Keep');
	});

	test('normalizes legacy bullets and renders the Hermes-style context', () => {
		const bank = { user: '- Legacy fact one\n- Legacy fact two', memory: 'Env note' };
		assert.deepStrictEqual(memoryEntries(bank.user), ['Legacy fact one', 'Legacy fact two']);
		const context = renderMemoryContext(bank);
		assert.match(context, /USER PROFILE/);
		assert.match(context, /MEMORY \(your notes\)/);
		assert.match(context, /Legacy fact one\n§\nLegacy fact two/);
		assert.match(context, /chars\]/);
		assert.match(memoryUsage(bank), /user: \d+% — \d+\/1400 chars · memory: \d+% — \d+\/2200 chars/);
	});

	test('edits the exact selected memory and keeps later edits when forgetting from an open picker', () => {
		const conversations = new AssistantConversationStore(store.add(new InMemoryStorageService()));
		conversations.writeMemoryBank('account', { user: 'Tea\nTea with milk', memory: 'Old note' });
		const edit = conversations.applyMemory('account', 'user', [{ action: 'patch', old_text: 'Tea', content: 'Coffee', exact: true }]);
		assert.ok(edit.ok);
		conversations.applyMemory('account', 'memory', [{ action: 'patch', old_text: 'Old note', content: 'Updated note' }]);
		conversations.applyMemory('account', 'user', [{ action: 'add', content: 'Uses metric' }]);
		assert.strictEqual(conversations.forgetMemoryEntries('account', { user: ['Tea with milk'], memory: ['Old note'] }), 1);
		assert.deepStrictEqual(conversations.readMemoryBank('account'), { user: 'Coffee\nUses metric', memory: 'Updated note' });
	});

	test('store keeps reviewed memories in the legacy bullet format and dedupes', () => {
		const storage = store.add(new InMemoryStorageService());
		const conversations = new AssistantConversationStore(storage);
		assert.deepStrictEqual(conversations.remember('account', ['Prefers short updates']), { added: 1, total: 1 });
		assert.deepStrictEqual(conversations.remember('account', ['Prefers short updates', 'Second fact']), { added: 1, total: 2 });
		assert.strictEqual(conversations.read('account').instructions, '- Prefers short updates\n- Second fact');
	});

	test('store applies memory batches atomically and persists both targets', () => {
		const storage = store.add(new InMemoryStorageService());
		const conversations = new AssistantConversationStore(storage);
		conversations.setInstructions('account', '- Legacy user fact');
		const failed = conversations.applyMemory('account', 'memory', [{ action: 'remove', old_text: 'absent' }]);
		assert.ok(!failed.ok);
		assert.strictEqual(conversations.read('account').agentNotes, undefined);
		const ok = conversations.applyMemory('account', 'memory', [{ action: 'add', content: 'Tool quirk noted' }]);
		assert.ok(ok.ok);
		const reopened = new AssistantConversationStore(storage);
		assert.deepStrictEqual(reopened.readMemoryBank('account'), { user: '- Legacy user fact', memory: 'Tool quirk noted' });
	});

	test('reviewed memories drop oldest entries when the user budget is exceeded', () => {
		const storage = store.add(new InMemoryStorageService());
		const conversations = new AssistantConversationStore(storage);
		conversations.setInstructions('account', '- ' + 'a'.repeat(ASSISTANT_MEMORY_LIMITS.user - 10));
		const { added } = conversations.remember('account', ['Fresh reviewed fact']);
		assert.strictEqual(added, 1);
		assert.deepStrictEqual(memoryEntries(conversations.read('account').instructions), ['Fresh reviewed fact']);
	});

	test('oversized reviewed facts preserve existing entries and only retained additions count', () => {
		const storage = store.add(new InMemoryStorageService());
		const conversations = new AssistantConversationStore(storage);
		const current = 'Original preference';
		const exactFit = 'a'.repeat(ASSISTANT_MEMORY_LIMITS.user - 2);
		const oversized = exactFit + 'x';
		conversations.setInstructions('account', current);
		assert.deepStrictEqual(conversations.remember('account', [oversized]), { added: 0, total: 1 });
		assert.strictEqual(conversations.readMemoryBank('account').user, current);
		assert.deepStrictEqual(conversations.remember('account', ['First suggestion', exactFit, oversized]), { added: 1, total: 1 });
		assert.strictEqual(new AssistantConversationStore(storage).readMemoryBank('account').user, `- ${exactFit}`);
	});

	test('counts the stored newline and both bullet prefixes when retaining reviewed notes', () => {
		const conversations = new AssistantConversationStore(store.add(new InMemoryStorageService()));
		const first = 'x'.repeat(ASSISTANT_MEMORY_LIMITS.user - 5);
		conversations.setInstructions('account', `- ${first}`);
		assert.deepStrictEqual(conversations.remember('account', ['Y']), { added: 1, total: 1 });
		assert.strictEqual(conversations.readMemoryBank('account').user, '- Y');
	});
});
