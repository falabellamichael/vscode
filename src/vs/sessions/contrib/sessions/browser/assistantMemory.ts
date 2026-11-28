/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Signal-C. All rights reserved.
 *--------------------------------------------------------------------------------------------*/

/**
 * Persistent Assistant memory modeled on the Hermes two-target memory system:
 * a `user` profile (who the user is) and `memory` notes (environment, conventions,
 * quirks), each under a hard character budget, injected into every conversation,
 * and edited through atomic batches of add/patch/remove operations.
 */

export type AssistantMemoryTarget = 'user' | 'memory';

export interface AssistantMemoryOperation {
	readonly action: 'add' | 'patch' | 'remove';
	readonly content?: string;
	readonly old_text?: string;
	/** The memory picker identifies an entire entry, even if its text appears in another entry. */
	readonly exact?: true;
}

export interface AssistantMemoryBank {
	readonly user: string;
	readonly memory: string;
}

export const ASSISTANT_MEMORY_LIMITS: Readonly<Record<AssistantMemoryTarget, number>> = { user: 1_400, memory: 2_200 };

export const ASSISTANT_MEMORY_TOOL = 'assistant.memory';

/** Entries are non-empty lines; legacy bullet prefixes are normalized away. */
export function memoryEntries(text: string): string[] {
	return text.split('\n').map(line => line.trim().replace(/^-(?:\s+|$)/, '')).filter(line => line.length > 0);
}

export function memoryUsage(bank: AssistantMemoryBank): string {
	const part = (target: AssistantMemoryTarget) => {
		const used = bank[target].length;
		const limit = ASSISTANT_MEMORY_LIMITS[target];
		return `${target}: ${Math.round(used / limit * 100)}% — ${used}/${limit} chars`;
	};
	return `${part('user')} · ${part('memory')}`;
}

/** The Hermes-style context block injected into the system prompt. */
export function renderMemoryContext(bank: AssistantMemoryBank): string {
	const section = (target: AssistantMemoryTarget, title: string) => {
		const entries = memoryEntries(bank[target]);
		const used = bank[target].length;
		const limit = ASSISTANT_MEMORY_LIMITS[target];
		return `══ ${title} [${Math.round(used / limit * 100)}% — ${used}/${limit} chars] ══\n${entries.length ? entries.join('\n§\n') : '(empty)'}`;
	};
	return `${section('user', 'USER PROFILE (who the user is)')}\n\n${section('memory', 'MEMORY (your notes)')}`;
}

export type AssistantMemoryResult =
	| { readonly ok: true; readonly bank: AssistantMemoryBank; readonly usage: string }
	| { readonly ok: false; readonly error: string; readonly usage: string };

/**
 * Applies a batch atomically: any invalid operation or a final over-budget result
 * leaves the bank untouched, so one batch can free room AND add new entries.
 */
export function applyMemoryOperations(bank: AssistantMemoryBank, target: AssistantMemoryTarget, operations: readonly AssistantMemoryOperation[]): AssistantMemoryResult {
	const usage = memoryUsage(bank);
	const fail = (error: string): AssistantMemoryResult => ({ ok: false, error, usage });
	if (!operations.length) { return fail('No operations were provided.'); }
	const entries = memoryEntries(bank[target]);
	for (const operation of operations) {
		if (operation.action === 'add') {
			const additions = memoryEntries(operation.content ?? '');
			if (!additions.length) { return fail('An add operation needs non-empty content.'); }
			for (const line of additions) {
				if (!entries.includes(line)) { entries.push(line); }
			}
			continue;
		}
		if (operation.action !== 'patch' && operation.action !== 'remove') {
			return fail(`Unknown memory action "${String(operation.action)}". Use add, patch, or remove.`);
		}
		const oldText = operation.old_text ?? '';
		if (!oldText) { return fail(`A ${operation.action} operation needs old_text.`); }
		const matches = entries.filter(entry => operation.exact ? entry === oldText : entry.includes(oldText));
		if (matches.length !== 1) { return fail(`old_text must identify exactly one ${target} entry (found ${matches.length}).`); }
		const [entry] = matches;
		const index = entries.indexOf(entry);
		if (operation.action === 'remove') {
			entries.splice(index, 1);
		} else {
			if (entry.indexOf(oldText) !== entry.lastIndexOf(oldText)) { return fail(`old_text must occur exactly once in the matching ${target} entry.`); }
			const replaced = entry.replace(oldText, () => operation.content ?? '');
			entries.splice(index, 1, ...memoryEntries(replaced));
		}
	}
	const text = entries.join('\n');
	const limit = ASSISTANT_MEMORY_LIMITS[target];
	if (text.length > limit) {
		return fail(`Applying all operations would put ${target} memory at ${text.length}/${limit} chars. Remove or shorten enough stale entries in the SAME batch, then retry.\nCurrent entries:\n${memoryEntries(bank[target]).join('\n§\n')}`);
	}
	const next: AssistantMemoryBank = { ...bank, [target]: text };
	return { ok: true, bank: next, usage: memoryUsage(next) };
}

/** Guidance mirrored from the Hermes memory tool, condensed for the tool description. */
export const ASSISTANT_MEMORY_GUIDANCE = 'Save durable facts to persistent memory that survives across conversations. Memory is injected into every turn, so keep entries compact and high-signal. TARGETS: "user" = who the user is (name, role, preferences, style); "memory" = your notes (environment, conventions, tool quirks, lessons). HOW: apply ALL changes to one target in ONE batch via operations (add needs content; patch replaces old_text, which must occur exactly once; remove deletes the entry containing old_text). The batch is atomic and the character budget is checked on the FINAL result, so a single batch can remove stale entries AND add new ones. WHEN: only facts that apply to every conversation regardless of task. SKIP: trivial or easily re-discovered info, raw data dumps, task progress, and temporary TODO state. Procedures and multi-step workflows belong in routines, not memory. When a target is full, replace or consolidate stale entries instead of skipping the save. Write entries as declarative facts, not instructions to yourself.';
