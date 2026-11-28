/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import type { ReachAssistantToolSpec } from '../../../native/common/native.js';
import type { ToolDefinition } from '../../common/state/sessionState.js';
import type { AttachedChatFolder } from './reachChatFolders.js';

export const findReachTools = 'reach_find_tools';
const searchSpec: ReachAssistantToolSpec = {
	name: findReachTools,
	description: 'Find enabled tools by name or purpose and make their functions available on the next step. Omit offset for a new query. To see more matches, repeat the same query with the returned nextOffset. An empty query lists all tools.',
	parameters: { type: 'object', properties: { query: { type: 'string' }, offset: { type: 'integer', minimum: 0 } }, required: ['query'] },
};

export interface ReachChatToolSelection {
	readonly includeTool?: (name: string) => boolean;
	readonly hasWorkspace?: boolean;
	readonly preferEditorForNewProject?: boolean;
	readonly attachedFolders?: readonly AttachedChatFolder[];
	readonly localRoot?: string;
	readonly editingAttachedFiles?: boolean;
	readonly previousTurn?: { readonly prompt: string; readonly reply: string };
	/** Native providers that unwrap text reads may budget the source instead of the serialized editor envelope. */
	readonly textReadBudgetBytes?: number;
}

const retainedWorkspaceFileTools = ['reach_create_file', 'reach_read_file', 'reach_edit_file'];
const workspaceFileTools = [...retainedWorkspaceFileTools, 'copilot_listDirectory', 'copilot_findFiles', 'copilot_findTextInFiles', 'copilot_getErrors', 'open_browser_page', 'read_page', 'click_element', 'type_in_page', 'run_playwright_code'];
const projectCreationTools = [...retainedWorkspaceFileTools, 'copilot_listDirectory', 'run_in_terminal', 'open_browser_page', 'run_playwright_code'];
const pageDiagnosticTools = ['list_browser_pages', 'open_browser_page', 'read_page', 'navigate_page', 'copilot_findFiles', 'reach_read_file', 'copilot_findTextInFiles', 'reach_edit_file', 'copilot_getErrors', 'copilot_listDirectory'];

export function isWebPageLoadIssue(prompt: string): boolean {
	return /\b(?:web\s?page|page|website|site|preview)\b[^.!?\n]{0,35}\b(?:(?:not|never|doesn['’]t|didn['’]t|won['’]t|can['’]t|fails? to|failed to)\s+(?:load|open|render)(?:ing)?|blank|404)\b|\bblank\s+(?:page|preview|website)\b/i.test(prompt) && !/\b(?:notebook|ipynb)\b/i.test(prompt);
}

export function isWorkspaceBuildRequest(prompt: string): boolean {
	return /\b(?:create|build|make|implement|write|develop|scaffold)\b[^.!?\n]{0,100}\b(?:version of|implementation of|clone of|app|application|game|website|webpage|web page|program|project|extension|plugin|cli)\b/i.test(prompt);
}
const workspaceFileRequest = /\b(?:this (?:folder|directory|workspace)|(?:html|css|javascript|typescript) files?|[\w.-]+\.(?:html?|css|[cm]?[jt]sx?|py|json))\b|\b(?:create|build|rebuild|make|edit|improve|update|fix)\b[^.!?\n]{0,80}\b(?:game|webpage|web page|script|source file|stylesheet)\b/i;

/** The gateway and smaller models accept bounded batches of function schemas. */
export class ReachChatToolCatalog {
	private readonly entries = new Map<string, ToolDefinition>();
	private readonly aliases = new Map<string, string>();
	private readonly reserved = new Set([findReachTools]);
	private definitions: readonly ToolDefinition[] | undefined;
	private selection: { query: string; offset: number; initial: boolean };
	private visible: string[] = [];
	private pageSize = 16;
	private keepWorkspaceFileTools = false;
	private readonly pinnedWorkspaceFileTools = new Set<string>();

	constructor(definitions: readonly ToolDefinition[], prompt: string, private readonly policy?: ReachChatToolSelection) {
		this.selection = { query: prompt, offset: 0, initial: true };
		this.refresh(definitions);
	}

	refresh(definitions: readonly ToolDefinition[]): void {
		if (this.definitions === definitions) { return; }
		this.definitions = definitions;
		this.entries.clear();
		for (const definition of definitions) {
			if (this.policy?.includeTool && !this.policy.includeTool(definition.name)) { continue; }
			let name = this.aliases.get(definition.name);
			if (!name) {
				const base = definition.name.replace(/[^a-zA-Z0-9_-]/g, '_').slice(0, 64) || 'tool';
				name = base;
				for (let i = 2; this.reserved.has(name); i++) {
					const suffix = '_' + i;
					name = base.slice(0, 64 - suffix.length) + suffix;
				}
				this.reserved.add(name);
				this.aliases.set(definition.name, name);
			}
			this.entries.set(name, definition);
		}
		this.selectPage(this.selection.query, this.selection.offset, this.selection.initial);
	}

	specs(wroteFile = false): ReachAssistantToolSpec[] {
		// For a new project the editor can inspect and create files directly. Keep the shell
		// discoverable for scaffolding, then advertise it alongside the editor for build and test.
		const useFileTools = this.policy?.preferEditorForNewProject && this.selection.initial && !wroteFile
			&& this.entries.has('reach_read_file') && !/\b(?:terminal|shell|powershell|bash|scaffold|npm|pnpm|yarn|cargo|gradle|test|install)\b/i.test(this.selection.query);
		const visible = useFileTools ? this.visible.filter(name => name !== 'run_in_terminal') : this.visible;
		const specs = visible.map(name => {
			const tool = this.entries.get(name)!;
			return { name, description: tool.description ?? tool.title ?? name, parameters: tool.inputSchema ?? { type: 'object' } };
		});
		return visible.length < this.entries.size ? [...specs, searchSpec] : specs;
	}

	reducePageSize(): boolean {
		if (this.visible.length <= 1) { return false; }
		this.pageSize = Math.ceil(this.visible.length / 2);
		this.visible = this.visible.slice(0, this.pageSize);
		return true;
	}

	originalName(name: string): string | undefined {
		return this.entries.get(name)?.name;
	}

	pinWorkspaceFileTool(name: string): void {
		if (retainedWorkspaceFileTools.includes(name) && this.entries.has(name)) {
			this.pinnedWorkspaceFileTools.add(name);
		}
	}

	selectPage(query: string, offset: number, initial = false): { tools: string[]; total: number; nextOffset?: number } {
		query = query.trim();
		if (!initial && (this.selection.initial || this.selection.query.toLowerCase() !== query.toLowerCase())) {
			offset = 0;
		}
		this.selection = { query, offset, initial };
		if (initial && isWebPageLoadIssue(query)) {
			const available = pageDiagnosticTools.filter(name => this.entries.has(name));
			if (available.length) {
				this.visible = available.slice(0, this.pageSize);
				return { tools: this.visible, total: this.entries.size };
			}
		}
		// Keep the editor's own file tools on the first page of a local coding task.
		// Explicit searches still page through the sender's entire enabled catalog.
		if (initial && this.policy?.hasWorkspace && !this.policy.editingAttachedFiles && isWorkspaceBuildRequest(query) && this.entries.has('reach_create_file')) {
			this.keepWorkspaceFileTools = true;
			this.visible = projectCreationTools.filter(name => this.entries.has(name)).slice(0, this.pageSize);
			return { tools: this.visible, total: this.entries.size };
		}
		if (initial && this.policy?.hasWorkspace && (this.entries.size > this.pageSize || this.policy.editingAttachedFiles)) {
			const previous = this.policy.previousTurn;
			const fileRequest = !!this.policy.attachedFolders?.length || workspaceFileRequest.test(query);
			if (this.policy.editingAttachedFiles || fileRequest || (previous && (workspaceFileRequest.test(previous.reply) || workspaceFileRequest.test(previous.prompt)))) {
				const local = workspaceFileTools.filter(name => this.entries.get(name)?.name === name && !(this.policy?.editingAttachedFiles && name === 'reach_create_file'));
				if (local.length) {
					this.keepWorkspaceFileTools = true;
					// A follow-up may omit the filename or switch to another tool. Keep tools the
					// user names directly alongside the file tools from the previous exchange.
					if (!fileRequest) {
						const words = new Set(query.toLowerCase().match(/[a-z0-9]{4,}/g));
						for (const name of this.entries.keys()) {
							const parts = name.replace(/([a-z])([A-Z])/g, '$1 $2').toLowerCase().split(/[^a-z0-9]+/);
							if (!workspaceFileTools.includes(name) && parts.some(part => words.has(part))) { local.push(name); }
						}
					}
					this.visible = local.slice(0, this.pageSize);
					return { tools: this.visible, total: this.entries.size };
				}
			}
		}
		const identifiers = new Set(query.toLowerCase().split(/[^a-z0-9_-]+/));
		const terms = [...new Set(query.toLowerCase().split(/[^a-z0-9]+/).filter(Boolean))];
		const ranked = [...this.entries].map(([name, tool]) => {
			const title = (name + ' ' + (tool.title ?? '')).toLowerCase();
			const description = tool.description?.toLowerCase() ?? '';
			const score = terms.reduce((sum, term) => sum + (title.includes(term) ? 4 : description.includes(term) ? 1 : 0), 0);
			return { name, score: score + (identifiers.has(name.toLowerCase()) ? 1 + 4 * terms.length : 0) };
		}).filter(entry => initial || !terms.length || entry.score > 0).sort((a, b) => b.score - a.score);
		// Keep a workspace file task's operations available while the model looks up a terminal or browser.
		// The search cursor counts only actual matches, including when the gateway reduces the page size.
		const baseRetained = !initial && this.keepWorkspaceFileTools
			? retainedWorkspaceFileTools.filter(name => this.entries.has(name) && !(this.policy?.editingAttachedFiles && name === 'reach_create_file')) : [];
		const retained = !initial
			? [...new Set([...baseRetained, ...this.pinnedWorkspaceFileTools])].filter(name => this.entries.has(name)).slice(0, this.pageSize - 1) : [];
		const matches = ranked.slice(offset, offset + this.pageSize - retained.length).map(entry => entry.name);
		this.visible = [...new Set([...matches, ...retained])];
		const next = offset + matches.length;
		return { tools: matches, total: ranked.length, ...(next < ranked.length ? { nextOffset: next } : {}) };
	}
}
