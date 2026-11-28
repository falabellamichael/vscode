/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { Codicon } from '../../../../base/common/codicons.js';
import { Event } from '../../../../base/common/event.js';
import { StringSHA1 } from '../../../../base/common/hash.js';
import { DisposableStore } from '../../../../base/common/lifecycle.js';
import { parse, stringify } from '../../../../base/common/marshalling.js';
import { Schemas } from '../../../../base/common/network.js';
import { equals } from '../../../../base/common/objects.js';
import { extUriBiasedIgnorePathCase } from '../../../../base/common/resources.js';
import { URI } from '../../../../base/common/uri.js';
import { generateUuid, isUUID } from '../../../../base/common/uuid.js';
import { REACH_ASSISTANT_PROMPT_LIMIT } from '../../../../platform/native/common/native.js';
import { IStorageService, StorageScope, StorageTarget } from '../../../../platform/storage/common/storage.js';
import { IChatDraft } from '../../../../workbench/contrib/chat/common/attachments/chatDraft.js';
import { toPasteVariableEntry } from '../../../../workbench/contrib/chat/common/attachments/chatVariableEntries.js';
import { fromStoredChatDraft, toStoredChatDraft } from '../../../../workbench/contrib/sessions/common/newChatDraftState.js';
import type { IAssistantAgentTarget } from './assistantAgentChats.js';
import type { AssistantRoutineDraft } from './assistantBriefingScheduler.js';
import { applyMemoryOperations, AssistantMemoryBank, AssistantMemoryOperation, AssistantMemoryResult, AssistantMemoryTarget, ASSISTANT_MEMORY_LIMITS, memoryEntries } from './assistantMemory.js';
import type { IAssistantWorktree } from './assistantWorktree.js';

export type AssistantPermissionMode = 'routine' | 'safeOnly' | 'allowAll';

export interface IAssistantActivity {
	readonly tool: string;
	readonly label: string;
	readonly status: 'pending' | 'awaitingApproval' | 'running' | 'done' | 'rejected' | 'failed';
	readonly detail?: string;
	readonly documentId?: string;
	readonly durationMs?: number;
}

export interface IAssistantMailDelivery {
	readonly accepted: number;
	readonly unconfirmed: boolean;
}

export interface IAssistantPlanItem {
	readonly step: string;
	readonly status: 'pending' | 'inProgress' | 'done';
}

interface IAssistantWorktreeSelection {
	readonly project: URI;
	readonly worktree: IAssistantWorktree;
}

function readUri(value: unknown): URI | undefined {
	if (typeof value !== 'string') { return undefined; }
	try {
		return URI.parse(value, true);
	} catch {
		return undefined;
	}
}

function readLocalPath(value: unknown): URI | undefined {
	const uri = readUri(value);
	return uri?.scheme === Schemas.file && uri.path.startsWith('/') && !uri.query && !uri.fragment ? uri : undefined;
}

interface IStoredWorktree {
	readonly uri?: unknown;
	readonly label?: unknown;
	readonly identity?: { readonly gitDirectory?: unknown; readonly createdAt?: unknown };
}

interface IStoredTaskTarget {
	readonly workspace?: string;
	readonly worktree?: IStoredWorktree;
	readonly fileRoots?: readonly string[];
	readonly fileRootNames?: readonly string[];
	readonly activeEditorFile?: string;
	readonly providerId?: string;
	readonly sessionTypeId?: string;
}

function storeWorktree(worktree: IAssistantWorktree): IStoredWorktree {
	return {
		uri: worktree.uri.toString(), label: worktree.label,
		identity: worktree.identity && { gitDirectory: worktree.identity.gitDirectory.toString(), createdAt: worktree.identity.createdAt },
	};
}

function readWorktree(raw: IStoredWorktree | undefined): IAssistantWorktree | undefined {
	const uri = readLocalPath(raw?.uri);
	if (!uri || typeof raw?.label !== 'string' || !raw.label) { return undefined; }
	const gitDirectory = readLocalPath(raw.identity?.gitDirectory);
	const createdAt = raw.identity?.createdAt;
	const identity = gitDirectory && typeof createdAt === 'number' && Number.isFinite(createdAt) ? { gitDirectory, createdAt } : undefined;
	return { uri, label: raw.label, identity };
}

export function storeAssistantTaskTarget(target: IAssistantAgentTarget): IStoredTaskTarget {
	return {
		workspace: target.workspace?.toString(), worktree: target.worktree && storeWorktree(target.worktree),
		fileRoots: target.fileRoots?.map(root => root.toString()), fileRootNames: target.fileRootNames, activeEditorFile: target.activeEditorFile?.toString(),
		providerId: target.providerId, sessionTypeId: target.sessionTypeId,
	};
}

export function readAssistantTaskTarget(task: IAssistantTask | undefined): IAssistantAgentTarget | undefined {
	return readAssistantTarget(task?.target);
}

export function readAssistantTarget(raw: IStoredTaskTarget | undefined): IAssistantAgentTarget | undefined {
	if (!raw || typeof raw !== 'object') { return undefined; }
	const fileRoots = Array.isArray(raw.fileRoots) ? raw.fileRoots.map(readUri).filter(uri => uri !== undefined) : [];
	const names = raw.fileRootNames;
	const fileRootNames = Array.isArray(names) && names.length === raw.fileRoots?.length && names.length === fileRoots.length
		&& names.every(name => typeof name === 'string') ? names : undefined;
	const activeEditorFile = readLocalPath(raw.activeEditorFile);
	return {
		workspace: readUri(raw.workspace), worktree: readWorktree(raw.worktree),
		fileRoots, ...(fileRootNames ? { fileRootNames } : {}),
		...(activeEditorFile ? { activeEditorFile } : {}),
		providerId: typeof raw.providerId === 'string' ? raw.providerId : undefined,
		sessionTypeId: typeof raw.sessionTypeId === 'string' ? raw.sessionTypeId : undefined,
	};
}

interface IAssistantQueuedPrompt {
	readonly id?: string;
	readonly prompt: string;
	readonly requestStart?: number;
	readonly imageUrls?: readonly string[];
	readonly serverIds?: readonly string[];
	readonly target: IStoredTaskTarget;
}

function isQueuedPrompt(item: IAssistantQueuedPrompt | undefined): item is IAssistantQueuedPrompt {
	return typeof item?.prompt === 'string' && !!item.prompt && !!item.target && typeof item.target === 'object';
}

export interface IAssistantTurn {
	readonly prompt: string;
	/** Starts the visible and model-facing history while retaining earlier turns for recall. */
	readonly compacted?: true;
	/** Offset of the user's typed request after the attached context, when recorded by the composer. */
	readonly requestStart?: number;
	/** Vision inputs sent with this prompt, as data or remote image URLs. */
	readonly imageUrls?: readonly string[];
	readonly reply: string;
	readonly routineId?: string;
	readonly routineRunId?: string;
	/** Legacy single-draft turns; newer turns keep every draft in routineDrafts. */
	readonly routineDraft?: AssistantRoutineDraft;
	readonly routineDrafts?: readonly AssistantRoutineDraft[];
	readonly unavailableConnections?: readonly { readonly id: string; readonly name: string; readonly needsSignIn?: true }[];
	readonly model?: string;
	readonly at?: number;
	readonly durationMs?: number;
	readonly activity?: readonly IAssistantActivity[];
	/** Provider outcomes carried across regenerations and send confirmations. */
	readonly mailDelivery?: IAssistantMailDelivery;
	readonly plan?: readonly IAssistantPlanItem[];
}

export function assistantSummaryIndex(turns: readonly IAssistantTurn[]): number {
	return turns.findLastIndex(turn => turn.compacted === true);
}

/** Regeneration may replace this branch only while the saved transcript still matches. */
export interface IAssistantTurnReplacement {
	readonly index: number;
	readonly original: readonly IAssistantTurn[];
}

export interface IAssistantConversation {
	readonly instructions: string;
	/** Hermes-style agent notes target; `instructions` remains the user-profile target. */
	readonly agentNotes?: string;
	/** Recent important-mail alerts reported by the passive watcher, newest first. */
	readonly mailAlerts?: readonly IAssistantMailAlert[];
	readonly draft?: string;
	readonly draftContext?: string;
	readonly queuedPrompts?: readonly IAssistantQueuedPrompt[];
	readonly startedQueuedPrompt?: IAssistantQueuedPrompt;
	readonly turns: readonly IAssistantTurn[];
	readonly modelId?: string;
	readonly tasks?: readonly IAssistantTask[];
	readonly agentChats?: readonly IAssistantAgentChat[];
	readonly permissionMode?: AssistantPermissionMode;
	readonly allowedTools?: readonly string[];
}

interface IAssistantAgentChat {
	readonly id: string;
	readonly resource: string;
	readonly title: string;
}

export interface IAssistantTask {
	readonly id: string;
	readonly title: string;
	readonly status: 'draft' | 'running' | 'completed' | 'stopped' | 'failed';
	readonly update?: string;
	readonly pinned?: boolean;
	readonly savedIdeaId?: string;
	/** The prompt the task last ran with; used to resume after an interruption. */
	readonly lastPrompt?: string;
	readonly target?: IStoredTaskTarget;
}

export interface IAssistantMailAlert {
	readonly id: string;
	readonly connectionId: string;
	readonly label: string;
	readonly subject: string;
	readonly from?: string;
	readonly received?: string;
	readonly url?: string;
	readonly at: number;
}

export interface IAssistantSavedIdea {
	readonly id: string;
	readonly instanceId?: string;
	readonly text: string;
	readonly sourceTitle: string;
	readonly taskId?: string;
	readonly createdAt: number;
}

interface IStoredSavedIdea extends IAssistantSavedIdea {
	/** Captured at save time, distinguishing a later save of a previously removed passage. */
	readonly continuationId?: string;
}

function savedIdeaTaskId(accountId: string, idea: IStoredSavedIdea): string {
	const generation = typeof idea.continuationId === 'string' && isUUID(idea.continuationId) ? idea.continuationId : undefined;
	const hash = new StringSHA1();
	hash.update(JSON.stringify([accountId, idea.id, generation]));
	return `idea-${hash.digest()}`;
}

export function assistantTextAttachment(name: string, text: string): string {
	return `<attachment name="${name.replace(/"/g, '')}">\n${text}\n</attachment>\n\n`;
}

export const ASSISTANT_SAVED_IDEAS_LIMIT = 100;
export const ASSISTANT_SAVED_IDEA_TEXT_LIMIT = 8_000;

export interface IAssistantDocument {
	readonly id: string;
	readonly title: string;
	readonly content: string;
	readonly createdAt: number;
	readonly updatedAt: number;
}

export const ASSISTANT_DOCUMENTS_LIMIT = 50;
export const ASSISTANT_DOCUMENT_TITLE_LIMIT = 160;
export const ASSISTANT_DOCUMENT_CONTENT_LIMIT = 64_000;

export type SaveAssistantDocumentResult =
	| { readonly status: 'saved'; readonly document: IAssistantDocument; readonly created: boolean }
	| { readonly status: 'empty' | 'tooLong' | 'full' | 'missing' };


type SaveIdeaResult = { status: 'saved' | 'duplicate'; idea: IAssistantSavedIdea } | { status: 'empty' | 'tooLong' | 'full' };

export function searchAssistantSavedIdeas(ideas: readonly IAssistantSavedIdea[], query: string): IAssistantSavedIdea[] {
	const words = query.toLowerCase().split(/\s+/).filter(Boolean);
	return ideas.filter(idea => {
		const fields = (idea.text + '\n' + idea.sourceTitle).toLowerCase();
		return words.every(word => fields.includes(word));
	}).sort((a, b) => b.createdAt - a.createdAt);
}

function encodedAssistantAccountId(accountId: string): string {
	if (!accountId) {
		throw new Error('A verified SignalREACH account id is required.');
	}
	return encodeURIComponent(accountId);
}

function savedIdeaPrefix(accountId: string): string {
	return `sessions.assistant.savedIdeas.${encodedAssistantAccountId(accountId)}/`;
}

function documentPrefix(accountId: string): string {
	return `sessions.assistant.documents.${encodedAssistantAccountId(accountId)}/`;
}

function isAssistantDocument(document: IAssistantDocument | undefined, id: string): document is IAssistantDocument {
	return !!document && document.id === id && /^[\w-]{1,64}$/.test(id)
		&& typeof document.title === 'string' && !!document.title.trim() && document.title.length <= ASSISTANT_DOCUMENT_TITLE_LIMIT
		&& typeof document.content === 'string' && !!document.content.trim() && document.content.length <= ASSISTANT_DOCUMENT_CONTENT_LIMIT
		&& typeof document.createdAt === 'number' && Number.isInteger(document.createdAt) && Number.isFinite(new Date(document.createdAt).getTime())
		&& typeof document.updatedAt === 'number' && Number.isInteger(document.updatedAt) && Number.isFinite(new Date(document.updatedAt).getTime());
}


function ideaTaskPrefix(accountId: string, kind: 'ideaTasks' | 'deletedIdeaTasks' | 'ideaTaskSeeds' | 'ideaTaskUpdates' = 'ideaTasks'): string {
	return `sessions.assistant.${kind}.${encodedAssistantAccountId(accountId)}/`;
}

function savedIdeaGenerationKey(accountId: string): string {
	return `sessions.assistant.savedIdeaGeneration.${encodedAssistantAccountId(accountId)}`;
}

function isSavedIdea(idea: IAssistantSavedIdea | undefined, id: string): idea is IAssistantSavedIdea {
	return !!idea && idea.id === id && /^[\w-]{1,64}$/.test(id)
		&& (idea.instanceId === undefined || (typeof idea.instanceId === 'string' && isUUID(idea.instanceId)))
		&& typeof idea.text === 'string' && idea.text.length <= ASSISTANT_SAVED_IDEA_TEXT_LIMIT && !!idea.text.trim()
		&& typeof idea.sourceTitle === 'string' && idea.sourceTitle.length <= ASSISTANT_SAVED_IDEA_TEXT_LIMIT
		&& (idea.taskId === undefined || (typeof idea.taskId === 'string' && idea.taskId.length <= ASSISTANT_SAVED_IDEA_TEXT_LIMIT))
		&& typeof idea.createdAt === 'number' && Number.isInteger(idea.createdAt) && Number.isFinite(new Date(idea.createdAt).getTime());
}

const PERMISSION_MODES: readonly AssistantPermissionMode[] = ['routine', 'safeOnly', 'allowAll'];

/** Stores one local conversation per verified SignalREACH account id. */
export class AssistantConversationStore {
	constructor(private readonly storage: IStorageService) { }

	onDidChange(accountId: string, disposables: DisposableStore, taskId?: string) {
		const key = this.key(accountId, taskId);
		if (taskId && !taskId.startsWith('idea-')) { return this.storage.onDidChangeValue(StorageScope.PROFILE, key, disposables); }
		const tasks = ideaTaskPrefix(accountId);
		const deleted = ideaTaskPrefix(accountId, 'deletedIdeaTasks');
		const updates = ideaTaskPrefix(accountId, 'ideaTaskUpdates');
		const seed = taskId && ideaTaskPrefix(accountId, 'ideaTaskSeeds') + taskId;
		return Event.filter(this.storage.onDidChangeValue(StorageScope.PROFILE, undefined, disposables),
			event => event.key === key || (taskId ? event.key === deleted + taskId || event.key === seed : event.key.startsWith(tasks) || event.key.startsWith(updates) || event.key.startsWith(deleted)), disposables);
	}

	listSavedIdeas(accountId: string): IAssistantSavedIdea[] {
		const prefix = savedIdeaPrefix(accountId);
		const ideas: IAssistantSavedIdea[] = [];
		for (const key of this.storage.keys(StorageScope.PROFILE)) {
			if (!key.startsWith(prefix)) { continue; }
			const idea = this.readSavedIdea(accountId, key.slice(prefix.length));
			if (idea) {
				const { id, instanceId, text, sourceTitle, taskId, createdAt } = idea;
				ideas.push({ id, ...(instanceId ? { instanceId } : {}), text, sourceTitle, ...(taskId ? { taskId } : {}), createdAt });
			}
		}
		return ideas.sort((a, b) => b.createdAt - a.createdAt);
	}

	listDocuments(accountId: string): IAssistantDocument[] {
		const prefix = documentPrefix(accountId);
		const documents: IAssistantDocument[] = [];
		for (const key of this.storage.keys(StorageScope.PROFILE)) {
			if (!key.startsWith(prefix)) { continue; }
			const document = this.readDocument(accountId, key.slice(prefix.length));
			if (document) { documents.push(document); }
		}
		return documents.sort((a, b) => b.updatedAt - a.updatedAt);
	}

	readDocument(accountId: string, id: string): IAssistantDocument | undefined {
		if (!/^[\w-]{1,64}$/.test(id)) { return undefined; }
		try {
			const document = this.storage.getObject<IAssistantDocument>(documentPrefix(accountId) + id, StorageScope.PROFILE);
			return isAssistantDocument(document, id) ? document : undefined;
		} catch {
			return undefined;
		}
	}

	saveDocument(accountId: string, title: string, content: string, id?: string): SaveAssistantDocumentResult {
		title = title.trim();
		if (!title || !content.trim()) { return { status: 'empty' }; }
		if (title.length > ASSISTANT_DOCUMENT_TITLE_LIMIT || content.length > ASSISTANT_DOCUMENT_CONTENT_LIMIT) { return { status: 'tooLong' }; }
		const existing = id === undefined ? undefined : this.readDocument(accountId, id);
		if (id !== undefined && !existing) { return { status: 'missing' }; }
		if (!existing && this.listDocuments(accountId).length >= ASSISTANT_DOCUMENTS_LIMIT) { return { status: 'full' }; }
		const now = Date.now();
		const document: IAssistantDocument = {
			id: existing?.id ?? generateUuid(), title, content,
			createdAt: existing?.createdAt ?? now, updatedAt: now,
		};
		this.storage.store(documentPrefix(accountId) + document.id, document, StorageScope.PROFILE, StorageTarget.MACHINE);
		return { status: 'saved', document, created: !existing };
	}

	removeDocument(accountId: string, id: string): boolean {
		if (!this.readDocument(accountId, id)) { return false; }
		this.storage.remove(documentPrefix(accountId) + id, StorageScope.PROFILE);
		return true;
	}

	private readSavedIdea(accountId: string, id: string): IStoredSavedIdea | undefined {
		const prefix = savedIdeaPrefix(accountId);
		if (!/^[\w-]{1,64}$/.test(id)) { return undefined; }
		try {
			const idea = this.storage.getObject<IStoredSavedIdea>(prefix + id, StorageScope.PROFILE);
			return isSavedIdea(idea, id) ? idea : undefined;
		} catch {
			return undefined;
		}
	}

	isCurrentSavedIdea(accountId: string, snapshot: IAssistantSavedIdea): boolean {
		const current = this.readSavedIdea(accountId, snapshot.id);
		return !!current && current.instanceId === snapshot.instanceId && current.createdAt === snapshot.createdAt;
	}

	saveIdea(accountId: string, text: string, sourceTitle: string, taskId?: string): SaveIdeaResult {
		const prefix = savedIdeaPrefix(accountId);
		taskId = taskId || undefined;
		if (!text.trim()) { return { status: 'empty' }; }
		if (text.length > ASSISTANT_SAVED_IDEA_TEXT_LIMIT || sourceTitle.length > ASSISTANT_SAVED_IDEA_TEXT_LIMIT || (taskId && taskId.length > ASSISTANT_SAVED_IDEA_TEXT_LIMIT)) { return { status: 'tooLong' }; }
		const ideas = this.listSavedIdeas(accountId);
		const duplicate = ideas.find(idea => idea.text === text && idea.taskId === taskId && (taskId !== undefined || idea.sourceTitle === sourceTitle));
		if (duplicate) { return { status: 'duplicate', idea: duplicate }; }
		if (ideas.length >= ASSISTANT_SAVED_IDEAS_LIMIT) { return { status: 'full' }; }

		// The same save in two windows addresses the same key. Unrelated saves write independent keys.
		const hash = new StringSHA1();
		hash.update(JSON.stringify([text, taskId, taskId ? undefined : sourceTitle]));
		let id = hash.digest();
		if (this.storage.get(prefix + id, StorageScope.PROFILE) !== undefined) { id = generateUuid(); }
		const idea = { id, instanceId: generateUuid(), text, sourceTitle, ...(taskId ? { taskId } : {}), createdAt: Date.now() };
		// Windows saving the same passage in the same generation address one continuation task.
		const continuationId = this.storage.get(savedIdeaGenerationKey(accountId), StorageScope.PROFILE);
		this.storage.store(prefix + id, { ...idea, continuationId }, StorageScope.PROFILE, StorageTarget.MACHINE);
		return { status: 'saved', idea };
	}

	removeSavedIdea(accountId: string, id: string): boolean {
		const prefix = savedIdeaPrefix(accountId);
		if (!this.readSavedIdea(accountId, id)) { return false; }
		// A new save after removal starts fresh, even if the old task still exists or is settling.
		this.storage.store(savedIdeaGenerationKey(accountId), generateUuid(), StorageScope.PROFILE, StorageTarget.MACHINE);
		this.storage.remove(prefix + id, StorageScope.PROFILE);
		return true;
	}

	read(accountId: string, taskId?: string): IAssistantConversation {
		if (taskId && this.isDeletedIdeaTask(accountId, taskId)) {
			this.clearDeletedIdeaTaskData(accountId, taskId);
			return { instructions: '', turns: [] };
		}
		const raw = this.storage.getObject<Partial<IAssistantConversation>>(this.key(accountId, taskId), StorageScope.PROFILE)
			?? (taskId?.startsWith('idea-') ? this.storage.getObject<Partial<IAssistantConversation>>(ideaTaskPrefix(accountId, 'ideaTaskSeeds') + taskId, StorageScope.PROFILE) : undefined);
		const tasks = taskId ? undefined : this.readTasks(accountId, raw?.tasks);
		return {
			instructions: typeof raw?.instructions === 'string' ? raw.instructions : '',
			...(typeof raw?.agentNotes === 'string' && raw.agentNotes ? { agentNotes: raw.agentNotes } : {}),
			...(Array.isArray(raw?.mailAlerts) ? { mailAlerts: raw.mailAlerts.filter((alert): alert is IAssistantMailAlert => typeof alert?.id === 'string' && typeof alert?.subject === 'string' && typeof alert?.label === 'string' && typeof alert?.at === 'number') } : {}),
			...(typeof raw?.draft === 'string' && raw.draft ? { draft: raw.draft } : {}),
			...(typeof raw?.draftContext === 'string' && raw.draftContext ? { draftContext: raw.draftContext } : {}),
			...(Array.isArray(raw?.queuedPrompts) ? { queuedPrompts: raw.queuedPrompts.filter(isQueuedPrompt) } : {}),
			...(isQueuedPrompt(raw?.startedQueuedPrompt) ? { startedQueuedPrompt: raw.startedQueuedPrompt } : {}),
			turns: Array.isArray(raw?.turns) ? raw.turns.filter(turn => typeof turn?.prompt === 'string' && typeof turn.reply === 'string') : [],
			...(typeof raw?.modelId === 'string' && raw.modelId ? { modelId: raw.modelId } : {}),
			...(tasks ? { tasks } : {}),
			...(!taskId && Array.isArray(raw?.agentChats) ? { agentChats: raw.agentChats.filter(chat => typeof chat?.id === 'string' && typeof chat.resource === 'string' && typeof chat.title === 'string') } : {}),
			...(typeof raw?.permissionMode === 'string' && (PERMISSION_MODES as readonly string[]).includes(raw.permissionMode) ? { permissionMode: raw.permissionMode } : {}),
			...(Array.isArray(raw?.allowedTools) ? { allowedTools: raw.allowedTools.filter((tool): tool is string => typeof tool === 'string') } : {}),
		};
	}

	getLastWorktreeProject(accountId: string): URI | undefined {
		return readLocalPath(this.storage.get(`sessions.assistant.lastWorktreeProject.${encodeURIComponent(accountId)}`, StorageScope.PROFILE));
	}

	setLastWorktreeProject(accountId: string, project: URI | undefined): void {
		const key = `sessions.assistant.lastWorktreeProject.${encodeURIComponent(accountId)}`;
		if (project?.scheme === Schemas.file) {
			this.storage.store(key, project.toString(), StorageScope.PROFILE, StorageTarget.MACHINE);
		} else {
			this.storage.remove(key, StorageScope.PROFILE);
		}
	}

	getLastOpenTask(accountId: string, project: URI | undefined): string | undefined {
		const saved = this.storage.getObject<{ id?: unknown; project?: unknown }>(this.openTaskKey(accountId), StorageScope.PROFILE);
		if (typeof saved?.id !== 'string') { return undefined; }
		const sameProject = saved.project === undefined ? project === undefined : !!project && extUriBiasedIgnorePathCase.isEqual(readUri(saved.project), project);
		return sameProject && this.read(accountId).tasks?.some(task => task.id === saved.id) ? saved.id : undefined;
	}

	setLastOpenTask(accountId: string, project: URI | undefined, taskId: string | undefined): void {
		const key = this.openTaskKey(accountId);
		if (taskId) {
			this.storage.store(key, { id: taskId, project: project?.toString() }, StorageScope.PROFILE, StorageTarget.MACHINE);
		} else {
			this.storage.remove(key, StorageScope.PROFILE);
		}
	}

	private openTaskKey(accountId: string): string {
		return `sessions.assistant.lastOpenTask.${encodeURIComponent(accountId)}`;
	}

	getWorktreeSelection(accountId: string, project: URI): IAssistantWorktreeSelection | undefined {
		const selections = this.readWorktreeSelections(accountId);
		return selections.find(selection => extUriBiasedIgnorePathCase.isEqual(project, selection.project))
			?? selections.find(selection => extUriBiasedIgnorePathCase.isEqual(project, selection.worktree.uri));
	}

	setWorktreeSelection(accountId: string, project: URI, worktree: IAssistantWorktree | undefined): void {
		const selections = this.readWorktreeSelections(accountId).filter(selection => !extUriBiasedIgnorePathCase.isEqual(project, selection.project));
		if (worktree) {
			selections.unshift({ project, worktree });
		}
		this.storage.store(`sessions.assistant.worktreeSelections.${encodeURIComponent(accountId)}`, selections.map(selection => ({
			project: selection.project.toString(), ...storeWorktree(selection.worktree),
		})), StorageScope.PROFILE, StorageTarget.MACHINE);
	}

	private readWorktreeSelections(accountId: string): IAssistantWorktreeSelection[] {
		const raw = this.storage.getObject<readonly (IStoredWorktree & { project?: unknown })[]>(`sessions.assistant.worktreeSelections.${encodeURIComponent(accountId)}`, StorageScope.PROFILE);
		if (!Array.isArray(raw)) { return []; }
		const selections: IAssistantWorktreeSelection[] = [];
		for (const entry of raw) {
			const project = readLocalPath(entry?.project);
			const worktree = readWorktree(entry);
			if (project && worktree) { selections.push({ project, worktree }); }
		}
		return selections;
	}

	setInstructions(accountId: string, instructions: string): void {
		this.write(accountId, { ...this.read(accountId), instructions });
	}

	readMemoryBank(accountId: string): AssistantMemoryBank {
		const conversation = this.read(accountId);
		return { user: conversation.instructions, memory: conversation.agentNotes ?? '' };
	}

	writeMemoryBank(accountId: string, bank: AssistantMemoryBank): void {
		this.write(accountId, { ...this.read(accountId), instructions: bank.user, agentNotes: bank.memory });
	}

	/** Removes only the entries unchecked in the picker, keeping notes added while it was open. */
	forgetMemoryEntries(accountId: string, unchecked: Readonly<Record<AssistantMemoryTarget, readonly string[]>>): number {
		const bank = this.readMemoryBank(accountId);
		const next = { ...bank };
		let removed = 0;
		for (const target of ['user', 'memory'] as const) {
			const entries = memoryEntries(bank[target]);
			const kept = entries.filter(entry => !unchecked[target].includes(entry));
			if (kept.length !== entries.length) {
				removed += entries.length - kept.length;
				next[target] = kept.join('\n');
			}
		}
		if (removed) { this.writeMemoryBank(accountId, next); }
		return removed;
	}

	applyMemory(accountId: string, target: AssistantMemoryTarget, operations: readonly AssistantMemoryOperation[]): AssistantMemoryResult {
		const result = applyMemoryOperations(this.readMemoryBank(accountId), target, operations);
		if (result.ok) {
			this.writeMemoryBank(accountId, result.bank);
		}
		return result;
	}

	recordMailAlerts(accountId: string, alerts: readonly IAssistantMailAlert[]): number {
		if (!alerts.length) { return 0; }
		const conversation = this.read(accountId);
		const existing = conversation.mailAlerts ?? [];
		const known = new Set(existing.map(alert => alert.id));
		const fresh = alerts.filter(alert => !known.has(alert.id));
		if (!fresh.length) { return 0; }
		this.write(accountId, { ...conversation, mailAlerts: [...fresh, ...existing].slice(0, 8) });
		return fresh.length;
	}

	remember(accountId: string, notes: readonly string[]): { added: number; total: number } {
		const bank = this.readMemoryBank(accountId);
		const entries = memoryEntries(bank.user);
		const originalLength = entries.length;
		for (const note of notes.flatMap(memoryEntries)) {
			if (note.length + 2 <= ASSISTANT_MEMORY_LIMITS.user && !entries.includes(note)) {
				entries.push(note);
			}
		}
		const added = entries.length - originalLength;
		if (!added) {
			return { added: 0, total: originalLength };
		}
		// Keep the newest entries that fit, including each stored bullet and newline.
		const lines = entries.map(entry => `- ${entry}`);
		let first = lines.length;
		let length = 0;
		while (first > 0) {
			const nextLength = length + lines[first - 1].length + (length ? 1 : 0);
			if (nextLength > ASSISTANT_MEMORY_LIMITS.user) { break; }
			first--;
			length = nextLength;
		}
		const kept = lines.slice(first);
		this.writeMemoryBank(accountId, { ...bank, user: kept.join('\n') });
		return { added: Math.min(added, kept.length), total: kept.length };
	}

	setDraft(accountId: string, draft: string, taskId?: string): void {
		this.setComposerDraft(accountId, { inputText: draft, attachments: [] }, taskId);
	}

	readComposerDraft(accountId: string, taskId?: string): IChatDraft {
		const current = this.read(accountId, taskId);
		if (current.draftContext) {
			try { return fromStoredChatDraft(parse(current.draftContext)); }
			catch { /* A malformed attachment draft can still restore its plain text. */ }
		}
		return { inputText: current.draft ?? '', attachments: [] };
	}

	/** The prepared prompt keys retry and completion while the editable composer keeps its attachment references. */
	setComposerDraft(accountId: string, input: IChatDraft, taskId?: string, preparedPrompt?: string): void {
		const current = this.read(accountId, taskId);
		const draftContext = input.attachments.length || input.references?.length ? stringify(toStoredChatDraft(input)) : undefined;
		const savedPrompt = draftContext && draftContext === current.draftContext ? current.draft : undefined;
		const draft = preparedPrompt ?? savedPrompt ?? input.inputText;
		if ((current.draft ?? '') !== draft || current.draftContext !== draftContext) {
			this.write(accountId, { ...current, draft, draftContext }, taskId);
		}
	}

	queuePrompt(accountId: string, taskId: string | undefined, prompt: string, target: IAssistantAgentTarget, requestStart?: number, imageUrls?: readonly string[], serverIds?: readonly string[]): void {
		const conversation = this.read(accountId, taskId);
		this.write(accountId, { ...conversation, queuedPrompts: [...(conversation.queuedPrompts ?? []), { id: generateUuid(), prompt, target: storeAssistantTaskTarget(target), ...(requestStart !== undefined ? { requestStart } : {}), ...(imageUrls?.length ? { imageUrls } : {}), ...(serverIds?.length ? { serverIds } : {}) }] }, taskId);
	}

	startQueuedPrompt(accountId: string, taskId: string | undefined, id: string): boolean {
		const conversation = this.read(accountId, taskId);
		if (conversation.startedQueuedPrompt) { return conversation.startedQueuedPrompt.id === id; }
		const [next, ...remaining] = conversation.queuedPrompts ?? [];
		if (!next || (next.id && next.id !== id)) { return false; }
		this.write(accountId, { ...conversation, queuedPrompts: remaining, startedQueuedPrompt: { ...next, id } }, taskId);
		return true;
	}

	takeQueuedPrompt(accountId: string, taskId: string | undefined, includeStarted: boolean): IAssistantQueuedPrompt | undefined {
		const conversation = this.read(accountId, taskId);
		if (includeStarted && conversation.startedQueuedPrompt) {
			this.write(accountId, { ...conversation, startedQueuedPrompt: undefined }, taskId);
			return conversation.startedQueuedPrompt;
		}
		const [next, ...remaining] = conversation.queuedPrompts ?? [];
		if (next) { this.write(accountId, { ...conversation, queuedPrompts: remaining }, taskId); }
		return next;
	}

	setModel(accountId: string, modelId: string | undefined, taskId?: string): void {
		this.write(accountId, { ...this.read(accountId, taskId), modelId }, taskId);
	}

	setPermissionMode(accountId: string, mode: AssistantPermissionMode, taskId?: string): void {
		this.write(accountId, { ...this.read(accountId, taskId), permissionMode: mode }, taskId);
	}

	allowTool(accountId: string, tool: string, taskId?: string): void {
		const current = this.read(accountId, taskId);
		const allowed = current.allowedTools ?? [];
		if (!allowed.includes(tool)) {
			this.write(accountId, { ...current, allowedTools: [...allowed, tool] }, taskId);
		}
	}

	setAllowedTools(accountId: string, allowedTools: readonly string[], taskId?: string): void {
		this.write(accountId, { ...this.read(accountId, taskId), allowedTools: [...allowedTools] }, taskId);
	}

	clearTurns(accountId: string, taskId?: string): void {
		this.write(accountId, { ...this.read(accountId, taskId), turns: [] }, taskId);
	}

	popTurn(accountId: string, taskId?: string): IAssistantTurn | undefined {
		const current = this.read(accountId, taskId);
		const last = current.turns.at(-1);
		if (last) { this.write(accountId, { ...current, turns: current.turns.slice(0, -1) }, taskId); }
		return last;
	}

	createTask(accountId: string, target?: IAssistantAgentTarget): string {
		const current = this.read(accountId);
		const id = generateUuid();
		return this.addTask(accountId, current, { id, title: '', status: 'draft', target: target && storeAssistantTaskTarget(target) });
	}

	getOrCreateSavedIdeaTask(accountId: string, ideaId: string, title: string, attachmentLabel: string, target: IAssistantAgentTarget): { taskId: string } | 'tooLong' | undefined {
		const idea = this.readSavedIdea(accountId, ideaId);
		if (!idea) { return undefined; }
		const current = this.read(accountId);
		const firstId = savedIdeaTaskId(accountId, idea);
		let id = firstId;
		for (let attempt = 1; this.isDeletedIdeaTask(accountId, id); attempt++) {
			id = `${firstId}-${attempt}`;
		}
		if (current.tasks?.some(task => task.id === id && task.savedIdeaId === ideaId)) { return { taskId: id }; }
		if (assistantTextAttachment(attachmentLabel, idea.text).length >= REACH_ASSISTANT_PROMPT_LIMIT) { return 'tooLong'; }
		const attachment = toPasteVariableEntry(attachmentLabel, idea.text, { id: `assistant-saved:${idea.id}`, icon: Codicon.bookmark });
		return { taskId: this.addTask(accountId, current, { id, title, savedIdeaId: ideaId, status: 'draft', target: storeAssistantTaskTarget(target) }, { inputText: '', attachments: [attachment] }) };
	}

	private addTask(accountId: string, current: IAssistantConversation, task: IAssistantTask, draft?: IChatDraft): string {
		if (task.savedIdeaId && draft) {
			// Another window may already be editing the child. Initial creation never writes mutable state.
			this.storage.store(ideaTaskPrefix(accountId, 'ideaTaskSeeds') + task.id, { instructions: '', turns: [], modelId: current.modelId, draftContext: stringify(toStoredChatDraft(draft)) }, StorageScope.PROFILE, StorageTarget.MACHINE);
			this.storage.store(ideaTaskPrefix(accountId) + task.id, task, StorageScope.PROFILE, StorageTarget.MACHINE);
		} else {
			if (current.modelId) { this.setModel(accountId, current.modelId, task.id); }
			this.write(accountId, { ...current, tasks: [...(current.tasks ?? []), task] });
		}
		return task.id;
	}

	addAgentChat(accountId: string, resource: string, title: string): string {
		const current = this.read(accountId);
		const id = generateUuid();
		this.write(accountId, { ...current, agentChats: [...(current.agentChats ?? []), { id, resource, title }] });
		return id;
	}

	updateTask(accountId: string, taskId: string, change: Partial<Omit<IAssistantTask, 'id'>>): void {
		if (this.isDeletedIdeaTask(accountId, taskId)) { return; }
		if (taskId.startsWith('idea-')) {
			const key = ideaTaskPrefix(accountId) + taskId;
			const task = this.storage.getObject<IAssistantTask>(key, StorageScope.PROFILE);
			if (task?.id === taskId) {
				const updates = ideaTaskPrefix(accountId, 'ideaTaskUpdates') + taskId;
				const previous = this.storage.getObject<Partial<IAssistantTask>>(updates, StorageScope.PROFILE);
				this.storage.store(updates, { ...task, ...previous, ...change }, StorageScope.PROFILE, StorageTarget.MACHINE);
				return;
			}
		}
		const current = this.read(accountId);
		if (!current.tasks?.some(task => task.id === taskId)) { return; }
		this.write(accountId, { ...current, tasks: current.tasks?.map(task => task.id === taskId ? { ...task, ...change } : task) });
	}

	deleteTask(accountId: string, taskId: string): void {
		if (this.isDeletedIdeaTask(accountId, taskId)) { this.clearDeletedIdeaTaskData(accountId, taskId); return; }
		const current = this.read(accountId);
		const ideaId = current.tasks?.find(task => task.id === taskId)?.savedIdeaId;
		if (ideaId) {
			// Keep deletion independent of the mutable records another window might still write.
			const dedicated = this.storage.get(ideaTaskPrefix(accountId) + taskId, StorageScope.PROFILE) !== undefined;
			this.storage.store(ideaTaskPrefix(accountId, 'deletedIdeaTasks') + taskId, true, StorageScope.PROFILE, StorageTarget.MACHINE);
			this.clearDeletedIdeaTaskData(accountId, taskId);
			if (dedicated) { return; }
		}
		this.write(accountId, { ...current, tasks: current.tasks?.filter(task => task.id !== taskId) });
		this.storage.remove(this.key(accountId, taskId), StorageScope.PROFILE);
	}

	appendTurn(accountId: string, turn: IAssistantTurn, taskId?: string, completedQueueId?: string, replacement?: IAssistantTurnReplacement): boolean {
		const current = this.read(accountId, taskId);
		if (replacement && !equals(current.turns, replacement.original)) { return false; }
		const settled = completedQueueId && current.startedQueuedPrompt?.id === completedQueueId;
		const earlier = replacement ? current.turns.slice(0, replacement.index) : current.turns;
		return this.write(accountId, { ...current, turns: [...earlier, { at: Date.now(), ...turn }], ...(settled ? { startedQueuedPrompt: undefined } : {}) }, taskId);
	}

	private readTasks(accountId: string, embedded: readonly IAssistantTask[] | undefined): IAssistantTask[] | undefined {
		const tasks = new Map<string, IAssistantTask>();
		if (Array.isArray(embedded)) {
			for (const task of embedded) {
				if (typeof task?.id === 'string' && typeof task.title === 'string' && !this.isDeletedIdeaTask(accountId, task.id)) { tasks.set(task.id, task); }
			}
		}
		const prefix = ideaTaskPrefix(accountId);
		const deleted = ideaTaskPrefix(accountId, 'deletedIdeaTasks');
		for (const key of this.storage.keys(StorageScope.PROFILE)) {
			if (key.startsWith(deleted)) {
				this.clearDeletedIdeaTaskData(accountId, key.slice(deleted.length));
			} else if (key.startsWith(prefix)) {
				const id = key.slice(prefix.length);
				if (this.isDeletedIdeaTask(accountId, id)) { continue; }
				try {
					const task = this.storage.getObject<IAssistantTask>(ideaTaskPrefix(accountId, 'ideaTaskUpdates') + id, StorageScope.PROFILE)
						?? this.storage.getObject<IAssistantTask>(key, StorageScope.PROFILE);
					if (task?.id === id && typeof task.title === 'string' && typeof task.savedIdeaId === 'string') { tasks.set(id, task); }
				} catch { /* A malformed individual task must not hide the other tasks. */ }
			}
		}
		return tasks.size || Array.isArray(embedded) ? [...tasks.values()] : undefined;
	}

	private isDeletedIdeaTask(accountId: string, taskId: string): boolean {
		return taskId.startsWith('idea-') && this.storage.getBoolean(ideaTaskPrefix(accountId, 'deletedIdeaTasks') + taskId, StorageScope.PROFILE, false);
	}

	private clearDeletedIdeaTaskData(accountId: string, taskId: string): void {
		if (!taskId.startsWith('idea-')) { return; }
		const keys = [ideaTaskPrefix(accountId) + taskId, ideaTaskPrefix(accountId, 'ideaTaskUpdates') + taskId, ideaTaskPrefix(accountId, 'ideaTaskSeeds') + taskId, this.key(accountId, taskId)];
		for (const key of keys) {
			if (this.storage.get(key, StorageScope.PROFILE) !== undefined) { this.storage.remove(key, StorageScope.PROFILE); }
		}
	}

	private write(accountId: string, conversation: IAssistantConversation, taskId?: string): boolean {
		if (taskId && this.isDeletedIdeaTask(accountId, taskId)) { return false; }
		if (!taskId && conversation.tasks?.some(task => task.savedIdeaId)) {
			// Legacy embedded tasks stay embedded. Never copy independent registrations back into the root.
			const prefix = ideaTaskPrefix(accountId);
			conversation = { ...conversation, tasks: conversation.tasks.filter(task => !task.savedIdeaId ||
				(!this.isDeletedIdeaTask(accountId, task.id) && this.storage.get(prefix + task.id, StorageScope.PROFILE) === undefined)) };
		}
		this.storage.store(this.key(accountId, taskId), conversation, StorageScope.PROFILE, StorageTarget.MACHINE);
		if (taskId?.startsWith('idea-')) { this.storage.remove(ideaTaskPrefix(accountId, 'ideaTaskSeeds') + taskId, StorageScope.PROFILE); }
		return true;
	}

	private key(accountId: string, taskId?: string): string {
		const owner = encodedAssistantAccountId(accountId);
		return taskId
			? `sessions.assistant.taskConversation.${owner}.${encodeURIComponent(taskId)}`
			: `sessions.assistant.conversation.${owner}`;
	}
}
