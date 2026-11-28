/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { raceTimeout } from '../../../../base/common/async.js';
import { CancellationToken, CancellationTokenSource } from '../../../../base/common/cancellation.js';
import { derived, IObservable, ISettableObservable, observableValue, waitForState } from '../../../../base/common/observable.js';
import { extUriBiasedIgnorePathCase } from '../../../../base/common/resources.js';
import { URI } from '../../../../base/common/uri.js';
import { localize } from '../../../../nls.js';
import { IFileService } from '../../../../platform/files/common/files.js';
import { InstantiationType, registerSingleton } from '../../../../platform/instantiation/common/extensions.js';
import { createDecorator } from '../../../../platform/instantiation/common/instantiation.js';
import { IAgentsEditorSelection } from '../../../../platform/native/common/native.js';
import { IStorageService } from '../../../../platform/storage/common/storage.js';
import { IWorkspaceTrustManagementService } from '../../../../platform/workspace/common/workspaceTrust.js';
import { IChatService } from '../../../../workbench/contrib/chat/common/chatService/chatService.js';
import { ChatAgentLocation, ChatPermissionLevel } from '../../../../workbench/contrib/chat/common/constants.js';
import { ISessionsNavigationService } from '../../../../workbench/contrib/sessions/browser/sessionsNavigationService.js';
import { ICreateNewSessionOptions, IProviderSessionType, ISessionsManagementService } from '../../../../workbench/contrib/sessions/common/sessionsManagement.js';
import { isActiveSessionStatus, SessionStatus, SessionTypeAuthRequirement } from '../../../../workbench/services/sessions/common/session.js';
import { AssistantConversationStore } from './assistantConversation.js';
import { IAssistantWorktree, readAssistantWorktreeIdentity } from './assistantWorktree.js';

export interface IAssistantAgentTarget {
	readonly workspace?: URI;
	readonly worktree?: IAssistantWorktree;
	readonly fileRoots?: readonly URI[];
	readonly fileRootNames?: readonly string[];
	readonly activeEditorFile?: URI;
	readonly providerId?: string;
	readonly sessionTypeId?: string;
}

interface IAssistantEditorSelection {
	readonly file: URI;
	readonly selection: IAgentsEditorSelection;
}

export function getAssistantEditorFolders(target: IAssistantAgentTarget | undefined): readonly URI[] {
	if (target?.worktree) {
		return [target.worktree.uri];
	}
	if (target?.fileRoots?.length) {
		return target.fileRoots;
	}
	return target?.workspace ? [target.workspace] : [];
}

export type AssistantWorktreeStatus = 'ready' | 'missing' | 'untrusted' | 'changed' | 'unverified';

export function assistantWorktreeError(status: AssistantWorktreeStatus): string | undefined {
	switch (status) {
		case 'missing': return localize('assistantWorktree.missing', "The selected Git worktree is unavailable. Choose another worktree or use project defaults.");
		case 'untrusted': return localize('assistantWorktree.untrusted', "The selected Git worktree needs trust. Select it again to review workspace trust.");
		case 'changed': return localize('assistantWorktree.changed', "The Git checkout at the selected path has changed. Choose a worktree or use project defaults.");
		case 'unverified': return localize('assistantWorktree.unverified', "Select this Git worktree again to verify its repository.");
		case 'ready': return undefined;
	}
}

export interface IAssistantDelegatedChat {
	readonly id: string;
	readonly title: string;
	readonly workspaceLabel: string | undefined;
	readonly status: SessionStatus;
}

export interface IAssistantAgentChats {
	readonly _serviceBrand: undefined;
	/** Captured before the Assistant view replaces the current project view. */
	target: IAssistantAgentTarget | undefined;
	readonly editorSelection: ISettableObservable<IAssistantEditorSelection | undefined>;
	/** Local IDE roots attached to the current New Chat workspace selection. */
	editorWorkspace?: {
		readonly selectedFolder: URI;
		readonly fileRoots: readonly URI[];
		readonly fileRootNames: readonly string[];
		readonly activeFile?: URI;
		readonly selectionVersion: number;
		readonly navigationVersion: number;
	};
	getWorktreeStatus(worktree: IAssistantWorktree): Promise<AssistantWorktreeStatus>;
	isResourceTrusted(uri: URI): Promise<boolean>;
	getChats(accountId: string): IAssistantDelegatedChat[];
	open(accountId: string, id: string): Promise<boolean>;
	list(accountId: string, target: IAssistantAgentTarget | undefined): string;
	start(accountId: string, target: IAssistantAgentTarget | undefined, prompt: string, title: string, agentType: string | undefined, token: CancellationToken): Promise<string>;
	send(accountId: string, id: string, prompt: string): Promise<string>;
	read(accountId: string, id: string, wait: boolean, token: CancellationToken): Promise<string>;
}

export const IAssistantAgentChats = createDecorator<IAssistantAgentChats>('assistantAgentChats');

const unknownAgent = 'No available agent chat with that ID. Use assistant.list_agents to see your agent chats.';
const replyWaitMs = 25_000;

async function waitUntil<T>(state: IObservable<T>, predicate: (value: T) => boolean, deadline: number, token: CancellationToken): Promise<void> {
	const remaining = deadline - Date.now();
	if (remaining <= 0 || predicate(state.get())) { return; }
	const source = new CancellationTokenSource(token);
	try {
		await raceTimeout(waitForState(state, predicate, undefined, source.token), remaining);
	} finally {
		source.dispose(true);
	}
}

function agentTypeKey({ providerId, sessionType }: IProviderSessionType): string {
	return `${providerId}/${sessionType.id}`;
}

export function describeAssistantAgentStatus(status: SessionStatus): string {
	switch (status) {
		case SessionStatus.InProgress: return localize('assistantAgent.statusWorking', "working");
		case SessionStatus.NeedsInput: return localize('assistantAgent.statusNeedsInput', "waiting for user input in its chat");
		case SessionStatus.Completed: return localize('assistantAgent.statusFinished', "finished");
		case SessionStatus.Error: return localize('assistantAgent.statusFailed', "failed");
		default: return localize('assistantAgent.statusStarting', "starting");
	}
}

export class AssistantAgentChats implements IAssistantAgentChats {
	declare readonly _serviceBrand: undefined;
	target: IAssistantAgentTarget | undefined;
	readonly editorSelection = observableValue<IAssistantEditorSelection | undefined>(this, undefined);
	editorWorkspace: IAssistantAgentChats['editorWorkspace'];
	private readonly conversations: AssistantConversationStore;
	private readonly sending = new Set<string>();

	constructor(
		@ISessionsManagementService private readonly sessions: ISessionsManagementService,
		@IChatService private readonly chatService: IChatService,
		@IStorageService storage: IStorageService,
		@ISessionsNavigationService private readonly navigation: ISessionsNavigationService,
		@IFileService private readonly fileService: IFileService,
		@IWorkspaceTrustManagementService private readonly workspaceTrust: IWorkspaceTrustManagementService,
	) {
		this.conversations = new AssistantConversationStore(storage);
	}

	async isResourceTrusted(uri: URI): Promise<boolean> {
		try {
			return (await this.workspaceTrust.getUriTrustInfo(uri)).trusted;
		} catch {
			return false;
		}
	}

	async getWorktreeStatus({ uri, identity }: IAssistantWorktree): Promise<AssistantWorktreeStatus> {
		if (!await this.fileService.exists(URI.joinPath(uri, '.git'))) {
			return 'missing';
		}
		if (!await this.isResourceTrusted(uri)) { return 'untrusted'; }
		const current = await readAssistantWorktreeIdentity(this.fileService, uri);
		if (!current) { return 'missing'; }
		if (!identity) { return 'unverified'; }
		return identity.createdAt === current.createdAt && extUriBiasedIgnorePathCase.isEqual(identity.gitDirectory, current.gitDirectory) ? 'ready' : 'changed';
	}

	private availableTypes(target: IAssistantAgentTarget | undefined): IProviderSessionType[] {
		const folder = target?.worktree?.uri ?? target?.workspace;
		const types = folder ? this.sessions.getSessionTypesForFolder(folder) : this.sessions.getQuickChatSessionTypes();
		return types.filter(type => type.sessionType.authRequirement !== SessionTypeAuthRequirement.Unusable
			&& (!target?.worktree || type.sessionType.supportsWorktreeConfiguration));
	}

	list(accountId: string, target: IAssistantAgentTarget | undefined): string {
		const types = this.availableTypes(target).map(type => `${type.sessionType.label}: ${agentTypeKey(type)}`);
		const chats = this.getChats(accountId).map(chat => `${chat.id}: ${chat.title} (${describeAssistantAgentStatus(chat.status)})`);
		return `Available agent types:\n${types.join('\n') || 'None available for this context.'}\n\nAgent chats:\n${chats.join('\n') || 'No agent chats started yet.'}`;
	}

	getChats(accountId: string): IAssistantDelegatedChat[] {
		const chats: IAssistantDelegatedChat[] = [];
		for (const entry of this.conversations.read(accountId).agentChats ?? []) {
			const chat = this.sessions.getSessionForChatResource(URI.parse(entry.resource))?.chat;
			if (!chat) {
				continue;
			}
			const model = this.chatService.getSession(chat.resource);
			const reportedStatus = chat.status.get();
			const status = model?.requestNeedsInput.get() ? SessionStatus.NeedsInput
				: reportedStatus !== SessionStatus.NeedsInput && model?.hasActiveRequest.get() ? SessionStatus.InProgress
					: reportedStatus;
			chats.push({ id: entry.id, title: chat.title.get() || entry.title, workspaceLabel: chat.workspace.get()?.label, status });
		}
		return chats;
	}

	async open(accountId: string, id: string): Promise<boolean> {
		const owned = this.find(accountId, id);
		if (!owned) {
			return false;
		}
		await this.navigation.openChat(owned.session, owned.chat.resource, { source: 'link' });
		return true;
	}

	async start(accountId: string, target: IAssistantAgentTarget | undefined, prompt: string, title: string, agentType: string | undefined, token: CancellationToken): Promise<string> {
		if (!prompt) { return 'Provide a task for the agent.'; }
		const selectedWorktree = target?.worktree;
		const worktree = selectedWorktree?.uri;
		if (selectedWorktree) {
			const error = assistantWorktreeError(await this.getWorktreeStatus(selectedWorktree));
			if (error) { return error; }
		}
		const types = this.availableTypes(target);
		const selected = agentType ? types.find(type => agentTypeKey(type) === agentType) :
			types.find(type => type.providerId === target?.providerId && type.sessionType.id === target?.sessionTypeId)
			?? types.find(type => type.sessionType.authRequirement === SessionTypeAuthRequirement.None) ?? types[0];
		if (!selected) { return 'No matching agent type. Use assistant.list_agents to see the options for the current context.'; }
		const name = title.slice(0, 100) || prompt.slice(0, 100);
		const request = { query: prompt, title: name, background: true, requireConnected: true };
		const options: ICreateNewSessionOptions = {
			providerId: selected.providerId,
			sessionTypeId: selected.sessionType.id,
			permissionLevel: ChatPermissionLevel.Default,
			...(selectedWorktree && {
				isolationMode: 'workspace',
				onSessionCreated: async session => {
					const workingDirectory = session.workspace.get()?.folders[0]?.workingDirectory;
					if (!extUriBiasedIgnorePathCase.isEqual(selectedWorktree.uri, workingDirectory)) {
						throw new Error('The agent could not open the selected Git worktree.');
					}
					const error = assistantWorktreeError(await this.getWorktreeStatus(selectedWorktree));
					if (error) { throw new Error(error); }
				},
			}),
		};
		const folder = worktree ?? target?.workspace;
		const session = folder
			? await this.sessions.createAndSendNewChatRequest(folder, request, options, token)
			: await this.sessions.createAndSendQuickChatRequest(request, options, token);
		if (!session) { return 'Could not confirm that the agent chat started. Check the regular Chats list before retrying.'; }
		const chat = session.mainChat.get();
		const id = this.conversations.addAgentChat(accountId, chat.resource.toString(), name);
		return `Started ${selected.sessionType.label} in the regular chat "${name}". Agent chat ID: ${id}. Use assistant.read_agent to check the reply.`;
	}

	async send(accountId: string, id: string, prompt: string): Promise<string> {
		if (!prompt) { return 'Provide a follow-up message for the agent.'; }
		const owned = this.find(accountId, id);
		if (!owned || owned.chat.isArchived.get() || owned.session.isArchived.get()) { return unknownAgent; }
		const resource = owned.chat.resource.toString();
		if (isActiveSessionStatus(owned.chat.status.get()) || this.sending.has(resource) || this.chatService.getSession(owned.chat.resource)?.hasActiveRequest.get()) {
			return 'The agent chat is busy. Read its status before sending another message.';
		}
		this.sending.add(resource);
		try {
			await this.sessions.sendRequestInBackground(owned.session, owned.chat, { query: prompt });
			return `Sent the follow-up to agent chat ${id}. Use assistant.read_agent to check its reply.`;
		} finally {
			this.sending.delete(resource);
		}
	}

	async read(accountId: string, id: string, wait: boolean, token: CancellationToken): Promise<string> {
		const owned = this.find(accountId, id);
		if (!owned) { return unknownAgent; }
		const { chat } = owned;
		const deadline = wait ? Date.now() + replyWaitMs : 0;
		const header = (status: SessionStatus) => `Agent chat ${id} is ${describeAssistantAgentStatus(status)}.`;
		await waitUntil(chat.status, status => status !== SessionStatus.InProgress, deadline, token);
		const reportedStatus = chat.status.get();
		if (reportedStatus !== SessionStatus.Completed && reportedStatus !== SessionStatus.Error) { return header(reportedStatus); }
		const reference = await this.chatService.acquireOrLoadSession(chat.resource, ChatAgentLocation.Chat, token, 'assistantAgentChats');
		try {
			const latest = reference?.object.getRequests().findLast(request => !request.isRequestHiddenFromTranscript)?.response;
			// The session list can still report completion of the previous turn when a follow-up starts.
			if (latest?.isIncomplete.get()) {
				const replyPending = derived(reader => latest.isIncomplete.read(reader) && !latest.isPendingConfirmation.read(reader));
				await waitUntil(replyPending, pending => !pending, deadline, token);
				if (latest.isPendingConfirmation.get()) { return header(SessionStatus.NeedsInput); }
				if (latest.isIncomplete.get()) { return header(SessionStatus.InProgress); }
			}
			const reply = latest?.isComplete ? latest.response.getFinalResponse().trim() : '';
			return reply ? `${header(reportedStatus)}\nLatest reply:\n${reply.slice(0, 5_000)}` : `${header(reportedStatus)} No completed reply is available.`;
		} finally {
			reference?.dispose();
		}
	}

	private find(accountId: string, id: string) {
		const entry = this.conversations.read(accountId).agentChats?.find(chat => chat.id === id);
		return entry ? this.sessions.getSessionForChatResource(URI.parse(entry.resource)) : undefined;
	}
}

registerSingleton(IAssistantAgentChats, AssistantAgentChats, InstantiationType.Delayed);
