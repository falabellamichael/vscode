/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { append, clearNode, $, EventHelper, EventType, addDisposableListener, isHTMLElement } from '../../../../base/browser/dom.js';
import { renderMarkdown } from '../../../../base/browser/markdownRenderer.js';
import { ActionBar } from '../../../../base/browser/ui/actionbar/actionbar.js';
import { Action, Separator } from '../../../../base/common/actions.js';
import { IntervalTimer, RunOnceScheduler } from '../../../../base/common/async.js';
import { encodeBase64, VSBuffer } from '../../../../base/common/buffer.js';
import { cancelOnDispose, CancellationToken } from '../../../../base/common/cancellation.js';
import { Codicon } from '../../../../base/common/codicons.js';
import { Emitter, Event } from '../../../../base/common/event.js';
import { hash } from '../../../../base/common/hash.js';
import { fromNow } from '../../../../base/common/date.js';
import { escapeIcons } from '../../../../base/common/iconLabels.js';
import { detectImageMimeType } from '../../../../base/common/image.js';
import { Disposable, DisposableStore, MutableDisposable } from '../../../../base/common/lifecycle.js';
import { equals } from '../../../../base/common/objects.js';
import { autorun, constObservable, ISettableObservable, observableValue } from '../../../../base/common/observable.js';
import { basename, extUriBiasedIgnorePathCase, joinPath } from '../../../../base/common/resources.js';
import { escapeRegExpCharacters } from '../../../../base/common/strings.js';
import { matchesSomeScheme, Schemas } from '../../../../base/common/network.js';
import { URI } from '../../../../base/common/uri.js';
import { ThemeIcon } from '../../../../base/common/themables.js';
import { isCodeEditor } from '../../../../editor/browser/editorBrowser.js';
import { TextLength } from '../../../../editor/common/core/text/textLength.js';
import { localize, localize2 } from '../../../../nls.js';
import { Action2, registerAction2 } from '../../../../platform/actions/common/actions.js';
import { ContextKeyExpr } from '../../../../platform/contextkey/common/contextkey.js';
import { IClipboardService } from '../../../../platform/clipboard/common/clipboardService.js';
import { CommandsRegistry, ICommandService } from '../../../../platform/commands/common/commands.js';
import { IConfigurationService } from '../../../../platform/configuration/common/configuration.js';
import { IDialogService, IFileDialogService } from '../../../../platform/dialogs/common/dialogs.js';
import { SyncDescriptor } from '../../../../platform/instantiation/common/descriptors.js';
import { IInstantiationService, ServicesAccessor } from '../../../../platform/instantiation/common/instantiation.js';
import { REACH_ASSISTANT_PROMPT_LIMIT, ReachAssistantModel } from '../../../../platform/native/common/native.js';
import { IContextMenuService } from '../../../../platform/contextview/browser/contextView.js';
import { IOpenerService } from '../../../../platform/opener/common/opener.js';
import { openLinkFromMarkdown } from '../../../../platform/markdown/browser/markdownRenderer.js';
import { IQuickInputService, IQuickPickItem } from '../../../../platform/quickinput/common/quickInput.js';
import { IFileService } from '../../../../platform/files/common/files.js';
import { IRequestService } from '../../../../platform/request/common/request.js';
import { IStorageService, StorageScope, StorageTarget } from '../../../../platform/storage/common/storage.js';
import { TerminalLocation } from '../../../../platform/terminal/common/terminal.js';
import { AccessibilitySignal, IAccessibilitySignalService } from '../../../../platform/accessibilitySignal/browser/accessibilitySignalService.js';
import { IWorkspaceContextService } from '../../../../platform/workspace/common/workspace.js';
import { IWorkspaceTrustRequestService } from '../../../../platform/workspace/common/workspaceTrust.js';
import { ActiveEditorContext, ActivePanelContext, IsAuxiliaryWindowContext, IsSessionsWindowContext, MainEditorAreaVisibleContext, PanelVisibleContext } from '../../../../workbench/common/contextkeys.js';
import { IWorkbenchLayoutService, Parts } from '../../../../workbench/services/layout/browser/layoutService.js';
import { BrowserEditorInput } from '../../../../workbench/contrib/browserView/common/browserEditorInput.js';
import { Menus } from '../../../../workbench/contrib/sessions/browser/menus.js';
import { DesktopLayoutContext } from '../../../../workbench/contrib/sessions/common/contextkeys.js';
import { ITerminalGroupService, ITerminalService } from '../../../../workbench/contrib/terminal/browser/terminal.js';
import { TERMINAL_VIEW_ID } from '../../../../workbench/contrib/terminal/common/terminal.js';
import { TerminalContextKeys } from '../../../../workbench/contrib/terminal/common/terminalContextKey.js';
import { OPEN_BROWSER_COMMAND_ID } from '../../editor/browser/addTabActions.js';
import { IChatDraft } from '../../../../workbench/contrib/chat/common/attachments/chatDraft.js';
import { toRelativeRange } from '../../../../workbench/contrib/chat/common/attachments/chatInlineAttachments.js';
import { IChatRequestVariableEntry, isImageVariableEntry } from '../../../../workbench/contrib/chat/common/attachments/chatVariableEntries.js';
import { coerceImageBuffer } from '../../../../workbench/contrib/chat/common/chatImageExtraction.js';
import { PromptsConfig } from '../../../../workbench/contrib/chat/common/promptSyntax/config/config.js';
import { PromptsType } from '../../../../workbench/contrib/chat/common/promptSyntax/promptTypes.js';
import { IPromptsService, PromptsStorage } from '../../../../workbench/contrib/chat/common/promptSyntax/service/promptsService.js';
import { ISearchService, QueryType } from '../../../../workbench/services/search/common/search.js';
import { ContributionEnablementState, isContributionEnabled } from '../../../../workbench/contrib/chat/common/enablement.js';
import { IGitService } from '../../../../workbench/contrib/git/common/gitService.js';
import { McpCommandIds } from '../../../../workbench/contrib/mcp/common/mcpCommandIds.js';
import { IMcpServer, IMcpService, McpConnectionState, McpToolVisibility } from '../../../../workbench/contrib/mcp/common/mcpTypes.js';
import { signalCDirectMailConnections } from '../../../../workbench/contrib/mcp/common/signalCAppConnections.js';
import { isSignalCMailAddress, ISignalCDirectMail } from '../../../../workbench/contrib/mcp/browser/signalCDirectMail.js';
import type { IComposerSlashCommand } from '../../../../workbench/contrib/sessions/browser/chat/newChatComposerInput.js';
import { SharedChatInputComposer } from '../../../../workbench/contrib/sessions/browser/chat/sharedChatInputComposer.js';
import { AbstractCustomView, ICustomViewService } from '../../../../workbench/contrib/sessions/browser/customView.js';
import { INewSessionComposerService } from '../../../../workbench/contrib/sessions/browser/newSessionComposerService.js';
import { ASSISTANT_CUSTOM_VIEW_ID, OPEN_ASSISTANT_ACTION_ID } from '../../../../workbench/contrib/sessions/common/chatConstants.js';
import { SESSION_WORKTREE_ICON } from '../../../../workbench/contrib/sessions/common/sessionWorktreeIcon.js';
import { IEditorGroupsService } from '../../../../workbench/services/editor/common/editorGroupsService.js';
import { IEditorService, SIDE_GROUP } from '../../../../workbench/services/editor/common/editorService.js';
import { IFocusedSessionHostService } from '../../../../workbench/services/sessions/browser/sessionHost.js';
import { IUntitledTextEditorService } from '../../../../workbench/services/untitled/common/untitledTextEditorService.js';
import { UntitledTextEditorInput } from '../../../../workbench/services/untitled/common/untitledTextEditorInput.js';
import { isSessionCreated, SessionStatus } from '../../../../workbench/services/sessions/common/session.js';
import { WorkspaceSelectionOrigin } from '../../../../workbench/services/sessions/common/workspaceSelection.js';
import { IViewsService } from '../../../../workbench/services/views/common/viewsService.js';
import { AgentsWindowMenus } from '../../../browser/agentsWindowMenus.js';
import { IReachSubscriptionConnection } from '../../accountMenu/browser/reachSubscriptionConnection.js';
import { AgentSubscriptionsAction } from '../../accountMenu/browser/subscriptionActions.js';
import { assistantWorktreeError, describeAssistantAgentStatus, IAssistantAgentChats, IAssistantAgentTarget } from './assistantAgentChats.js';
import { ASSISTANT_BRIEFINGS, AssistantBriefingPeriod } from './assistantBriefings.js';
import { IAssistantBriefingScheduler, IAssistantBriefingTime, IAssistantRoutine, IAssistantRoutineSettings, parseAssistantBriefingTime } from './assistantBriefingScheduler.js';
import { ASSISTANT_SAVED_IDEA_TEXT_LIMIT, ASSISTANT_SAVED_IDEAS_LIMIT, AssistantConversationStore, AssistantPermissionMode, assistantSummaryIndex, assistantTextAttachment, IAssistantActivity, IAssistantSavedIdea, IAssistantTask, IAssistantTurn, readAssistantTaskTarget, searchAssistantSavedIdeas, storeAssistantTaskTarget } from './assistantConversation.js';
import { IAssistantMailWatcher } from './assistantMailWatcher.js';
import { memoryEntries, memoryUsage } from './assistantMemory.js';
import { IAssistantApprovalRequest, IAssistantRuns } from './assistantRuns.js';
import { listAssistantSkills } from './assistantSkills.js';
import { assistantConnectionKey, assistantToolApprovalKey, captureAssistantTaskTarget, collectAssistantTools, findAssistantAppConnection, formatAssistantWorkspacePath, getAssistantSkillScope, getAssistantWorkspaceFolders, IAssistantWorkspaceFolder, isWorkIqMcpServer, resolveAssistantWorkspacePath, safeMcpTool } from './assistantTools.js';
import { readAssistantCommonGitDirectory, readAssistantWorktreeIdentity } from './assistantWorktree.js';
import './media/assistantSession.css';

export const MANAGE_ASSISTANT_ROUTINES_ACTION_ID = 'sessionsView.manageAssistantRoutines';
export const CLEAR_ASSISTANT_CHAT_ACTION_ID = 'sessionsView.clearAssistantChat';

const ASSISTANT_IMAGE_ATTACHMENT_LIMIT = 4;
const ASSISTANT_IMAGE_ATTACHMENT_BYTES_LIMIT = 12 * 1024 * 1024;

/** Converts a composer image attachment into the vision input format used by the REACH transport. */
export function assistantImageAttachmentDataUrl(attachment: IChatRequestVariableEntry): { readonly url: string; readonly bytes: number } | undefined {
	if (!isImageVariableEntry(attachment)) { return undefined; }
	const bytes = coerceImageBuffer(attachment.value);
	if (!bytes?.byteLength) { return undefined; }
	// The bytes are authoritative: image entries often omit mimeType, and resizing may re-encode them.
	const mimeType = detectImageMimeType(bytes) ?? attachment.mimeType;
	if (!mimeType?.startsWith('image/')) { return undefined; }
	return { url: `data:${mimeType};base64,${encodeBase64(VSBuffer.wrap(bytes))}`, bytes: bytes.byteLength };
}

/** Document-like reply metadata used for the transcript's document card. */
export function assistantReplyDocument(reply: string): { readonly title?: string; readonly words: number } | undefined {
	const headings = [...reply.matchAll(/^#{1,3}\s+(.+?)\s*$/gm)].map(match => match[1].replace(/[#*_`\[\]]/g, '').trim()).filter(Boolean);
	const words = (reply.match(/\S+/g) ?? []).length;
	if (reply.length < 1200 && headings.length < 2 && words < 180) { return undefined; }
	return { title: headings[0], words };
}

function assistantPromptWithAttachmentContext(input: string, label: string, content: string, truncation: string): string {
	if (!content.trim()) { return input; }
	const emptyAttachmentLength = assistantTextAttachment(label, '').length;
	const available = REACH_ASSISTANT_PROMPT_LIMIT - input.length - emptyAttachmentLength;
	if (available < 240) { return input; }
	let excerpt = content;
	if (excerpt.length > available) {
		if (available < truncation.length + 80) { return input; }
		excerpt = `${content.slice(0, available - truncation.length - 1)}\n${truncation}`;
	}
	return assistantTextAttachment(label, excerpt) + input;
}

/** Adds as much open-document context as fits in the fixed-size prompt without pushing the user's typed request over the limit. */
export function assistantPromptWithDocumentContext(input: string, title: string, content: string): string {
	return assistantPromptWithAttachmentContext(input, localize('assistantView.documentContextAttachment', "Document · {0}", title), content, '…(document truncated; use document tools for the rest)');
}

/** Adds the visible document selection as the most specific editing target, bounded by the fixed-size prompt. */
export function assistantPromptWithSelectionContext(input: string, title: string, selection: { readonly text: string; readonly startLine: number; readonly endLine: number }): string {
	const label = selection.startLine === selection.endLine
		? localize('assistantView.selectionContextAttachment', "Selected text · {0} · line {1}", title, selection.startLine)
		: localize('assistantView.selectionContextAttachmentRange', "Selected text · {0} · lines {1}-{2}", title, selection.startLine, selection.endLine);
	return assistantPromptWithAttachmentContext(input, label, selection.text, '…(selection truncated)');
}

/** Finds the current selection in a visible editor for an open Assistant document. */
export function assistantVisibleDocumentSelection(controls: readonly unknown[], resource: string): { readonly text: string; readonly startLine: number; readonly endLine: number } | undefined {
	for (const control of controls) {
		if (!isCodeEditor(control)) { continue; }
		const model = control.getModel();
		if (model?.uri.toString() !== resource) { continue; }
		const selection = control.getSelection();
		if (!selection || selection.isEmpty()) { return undefined; }
		const text = model.getValueInRange(selection);
		return text.trim() ? { text, startLine: selection.startLineNumber, endLine: selection.endLineNumber } : undefined;
	}
	return undefined;
}

/** Reads only explicit self-naming preferences from the user profile; descriptive profile lines are not guessed as names. */
export function assistantGreetingName(userProfile: string): string | undefined {
	for (const line of memoryEntries(userProfile)) {
		const match = /^(?:my name is|name:|call me)\s+([^.,;\n]{2,40})/i.exec(line);
		const name = match?.[1].trim().replace(/["'`]+/g, '');
		if (name && /^[\p{L}\p{M}'’._-]+(?: [\p{L}\p{M}'’._-]+){0,2}$/u.test(name)) { return name; }
	}
	return undefined;
}

/** Finds an exact @routine mention. The longest names win so overlapping routine names remain unambiguous. */
export function assistantRoutineMention(text: string, routines: readonly IAssistantRoutine[]): { readonly routine: IAssistantRoutine; readonly start: number; readonly end: number } | undefined {
	const candidates = [...routines]
		.filter(routine => routine.name.trim().length > 0)
		.sort((a, b) => b.name.length - a.name.length);
	for (let at = text.indexOf('@'); at >= 0; at = text.indexOf('@', at + 1)) {
		if (at > 0 && !/[\s([{]/.test(text[at - 1])) { continue; }
		const rest = text.slice(at + 1);
		for (const routine of candidates) {
			const source = routine.name.trim().split(/\s+/).map(escapeRegExpCharacters).join('\\s+');
			const match = new RegExp(`^${source}(?![\\p{L}\\p{M}\\p{N}_])`, 'iu').exec(rest);
			if (match) { return { routine, start: at, end: at + 1 + match[0].length }; }
		}
	}
	return undefined;
}

/** Removes the typed /task command and keeps inline references aligned in the child draft. */
export function getAssistantLinkedTaskDraft(draft: IChatDraft): IChatDraft | undefined {
	const command = draft.inputText.trim();
	if (!command.startsWith('/task ')) { return undefined; }
	const inputText = command.slice('/task '.length).trimStart();
	if (!inputText) { return undefined; }
	const prefixLength = draft.inputText.trimEnd().length - inputText.length;
	const origin = TextLength.ofText(draft.inputText.slice(0, prefixLength)).toRange().getEndPosition();
	return {
		inputText,
		attachments: draft.attachments.map(attachment => attachment.range ? {
			...attachment, range: { start: attachment.range.start - prefixLength, endExclusive: attachment.range.endExclusive - prefixLength },
		} : attachment),
		references: draft.references?.map(reference => ({ ...reference, range: toRelativeRange(reference.range, origin) })),
	};
}

/** Return the exact rendered selection, or the original Markdown for a whole-reply save. */
export function assistantReplyPassage(element: HTMLElement, markdown: string): string | undefined {
	const selection = element.ownerDocument.getSelection();
	if (!selection || selection.isCollapsed || !selection.rangeCount) { return markdown; }
	const range = selection.getRangeAt(0);
	if (!range.intersectsNode(element)) { return markdown; }
	return selection.rangeCount === 1 && element.contains(range.startContainer) && element.contains(range.endContainer) ? selection.toString() : undefined;
}

function ideaExcerpt(text: string): string {
	const line = text.trim().replace(/\s+/g, ' ');
	return line.length > 80 ? line.slice(0, 79) + '…' : line;
}

async function pickSavedIdea(quickInput: IQuickInputService, ideas: readonly IAssistantSavedIdea[], query: string, token: CancellationToken): Promise<IAssistantSavedIdea | undefined> {
	const disposables = new DisposableStore();
	const picker = disposables.add(quickInput.createQuickPick<IQuickPickItem & { idea: IAssistantSavedIdea }>());
	try {
		picker.title = localize('assistantView.savedIdeas', "Saved ideas");
		picker.placeholder = localize('assistantView.savedIdeasHint', "Search passages and source chats");
		picker.sortByLabel = false;
		picker.value = query;
		const filter = (value: string) => {
			picker.items = searchAssistantSavedIdeas(ideas, value).map(idea => ({
				label: escapeIcons(ideaExcerpt(idea.text)), description: escapeIcons(idea.sourceTitle),
				detail: new Date(idea.createdAt).toLocaleString(), alwaysShow: true, idea,
			}));
		};
		filter(query);
		disposables.add(picker.onDidChangeValue(filter));
		return await new Promise<IAssistantSavedIdea | undefined>(resolve => {
			disposables.add(picker.onDidAccept(() => { resolve(picker.selectedItems[0]?.idea); picker.hide(); }));
			disposables.add(picker.onDidHide(() => resolve(undefined)));
			disposables.add(token.onCancellationRequested(() => picker.hide()));
			picker.show();
		});
	} finally {
		disposables.dispose();
	}
}

const assistantSlashCommands: readonly IComposerSlashCommand[] = [
	{ label: '/help', description: localize('assistantView.helpHelp', "Browse Assistant commands") },
	{ label: '/task', description: localize('assistantView.helpTask', "Start a linked task: /task <description>"), acceptsArguments: true },
	{ label: '/tasks', description: localize('assistantView.helpTasks', "Jump to a linked task") },
	{ label: '/agents', description: localize('assistantView.helpAgents', "Open a regular chat started by Assistant") },
	{ label: '/worktree', description: localize('assistantView.helpWorktree', "Choose an existing Git worktree for Assistant") },
	{ label: '/attach', description: localize('assistantView.helpAttach', "Attach a workspace file (picker or path)"), acceptsArguments: true },
	{ label: '/model', description: localize('assistantView.helpModel', "Choose a REACH model") },
	{ label: '/connect', description: localize('assistantView.helpConnect', "Connect an app to Assistant, for example /connect Slack"), acceptsArguments: true },
	{ label: '/send-email', description: localize('assistantView.helpSendEmail', "Send through Gmail or Outlook without using a model"), acceptsArguments: true },
	{ label: '/prompts', description: localize('assistantView.helpPrompts', "Use a prompt template from a connected app") },
	{ label: '/skill', description: localize('assistantView.helpSkill', "Choose a skill or enter /skill <name> <request>"), acceptsArguments: true },
	{ label: '/save', description: localize('assistantView.helpSaveIdea', "Save a passage from this draft with /save <text>"), acceptsArguments: true },
	{ label: '/saved', description: localize('assistantView.helpSavedIdeas', "Find a saved idea with /saved [words]"), acceptsArguments: true },
	{ label: '/resources', description: localize('assistantView.helpResources', "Attach readable content published by a connected app") },
	{ label: '/tools', description: localize('assistantView.helpTools', "List the tools Assistant can use") },
	{ label: '/briefing', description: localize('assistantView.helpBriefing', "Run or schedule a morning or evening briefing"), acceptsArguments: true },
	{ label: '/weather', description: localize('assistantView.helpWeather', "Set up The Weather Network forecasts") },
	{ label: '/routine', description: localize('assistantView.helpRoutine', "Create, run, or manage recurring Assistant routines") },
	{ label: '/learn', description: localize('assistantView.helpLearn', "Learn from connected history and review proposed memories"), acceptsArguments: true },
	{ label: '/memory', description: localize('assistantView.helpMemory', "Show saved memory; /forget edits it") },
	{ label: '/forget', description: localize('assistantView.helpForget', "Choose which saved memories to keep") },
	{ label: '/documents', description: localize('assistantView.helpDocuments', "Open or remove persistent Assistant documents") },
	{ label: '/export', description: localize('assistantView.helpExport', "Copy the transcript as Markdown") },
	{ label: '/search', description: localize('assistantView.helpSearch', "Filter the transcript — /search alone clears"), acceptsArguments: true },
	{ label: '/copy', description: localize('assistantView.helpCopy', "Copy the last reply") },
	{ label: '/retry', description: localize('assistantView.helpRetry', "Resend the last failed prompt") },
	{ label: '/regenerate', description: localize('assistantView.helpRegenerate', "Drop and resend the last turn") },
	{ label: '/undo', description: localize('assistantView.helpUndo', "Drop the last turn and restore its prompt to the composer") },
	{ label: '/compact', description: localize('assistantView.helpCompact', "Summarize and collapse this thread's history") },
	{ label: '/context', description: localize('assistantView.helpContext', "Show how much history the next request carries") },
	{ label: '/stop', description: localize('assistantView.helpStop', "Cancel the running reply (same as Escape)") },
	{ label: '/clear', description: localize('assistantView.helpClear', "Clear this conversation") },
];

export interface IAssistantRoutineTemplate {
	readonly id: string;
	readonly name: string;
	readonly description: string;
	readonly prompt: string;
	readonly cadence: IAssistantRoutineSettings['cadence'];
	readonly hour: number;
	readonly minute: number;
	readonly sourceNames: readonly string[];
}

export const assistantRoutineTemplates: readonly IAssistantRoutineTemplate[] = [
	{
		id: 'morning-briefing',
		name: localize('assistantView.templateMorningBriefing', "Morning briefing"),
		description: localize('assistantView.templateMorningBriefingDescription', "Today's schedule, urgent messages, and the few things that need you first"),
		prompt: 'Review today\'s calendar and important unread mail. Write a concise morning briefing with: first, the three most important things needing the user; second, today\'s schedule with preparation notes; third, messages that appear to need a reply; and fourth, one useful need-to-know update. Cite the calendar event or message source for each important item. Do not send, label, move, or change anything.',
		cadence: 'daily', hour: 7, minute: 30,
		sourceNames: ['calendar', 'gmail', 'outlook', 'microsoft 365'],
	},
	{
		id: 'meeting-briefing',
		name: localize('assistantView.templateMeetingBriefing', "Meeting briefing"),
		description: localize('assistantView.templateMeetingBriefingDescription', "Research and preparation notes for the meetings ahead"),
		prompt: 'Review the meetings scheduled for the next 24 hours. Skip low-signal holds and focus on meetings with other people. For each useful meeting, write a concise briefing with purpose, participants, relevant recent email or documents, open questions, and what the user should decide or prepare. Cite the event and source material. Do not change the calendar or send invitations.',
		cadence: 'weekdays', hour: 7, minute: 0,
		sourceNames: ['calendar', 'gmail', 'outlook', 'microsoft 365', 'drive', 'docs'],
	},
	{
		id: 'inbox-triage',
		name: localize('assistantView.templateInboxTriage', "Inbox triage"),
		description: localize('assistantView.templateInboxTriageDescription', "Unread mail grouped by what needs a reply, a decision, or no action"),
		prompt: 'Read unread and recent important mail. Group it into: needs a reply today, needs a decision, waiting on someone else, newsletters or low priority, and clearly no action. For each important message include sender, subject, why it matters, and a suggested next step. Cite message IDs or links. Do not label, archive, draft, or send mail.',
		cadence: 'weekdays', hour: 8, minute: 0,
		sourceNames: ['gmail', 'outlook', 'microsoft 365'],
	},
	{
		id: 'github-report',
		name: localize('assistantView.templateGithubReport', "GitHub report"),
		description: localize('assistantView.templateGithubReportDescription', "Plain-English shipped, blocked, and review activity from open repositories"),
		prompt: 'Review recent GitHub activity relevant to the user. Summarize what shipped, what is blocked, pull requests waiting on the user, important review comments, and notable failures or incidents. Group related work and explain it in plain English for a team status update. Include repository and pull request links. Do not comment, merge, close, or change anything.',
		cadence: 'weekdays', hour: 16, minute: 0,
		sourceNames: ['github'],
	},
	{
		id: 'newsletter-digest',
		name: localize('assistantView.templateNewsletterDigest', "Newsletter digest"),
		description: localize('assistantView.templateNewsletterDigestDescription', "A concise digest of newsletters and low-urgency subscriptions"),
		prompt: 'Find recent newsletters and subscription-style mail. Write a concise digest grouped by topic, with the few items actually worth reading, one-line takeaways, and links or message IDs. Omit repetitive promotions and say when nothing important was found. Do not label, archive, delete, or unsubscribe anything.',
		cadence: 'daily', hour: 17, minute: 30,
		sourceNames: ['gmail', 'outlook', 'microsoft 365'],
	},
	{
		id: 'relationship-reconnect',
		name: localize('assistantView.templateReconnect', "Relationship reconnect"),
		description: localize('assistantView.templateReconnectDescription', "People worth reconnecting with based on recent contact history"),
		prompt: 'Review recent mail and calendar activity for people the user has meaningfully interacted with but not contacted recently. Suggest up to five thoughtful reconnections, explaining the relationship context, why now is useful, and a short natural opening line. Do not send messages or create calendar events.',
		cadence: 'weekdays', hour: 9, minute: 0,
		sourceNames: ['calendar', 'gmail', 'outlook', 'microsoft 365', 'contacts'],
	},
];

export function assistantRoutineTemplateServerIds(template: IAssistantRoutineTemplate, sources: readonly { readonly id: string; readonly label: string }[]): readonly string[] {
	const wanted = template.sourceNames.map(name => name.toLowerCase());
	return sources.filter(source => {
		const label = source.label.toLowerCase();
		return wanted.some(name => label.includes(name));
	}).map(source => source.id);
}

const formatScheduleTime = ({ hour, minute }: IAssistantBriefingTime) => `${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}`;

function describeAssistantRunStatus(status: NonNullable<IAssistantRoutine['lastRun']>['status']): string {
	const statuses = {
		running: localize('assistantView.routineRunning', "Running"),
		completed: localize('assistantView.routineReportReady', "Report ready"),
		needsAttention: localize('assistantView.routineNeedsAttention', "Needs attention"),
		failed: localize('assistantView.routineRunFailed', "Failed"),
		stopped: localize('assistantView.routineRunStopped', "Stopped"),
		interrupted: localize('assistantView.routineRunInterrupted', "Interrupted"),
	};
	return statuses[status];
}

function describeLastRoutineRun(routine: IAssistantRoutine): string {
	const run = routine.lastRun;
	if (!run) { return localize('assistantView.routineNeverRun', "No recorded runs"); }
	const trigger = run.trigger === 'scheduled'
		? localize('assistantView.routineScheduledRun', "Scheduled run")
		: localize('assistantView.routineManualRun', "Manual run");
	return localize('assistantView.routineLastRun', "{0} · {1} · {2}", describeAssistantRunStatus(run.status), trigger, new Date(run.startedAt).toLocaleString(undefined, { timeZone: routine.timeZone }));
}

class AssistantView extends AbstractCustomView {
	override readonly maxWidth = Number.MAX_SAFE_INTEGER;
	readonly title = constObservable(localize('assistantView.title', "Assistant"));
	private readonly canSend = observableValue(this, false);
	private accountCheck = 0;
	private content: HTMLElement | undefined;
	private composerHost: HTMLElement | undefined;
	private composer: SharedChatInputComposer | undefined;
	private saveDraft: (() => void) | undefined;

	constructor(
		private readonly routinesRequested: ISettableObservable<boolean>,
		private readonly onDidRequestClearChat: Event<void>,
		@IReachSubscriptionConnection private readonly reach: IReachSubscriptionConnection,
		@ICommandService private readonly commands: ICommandService,
		@IStorageService private readonly storage: IStorageService,
		@IClipboardService private readonly clipboard: IClipboardService,
		@IInstantiationService private readonly instantiation: IInstantiationService,
		@IQuickInputService private readonly quickInput: IQuickInputService,
		@IContextMenuService private readonly contextMenu: IContextMenuService,
		@IOpenerService private readonly openerService: IOpenerService,
		@IAssistantRuns private readonly runs: IAssistantRuns,
		@IAssistantAgentChats private readonly agentChats: IAssistantAgentChats,
		@IMcpService private readonly mcpService: IMcpService,
		@IDialogService private readonly dialogs: IDialogService,
		@IFileDialogService private readonly fileDialogs: IFileDialogService,
		@IRequestService private readonly requestService: IRequestService,
		@IFileService private readonly fileService: IFileService,
		@IWorkspaceContextService private readonly workspaceService: IWorkspaceContextService,
		@ISearchService private readonly searchService: ISearchService,
		@IAccessibilitySignalService private readonly signalService: IAccessibilitySignalService,
		@IGitService private readonly gitService: IGitService,
		@IWorkspaceTrustRequestService private readonly workspaceTrust: IWorkspaceTrustRequestService,
		@ICustomViewService private readonly customViews: ICustomViewService,
		@ISignalCDirectMail private readonly directMail: ISignalCDirectMail,
		@IConfigurationService private readonly configuration: IConfigurationService,
		@IPromptsService private readonly promptsService: IPromptsService,
		@IEditorService private readonly editorService: IEditorService,
		@IUntitledTextEditorService private readonly untitledTextEditors: IUntitledTextEditorService,
	) {
		super();
	}

	private async pickRoutineSettings(current: () => boolean, defaults?: Partial<IAssistantRoutineSettings>, editing = false): Promise<IAssistantRoutineSettings | undefined> {
		const title = editing ? localize('assistantView.editRoutine', "Edit routine") : localize('assistantView.newRoutine', "New routine");
		const name = (await this.quickInput.input({
			title,
			value: defaults?.name,
			prompt: localize('assistantView.nameRoutine', "Name this routine"),
			validateInput: async value => value.trim().length > 0 && value.trim().length <= 80 ? undefined
				: localize('assistantView.routineNameLength', "Enter a name of 1 to 80 characters."),
		}))?.trim();
		if (!name || !current()) { return undefined; }
		const prompt = (await this.quickInput.input({
			title,
			value: defaults?.prompt,
			prompt: localize('assistantView.routineInstructions', "What should Assistant check and report?"),
			validateInput: async value => value.trim().length > 0 && value.trim().length <= 2000 ? undefined
				: localize('assistantView.routineInstructionsLength', "Enter instructions of 1 to 2,000 characters."),
		}))?.trim();
		if (!prompt || !current()) { return undefined; }
		const senders = this.directMail.connections.get().filter(mail => mail.canSend && !mail.needsSignIn)
			.map(mail => ({ id: mail.id, label: mail.label }));
		const previousSender = defaults?.emailDelivery?.connectionId;
		if (previousSender && !senders.some(sender => sender.id === previousSender)) {
			senders.push({ id: previousSender, label: localize('assistantView.savedSender', "Saved mailbox (reconnect to send)") });
		}
		let emailDelivery: IAssistantRoutineSettings['emailDelivery'];
		if (senders.length) {
			const choices = [
				{ label: localize('assistantView.routineInAssistant', "Show the result in Assistant"), email: false },
				{ label: localize('assistantView.routineEmailDelivery', "Email the result after each completed run"), email: true },
			];
			const delivery = await this.quickInput.pick(choices, {
				title: localize('assistantView.routineDelivery', "Routine delivery"),
				activeItem: choices[previousSender ? 1 : 0],
			});
			if (!delivery || !current()) { return undefined; }
			if (delivery.email) {
				const sender = await this.quickInput.pick(senders, {
					title: localize('assistantView.routineSender', "Choose the sending mailbox"),
					activeItem: senders.find(item => item.id === previousSender),
				});
				if (!sender || !current()) { return undefined; }
				const recipient = (await this.quickInput.input({
					title: localize('assistantView.routineRecipient', "Routine email recipient"),
					value: defaults?.emailDelivery?.recipient,
					prompt: localize('assistantView.routineRecipientPrompt', "Enter one address to receive each completed routine. The routine name is the subject."),
					validateInput: async value => isSignalCMailAddress(value.trim()) ? undefined
						: localize('assistantView.routineRecipientInvalid', "Enter one complete email address."),
				}))?.trim();
				if (!recipient || !current()) { return undefined; }
				emailDelivery = { connectionId: sender.id, recipient };
			}
		}
		const sources = this.mcpService.servers.get()
			.filter(server => isContributionEnabled(server.enablement.get()))
			.map(server => {
				const id = assistantConnectionKey(server);
				return { id, label: server.definition.label, description: server.collection.label, picked: defaults?.serverIds?.includes(id) };
			});
		for (const mail of this.directMail.connections.get()) {
			sources.push({
				id: mail.id, label: mail.label, picked: defaults?.serverIds?.includes(mail.id),
				description: mail.needsSignIn ? localize('assistantView.directMailSignIn', "Sign-in required") : localize('assistantView.directMailReadOnly', "Read-only"),
			});
		}
		const selected = sources.length ? await this.quickInput.pick(sources, {
			canPickMany: true,
			title: localize('assistantView.routineSources', "Choose connections for this routine"),
			placeHolder: emailDelivery
				? localize('assistantView.routineOptionalSourcesHint', "Choose read-only sources for the email, or continue with none.")
				: localize('assistantView.routineSourcesHint', "Choose one or more connections. Each run uses read-only tools."),
		}) : [];
		if (!selected || (!selected.length && !emailDelivery) || !current()) { return undefined; }
		const days = [
			{ label: localize('assistantView.routineDaily', "Daily"), cadence: 'daily' as const },
			{ label: localize('assistantView.routineWeekdays', "Weekdays"), cadence: 'weekdays' as const },
		];
		const schedule = await this.quickInput.pick(days, {
			title: localize('assistantView.routineFrequency', "How often should it run?"),
			activeItem: days.find(day => day.cadence === defaults?.cadence),
		});
		if (!schedule || !current()) { return undefined; }
		const timeZone = Intl.DateTimeFormat().resolvedOptions().timeZone;
		const value = await this.quickInput.input({
			title: localize('assistantView.routineTime', "Confirm routine schedule"),
			value: formatScheduleTime({ hour: defaults?.hour ?? 8, minute: defaults?.minute ?? 0 }),
			prompt: emailDelivery
				? localize('assistantView.routineEmailTimePrompt', "Confirm a 24-hour time in {0} to save automatic delivery to {1}. Routines run while Agents is open.", timeZone, emailDelivery.recipient)
				: localize('assistantView.routineTimePrompt', "Confirm a 24-hour time in {0} to save. Routines run while Agents is open.", timeZone),
			validateInput: async value => parseAssistantBriefingTime(value) ? undefined : localize('assistantView.briefingInvalidTime', "Enter a 24-hour time, for example 08:00."),
		});
		const time = value === undefined ? undefined : parseAssistantBriefingTime(value);
		return time && current() ? { ...time, name, prompt, cadence: schedule.cadence, serverIds: selected.map(source => source.id), emailDelivery } : undefined;
	}

	render(container: HTMLElement): void {
		const content = this.content = append(container, $('.reach-assistant'));
		const gate = append(content, $('p.reach-assistant-gate', { role: 'status' }));
		if (!this.reach.assistantSupported) {
			gate.textContent = localize('assistantView.desktopOnly', "Open Signal-C desktop to use Assistant with your REACH Subscription.");
			return;
		}
		gate.textContent = localize('assistantView.checkingSubscription', "Checking REACH Subscription…");
		const manage = append(content, $('button.reach-assistant-link', { type: 'button', hidden: true }, localize('assistantView.manageSubscription', "Manage REACH Subscription")));
		this._register(addDisposableListener(manage, EventType.CLICK, () => void this.commands.executeCommand(AgentSubscriptionsAction.ID)));

		const store = new AssistantConversationStore(this.storage);
		const openedTarget = this.agentChats.target;
		let rootTarget = openedTarget;
		let rootProject = openedTarget?.workspace;
		const restoreProject = !openedTarget && this.customViews.activeCustomViewOpen.get()?.source === 'restore';
		let restoreTaskOnSignIn = restoreProject;
		const createTask = (account: string) => store.createTask(account, captureAssistantTaskTarget(this.workspaceService, this.agentChats.target));
		const collectThreadTools = (accountId: string, taskId: string | undefined) => collectAssistantTools({
			accountId, taskId, store,
			requestService: this.requestService, mcpService: this.mcpService, directMail: this.directMail, fileService: this.fileService,
			readNewsFeeds: scope => this.reach.readAssistantNewsFeeds(scope),
			readWeatherForecast: (day, location) => this.reach.readAssistantWeatherForecast(day, location),
			configuration: this.configuration,
			promptsService: this.configuration.getValue<boolean>(PromptsConfig.USE_AGENT_SKILLS) ? this.promptsService : undefined,
			workspaceService: this.workspaceService, searchService: this.searchService, openerService: this.openerService,
			agentChats: this.agentChats, agentChatTarget: this.agentChats.target, connectApp: appId => this.runs.openAppSetup(appId),
		});
		const chat = append(content, $('.reach-assistant-chat', { hidden: true }));
		const breadcrumb = append(chat, $('.reach-assistant-breadcrumb', { hidden: true }));
		const back = append(breadcrumb, $('button.reach-assistant-link', { type: 'button' }, localize('assistantView.back', "‹ Assistant")));
		const taskName = append(breadcrumb, $('span'));
		const apps = append(chat, $('.reach-assistant-apps', { 'aria-label': localize('assistantView.apps', "Connected apps"), hidden: true }));
		const thread = append(chat, $('.reach-assistant-thread'));
		const welcome = append(thread, $('.reach-assistant-welcome'));
		append(welcome, $('span' + ThemeIcon.asCSSSelector(Codicon.sparkle), { 'aria-hidden': 'true' }));
		const welcomeTitle = append(welcome, $('h2'));
		const welcomeDetail = append(welcome, $('p.reach-assistant-muted.reach-assistant-greeting-detail', { hidden: true }));
		const suggestions = append(welcome, $('.reach-assistant-suggestions'));
		append(welcome, $('p.reach-assistant-muted.reach-assistant-commands', {}, localize('assistantView.commands', "Commands: {0}", assistantSlashCommands.map(command => command.label).join(' · '))));
		const taskUpdates = append(thread, $('section.reach-assistant-task-updates', { 'aria-label': localize('assistantView.tasksAria', "Linked tasks"), hidden: true }));
		const transcript = append(thread, $('.reach-assistant-transcript', { role: 'log', 'aria-label': localize('assistantView.transcript', "Assistant conversation") }));
		const followups = append(thread, $('.reach-assistant-suggestions.reach-assistant-followups', { hidden: true }));
		const pending = append(thread, $('article.reach-assistant-pending', { hidden: true }));
		const pendingPrompt = append(pending, $('p.reach-assistant-message'));
		const pendingReply = append(pending, $('div.reach-assistant-streaming'));
		const pendingStatus = append(pending, $('p.reach-assistant-muted', { role: 'status' }));
		const pendingPlan = append(pending, $('ul.reach-assistant-plan', { hidden: true, 'aria-label': localize('assistantView.plan', "Plan") }));
		const pendingActivity = append(pending, $('.reach-assistant-activity'));
		const approval = append(pending, $('.reach-assistant-approval', { hidden: true, role: 'alert' }));
		const approvalText = append(approval, $('p.reach-assistant-approval-text'));
		const approvalDetails = append(approval, $('details.reach-assistant-approval-details'));
		append(approvalDetails, $('summary', {}, localize('assistantView.approvalDetails', "Arguments")));
		const approvalArgs = append(approvalDetails, $('pre.reach-assistant-approval-args'));
		const approvalButtons = append(approval, $('.reach-assistant-approval-buttons'));
		const approvalAllow = append(approvalButtons, $('button.reach-assistant-control', { type: 'button', title: localize('assistantView.allowOnceTip', "Run this action one time") }, localize('assistantView.allowOnce', "Allow once")));
		const approvalAlways = append(approvalButtons, $('button.reach-assistant-control', { type: 'button', title: localize('assistantView.alwaysAllowTip', "Always run this tool in this thread without asking again") }, localize('assistantView.alwaysAllow', "Always allow")));
		const approvalReject = append(approvalButtons, $('button.reach-assistant-control', { type: 'button', title: localize('assistantView.rejectTip', "Decline this action; the assistant continues without it") }, localize('assistantView.reject', "Reject")));
		const approvalRejectNote = append(approvalButtons, $('button.reach-assistant-link', { type: 'button', title: localize('assistantView.rejectNoteTip', "Decline and tell the assistant why") }, localize('assistantView.rejectNote', "Reject with note…")));
		const approvalTrust = append(approvalButtons, $('button.reach-assistant-link', { type: 'button', hidden: true }));

		const dock = append(chat, $('.reach-assistant-dock'));
		const queuedRow = append(dock, $('p.reach-assistant-queued', { hidden: true }));
		append(queuedRow, $('span' + ThemeIcon.asCSSSelector(Codicon.comment), { 'aria-hidden': 'true' }));
		const queuedText = append(queuedRow, $('span'));
		const sendQueuedLabel = localize('assistantView.sendQueued', "Send queued");
		const retryQueuedLabel = localize('assistantView.retryQueued', "Retry queued message");
		const queuedSend = append(queuedRow, $('button.reach-assistant-link', { type: 'button' }, sendQueuedLabel));
		const queuedEdit = append(queuedRow, $('button.reach-assistant-link', { type: 'button' }, localize('assistantView.editQueue', "Edit")));
		const queuedCancel = append(queuedRow, $('button.reach-assistant-link', { type: 'button' }, localize('assistantView.unqueue', "Remove from queue")));
		const composerHost = this.composerHost = append(dock, $('.reach-assistant-composer'));
		const controls = $('.reach-assistant-controls', { role: 'group', 'aria-label': localize('assistantView.composerControls', "Assistant controls") });
		const permissionButton = append(controls, $('button.reach-assistant-control.reach-assistant-permission', { type: 'button', 'aria-haspopup': 'listbox' }));
		append(permissionButton, $('span' + ThemeIcon.asCSSSelector(Codicon.shield), { 'aria-hidden': 'true' }));
		const permissionLabel = append(permissionButton, $('span.reach-assistant-control-label'));
		const modelButton = append(controls, $('button.reach-assistant-control.reach-assistant-model', { type: 'button', 'aria-haspopup': 'listbox' }));
		append(modelButton, $('span.reach-assistant-model-icon' + ThemeIcon.asCSSSelector(Codicon.sparkle), { 'aria-hidden': 'true' }));
		const modelLabel = append(modelButton, $('span.reach-assistant-control-label'));
		append(modelButton, $('span' + ThemeIcon.asCSSSelector(Codicon.chevronDown), { 'aria-hidden': 'true' }));
		const worktreeButton = append(controls, $<HTMLButtonElement>('button.reach-assistant-control.reach-assistant-worktree', { type: 'button', 'aria-haspopup': 'listbox' }));
		append(worktreeButton, $('span' + ThemeIcon.asCSSSelector(SESSION_WORKTREE_ICON), { 'aria-hidden': 'true' }));
		const worktreeLabel = append(worktreeButton, $('span.reach-assistant-control-label'));
		const newTaskLabel = localize('assistantView.newTask', "New task");
		const resumeTaskLabel = localize('assistantView.resumeTask', "Resume task");
		const newTask = append(controls, $('button.reach-assistant-control', { type: 'button', 'aria-label': newTaskLabel }));
		newTask.title = localize('assistantView.newTaskTip', "Open a linked task conversation, or send /task followed by a task description.");
		append(newTask, $('span' + ThemeIcon.asCSSSelector(Codicon.add), { 'aria-hidden': 'true' }));
		append(newTask, $('span.reach-assistant-control-label', {}, newTaskLabel));
		const agentChatsLabel = localize('assistantView.agentChats', "Agent Chats");
		const agentChatsButton = append(controls, $('button.reach-assistant-control.reach-assistant-agent-chats', { type: 'button', 'aria-haspopup': 'listbox', 'aria-label': agentChatsLabel, title: agentChatsLabel }));
		append(agentChatsButton, $('span' + ThemeIcon.asCSSSelector(Codicon.commentDiscussion), { 'aria-hidden': 'true' }));
		append(agentChatsButton, $('span.reach-assistant-control-label', {}, agentChatsLabel));
		const savedIdeasLabel = localize('assistantView.savedIdeas', "Saved ideas");
		const savedIdeasButton = append(controls, $('button.reach-assistant-control', { type: 'button', 'aria-haspopup': 'listbox', 'aria-label': savedIdeasLabel, title: savedIdeasLabel }));
		append(savedIdeasButton, $('span' + ThemeIcon.asCSSSelector(Codicon.bookmark), { 'aria-hidden': 'true' }));
		append(savedIdeasButton, $('span.reach-assistant-control-label', {}, savedIdeasLabel));
		const feedback = append(dock, $('p.reach-assistant-feedback', { role: 'status', hidden: true }));
		const editingRow = append(dock, $('p.reach-assistant-editing', { role: 'status', hidden: true }));
		append(editingRow, $('span' + ThemeIcon.asCSSSelector(Codicon.edit), { 'aria-hidden': 'true' }));
		const editingText = append(editingRow, $('span'));
		const editingCancel = append(editingRow, $('button.reach-assistant-link', { type: 'button' }, localize('assistantView.cancelEdit', "Cancel edit")));
		const counter = append(dock, $('span.reach-assistant-counter', { role: 'status', hidden: true }));

		const threadWatch = this._register(new MutableDisposable<DisposableStore>());
		const savedIdeasPicker = this._register(new MutableDisposable<DisposableStore>());
		const renderedTurns = this._register(new DisposableStore());
		const taskListeners = this._register(new DisposableStore());
		const suggestionListeners = this._register(new DisposableStore());
		const connectionListeners = this._register(new DisposableStore());
		const activityListeners = this._register(new DisposableStore());
		const elapsedTimer = this._register(new IntervalTimer());
		let accountId: string | undefined;
		let taskId: string | undefined;
		const resumeTask = (task: IAssistantTask) => {
			const prompt = task.lastPrompt ?? task.title;
			if (!accountId || !prompt || this.runs.getPending(accountId, task.id)) { return; }
			const conversation = store.read(accountId, task.id);
			if (conversation.startedQueuedPrompt) { openThread(task.id); return; }
			if (taskId === task.id) { this.saveDraft?.(); }
			const { draft, draftContext } = store.read(accountId, task.id);
			const unsentContext = !!draftContext && store.readComposerDraft(accountId, task.id).inputText === draft;
			void this.runs.send(accountId, prompt, task.id, { preserveDraft: unsentContext || (!!draft?.trim() && draft.trim() !== prompt.trim()) });
		};
		let models: readonly ReachAssistantModel[] = [];
		let turnCount = -1;
		let lastRenderedTurn: IAssistantTurn | undefined;
		let searchQuery = '';
		let wasPending = false;
		let settingInput = false;
		let draftRevision = 0;
		let activeApproval: IAssistantApprovalRequest | undefined;
		const requestInProgress = observableValue('assistantRequestInProgress', false);
		const stopRequest = () => {
			if (accountId) { this.runs.stop(accountId, taskId); }
		};
		const lastTaskStatuses = new Map<string, string>();
		let editingFrom: { index: number; original: readonly IAssistantTurn[]; accountId: string; taskId: string | undefined } | undefined;
		const setEditingFrom = (value: typeof editingFrom) => {
			editingFrom = value;
			editingRow.hidden = !value;
			if (value) { editingText.textContent = localize('assistantView.editingBanner', "Editing earlier prompt — sending replaces it and all later replies"); }
		};

		let feedbackAction: (() => void) | undefined;
		this._register(addDisposableListener(feedback, EventType.CLICK, event => {
			if ((event.target as HTMLElement).classList.contains('reach-assistant-feedback-action')) {
				event.preventDefault();
				feedbackAction?.();
			}
		}));
		const setFeedback = (message: string, action?: { label: string; run: () => void }) => {
			clearNode(feedback);
			feedbackAction = action?.run;
			if (message) { append(feedback, $('span', {}, message)); }
			if (action) {
				const link = append(feedback, $('a.reach-assistant-feedback-action', { href: '#' }, action.label));
				link.setAttribute('role', 'button');
			}
			feedback.hidden = !message && !action;
		};
		let clearChatDialogOpen = false;
		const clearChat = async () => {
			const owner = accountId;
			const id = taskId;
			if (!owner || clearChatDialogOpen) { return; }
			if (this.runs.getPendingDetail(owner, id)) {
				setFeedback(localize('assistantView.clearChatBusy', "Wait for the current reply before clearing this chat."));
				return;
			}
			const turns = store.read(owner, id).turns;
			if (!turns.length) {
				setFeedback(localize('assistantView.clearChatEmpty', "This chat is already empty."));
				return;
			}
			clearChatDialogOpen = true;
			try {
				const { confirmed } = await this.dialogs.confirm({
					custom: true,
					type: 'warning',
					message: localize('assistantView.clearChatConfirm', "Are you sure you want to clear this chat?"),
					detail: localize('assistantView.clearChatDetail', "This permanently removes the messages in the current chat."),
					primaryButton: localize('assistantView.clearChatButton', "Clear Chat"),
				});
				if (!confirmed || this._store.isDisposed || accountId !== owner || taskId !== id) { return; }
				const current = store.read(owner, id).turns;
				if (this.runs.getPendingDetail(owner, id) || current.length !== turns.length || current.at(-1)?.at !== turns.at(-1)?.at) {
					setFeedback(localize('assistantView.clearChatChanged', "This chat changed while the dialog was open. Open Clear Chat again to review."));
					return;
				}
				searchQuery = '';
				setEditingFrom(undefined);
				store.clearTurns(owner, id);
				setFeedback(localize('assistantView.cleared', "Conversation cleared."));
			} finally {
				clearChatDialogOpen = false;
			}
		};
		this._register(this.onDidRequestClearChat(() => void clearChat()));
		const worktreeTitle = localize('assistantView.worktreePicker', "Choose Assistant worktree");
		const getWorktreeProject = () => {
			const target = this.agentChats.target;
			return target?.workspace ?? (target?.fileRoots ? target.fileRoots[0] : this.workspaceService.getWorkspace().folders[0]?.uri);
		};
		const worktreeError = async (target: IAssistantAgentTarget | undefined) => target?.worktree && assistantWorktreeError(await this.agentChats.getWorktreeStatus(target.worktree));
		const attachmentFolders = async () => {
			const target = this.agentChats.target;
			const thread = taskId;
			const error = await worktreeError(target);
			if (this._store.isDisposed || target !== this.agentChats.target || thread !== taskId) { return undefined; }
			if (error) { setFeedback(error); return undefined; }
			return getAssistantWorkspaceFolders(this.workspaceService, target);
		};
		const updateWorktreeLabel = () => {
			const selected = this.agentChats.target?.worktree;
			worktreeLabel.textContent = selected?.label ?? localize('assistantView.worktree', "Worktree");
			worktreeButton.title = selected ? `${worktreeTitle}\n${selected.uri.fsPath}` : worktreeTitle;
			worktreeButton.setAttribute('aria-label', selected ? localize('assistantView.activeWorktree', "Choose Assistant worktree. Current worktree: {0}", selected.label) : worktreeTitle);
		};
		const pickWorktree = async () => {
			if (!accountId || worktreeButton.disabled) { return; }
			const target = this.agentChats.target;
			const project = getWorktreeProject();
			const localProject = project?.scheme === Schemas.file ? project : undefined;
			const currentWorktree = localProject ? target?.worktree : undefined;
			const account = accountId;
			const task = taskId;
			const taskRunning = () => !!task && !!this.runs.getPending(account, task);
			if (taskRunning()) {
				setFeedback(localize('assistantView.worktreeTaskRunning', "Stop the task or wait for its reply before changing its worktree."));
				return;
			}
			const stillCurrent = () => !this._store.isDisposed && accountId === account && taskId === task && this.agentChats.target === target && !taskRunning();
			worktreeButton.disabled = true;
			try {
				const workspace = localProject ?? (await this.fileDialogs.showOpenDialog({
					title: localize('assistantView.worktreeLocalProject', "Choose a local Git project"),
					canSelectFiles: false,
					canSelectFolders: true,
					canSelectMany: false,
					availableFileSystems: [Schemas.file],
				}))?.[0];
				if (workspace?.scheme !== Schemas.file || !stillCurrent()) { return; }
				setFeedback(localize('assistantView.loadingWorktrees', "Finding Git worktrees…"));
				const repo = await this.gitService.openRepository(workspace);
				if (!stillCurrent()) { return; }
				if (!repo && !currentWorktree) {
					setFeedback(localize('assistantView.noGitRepository', "Couldn’t open the Git repository for this project."));
					return;
				}
				if (repo) { await this.commands.executeCommand('git.refresh', repo.rootUri); }
				const state = repo && await this.commands.executeCommand<{ worktrees: readonly { name: string; path: string; ref: string; detached: boolean; main: boolean }[] } | null>('git.api.getRepositoryState', repo.rootUri.toString());
				const commonDirectory = repo && await readAssistantCommonGitDirectory(this.fileService, repo.rootUri);
				const worktrees = await Promise.all((state?.worktrees ?? []).map(async worktree => {
					if (!commonDirectory) { return undefined; }
					const uri = URI.file(worktree.path);
					const identity = await readAssistantWorktreeIdentity(this.fileService, uri);
					const expected = worktree.main ? commonDirectory : URI.joinPath(commonDirectory, 'worktrees', worktree.name);
					if (!identity || !extUriBiasedIgnorePathCase.isEqual(identity.gitDirectory, expected)) { return undefined; }
					const label = worktree.detached ? localize('assistantView.detachedWorktree', "Detached HEAD ({0})", worktree.ref.slice(0, 7)) : worktree.ref.replace(/^refs\/heads\//, '') || basename(uri);
					return { uri, label, identity, main: worktree.main };
				}));
				if (!stillCurrent()) { return; }
				const choices = worktrees.filter(worktree => worktree !== undefined);
				if (!choices.length && !currentWorktree) {
					setFeedback(localize('assistantView.noWorktrees', "No Git worktrees were found for this project."));
					return;
				}
				setFeedback('');
				const items: (IQuickPickItem & { worktree?: IAssistantAgentTarget['worktree'] })[] = choices.map(worktree => ({
					label: escapeIcons(worktree.label),
					description: extUriBiasedIgnorePathCase.isEqual(worktree.uri, currentWorktree?.uri) ? localize('assistantView.selectedWorktree', "Selected") : worktree.main ? localize('assistantView.mainWorktree', "Main checkout") : undefined,
					detail: escapeIcons(worktree.uri.fsPath),
					worktree,
				}));
				if (currentWorktree) {
					items.push({ label: localize('assistantView.useProjectDefaults', "Use project defaults"), detail: localize('assistantView.clearWorktreeDetail', "Clear the explicit worktree selection") });
				}
				const selected = await this.quickInput.pick(items, { title: worktreeTitle, placeHolder: localize('assistantView.worktreeSearch', "Filter by branch or path"), matchOnDescription: true, matchOnDetail: true });
				if (!selected || !stillCurrent()) { return; }
				const worktree = selected.worktree;
				if (worktree) {
					const trusted = await this.workspaceTrust.requestResourcesTrust({ uri: worktree.uri });
					if (!trusted || !stillCurrent()) { return; }
					const error = assistantWorktreeError(await this.agentChats.getWorktreeStatus(worktree));
					if (!stillCurrent()) { return; }
					if (error) { setFeedback(error); return; }
				}
				const nextTarget = this.agentChats.target = { ...(localProject ? target : undefined), workspace: localProject ?? repo?.rootUri ?? workspace, worktree };
				if (task) {
					store.updateTask(account, task, { target: storeAssistantTaskTarget(captureAssistantTaskTarget(this.workspaceService, nextTarget)) });
				} else {
					rootTarget = nextTarget;
					rootProject = nextTarget.workspace;
					store.setWorktreeSelection(account, rootProject, worktree);
					store.setLastWorktreeProject(account, rootProject);
				}
				updateWorktreeLabel();
				renderSuggestions();
				setFeedback(worktree
					? localize('assistantView.worktreePicked', "New Assistant requests will use {0}.", worktree.uri.fsPath)
					: localize('assistantView.worktreeCleared', "New Assistant requests will use project defaults."));
			} catch {
				if (stillCurrent()) { setFeedback(localize('assistantView.listWorktreesFailed', "Couldn’t list Git worktrees for this project.")); }
			} finally {
				worktreeButton.disabled = false;
			}
		};
		const showAgentChats = async () => {
			const currentAccount = accountId;
			const check = this.accountCheck;
			if (!currentAccount) {
				return;
			}
			const chats = this.agentChats.getChats(currentAccount).reverse().sort((a, b) => Number(b.status === SessionStatus.NeedsInput) - Number(a.status === SessionStatus.NeedsInput));
			if (!chats.length) {
				setFeedback(localize('assistantView.noAgentChats', "No agent chats are available. Ask Assistant to start a regular agent chat."));
				return;
			}
			const pick = await this.quickInput.pick(chats.map(chat => ({
				label: escapeIcons(chat.title),
				description: describeAssistantAgentStatus(chat.status),
				detail: chat.workspaceLabel && escapeIcons(chat.workspaceLabel),
				id: chat.id,
			})), { title: localize('assistantView.agentChats', "Agent Chats"), placeHolder: localize('assistantView.pickAgentChat', "Choose a chat to open"), matchOnDescription: true, matchOnDetail: true });
			if (!pick || this.accountCheck !== check || accountId !== currentAccount) {
				return;
			}
			this.saveDraft?.();
			const opened = await this.agentChats.open(currentAccount, pick.id).catch(() => false);
			if (!opened && this.accountCheck === check && accountId === currentAccount) {
				setFeedback(localize('assistantView.openAgentChatFailed', "This agent chat couldn’t be opened. Choose another chat or check the Chats list."));
			}
		};
		const updateCounter = () => {
			const length = this.composer?.getInputValue().length ?? 0;
			counter.hidden = length <= REACH_ASSISTANT_PROMPT_LIMIT * 0.8;
			if (!counter.hidden) {
				counter.textContent = localize('assistantView.counter', "{0} / {1}", length.toLocaleString(), REACH_ASSISTANT_PROMPT_LIMIT.toLocaleString());
				counter.classList.toggle('over', length > REACH_ASSISTANT_PROMPT_LIMIT);
			}
		};
		const setInput = (text: string) => {
			draftRevision++;
			settingInput = true;
			try { this.composer?.setInputValue(text); } finally { settingInput = false; }
			updateCounter();
		};
		const setComposerDraft = (draft: IChatDraft = { inputText: '', attachments: [] }) => {
			draftRevision++;
			settingInput = true;
			try { this.composer?.setDraft(draft); } finally { settingInput = false; }
			updateCounter();
		};
		this.saveDraft = () => {
			if (!accountId || !this.composer) { return; }
			const draft = this.composer.getDraft();
			const pending = this.runs.getPending(accountId, taskId);
			if (!draft.inputText && !draft.attachments.length && !draft.references?.length && pending && store.read(accountId, taskId).draft?.trim() === pending) { return; }
			store.setComposerDraft(accountId, draft, taskId);
		};
		const saveDraft = this._register(new RunOnceScheduler(() => this.saveDraft?.(), 250));
		this._register(this.storage.onWillSaveState(() => this.saveDraft?.()));
		let verifiedAccountCheck = this.accountCheck;
		const saveIdea = (owner: string, thread: string | undefined, text: string): boolean => {
			if (this.accountCheck !== verifiedAccountCheck) { return false; }
			const source = thread ? store.read(owner).tasks?.find(task => task.id === thread)?.title || localize('assistantView.newTask', "New task") : localize('assistantView.title', "Assistant");
			const result = store.saveIdea(owner, text, source, thread);
			switch (result.status) {
				case 'saved': setFeedback(localize('assistantView.ideaSaved', "Saved to Saved ideas.")); return true;
				case 'duplicate': setFeedback(localize('assistantView.ideaDuplicate', "This passage is already in Saved ideas.")); return true;
				case 'empty': setFeedback(localize('assistantView.ideaEmpty', "Enter text after /save, or select a passage in a reply.")); return false;
				case 'tooLong': setFeedback(localize('assistantView.ideaLong', "Choose a shorter passage. A saved idea can contain up to {0} characters.", ASSISTANT_SAVED_IDEA_TEXT_LIMIT)); return false;
				case 'full': setFeedback(localize('assistantView.ideaFull', "Saved ideas has reached its limit of {0}. Remove an idea to make room.", ASSISTANT_SAVED_IDEAS_LIMIT)); return false;
			}
		};
		const browseSavedIdeas = async (query = '', fromCommand = false): Promise<void> => {
			const owner = accountId;
			const thread = taskId;
			const composer = this.composer;
			const revision = draftRevision;
			const accountCheck = this.accountCheck;
			const activeTarget = this.agentChats.target;
			if (!owner || !composer || accountCheck !== verifiedAccountCheck) { return; }
			const original = composer.getInputValue();
			const session = savedIdeasPicker.value = new DisposableStore();
			const token = cancelOnDispose(session);
			const current = () => !token.isCancellationRequested && this.accountCheck === accountCheck && accountId === owner && taskId === thread
				&& this.agentChats.target === activeTarget && draftRevision === revision && this.composer?.getInputValue() === original;
			try {
				const ideas = store.listSavedIdeas(owner);
				if (!ideas.length) {
					setFeedback(localize('assistantView.noSavedIdeas', "No saved ideas yet. Use Save reply or enter /save followed by a passage."));
					return;
				}
				const idea = await pickSavedIdea(this.quickInput, ideas, query, token);
				if (!idea || !current()) { return; }
				const action = await this.quickInput.pick([
					{ label: localize('assistantView.attachSavedIdea', "Attach to draft"), id: 'attach' },
					{ label: localize('assistantView.continueSavedIdea', "Continue in task"), id: 'continue' },
					{ label: localize('assistantView.copySavedIdea', "Copy passage"), id: 'copy' },
					{ label: localize('assistantView.removeSavedIdea', "Remove saved idea"), id: 'remove' },
				], { title: localize('assistantView.useSavedIdea', "Use saved idea"), placeHolder: ideaExcerpt(idea.text) }, token);
				if (!action || !current()) { return; }
				if (!store.isCurrentSavedIdea(owner, idea)) {
					setFeedback(localize('assistantView.ideaGone', "This idea has already been removed."));
					return;
				}
				const label = localize('assistantView.ideaAttachment', "Saved idea · {0}", ideaExcerpt(idea.text).replace(/["<>]/g, ''));
				if (action.id === 'continue') {
					const target = captureAssistantTaskTarget(this.workspaceService, activeTarget);
					const destination = store.getOrCreateSavedIdeaTask(owner, idea.id, ideaExcerpt(idea.text), label, target);
					if (destination === 'tooLong') {
						setFeedback(localize('assistantView.ideaTaskTooLong', "This passage leaves no room for a message in a new task. Copy it or save a shorter passage."));
						return;
					}
					if (!destination) { return; }
					if (fromCommand) { setInput(''); this.saveDraft?.(); }
					if (destination.taskId !== taskId) { openThread(destination.taskId); } else { composer.focus(); }
					return;
				}
				if (action.id === 'copy') {
					await this.clipboard.writeText(idea.text);
					if (!current()) { return; }
					setFeedback(localize('assistantView.ideaCopied', "Passage copied."));
				} else if (action.id === 'remove') {
					store.removeSavedIdea(owner, idea.id);
					setFeedback(localize('assistantView.ideaRemoved', "Saved idea removed."));
				} else {
					const id = `assistant-saved:${idea.id}`;
					const draft = composer.getDraft();
					const alreadyAttached = draft.attachments.some(item => item.id === id);
					if (!alreadyAttached) {
						const attachedLength = draft.attachments.reduce((length, item) => typeof item.value === 'string' ? length + assistantTextAttachment(item.fullName ?? item.name, item.value).length : length, 0);
						if ((fromCommand ? 0 : original.length) + attachedLength + assistantTextAttachment(label, idea.text).length > REACH_ASSISTANT_PROMPT_LIMIT) {
							setFeedback(localize('assistantView.ideaAttachTooLong', "This idea and draft exceed the {0}-character message limit. Copy the passage or shorten the draft.", REACH_ASSISTANT_PROMPT_LIMIT));
							return;
						}
						if (fromCommand) { setInput(''); }
						composer.attachTextContext(label, idea.text, Codicon.bookmark, id);
						this.saveDraft?.();
						composer.focus();
						setFeedback(localize('assistantView.ideaAttached', "Saved idea attached to the draft."));
						return;
					}
					setFeedback(localize('assistantView.ideaAlreadyAttached', "This idea is already attached to the draft."));
				}
				if (fromCommand) { setInput(''); this.saveDraft?.(); }
				this.composer?.focus();
			} finally {
				if (savedIdeasPicker.value === session) { savedIdeasPicker.clear(); }
			}
		};
		const attachSkill = async (name?: string, request = ''): Promise<boolean> => {
			const owner = accountId;
			const thread = taskId;
			const revision = draftRevision;
			const original = this.composer?.getInputValue();
			const target = this.agentChats.target;
			const current = () => !this._store.isDisposed && accountId === owner && taskId === thread && draftRevision === revision
				&& this.agentChats.target === target && this.composer?.getInputValue() === original;
			try {
				const folders = await attachmentFolders();
				if (!folders || !current()) { return false; }
				const scope = await getAssistantSkillScope(folders, this.workspaceService, this.agentChats);
				const skills = this.configuration.getValue<boolean>(PromptsConfig.USE_AGENT_SKILLS)
					? await listAssistantSkills(this.promptsService, scope, 'user', CancellationToken.None) : [];
				if (!current()) { return false; }
				if (!skills.length) {
					setFeedback(localize('assistantView.noSkills', "No enabled skills are available for this Assistant chat."));
					return true;
				}
				const skill = name ? skills.find(skill => skill.name.toLowerCase() === name.toLowerCase()) : (await this.quickInput.pick(
					skills.map(skill => ({ label: skill.name, description: skill.pluginLabel ?? skill.extension?.displayName, detail: skill.description, skill })),
					{ title: localize('assistantView.skillsTitle', "Choose an Assistant skill"), matchOnDetail: true },
				))?.skill;
				if (!current()) { return false; }
				if (!skill) {
					if (name) { setFeedback(localize('assistantView.unknownSkill', "Skill {0} is unavailable. Use /skill to choose an enabled skill.", name)); }
					return true;
				}
				if (scope.roots && skill.storage === PromptsStorage.local && !await this.agentChats.isResourceTrusted(skill.uri)) {
					if (current()) { setFeedback(localize('assistantView.skillUntrusted', "This skill’s folder is no longer trusted.")); }
					return true;
				}
				const parsed = await this.promptsService.parseNew(skill.uri, CancellationToken.None);
				if (!current()) { return false; }
				const instructions = parsed.body?.getContent().trim();
				if (!instructions || !this.configuration.getValue<boolean>(PromptsConfig.USE_AGENT_SKILLS) || parsed.header?.userInvocable === false || this.promptsService.getDisabledPromptFiles(PromptsType.skill).has(skill.uri)) {
					setFeedback(localize('assistantView.skillUnavailable', "This skill has no available instructions."));
					return true;
				}
				const label = localize('assistantView.skillAttachment', "Skill · {0}", skill.name);
				if ((assistantTextAttachment(label, instructions) + request).length > REACH_ASSISTANT_PROMPT_LIMIT) {
					setFeedback(localize('assistantView.skillTooLong', "This skill and request exceed the {0}-character message limit.", REACH_ASSISTANT_PROMPT_LIMIT));
					return true;
				}
				setInput(request);
				this.composer?.attachTextContext(label, instructions, Codicon.book, `assistant-skill:${skill.uri}`);
				this.saveDraft?.();
				setFeedback(localize('assistantView.skillAttached', "Attached {0}. Review your request and send it when ready.", skill.name));
				this.composer?.focus();
			} catch {
				if (current()) { setFeedback(localize('assistantView.skillFailed', "Could not read the skill. Reopen /skill to check the available skills.")); }
			}
			return true;
		};
		const browsePrompts = async (onlyServer?: IMcpServer): Promise<void> => {
			const items: { label: string; description?: string; detail?: string; prompt: { arguments: readonly { name: string; description?: string; required?: boolean }[]; resolve(args: Record<string, string | undefined>): Promise<readonly { content: { type: string; text?: string } }[]> } }[] = [];
			for (const server of this.mcpService.servers.get()) {
				if (!isContributionEnabled(server.enablement.get()) || (onlyServer && server !== onlyServer)) { continue; }
				const serverLabel = server.serverMetadata.get()?.serverName ?? server.definition.label;
				for (const prompt of server.prompts.get()) {
					items.push({ label: prompt.title ?? prompt.name, description: serverLabel, detail: prompt.description, prompt });
				}
			}
			if (!items.length) {
				setFeedback(onlyServer
					? localize('assistantView.noPromptsServer', "\"{0}\" doesn't provide prompts.", onlyServer.serverMetadata.get()?.serverName ?? onlyServer.definition.label)
					: localize('assistantView.noPrompts', "No prompts available — connect an app that provides them."));
				setInput('');
				return;
			}
			const pick = await this.quickInput.pick(items, { title: localize('assistantView.promptsTitle', "App prompts"), placeHolder: localize('assistantView.promptsPlaceHolder', "Choose a prompt provided by a connected app"), matchOnDescription: true, matchOnDetail: true });
			if (!pick) { setInput(''); return; }
			const args: Record<string, string | undefined> = {};
			for (const arg of pick.prompt.arguments) {
				const value = await this.quickInput.input({ title: `${pick.label} — ${arg.name}${arg.required ? '*' : ''}`, placeHolder: arg.description ?? (arg.required ? localize('assistantView.promptArgRequired', "Required") : localize('assistantView.promptArgOptional', "Optional")) });
				if (value === undefined) { setInput(''); return; }
				args[arg.name] = value || undefined;
			}
			try {
				const messages = await pick.prompt.resolve(args);
				const text = messages.map(message => message.content.type === 'text' ? message.content.text ?? '' : '').filter(Boolean).join('\n\n').trim();
				setInput(text || pick.label);
			} catch {
				setFeedback(localize('assistantView.promptResolveFailed', "The app failed to resolve that prompt."));
				setInput('');
			}
			this.composer?.focus();
		};
		const browseResources = async (onlyServer?: IMcpServer): Promise<void> => {
			const items: { label: string; description?: string; detail?: string; name: string; uri: URI }[] = [];
			for (const server of this.mcpService.servers.get()) {
				if (!isContributionEnabled(server.enablement.get()) || (onlyServer && server !== onlyServer)) { continue; }
				const serverLabel = server.serverMetadata.get()?.serverName ?? server.definition.label;
				try {
					for await (const page of server.resources()) {
						for (const resource of page) {
							items.push({ label: resource.title ?? resource.name, description: `${serverLabel}${resource.mimeType ? ` · ${resource.mimeType}` : ''}`, detail: resource.description, name: resource.name, uri: resource.uri });
						}
					}
				} catch { /* A server that fails to list resources shouldn't block the others. */ }
			}
			if (!items.length) {
				setFeedback(onlyServer
					? localize('assistantView.noResourcesServer', "\"{0}\" doesn't publish resources.", onlyServer.serverMetadata.get()?.serverName ?? onlyServer.definition.label)
					: localize('assistantView.noResources', "No resources available — connected apps don't publish readable content."));
				setInput('');
				return;
			}
			const picks = await this.quickInput.pick(items, { canPickMany: true, title: localize('assistantView.resourcesTitle', "App resources"), placeHolder: localize('assistantView.resourcesPlaceHolder', "Choose resources to attach"), matchOnDescription: true, matchOnDetail: true });
			if (!picks?.length) { setInput(''); return; }
			const blocks: string[] = [];
			const names: string[] = [];
			for (const pick of picks) {
				try {
					const content = (await this.fileService.readFile(pick.uri, { limits: { size: 32 * 1024 } })).value.toString();
					const truncated = content.length > 12_000;
					blocks.push(`<resource name="${pick.name.replace(/"/g, '')}">\n${truncated ? `${content.slice(0, 12_000)}\n…(truncated)` : content}\n</resource>\n\n`);
					names.push(pick.label);
				} catch {
					setFeedback(localize('assistantView.attachFailed', "Could not read \"{0}\".", pick.label));
					setInput('');
					return;
				}
			}
			const combined = blocks.join('');
			if (combined.length > REACH_ASSISTANT_PROMPT_LIMIT) {
				setFeedback(localize('assistantView.attachTooLong', "\"{0}\" is too large to attach.", names.join(', ')));
				setInput('');
			} else {
				setInput(combined);
				this.saveDraft?.();
				setFeedback(names.length === 1
					? localize('assistantView.attached', "Attached \"{0}\".", names[0])
					: localize('assistantView.attachedMany', "Attached {0} files.", names.length));
			}
		};
		const routineScheduler = () => this.instantiation.invokeFunction(accessor => accessor.get(IAssistantBriefingScheduler));
		this.instantiation.invokeFunction(accessor => accessor.get(IAssistantMailWatcher)).start();
		const saveRoutine = async (owner: string, defaults?: Partial<IAssistantRoutineSettings>, routineId?: string) => {
			const current = () => !this._store.isDisposed && accountId === owner && !taskId;
			if (!this.directMail.connections.get().length && !this.mcpService.servers.get().some(server => isContributionEnabled(server.enablement.get()))) {
				setFeedback(localize('assistantView.routineNeedsConnection', "Connect an app before saving a routine."), {
					label: localize('assistantView.connectApp', "Connect app"), run: () => this.runs.openAppSetup(),
				});
				return false;
			}
			try {
				const template = assistantRoutineTemplates.find(candidate => candidate.id === (defaults as IAssistantRoutineTemplate | undefined)?.id);
				const serverIds = template && !defaults?.serverIds ? assistantRoutineTemplateServerIds(template, [
					...this.mcpService.servers.get().filter(server => isContributionEnabled(server.enablement.get()))
						.map(server => ({ id: assistantConnectionKey(server), label: server.definition.label })),
					...this.directMail.connections.get().map(mail => ({ id: mail.id, label: mail.label })),
				]) : undefined;
				const settings = await this.pickRoutineSettings(current, serverIds ? { ...defaults, serverIds } : defaults, !!routineId);
				if (!settings || !current()) { return false; }
				const saved = await routineScheduler().saveRoutine(owner, settings, routineId);
				if (current()) {
					setFeedback(saved ? localize('assistantView.routineSaved', "Routine saved. Scheduled runs happen while Agents is open.")
						: localize('assistantView.routineChanged', "This routine has changed. Open /routine again to refresh the list."));
				}
				return saved;
			} catch {
				if (current()) { setFeedback(localize('assistantView.routinesFailed', "Couldn’t access Assistant routines.")); }
				return false;
			}
		};
		const runRoutine = async (owner: string, id: string) => {
			if (this.runs.getPending(owner)) {
				setFeedback(localize('assistantView.waitForReply', "Wait for the current reply first."));
				return;
			}
			try {
				const started = await routineScheduler().runRoutineNow(owner, id);
				if (!started && accountId === owner && !taskId && !this.runs.getError(owner)) {
					setFeedback(localize('assistantView.routineUnavailable', "Couldn’t start this routine. It may have been removed or Assistant may be busy."));
				}
			} catch {
				if (accountId === owner && !taskId) { setFeedback(localize('assistantView.routinesFailed', "Couldn’t access Assistant routines.")); }
			}
		};
		const manageRoutines = async (owner: string, fromSlashCommand = false) => {
			const current = () => !this._store.isDisposed && accountId === owner && !taskId;
			const clearRoutineCommand = () => {
				if (fromSlashCommand && this.composer?.getInputValue().trim().toLowerCase() === '/routine') {
					setInput('');
					this.saveDraft?.();
				}
			};
			try {
				const scheduler = routineScheduler();
				const routines = await scheduler.getRoutines(owner);
				if (!current()) { return false; }
				const item = await this.quickInput.pick([
					{ label: localize('assistantView.newRoutine', "New routine"), routine: undefined },
					...routines.map(routine => ({
						routine,
						label: escapeIcons(routine.name),
						description: routine.enabled
							? localize('assistantView.routineNext', "Next run {0}", new Date(routine.nextAt).toLocaleString(undefined, { timeZone: routine.timeZone }))
							: localize('assistantView.routinePaused', "Paused"),
						detail: localize('assistantView.routineDetails', "{0} at {1} ({2}) · {3} connection(s) · {4}", routine.cadence === 'daily'
							? localize('assistantView.routineDaily', "Daily") : localize('assistantView.routineWeekdays', "Weekdays"),
							formatScheduleTime(routine), routine.timeZone, routine.serverIds.length, describeLastRoutineRun(routine)),
					})),
				], { title: localize('assistantView.routines', "Assistant routines") });
				if (!item || !current()) { return false; }
				const { routine } = item;
				const reportId = routine?.lastRun?.id;
				const hasReport = reportId && store.read(owner).turns.some(turn => turn.routineRunId === reportId);
				const action = routine ? await this.quickInput.pick([
					{ label: localize('assistantView.runRoutineNow', "Run now"), id: 'run' },
					...(hasReport ? [{ label: localize('assistantView.openRoutineReport', "Open Last Report"), id: 'report' }] : []),
					{ label: localize('assistantView.editRoutine', "Edit routine"), id: 'edit' },
					{ label: routine.enabled ? localize('assistantView.pauseRoutine', "Pause routine") : localize('assistantView.resumeRoutine', "Resume routine"), id: 'toggle' },
					{ label: localize('assistantView.deleteRoutine', "Delete routine"), id: 'delete' },
				], { title: routine.name }) : { id: 'edit' };
				if (!action || !current()) { return false; }
				let message: string;
				let updated: boolean;
				switch (action.id) {
					case 'run':
						if (!routine) { return false; }
						clearRoutineCommand();
						void runRoutine(owner, routine.id);
						return true;
					case 'report': {
						if (!reportId) { return false; }
						if (searchQuery) {
							searchQuery = '';
							turnCount = -1;
						}
						renderTurns();
						const report = Array.from(transcript.children).find((element): element is HTMLElement =>
							isHTMLElement(element) && element.dataset.routineRunId === reportId);
						if (!report) {
							setFeedback(localize('assistantView.routineReportGone', "This report is no longer in the conversation."));
							return false;
						}
						clearRoutineCommand();
						setFeedback('');
						report.tabIndex = -1;
						report.focus({ preventScroll: true });
						report.scrollIntoView({ block: 'center' });
						return true;
					}
					case 'toggle':
						if (!routine) { return false; }
						updated = await scheduler.setRoutineEnabled(owner, routine.id, !routine.enabled);
						message = routine.enabled ? localize('assistantView.routinePausedFeedback', "Routine paused.") : localize('assistantView.routineResumed', "Routine resumed.");
						break;
					case 'delete': {
						if (!routine) { return false; }
						const answer = await this.dialogs.confirm({
							message: localize('assistantView.deleteRoutineConfirm', "Delete the routine “{0}”?", routine.name),
							primaryButton: localize('assistantView.deleteRoutineButton', "Delete"),
						});
						if (!answer.confirmed || !current()) { return false; }
						updated = await scheduler.deleteRoutine(owner, routine.id);
						message = localize('assistantView.routineDeleted', "Routine deleted.");
						break;
					}
					default: {
						if (routine) {
							updated = await saveRoutine(owner, routine, routine.id);
						} else {
							const blank = { label: localize('assistantView.blankRoutine', "Start from scratch"), template: undefined };
							const choice = await this.quickInput.pick([
								blank,
								...assistantRoutineTemplates.map(template => ({
									label: template.name,
									description: template.description,
									detail: localize('assistantView.routineTemplateDetail', "Suggested: {0} at {1} · Uses matching {2} connections when available", template.cadence === 'daily' ? localize('assistantView.routineDaily', "Daily") : localize('assistantView.routineWeekdays', "Weekdays"), formatScheduleTime(template), template.sourceNames.join(', ')),
									template,
								})),
							], { title: localize('assistantView.routineTemplates', "Choose a routine") });
							if (!choice || !current()) { return false; }
							updated = await saveRoutine(owner, choice.template, undefined);
						}
						if (updated && current()) { clearRoutineCommand(); }
						return updated;
					}
				}
				if (current()) {
					setFeedback(updated ? message : localize('assistantView.routineChanged', "This routine has changed. Open /routine again to refresh the list."));
					if (updated) { clearRoutineCommand(); }
				}
				return updated;
			} catch {
				if (current()) { setFeedback(localize('assistantView.routinesFailed', "Couldn’t access Assistant routines.")); }
				return false;
			}
		};
		let submittingEmail = false;
		const submit = async (): Promise<boolean> => {
			let input = this.composer?.getInputValue().trim();
			if (!accountId || !input) { return false; }
			const initialInput = input;
			if (/^\/save(?:\s|$)/i.test(input)) {
				const passage = this.composer!.getInputValue().trimStart().replace(/^\/save(?:\r\n|\s)?/i, '');
				if (!saveIdea(accountId, taskId, passage)) { return false; }
				setInput('');
				this.saveDraft?.();
				return true;
			}
			if (/^\/saved(?:\s|$)/i.test(input)) {
				await browseSavedIdeas(input.slice('/saved'.length).trim(), true);
				return true;
			}
			if (/^\/weather(?:\s|$)/i.test(input)) {
				const owner = accountId;
				const thread = taskId;
				const current = () => accountId === owner && taskId === thread && !this._store.isDisposed;
				const title = localize('assistantView.weatherNetworkTitle', "The Weather Network");
				setInput('');
				if (input.toLowerCase() !== '/weather') {
					setFeedback(localize('assistantView.weatherNetworkUseSetup', "Run /weather to enter your key in the secure setup field."));
					return true;
				}
				try {
					const configured = await this.reach.hasWeatherNetworkKey();
					if (!current()) { return false; }
					if (configured) {
						const action = await this.quickInput.pick([
							{ id: 'change', label: localize('assistantView.weatherNetworkChange', "Replace API key") },
							{ id: 'remove', label: localize('assistantView.weatherNetworkRemove', "Remove API key") },
						], { title });
						if (!action || !current()) { return false; }
						if (action.id === 'remove') {
							await this.reach.removeWeatherNetworkKey();
							if (current()) { setFeedback(localize('assistantView.weatherNetworkRemoved', "The Weather Network key was removed.")); }
							return true;
						}
					}
					const key = await this.quickInput.input({
						title, password: true,
						prompt: localize('assistantView.weatherNetworkPrompt', "Enter a Pelmorex API key with access to Long Term Forecasts."),
					});
					if (!key || !current()) { return false; }
					await this.reach.saveWeatherNetworkKey(key);
					if (current()) { setFeedback(localize('assistantView.weatherNetworkConnected', "The Weather Network is ready. Ask for a forecast by city.")); }
				} catch (error) {
					if (current()) { setFeedback(error instanceof Error ? error.message : localize('assistantView.weatherNetworkFailed', "Could not set up The Weather Network.")); }
				}
				return true;
			}
			if (input.toLowerCase() === '/routine') {
				if (taskId) {
					setFeedback(localize('assistantView.routinesFromMain', "Return to Assistant to manage routines."));
					return false;
				}
				return manageRoutines(accountId, true);
			}
			if (input.toLowerCase() === '/briefing') {
				const owner = accountId;
				const thread = taskId;
				const periods: readonly AssistantBriefingPeriod[] = ['morning', 'evening'];
				const scheduler = this.instantiation.invokeFunction(accessor => accessor.get(IAssistantBriefingScheduler));
				const schedules = await scheduler.getSchedules(owner).catch(() => []);
				if (this._store.isDisposed || accountId !== owner || taskId !== thread) { return false; }
				const pick = await this.quickInput.pick([
					...periods.map(period => ({ label: ASSISTANT_BRIEFINGS[period].name, description: localize('assistantView.briefingRunNow', "Run now"), detail: ASSISTANT_BRIEFINGS[period].description, period, action: 'run' as const })),
					...periods.map(period => {
						const schedule = schedules.find(item => item.period === period);
						return {
							label: localize('assistantView.briefingSchedule', "Schedule {0}", ASSISTANT_BRIEFINGS[period].name),
							description: schedule ? localize('assistantView.briefingScheduledTime', "Daily at {0} ({1})", formatScheduleTime(schedule), schedule.timeZone) : localize('assistantView.briefingDaily', "Daily while Agents is open"),
							detail: schedule?.lastRun ? localize('assistantView.briefingLastRun', "Last run: {0} · {1}",
								schedule.lastRun.retryAt === undefined ? describeAssistantRunStatus(schedule.lastRun.status) : localize('assistantView.briefingRetryScheduled', "Failed; retry scheduled"),
								new Date(schedule.lastRun.startedAt).toLocaleString(undefined, { timeZone: schedule.timeZone })) : undefined,
							period, action: 'schedule' as const,
						};
					}),
					...schedules.map(schedule => ({ label: localize('assistantView.briefingDisable', "Turn off {0}", ASSISTANT_BRIEFINGS[schedule.period].name), period: schedule.period, action: 'disable' as const })),
				], { title: localize('assistantView.briefingTitle', "Assistant briefings") });
				if (!pick || this._store.isDisposed || accountId !== owner || taskId !== thread) { return false; }
				if (pick.action === 'run') {
					setInput(`/briefing ${pick.period}`);
					return submit();
				}
				const { name, schedule: preset } = ASSISTANT_BRIEFINGS[pick.period];
				let time: IAssistantBriefingTime | undefined;
				if (pick.action === 'schedule') {
					const previous = schedules.find(item => item.period === pick.period);
					const timeZone = Intl.DateTimeFormat().resolvedOptions().timeZone;
					const value = await this.quickInput.input({
						title: pick.label,
						value: formatScheduleTime(previous ?? { hour: preset.scheduleHour, minute: preset.scheduleMinute }),
						prompt: localize('assistantView.briefingTimePrompt', "Enter a daily time in HH:mm ({0}). Assistant uses read-only tools while Agents is open.", timeZone),
						validateInput: async value => parseAssistantBriefingTime(value) ? undefined : localize('assistantView.briefingInvalidTime', "Enter a 24-hour time, for example 08:00."),
					});
					time = value === undefined ? undefined : parseAssistantBriefingTime(value);
					if (!time || this._store.isDisposed || accountId !== owner || taskId !== thread) { return false; }
				}
				try {
					await scheduler.setSchedule(owner, pick.period, time);
				} catch {
					setFeedback(localize('assistantView.briefingScheduleFailed', "Couldn’t save the briefing schedule."));
					return false;
				}
				if (this._store.isDisposed || accountId !== owner || taskId !== thread) { return true; }
				setInput('');
				this.saveDraft?.();
				setFeedback(time ? localize('assistantView.briefingScheduled', "{0} scheduled daily at {1} while Agents is open.", name, formatScheduleTime(time))
					: localize('assistantView.briefingDisabled', "{0} schedule turned off.", name));
				return true;
			}
			if (input === '/task' || (taskId && input.startsWith('/task '))) {
				setFeedback(taskId ? localize('assistantView.taskFromMain', "Return to Assistant to start a linked task.") : localize('assistantView.taskNeedsDescription', "Add a description after /task to start a linked task."));
				return false;
			}
			const owner = accountId;
			const threadId = taskId;
			const commandRevision = draftRevision;
			const commandIsCurrent = () => !this._store.isDisposed && accountId === owner && taskId === threadId
				&& draftRevision === commandRevision && this.composer?.getInputValue().trim() === initialInput;
			if (input === '/documents') {
				const documents = store.listDocuments(owner);
				if (!documents.length) {
					setFeedback(localize('assistantView.noDocuments', "No Assistant documents yet. Ask for a report, proposal, plan, or draft and Assistant can save one."));
				} else {
					const pick = await this.quickInput.pick(documents.map(document => {
						const firstLine = document.content.trim().split('\n').map(line => line.replace(/^#+\s*/, '').trim()).find(line => !!line) ?? '';
						return {
							document,
							label: document.title,
							description: localize('assistantView.documentUpdated', "Updated {0}", new Date(document.updatedAt).toLocaleString()),
							detail: firstLine.length > 160 ? `${firstLine.slice(0, 160)}…` : firstLine,
						};
					}), { title: localize('assistantView.documentsTitle', "Assistant documents"), matchOnDescription: true, matchOnDetail: true });
					if (!commandIsCurrent()) { return false; }
					if (pick) {
						const action = await this.quickInput.pick([
							{ label: localize('assistantView.openDocumentAction', "Open"), id: 'open' },
							{ label: localize('assistantView.copyDocumentAction', "Copy"), id: 'copy' },
							{ label: localize('assistantView.exportDocumentAction', "Export as Markdown…"), id: 'export' },
							{ label: localize('assistantView.deleteDocumentAction', "Delete"), id: 'delete' },
						], { title: pick.document.title });
						if (!commandIsCurrent()) { return false; }
						if (action?.id === 'open') {
							const latest = store.readDocument(owner, pick.document.id);
							if (latest) { await openReplyDocument(latest.content, { accountId: owner, documentId: pick.document.id }); }
							else { setFeedback(localize('assistantView.documentMissing', "This document is no longer available.")); }
						} else if (action?.id === 'copy') {
							const latest = store.readDocument(owner, pick.document.id);
							if (latest) {
								await this.clipboard.writeText(latest.content);
								setFeedback(localize('assistantView.documentCopied', "Document copied."));
							} else { setFeedback(localize('assistantView.documentMissing', "This document is no longer available.")); }
						} else if (action?.id === 'export') {
							const latest = store.readDocument(owner, pick.document.id);
							if (!latest) { setFeedback(localize('assistantView.documentMissing', "This document is no longer available.")); }
							else {
								const baseName = (latest.title.replace(/[\/:*?"<>|]/g, '-').replace(/\s+/g, ' ').trim() || 'assistant-document').slice(0, 80);
								const target = await this.fileDialogs.showSaveDialog({
									title: localize('assistantView.exportDocumentTitle', "Export Assistant document"),
									defaultUri: joinPath(await this.fileDialogs.defaultFilePath(), `${baseName}.md`),
								});
								if (!commandIsCurrent()) { return false; }
								if (target) {
									try {
										await this.fileService.writeFile(target, VSBuffer.fromString(latest.content));
										setFeedback(localize('assistantView.documentExported', "Document exported."));
									} catch {
										setFeedback(localize('assistantView.documentExportFailed', "The document could not be exported."));
									}
								}
							}
						} else if (action?.id === 'delete') {
							const answer = await this.dialogs.confirm({
								message: localize('assistantView.deleteDocumentConfirm', "Delete the document “{0}”?", pick.document.title),
								primaryButton: localize('assistantView.deleteDocumentButton', "Delete"),
							});
							if (!commandIsCurrent()) { return false; }
							if (answer.confirmed) {
								for (const [thread, active] of activeArtifacts) {
									if (active.documentId === pick.document.id) { activeArtifacts.delete(thread); }
								}
								setFeedback(store.removeDocument(owner, pick.document.id)
									? localize('assistantView.documentDeleted', "Document deleted.")
									: localize('assistantView.documentMissing', "This document is no longer available."));
							}
						}
					}
				}
				setInput('');
				return true;
			}
			if (input === '/memory') {
				const bank = store.readMemoryBank(accountId);
				const userTag = localize('assistantView.memoryUserTag', "user profile");
				const notesTag = localize('assistantView.memoryNotesTag', "agent notes");
				const addLabel = localize('assistantView.memoryAdd', "Add a memory…");
				const pick = await this.quickInput.pick([
					{ label: addLabel, target: undefined },
					...memoryEntries(bank.user).map(entry => ({ label: entry, description: userTag, target: 'user' as const })),
					...memoryEntries(bank.memory).map(entry => ({ label: entry, description: notesTag, target: 'memory' as const })),
				], {
					title: localize('assistantView.memoryTitle', "Assistant memory ({0})", memoryUsage(bank)),
					placeHolder: localize('assistantView.memoryPlaceHolder', "Select a memory to edit or remove, or add a new one"),
				});
				if (!commandIsCurrent()) { return false; }
				if (pick) {
					if (pick.target === undefined) {
						const target = await this.quickInput.pick([
							{ label: userTag, description: localize('assistantView.memoryUserHint', "Who the user is: name, role, preferences"), target: 'user' as const },
							{ label: notesTag, description: localize('assistantView.memoryNotesHint', "Your notes: environment, conventions, lessons"), target: 'memory' as const },
						], { title: localize('assistantView.memoryTarget', "Save to which memory?") });
						if (!commandIsCurrent()) { return false; }
						if (target) {
							const content = await this.quickInput.input({
								title: localize('assistantView.memoryNew', "New memory"),
								prompt: localize('assistantView.memoryNewPrompt', "One durable fact, written as a statement."),
							});
							if (!commandIsCurrent()) { return false; }
							if (content?.trim()) {
								const result = store.applyMemory(accountId, target.target, [{ action: 'add', content: content.trim() }]);
								setFeedback(result.ok ? localize('assistantView.memorySaved', "Memory saved. {0}", result.usage) : result.error);
							}
						}
					} else {
						const target = pick.target;
						const action = await this.quickInput.pick([
							{ label: localize('assistantView.memoryEdit', "Edit…"), edit: true },
							{ label: localize('assistantView.memoryRemove', "Remove"), edit: false },
						], { title: pick.label });
						if (!commandIsCurrent()) { return false; }
						if (action?.edit) {
							const content = await this.quickInput.input({
								title: localize('assistantView.memoryEditTitle', "Edit memory"),
								value: pick.label,
								prompt: localize('assistantView.memoryNewPrompt', "One durable fact, written as a statement."),
							});
							if (!commandIsCurrent()) { return false; }
							if (content !== undefined) {
								const result = store.applyMemory(accountId, target, [{ action: 'patch', old_text: pick.label, content, exact: true }]);
								setFeedback(result.ok ? localize('assistantView.memorySaved', "Memory saved. {0}", result.usage) : result.error);
							}
						} else if (action) {
							const result = store.applyMemory(accountId, target, [{ action: 'remove', old_text: pick.label, exact: true }]);
							setFeedback(result.ok ? localize('assistantView.memoryRemoved', "Memory removed. {0}", result.usage) : result.error);
						}
					}
				}
				setInput('');
				return true;
			}
			if (input === '/forget') {
				const bank = store.readMemoryBank(accountId);
				const userEntries = memoryEntries(bank.user);
				const noteEntries = memoryEntries(bank.memory);
				const total = userEntries.length + noteEntries.length;
				if (!total) {
					setFeedback(localize('assistantView.noMemory', "No memory saved yet. Ask the assistant to remember something."));
				} else {
					const userTag = localize('assistantView.memoryUserTag', "user profile");
					const notesTag = localize('assistantView.memoryNotesTag', "agent notes");
					const picks = await this.quickInput.pick([
						...userEntries.map(entry => ({ label: entry, description: userTag, picked: true })),
						...noteEntries.map(entry => ({ label: entry, description: notesTag, picked: true })),
					], {
						canPickMany: true,
						title: localize('assistantView.forgetTitle', "Assistant memory"),
						placeHolder: localize('assistantView.forgetPlaceHolder', "Uncheck memories to forget, then confirm"),
					});
					if (!commandIsCurrent()) { return false; }
					if (picks) {
						const keptUser = new Set(picks.filter(pick => pick.description === userTag).map(pick => pick.label));
						const keptNotes = new Set(picks.filter(pick => pick.description === notesTag).map(pick => pick.label));
						const forgot = store.forgetMemoryEntries(owner, {
							user: userEntries.filter(entry => !keptUser.has(entry)),
							memory: noteEntries.filter(entry => !keptNotes.has(entry)),
						});
						setFeedback(forgot === 0
							? localize('assistantView.memoryKept', "Memory unchanged.")
							: localize('assistantView.memoryForgot', "Forgot {0} memor{1}.", forgot, forgot === 1 ? 'y' : 'ies'));
					}
				}
				setInput('');
				return true;
			}
			if (input === '/watch' || input.startsWith('/watch ')) {
				const watcher = this.instantiation.invokeFunction(accessor => accessor.get(IAssistantMailWatcher));
				const arg = input.slice('/watch'.length).trim();
				if (arg === 'on' || arg === 'off') {
					await watcher.setEnabled(arg === 'on');
					setFeedback(arg === 'on'
						? localize('assistantView.watchOn', "Important-mail watch enabled. Connected mailboxes are checked quietly while Agents is open.")
						: localize('assistantView.watchOff', "Important-mail watch disabled."));
				} else if (arg === 'check') {
					const found = await watcher.checkNow(accountId);
					setFeedback(found
						? localize('assistantView.watchFound', "Reported {0} new important email(s).", found)
						: localize('assistantView.watchNone', "No new important emails in the connected mailboxes."));
				} else {
					const status = watcher.status(accountId);
					watcher.markSeen(accountId);
					const connections = status.connections.map(connection => `${connection.label}: ${connection.connected
						? localize('assistantView.watchConnected', "connected")
						: localize('assistantView.watchSkipped', "not connected — skipped")}`).join('; ');
					setFeedback(localize('assistantView.watchStatus', "Important-mail watch is {0}. {1}.{2}",
						status.enabled ? localize('assistantView.watchOnWord', "on") : localize('assistantView.watchOffWord', "off"),
						connections || localize('assistantView.watchNoConnections', "no mail connections"),
						status.alerts.length ? ` ${localize('assistantView.watchAlerts', "{0} recent alert(s) — ask me about them.", status.alerts.length)}` : ''));
				}
				setInput('');
				return true;
			}
			if (input === '/tools') {
				const catalog = collectThreadTools(accountId, taskId);
				void this.quickInput.pick([...catalog.entries.values()].map(entry => ({
					label: entry.name,
					description: entry.label,
					detail: entry.safe
						? localize('assistantView.toolSafe', "Read-only — runs automatically in Auto mode")
						: localize('assistantView.toolSideEffects', "Can have side effects — may ask before running"),
				})), { title: localize('assistantView.toolsTitle', "Tools available in this conversation"), matchOnDescription: true });
				setInput('');
				return true;
			}
			if (input === '/clear') {
				setInput('');
				this.saveDraft?.();
				await clearChat();
				return true;
			}
			if (input === '/stop') {
				stopRequest();
				setInput('');
				return true;
			}
			if (input === '/tasks') {
				const tasks = [...(store.read(accountId).tasks ?? [])].sort((a, b) => Number(!!b.pinned) - Number(!!a.pinned));
				if (!tasks.length) {
					setFeedback(localize('assistantView.noTasks', "No linked tasks yet — /task <description> starts one."));
				} else {
					const pick = await this.quickInput.pick(tasks.map(task => ({
						label: task.title || localize('assistantView.newTask', "New task"),
						description: `${task.status}${task.update ? ` — ${task.update}` : ''}`,
						task,
					})), { title: localize('assistantView.tasksTitle', "Linked tasks"), matchOnDescription: true });
					if (pick && accountId) { openThread(pick.task.id); }
				}
				setInput('');
				return true;
			}
			if (input === '/agents') {
				setInput('');
				await showAgentChats();
				return true;
			}
			if (input === '/worktree') {
				setInput('');
				await pickWorktree();
				return true;
			}
			if (input === '/model') {
				modelButton.click();
				setInput('');
				return true;
			}
			if (input === '/connect' || input.startsWith('/connect ')) {
				const name = input.slice('/connect'.length).trim();
				const app = name && findAssistantAppConnection(name);
				if (name && !app) {
					setFeedback(localize('assistantView.unknownConnection', "No app named {0}. Use /connect to browse apps.", name));
					return true;
				}
				setInput('');
				this.runs.openAppSetup(app ? app.id : undefined);
				return true;
			}
			if (/^\/send-email(?:\s|$)/i.test(input)) {
				if (submittingEmail) { return true; }
				submittingEmail = true;
				const owner = accountId;
				const thread = taskId;
				const current = () => accountId === owner && taskId === thread && !this._store.isDisposed
					&& this.composer?.getInputValue().trim() === input;
				try {
					const availableMailboxes = () => this.directMail.connections.get().filter(mail => mail.canSend && !mail.needsSignIn);
					let mailboxes = availableMailboxes();
					if (!mailboxes.length) {
						const choice = await this.quickInput.pick([
							{ label: signalCDirectMailConnections.gmail.name, provider: 'gmail' as const },
							{ label: signalCDirectMailConnections.outlook.name, provider: 'outlook' as const },
						], { title: localize('assistantView.sendEmailConnect', "Choose a mailbox and sign in to read and send") });
						if (!choice || !current()) { return true; }
						await this.directMail.manage(choice.provider);
						if (!current()) { return true; }
						mailboxes = availableMailboxes();
						if (!mailboxes.length) {
							setFeedback(localize('assistantView.sendEmailAfterConnect', "Choose Sign In to Read and Send and grant mailing permission in the browser."));
							return true;
						}
					}
					const [header, ...lines] = input.slice('/send-email'.length).trim().split(/\r?\n/);
					const separator = header.indexOf('|');
					const to = separator < 0 ? '' : header.slice(0, separator).trim();
					const subject = separator < 0 ? '' : header.slice(separator + 1).trim();
					const body = lines.join('\n').trim();
					if (!isSignalCMailAddress(to) || !subject || !body) {
						setFeedback(localize('assistantView.sendEmailFormat', "Enter /send-email address | subject, then press Shift+Enter and write the message."));
						return true;
					}
					const mailbox = mailboxes.length === 1 ? mailboxes[0] : (await this.quickInput.pick(
						mailboxes.map(mail => ({ label: mail.label, mail })), { title: localize('assistantView.sendEmailSender', "Choose the sending mailbox") }))?.mail;
					if (!mailbox || !current()) { return true; }
					const confirmation = await this.dialogs.confirm({
						type: 'info',
						message: localize('assistantView.sendEmailConfirm', "Send this email to {0}?", to),
						detail: localize('assistantView.sendEmailDetails', "From {0}\nSubject: {1}\n\n{2}", mailbox.label, subject, body),
						primaryButton: localize('assistantView.sendEmailButton', "Send email"),
					});
					if (!confirmation.confirmed || !current()) { return true; }
					setFeedback(localize('assistantView.sendingEmail', "Sending email…"));
					const result = await this.directMail.send(mailbox, { to, subject, body }, CancellationToken.None);
					if (current()) {
						setInput('');
						setFeedback(result);
					}
				} catch (error) {
					if (current()) { setFeedback(error instanceof Error ? error.message : localize('assistantView.sendEmailFailed', "Could not confirm email submission. Check Sent before retrying.")); }
				} finally {
					submittingEmail = false;
				}
				return true;
			}
			const skillCommand = /^\/skill(?:\s+(\S+))?(?:\s+([\s\S]+))?$/i.exec(input);
			if (skillCommand) { return attachSkill(skillCommand[1], skillCommand[2]); }
			if (input === '/prompts') {
				await browsePrompts();
				return true;
			}
			if (input === '/help') {
				const commands = assistantSlashCommands.filter(command => command.label !== '/help');
				const pick = await this.quickInput.pick(commands, { title: localize('assistantView.helpTitle', "Assistant commands"), matchOnDescription: true });
				setInput(pick ? `${pick.label}${pick.acceptsArguments ? ' ' : ''}` : '');
				if (pick) { this.composer?.focus(); }
				return true;
			}
			if (input === '/search' || input.startsWith('/search ')) {
				searchQuery = input === '/search' ? '' : input.slice('/search'.length).trim().toLowerCase();
				turnCount = -1;
				renderTurns();
				const matches = searchQuery
					? store.read(accountId, taskId).turns.filter(turn => turn.prompt.toLowerCase().includes(searchQuery) || turn.reply.toLowerCase().includes(searchQuery)).length
					: 0;
				setFeedback(searchQuery
					? localize('assistantView.searchResults', "{0} turn(s) match \"{1}\" — /search clears the filter.", matches, searchQuery)
					: '');
				setInput('');
				return true;
			}
			if (input === '/copy') {
				const last = [...store.read(accountId, taskId).turns].pop();
				if (last) {
					await this.clipboard.writeText(last.reply);
					setFeedback(localize('assistantView.copied', "Reply copied."));
				} else {
					setFeedback(localize('assistantView.nothingToCopy', "Nothing to copy yet."));
				}
				setInput('');
				return true;
			}
			if (input === '/context') {
				const stats = this.runs.describeContext(accountId, taskId);
				setFeedback(localize('assistantView.searchableContextInfo', "The next reply includes {0} of {1} turns, about {2} of {3} characters. Ask about a date or phrase to search earlier turns.", stats.kept, stats.total, stats.chars.toLocaleString(), stats.limit.toLocaleString()));
				setInput('');
				return true;
			}
			if (input === '/undo') {
				const last = store.popTurn(accountId, taskId);
				setFeedback(last?.routineId ? localize('assistantView.routineUndone', "Removed the routine’s last report.")
					: last ? localize('assistantView.undone', "Removed the last exchange — its prompt is back in the composer.") : localize('assistantView.nothingToUndo', "Nothing to undo yet."));
				setInput(last?.routineId ? '' : last?.prompt ?? '');
				if (last) { this.composer?.focus(); }
				return true;
			}
			if (input === '/retry') {
				const failed = this.runs.getFailedPrompt(accountId, taskId);
				const routineId = !taskId ? this.runs.getFailedRoutine(accountId) : undefined;
				const retryQueued = !failed && !routineId && !!store.read(accountId, taskId).startedQueuedPrompt && !this.runs.getPending(accountId, taskId);
				setFeedback(failed || routineId || retryQueued ? '' : localize('assistantView.nothingToRetry', "Nothing to retry."));
				setInput('');
				if (routineId) { void runRoutine(accountId, routineId); }
				if (failed) { void this.runs.send(accountId, failed, taskId); }
				if (retryQueued) { this.runs.sendQueuedPrompt(accountId, taskId); }
				return true;
			}
			if (input === '/regenerate') {
				const routineId = store.read(accountId, taskId).turns.at(-1)?.routineId;
				if (routineId) {
					setInput('');
					void runRoutine(accountId, routineId);
					return true;
				}
				const turns = store.read(accountId, taskId).turns;
				const last = turns.at(-1);
				if (this.runs.getPending(accountId, taskId)) {
					setFeedback(localize('assistantView.waitForReply', "Wait for the current reply first."));
					return true;
				}
				setFeedback(last ? '' : localize('assistantView.nothingToRegenerate', "Nothing to regenerate yet."));
				setInput('');
				if (last) { void this.runs.send(accountId, last.prompt, taskId, { requestStart: last.requestStart, replaceFrom: turns.length - 1, preserveDraft: true }); }
				return true;
			}
			if (input === '/compact') {
				const owner = accountId;
				const thread = taskId;
				const { turns } = store.read(owner, thread);
				if (turns.length < 2 || turns.at(-1)?.compacted) {
					setFeedback(localize('assistantView.nothingToCompact', "Nothing to compact yet."));
					setInput('');
					return true;
				}
				if (this.runs.getPending(owner, thread)) {
					setFeedback(localize('assistantView.compactBusy', "Wait for the current reply first."));
					return false;
				}
				const label = localize('assistantView.compactLabel', "Summarize the conversation");
				const stats = this.runs.describeContext(owner, thread, label.length);
				setInput('');
				this.saveDraft?.();
				const done = await this.runs.send(owner, label, thread, { compact: true, preserveDraft: true });
				if (!this._store.isDisposed && accountId === owner && taskId === thread) {
					setFeedback(done ? localize('assistantView.compactRetained', "Saved a summary using {0} of {1} turns as initial context. All exchanges are available in search and export.", stats.kept, stats.total)
						: localize('assistantView.compactFailed', "The summary could not be saved. Your exchanges are still available."));
				}
				return true;
			}
			if (input === '/export') {
				const turns = store.read(accountId, taskId).turns;
				if (!turns.length) {
					setFeedback(localize('assistantView.nothingToExport', "Nothing to export yet."));
				} else {
					await this.clipboard.writeText(turns.map(turn => {
						const tools = turn.activity?.length
							? `\n\n_Tools: ${turn.activity.map(a => `${a.label} (${a.status})`).join(', ')}_`
							: '';
						const stamp = turn.at ? ` · ${new Date(turn.at).toLocaleString()}` : '';
						return `**You**${stamp} — ${turn.prompt}\n\n${turn.reply}${tools}`;
					}).join('\n\n---\n\n'));
					setFeedback(localize('assistantView.exported', "Copied {0} turn(s) to the clipboard.", turns.length));
				}
				setInput('');
				return true;
			}
			if (input === '/attach' || input.startsWith('/attach ')) {
				const revision = draftRevision;
				const folders = await attachmentFolders();
				if (!folders || revision !== draftRevision) { return false; }
				let rels = input.slice('/attach'.length).match(/"[^"]+"|\S+/g)?.map(path => path.replace(/^"|"$/g, '')) ?? [];
				if (!rels.length) {
					rels = await pickWorkspaceFiles(folders);
					if (this._store.isDisposed || revision !== draftRevision) { return false; }
					if (!rels.length) {
						setInput('');
						return true;
					}
				}
				const blocks: string[] = [];
				const names: string[] = [];
				for (const rel of rels) {
					const uri = resolveAssistantWorkspacePath(folders, rel);
					if (typeof uri === 'string') {
						setFeedback(uri);
						setInput('');
						return true;
					}
					try {
						const content = (await this.fileService.readFile(uri, { limits: { size: 32 * 1024 } })).value.toString();
						if (this._store.isDisposed || revision !== draftRevision) { return false; }
						const truncated = content.length > 12_000;
						const label = formatAssistantWorkspacePath(folders, uri) ?? rel;
						blocks.push(`<file path="${label.replace(/"/g, '')}">\n${truncated ? `${content.slice(0, 12_000)}\n…(truncated)` : content}\n</file>\n\n`);
						names.push(label);
					} catch {
						if (this._store.isDisposed || revision !== draftRevision) { return false; }
						setFeedback(localize('assistantView.attachFailed', "Could not read \"{0}\".", rel));
						setInput('');
						return true;
					}
				}
				const combined = blocks.join('');
				if (combined.length > REACH_ASSISTANT_PROMPT_LIMIT) {
					setFeedback(localize('assistantView.attachTooLong', "\"{0}\" is too large to attach.", names.join(', ')));
					setInput('');
				} else {
					setInput(combined);
					this.saveDraft?.();
					setFeedback(names.length === 1
						? localize('assistantView.attached', "Attached \"{0}\".", names[0])
						: localize('assistantView.attachedMany', "Attached {0} files.", names.length));
				}
				return true;
			}
			if (input === '/resources') {
				await browseResources();
				return true;
			}
			const submittedDraft = this.composer?.getDraft();
			if (!submittedDraft) { return false; }
			const taskDraft = !taskId ? getAssistantLinkedTaskDraft(submittedDraft) : undefined;
			let typedRequest = taskDraft?.inputText ?? input;
			input = typedRequest;
			let routineServerIds: readonly string[] | undefined;
			if (/(^|[\s([{])@/.test(typedRequest)) {
				const scheduler = this.instantiation.invokeFunction(accessor => accessor.get(IAssistantBriefingScheduler));
				const routines = await scheduler.getRoutines(accountId).catch(() => []);
				if (!commandIsCurrent()) { return false; }
				const mention = assistantRoutineMention(typedRequest, routines);
				if (mention) {
					const before = typedRequest.slice(0, mention.start).trimEnd();
					const after = typedRequest.slice(mention.end).trimStart();
					typedRequest = [before, after].filter(Boolean).join(' ') || localize('assistantView.routineMentionDefault', "Run this routine now.");
					input = assistantTextAttachment(localize('assistantView.routineAttachment', "Routine · {0}", mention.routine.name), mention.routine.prompt) + typedRequest;
					routineServerIds = mention.routine.serverIds;
				}
			}
			const imageUrls: string[] = [];
			if (submittedDraft.attachments.length) {
				const revision = draftRevision;
				const target = this.agentChats.target;
				const folders = await attachmentFolders();
				if (!folders || revision !== draftRevision) { return false; }
				const blocks: string[] = [];
				const skipped: string[] = [];
				let imageBytes = 0;
				for (const attachment of submittedDraft.attachments) {
					const name = attachment.fullName ?? attachment.name;
					if (isImageVariableEntry(attachment)) {
						const image = assistantImageAttachmentDataUrl(attachment);
						if (image && imageUrls.length < ASSISTANT_IMAGE_ATTACHMENT_LIMIT && imageBytes + image.bytes <= ASSISTANT_IMAGE_ATTACHMENT_BYTES_LIMIT) {
							imageUrls.push(image.url);
							imageBytes += image.bytes;
						} else {
							skipped.push(name);
						}
						continue;
					}
					if (URI.isUri(attachment.value)) {
						const label = formatAssistantWorkspacePath(folders, attachment.value);
						if (label === undefined) {
							skipped.push(name);
							continue;
						}
						try {
							const content = (await this.fileService.readFile(attachment.value, { limits: { size: 32 * 1024 } })).value.toString();
							const truncated = content.length > 12_000;
							blocks.push(`<file path="${label.replace(/"/g, '')}">\n${truncated ? `${content.slice(0, 12_000)}\n…(truncated)` : content}\n</file>\n\n`);
						} catch {
							skipped.push(name);
							continue;
						}
					} else if (typeof attachment.value === 'string') {
						const content = attachment.value.length > 12_000 ? `${attachment.value.slice(0, 12_000)}\n…(truncated)` : attachment.value;
						blocks.push(assistantTextAttachment(name, content));
					} else {
						skipped.push(name);
						continue;
					}
					if (this._store.isDisposed || revision !== draftRevision) { return false; }
				}
				if (this._store.isDisposed || revision !== draftRevision || this.agentChats.target !== target) { return false; }
				if (!blocks.length && !imageUrls.length) {
					setFeedback(localize('assistantView.textOnly', "Attachments couldn't be read as text or supported images — paste the content you want to share."));
					return false;
				}
				input = blocks.join('') + input;
				if (skipped.length) {
					setFeedback(localize('assistantView.attachmentsSkipped', "Attached {0} item(s); skipped {1}: only text content and supported images can be sent.", String(blocks.length + imageUrls.length), skipped.join(', ')));
				} else if (blocks.length) {
					setFeedback(localize('assistantView.attachmentsInlined', "Attached {0} item(s) as text.", blocks.length));
				} else {
					setFeedback(localize('assistantView.imagesAttached', "Attached {0} image(s).", imageUrls.length));
				}
			}
			const activeArtifact = activeArtifacts.get(artifactThreadKey(accountId, taskId));
			const activeDocument = activeArtifact && store.readDocument(accountId, activeArtifact.documentId);
			if (activeArtifact && activeDocument) {
				const selection = assistantVisibleDocumentSelection(this.editorService.visibleTextEditorControls, activeArtifact.resource);
				if (selection) {
					input = assistantPromptWithSelectionContext(input, activeDocument.title, selection);
				}
				input = assistantPromptWithDocumentContext(input, activeDocument.title, activeDocument.content);
			}
			if (input.length > REACH_ASSISTANT_PROMPT_LIMIT) {
				setFeedback(localize('assistantView.messageTooLong', "Assistant messages can be up to {0} characters.", REACH_ASSISTANT_PROMPT_LIMIT));
				return false;
			}
			const editing = !taskDraft && editingFrom?.accountId === accountId && editingFrom.taskId === taskId ? editingFrom : undefined;
			if (editing && this.runs.getPending(accountId, taskId)) {
				setFeedback(localize('assistantView.waitForEdit', "Wait for the current reply before replacing an earlier prompt."));
				return false;
			}
			if (editing && !equals(editing.original, store.read(accountId, taskId).turns)) {
				setEditingFrom(undefined);
				setFeedback(localize('assistantView.editHistoryChanged', "This conversation changed. Your draft is here. Select the prompt you want to replace, or send this as a new message."));
				return false;
			}
			setEditingFrom(undefined);
			const requestStart = input.length > typedRequest.length ? input.length - typedRequest.length : undefined;
			if (taskDraft) {
				saveDraft.cancel();
				setComposerDraft();
				const linkedTask = createTask(accountId);
				openThread(linkedTask);
				store.setComposerDraft(accountId, taskDraft, linkedTask, input);
				void this.runs.send(accountId, input, linkedTask, { requestStart, imageUrls, serverIds: routineServerIds });
				return true;
			}
			if (this.runs.getPending(accountId, taskId)) {
				if (!this.runs.queuePrompt(accountId, input, taskId, requestStart, imageUrls, routineServerIds)) { return false; }
				saveDraft.cancel();
				setComposerDraft();
				this.saveDraft?.();
				return true;
			}
			saveDraft.cancel();
			store.setComposerDraft(accountId, submittedDraft, taskId, submittedDraft.attachments.length ? input : undefined);
			setComposerDraft();
			void this.runs.send(accountId, input, taskId, { replaceFrom: editing?.index, requestStart, imageUrls, serverIds: routineServerIds });
			return true;
		};
		const addQuote = (reply: string) => {
			if (!accountId || this.runs.getPending(accountId, taskId)) { return; }
			const existing = this.composer?.getInputValue() ?? '';
			const quote = reply.split('\n').map(line => '> ' + line).join('\n');
			const next = (existing ? existing + '\n\n' : '') + quote + '\n\n';
			if (next.length > REACH_ASSISTANT_PROMPT_LIMIT) {
				setFeedback(localize('assistantView.quoteTooLong', "This reply is too long to quote. Copy the part you want to use."));
				return;
			}
			setInput(next);
			this.saveDraft?.();
			this.composer?.focus();
		};
		const pickWorkspaceFiles = async (folders: readonly IAssistantWorkspaceFolder[]): Promise<string[]> => {
			if (!folders.length) {
				setFeedback(localize('assistantView.noWorkspace', "Open a workspace folder to attach files."));
				return [];
			}
			const result = await this.searchService.fileSearch({
				type: QueryType.File,
				folderQueries: folders.map(({ uri }) => ({ folder: uri })),
				maxResults: 250,
				sortByScore: true,
			});
			const items = result.results.map(match => {
				const label = formatAssistantWorkspacePath(folders, match.resource);
				return label ? { label } : undefined;
			}).filter((item): item is { label: string } => !!item);
			const activeFile = this.agentChats.target?.activeEditorFile;
			const activePath = activeFile && formatAssistantWorkspacePath(folders, activeFile);
			const picks = await this.quickInput.pick(activePath ? [
				{ label: activePath, description: localize('assistantView.attachActiveEditorDetail', "Active in the IDE when you switched") },
				...items.filter(item => item.label !== activePath),
			] : items, {
				canPickMany: true,
				title: localize('assistantView.attachPickTitle', "Attach workspace files"),
				placeHolder: localize('assistantView.attachPick', "Type to filter; check files to attach them together"),
			});
			return picks?.map(pick => pick.label) ?? [];
		};
		const updateModelLabel = () => {
			const selected = accountId && store.read(accountId, taskId).modelId;
			const label = selected ? models.find(model => model.id === selected)?.name ?? selected : localize('assistantView.autoModel', "Auto");
			modelLabel.textContent = label;
			modelButton.title = localize('assistantView.modelPicker', "Choose a REACH model");
			modelButton.setAttribute('aria-label', localize('assistantView.currentModel', "Choose a REACH model. Current model: {0}", label));
		};
		const updatePermissionLabel = () => {
			const mode = accountId ? store.read(accountId, taskId).permissionMode ?? 'routine' : 'routine';
			const labels: Record<AssistantPermissionMode, string> = {
				routine: localize('assistantView.permissionRoutine', "Auto"),
				safeOnly: localize('assistantView.permissionSafeOnly', "Safe only"),
				allowAll: localize('assistantView.permissionAllowAll', "Allow all"),
			};
			permissionLabel.textContent = labels[mode];
			permissionButton.title = localize('assistantView.permissionPicker', "Choose which assistant actions run without asking");
			permissionButton.setAttribute('aria-label', localize('assistantView.currentPermission', "Assistant permissions. Current mode: {0}", labels[mode]));
		};
		const activityIcon = (status: IAssistantActivity['status']) =>
			status === 'done' ? Codicon.check : status === 'rejected' ? Codicon.close : status === 'failed' ? Codicon.warning : status === 'running' ? Codicon.sync : Codicon.circle;
		const toolKindIcon = (tool: string) =>
			/^assistant\.(read_file|list_files)$/.test(tool) ? Codicon.file
				: /^assistant\.search_files$/.test(tool) ? Codicon.search
					: /^assistant\.(write_file|edit_file)$/.test(tool) ? Codicon.edit
						: /^assistant\.(web_fetch|get_news_headlines|get_weather_forecast)$/.test(tool) ? Codicon.globe
							: /^assistant\.(create_task|start_task|list_tasks|read_task|rename_task|report_progress)$/.test(tool) ? Codicon.checklist
								: tool.startsWith('assistant.') ? Codicon.tools
									: Codicon.plug;
		const formatDuration = (ms: number): string => {
			const s = Math.round(ms / 1000);
			if (s < 60) { return `${s}s`; }
			const m = Math.floor(s / 60);
			return m < 60 ? `${m}m ${s % 60}s` : `${Math.floor(m / 60)}h ${m % 60}m`;
		};
		const renderActivity = (container: HTMLElement, activity: readonly IAssistantActivity[], store: DisposableStore) => {
			store.clear();
			clearNode(container);
			const collapsed = activity.length > 4;
			const rest = collapsed ? append(container, $('span.reach-assistant-activity-more', { hidden: true })) : undefined;
			const renderOne = (item: IAssistantActivity, host: HTMLElement) => {
				const hasDetail = !!item.detail;
				const chip = append(host, $(`${hasDetail ? 'button.reach-assistant-activity-chip-button' : 'span'}.reach-assistant-activity-chip.status-${item.status}`));
				const duration = item.durationMs !== undefined ? ` · ${formatDuration(item.durationMs)}` : '';
				if (hasDetail) {
					chip.setAttribute('type', 'button');
					chip.setAttribute('aria-expanded', 'false');
					chip.title = `${item.label} — ${item.status}${duration}. ${localize('assistantView.activityExpand', "Click for details.")}`;
				} else {
					chip.title = `${item.label} — ${item.status}${duration}`;
				}
				append(chip, $('span' + ThemeIcon.asCSSSelector(toolKindIcon(item.tool)), { 'aria-hidden': 'true' }));
				append(chip, $('span', {}, item.status === 'rejected' ? `${item.label} (${localize('assistantView.declined', "declined")})` : item.label));
				append(chip, $('span' + ThemeIcon.asCSSSelector(activityIcon(item.status)), { 'aria-hidden': 'true' }));
				if (item.durationMs !== undefined && item.durationMs >= 1000) {
					append(chip, $('span.reach-assistant-activity-duration', {}, formatDuration(item.durationMs)));
				}
				if (hasDetail) {
					const detailRow = append(host, $('div.reach-assistant-activity-detail', { hidden: true }, item.detail));
					store.add(addDisposableListener(chip, EventType.CLICK, () => {
						const show = detailRow.hidden;
						detailRow.hidden = !show;
						chip.setAttribute('aria-expanded', String(show));
					}));
				}
			};
			activity.forEach((item, i) => renderOne(item, i >= 3 && rest ? rest : container));
			if (rest) {
				const extra = activity.length - 3;
				const toggle = append(container, $('button.reach-assistant-activity-chip.reach-assistant-activity-toggle',
					{ type: 'button', 'aria-expanded': 'false' },
					localize('assistantView.moreTools', "+{0} more", extra)));
				store.add(addDisposableListener(toggle, EventType.CLICK, () => {
					const show = rest.hidden;
					rest.hidden = !show;
					toggle.setAttribute('aria-expanded', String(show));
					toggle.textContent = show ? localize('assistantView.lessTools', "Show less") : localize('assistantView.moreTools', "+{0} more", extra);
				}));
			}
		};
		const renderSuggestions = () => {
			suggestionListeners.clear();
			clearNode(suggestions);
			clearNode(followups);
			followups.hidden = true;
			if (!accountId) { return; }
			const editorTarget = this.agentChats.target;
			const folders = getAssistantWorkspaceFolders(this.workspaceService, editorTarget);
			const capturedSelection = this.agentChats.editorSelection.get();
			const selectionPath = capturedSelection && formatAssistantWorkspacePath(folders, capturedSelection.file);
			const editorChips: { label: string; title?: string; run: () => void }[] = [];
			if (capturedSelection && selectionPath) {
				const { file, selection: { text, startLine, endLine, truncated } } = capturedSelection;
				const lines = startLine === endLine ? String(startLine) : `${startLine}–${endLine}`;
				const fileLabel = `${basename(file)}:${lines}`;
				const pathLabel = `${selectionPath}:${lines}`;
				const contextId = `ide-selection:${file.toString()}:${lines}:${hash(text)}`;
				editorChips.push({
					label: localize('assistantView.attachIDESelection', "Attach IDE selection · {0}", fileLabel),
					title: truncated
						? localize('assistantView.attachIDESelectionTruncated', "Attach the first {0} characters of the selection in {1}", text.length, pathLabel)
						: localize('assistantView.attachIDESelectionPath', "Attach selected text from {0}", pathLabel),
					run: () => {
						if (this.agentChats.target === editorTarget && this.agentChats.editorSelection.get() === capturedSelection) {
							this.composer?.attachTextContext(pathLabel, text, Codicon.code, contextId);
							this.composer?.focus();
						}
					},
				});
			}
			const editorFile = selectionPath ? capturedSelection?.file : editorTarget?.activeEditorFile;
			const editorPath = editorFile && formatAssistantWorkspacePath(folders, editorFile);
			if (editorFile && editorPath) {
				editorChips.push({
					label: localize('assistantView.attachActiveEditor', "Attach {0} from IDE", basename(editorFile)),
					title: localize('assistantView.attachActiveEditorPath', "Attach {0} from IDE", editorPath),
					run: () => {
						if (this.agentChats.target === editorTarget) {
							this.composer?.attach([editorFile]);
							this.composer?.focus();
						}
					},
				});
			}
			const wire = (container: HTMLElement, chips: { label: string; title?: string; prompt?: string; run?: () => void }[]) => {
				for (const chip of chips.slice(0, 3)) {
					const button = append(container, $('button.reach-assistant-suggestion', { type: 'button' }, chip.label));
					if (chip.title) { button.title = chip.title; }
					suggestionListeners.add(addDisposableListener(button, EventType.CLICK, () => {
						if (!accountId || this.runs.getPending(accountId, taskId)) { return; }
						if (chip.run) { chip.run(); return; }
						if (chip.prompt) {
							this.saveDraft?.();
							void this.runs.send(accountId, chip.prompt, taskId, { preserveDraft: true });
						}
					}));
				}
			};
			if (taskId || turnCount > 0) {
				if (this.runs.getPending(accountId, taskId)) { return; }
				const task = taskId ? store.read(accountId).tasks?.find(task => task.id === taskId) : undefined;
				if (task?.status === 'running' && (task.lastPrompt ?? task.title) && !store.read(accountId, taskId).startedQueuedPrompt) {
					wire(followups, [{ label: resumeTaskLabel, run: () => resumeTask(task) }, ...editorChips]);
					followups.hidden = false;
					return;
				}
				if (this.runs.getFailedPrompt(accountId, taskId)) {
					wire(followups, [{ label: localize('assistantView.followupDiagnose', "What went wrong?"), prompt: localize('assistantView.followupDiagnosePrompt', "The previous attempt failed — explain what went wrong and how to fix it.") }, ...editorChips]);
					followups.hidden = false;
					return;
				}
				const last = store.read(accountId, taskId).turns.at(-1);
				const done = new Set((last?.activity ?? []).filter(item => item.status === 'done').map(item => item.tool));
				const chips: { label: string; title?: string; prompt?: string; run?: () => void }[] = [...editorChips];
				if (done.has('assistant.write_file') || done.has('assistant.edit_file')) {
					chips.push({ label: localize('assistantView.followupChanges', "Summarize the changes"), prompt: localize('assistantView.followupChangesPrompt', "Summarize the changes you just made to my files.") });
				}
				if (done.has('assistant.create_task')) {
					chips.push({ label: localize('assistantView.followupTasks', "Check on my tasks"), prompt: localize('assistantView.followupTasksPrompt', "What are my linked tasks up to?") });
				}
				if (done.has('assistant.search_files') || done.has('assistant.read_file') || done.has('assistant.find_files')) {
					chips.push({ label: localize('assistantView.followupRelated', "Find related files"), prompt: localize('assistantView.followupRelatedPrompt', "What other files in the workspace look related?") });
				}
				wire(followups, chips);
				followups.hidden = chips.length === 0;
				return;
			}
			const hour = new Date().getHours();
			const daypart = hour < 12 ? 'morning' : hour < 18 ? 'afternoon' : 'evening';
			const greetingName = assistantGreetingName(store.readMemoryBank(accountId).user);
			welcomeTitle.textContent = daypart === 'morning'
				? greetingName ? localize('assistantView.greetingMorningNamed', "Good morning, {0}. What should we tackle?", greetingName) : localize('assistantView.greetingMorning', "Good morning. What should we tackle?")
				: daypart === 'afternoon'
					? greetingName ? localize('assistantView.greetingAfternoonNamed', "Good afternoon, {0}. What should we tackle?", greetingName) : localize('assistantView.greetingAfternoon', "Good afternoon. What should we tackle?")
					: greetingName ? localize('assistantView.greetingEveningNamed', "Good evening, {0}. What should we tackle?", greetingName) : localize('assistantView.greetingEvening', "Good evening. What should we tackle?");
			const runningTasks = store.read(accountId).tasks?.filter(task => task.status === 'running').length ?? 0;
			const mailAlerts = store.read(accountId).mailAlerts?.length ?? 0;
			const attention = [
				runningTasks ? localize('assistantView.greetingTasks', "{0} linked task(s) running", runningTasks) : '',
				mailAlerts ? localize('assistantView.greetingMail', "{0} important email alert(s)", mailAlerts) : '',
			].filter(Boolean).join(' · ');
			welcomeDetail.textContent = attention || (memoryEntries(store.readMemoryBank(accountId).user).length ? localize('assistantView.greetingMemory', "I’ll keep your saved preferences in mind.") : '');
			welcomeDetail.hidden = !welcomeDetail.textContent;
			const enabledServers = this.mcpService.servers.get()
				.filter(server => isContributionEnabled(server.enablement.get()));
			const serverNames = enabledServers
				.map(server => (server.serverMetadata.get()?.serverName ?? server.definition.label).toLowerCase())
				.join(' ');
			const chips: { label: string; title?: string; prompt?: string; run?: () => void }[] = [...editorChips];
			for (const server of enabledServers.filter(server => server.prompts.get().length > 0).slice(0, 2)) {
				const name = server.serverMetadata.get()?.serverName ?? server.definition.label;
				chips.push({ label: localize('assistantView.suggestAppPrompts', "{0} prompts…", name), run: () => void browsePrompts(server) });
			}
			if (/mail|inbox/.test(serverNames) || this.directMail.connections.get().some(mail => !mail.needsSignIn)) { chips.push({ label: localize('assistantView.suggestEmail', "Summarize my recent emails"), prompt: localize('assistantView.suggestEmailPrompt', "Summarize my 5 most recent emails, with senders and subjects.") }); }
			if (/calendar/.test(serverNames)) { chips.push({ label: localize('assistantView.suggestCalendar', "What's on my calendar today?"), prompt: localize('assistantView.suggestCalendarPrompt', "What's on my calendar today? List the events with times.") }); }
			if (/slack|teams|chat|message/.test(serverNames)) { chips.push({ label: localize('assistantView.suggestMessages', "Catch me up on unread messages"), prompt: localize('assistantView.suggestMessagesPrompt', "Catch me up on my unread messages.") }); }
			if (this.configuration.getValue<boolean>(PromptsConfig.USE_AGENT_SKILLS)) {
				chips.splice(2, 0, { label: localize('assistantView.suggestSkill', "Use a skill…"), run: () => void attachSkill(undefined, this.composer?.getInputValue() ?? '') });
			}
			chips.push({ label: localize('assistantView.suggestTools', "What can you do?"), prompt: localize('assistantView.suggestToolsPrompt', "List the tools you can use and what they can do.") });
			wire(suggestions, chips);
		};
		const connectionDot = (server: IMcpServer): { css: string; text: string; detail?: string } => {
			if (!isContributionEnabled(server.enablement.get())) {
				return { css: 'disabled', text: localize('assistantView.appDisabled', "Disabled") };
			}
			const state = server.connectionState.get();
			switch (state.state) {
				case McpConnectionState.Kind.Running: return { css: 'running', text: localize('assistantView.appConnected', "Connected") };
				case McpConnectionState.Kind.Starting: return { css: 'starting', text: localize('assistantView.appConnecting', "Connecting…") };
				case McpConnectionState.Kind.Error: return { css: 'error', text: localize('assistantView.appError', "Connection error"), detail: state.message };
				default: return McpConnectionState.needsSignIn(state)
					? { css: 'attention', text: McpConnectionState.toString(state) }
					: { css: 'stopped', text: localize('assistantView.appOnDemand', "On-demand") };
			}
		};
		const PINNED_SERVERS_KEY = 'sessions.assistant.pinnedServers';
		const pinnedServers = (): Set<string> => new Set(this.storage.get(PINNED_SERVERS_KEY, StorageScope.PROFILE, '').split(',').filter(Boolean));
		const setPinned = (id: string, pinned: boolean) => {
			const set = pinnedServers();
			if (pinned) { set.add(id); } else { set.delete(id); }
			this.storage.store(PINNED_SERVERS_KEY, [...set].join(','), StorageScope.PROFILE, StorageTarget.USER);
			renderConnections();
		};
		const renderConnections = () => {
			connectionListeners.clear();
			clearNode(apps);
			apps.hidden = !accountId;
			if (!accountId) { return; }
			const pinned = pinnedServers();
			const rank = (dot: { css: string }) => ({ running: 0, starting: 1, attention: 2, error: 3, stopped: 4, disabled: 5 } as const)[dot.css] ?? 6;
			for (const mail of this.directMail.connections.get()) {
				const css = mail.needsSignIn ? 'attention' : 'running';
				const status = mail.needsSignIn ? localize('assistantView.directMailSignIn', "Sign-in required")
					: mail.canSend ? localize('assistantView.directMailReadSend', "Read and send") : localize('assistantView.directMailReadOnly', "Read-only");
				const label = localize('assistantView.appAria', "{0} app, {1}. Manage connection.", mail.label, status);
				const tile = append(apps, $('button.reach-assistant-app.reach-assistant-app-' + css, { type: 'button', title: label, 'aria-label': label }));
				append(tile, $('span' + ThemeIcon.asCSSSelector(Codicon.mail), { 'aria-hidden': 'true' }));
				append(tile, $('span.reach-assistant-app-name', {}, mail.label));
				append(tile, $('span.reach-assistant-app-count', {}, status));
				append(tile, $('span.reach-assistant-app-dot.' + css, { 'aria-hidden': 'true' }));
				connectionListeners.add(addDisposableListener(tile, EventType.CLICK, () => this.runs.openAppSetup(signalCDirectMailConnections[mail.provider].id)));
			}
			for (const server of [...this.mcpService.servers.get()].sort((a, b) => Number(pinned.has(b.definition.id)) - Number(pinned.has(a.definition.id)) || rank(connectionDot(a)) - rank(connectionDot(b)))) {
				const dot = connectionDot(server);
				const name = server.serverMetadata.get()?.serverName ?? server.definition.label;
				const modelTools = isContributionEnabled(server.enablement.get())
					? server.tools.get().filter(tool => !!(tool.visibility & McpToolVisibility.Model))
					: [];
				const tools = modelTools.length;
				const toolNames = modelTools.slice(0, 4).map(tool => tool.definition.title ?? tool.definition.name).join(', ');
				const stateText = dot.detail ? `${dot.text}: ${dot.detail}` : dot.text;
				const prompts = server.prompts.get().length;
				const tooltipParts = [`${name} — ${stateText}`];
				if (tools) { tooltipParts.push(localize('assistantView.appTooltipTools', "{0} tools: {1}{2}", tools, toolNames, tools > 4 ? '…' : '')); }
				if (prompts) { tooltipParts.push(localize('assistantView.appTooltipPrompts', "{0} prompts", prompts)); }
				const signIn = dot.css === 'attention';
				const options = () => this.commands.executeCommand(McpCommandIds.ServerOptions, server.definition.id, { collectionId: server.collection.id });
				const tile = append(apps, $('button.reach-assistant-app.reach-assistant-app-' + dot.css, {
					type: 'button',
					title: signIn ? localize('assistantView.appSignInTooltip', "{0}. Click to sign in.", tooltipParts.join(', ')) : localize('assistantView.appTooltip', "{0}. Click to manage.", tooltipParts.join(', ')),
					'aria-label': signIn ? localize('assistantView.appSignInAria', "Sign in to {0}", name) : localize('assistantView.appAria', "{0} app, {1}. Manage connection.", name, stateText),
				}));
				const iconUrl = server.serverMetadata.get()?.icons?.getUrl(16)?.dark;
				if (iconUrl && (iconUrl.scheme === 'data' || iconUrl.scheme === 'https')) {
					append(tile, $('img.reach-assistant-app-icon', { src: iconUrl.toString(true), 'aria-hidden': 'true' }));
				} else {
					append(tile, $('span' + ThemeIcon.asCSSSelector(Codicon.plug), { 'aria-hidden': 'true' }));
				}
				append(tile, $('span.reach-assistant-app-name', {}, name));
				if (pinned.has(server.definition.id)) { append(tile, $('span' + ThemeIcon.asCSSSelector(Codicon.pinned) + '.reach-assistant-app-pin', { 'aria-hidden': 'true' })); }
				if (dot.css === 'attention') {
					append(tile, $('span.reach-assistant-app-count', {}, dot.text));
				} else if (tools) {
					append(tile, $('span.reach-assistant-app-count', {}, String(tools)));
				}
				append(tile, $('span.reach-assistant-app-dot.' + dot.css, { 'aria-hidden': 'true' }));
				connectionListeners.add(addDisposableListener(tile, EventType.CLICK, () => {
					if (McpConnectionState.needsSignIn(server.connectionState.get()) && isContributionEnabled(server.enablement.get())) {
						void server.start({ promptType: 'all-untrusted' });
					} else {
						void options();
					}
				}));
				connectionListeners.add(addDisposableListener(tile, EventType.CONTEXT_MENU, event => {
					EventHelper.stop(event, true);
					const state = server.connectionState.get();
					const running = McpConnectionState.isRunning(state);
					const needsSignIn = McpConnectionState.needsSignIn(state);
					const enabled = isContributionEnabled(server.enablement.get());
					this.contextMenu.showContextMenu({
						getAnchor: () => tile,
						getActions: () => [
							new Action('assistant.app.options', localize('assistantView.appOptions', "Options…"), undefined, true,
								options),
							running
								? new Action('assistant.app.stop', localize('assistantView.appStop', "Stop"), undefined, true,
									() => this.commands.executeCommand(McpCommandIds.StopServer, server.definition.id))
								: new Action('assistant.app.start', needsSignIn ? localize('assistantView.appSignIn', "Sign In…") : localize('assistantView.appStart', "Start"), undefined, enabled,
									() => server.start({ promptType: 'all-untrusted' })),
							new Action('assistant.app.restart', localize('assistantView.appRestart', "Restart"), undefined, running,
								() => this.commands.executeCommand(McpCommandIds.RestartServer, server.definition.id)),
							new Action('assistant.app.toggleEnabled', enabled ? localize('assistantView.appDisable', "Disable") : localize('assistantView.appEnable', "Enable"), undefined, true,
								async () => {
									this.mcpService.enablementModel.setEnabled(server.definition.id, enabled ? ContributionEnablementState.DisabledProfile : ContributionEnablementState.EnabledProfile);
									if (enabled) { await server.stop(); }
								}),
							new Separator(),
							new Action('assistant.app.output', localize('assistantView.appOutput', "Show Output"), undefined, true,
								() => server.showOutput()),
							new Action('assistant.app.usePrompt', localize('assistantView.appUsePrompt', "Use prompt…"), undefined, server.prompts.get().length > 0,
								() => browsePrompts(server)),
							new Action('assistant.app.attachResource', localize('assistantView.appAttachResource', "Attach resource…"), undefined, true,
								() => browseResources(server)),
							new Action('assistant.app.listTools', localize('assistantView.appListTools', "List tools…"), undefined, modelTools.length > 0,
								() => this.quickInput.pick(modelTools.map(tool => ({
									label: tool.definition.title ?? tool.definition.name,
									description: tool.definition.description,
									detail: safeMcpTool(tool, isWorkIqMcpServer(server))
										? localize('assistantView.toolSafe', "Read-only — runs automatically in Auto mode")
										: localize('assistantView.toolSideEffects', "Can have side effects — may ask before running"),
								})), { title: localize('assistantView.appToolsTitle', "{0} tools", name), matchOnDescription: true }).then(() => undefined)),
							(() => {
								const allowed = new Set(accountId ? store.read(accountId, taskId).allowedTools ?? [] : []);
								const trusted = modelTools.length > 0 && modelTools.every(tool => allowed.has(tool.id));
								return new Action('assistant.app.trust', trusted ? localize('assistantView.appUntrust', "Require approval again") : localize('assistantView.appTrust', "Trust all tools"), undefined, modelTools.length > 0 && !!accountId,
									() => {
										if (!accountId) { return; }
										const next = new Set(store.read(accountId, taskId).allowedTools ?? []);
										for (const tool of modelTools) { trusted ? next.delete(tool.id) : next.add(tool.id); }
										store.setAllowedTools(accountId, [...next], taskId);
										setFeedback(trusted
											? localize('assistantView.appUntrusted', "{0} tools will ask before running again.", name)
											: localize('assistantView.appTrusted', "{0} tools will run without asking in this thread.", name));
									});
							})(),
							new Separator(),
							new Action('assistant.app.pin', pinned.has(server.definition.id) ? localize('assistantView.appUnpin', "Unpin app") : localize('assistantView.appPin', "Pin app"), undefined, true,
								() => setPinned(server.definition.id, !pinned.has(server.definition.id))),
							new Action('assistant.app.connectMore', localize('assistantView.connectApp', "Connect app"), undefined, true,
								() => this.runs.openAppSetup()),
						],
					});
				}));
			}
			const add = append(apps, $('button.reach-assistant-app.reach-assistant-app-add', {
				type: 'button',
				title: localize('assistantView.connectAppTip', "Connect more apps so Assistant can use their tools"),
				'aria-label': localize('assistantView.connectAppAria', "Connect an app to Assistant"),
			}));
			append(add, $('span' + ThemeIcon.asCSSSelector(Codicon.add), { 'aria-hidden': 'true' }));
			append(add, $('span', {}, localize('assistantView.connectApp', "Connect app")));
			connectionListeners.add(addDisposableListener(add, EventType.CLICK,
				() => this.runs.openAppSetup()));
		};
		const renderAssistantMarkdown = (value: string) => renderMarkdown({ value, isTrusted: false }, {
			sanitizerConfig: { remoteImageIsAllowed: () => false },
			transformUri: (href, kind) => kind === 'image' ? '' : href,
			actionHandler: href => {
				if (matchesSomeScheme(href, Schemas.https, Schemas.http)) {
					void openLinkFromMarkdown(this.openerService, href, false);
				}
			},
		});
		const artifactEditorListeners = new Map<string, DisposableStore>();
		const artifactEditorDisposeWatcher = this._store.add(new MutableDisposable());
		const artifactThreadKey = (owner: string, thread: string | undefined) => `${owner}
${thread ?? ''}`;
		const activeArtifacts = new Map<string, { readonly documentId: string; readonly resource: string }>();
		const openReplyDocument = async (reply: string, artifact?: { readonly accountId: string; readonly documentId: string }) => {
			const untitledResource = artifact && URI.from({ scheme: Schemas.untitled, path: `/Assistant Document-${artifact.documentId}.md` });
			const documentModel = this.untitledTextEditors.create({
				initialValue: reply, languageId: 'markdown',
				...(untitledResource ? { untitledResource } : {}),
			});
			if (artifact) {
				artifactEditorDisposeWatcher.value ??= this.untitledTextEditors.onWillDispose(model => {
					const key = model.resource.toString();
					artifactEditorListeners.get(key)?.dispose();
					artifactEditorListeners.delete(key);
					for (const [thread, active] of activeArtifacts) {
						if (active.resource === key) { activeArtifacts.delete(thread); }
					}
				});
				const key = documentModel.resource.toString();
				if (!artifactEditorListeners.has(key)) {
					const listeners = new DisposableStore();
					const sync = listeners.add(new RunOnceScheduler(() => {
						const value = this.untitledTextEditors.getValue(documentModel.resource);
						const current = store.readDocument(artifact.accountId, artifact.documentId);
						if (value !== undefined && current && value !== current.content) {
							store.saveDocument(artifact.accountId, current.title, value, artifact.documentId);
						}
					}, 500));
					listeners.add(documentModel.onDidChangeContent(() => sync.schedule()));
					artifactEditorListeners.set(key, listeners);
				}
				if (accountId) {
					activeArtifacts.set(artifactThreadKey(accountId, taskId), { documentId: artifact.documentId, resource: key });
				}
			}
			const input = this.instantiation.createInstance(UntitledTextEditorInput, documentModel);
			await this.editorService.openEditor(input, { pinned: true }, artifact ? SIDE_GROUP : undefined);
		};
		const renderTurns = () => {
			if (!accountId) { return; }
			const owner = accountId;
			const thread = taskId;
			const turns = store.read(owner, thread).turns;
			const last = turns.at(-1);
			if (turns.length === turnCount && equals(last, lastRenderedTurn)) { return; }
			lastRenderedTurn = last;
			renderedTurns.clear();
			clearNode(transcript);
			const first = searchQuery ? 0 : Math.max(0, assistantSummaryIndex(turns));
			for (let turnIndex = first; turnIndex < turns.length; turnIndex++) {
				const turn = turns[turnIndex];
				const item = append(transcript, $('article'));
				if (turn.routineRunId) { item.dataset.routineRunId = turn.routineRunId; }
				item.hidden = !!searchQuery && !turn.prompt.toLowerCase().includes(searchQuery) && !turn.reply.toLowerCase().includes(searchQuery);
				const attachments: { kind: string; label: string }[] = [];
				const messageText = turn.prompt.replace(/<(file|attachment|resource)\s+(?:path|name)="([^"]*)">[\s\S]*?<\/\1>\s*/g, (_m, kind: string, label: string) => { attachments.push({ kind, label }); return ''; }).trim();
				if (attachments.length) {
					const row = append(item, $('.reach-assistant-attachments'));
					row.setAttribute('role', 'list');
					for (const attachment of attachments) {
						const el = append(row, $('span.reach-assistant-attachment'));
						el.setAttribute('role', 'listitem');
						el.setAttribute('title', attachment.label);
						append(el, $(`span.codicon.codicon-${attachment.kind === 'file' ? 'file' : attachment.kind === 'resource' ? 'database' : 'note'}`));
						append(el, $('span', {}, attachment.label.split(/[\\/]/).pop() ?? attachment.label));
					}
				}
				if (turn.imageUrls?.length) {
					const row = append(item, $('.reach-assistant-images'));
					row.setAttribute('role', 'list');
					row.setAttribute('aria-label', localize('assistantView.attachedImages', "Attached images"));
					turn.imageUrls.forEach((url, index) => {
						const image = document.createElement('img');
						image.className = 'reach-assistant-image';
						image.src = url;
						image.alt = localize('assistantView.attachedImage', "Attached image {0}", index + 1);
						image.loading = 'lazy';
						row.appendChild(image);
					});
				}
				if (messageText) {
					append(item, $('p.reach-assistant-message', {}, messageText)).setAttribute('aria-label', localize('assistantView.userMessage', "Your message"));
				}
				const reply = renderedTurns.add(renderAssistantMarkdown(turn.reply));
				reply.element.classList.add('reach-assistant-reply');
				append(item, reply.element);
				const libraryDocumentIds = [...new Set((turn.activity ?? []).map(activity => activity.documentId).filter((id): id is string => typeof id === 'string'))];
				for (const documentId of libraryDocumentIds) {
					const artifact = store.readDocument(owner, documentId);
					if (!artifact) { continue; }
					const documentButton = append(item, $('button.reach-assistant-document', {
						type: 'button',
						title: localize('assistantView.openSavedDocumentTip', "Open the saved Markdown document"),
					}));
					append(documentButton, $('span' + ThemeIcon.asCSSSelector(Codicon.file), { 'aria-hidden': 'true' }));
					const documentText = append(documentButton, $('span.reach-assistant-document-text'));
					append(documentText, $('span.reach-assistant-document-title', {}, artifact.title));
					append(documentText, $('span.reach-assistant-document-meta', {}, localize('assistantView.savedDocumentMeta', "Saved document · {0} words · Open", artifact.content.match(/\S+/g)?.length ?? 0)));
					renderedTurns.add(addDisposableListener(documentButton, EventType.CLICK, () => {
						const latest = store.readDocument(owner, documentId);
						if (latest) { void openReplyDocument(latest.content, { accountId: owner, documentId }); }
						else if (accountId === owner && taskId === thread) { setFeedback(localize('assistantView.documentMissing', "This document is no longer available.")); }
					}));
				}
				const replyDocument = assistantReplyDocument(turn.reply);
				if (replyDocument) {
					const documentButton = append(item, $('button.reach-assistant-document', {
						type: 'button',
						title: localize('assistantView.openDocumentTip', "Open this reply as a Markdown document"),
					}));
					append(documentButton, $('span' + ThemeIcon.asCSSSelector(Codicon.file), { 'aria-hidden': 'true' }));
					const documentText = append(documentButton, $('span.reach-assistant-document-text'));
					append(documentText, $('span.reach-assistant-document-title', {}, replyDocument.title ?? localize('assistantView.documentTitle', "Assistant document")));
					append(documentText, $('span.reach-assistant-document-meta', {}, localize('assistantView.documentMeta', "{0} words · Open document", replyDocument.words)));
					renderedTurns.add(addDisposableListener(documentButton, EventType.CLICK, () => void openReplyDocument(turn.reply)));
				}
				const recoveryConnections = Array.isArray(turn.unavailableConnections)
					? turn.unavailableConnections.filter(connection => typeof connection?.id === 'string' && typeof connection.name === 'string')
					: [];
				const routineId = recoveryConnections.length && !thread ? turn.routineId : undefined;
				let recovery: HTMLElement | undefined;
				const addRecoveryAction = (label: string, run: () => void) => {
					recovery ??= append(item, $('.reach-assistant-reply-tools', {
						role: 'group', 'aria-label': localize('assistantView.reportActions', "Report recovery"),
					}));
					const button = append(recovery, $('button.reach-assistant-link', { type: 'button' }, label));
					renderedTurns.add(addDisposableListener(button, EventType.CLICK, () => {
						if (accountId === owner && taskId === thread) { run(); }
					}));
				};
				for (const { id, name, needsSignIn } of recoveryConnections) {
					addRecoveryAction(localize('assistantView.manageReportConnection', "Manage {0}", name), () => {
						const direct = this.directMail.connections.get().find(connection => connection.id === id);
						if (direct) { this.runs.openAppSetup(signalCDirectMailConnections[direct.provider].id); return; }
						const server = this.mcpService.servers.get().find(candidate => assistantConnectionKey(candidate) === id);
						if (!server) {
							setFeedback(localize('assistantView.reportConnectionRemoved', "This connection has been removed."));
							return;
						}
						void this.commands.executeCommand(McpCommandIds.ServerOptions, server.definition.id, { collectionId: server.collection.id, needsSignIn: needsSignIn === true });
					});
				}
				if (routineId) {
					addRecoveryAction(localize('assistantView.rerunRoutine', "Run routine again"), () => void runRoutine(owner, routineId));
				}
				const routineDrafts = Array.isArray(turn.routineDrafts) ? turn.routineDrafts : turn.routineDraft ? [turn.routineDraft] : [];
				for (const routineDraft of routineDrafts) {
					const review = append(item, $('button.reach-assistant-link', { type: 'button' }, localize('assistantView.reviewRoutine', "Review routine: {0}", routineDraft.name)));
					renderedTurns.add(addDisposableListener(review, EventType.CLICK, () => {
						if (accountId && !taskId) { void saveRoutine(accountId, routineDraft); }
					}));
				}
				const tools = append(item, $('.reach-assistant-reply-tools'));
				const actions = renderedTurns.add(new ActionBar(tools, { ariaLabel: localize('assistantView.replyActions', "Reply actions") }));
				let pointerPassage: { text: string | undefined } | undefined;
				renderedTurns.add(addDisposableListener(tools, EventType.POINTER_DOWN, event => {
					pointerPassage = isHTMLElement(event.target) && actions.getAction(event.target)?.id === 'assistant.saveIdea'
						? { text: assistantReplyPassage(reply.element, turn.reply) } : undefined;
				}, true));
				renderedTurns.add(addDisposableListener(tools, EventType.KEY_DOWN, () => { pointerPassage = undefined; }));
				renderedTurns.add(addDisposableListener(tools, 'pointercancel', () => { pointerPassage = undefined; }));
				actions.push([
					renderedTurns.add(new Action('assistant.saveIdea', localize('assistantView.saveReply', "Save reply"), ThemeIcon.asClassName(Codicon.bookmark), true, () => {
						if (accountId !== owner || taskId !== thread) { return; }
						const text = pointerPassage ? pointerPassage.text : assistantReplyPassage(reply.element, turn.reply);
						pointerPassage = undefined;
						if (text === undefined) { setFeedback(localize('assistantView.ideaSelection', "Select a passage within a single reply.")); return; }
						saveIdea(owner, thread, text);
					})),
					renderedTurns.add(new Action('assistant.openDocument', localize('assistantView.openDocument', "Open as document"), ThemeIcon.asClassName(Codicon.file), true, () => openReplyDocument(turn.reply))),
					renderedTurns.add(new Action('assistant.copyReply', localize('assistantView.copy', "Copy reply"), ThemeIcon.asClassName(Codicon.copy), true, async () => {
						await this.clipboard.writeText(turn.reply);
						setFeedback(localize('assistantView.copied', "Reply copied."));
					})),
					renderedTurns.add(new Action('assistant.copyPrompt', localize('assistantView.copyPrompt', "Copy prompt"), ThemeIcon.asClassName(Codicon.clippy), true, async () => {
						await this.clipboard.writeText(turn.prompt);
						setFeedback(localize('assistantView.promptCopied', "Prompt copied."));
					})),
					...(!routineId ? [renderedTurns.add(new Action('assistant.regenHere', turn.routineId ? localize('assistantView.rerunRoutine', "Run routine again") : localize('assistantView.regenHere', "Regenerate from here"), ThemeIcon.asClassName(Codicon.refresh), true, () => {
						if (!accountId) { return; }
						if (turn.routineId) {
							void runRoutine(accountId, turn.routineId);
							return;
						}
						if (this.runs.getPending(accountId, taskId)) {
							setFeedback(localize('assistantView.waitForReply', "Wait for the current reply first."));
							return;
						}
						void this.runs.send(accountId, turn.prompt, taskId, { requestStart: turn.requestStart, replaceFrom: turnIndex, preserveDraft: true });
					}))] : []),
					renderedTurns.add(new Action('assistant.editPrompt', turn.routineId ? localize('assistantView.routines', "Assistant routines") : localize('assistantView.editPrompt', "Edit prompt in composer"), ThemeIcon.asClassName(Codicon.edit), true, () => {
						if (turn.routineId && accountId) {
							void manageRoutines(accountId);
							return;
						}
						if (accountId) { setEditingFrom({ index: turnIndex, original: turns, accountId, taskId }); }
						setInput(turn.prompt);
						this.composer?.focus();
					})),
					renderedTurns.add(new Action('assistant.quoteReply', localize('assistantView.quote', "Quote in next message"), ThemeIcon.asClassName(Codicon.quote), true, () => addQuote(turn.reply))),
				], { icon: true, label: false });
				if (turn.model || turn.at || turn.durationMs !== undefined) {
					const parts: string[] = [];
					if (turn.model) { parts.push(turn.model); }
					if (turn.at) { parts.push(fromNow(turn.at, true)); }
					if (turn.durationMs !== undefined) { parts.push(localize('assistantView.workedFor', "worked {0}", formatDuration(turn.durationMs))); }
					const meta = append(tools, $('span.reach-assistant-muted', {}, parts.join(' · ')));
					if (turn.at) { meta.title = new Date(turn.at).toLocaleString(); }
				}
				if (turn.activity?.length) {
					renderActivity(append(item, $('.reach-assistant-activity')), turn.activity, renderedTurns.add(new DisposableStore()));
				}
				if (turn.plan?.length) {
					const done = turn.plan.filter(step => step.status === 'done').length;
					const planDetails = append(item, $('details.reach-assistant-plan-details'));
					append(planDetails, $('summary', {}, localize('assistantView.planSummary', "Plan — {0}/{1} done", done, turn.plan.length)));
					const planList = append(planDetails, $('ul.reach-assistant-plan'));
					for (const step of turn.plan) {
						const li = append(planList, $('li.reach-assistant-plan-item.' + step.status));
						append(li, $('span' + ThemeIcon.asCSSSelector(step.status === 'done' ? Codicon.check : step.status === 'inProgress' ? Codicon.sync : Codicon.circleOutline), { 'aria-hidden': 'true' }));
						append(li, $('span', {}, step.step));
					}
				}
			}
			turnCount = turns.length;
		};
		const renderTasks = () => {
			taskListeners.clear();
			clearNode(taskUpdates);
			const tasks = accountId ? store.read(accountId).tasks ?? [] : [];
			for (const task of tasks) {
				const previous = lastTaskStatuses.get(task.id);
				lastTaskStatuses.set(task.id, task.status);
				if (previous === 'running' && task.status !== 'running' && task.status !== 'draft' && task.id !== taskId) {
					void this.signalService.playSignal(task.status === 'completed' ? AccessibilitySignal.taskCompleted : AccessibilitySignal.taskFailed, { source: 'assistantSession' });
					setFeedback(task.status === 'completed'
						? localize('assistantView.taskFinished', "Task \"{0}\" finished — open it to read the reply.", task.title || localize('assistantView.newTask', "New task"))
						: localize('assistantView.taskEnded', "Task \"{0}\" ended with status: {1}.", task.title || localize('assistantView.newTask', "New task"), task.status));
				}
			}
			taskUpdates.hidden = !!taskId || tasks.length === 0;
			if (taskUpdates.hidden || !accountId) { return; }
			append(taskUpdates, $('h3', {}, localize('assistantView.tasks', "Linked tasks")));
			const statuses = {
				draft: localize('assistantTask.draft', "Draft"),
				running: localize('assistantTask.running', "Running"),
				completed: localize('assistantTask.completed', "Complete"),
				stopped: localize('assistantTask.stopped', "Stopped"),
				failed: localize('assistantTask.failed', "Needs attention"),
			};
			for (const task of [...tasks].sort((a, b) => Number(!!b.pinned) - Number(!!a.pinned))) {
				const row = append(taskUpdates, $('.reach-assistant-task'));
				const link = append(row, $('button.reach-assistant-link', { type: 'button' }, task.title || localize('assistantView.newTask', "New task")));
				taskListeners.add(addDisposableListener(link, EventType.CLICK, () => openThread(task.id)));
				const interrupted = task.status === 'running' && !this.runs.getPending(accountId, task.id);
				const needsApproval = !!this.runs.getPendingApproval(accountId, task.id);
				const taskStatus = needsApproval
					? localize('assistantTask.needsApproval', "Needs approval")
					: interrupted ? localize('assistantTask.interrupted', "Interrupted") : statuses[task.status];
				append(row, $('span.reach-assistant-muted', { role: 'status' }, taskStatus));
				if (task.update) { append(row, $('span.reach-assistant-muted.reach-assistant-task-update', { title: task.update }, task.update)); }
				const actions = append(row, $('span.reach-assistant-task-actions'));
				if (interrupted) {
					const resume = append(actions, $('button.reach-assistant-task-action', { type: 'button', title: resumeTaskLabel }));
					append(resume, $('span' + ThemeIcon.asCSSSelector(Codicon.play), { 'aria-hidden': 'true' }));
					taskListeners.add(addDisposableListener(resume, EventType.CLICK, () => resumeTask(task)));
				}
				const pin = append(actions, $('button.reach-assistant-task-action', { type: 'button', title: task.pinned ? localize('assistantView.unpinTask', "Unpin task") : localize('assistantView.pinTask', "Pin task") }));
				append(pin, $('span' + ThemeIcon.asCSSSelector(task.pinned ? Codicon.pinned : Codicon.pin), { 'aria-hidden': 'true' }));
				taskListeners.add(addDisposableListener(pin, EventType.CLICK, () => accountId && store.updateTask(accountId, task.id, { pinned: !task.pinned })));
				const rename = append(actions, $('button.reach-assistant-task-action', { type: 'button', title: localize('assistantView.renameTask', "Rename task") }));
				append(rename, $('span' + ThemeIcon.asCSSSelector(Codicon.edit), { 'aria-hidden': 'true' }));
				taskListeners.add(addDisposableListener(rename, EventType.CLICK, async () => {
					const owner = accountId;
					const check = this.accountCheck;
					if (!owner || check !== verifiedAccountCheck) { return; }
					const title = await this.quickInput.input({ title: localize('assistantView.renameTask', "Rename task"), value: task.title });
					if (title !== undefined && owner === accountId && check === this.accountCheck) { store.updateTask(owner, task.id, { title: title.trim() }); }
				}));
				const remove = append(actions, $('button.reach-assistant-task-action', { type: 'button', title: localize('assistantView.deleteTask', "Delete task") }));
				append(remove, $('span' + ThemeIcon.asCSSSelector(Codicon.trash), { 'aria-hidden': 'true' }));
				taskListeners.add(addDisposableListener(remove, EventType.CLICK, async () => {
					const owner = accountId;
					const check = this.accountCheck;
					if (!owner || check !== verifiedAccountCheck) { return; }
					const confirmed = await this.dialogs.confirm({
						message: localize('assistantView.deleteTaskConfirm', "Delete this task and its conversation?"),
						detail: task.title || undefined,
						primaryButton: localize('assistantView.deleteTaskButton', "Delete"),
					});
					if (confirmed.confirmed && owner === accountId && check === this.accountCheck) {
						this.runs.stop(owner, task.id);
						store.deleteTask(owner, task.id);
						if (taskId === task.id) { openThread(); }
					}
				}));
				if (task.update) { append(row, $('p.reach-assistant-task-update', {}, task.update)); }
			}
		};
		const pendingStatusText = (status: string | undefined) => {
			switch (status) {
				case 'thinking': return localize('assistantView.waiting', "Waiting for REACH…");
				case 'approval': return localize('assistantView.waitingApproval', "Waiting for your approval…");
				case 'finishing': return localize('assistantView.finishing', "Finishing…");
				default: return status?.startsWith('step ')
					? localize('assistantView.stepProgress', "Working (step {0})…", status.slice(5))
					: status ? localize('assistantView.usingTool', "Using {0}…", status) : '';
			}
		};
		const pendingMd = this._register(new MutableDisposable<ReturnType<typeof renderMarkdown>>());
		let streamedPartial = '';
		const streamRender = this._register(new RunOnceScheduler(() => {
			clearNode(pendingReply);
			pendingMd.clear();
			if (streamedPartial) {
				pendingMd.value = renderAssistantMarkdown(streamedPartial);
				append(pendingReply, pendingMd.value.element);
			}
		}, 350));
		const renderRun = () => {
			if (!accountId) { return; }
			const detail = this.runs.getPendingDetail(accountId, taskId);
			const busy = detail !== undefined;
			const nextApproval = this.runs.getPendingApproval(accountId, taskId);
			if (nextApproval && nextApproval.id !== activeApproval?.id) { void this.signalService.playSignal(AccessibilitySignal.chatUserActionRequired, { source: 'assistantSession' }); }
			activeApproval = nextApproval;
			pending.hidden = !busy;
			pendingPrompt.textContent = detail?.prompt ?? '';
			const partial = detail?.partial ?? '';
			pendingReply.hidden = !partial;
			if (partial !== streamedPartial) {
				streamedPartial = partial;
				streamRender.schedule();
			}
			if (partial) { thread.scrollTop = thread.scrollHeight; }
			const tick = () => {
				const base = pendingStatusText(detail?.status);
				const seconds = detail ? Math.max(0, Math.floor((Date.now() - detail.startedAt) / 1000)) : 0;
				pendingStatus.textContent = base && seconds > 2 ? `${base} · ${seconds}s` : base;
			};
			tick();
			elapsedTimer.cancel();
			if (busy) { elapsedTimer.cancelAndSet(tick, 1000); }
			const conversation = store.read(accountId, taskId);
			const waiting = conversation.queuedPrompts ?? [];
			const started = detail?.fromQueue ? undefined : conversation.startedQueuedPrompt;
			const nextQueued = started ?? waiting[0];
			const remaining = started ? waiting.length : waiting.length - 1;
			queuedRow.hidden = !nextQueued;
			queuedSend.hidden = busy;
			queuedSend.textContent = started ? retryQueuedLabel : sendQueuedLabel;
			const preview = nextQueued?.prompt.split('\n')[0].slice(0, 80) ?? '';
			const more = remaining > 0 ? localize('assistantView.queuedMore', " (+{0} more)", remaining) : '';
			queuedText.textContent = started ? localize('assistantView.queuedRetryPreview', "Ready to retry: {0}{1}", preview, more)
				: nextQueued ? localize('assistantView.queuedPreview', "Queued: {0}{1}", preview, more) : '';
			clearNode(pendingPlan);
			pendingPlan.hidden = !detail?.plan?.length;
			for (const item of detail?.plan ?? []) {
				const li = append(pendingPlan, $('li.reach-assistant-plan-item.' + item.status));
				append(li, $('span' + ThemeIcon.asCSSSelector(item.status === 'done' ? Codicon.check : item.status === 'inProgress' ? Codicon.sync : Codicon.circleOutline), { 'aria-hidden': 'true' }));
				append(li, $('span', {}, item.step));
			}
			renderActivity(pendingActivity, detail?.activity ?? [], activityListeners);
			if (activeApproval) {
				clearNode(approvalText);
				append(approvalText, $('span' + ThemeIcon.asCSSSelector(toolKindIcon(activeApproval.tool)), { 'aria-hidden': 'true' }));
				append(approvalText, $('span', {}, activeApproval.label));
				approvalArgs.textContent = activeApproval.argsPreview;
			}
			approval.hidden = !activeApproval;
			if (activeApproval) {
				const owner = this.mcpService.servers.get().find(server =>
					isContributionEnabled(server.enablement.get()) && server.tools.get().some(tool => tool.id === activeApproval!.tool));
				approvalTrust.hidden = !owner;
				if (owner) {
					const ownerName = owner.serverMetadata.get()?.serverName ?? owner.definition.label;
					approvalTrust.textContent = localize('assistantView.trustApp', "Always allow {0} tools", ownerName);
					approvalTrust.title = localize('assistantView.trustAppTip', "Run all current {0} tools in this thread without asking again", ownerName);
				}
			} else {
				approvalTrust.hidden = true;
			}
			approvalDetails.hidden = !activeApproval?.argsPreview;
			approvalDetails.toggleAttribute('open', (activeApproval?.argsPreview.length ?? 0) < 200);
			requestInProgress.set(busy, undefined);
			(modelButton as HTMLButtonElement).disabled = busy;
			const wasBusy = wasPending;
			if (wasBusy !== busy) {
				this.composer?.setInputEnabled(true);
				if (!busy) {
					if (wasBusy && !this.runs.getError(accountId, taskId)) { void this.signalService.playSignal(AccessibilitySignal.chatResponseReceived, { source: 'assistantSession' }); }
					streamedPartial = '';
					streamRender.cancel();
					pendingMd.clear();
					clearNode(pendingReply);
				}
				if (wasBusy && !this.composer?.getInputValue() && !this.composer?.attachments.length) {
					const draft = store.readComposerDraft(accountId, taskId);
					setComposerDraft(draft.inputText || draft.attachments.length || draft.references?.length ? draft : {
						inputText: this.runs.getFailedPrompt(accountId, taskId) ?? '', attachments: [],
					});
				}
				wasPending = busy;
			}
			this.canSend.set(true, undefined);
			welcome.hidden = busy || turnCount > 0 || !taskUpdates.hidden || !!taskId;
			const error = this.runs.getError(accountId, taskId);
			if (error || busy || wasBusy) {
				const routineId = !taskId ? this.runs.getFailedRoutine(accountId) : undefined;
				const messages = {
					signedOut: localize('assistantView.expired', "Your REACH sign-in expired. Reconnect from Agent Subscriptions."),
					subscriptionRequired: localize('assistantView.subscriptionRequired', "Assistant requires an active REACH Subscription."),
					allowanceExhausted: localize('assistantView.allowanceExhausted', "Your REACH subscription request allowance is exhausted."),
					noModel: localize('assistantView.noChatModel', "REACH didn’t list an eligible chat model for this account."),
					modelUnavailable: localize('assistantView.modelUnavailable', "This model is no longer available. Choose another model and send again."),
					historyChanged: localize('assistantView.historyChanged', "This conversation changed during regeneration. Select the reply you want to regenerate again."),
					busy: localize('assistantView.busy', "This conversation is already waiting for a reply."),
					stopped: routineId ? localize('assistantView.stoppedRoutine', "Routine stopped. You can run it again.")
						: started ? localize('assistantView.stoppedQueued', "Stopped. Your queued message is saved for retry.") : localize('assistantView.stopped', "Stopped. Your draft is ready to send again."),
					unavailable: localize('assistantView.failed', "REACH couldn’t complete this request. Send again to retry."),
					outputLimit: localize('assistantView.outputLimit', "The model reached its output limit before finishing. Ask for a shorter response or choose another model."),
				};
				setFeedback(error ? messages[error] : '', error && (routineId || this.runs.getFailedPrompt(accountId, taskId)) ? {
					label: localize('assistantView.retryLink', "Retry"),
					run: () => {
						if (!accountId) { return; }
						if (routineId) { void runRoutine(accountId, routineId); return; }
						const prompt = this.runs.getFailedPrompt(accountId, taskId);
						if (prompt) { setFeedback(''); void this.runs.send(accountId, prompt, taskId); }
					},
				} : undefined);
			}
			if (wasBusy !== busy) { thread.scrollTop = thread.scrollHeight; }
		};
		const renderState = () => {
			renderTurns();
			renderTasks();
			renderRun();
			renderSuggestions();
			renderConnections();
			if (taskId && accountId) {
				taskName.textContent = store.read(accountId).tasks?.find(task => task.id === taskId)?.title || localize('assistantView.newTask', "New task");
			}
			updateModelLabel();
			updatePermissionLabel();
			updateWorktreeLabel();
		};
		const openThread = (target?: string, focus = true, savePrevious = true) => {
			if (!accountId) { return; }
			savedIdeasPicker.clear();
			const account = accountId;
			if (target !== taskId) { this.composer?.cancelDictation(); }
			saveDraft.cancel();
			if (savePrevious) { this.saveDraft?.(); }
			const savedTask = target ? store.read(account).tasks?.find(task => task.id === target) : undefined;
			const id = taskId = savedTask?.id;
			store.setLastOpenTask(account, rootProject, id);
			const activeTarget = this.agentChats.target = readAssistantTaskTarget(savedTask) ?? rootTarget;
			turnCount = -1;
			searchQuery = '';
			wasPending = false;
			activeApproval = undefined;
			setEditingFrom(undefined);
			lastTaskStatuses.clear();
			setFeedback('');
			threadWatch.clear();
			if (id) {
				const disposables = new DisposableStore();
				threadWatch.value = disposables;
				store.onDidChange(account, disposables, id)(renderState, undefined, disposables);
			}
			breadcrumb.hidden = !id;
			newTask.hidden = !!id;
			const pending = this.runs.getPending(account, id);
			const submitted = pending && store.read(account, id).draft?.trim() === pending;
			setComposerDraft(submitted ? undefined : store.readComposerDraft(account, id));
			this.composer?.setInputEnabled(true);
			renderState();
			void worktreeError(activeTarget).then(error => {
				if (error && !this._store.isDisposed && account === accountId && id === taskId && this.agentChats.target === activeTarget) { setFeedback(error); }
			});
			thread.scrollTop = thread.scrollHeight;
			if (focus) { this.composer?.focus(); }
		};
		const rootWatch = this._register(new MutableDisposable<DisposableStore>());
		const refresh = async () => {
			const check = ++this.accountCheck;
			savedIdeasPicker.clear();
			const account = await this.reach.getAccount();
			if (check !== this.accountCheck) { return; }
			verifiedAccountCheck = check;
			const next = account?.subscriptionActive ? account.id : undefined;
			gate.hidden = !!next;
			manage.hidden = !!next;
			chat.hidden = !next;
			if (!next) {
				this.canSend.set(false, undefined);
				requestInProgress.set(false, undefined);
				gate.textContent = account ? localize('assistantView.subscriptionRequired', "Assistant requires an active REACH Subscription.")
					: localize('assistantView.signInRequired', "Sign in with your SignalREACH account to connect Assistant.");
			}
			if (next === accountId) { return; }
			this.composer?.cancelDictation();
			this.saveDraft?.();
			accountId = next;
			taskId = undefined;
			activeApproval = undefined;
			models = [];
			rootWatch.clear();
			threadWatch.clear();
			renderedTurns.clear();
			taskListeners.clear();
			clearNode(transcript);
			clearNode(taskUpdates);
			setComposerDraft();
			const ambientProject = openedTarget?.workspace ?? this.workspaceService.getWorkspace().folders[0]?.uri;
			const project = ambientProject ?? (next && restoreProject ? store.getLastWorktreeProject(next) : undefined);
			const selection = next && project?.scheme === Schemas.file ? store.getWorktreeSelection(next, project) : undefined;
			rootProject = selection?.project ?? project;
			rootTarget = openedTarget || selection || (!ambientProject && project) ? {
				...openedTarget,
				workspace: rootProject,
				worktree: selection?.worktree,
			} : undefined;
			this.agentChats.target = rootTarget;
			if (!next) { return; }
			store.setLastWorktreeProject(next, rootProject);
			const restoredTaskId = restoreTaskOnSignIn ? store.getLastOpenTask(next, rootProject) : undefined;
			restoreTaskOnSignIn = false;
			if (!this.composer) {
				const composer = this.composer = this._register(this.instantiation.createInstance(SharedChatInputComposer, {
					session: constObservable(undefined),
					loading: constObservable(false),
					canSendRequest: this.canSend,
					canSubmitWithoutSession: this.canSend,
					textOnly: true,
					hostRequest: { isRunning: requestInProgress, stop: stopRequest },
					slashCommands: assistantSlashCommands,
					mentionCompletions: async () => {
						const owner = accountId;
						if (!owner) { return []; }
						const scheduler = this.instantiation.invokeFunction(accessor => accessor.get(IAssistantBriefingScheduler));
						const routines = await scheduler.getRoutines(owner).catch(() => []);
						if (owner !== accountId) { return []; }
						return [...routines].sort((a, b) => a.name.localeCompare(b.name)).map(routine => ({
							label: routine.name,
							description: `${routine.enabled ? '' : `${localize('assistantView.routineDisabled', "Disabled")} · `}${routine.cadence === 'weekdays' ? localize('assistantView.routineWeekdays', "Weekdays") : localize('assistantView.routineDaily', "Daily")} · ${formatScheduleTime(routine)}`,
						}));
					},
					renderRepositoryControls: false,
					renderChatPet: false,
					voiceTarget: false,
					shouldPersistInput: () => false,
					onDidChangeDraft: () => {
						if (!settingInput) {
							draftRevision++;
							saveDraft.schedule();
						}
					},
					placeholder: localize('assistantView.prompt', "Message Assistant — / for commands"),
					submitOverride: submit,
					sendRequest: async () => false,
				}));
				composer.render(composerHost, chat, undefined, controls);
				this._register(composer.onDidChangeInput(() => {
					if (!settingInput) { saveDraft.schedule(); }
					updateCounter();
				}));
			}
			const disposables = new DisposableStore();
			rootWatch.value = disposables;
			store.onDidChange(next, disposables)(renderState, undefined, disposables);
			openThread(restoredTaskId, false, false);
			void this.reach.listAssistantModels(next).then(entries => {
				if (accountId === next) { models = entries; updateModelLabel(); }
			}, () => { });
		};

		this._register(addDisposableListener(permissionButton, EventType.CLICK, async () => {
			const currentAccount = accountId;
			const currentTask = taskId;
			if (!currentAccount) { return; }
			const picks: { label: string; description?: string; mode?: AssistantPermissionMode; manageTools?: boolean }[] = [
				{ label: localize('assistantView.permissionRoutine', "Auto"), description: localize('assistantView.permissionRoutineDescription', "Ask before actions with side effects, except for tools you’ve allowed in this thread"), mode: 'routine' },
				{ label: localize('assistantView.permissionSafeOnly', "Safe only"), description: localize('assistantView.permissionSafeOnlyDescription', "Automatically decline actions with side effects in this thread"), mode: 'safeOnly' },
				{ label: localize('assistantView.permissionAllowAll', "Allow all"), description: localize('assistantView.permissionAllowAllDescription', "Approve every tool call in this thread"), mode: 'allowAll' },
			];
			picks.push({ label: localize('assistantView.manageAllowedTools', "Allowed tools…"), description: localize('assistantView.manageAllowedToolsDescription', "Choose which tools may run without asking"), manageTools: true });
			const choice = await this.quickInput.pick(picks, { title: localize('assistantView.permissionPickerTitle', "Assistant permissions"), placeHolder: localize('assistantView.permissionPicker', "Choose which assistant actions run without asking") });
			if (!choice || accountId !== currentAccount || taskId !== currentTask) { return; }
			if (choice.manageTools) {
				const allowed = new Set(store.read(currentAccount, currentTask).allowedTools ?? []);
				const catalog = collectThreadTools(currentAccount, currentTask);
				const approvable = [...catalog.entries.values()].filter(entry => !entry.safe);
				if (!approvable.length) {
					setFeedback(localize('assistantView.noApprovableTools', "No tools need approval — everything available is read-only."));
					return;
				}
				const selected = await this.quickInput.pick(approvable.map(entry => ({
					label: entry.label,
					description: `${entry.name}${entry.kind === 'mcp' ? ' · MCP' : ''}`,
					picked: allowed.has(assistantToolApprovalKey(entry)),
					tool: assistantToolApprovalKey(entry),
				})), {
					canPickMany: true,
					title: localize('assistantView.allowedToolsTitle', "Tools that may run without asking"),
					placeHolder: localize('assistantView.allowedToolsPlaceHolder', "In Auto mode, checked tools can run without asking"),
				});
				if (selected && accountId === currentAccount && taskId === currentTask) {
					store.setAllowedTools(currentAccount, selected.map(item => item.tool), currentTask);
				}
				return;
			}
			if (choice.mode) {
				store.setPermissionMode(currentAccount, choice.mode, currentTask);
			}
		}));
		const respondApproval = (resolution: 'allow' | 'always' | 'reject', note?: string) => {
			if (activeApproval) {
				this.runs.respondApproval(activeApproval.id, resolution, note);
				activeApproval = undefined;
				renderRun();
			}
		};
		this._register(addDisposableListener(approvalAllow, EventType.CLICK, () => respondApproval('allow')));
		this._register(addDisposableListener(approvalAlways, EventType.CLICK, () => respondApproval('always')));
		this._register(addDisposableListener(approvalReject, EventType.CLICK, () => respondApproval('reject')));
		this._register(addDisposableListener(approvalRejectNote, EventType.CLICK, async () => {
			const note = await this.quickInput.input({ title: localize('assistantView.rejectNoteTitle', "Why reject? (sent to the assistant)"), placeHolder: localize('assistantView.rejectNotePlaceholder', "Optional reason — leave empty to reject silently") });
			if (note !== undefined) { respondApproval('reject', note || undefined); }
		}));
		this._register(addDisposableListener(approvalTrust, EventType.CLICK, () => {
			if (!activeApproval || !accountId) { return; }
			const owner = this.mcpService.servers.get().find(server =>
				isContributionEnabled(server.enablement.get()) && server.tools.get().some(tool => tool.id === activeApproval!.tool));
			if (!owner) { respondApproval('allow'); return; }
			const ownerName = owner.serverMetadata.get()?.serverName ?? owner.definition.label;
			const merged = new Set(store.read(accountId, taskId).allowedTools ?? []);
			for (const tool of owner.tools.get()) {
				if (!!(tool.visibility & McpToolVisibility.Model)) { merged.add(tool.id); }
			}
			store.setAllowedTools(accountId, [...merged], taskId);
			setFeedback(localize('assistantView.trustedApp', "{0} tools will run without asking in this thread.", ownerName));
			respondApproval('allow');
		}));
		this._register(this.runs.onDidRequestApproval(request => {
			if (request.accountId !== accountId || request.taskId !== taskId) { return; }
			renderRun();
			thread.scrollTop = thread.scrollHeight;
		}));
		this._register(addDisposableListener(newTask, EventType.CLICK, () => {
			if (accountId) { openThread(createTask(accountId)); }
		}));
		this._register(addDisposableListener(agentChatsButton, EventType.CLICK, () => void showAgentChats()));
		this._register(addDisposableListener(savedIdeasButton, EventType.CLICK, () => void browseSavedIdeas()));
		this._register(addDisposableListener(worktreeButton, EventType.CLICK, () => void pickWorktree()));
		this._register(addDisposableListener(back, EventType.CLICK, () => openThread()));
		this._register(addDisposableListener(queuedSend, EventType.CLICK, () => {
			if (accountId) { this.runs.sendQueuedPrompt(accountId, taskId); }
		}));
		this._register(addDisposableListener(queuedEdit, EventType.CLICK, () => {
			if (!accountId) { return; }
			if (this.composer?.getInputValue().trim()) {
				setFeedback(localize('assistantView.editQueueExistingDraft', "Finish or clear the current draft before editing a queued message."));
				return;
			}
			const next = store.takeQueuedPrompt(accountId, taskId, !this.runs.getPendingDetail(accountId, taskId)?.fromQueue);
			if (next) { setInput(next.prompt); this.saveDraft?.(); this.composer?.focus(); }
		}));
		this._register(addDisposableListener(queuedCancel, EventType.CLICK, () => {
			if (!accountId || !store.takeQueuedPrompt(accountId, taskId, !this.runs.getPendingDetail(accountId, taskId)?.fromQueue)) { return; }
			setFeedback(localize('assistantView.queueCleared', "Queued message removed."));
		}));
		this._register(addDisposableListener(editingCancel, EventType.CLICK, () => setEditingFrom(undefined)));
		this._register(addDisposableListener(chat, EventType.KEY_DOWN, event => {
			if (event.key === 'Escape' && accountId && this.runs.getPending(accountId, taskId)) {
				stopRequest();
				event.preventDefault();
				event.stopPropagation();
			} else if (event.key === 'Escape' && editingFrom) {
				setEditingFrom(undefined);
				event.preventDefault();
				event.stopPropagation();
			}
		}));
		this._register(addDisposableListener(modelButton, EventType.CLICK, async () => {
			const currentAccount = accountId;
			const currentTask = taskId;
			if (!currentAccount) { return; }
			try {
				models = await this.reach.listAssistantModels(currentAccount);
			} catch {
				setFeedback(localize('assistantView.modelsFailed', "Couldn’t load REACH models. Try opening the picker again."));
				return;
			}
			if (accountId !== currentAccount || taskId !== currentTask) { return; }
			const selectedModelId = store.read(currentAccount, currentTask).modelId;
			const items: (IQuickPickItem & { modelId?: string })[] = [
				{ label: `${selectedModelId ? '' : '$(check) '}${localize('assistantView.autoModel', "Auto")}`, description: localize('assistantView.autoModelDescription', "Use the first available REACH chat model") },
				...models.map(model => ({
					label: `${model.id === selectedModelId ? '$(check) ' : ''}${model.name}`,
					description: [
						model.id === model.name ? undefined : model.id,
						model.supportsTools === false ? localize('assistantView.modelTextTools', "text-based tool calls") : undefined,
					].filter(Boolean).join(' — ') || undefined,
					modelId: model.id,
				})),
			];
			const choice = await this.quickInput.pick<IQuickPickItem & { modelId?: string }>(items, { title: localize('assistantView.modelPickerTitle', "REACH models"), placeHolder: localize('assistantView.modelPicker', "Choose a REACH model"), matchOnDescription: true, activeItem: items.find(item => item.modelId === selectedModelId) ?? items[0] });
			if (choice && accountId === currentAccount && taskId === currentTask) {
				store.setModel(currentAccount, choice.modelId, currentTask);
				setFeedback('');
			}
		}));
		this._register(this.runs.onDidChange(changed => {
			if (changed === accountId) { renderState(); }
		}));
		this._register(autorun(reader => {
			this.agentChats.editorSelection.read(reader);
			renderSuggestions();
		}));
		const connectionRefresh = this._register(new RunOnceScheduler(renderConnections, 100));
		this._register(autorun(reader => {
			this.directMail.connections.read(reader);
			for (const server of this.mcpService.servers.read(reader)) {
				server.connectionState.read(reader);
				server.enablement.read(reader);
				server.tools.read(reader);
				server.prompts.read(reader);
				server.serverMetadata.read(reader);
			}
			connectionRefresh.schedule();
		}));
		this._register(this.reach.onDidChangeAccount(() => void refresh()));
		this._register(autorun(reader => {
			if (!this.routinesRequested.read(reader) || !this.canSend.read(reader)) { return; }
			// Let the new view finish laying out, or the account refresh finish restoring its draft.
			queueMicrotask(() => {
				if (this._store.isDisposed || !accountId || !this.routinesRequested.read(undefined)) { return; }
				this.routinesRequested.set(false, undefined);
				if (taskId) { openThread(undefined, false); }
				void manageRoutines(accountId);
			});
		}));
		void refresh();
	}

	layout(_width: number, height: number): void {
		if (this.content && this.composerHost) {
			this.content.style.height = Math.max(220, height - 20) + 'px';
			this.composer?.layout(height, this.composerHost.clientWidth);
		}
	}

	override focus(): void { this.composer?.focus(); }

	override dispose(): void {
		this.saveDraft?.();
		this.saveDraft = undefined;
		this.accountCheck++;
		super.dispose();
	}
}

export class AssistantSessionContribution extends Disposable {
	static readonly ID = 'sessions.assistantConversation';

	constructor(@ICustomViewService customViews: ICustomViewService) {
		super();
		const routinesRequested = observableValue(this, false);
		const clearChatRequested = this._register(new Emitter<void>());
		this._register(customViews.registerCustomView({
			id: ASSISTANT_CUSTOM_VIEW_ID,
			ctor: new SyncDescriptor(AssistantView, [routinesRequested, clearChatRequested.event]),
			actions: { style: 'buttonBar', menuId: AgentsWindowMenus.CustomViewAssistant },
			// The editor and the bottom panel stay usable beside the Assistant view, which covers only
			// the sessions grid and the side panel: its activity bar's tabs open in them, as a chat's do.
			coveredParts: [Parts.SESSIONS_PART, Parts.AUXILIARYBAR_PART],
			activityBar: { menuId: Menus.AssistantActivityBar, contextMenuId: Menus.AssistantActivityBarContext },
		}, { restore: true }));
		this._register(autorun(reader => {
			if (customViews.activeCustomView.read(reader)?.id !== ASSISTANT_CUSTOM_VIEW_ID) {
				routinesRequested.set(false, undefined);
			}
		}));
		this._register(CommandsRegistry.registerCommand(OPEN_ASSISTANT_ACTION_ID, accessor => {
			if (customViews.activeCustomView.get()?.id !== ASSISTANT_CUSTOM_VIEW_ID) {
				const host = accessor.get(IFocusedSessionHostService).getFocusedHost();
				const session = host.session.get();
				const workspace = (host.chat.get() ?? session?.mainChat.get())?.workspace.get() ?? session?.workspace.get();
				const composerService = accessor.get(INewSessionComposerService);
				const selection = composerService.getComposer(host)?.workspaceSelection;
				const target = accessor.get(IAssistantAgentChats);
				const project = workspace?.uri ?? (selection?.state === 'selected' ? selection.folderUri : undefined);
				const editor = target.editorWorkspace;
				const useEditorRoots = editor && !isSessionCreated(session) && !workspace?.isVirtualWorkspace && selection?.state === 'selected'
					&& (selection.origin === WorkspaceSelectionOrigin.WindowContext || selection.origin === WorkspaceSelectionOrigin.WindowOpen)
					&& extUriBiasedIgnorePathCase.isEqual(selection.folderUri, editor.selectedFolder)
					&& editor.selectionVersion === composerService.userWorkspaceSelectionVersion.get()
					&& editor.navigationVersion === composerService.userNavigationVersion.get();
				const selectedLocalFolder = !workspace && project?.scheme === Schemas.file
					&& (!editor || selection?.origin === WorkspaceSelectionOrigin.User) ? project : undefined;
				const fileRoots = useEditorRoots ? editor.fileRoots
					: workspace && !workspace.isVirtualWorkspace ? workspace.folders.map(folder => folder.workingDirectory)
						: selectedLocalFolder ? [selectedLocalFolder] : undefined;
				const fileRootNames = useEditorRoots ? editor.fileRootNames : workspace && !workspace.isVirtualWorkspace ? workspace.folders.map(folder => folder.name) : undefined;
				target.target = {
					workspace: project,
					...(fileRoots?.length ? { fileRoots, fileRootNames } : {}),
					...(useEditorRoots && editor.activeFile ? { activeEditorFile: editor.activeFile } : {}),
					providerId: session?.providerId,
					sessionTypeId: session?.sessionType,
				};
			}
			customViews.showCustomView(ASSISTANT_CUSTOM_VIEW_ID);
		}));
		this._register(registerAction2(class ManageAssistantRoutinesAction extends Action2 {
			constructor() {
				super({
					id: MANAGE_ASSISTANT_ROUTINES_ACTION_ID,
					title: localize2('assistantView.manageRoutinesAction', "Assistant Routines"),
					icon: Codicon.calendar,
					f1: true,
					precondition: IsSessionsWindowContext,
					menu: [
						{ id: AgentsWindowMenus.CustomViewAssistant, group: 'navigation', order: 1 },
						{ id: AgentsWindowMenus.CustomViewAutomations, group: 'navigation', order: 4, when: IsSessionsWindowContext },
					],
				});
			}

			override async run(accessor: ServicesAccessor): Promise<void> {
				if (customViews.activeCustomView.get()?.id !== ASSISTANT_CUSTOM_VIEW_ID) {
					await accessor.get(ICommandService).executeCommand(OPEN_ASSISTANT_ACTION_ID);
				}
				if (customViews.activeCustomView.get()?.id === ASSISTANT_CUSTOM_VIEW_ID) {
					routinesRequested.set(true, undefined);
				}
			}
		}));
		this._register(registerAction2(class ClearAssistantChatAction extends Action2 {
			constructor() {
				super({
					id: CLEAR_ASSISTANT_CHAT_ACTION_ID,
					title: localize2('assistantView.clearChatAction', "Clear Chat"),
					icon: Codicon.trash,
					precondition: IsSessionsWindowContext,
					menu: { id: AgentsWindowMenus.CustomViewAssistant, group: 'navigation', order: 2 },
				});
			}

			override run(): void {
				if (customViews.activeCustomView.get()?.id === ASSISTANT_CUSTOM_VIEW_ID) {
					clearChatRequested.fire();
				}
			}
		}));
	}
}

//#region Assistant Activity Bar

export const TOGGLE_ASSISTANT_BROWSER_ACTION_ID = 'sessions.assistant.toggleBrowser';
export const TOGGLE_ASSISTANT_TERMINAL_ACTION_ID = 'sessions.assistant.toggleTerminal';
const assistantActivityBarWhen = ContextKeyExpr.and(IsSessionsWindowContext, IsAuxiliaryWindowContext.toNegated());

registerAction2(class ToggleAssistantBrowserAction extends Action2 {
	constructor() {
		const when = ContextKeyExpr.and(assistantActivityBarWhen, DesktopLayoutContext);
		super({
			id: TOGGLE_ASSISTANT_BROWSER_ACTION_ID,
			title: localize2('assistantActivityBar.browser', "Browser"),
			icon: Codicon.globe,
			precondition: when,
			toggled: ContextKeyExpr.and(MainEditorAreaVisibleContext, ActiveEditorContext.isEqualTo(BrowserEditorInput.EDITOR_ID)),
			menu: { id: Menus.AssistantActivityBar, group: '1_browser', order: 1, when },
		});
	}

	override async run(accessor: ServicesAccessor): Promise<void> {
		const layout = accessor.get(IWorkbenchLayoutService);
		const group = accessor.get(IEditorGroupsService).mainPart.activeGroup;
		if (layout.isVisible(Parts.EDITOR_PART) && group.activeEditor instanceof BrowserEditorInput) {
			layout.setPartHidden(true, Parts.EDITOR_PART);
			layout.focusPart(Parts.CUSTOM_VIEW_GRID_PART);
			return;
		}
		await accessor.get(ICommandService).executeCommand(OPEN_BROWSER_COMMAND_ID);
	}
});

registerAction2(class ToggleAssistantTerminalAction extends Action2 {
	constructor() {
		const when = ContextKeyExpr.and(assistantActivityBarWhen, ContextKeyExpr.or(TerminalContextKeys.processSupported, TerminalContextKeys.terminalHasBeenCreated));
		super({
			id: TOGGLE_ASSISTANT_TERMINAL_ACTION_ID,
			title: localize2('assistantActivityBar.terminal', "Terminal"),
			icon: Codicon.terminal,
			precondition: when,
			toggled: ContextKeyExpr.and(PanelVisibleContext, ActivePanelContext.isEqualTo(TERMINAL_VIEW_ID)),
			menu: { id: Menus.AssistantActivityBar, group: '2_terminal', order: 1, when },
		});
	}

	override async run(accessor: ServicesAccessor): Promise<void> {
		const layout = accessor.get(IWorkbenchLayoutService);
		if (layout.isVisible(Parts.PANEL_PART) && accessor.get(IViewsService).isViewVisible(TERMINAL_VIEW_ID)) {
			layout.setPartHidden(true, Parts.PANEL_PART);
			layout.focusPart(Parts.CUSTOM_VIEW_GRID_PART);
			return;
		}
		const groups = accessor.get(ITerminalGroupService);
		if (!groups.activeInstance) {
			await accessor.get(ITerminalService).createTerminal({ location: TerminalLocation.Panel });
		}
		await groups.showPanel(true);
	}
});

//#endregion
