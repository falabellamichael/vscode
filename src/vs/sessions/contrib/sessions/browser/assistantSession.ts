/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { append, clearNode, $, EventHelper, EventType, addDisposableListener, isHTMLElement } from '../../../../base/browser/dom.js';
import { renderMarkdown } from '../../../../base/browser/markdownRenderer.js';
import { ActionBar } from '../../../../base/browser/ui/actionbar/actionbar.js';
import { Action, Separator } from '../../../../base/common/actions.js';
import { IntervalTimer, RunOnceScheduler } from '../../../../base/common/async.js';
import { Codicon } from '../../../../base/common/codicons.js';
import { hash } from '../../../../base/common/hash.js';
import { fromNow } from '../../../../base/common/date.js';
import { escapeIcons } from '../../../../base/common/iconLabels.js';
import { Disposable, DisposableStore, MutableDisposable } from '../../../../base/common/lifecycle.js';
import { autorun, constObservable, ISettableObservable, observableValue } from '../../../../base/common/observable.js';
import { basename, extUriBiasedIgnorePathCase } from '../../../../base/common/resources.js';
import { matchesSomeScheme, Schemas } from '../../../../base/common/network.js';
import { URI } from '../../../../base/common/uri.js';
import { ThemeIcon } from '../../../../base/common/themables.js';
import { localize, localize2 } from '../../../../nls.js';
import { Action2, registerAction2 } from '../../../../platform/actions/common/actions.js';
import { IClipboardService } from '../../../../platform/clipboard/common/clipboardService.js';
import { CommandsRegistry, ICommandService } from '../../../../platform/commands/common/commands.js';
import { IConfigurationService } from '../../../../platform/configuration/common/configuration.js';
import { IDialogService } from '../../../../platform/dialogs/common/dialogs.js';
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
import { AccessibilitySignal, IAccessibilitySignalService } from '../../../../platform/accessibilitySignal/browser/accessibilitySignalService.js';
import { IWorkspaceContextService } from '../../../../platform/workspace/common/workspace.js';
import { IWorkspaceTrustRequestService } from '../../../../platform/workspace/common/workspaceTrust.js';
import { IsSessionsWindowContext } from '../../../../workbench/common/contextkeys.js';
import { IChatDraft } from '../../../../workbench/contrib/chat/common/attachments/chatDraft.js';
import { ISearchService, QueryType } from '../../../../workbench/services/search/common/search.js';
import { isContributionEnabled } from '../../../../workbench/contrib/chat/common/enablement.js';
import { IGitService } from '../../../../workbench/contrib/git/common/gitService.js';
import { McpCommandIds } from '../../../../workbench/contrib/mcp/common/mcpCommandIds.js';
import { IMcpServer, IMcpService, McpConnectionState, McpToolVisibility } from '../../../../workbench/contrib/mcp/common/mcpTypes.js';
import { signalCDirectMailConnections } from '../../../../workbench/contrib/mcp/common/signalCAppConnections.js';
import { ISignalCDirectMail } from '../../../../workbench/contrib/mcp/browser/signalCDirectMail.js';
import type { IComposerSlashCommand } from '../../../../workbench/contrib/sessions/browser/chat/newChatComposerInput.js';
import { SharedChatInputComposer } from '../../../../workbench/contrib/sessions/browser/chat/sharedChatInputComposer.js';
import { AbstractCustomView, ICustomViewService } from '../../../../workbench/contrib/sessions/browser/customView.js';
import { INewSessionComposerService } from '../../../../workbench/contrib/sessions/browser/newSessionComposerService.js';
import { ASSISTANT_CUSTOM_VIEW_ID, OPEN_ASSISTANT_ACTION_ID } from '../../../../workbench/contrib/sessions/common/chatConstants.js';
import { SESSION_WORKTREE_ICON } from '../../../../workbench/contrib/sessions/common/sessionWorktreeIcon.js';
import { IFocusedSessionHostService } from '../../../../workbench/services/sessions/browser/sessionHost.js';
import { isSessionCreated, SessionStatus } from '../../../../workbench/services/sessions/common/session.js';
import { WorkspaceSelectionOrigin } from '../../../../workbench/services/sessions/common/workspaceSelection.js';
import { AgentsWindowMenus } from '../../../browser/agentsWindowMenus.js';
import { IReachSubscriptionConnection } from '../../accountMenu/browser/reachSubscriptionConnection.js';
import { AgentSubscriptionsAction } from '../../accountMenu/browser/subscriptionActions.js';
import { assistantWorktreeError, describeAssistantAgentStatus, IAssistantAgentChats, IAssistantAgentTarget } from './assistantAgentChats.js';
import { ASSISTANT_BRIEFINGS, AssistantBriefingPeriod } from './assistantBriefings.js';
import { IAssistantBriefingScheduler, IAssistantBriefingTime, IAssistantRoutine, IAssistantRoutineSettings, parseAssistantBriefingTime } from './assistantBriefingScheduler.js';
import { AssistantConversationStore, AssistantPermissionMode, IAssistantActivity, IAssistantTask, readAssistantTaskTarget, storeAssistantTaskTarget } from './assistantConversation.js';
import { IAssistantApprovalRequest, IAssistantRuns } from './assistantRuns.js';
import { assistantConnectionKey, captureAssistantTaskTarget, collectAssistantTools, findAssistantAppConnection, formatAssistantWorkspacePath, getAssistantWorkspaceFolders, IAssistantWorkspaceFolder, isWorkIqMcpServer, resolveAssistantWorkspacePath, safeMcpTool } from './assistantTools.js';
import { readAssistantCommonGitDirectory, readAssistantWorktreeIdentity } from './assistantWorktree.js';
import './media/assistantSession.css';

export const MANAGE_ASSISTANT_ROUTINES_ACTION_ID = 'sessionsView.manageAssistantRoutines';

const assistantSlashCommands: readonly IComposerSlashCommand[] = [
	{ label: '/help', description: localize('assistantView.helpHelp', "Browse Assistant commands") },
	{ label: '/task', description: localize('assistantView.helpTask', "Start a linked task: /task <description>"), acceptsArguments: true },
	{ label: '/tasks', description: localize('assistantView.helpTasks', "Jump to a linked task") },
	{ label: '/agents', description: localize('assistantView.helpAgents', "Open a regular chat started by Assistant") },
	{ label: '/worktree', description: localize('assistantView.helpWorktree', "Choose an existing Git worktree for Assistant") },
	{ label: '/attach', description: localize('assistantView.helpAttach', "Attach a workspace file (picker or path)"), acceptsArguments: true },
	{ label: '/model', description: localize('assistantView.helpModel', "Choose a REACH model") },
	{ label: '/connect', description: localize('assistantView.helpConnect', "Connect an app to Assistant, for example /connect Slack"), acceptsArguments: true },
	{ label: '/prompts', description: localize('assistantView.helpPrompts', "Use a prompt template from a connected app") },
	{ label: '/resources', description: localize('assistantView.helpResources', "Attach readable content published by a connected app") },
	{ label: '/tools', description: localize('assistantView.helpTools', "List the tools Assistant can use") },
	{ label: '/briefing', description: localize('assistantView.helpBriefing', "Run or schedule a morning or evening briefing"), acceptsArguments: true },
	{ label: '/weather', description: localize('assistantView.helpWeather', "Set up The Weather Network forecasts") },
	{ label: '/routine', description: localize('assistantView.helpRoutine', "Create, run, or manage recurring Assistant routines") },
	{ label: '/learn', description: localize('assistantView.helpLearn', "Learn from connected history and review proposed memories"), acceptsArguments: true },
	{ label: '/memory', description: localize('assistantView.helpMemory', "Show saved memory; /forget edits it") },
	{ label: '/forget', description: localize('assistantView.helpForget', "Choose which saved memories to keep") },
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

const formatScheduleTime = ({ hour, minute }: IAssistantBriefingTime) => `${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}`;

function describeLastRoutineRun(routine: IAssistantRoutine): string {
	const run = routine.lastRun;
	if (!run) { return localize('assistantView.routineNeverRun', "No recorded runs"); }
	const statuses = {
		running: localize('assistantView.routineRunning', "Running"),
		completed: localize('assistantView.routineReportReady', "Report ready"),
		needsAttention: localize('assistantView.routineNeedsAttention', "Needs attention"),
		failed: localize('assistantView.routineRunFailed', "Failed"),
		stopped: localize('assistantView.routineRunStopped', "Stopped"),
		interrupted: localize('assistantView.routineRunInterrupted', "Interrupted"),
	};
	const trigger = run.trigger === 'scheduled'
		? localize('assistantView.routineScheduledRun', "Scheduled run")
		: localize('assistantView.routineManualRun', "Manual run");
	return localize('assistantView.routineLastRun', "{0} · {1} · {2}", statuses[run.status], trigger, new Date(run.startedAt).toLocaleString(undefined, { timeZone: routine.timeZone }));
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
		const selected = await this.quickInput.pick(sources, {
			canPickMany: true,
			title: localize('assistantView.routineSources', "Choose connections for this routine"),
			placeHolder: localize('assistantView.routineSourcesHint', "Choose one or more connections. Each run uses read-only tools."),
		});
		if (!selected?.length || !current()) { return undefined; }
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
			prompt: localize('assistantView.routineTimePrompt', "Confirm a 24-hour time in {0} to save. Routines run while Agents is open.", timeZone),
			validateInput: async value => parseAssistantBriefingTime(value) ? undefined : localize('assistantView.briefingInvalidTime', "Enter a 24-hour time, for example 08:00."),
		});
		const time = value === undefined ? undefined : parseAssistantBriefingTime(value);
		return time && current() ? { ...time, name, prompt, cadence: schedule.cadence, serverIds: selected.map(source => source.id) } : undefined;
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
		append(welcome, $('h2', {}, localize('assistantView.welcome', "How can I help?")));
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
		const worktreeButton = append(controls, $<HTMLButtonElement>('button.reach-assistant-control.reach-assistant-worktree', { type: 'button', 'aria-haspopup': 'listbox', hidden: true }));
		append(worktreeButton, $('span' + ThemeIcon.asCSSSelector(SESSION_WORKTREE_ICON), { 'aria-hidden': 'true' }));
		const worktreeLabel = append(worktreeButton, $('span.reach-assistant-control-label'));
		const newTaskLabel = localize('assistantView.newTask', "New task");
		const resumeTaskLabel = localize('assistantView.resumeTask', "Resume task");
		const newTask = append(controls, $('button.reach-assistant-control', { type: 'button', 'aria-label': newTaskLabel }));
		newTask.title = localize('assistantView.newTaskTip', "Open a linked task conversation, or send /task followed by a task description.");
		append(newTask, $('span' + ThemeIcon.asCSSSelector(Codicon.add), { 'aria-hidden': 'true' }));
		append(newTask, $('span.reach-assistant-control-label', {}, newTaskLabel));
		const agentChatsLabel = localize('assistantView.agentChats', "Agent Chats");
		const agentChatsButton = append(controls, $('button.reach-assistant-control.reach-assistant-agent-chats', { type: 'button', 'aria-haspopup': 'listbox', 'aria-label': agentChatsLabel, title: agentChatsLabel, hidden: true }));
		append(agentChatsButton, $('span' + ThemeIcon.asCSSSelector(Codicon.commentDiscussion), { 'aria-hidden': 'true' }));
		append(agentChatsButton, $('span.reach-assistant-control-label', {}, agentChatsLabel));
		const feedback = append(dock, $('p.reach-assistant-feedback', { role: 'status', hidden: true }));
		const editingRow = append(dock, $('p.reach-assistant-editing', { role: 'status', hidden: true }));
		append(editingRow, $('span' + ThemeIcon.asCSSSelector(Codicon.edit), { 'aria-hidden': 'true' }));
		const editingText = append(editingRow, $('span'));
		const editingCancel = append(editingRow, $('button.reach-assistant-link', { type: 'button' }, localize('assistantView.cancelEdit', "Cancel edit")));
		const counter = append(dock, $('span.reach-assistant-counter', { role: 'status', hidden: true }));

		const threadWatch = this._register(new MutableDisposable<DisposableStore>());
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
		let editingFrom: { index: number; accountId: string; taskId: string | undefined } | undefined;
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
			worktreeButton.hidden = getWorktreeProject()?.scheme !== Schemas.file;
			worktreeLabel.textContent = selected?.label ?? localize('assistantView.worktree', "Worktree");
			worktreeButton.title = selected ? `${worktreeTitle}\n${selected.uri.fsPath}` : worktreeTitle;
			worktreeButton.setAttribute('aria-label', selected ? localize('assistantView.activeWorktree', "Choose Assistant worktree. Current worktree: {0}", selected.label) : worktreeTitle);
		};
		const pickWorktree = async () => {
			const workspace = getWorktreeProject();
			if (workspace?.scheme !== Schemas.file) {
				setFeedback(localize('assistantView.worktreeLocalProject', "Select a local Git project in Agents to choose a worktree."));
				return;
			}
			if (!accountId || worktreeButton.disabled) { return; }
			const target = this.agentChats.target;
			const account = accountId;
			const task = taskId;
			const taskRunning = () => !!task && !!this.runs.getPending(account, task);
			if (taskRunning()) {
				setFeedback(localize('assistantView.worktreeTaskRunning', "Stop the task or wait for its reply before changing its worktree."));
				return;
			}
			const stillCurrent = () => !this._store.isDisposed && accountId === account && taskId === task && this.agentChats.target === target && !taskRunning();
			worktreeButton.disabled = true;
			setFeedback(localize('assistantView.loadingWorktrees', "Finding Git worktrees…"));
			try {
				const repo = await this.gitService.openRepository(workspace);
				if (!stillCurrent()) { return; }
				if (!repo && !target?.worktree) {
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
				if (!choices.length && !target?.worktree) {
					setFeedback(localize('assistantView.noWorktrees', "No Git worktrees were found for this project."));
					return;
				}
				setFeedback('');
				const items: (IQuickPickItem & { worktree?: IAssistantAgentTarget['worktree'] })[] = choices.map(worktree => ({
					label: escapeIcons(worktree.label),
					description: extUriBiasedIgnorePathCase.isEqual(worktree.uri, target?.worktree?.uri) ? localize('assistantView.selectedWorktree', "Selected") : worktree.main ? localize('assistantView.mainWorktree', "Main checkout") : undefined,
					detail: escapeIcons(worktree.uri.fsPath),
					worktree,
				}));
				if (target?.worktree) {
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
				const nextTarget = this.agentChats.target = { ...target, workspace, worktree };
				if (task) {
					store.updateTask(account, task, { target: storeAssistantTaskTarget(captureAssistantTaskTarget(this.workspaceService, nextTarget)) });
				} else {
					rootTarget = nextTarget;
					rootProject = workspace;
					store.setWorktreeSelection(account, workspace, worktree);
					store.setLastWorktreeProject(account, workspace);
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
		const saveRoutine = async (owner: string, defaults?: Partial<IAssistantRoutineSettings>, routineId?: string) => {
			const current = () => !this._store.isDisposed && accountId === owner && !taskId;
			if (!this.directMail.connections.get().length && !this.mcpService.servers.get().some(server => isContributionEnabled(server.enablement.get()))) {
				setFeedback(localize('assistantView.routineNeedsConnection', "Connect an app before saving a routine."), {
					label: localize('assistantView.connectApp', "Connect app"), run: () => this.runs.openAppSetup(),
				});
				return false;
			}
			try {
				const settings = await this.pickRoutineSettings(current, defaults, !!routineId);
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
					default:
						updated = await saveRoutine(owner, routine, routine?.id);
						if (updated && current()) { clearRoutineCommand(); }
						return updated;
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
		const submit = async (): Promise<boolean> => {
			let input = this.composer?.getInputValue().trim();
			if (!accountId || !input) { return false; }
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
			if (input === '/memory') {
				const memory = store.read(accountId).instructions;
				setFeedback(memory ? localize('assistantView.memory', "Memory: {0}", memory.slice(0, 300)) : localize('assistantView.noMemory', "No memory saved yet. Ask the assistant to remember something."));
				setInput('');
				return true;
			}
			if (input === '/forget') {
				const memory = store.read(accountId).instructions;
				const lines = memory.split('\n').filter(line => line.trim());
				if (!lines.length) {
					setFeedback(localize('assistantView.noMemory', "No memory saved yet. Ask the assistant to remember something."));
				} else {
					const picks = await this.quickInput.pick(lines.map(line => ({ label: line.replace(/^-\s*/, ''), picked: true })), {
						canPickMany: true,
						title: localize('assistantView.forgetTitle', "Assistant memory"),
						placeHolder: localize('assistantView.forgetPlaceHolder', "Uncheck memories to forget, then confirm"),
					});
					if (picks) {
						store.setInstructions(accountId, picks.map(pick => `- ${pick.label}`).join('\n'));
						setFeedback(picks.length === lines.length
							? localize('assistantView.memoryKept', "Memory unchanged.")
							: localize('assistantView.memoryForgot', "Forgot {0} memor{1}.", lines.length - picks.length, lines.length - picks.length === 1 ? 'y' : 'ies'));
					}
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
				store.clearTurns(accountId, taskId);
				setFeedback(localize('assistantView.cleared', "Conversation cleared."));
				setInput('');
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
				setFeedback(localize('assistantView.contextInfo', "Next reply sees {0} of {1} turn(s), about {2} of {3} characters. Older turns drop out — /compact consolidates them.", stats.kept, stats.total, stats.chars.toLocaleString(), stats.limit.toLocaleString()));
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
				const last = store.popTurn(accountId, taskId);
				setFeedback(last ? '' : localize('assistantView.nothingToRegenerate', "Nothing to regenerate yet."));
				setInput('');
				if (last) { void this.runs.send(accountId, last.prompt, taskId); }
				return true;
			}
			if (input === '/compact') {
				const count = store.read(accountId, taskId).turns.length;
				if (count < 2) {
					setFeedback(localize('assistantView.nothingToCompact', "Nothing to compact yet."));
				} else {
					setInput('');
					const done = await this.runs.send(accountId, localize('assistantView.compactPrompt', "Summarize this conversation so far — key facts, decisions made, file paths touched and open questions — compactly, for future reference."), taskId);
					const summary = done ? store.read(accountId, taskId).turns.at(-1)?.reply : undefined;
					if (summary) {
						store.clearTurns(accountId, taskId);
						store.appendTurn(accountId, { prompt: localize('assistantView.compactLabel', "Summarize the conversation"), reply: summary }, taskId);
						setFeedback(localize('assistantView.compacted', "Compacted {0} turns into a summary.", count));
					} else if (!done) {
						setFeedback(localize('assistantView.compactBusy', "Wait for the current reply first."));
					}
				}
				setInput('');
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
			if (submittedDraft.attachments.length) {
				const revision = draftRevision;
				const folders = await attachmentFolders();
				if (!folders || revision !== draftRevision) { return false; }
				const blocks: string[] = [];
				const skipped: string[] = [];
				for (const attachment of submittedDraft.attachments) {
					const name = attachment.fullName ?? attachment.name;
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
						blocks.push(`<attachment name="${name.replace(/"/g, '')}">\n${content}\n</attachment>\n\n`);
					} else {
						skipped.push(name);
						continue;
					}
					if (this._store.isDisposed || revision !== draftRevision) { return false; }
				}
				if (this._store.isDisposed || revision !== draftRevision) { return false; }
				if (!blocks.length) {
					setFeedback(localize('assistantView.textOnly', "Attachments couldn't be read as text — paste the content you want to share."));
					return false;
				}
				input = blocks.join('') + input;
				if (skipped.length) {
					setFeedback(localize('assistantView.attachmentsSkipped', "Attached {0} item(s) as text; skipped {1}: only text content can be sent.", String(blocks.length), skipped.join(', ')));
				} else {
					setFeedback(localize('assistantView.attachmentsInlined', "Attached {0} item(s) as text.", blocks.length));
				}
			}
			if (input.length > REACH_ASSISTANT_PROMPT_LIMIT) {
				setFeedback(localize('assistantView.messageTooLong', "Assistant messages can be up to {0} characters.", REACH_ASSISTANT_PROMPT_LIMIT));
				return false;
			}
			if (editingFrom && editingFrom.accountId === accountId && editingFrom.taskId === taskId) {
				store.truncateTurns(accountId, editingFrom.index, taskId);
			}
			setEditingFrom(undefined);
			if (this.runs.getPending(accountId, taskId)) {
				if (!this.runs.queuePrompt(accountId, input, taskId)) { return false; }
				saveDraft.cancel();
				setComposerDraft();
				this.saveDraft?.();
				return true;
			}
			saveDraft.cancel();
			if (!taskId && input.startsWith('/task ')) {
				const taskPrompt = input.slice('/task '.length).trim();
				if (!taskPrompt) { return false; }
				setInput('');
				store.setDraft(accountId, '');
				openThread(createTask(accountId));
				setInput(taskPrompt);
				this.saveDraft?.();
				setInput('');
				void this.runs.send(accountId, taskPrompt, taskId);
				return true;
			}
			store.setComposerDraft(accountId, submittedDraft, taskId, submittedDraft.attachments.length ? input : undefined);
			setComposerDraft();
			void this.runs.send(accountId, input, taskId);
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
						setInput(chip.prompt ?? '');
						void submit();
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
					: { css: 'stopped', text: localize('assistantView.appDisconnected', "Disconnected") };
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
				const status = mail.needsSignIn ? localize('assistantView.directMailSignIn', "Sign-in required") : localize('assistantView.directMailReadOnly', "Read-only");
				const label = localize('assistantView.appAria', "{0} app, {1}. Manage connection.", mail.label, status);
				const tile = append(apps, $('button.reach-assistant-app.reach-assistant-app-' + css, { type: 'button', title: label, 'aria-label': label }));
				append(tile, $('span' + ThemeIcon.asCSSSelector(Codicon.mail), { 'aria-hidden': 'true' }));
				append(tile, $('span.reach-assistant-app-name', {}, mail.label));
				if (mail.needsSignIn) { append(tile, $('span.reach-assistant-app-count', {}, status)); }
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
				const tile = append(apps, $('button.reach-assistant-app.reach-assistant-app-' + dot.css, {
					type: 'button',
					title: localize('assistantView.appTooltip', "{0}. Click to manage.", tooltipParts.join(', ')),
					'aria-label': localize('assistantView.appAria', "{0} app, {1}. Manage connection.", name, stateText),
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
				connectionListeners.add(addDisposableListener(tile, EventType.CLICK,
					() => void this.commands.executeCommand(McpCommandIds.ServerOptions, server.definition.id)));
				connectionListeners.add(addDisposableListener(tile, EventType.CONTEXT_MENU, event => {
					EventHelper.stop(event, true);
					const state = server.connectionState.get();
					const running = McpConnectionState.isRunning(state);
					const needsSignIn = McpConnectionState.needsSignIn(state);
					this.contextMenu.showContextMenu({
						getAnchor: () => tile,
						getActions: () => [
							new Action('assistant.app.options', localize('assistantView.appOptions', "Options…"), undefined, true,
								() => this.commands.executeCommand(McpCommandIds.ServerOptions, server.definition.id)),
							running
								? new Action('assistant.app.stop', localize('assistantView.appStop', "Stop"), undefined, true,
									() => this.commands.executeCommand(McpCommandIds.StopServer, server.definition.id))
								: new Action('assistant.app.start', needsSignIn ? localize('assistantView.appSignIn', "Sign In…") : localize('assistantView.appStart', "Start"), undefined, true,
									() => this.commands.executeCommand(McpCommandIds.StartServer, server.definition.id)),
							new Action('assistant.app.restart', localize('assistantView.appRestart', "Restart"), undefined, running,
								() => this.commands.executeCommand(McpCommandIds.RestartServer, server.definition.id)),
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
		const renderTurns = () => {
			if (!accountId) { return; }
			const owner = accountId;
			const thread = taskId;
			const turns = store.read(owner, thread).turns;
			if (turns.length === turnCount) { return; }
			renderedTurns.clear();
			clearNode(transcript);
			for (const [turnIndex, turn] of turns.entries()) {
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
				if (messageText) {
					append(item, $('p.reach-assistant-message', {}, messageText)).setAttribute('aria-label', localize('assistantView.userMessage', "Your message"));
				}
				const reply = renderedTurns.add(renderAssistantMarkdown(turn.reply));
				reply.element.classList.add('reach-assistant-reply');
				append(item, reply.element);
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
				const routineDraft = turn.routineDraft;
				if (routineDraft) {
					const review = append(item, $('button.reach-assistant-link', { type: 'button' }, localize('assistantView.reviewRoutine', "Review routine: {0}", routineDraft.name)));
					renderedTurns.add(addDisposableListener(review, EventType.CLICK, () => {
						if (accountId && !taskId) { void saveRoutine(accountId, routineDraft); }
					}));
				}
				const tools = append(item, $('.reach-assistant-reply-tools'));
				const actions = renderedTurns.add(new ActionBar(tools, { ariaLabel: localize('assistantView.replyActions', "Reply actions") }));
				actions.push([
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
						store.truncateTurns(accountId, turnIndex, taskId);
						void this.runs.send(accountId, turn.prompt, taskId);
					}))] : []),
					renderedTurns.add(new Action('assistant.editPrompt', turn.routineId ? localize('assistantView.routines', "Assistant routines") : localize('assistantView.editPrompt', "Edit prompt in composer"), ThemeIcon.asClassName(Codicon.edit), true, () => {
						if (turn.routineId && accountId) {
							void manageRoutines(accountId);
							return;
						}
						if (accountId) { setEditingFrom({ index: turnIndex, accountId, taskId }); }
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
					if (!accountId) { return; }
					const title = await this.quickInput.input({ title: localize('assistantView.renameTask', "Rename task"), value: task.title });
					if (title !== undefined) { store.updateTask(accountId, task.id, { title: title.trim() }); }
				}));
				const remove = append(actions, $('button.reach-assistant-task-action', { type: 'button', title: localize('assistantView.deleteTask', "Delete task") }));
				append(remove, $('span' + ThemeIcon.asCSSSelector(Codicon.trash), { 'aria-hidden': 'true' }));
				taskListeners.add(addDisposableListener(remove, EventType.CLICK, async () => {
					if (!accountId) { return; }
					const confirmed = await this.dialogs.confirm({
						message: localize('assistantView.deleteTaskConfirm', "Delete this task and its conversation?"),
						detail: task.title || undefined,
						primaryButton: localize('assistantView.deleteTaskButton', "Delete"),
					});
					if (confirmed.confirmed) {
						this.runs.stop(accountId, task.id);
						store.deleteTask(accountId, task.id);
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
					busy: localize('assistantView.busy', "This conversation is already waiting for a reply."),
					stopped: routineId ? localize('assistantView.stoppedRoutine', "Routine stopped. You can run it again.")
						: started ? localize('assistantView.stoppedQueued', "Stopped. Your queued message is saved for retry.") : localize('assistantView.stopped', "Stopped. Your draft is ready to send again."),
					unavailable: localize('assistantView.failed', "REACH couldn’t complete this request. Send again to retry."),
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
			agentChatsButton.hidden = !accountId || !store.read(accountId).agentChats?.length;
			if (taskId && accountId) {
				taskName.textContent = store.read(accountId).tasks?.find(task => task.id === taskId)?.title || localize('assistantView.newTask', "New task");
			}
			updateModelLabel();
			updatePermissionLabel();
			updateWorktreeLabel();
		};
		const openThread = (target?: string, focus = true, savePrevious = true) => {
			if (!accountId) { return; }
			const account = accountId;
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
			const account = await this.reach.getAccount();
			if (check !== this.accountCheck) { return; }
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
					picked: allowed.has(entry.name),
					tool: entry.name,
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
						model.supportsTools === false ? localize('assistantView.modelNoTools', "no tool use") : undefined,
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
		this._register(customViews.registerCustomView({
			id: ASSISTANT_CUSTOM_VIEW_ID,
			ctor: new SyncDescriptor(AssistantView, [routinesRequested]),
			actions: { style: 'buttonBar', menuId: AgentsWindowMenus.CustomViewAssistant },
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
				const fileRoots = useEditorRoots ? editor.fileRoots : workspace && !workspace.isVirtualWorkspace ? workspace.folders.map(folder => folder.workingDirectory) : undefined;
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
						{ id: AgentsWindowMenus.CustomViewAssistant, group: 'navigation' },
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
	}
}
