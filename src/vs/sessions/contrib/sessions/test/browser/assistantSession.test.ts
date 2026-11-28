/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import assert from 'assert';
import { spy } from 'sinon';
import { DeferredPromise, timeout } from '../../../../../base/common/async.js';
import { VSBuffer } from '../../../../../base/common/buffer.js';
import { CancellationToken } from '../../../../../base/common/cancellation.js';
import { Codicon } from '../../../../../base/common/codicons.js';
import { Emitter, Event } from '../../../../../base/common/event.js';
import { Disposable } from '../../../../../base/common/lifecycle.js';
import { ResourceSet } from '../../../../../base/common/map.js';
import { constObservable, observableValue } from '../../../../../base/common/observable.js';
import { isWindows } from '../../../../../base/common/platform.js';
import { URI } from '../../../../../base/common/uri.js';
import { mock, upcastPartial } from '../../../../../base/test/common/mock.js';
import { runWithFakedTimers } from '../../../../../base/test/common/timeTravelScheduler.js';
import { ensureNoDisposablesAreLeakedInTestSuite } from '../../../../../base/test/common/utils.js';
import { isIMenuItem, MenuRegistry } from '../../../../../platform/actions/common/actions.js';
import { IClipboardService } from '../../../../../platform/clipboard/common/clipboardService.js';
import { CommandsRegistry, ICommandService } from '../../../../../platform/commands/common/commands.js';
import { ConfigurationTarget, IConfigurationService } from '../../../../../platform/configuration/common/configuration.js';
import { IDialogService } from '../../../../../platform/dialogs/common/dialogs.js';
import { TestInstantiationService } from '../../../../../platform/instantiation/test/common/instantiationServiceMock.js';
import { AssistantNewsScope, AssistantWeatherDay, AssistantWeatherForecast, AssistantWeatherLocation, REACH_ASSISTANT_PROMPT_LIMIT, ReachAssistantRequest, ReachAssistantResponse } from '../../../../../platform/native/common/native.js';
import { toUpstreamMessage } from '../../../../../platform/native/common/reachChatTransport.js';
import { IQuickInputHideEvent, IQuickInputService, IQuickPick, IQuickPickItem, QuickInputHideReason } from '../../../../../platform/quickinput/common/quickInput.js';
import { IOpenerService } from '../../../../../platform/opener/common/opener.js';
import { IFileService } from '../../../../../platform/files/common/files.js';
import { IRequestService } from '../../../../../platform/request/common/request.js';
import { InMemoryStorageService, IStorageService, StorageScope, StorageTarget, WillSaveStateReason } from '../../../../../platform/storage/common/storage.js';
import { IWorkspaceContextService } from '../../../../../platform/workspace/common/workspace.js';
import { IWorkspaceTrustRequestService } from '../../../../../platform/workspace/common/workspaceTrust.js';
import { ISearchService } from '../../../../../workbench/services/search/common/search.js';
import { AccessibilitySignal, IAccessibilitySignalService } from '../../../../../platform/accessibilitySignal/browser/accessibilitySignalService.js';
import { toFileVariableEntry, toPasteVariableEntry } from '../../../../../workbench/contrib/chat/common/attachments/chatVariableEntries.js';
import type { IChatDraft } from '../../../../../workbench/contrib/chat/common/attachments/chatDraft.js';
import { ContributionEnablementState } from '../../../../../workbench/contrib/chat/common/enablement.js';
import { PromptsConfig } from '../../../../../workbench/contrib/chat/common/promptSyntax/config/config.js';
import { PromptFileParser } from '../../../../../workbench/contrib/chat/common/promptSyntax/promptFileParser.js';
import { IPromptsService, PromptsStorage } from '../../../../../workbench/contrib/chat/common/promptSyntax/service/promptsService.js';
import { IGitService } from '../../../../../workbench/contrib/git/common/gitService.js';
import { McpCommandIds } from '../../../../../workbench/contrib/mcp/common/mcpCommandIds.js';
import { IAutostartResult, IMcpServer, IMcpService, IMcpTool, McpConnectionState, McpServerCacheState, McpServerDefinition, McpServerTransportType, McpToolVisibility } from '../../../../../workbench/contrib/mcp/common/mcpTypes.js';
import { microsoft365WorkIq } from '../../../../../workbench/contrib/mcp/common/signalCAppConnections.js';
import { ISignalCDirectMail } from '../../../../../workbench/contrib/mcp/browser/signalCDirectMail.js';
import { SharedChatInputComposer } from '../../../../../workbench/contrib/sessions/browser/chat/sharedChatInputComposer.js';
import { INewSessionComposer, INewSessionComposerService } from '../../../../../workbench/contrib/sessions/browser/newSessionComposerService.js';
import { AbstractCustomView, type ICustomViewDescriptor, ICustomViewService } from '../../../../../workbench/contrib/sessions/browser/customView.js';
import { Menus } from '../../../../../workbench/contrib/sessions/browser/menus.js';
import { Parts } from '../../../../../workbench/services/layout/browser/layoutService.js';
import { ASSISTANT_CUSTOM_VIEW_ID, OPEN_ASSISTANT_ACTION_ID } from '../../../../../workbench/contrib/sessions/common/chatConstants.js';
import { ASSISTANT_WEATHER_LOCATION_SETTING } from '../../../../../workbench/contrib/sessions/common/sessionConfig.js';
import { IFocusedSessionHostService, ISessionHost } from '../../../../../workbench/services/sessions/browser/sessionHost.js';
import { ISession, ISessionFolder, ISessionWorkspace } from '../../../../../workbench/services/sessions/common/session.js';
import { WorkspaceSelectionOrigin } from '../../../../../workbench/services/sessions/common/workspaceSelection.js';
import { AgentsWindowMenus } from '../../../../browser/agentsWindowMenus.js';
import { IReachAccount, IReachSubscriptionConnection } from '../../../accountMenu/browser/reachSubscriptionConnection.js';
import { IAssistantAgentChats } from '../../browser/assistantAgentChats.js';
import { IAssistantBriefingScheduler, IAssistantRoutine } from '../../browser/assistantBriefingScheduler.js';
import { ASSISTANT_BRIEFINGS } from '../../browser/assistantBriefings.js';
import { AssistantConversationStore, assistantSummaryIndex, readAssistantTaskTarget } from '../../browser/assistantConversation.js';
import { IAssistantMailWatcher } from '../../browser/assistantMailWatcher.js';
import { AssistantRuns, IAssistantApprovalRequest, IAssistantRuns } from '../../browser/assistantRuns.js';
import { assistantGreetingName, assistantImageAttachmentDataUrl, assistantPromptWithDocumentContext, assistantPromptWithSelectionContext, assistantReplyDocument, assistantVisibleDocumentSelection, assistantRoutineMention, assistantRoutineTemplates, assistantRoutineTemplateServerIds, assistantReplyPassage, AssistantSessionContribution, CLEAR_ASSISTANT_CHAT_ACTION_ID, getAssistantLinkedTaskDraft, MANAGE_ASSISTANT_ROUTINES_ACTION_ID, TOGGLE_ASSISTANT_BROWSER_ACTION_ID, TOGGLE_ASSISTANT_TERMINAL_ACTION_ID } from '../../browser/assistantSession.js';
import { assistantConnectionKey, collectAssistantTools, findAssistantAppConnection, formatAssistantWorkspacePath, getAssistantWorkspaceFolders, parseTextToolCalls, resolveAssistantWorkspacePath } from '../../browser/assistantTools.js';

function emptyMcpService(): IMcpService {
	return new class extends mock<IMcpService>() {
		override readonly servers = constObservable([]);
		override autostart() { return constObservable(IAutostartResult.Empty); }
	};
}

function fileServiceMock(files: Record<string, string> = {}): IFileService {
	return new class extends mock<IFileService>() {
		override async exists(resource: URI) { return Object.hasOwn(files, resource.path); }
		override async readFile(resource: URI) {
			const content = files[resource.path];
			if (content !== undefined) {
				return { value: VSBuffer.fromString(content) } as never;
			}
			throw new Error('File not found');
		}
		override async writeFile(resource: URI, buffer: VSBuffer) {
			files[resource.path] = buffer.toString();
			return undefined as never;
		}
	};
}

function workspaceServiceMock(): IWorkspaceContextService {
	return new class extends mock<IWorkspaceContextService>() {
		override getWorkspace() { return { folders: [{ uri: URI.parse('file:///ws') }] } as never; }
	};
}

function searchServiceMock(): ISearchService {
	return new class extends mock<ISearchService>() { };
}

const quickInputStub = new class extends mock<IQuickInputService>() { };

const openerStub = new class extends mock<IOpenerService>() { };
const agentChatsStub = upcastPartial<IAssistantAgentChats>({ target: undefined });
const assistantCommandStub = new class extends mock<ICommandService>() { };
const directMailStub = upcastPartial<ISignalCDirectMail>({ connections: constObservable([]) });
const configurationStub = upcastPartial<IConfigurationService>({ getValue: () => undefined });
const mailWatcherStub = upcastPartial<IAssistantMailWatcher>({
	start: () => undefined, isEnabled: () => true, onDidChange: Event.None,
	status: () => ({ enabled: true, connections: [], alerts: [] }),
	checkNow: async () => 0, setEnabled: async () => undefined, markSeen: () => undefined,
});
const promptsStub = upcastPartial<IPromptsService>({ findAgentSkills: async () => [], getDisabledPromptFiles: () => new ResourceSet() });

function mcpServiceWith(...servers: IMcpServer[]): IMcpService {
	return new class extends mock<IMcpService>() {
		override readonly servers = constObservable(servers);
		override autostart() { return constObservable(IAutostartResult.Empty); }
	};
}

function fakeServer(overrides: { name?: string; id?: string; collectionId?: string; url?: string; running?: boolean; state?: McpConnectionState; toolCount?: number; tools?: readonly IMcpTool[]; instructions?: string; onStart?: (setTools: (tools: readonly IMcpTool[]) => void) => void } = {}): IMcpServer {
	const tools = overrides.tools ?? Array.from({ length: overrides.toolCount ?? 2 }, (_, i) => new class extends mock<IMcpTool>() {
		override readonly visibility = McpToolVisibility.Model;
		override readonly id = `mcp_fake_${i}`;
		override readonly definition = { name: `mcp_fake_${i}`, description: 'Fake tool', inputSchema: { type: 'object' } } as never;
	});
	const initialState = overrides.state ?? { state: overrides.running === false ? McpConnectionState.Kind.Stopped : McpConnectionState.Kind.Running };
	const toolsStore = observableValue<readonly IMcpTool[]>('serverTools', tools);
	return new class extends mock<IMcpServer>() {
		override readonly collection = { id: overrides.collectionId ?? 'user', label: 'User' } as never;
		override readonly definition = { label: overrides.name ?? 'Notes', id: overrides.id ?? 'notes-server' } as never;
		override readonly enablement = constObservable(ContributionEnablementState.EnabledWorkspace);
		override readonly connectionState = observableValue<McpConnectionState>('connectionState', initialState);
		override readonly cacheState = observableValue('cacheState', initialState.state === McpConnectionState.Kind.Running ? McpServerCacheState.Live : McpServerCacheState.Unknown);
		override readonly serverMetadata = constObservable({ serverName: overrides.name ?? 'Notes', serverInstructions: overrides.instructions, icons: undefined as never });
		override readonly tools = toolsStore;
		override readonly prompts = constObservable([]);
		override async start() {
			overrides.onStart?.(next => toolsStore.set(next, undefined));
			if (this.connectionState.get().state === McpConnectionState.Kind.Error || McpConnectionState.needsSignIn(this.connectionState.get())) { return this.connectionState.get(); }
			this.connectionState.set({ state: McpConnectionState.Kind.Running }, undefined);
			this.cacheState.set(McpServerCacheState.Live, undefined);
			return this.connectionState.get();
		}
		override readDefinitions() {
			const server = overrides.url ? upcastPartial<McpServerDefinition>({
				launch: { type: McpServerTransportType.HTTP, uri: URI.parse(overrides.url), headers: [] },
			}) : undefined;
			return constObservable({ server, collection: undefined });
		}
	};
}

suite('AssistantSession', () => {
	const store = ensureNoDisposablesAreLeakedInTestSuite();

	test('uses only explicit naming preferences for the personalized greeting', () => {
		assert.strictEqual(assistantGreetingName('My name is Ada Lovelace\nPrefers concise replies.'), 'Ada Lovelace');
		assert.strictEqual(assistantGreetingName('Call me Rowan.'), 'Rowan');
		assert.strictEqual(assistantGreetingName('The user is a product manager.'), undefined);
	});

	test('identifies document-like replies and extracts their first heading', () => {
		assert.strictEqual(assistantReplyDocument('Short answer.'), undefined);
		assert.deepStrictEqual(assistantReplyDocument('# Launch Plan\n\n' + 'Detailed rollout notes. '.repeat(70)), { title: 'Launch Plan', words: 213 });
		assert.strictEqual(assistantReplyDocument('## One\n\nbody\n\n## Two\n\nmore')?.title, 'One');
	});

	test('fits open-document context into the fixed-size prompt without hiding the typed request', () => {
		const short = assistantPromptWithDocumentContext('Revise the opening.', 'Launch plan', '# Launch plan\n\nOpen with the customer outcome.');
		assert.match(short, /^<attachment name="Document · Launch plan">/);
		assert.ok(short.endsWith('Revise the opening.'));
		const long = assistantPromptWithDocumentContext('Revise the opening.', 'Launch plan', 'x'.repeat(20_000));
		assert.ok(long.length <= REACH_ASSISTANT_PROMPT_LIMIT);
		assert.match(long, /document truncated/);
		assert.ok(long.endsWith('Revise the opening.'));
		const crowded = 'x'.repeat(REACH_ASSISTANT_PROMPT_LIMIT - 100);
		assert.strictEqual(assistantPromptWithDocumentContext(crowded, 'Launch plan', 'document'), crowded);
	});

	test('captures a visible document selection and gives it bounded prompt context', () => {
		const resource = 'untitled:/Assistant Document-document-1.md';
		const control = {
			getEditorType: () => 'vs.editor.ICodeEditor',
			getModel: () => ({ uri: { toString: () => resource }, getValueInRange: () => 'selected paragraph' }),
			getSelection: () => ({ isEmpty: () => false, startLineNumber: 3, endLineNumber: 5 }),
		};
		assert.deepStrictEqual(assistantVisibleDocumentSelection([control], resource), { text: 'selected paragraph', startLine: 3, endLine: 5 });
		assert.strictEqual(assistantVisibleDocumentSelection([control], 'untitled:/other.md'), undefined);
		const empty = { ...control, getSelection: () => ({ isEmpty: () => true, startLineNumber: 3, endLineNumber: 3 }) };
		assert.strictEqual(assistantVisibleDocumentSelection([empty], resource), undefined);
		const input = assistantPromptWithSelectionContext('Tighten this.', 'Launch plan', { text: 'selected paragraph', startLine: 3, endLine: 5 });
		assert.match(input, /^<attachment name="Selected text · Launch plan · lines 3-5">/);
		assert.ok(input.endsWith('Tighten this.'));
	});

	test('preselects routine template connections by matching app labels', () => {
		const template = assistantRoutineTemplates.find(candidate => candidate.id === 'morning-briefing')!;
		assert.deepStrictEqual(assistantRoutineTemplateServerIds(template, [
			{ id: 'google-calendar', label: 'Google Calendar' },
			{ id: 'gmail', label: 'Gmail' },
			{ id: 'tracker', label: 'Project tracker' },
		]), ['google-calendar', 'gmail']);
	});

	test('resolves exact routine mentions without treating email addresses as routines', () => {
		const routine = (name: string, prompt = 'Read the inbox and summarize important mail.') => upcastPartial<IAssistantRoutine>({
			name, prompt, serverIds: ['gmail-server'], enabled: true, cadence: 'daily', hour: 8, minute: 30, timeZone: 'UTC', nextAt: 0,
		});
		const weekly = routine('Weekly report');
		const report = routine('Report');
		assert.deepStrictEqual(assistantRoutineMention('@Weekly report, include blockers', [report, weekly]), { routine: weekly, start: 0, end: 14 });
		assert.deepStrictEqual(assistantRoutineMention('Please use @weekly   REPORT now', [weekly]), { routine: weekly, start: 11, end: 27 });
		assert.strictEqual(assistantRoutineMention('Email person@example.com', [weekly]), undefined);
		assert.strictEqual(assistantRoutineMention('@Weekly reporting notes', [weekly]), undefined);
	});

	test('routine mention selections start only their selected connection', async () => {
		const storage = store.add(new InMemoryStorageService());
		let autostarts = 0;
		let gmailStarts = 0;
		let unrelatedStarts = 0;
		const gmail = fakeServer({ name: 'Gmail', id: 'gmail-server', running: false, tools: [], onStart: () => { gmailStarts++; } });
		const unrelated = fakeServer({ name: 'Project tracker', id: 'tracker-server', running: false, tools: [], onStart: () => { unrelatedStarts++; } });
		const mcp = new class extends mock<IMcpService>() {
			override readonly servers = observableValue<readonly IMcpServer[]>('fixtureServers', [gmail, unrelated]);
			override autostart() { autostarts++; return constObservable(IAutostartResult.Empty); }
		};
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(): Promise<ReachAssistantResponse> { return { kind: 'reply', text: 'Routine complete.' }; }
		};
		const runs = store.add(new AssistantRuns(reach, storage, mcp, upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		assert.strictEqual(await runs.send('account', '<attachment name="Routine · Weekly report">Summarize mail.</attachment>\n\nRun this routine now.', undefined, { serverIds: [assistantConnectionKey(gmail)] }), true);
		assert.deepStrictEqual([autostarts, gmailStarts, unrelatedStarts], [0, 1, 0]);
		assert.strictEqual(unrelated.connectionState.get().state, McpConnectionState.Kind.Stopped);
	});

	test('docks Browser and Terminal tabs on the Assistant view activity bar', () => {
		const descriptors: ICustomViewDescriptor[] = [];
		const views = new class extends mock<ICustomViewService>() {
			override readonly activeCustomView = constObservable<ICustomViewDescriptor | undefined>(undefined);
			override registerCustomView(value: ICustomViewDescriptor) { descriptors.push(value); return Disposable.None; }
		};
		store.add(new AssistantSessionContribution(views));
		const [descriptor] = descriptors;
		assert.deepStrictEqual(descriptor.coveredParts, [Parts.SESSIONS_PART, Parts.AUXILIARYBAR_PART]);
		assert.deepStrictEqual(descriptor.activityBar, { menuId: Menus.AssistantActivityBar, contextMenuId: Menus.AssistantActivityBarContext });
		const tabs = MenuRegistry.getMenuItems(Menus.AssistantActivityBar).filter(isIMenuItem);
		assert.deepStrictEqual(tabs.map(item => item.command.id), [TOGGLE_ASSISTANT_BROWSER_ACTION_ID, TOGGLE_ASSISTANT_TERMINAL_ACTION_ID]);
		assert.ok(tabs.every(item => CommandsRegistry.getCommand(item.command.id)));
		assert.strictEqual(tabs[0].command.icon, Codicon.globe);
		assert.strictEqual(tabs[1].command.icon, Codicon.terminal);
	});

	test('converts image attachments into upstream vision content', () => {
		const pngBytes = new Uint8Array([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 1, 2, 3, 4]);
		const image = assistantImageAttachmentDataUrl({ kind: 'image', id: 'image-1', name: 'whiteboard.png', value: pngBytes });
		assert.ok(image);
		assert.ok(image.url.startsWith('data:image/png;base64,'));
		assert.strictEqual(image.bytes, pngBytes.byteLength);
		assert.deepStrictEqual(toUpstreamMessage({ role: 'user', content: 'What does this show?', imageUrls: [image.url] }).content, [
			{ type: 'text', text: 'What does this show?' },
			{ type: 'image_url', image_url: { url: image.url } },
		]);
	});

	test('sends attached images and replays them with their follow-up history', async () => {
		const storage = store.add(new InMemoryStorageService());
		const conversation = new AssistantConversationStore(storage);
		const imageUrl = 'data:image/png;base64,iVBORw0KGgo=';
		const requests: ReachAssistantRequest[] = [];
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_account: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				requests.push({ ...request, messages: [...request.messages] });
				return { kind: 'reply', text: requests.length === 1 ? 'A whiteboard sketch.' : 'Mostly blue.' };
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, emptyMcpService(), upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		assert.strictEqual(await runs.send('account', 'What does this image show?', undefined, { imageUrls: [imageUrl] }), true);
		assert.match(requests[0].messages[0].content, /Lead with the completed outcome in first person/);
		assert.match(requests[0].messages[0].content, /waiting for the user’s approval/);
		assert.deepStrictEqual(requests[0].messages.at(-1)?.imageUrls, [imageUrl]);
		assert.deepStrictEqual(conversation.read('account').turns[0].imageUrls, [imageUrl]);
		assert.strictEqual(await runs.send('account', 'What color dominates?'), true);
		assert.strictEqual(requests[1].messages[1].role, 'user');
		assert.deepStrictEqual(requests[1].messages[1].imageUrls, [imageUrl]);
		assert.deepStrictEqual(requests[1].messages.at(-1)?.imageUrls, undefined);
		assert.strictEqual(await runs.send('account', 'What does this image show?', undefined, { replaceFrom: 0 }), true);
		assert.deepStrictEqual(requests[2].messages.at(-1)?.imageUrls, [imageUrl]);
	});

	test('retrieves requested saved ideas across chats and limits follow-up reads to matching snapshots', async () => {
		const storage = store.add(new InMemoryStorageService());
		const conversations = new AssistantConversationStore(storage);
		const taskId = conversations.createTask('account');
		const source = 'Copper shade with paper ribs. ' + 'x'.repeat(2_000) + '\nUse a folding base.';
		const saved = conversations.saveIdea('account', source, 'Lamp sketches', taskId);
		assert.ok(saved.status === 'saved');
		for (let i = 0; i < 4; i++) { conversations.saveIdea('account', 'Copper option ' + i, 'Lamp sketches', taskId); }
		for (const [id, sourceTitle] of [['detail', '[Cutouts](file:///layout.md)'], ['never-read', 'Archived source']]) {
			storage.store('sessions.assistant.savedIdeas.account/' + id, { id, text: 'Copper stand.', sourceTitle, createdAt: 1 }, StorageScope.PROFILE, StorageTarget.MACHINE);
		}
		const unrelated = conversations.saveIdea('account', 'Unrelated vintage sailboat.', 'Boats');
		assert.ok(unrelated.status === 'saved');
		conversations.saveIdea('other', 'Copper private account sentinel.', 'Other');
		conversations.appendTurn('account', { prompt: 'Previous question', reply: 'Unrelated earlier exchange.' });
		const requests: ReachAssistantRequest[] = [];
		let replyOverride: (() => Promise<ReachAssistantResponse>) | undefined;
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_account: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				requests.push({ ...request, messages: [...request.messages] });
				if (replyOverride) { return replyOverride(); }
				if (requests.length <= 2) {
					const id = requests.length === 1 ? unrelated.idea.id : saved.idea.id;
					return { kind: 'reply', text: '', toolCalls: [{ id: 'read' + requests.length, name: 'assistant.read_saved_ideas', arguments: JSON.stringify({ id, page: 2 }) }] };
				}
				if (requests.length === 3) {
					return { kind: 'reply', text: '', toolCalls: [{ id: 'detail', name: 'assistant.read_saved_ideas', arguments: JSON.stringify({ id: 'detail' }) }] };
				}
				return { kind: 'reply', text: 'Try a folding base.' };
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, emptyMcpService(), upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		assert.strictEqual(await runs.send('account', 'Using my saved ideas about copper, suggest a prototype.'), true);
		assert.deepStrictEqual(requests.map(request => request.tools?.map(tool => tool.name)), Array(4).fill(['assistant.read_saved_ideas']));
		assert.match(requests[0].messages.find(message => message.content.includes('Lamp sketches'))?.content ?? '', /Copper shade with paper ribs/);
		assert.match(requests[1].messages.at(-1)?.content ?? '', /No matching saved idea/);
		assert.match(requests[2].messages.at(-1)?.content ?? '', /Use a folding base/);
		assert.match(requests[3].messages.at(-1)?.content ?? '', /Copper stand/);
		assert.doesNotMatch(JSON.stringify(requests), /Archived source|vintage sailboat|private account sentinel|Unrelated earlier exchange/);
		assert.strictEqual(conversations.read('account').turns.at(-1)?.reply, 'Try a folding base.\n\n**Source chats consulted** · Lamp sketches · \\[Cutouts\\]\\(file\\:///layout\\.md\\)');

		await runs.send('account', 'What ideas did I save about bicycles?', taskId);
		assert.match(conversations.read('account', taskId).turns.at(-1)?.reply ?? '', /couldn’t find a saved idea/);
		assert.doesNotMatch(conversations.read('account', taskId).turns.at(-1)?.reply ?? '', /Source chats consulted/);
		assert.strictEqual(requests.length, 4);
		const attachment = '<attachment name="example">What are my saved ideas about copper?</attachment>\n\n';
		await runs.send('account', attachment + 'Explain this example.', taskId, { requestStart: attachment.length });
		assert.strictEqual(requests.length, 5);
		assert.ok(!requests.at(-1)?.tools?.some(tool => tool.name === 'assistant.read_saved_ideas'));
		assert.doesNotMatch(JSON.stringify(requests.at(-1)?.messages), /Copper shade with paper ribs/);

		const before = conversations.read('account', taskId).turns;
		const question = 'Find my saved ideas about copper.';
		for (const response of [{ kind: 'reply', text: '  ' }, { kind: 'unavailable' }] as const) {
			replyOverride = async () => response;
			assert.strictEqual(await runs.send('account', question, taskId), false);
			assert.deepStrictEqual(conversations.read('account', taskId).turns, before);
		}
		const started = new DeferredPromise<void>();
		const finish = new DeferredPromise<ReachAssistantResponse>();
		replyOverride = () => { started.complete(); return finish.p; };
		const pending = runs.send('account', question, taskId);
		await started.p;
		runs.stop('account', taskId);
		await finish.complete({ kind: 'reply', text: 'Cancelled result.' });
		assert.strictEqual(await pending, false);
		assert.deepStrictEqual(conversations.read('account', taskId).turns, before);
	});

	test('loads workflow guidance on a native tool call and respects the global skills setting on the next turn', async () => {
		const storage = store.add(new InMemoryStorageService());
		const requests: ReachAssistantRequest[] = [];
		const uri = URI.parse('file:///skills/meeting-follow-up/SKILL.md');
		let reads = 0;
		let enabled = true;
		const prompts = upcastPartial<IPromptsService>({
			findAgentSkills: async () => [{ name: 'meeting-follow-up', description: 'Recap meeting notes', uri, storage: PromptsStorage.builtIn, userInvocable: true, disableModelInvocation: false }],
			getDisabledPromptFiles: () => new ResourceSet(),
			parseNew: async uri => { reads++; return new PromptFileParser().parse(uri, 'Keep proposals under open questions.'); },
		});
		const configuration = upcastPartial<IConfigurationService>({ getValue: <T>(key?: unknown) => (key === PromptsConfig.USE_AGENT_SKILLS ? enabled : undefined) as T });
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_account: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				requests.push(request);
				return requests.length === 1
					? { kind: 'reply', text: '', toolCalls: [{ id: 'skill', name: 'assistant.find_skill', arguments: '{"name":"meeting-follow-up"}' }] }
					: { kind: 'reply', text: 'The launch date is still open.' };
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, emptyMcpService(), upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configuration, mailWatcherStub, prompts));
		await runs.send('account', 'Recap this meeting. We discussed a launch date.');
		enabled = false;
		await runs.send('account', 'Thanks.');
		assert.deepStrictEqual(requests.map(request => request.tools?.some(tool => tool.name === 'assistant.find_skill')), [true, true, false]);
		assert.match(requests[1].messages.find(message => message.role === 'tool')?.content ?? '', /Keep proposals under open questions/);
		assert.strictEqual(reads, 1);
	});

	test('reads one eligible named skill before answering a reference question and stops if there is no evidence', async () => {
		const storage = store.add(new InMemoryStorageService());
		const uri = URI.parse('file:///skills/meeting-follow-up/SKILL.md');
		const question = 'Read the installed meeting-follow-up skill. What does it say to do when a deadline is missing?';
		const content = 'Write “unspecified” for an absent deadline.';
		let hidden = false;
		let enabled = true;
		let starts = 0;
		let reads = 0;
		const reading = new DeferredPromise<void>();
		const resume = new DeferredPromise<void>();
		let pause = false;
		const prompts = upcastPartial<IPromptsService>({
			findAgentSkills: async () => [{ name: 'meeting-follow-up', description: '', uri, storage: PromptsStorage.builtIn, userInvocable: false, disableModelInvocation: hidden }],
			getDisabledPromptFiles: () => new ResourceSet(),
			parseNew: async uri => {
				reads++;
				if (pause) { await reading.complete(); await resume.p; }
				return new PromptFileParser().parse(uri, content);
			},
		});
		const mcp = new class extends mock<IMcpService>() {
			override readonly servers = constObservable([]);
			override autostart() { starts++; return constObservable(IAutostartResult.Empty); }
		};
		const configuration = upcastPartial<IConfigurationService>({ getValue: <T>(key?: unknown) => (key === PromptsConfig.USE_AGENT_SKILLS ? enabled : undefined) as T });
		const requests: ReachAssistantRequest[] = [];
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_account: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				requests.push(request);
				return { kind: 'reply', text: 'It says to write “unspecified”.' };
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, mcp, upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configuration, mailWatcherStub, prompts));
		const conversation = new AssistantConversationStore(storage);
		conversation.appendTurn('account', { prompt: 'Unrelated history', reply: 'Old context' });
		await runs.send('account', question);
		assert.deepStrictEqual(requests[0].tools ?? [], []);
		assert.match(requests[0].messages[0].content, /reference material/);
		assert.deepStrictEqual(requests[0].messages.at(-1), { role: 'user', content: question });
		assert.ok(requests[0].messages.some(message => message.role === 'user' && message.content.includes(content)));
		assert.ok(!requests[0].messages.some(message => message.content.includes('Unrelated history')));
		assert.deepStrictEqual(conversation.read('account').turns.at(-1)?.activity?.map(a => [a.tool, a.status]), [['assistant.find_skill', 'done']]);
		hidden = true;
		await runs.send('account', question);
		assert.match(conversation.read('account').turns.at(-1)!.reply, /couldn’t read the instructions/);
		hidden = false;
		enabled = false;
		await runs.send('account', question);
		assert.strictEqual(reads, 1);
		enabled = true;
		await runs.send('account', question.replace('meeting-follow-up', 'unknown-skill'));
		pause = true;
		const pending = runs.send('account', question);
		await reading.p;
		runs.stop('account');
		await resume.complete();
		assert.strictEqual(await pending, false);
		assert.strictEqual(conversation.read('account').turns.length, 5);
		assert.strictEqual(requests.length, 1);
		assert.strictEqual(starts, 0);
	});

	test('offers only same-skill page reads to a text-only model for a multi-page reference', async () => {
		const storage = store.add(new InMemoryStorageService());
		const uri = URI.parse('file:///skills/meeting-follow-up/SKILL.md');
		const content = 'x'.repeat(4000) + 'Retain the relative deadline.';
		const reads: string[] = [];
		const prompts = upcastPartial<IPromptsService>({
			findAgentSkills: async () => ['meeting-follow-up', 'other'].map(name => ({ name, description: '', uri: URI.joinPath(uri, name), storage: PromptsStorage.builtIn, userInvocable: true, disableModelInvocation: false })),
			getDisabledPromptFiles: () => new ResourceSet(),
			parseNew: async uri => { reads.push(uri.path); return new PromptFileParser().parse(uri, content); },
		});
		const config = upcastPartial<IConfigurationService>({ getValue: <T>(key?: unknown) => (key === PromptsConfig.USE_AGENT_SKILLS ? true : undefined) as T });
		const offers: string[][] = [];
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_account: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				offers.push(request.tools?.map(tool => tool.name) ?? []);
				assert.match(request.textToolIndex ?? '', /assistant\.find_skill/);
				const result = request.messages.at(-1)!.content;
				assert.ok(result.length <= REACH_ASSISTANT_PROMPT_LIMIT);
				const call = (args: object) => ({ kind: 'reply' as const, text: '\`\`\`tool_call\n' + JSON.stringify({ name: 'assistant.find_skill', arguments: args }) + '\n\`\`\`' });
				switch (offers.length) {
					case 1:
						assert.strictEqual(result.length, REACH_ASSISTANT_PROMPT_LIMIT);
						assert.match(request.messages.at(-2)!.content, /page 1 of 2/);
						return call({ name: 'other' });
					case 2: assert.match(result, /Only pages of the requested skill/); return call({ page: 2 });
					default: assert.match(result, /Retain the relative deadline/); return { kind: 'reply', text: 'Keep the relative deadline.' };
				}
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, emptyMcpService(), upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, config, mailWatcherStub, prompts));
		const question = 'What does the installed meeting-follow-up skill say about deadlines'.padEnd(REACH_ASSISTANT_PROMPT_LIMIT - 1, ' ') + '?';
		await runs.send('account', question);
		assert.deepStrictEqual(offers, Array.from({ length: 3 }, () => ['assistant.find_skill']));
		assert.deepStrictEqual(reads, [URI.joinPath(uri, 'meeting-follow-up').path, URI.joinPath(uri, 'meeting-follow-up').path]);
	});

	test('writes from supplied notes with a small tool catalog and skips connection startup', async () => {
		const storage = store.add(new InMemoryStorageService());
		const conversation = new AssistantConversationStore(storage);
		const bank = { user: '- Replies in French', memory: 'Old environment note' };
		conversation.writeMemoryBank('account', bank);
		conversation.recordMailAlerts('account', [{ id: 'fixture', connectionId: 'mail', label: 'Mail', subject: 'Unrelated subject', at: 1 }]);
		let startups = 0;
		const mcp = new class extends mock<IMcpService>() {
			override readonly servers = constObservable([]);
			override autostart() { startups++; return constObservable(IAutostartResult.Empty); }
		};
		const offers: { tools: string[]; system: string; textIndex?: string }[] = [];
		let latestMessages: ReachAssistantRequest['messages'] = [];
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_account: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				offers.push({ tools: request.tools?.map(tool => tool.name) ?? [], system: request.messages[0].content, textIndex: request.textToolIndex });
				latestMessages = request.messages;
				return { kind: 'reply', text: 'Mira will send the draft Friday.' };
			}
		};
		const configuration = upcastPartial<IConfigurationService>({ getValue: <T>(key?: unknown) => (key === PromptsConfig.USE_AGENT_SKILLS ? true : undefined) as T });
		const runs = store.add(new AssistantRuns(reach, storage, mcp, upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configuration, mailWatcherStub, promptsStub));
		const notes = 'Summarize these notes: Mira will send the draft Friday.';
		await runs.send('account', '<attachment name="Skill · meeting-follow-up">\nCapture agreed actions.\n</attachment>\n\n' + notes);
		await runs.send('account', 'Summarize these notes: Mira will send the email about the repository Friday.');
		await runs.send('account', '<file path="minutes.txt">\nA decision was deferred.\n</file>\n\nRecap the attached transcript.');
		await runs.send('account', 'For a meeting recap, summarize these notes for the team. Rowan will send the budget Monday. Keep the answer to one sentence.');
		await runs.send('account', 'Turn these meeting notes into a follow-up for the attendees. Mira will send the draft Friday.');
		assert.deepStrictEqual(offers.map(offer => offer.tools), [
			['assistant.find_app_tools'], ...Array.from({ length: 4 }, () => ['assistant.find_app_tools', 'assistant.find_skill']),
		]);
		for (const { system, textIndex } of offers) {
			assert.match(system, /material the user supplied/);
			assert.match(system, /Replies in French/);
			assert.doesNotMatch(system, /news|weather|direct mail|workspace folder|Old environment note|Unrelated subject|```tool_call/i);
			assert.match(textIndex ?? '', /```tool_call\n\{"name":"assistant\.find_app_tools"/);
		}
		assert.strictEqual(startups, 0);
		await runs.send('account', 'Summarize these notes from my calendar.');
		await runs.send('account', 'Summarize these notes. Find them in my calendar.');
		await runs.send('account', 'Summarize these notes. Email the recap to the team.');
		await runs.send('account', 'Summarize these notes: Mira will send the email about the repository. Then check my calendar.');
		assert.ok(offers.slice(5).every(offer => offer.tools.includes('assistant.list_connection_sources')));
		assert.strictEqual(startups, 0);
		assert.strictEqual(conversation.read('account').turns.at(-1)?.reply, 'Mira will send the draft Friday.');
		assert.deepStrictEqual(conversation.readMemoryBank('account'), bank);
		const idea = '<attachment name="Skill · idea-exploration">\nSuggest an experiment.\n</attachment>\n\n';
		const request = 'I have an hour with my grandmother. How could we start a recipe collection?';
		assert.strictEqual(await runs.send('account', idea + request, undefined, { requestStart: idea.length }), true);
		assert.deepStrictEqual(latestMessages.slice(-2), [{ role: 'user', content: idea.trimEnd() }, { role: 'user', content: request }]);
		assert.strictEqual(conversation.read('account').turns.at(-1)?.prompt, idea + request);
		assert.deepStrictEqual(offers.at(-1)?.tools, ['assistant.find_app_tools']);
		assert.match(offers.at(-1)?.system ?? '', /Use the workflow attached/);
		assert.strictEqual(startups, 0);
	});

	test('activates only a matching cold connection on explicit tool discovery during a supplied-text turn', async () => {
		const storage = store.add(new InMemoryStorageService());
		let startups = 0;
		let autostarts = 0;
		let reads = 0;
		const readNote = upcastPartial<IMcpTool>({
			id: 'read_note', visibility: McpToolVisibility.Model,
			definition: { name: 'read_note', inputSchema: { type: 'object' }, annotations: { readOnlyHint: true } },
			call: async () => { reads++; return { content: [{ type: 'text', text: 'Fixture reference.' }] }; },
		});
		const notes = fakeServer({ name: 'Notes', running: false, tools: [], onStart: setTools => { startups++; setTools([readNote]); } });
		const servers = observableValue<readonly IMcpServer[]>('fixtureServers', [notes]);
		const mcp = new class extends mock<IMcpService>() {
			override readonly servers = servers;
			override autostart() {
				autostarts++;
				return constObservable(IAutostartResult.Empty);
			}
		};
		let step = 0;
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_account: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				step++;
				if (step === 1) {
					assert.strictEqual(startups, 0);
					return { kind: 'reply', text: '', toolCalls: [{ id: 'memory', name: 'assistant.find_app_tools', arguments: '{"name":"assistant.memory"}' }] };
				}
				if (step === 2) {
					assert.strictEqual(startups, 0);
					assert.match(request.messages.at(-1)?.content ?? '', /preferences are read-only/);
					assert.ok(!request.tools?.some(tool => tool.name === 'assistant.memory'));
					return { kind: 'reply', text: '', toolCalls: [{ id: 'builtin', name: 'assistant.find_app_tools', arguments: '{"name":"assistant.get_current_time"}' }] };
				}
				if (step === 3) {
					assert.ok(request.tools?.some(tool => tool.name === 'assistant.get_current_time'));
					assert.strictEqual(startups, 0);
					return { kind: 'reply', text: '', toolCalls: [{ id: 'app', name: 'assistant.find_app_tools', arguments: '{"app":"Notes","name":"read_note"}' }] };
				}
				if (step === 4) {
					assert.ok(request.tools?.some(tool => tool.name === 'read_note'));
					return { kind: 'reply', text: '', toolCalls: [{ id: 'read', name: 'read_note', arguments: '{}' }] };
				}
				return { kind: 'reply', text: 'Done.' };
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, mcp, upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		assert.ok(await runs.send('account', 'Proofread this text: The draft is ready.'));
		assert.deepStrictEqual([autostarts, startups, reads, step], [0, 1, 1, 5]);
	});

	test('offers cached connection tools without starting every configured app', async () => {
		const storage = store.add(new InMemoryStorageService());
		const gmail = fakeServer({
			name: 'Gmail', id: 'gmail-server', running: false, tools: [
				upcastPartial<IMcpTool>({
					id: 'gmail_search', visibility: McpToolVisibility.Model,
					definition: { name: 'gmail_search', inputSchema: { type: 'object' }, annotations: { readOnlyHint: true } },
					call: async () => ({ content: [{ type: 'text', text: 'Fixture message ID' }] }),
				}),
			]
		});
		const unrelated = fakeServer({ name: 'Project tracker', id: 'tracker-server', running: false, tools: [] });
		let autostarts = 0;
		const mcp = new class extends mock<IMcpService>() {
			override readonly servers = observableValue<readonly IMcpServer[]>('fixtureServers', [gmail, unrelated]);
			override autostart() { autostarts++; return constObservable(IAutostartResult.Empty); }
		};
		const requests: ReachAssistantRequest[] = [];
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_account: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				requests.push(request);
				return requests.length === 1
					? { kind: 'reply', text: '', toolCalls: [{ id: 'search', name: 'gmail_search', arguments: '{}' }] }
					: { kind: 'reply', text: 'Found the test message.' };
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, mcp, upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		assert.strictEqual(await runs.send('account', 'Search my Gmail.'), true);
		assert.strictEqual(autostarts, 0);
		assert.strictEqual(gmail.connectionState.get().state, McpConnectionState.Kind.Running);
		assert.strictEqual(unrelated.connectionState.get().state, McpConnectionState.Kind.Stopped);
		assert.ok(requests[0].tools?.some(tool => tool.name === 'gmail_search'));
		assert.strictEqual(requests[1].messages.at(-1)?.content, 'Fixture message ID');
	});

	test('activates a named cold connection before the first model request', async () => {
		const storage = store.add(new InMemoryStorageService());
		let autostarts = 0;
		let gmailStarts = 0;
		let unrelatedStarts = 0;
		const gmailSearch = upcastPartial<IMcpTool>({
			id: 'gmail_search', visibility: McpToolVisibility.Model,
			definition: { name: 'gmail_search', inputSchema: { type: 'object' }, annotations: { readOnlyHint: true } },
			call: async () => ({ content: [{ type: 'text', text: 'Fixture message ID' }] }),
		});
		const gmail = fakeServer({ name: 'Gmail', id: 'gmail-server', running: false, tools: [], onStart: setTools => { gmailStarts++; setTools([gmailSearch]); } });
		const unrelated = fakeServer({ name: 'Project tracker', id: 'tracker-server', running: false, tools: [], onStart: () => { unrelatedStarts++; } });
		const mcp = new class extends mock<IMcpService>() {
			override readonly servers = observableValue<readonly IMcpServer[]>('fixtureServers', [gmail, unrelated]);
			override autostart() { autostarts++; return constObservable(IAutostartResult.Empty); }
		};
		let requests = 0;
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_account: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				requests++;
				if (requests === 1) {
					assert.ok(request.tools?.some(tool => tool.name === 'gmail_search'));
					return { kind: 'reply', text: '', toolCalls: [{ id: 'search', name: 'gmail_search', arguments: '{}' }] };
				}
				assert.strictEqual(request.messages.at(-1)?.content, 'Fixture message ID');
				return { kind: 'reply', text: 'Found the test message.' };
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, mcp, upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		assert.strictEqual(await runs.send('account', 'Search my Gmail.'), true);
		assert.deepStrictEqual([autostarts, gmailStarts, unrelatedStarts, requests], [0, 1, 0, 2]);
		assert.strictEqual(unrelated.connectionState.get().state, McpConnectionState.Kind.Stopped);
	});

	test('can stop Assistant before the first model request', async () => {
		const storage = store.add(new InMemoryStorageService());
		const mcp = emptyMcpService();
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(): Promise<ReachAssistantResponse> { assert.fail('A stopped run must not reach the model'); }
		};
		const runs = store.add(new AssistantRuns(reach, storage, mcp, upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		const request = runs.send('account', 'Search my Gmail.');
		runs.stop('account');
		assert.strictEqual(await request, false);
	});

	test('activates only Gmail for cold discovery and executes its tool only once offered', async () => {
		const storage = store.add(new InMemoryStorageService());
		let toolCalls = 0;
		let gmailStarts = 0;
		let unrelatedStarts = 0;
		let autostarts = 0;
		const gmailSearch = upcastPartial<IMcpTool>({
			id: 'gmail_search', visibility: McpToolVisibility.Model,
			definition: { name: 'gmail_search', inputSchema: { type: 'object' }, annotations: { readOnlyHint: true } },
			call: async () => {
				toolCalls++;
				return { content: [{ type: 'text', text: 'Late fixture message ID' }] };
			},
		});
		const gmail = fakeServer({ name: 'Gmail', id: 'gmail-server', running: false, tools: [], onStart: setTools => { gmailStarts++; setTools([gmailSearch]); } });
		const unrelated = fakeServer({ name: 'Project tracker', id: 'tracker-server', running: false, tools: [], onStart: () => { unrelatedStarts++; } });
		const mcp = new class extends mock<IMcpService>() {
			override readonly servers = observableValue<readonly IMcpServer[]>('fixtureServers', [gmail, unrelated]);
			override autostart() { autostarts++; return constObservable(IAutostartResult.Empty); }
		};
		let step = 0;
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_account: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				if (++step === 1) {
					assert.ok(!request.tools?.some(tool => tool.name === 'gmail_search'));
					return {
						kind: 'reply', text: '', toolCalls: [
							{ id: 'find', name: 'assistant.find_app_tools', arguments: '{"app":"Gmail"}' },
							{ id: 'too-soon', name: 'gmail_search', arguments: '{}' },
						]
					};
				}
				assert.ok(request.tools?.some(tool => tool.name === 'gmail_search'));
				if (step === 2) {
					assert.match(request.messages.find(message => message.toolCallId === 'find')?.content ?? '', /Activated on demand: Gmail/);
					assert.match(request.messages.find(message => message.toolCallId === 'find')?.content ?? '', /gmail_search/);
					assert.match(request.messages.at(-1)?.content ?? '', /unknown tool "gmail_search"/);
					assert.strictEqual(toolCalls, 0);
					return { kind: 'reply', text: '', toolCalls: [{ id: 'search', name: 'gmail_search', arguments: '{}' }] };
				}
				assert.strictEqual(request.messages.at(-1)?.content, 'Late fixture message ID');
				return { kind: 'reply', text: 'Found the message.' };
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, mcp, upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		assert.strictEqual(await runs.send('account', 'Search my inbox.'), true);
		assert.deepStrictEqual([autostarts, gmailStarts, unrelatedStarts, toolCalls, step], [0, 1, 0, 1, 3]);
		assert.strictEqual(gmail.connectionState.get().state, McpConnectionState.Kind.Running);
		assert.strictEqual(unrelated.connectionState.get().state, McpConnectionState.Kind.Stopped);
	});

	test('includes a late read-only account when listing sources for a briefing', async () => {
		await runWithFakedTimers({ useFakeTimers: true }, async () => {
			const storage = store.add(new InMemoryStorageService());
			const readTool = (name: string) => upcastPartial<IMcpTool>({
				id: name, visibility: McpToolVisibility.Model,
				definition: { name, inputSchema: { type: 'object' }, annotations: { readOnlyHint: true } },
			});
			const github = fakeServer({ name: 'GitHub', tools: [readTool('read_prs')] });
			const gmail = fakeServer({
				name: 'Gmail', tools: [readTool('read_email'), upcastPartial<IMcpTool>({
					id: 'send_email', visibility: McpToolVisibility.Model,
					definition: { name: 'send_email', inputSchema: { type: 'object' } },
				})]
			});
			const startup = observableValue<IAutostartResult>('fixtureStartup', { ...IAutostartResult.Empty, working: true });
			const mcp = new class extends mock<IMcpService>() {
				override readonly servers = observableValue<readonly IMcpServer[]>('fixtureServers', [github]);
				override autostart() { return startup; }
			};
			let requests = 0;
			const reach = new class extends mock<IReachSubscriptionConnection>() {
				override readonly onDidChangeAccount = Event.None;
				override async sendAssistantPrompt(_account: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
					const names = request.tools?.map(tool => tool.name) ?? [];
					if (++requests === 1) {
						assert.ok(names.includes('read_prs') && !names.includes('read_email'));
						void timeout(50).then(() => {
							mcp.servers.set([github, gmail], undefined);
							startup.set(IAutostartResult.Empty, undefined);
						});
						return { kind: 'reply', text: '', toolCalls: [{ id: 'sources', name: 'assistant.list_connection_sources', arguments: '{}' }] };
					}
					assert.deepStrictEqual(JSON.parse(request.messages.at(-1)?.content ?? ''), { available: ['GitHub', 'Gmail'], inactive: [], unavailable: [] });
					assert.ok(names.includes('read_prs') && names.includes('read_email') && !names.includes('send_email'));
					return { kind: 'reply', text: 'Briefing with both fixture sources.' };
				}
			};
			const runs = store.add(new AssistantRuns(reach, storage, mcp, upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
			assert.strictEqual(await runs.send('account', '/briefing morning'), true);
			assert.strictEqual(requests, 2);
		});
	});

	test('briefs on demand with connected read-only tools and the same prompts as the daily presets', async () => {
		const storage = store.add(new InMemoryStorageService());
		const requests: ReachAssistantRequest[] = [];
		const conversations = new AssistantConversationStore(storage);
		conversations.appendTurn('account', { prompt: 'Unrelated private context', reply: 'Old reply' });
		const mcpService = mcpServiceWith(fakeServer({
			name: 'GitHub', instructions: 'Untrusted fixture instructions', tools: [
				upcastPartial<IMcpTool>({
					id: 'read_prs', visibility: McpToolVisibility.Model,
					definition: { name: 'read_prs', inputSchema: { type: 'object' }, annotations: { readOnlyHint: true } },
				}),
				upcastPartial<IMcpTool>({
					id: 'write_pr', visibility: McpToolVisibility.Model,
					definition: { name: 'write_pr', inputSchema: { type: 'object' } },
				}),
			]
		}));
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_account: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				requests.push(request);
				return { kind: 'reply', text: 'Briefing with fixture sources.' };
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, mcpService, upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		for (const period of ['morning', 'evening'] as const) {
			assert.strictEqual(await runs.send('account', `/briefing ${period}`), true);
			const request = requests.at(-1)!;
			assert.strictEqual(request.messages.length, 2);
			assert.doesNotMatch(request.messages[0].content, /Untrusted fixture instructions/);
			assert.strictEqual(request.messages[1].content, ASSISTANT_BRIEFINGS[period].prompt);
			assert.match(request.messages[1].content, /received and sent email separately/);
			assert.deepStrictEqual(request.tools?.map(tool => tool.name).sort(),
				['assistant.find_app_tools', 'assistant.get_current_time', 'assistant.get_news_headlines', 'assistant.get_weather_forecast', 'assistant.list_connection_sources', 'read_prs']);
		}
		assert.deepStrictEqual(conversations.read('account').turns.slice(1).map(turn => turn.prompt), ['/briefing morning', '/briefing evening']);
		assert.deepStrictEqual([ASSISTANT_BRIEFINGS.morning.schedule.scheduleHour, ASSISTANT_BRIEFINGS.evening.schedule.scheduleHour], [8, 18]);
	});

	test('runs scheduled briefings through the same Assistant with quiet read-only tools and preserves the draft', async () => {
		const storage = store.add(new InMemoryStorageService());
		const conversation = new AssistantConversationStore(storage);
		conversation.setDraft('account', '/briefing morning');
		let called = false;
		let newsScope: AssistantNewsScope | undefined;
		let weatherDay: AssistantWeatherDay | undefined;
		let weatherLocation: AssistantWeatherLocation | undefined;
		const forecast: AssistantWeatherForecast = {
			source: 'The Weather Network', sourceUrl: 'https://www.theweathernetwork.com/en/city/ca/british-columbia/vancouver/7-days',
			location: 'Vancouver, BC, CA', date: '2026-11-19', highC: 8, lowC: 2, precipitationChancePercent: 70, conditions: 'Showers',
		};
		const published = new Date(Date.now() - 60_000);
		const stale = new Date(Date.now() - 3 * 24 * 60 * 60 * 1000);
		const mcp = mcpServiceWith(fakeServer({
			tools: [
				upcastPartial<IMcpTool>({
					id: 'read_mail', visibility: McpToolVisibility.Model,
					definition: { name: 'read_mail', inputSchema: { type: 'object' }, annotations: { readOnlyHint: true } },
					call: async (_args, context) => {
						assert.strictEqual(context?.errorOnUserInteraction, true);
						called = true;
						return { content: [{ type: 'text', text: 'Fixture email summary' }] };
					},
				}),
				upcastPartial<IMcpTool>({
					id: 'send_mail', visibility: McpToolVisibility.Model,
					definition: { name: 'send_mail', inputSchema: { type: 'object' } },
				}),
			]
		}));
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async readAssistantWeatherForecast(day: AssistantWeatherDay, location: AssistantWeatherLocation) {
				weatherDay = day;
				weatherLocation = location;
				return forecast;
			}
			override async readAssistantNewsFeeds(scope: AssistantNewsScope) {
				newsScope = scope;
				return [
					{ publisher: 'CBC News', error: 'HTTP 503' },
					{
						publisher: 'Global News', xml: `<rss><channel>
						<item><title>Toronto &amp; area</title><link>https://globalnews.ca/news/current</link><pubDate>${published.toUTCString()}</pubDate></item>
						<item><title>Duplicate</title><link>https://globalnews.ca/news/current#section</link><pubDate>${published.toUTCString()}</pubDate></item>
						<item><title>Old Toronto news</title><link>https://globalnews.ca/news/old</link><pubDate>${stale.toUTCString()}</pubDate></item>
						<item><title>Invalid link</title><link>javascript:alert(1)</link><pubDate>${published.toUTCString()}</pubDate></item>
					</channel></rss>` },
				];
			}
			override async sendAssistantPrompt(_account: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				assert.ok(request.tools?.some(tool => tool.name === 'read_mail'));
				assert.ok(!request.tools?.some(tool => tool.name === 'send_mail' || tool.name === 'assistant.web_fetch' || tool.name === 'assistant.set_weather_location'));
				if (request.messages.at(-1)?.role === 'tool') {
					const results = request.messages.filter(message => message.role === 'tool');
					assert.strictEqual(results[0].content, 'Fixture email summary');
					assert.deepStrictEqual(JSON.parse(results[1].content), {
						scope: 'toronto',
						headlines: [{ title: 'Toronto & area', url: 'https://globalnews.ca/news/current', publisher: 'Global News', publishedAt: new Date(published.toUTCString()).toISOString() }],
						sources: [{ publisher: 'CBC News', status: 'HTTP 503' }, { publisher: 'Global News', status: 'Available' }],
					});
					assert.deepStrictEqual(JSON.parse(results[2].content), forecast);
					return { kind: 'reply', text: 'Your morning briefing' };
				}
				return {
					kind: 'reply', text: '', toolCalls: [
						{ id: 'read', name: 'read_mail', arguments: '{}' },
						{ id: 'news', name: 'assistant.get_news_headlines', arguments: '{"scope":"toronto"}' },
						{ id: 'weather', name: 'assistant.get_weather_forecast', arguments: '{"day":"tomorrow"}' },
					]
				};
			}
		};
		const configuration = upcastPartial<IConfigurationService>({ getValue: <T>(key?: unknown) => (key === ASSISTANT_WEATHER_LOCATION_SETTING ? 'Vancouver, BC, CA' : undefined) as T });
		const runs = store.add(new AssistantRuns(reach, storage, mcp, upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configuration, mailWatcherStub, promptsStub));
		assert.strictEqual(await runs.sendScheduledBriefing('account', 'morning'), 'completed');
		assert.strictEqual(called, true);
		assert.strictEqual(newsScope, 'toronto');
		assert.strictEqual(weatherDay, 'tomorrow');
		assert.deepStrictEqual(weatherLocation, { city: 'Vancouver', country: 'CA', subdivision: 'BC' });
		assert.strictEqual(conversation.read('account').draft, '/briefing morning');
		assert.strictEqual(conversation.read('account').turns.at(-1)?.reply, 'Your morning briefing');
	});

	test('retries a scheduled briefing only for transient first-round failures', async () => {
		const storage = store.add(new InMemoryStorageService());
		const conversation = new AssistantConversationStore(storage);
		const responses: ReachAssistantResponse[] = [];
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(): Promise<ReachAssistantResponse> {
				return responses.shift() ?? { kind: 'reply', text: 'Briefing done.' };
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, emptyMcpService(), upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		responses.push({ kind: 'unavailable', reason: 'preflightUnavailable' });
		assert.strictEqual(await runs.sendScheduledBriefing('account', 'morning'), 'retryable');
		assert.strictEqual(conversation.read('account').turns.length, 0);
		responses.push(
			{ kind: 'reply', text: '', toolCalls: [{ id: 'clock', name: 'assistant.get_current_time', arguments: '{}' }] },
			{ kind: 'unavailable', reason: 'preflightUnavailable' });
		assert.strictEqual(await runs.sendScheduledBriefing('account', 'morning'), 'failed');
		assert.strictEqual(conversation.read('account').turns.length, 0);
	});

	test('uses AI feeds for scheduled AI news and withholds an unverified Outlook inbox result', async () => {
		const storage = store.add(new InMemoryStorageService());
		const scopes: AssistantNewsScope[] = [];
		const reads: Record<string, unknown>[] = [];
		const connection = { id: 'outlook-fixture', provider: 'outlook' as const, label: 'Outlook', needsSignIn: false };
		const directMail = upcastPartial<ISignalCDirectMail>({
			connections: constObservable([connection]),
			read: async (_selected, args) => {
				reads.push(args);
				return '{"messages":[]}';
			},
		});
		const hour = 60 * 60 * 1000;
		const story = (title: string, url: string, hoursAgo: number) =>
			`<item><title>${title}</title><link>${url}</link><pubDate>${new Date(Date.now() - hoursAgo * hour).toUTCString()}</pubDate></item>`;
		const feed = (publisher: string, ...items: string[]) => ({ publisher, xml: `<rss><channel>${items.join('')}</channel></rss>` });
		let aiOutage = false;
		let useOutlook = false;
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async readAssistantNewsFeeds(scope: AssistantNewsScope) {
				scopes.push(scope);
				if (scope === 'canada') { return [feed('CBC News', story('Fixture Canadian headline', 'https://www.cbc.ca/news/canada/fixture', 1))]; }
				assert.strictEqual(scope, 'ai');
				return aiOutage ? [{ publisher: 'OpenAI', error: 'HTTP 503' }] : [
					feed('OpenAI', story('OpenAI model release', 'https://openai.com/index/fixture-model/', 1),
						story('OpenAI second announcement', 'https://openai.com/index/fixture-second/', 2),
						story('OpenAI third announcement', 'https://openai.com/index/fixture-third/', 3)),
					feed('Google DeepMind', story('Introducing Gemini Fixture', 'https://deepmind.google/blog/fixture-gemini/', 50),
						story('Expired AI story', 'https://deepmind.google/blog/expired/', 8 * 24)),
					feed('Hugging Face', story('New open-weight model', 'https://huggingface.co/blog/fixture-release', 80)),
				];
			}
			override async sendAssistantPrompt(_account: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				assert.match(request.messages[1].content, /Signal-C will add the AI news section/);
				const results = request.messages.filter(message => message.role === 'tool');
				if (!results.length) {
					const toolCalls = [{ id: 'news', name: 'assistant.get_news_headlines', arguments: '{"scope":"canada"}' }];
					if (useOutlook) { toolCalls.push({ id: 'inbox', name: 'assistant.outlook_read', arguments: '{"folder":"inbox","unreadOnly":true}' }); }
					return { kind: 'reply', text: '', toolCalls };
				}
				assert.match(results[0].content, /Fixture Canadian headline/);
				return { kind: 'reply', text: '## Weather\n\nSunny.\n\n**2) Canadian headlines**\n\nFixture Canadian headline.\n\n**3) Top 3 AI News Stories of the Day**\n\n1. **Unrelated crime story** — CBC News\n\n**4) Outlook Inbox Check**\n\n- No unread messages found in the inbox.' };
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, emptyMcpService(), upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMail, configurationStub, mailWatcherStub, promptsStub));
		const routine = { id: 'ai-briefing', name: 'Afternoon Briefing', prompt: 'Cover Toronto weather, Canadian headlines, top 3 A.I. news stories, and check my Outlook inbox.', serverIds: [connection.id] };
		const conversation = new AssistantConversationStore(storage);

		assert.strictEqual(await runs.sendScheduledRoutine('account', routine, 'ai-news-report'), 'needsAttention');
		assert.deepStrictEqual(scopes, ['ai', 'canada']);
		const turn = conversation.read('account').turns.at(-1)!;
		assert.match(turn.reply, /Fixture Canadian headline/);
		assert.match(turn.reply, /Outlook inbox could not be verified/);
		assert.match(turn.reply, /## AI news[\s\S]*OpenAI model release[\s\S]*Introducing Gemini Fixture[\s\S]*New open-weight model/);
		assert.doesNotMatch(turn.reply, /Unrelated crime story|No unread messages|Expired AI story|second announcement/);
		assert.strictEqual(turn.activity?.filter(item => item.tool === 'assistant.get_news_headlines').length, 2);

		aiOutage = true;
		useOutlook = true;
		assert.strictEqual(await runs.sendScheduledRoutine('account', routine, 'ai-news-outage'), 'needsAttention');
		assert.deepStrictEqual(reads, [{ folder: 'inbox', unreadOnly: true }]);
		const outage = conversation.read('account').turns.at(-1)!.reply;
		assert.match(outage, /No current ai headlines/);
		assert.match(outage, /No unread messages found in the inbox/);
		assert.doesNotMatch(outage, /Unrelated crime story|could not be verified/);
	});

	test('reports the local clock and UTC offset for the device or a requested time zone', async () => {
		const catalog = collectAssistantTools({
			accountId: 'account', store: new AssistantConversationStore(store.add(new InMemoryStorageService())), mcpService: emptyMcpService(),
			requestService: upcastPartial<IRequestService>({}), fileService: fileServiceMock(), workspaceService: workspaceServiceMock(), searchService: searchServiceMock(), openerService: openerStub,
		});
		const clock = catalog.entries.get('assistant.get_current_time');
		assert.ok(clock?.kind === 'builtin');
		const device = JSON.parse(await clock.run({}, CancellationToken.None));
		assert.strictEqual(device.timeZone, Intl.DateTimeFormat().resolvedOptions().timeZone);
		const nepal = JSON.parse(await clock.run({ timeZone: 'Asia/Kathmandu' }, CancellationToken.None));
		assert.match(nepal.local, /GMT\+05:45/);
		assert.strictEqual(new Date(nepal.utc).toISOString(), nepal.utc);
		await assert.rejects(clock.run({ timeZone: 'Not/AZone' }, CancellationToken.None), /Choose a time zone/);
	});

	test('saves an approved weather location in User Settings and supports one-off forecast locations', async () => {
		const storage = store.add(new InMemoryStorageService());
		let setting = 'Toronto, ON, CA';
		const readLocations: AssistantWeatherLocation[] = [];
		const configuration = upcastPartial<IConfigurationService>({
			getValue: <T>(key?: unknown) => (key === ASSISTANT_WEATHER_LOCATION_SETTING ? setting : undefined) as T,
			updateValue: async (key: string, value: unknown, target?: unknown) => {
				assert.strictEqual(key, ASSISTANT_WEATHER_LOCATION_SETTING);
				assert.strictEqual(target, ConfigurationTarget.USER);
				assert.ok(typeof value === 'string');
				setting = value;
			},
		});
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async readAssistantWeatherForecast(_day: AssistantWeatherDay, location: AssistantWeatherLocation): Promise<AssistantWeatherForecast> {
				readLocations.push(location);
				return { source: 'The Weather Network', sourceUrl: 'https://www.theweathernetwork.com/en/city', location: location.city, date: '2026-11-19' };
			}
			override async sendAssistantPrompt(_account: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				const last = request.messages.at(-1);
				if (last?.role === 'tool') { return { kind: 'reply', text: last.content }; }
				const call = last?.content.includes('Set')
					? { name: 'assistant.set_weather_location', arguments: '{"city":"Vancouver","country":"ca","subdivision":"bc"}' }
					: { name: 'assistant.get_weather_forecast', arguments: last?.content.includes('Paris') ? '{"city":"Paris","country":"FR"}' : '{}' };
				assert.ok(request.tools?.some(tool => tool.name === call.name));
				return { kind: 'reply', text: '', toolCalls: [{ id: 'weather', ...call }] };
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, emptyMcpService(), upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configuration, mailWatcherStub, promptsStub));
		let allow = false;
		store.add(runs.onDidRequestApproval(request => {
			assert.strictEqual(request.tool, 'assistant.set_weather_location');
			runs.respondApproval(request.id, allow ? 'allow' : 'reject');
		}));
		assert.strictEqual(await runs.send('account', 'Set my default weather location to Vancouver'), true);
		assert.strictEqual(setting, 'Toronto, ON, CA');
		allow = true;
		assert.strictEqual(await runs.send('account', 'Set my default weather location to Vancouver'), true);
		assert.strictEqual(setting, 'Vancouver, BC, CA');
		assert.strictEqual(await runs.send('account', 'What is my weather?'), true);
		assert.strictEqual(await runs.send('account', 'What is the weather in Paris?'), true);
		assert.deepStrictEqual(readLocations, [
			{ city: 'Vancouver', country: 'CA', subdivision: 'BC' },
			{ city: 'Paris', country: 'FR', subdivision: undefined },
		]);
		assert.strictEqual(setting, 'Vancouver, BC, CA');
	});

	test('uses only the selected connection and its read-only tools for a routine even in Allow all mode', async () => {
		const storage = store.add(new InMemoryStorageService());
		const conversation = new AssistantConversationStore(storage);
		const flushedRunIds: (string | undefined)[] = [];
		store.add(storage.onWillSaveState(() => flushedRunIds.push(conversation.read('account').turns.at(-1)?.routineRunId)));
		conversation.setDraft('account', 'Unsent draft');
		conversation.remember('account', ['Private memory from another connection']);
		conversation.setPermissionMode('account', 'allowAll');
		const called: string[] = [];
		const makeTool = (name: string, readOnlyHint: boolean) => upcastPartial<IMcpTool>({
			id: name, visibility: McpToolVisibility.Model,
			definition: { name, inputSchema: { type: 'object' }, annotations: { readOnlyHint } },
			call: async (_args, context) => {
				assert.strictEqual(context?.errorOnUserInteraction, true);
				called.push(name);
				return { content: [{ type: 'text', text: 'Fixture request A-123' }] };
			},
		});
		// The server IDs intentionally match; collections distinguish the accounts.
		const work = fakeServer({ name: 'Work Mail', collectionId: 'work', tools: [makeTool('read_work', true), makeTool('send_work', false)] });
		const personal = fakeServer({ name: 'Personal Mail', collectionId: 'personal', tools: [makeTool('read_personal', true)] });
		const mcp = mcpServiceWith(work, personal);
		let requestCount = 0;
		let newsOutage = false;
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async readAssistantNewsFeeds() {
				assert.ok(newsOutage);
				return [{ publisher: 'CBC News', error: 'HTTP 503' }, { publisher: 'Global News', xml: '<html />' }];
			}
			override async sendAssistantPrompt(_account: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				const names = request.tools?.map(tool => tool.name) ?? [];
				assert.ok(names.includes('read_work'));
				assert.ok(!names.includes('read_personal') && !names.includes('send_work') && !names.includes('assistant.start_agent'));
				assert.doesNotMatch(request.messages[0].content, /Private memory from another connection/);
				assert.match(request.messages[1].content, /Summarize open support requests/);
				if (++requestCount === 1) {
					return {
						kind: 'reply', text: '', toolCalls: [
							{ id: 'sources', name: 'assistant.list_connection_sources', arguments: '{}' },
							newsOutage
								? { id: 'news', name: 'assistant.get_news_headlines', arguments: '{}' }
								: { id: 'foreign', name: 'read_personal', arguments: '{}' },
							{ id: 'write', name: 'send_work', arguments: '{}' },
							{ id: 'read', name: 'read_work', arguments: '{}' },
						]
					};
				}
				const results = request.messages.filter(message => message.role === 'tool').map(message => message.content);
				assert.deepStrictEqual(JSON.parse(results[0]).available, ['Work Mail']);
				if (newsOutage) {
					assert.match(results[1], /No current canada headlines.*CBC News: HTTP 503.*Global News: Invalid RSS response/u);
				} else {
					assert.strictEqual(results[1], 'Error: unknown tool "read_personal"');
				}
				assert.deepStrictEqual(results.slice(2), ['Error: unknown tool "send_work"', 'Fixture request A-123']);
				return { kind: 'reply', text: 'Open requests report' };
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, mcp, upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		const settings = {
			id: 'fixture-routine', name: 'Support triage', prompt: 'Summarize open support requests', serverIds: [JSON.stringify(['work', 'notes-server'])],
		};
		assert.strictEqual(await runs.sendScheduledRoutine('account', settings, 'full-report'), 'completed');
		assert.strictEqual(requestCount, 2);
		assert.deepStrictEqual(called, ['read_work']);
		const state = conversation.read('account');
		assert.strictEqual(state.draft, 'Unsent draft');
		assert.strictEqual(state.turns.at(-1)?.routineId, 'fixture-routine');
		assert.strictEqual(state.turns.at(-1)?.routineRunId, 'full-report');
		assert.strictEqual(state.turns.at(-1)?.reply, 'Open requests report');
		assert.deepStrictEqual(flushedRunIds, ['full-report']);
		requestCount = 0;
		assert.strictEqual(await runs.sendScheduledRoutine('account', {
			...settings, serverIds: [...settings.serverIds, 'missing-connection'],
		}, 'partial-report'), 'needsAttention');
		assert.strictEqual(conversation.read('account').turns.at(-1)?.routineRunId, 'partial-report');
		assert.match(conversation.read('account').turns.at(-1)?.reply ?? '', /1 saved connection — No longer configured/);
		assert.strictEqual(conversation.read('account').turns.at(-1)?.unavailableConnections, undefined);
		assert.deepStrictEqual(flushedRunIds, ['full-report', 'partial-report']);
		requestCount = 0;
		newsOutage = true;
		assert.strictEqual(await runs.sendScheduledRoutine('account', {
			...settings, prompt: 'Summarize open support requests and current headlines',
		}, 'news-outage'), 'needsAttention');
		const outage = conversation.read('account').turns.at(-1);
		assert.strictEqual(outage?.reply, 'Open requests report');
		assert.strictEqual(outage?.activity?.find(item => item.tool === 'assistant.get_news_headlines')?.status, 'failed');
	});

	test('a routine can use a selected direct mailbox without starting an MCP connection', async () => {
		const storage = store.add(new InMemoryStorageService());
		const connection = { id: 'direct-outlook-fixture', provider: 'outlook' as const, label: 'Outlook (Direct)', needsSignIn: false };
		const gmail = { id: 'direct-gmail-fixture', provider: 'gmail' as const, label: 'Gmail (Direct)', needsSignIn: false };
		const directMail = upcastPartial<ISignalCDirectMail>({
			connections: constObservable([connection, gmail]),
			read: async (selected, args) => {
				assert.deepStrictEqual(selected, connection);
				assert.deepStrictEqual(args, { folder: 'sent' });
				return '{"messages":[{"id":"mail-1","subject":"Sent test message"}]}';
			},
		});
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_account: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				assert.ok(request.tools?.some(tool => tool.name === 'assistant.outlook_read'));
				assert.ok(!request.tools?.some(tool => tool.name === 'assistant.gmail_read'));
				if (request.messages.at(-1)?.role === 'tool') {
					assert.match(request.messages.at(-1)?.content ?? '', /Sent test message/);
					return { kind: 'reply', text: 'Outlook sent-mail report' };
				}
				return { kind: 'reply', text: '', toolCalls: [{ id: 'outlook', name: 'assistant.outlook_read', arguments: '{"folder":"sent"}' }] };
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, emptyMcpService(), upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMail, configurationStub, mailWatcherStub, promptsStub));
		assert.strictEqual(await runs.sendScheduledRoutine('account', { id: 'check-mail', name: 'Sent mail', prompt: 'Summarize sent mail', serverIds: [connection.id] }, 'direct-report'), 'completed');
		const turn = new AssistantConversationStore(storage).read('account').turns.at(-1);
		assert.strictEqual(turn?.reply, 'Outlook sent-mail report');
		assert.strictEqual(turn?.activity?.find(item => item.tool === 'assistant.outlook_read')?.detail, undefined);
	});

	test('reports signed-out Outlook even when the model omits it and keeps reading the other selected source', async () => {
		const storage = store.add(new InMemoryStorageService());
		const conversation = new AssistantConversationStore(storage);
		const readCalls: string[] = [];
		const outlook = fakeServer({
			name: 'Work Outlook', collectionId: 'outlook', url: microsoft365WorkIq.url,
			state: { state: McpConnectionState.Kind.Stopped, reason: 'needs-user-interaction' },
			tools: [upcastPartial<IMcpTool>({
				id: 'cached_outlook_read', visibility: McpToolVisibility.Model, definition: { name: 'fetch', inputSchema: { type: 'object' } },
				call: async () => assert.fail('A cached read must not retry a known signed-out connection'),
			})],
		});
		const mail = fakeServer({
			name: 'Other work mail', collectionId: 'mail', tools: [upcastPartial<IMcpTool>({
				id: 'read_mail', visibility: McpToolVisibility.Model,
				definition: { name: 'read_mail', inputSchema: { type: 'object' }, annotations: { readOnlyHint: true } },
				call: async (_args, context) => {
					assert.strictEqual(context?.errorOnUserInteraction, true);
					readCalls.push('mail');
					return { content: [{ type: 'text', text: 'Fixture case A-100' }] };
				},
			})]
		});
		const ignored = fakeServer({ name: 'Unselected account', collectionId: 'unselected', state: { state: McpConnectionState.Kind.Error, message: 'Private provider error' } });
		const mcp = mcpServiceWith(outlook, mail, ignored);
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_account: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				assert.ok(request.tools?.some(tool => tool.name === 'read_mail'));
				assert.ok(!request.tools?.some(tool => tool.name === 'cached_outlook_read'));
				const results = request.messages.filter(message => message.role === 'tool');
				if (!results.length) {
					return {
						kind: 'reply', text: '', toolCalls: [
							{ id: 'sources', name: 'assistant.list_connection_sources', arguments: '{}' },
							{ id: 'mail', name: 'read_mail', arguments: '{}' },
						]
					};
				}
				assert.deepStrictEqual(JSON.parse(results[0].content), {
					available: ['Other work mail'], inactive: [], unavailable: [{ name: 'Work Outlook', reason: 'Sign-in required' }],
				});
				assert.strictEqual(results[1].content, 'Fixture case A-100');
				return { kind: 'reply', text: 'Found case A-100.' };
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, mcp, upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		const routine = {
			id: 'fixture-routine', name: 'Inbox review', prompt: 'Report inbox cases.',
			serverIds: [JSON.stringify(['outlook', 'notes-server']), JSON.stringify(['mail', 'notes-server'])],
		};
		assert.strictEqual(await runs.sendScheduledRoutine('account', routine, 'partial-outlook-report'), 'needsAttention');
		assert.deepStrictEqual(readCalls, ['mail']);
		assert.strictEqual(conversation.read('account').turns.at(-1)?.reply, 'Found case A-100.\n\n**Unavailable connections**\n\n- Work Outlook — Sign-in required');
		assert.deepStrictEqual(conversation.read('account').turns.at(-1)?.unavailableConnections, [{ id: JSON.stringify(['outlook', 'notes-server']), name: 'Work Outlook', needsSignIn: true }]);
	});

	test('discovers Outlook by name and reads structured Work IQ mail in a briefing', async () => {
		const storage = store.add(new InMemoryStorageService());
		const mailPath = '/me/mailFolders/inbox/messages?$top=3';
		let reads = 0;
		const workIq = fakeServer({
			name: 'Work IQ', url: microsoft365WorkIq.url, tools: [
				upcastPartial<IMcpTool>({
					id: 'workiq_fetch', visibility: McpToolVisibility.Model,
					definition: { name: 'fetch', inputSchema: { type: 'object', properties: { entityUrls: { type: 'array', items: { type: 'string' } } } } },
					call: async args => {
						assert.deepStrictEqual(args, { entityUrls: [mailPath] });
						reads++;
						return { content: [], structuredContent: { results: [{ statusCode: 200, data: { value: [{ subject: 'Fixture mail in the inbox' }] } }] } };
					},
				}),
				upcastPartial<IMcpTool>({
					id: 'workiq_action', visibility: McpToolVisibility.Model,
					definition: { name: 'do_action', inputSchema: { type: 'object' }, annotations: { readOnlyHint: true } },
					call: async () => assert.fail('A Work IQ action must never be exposed in a briefing'),
				}),
			]
		});
		const unrelated = fakeServer({
			name: 'Work IQ', url: 'https://example.com/mcp', tools: [upcastPartial<IMcpTool>({
				id: 'unrelated_fetch', visibility: McpToolVisibility.Model,
				definition: { name: 'fetch', inputSchema: { type: 'object' } },
				call: async () => assert.fail('A matching name cannot inherit Work IQ read-only treatment'),
			})]
		});
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_account: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				const tools = request.tools?.map(tool => tool.name) ?? [];
				assert.ok(tools.includes('workiq_fetch'));
				assert.ok(!tools.includes('workiq_action') && !tools.includes('unrelated_fetch'));
				const result = request.messages.at(-1)?.content ?? '';
				if (result.includes('Fixture mail in the inbox')) { return { kind: 'reply', text: 'Outlook briefing ready.' }; }
				if (request.messages.at(-1)?.role === 'tool') {
					assert.match(result, /Available in the next request.*workiq_fetch/s);
					return { kind: 'reply', text: '', toolCalls: [{ id: 'read', name: 'workiq_fetch', arguments: JSON.stringify({ entityUrls: [mailPath] }) }] };
				}
				return { kind: 'reply', text: '', toolCalls: [{ id: 'discover', name: 'assistant.find_app_tools', arguments: '{"app":"Outlook","query":"search email"}' }] };
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, mcpServiceWith(workIq, unrelated), upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		assert.strictEqual(await runs.send('account', '/briefing morning'), true);
		assert.strictEqual(reads, 1);
		assert.strictEqual(new AssistantConversationStore(storage).read('account').turns.at(-1)?.reply, 'Outlook briefing ready.');
	});

	test('stages a routine proposal in a chat turn, including editable schedule defaults', async () => {
		const storage = store.add(new InMemoryStorageService());
		const requests: ReachAssistantRequest[] = [];
		const proposal = { name: 'Calendar review', prompt: 'Summarize tomorrow’s meetings', cadence: 'weekdays', hour: 16, minute: 30 };
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_account: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				requests.push(request);
				if (requests.length === 1) {
					assert.ok(request.tools?.some(tool => tool.name === 'assistant.propose_routine'));
					return {
						kind: 'reply', text: '', toolCalls: [{
							id: 'proposal', name: 'assistant.propose_routine', arguments: JSON.stringify({ ...proposal, serverIds: ['injected'] }),
						}]
					};
				}
				assert.match(request.messages.at(-1)?.content ?? '', /Review routine/);
				return { kind: 'reply', text: 'Here is a routine to review.' };
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, emptyMcpService(), upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		assert.strictEqual(await runs.send('account', 'Set up a weekday calendar review at 16:30'), true);
		assert.strictEqual(requests.length, 2);
		assert.deepStrictEqual(new AssistantConversationStore(storage).read('account').turns[0].routineDrafts, [proposal]);
	});

	test('stages every routine drafted in one turn so each draft keeps its own review', async () => {
		const storage = store.add(new InMemoryStorageService());
		const requests: ReachAssistantRequest[] = [];
		const morning = { name: 'Morning Briefing', prompt: 'Toronto weather, Canadian headlines, AI stories, Outlook inbox', cadence: 'weekdays', hour: 7, minute: 0 };
		const evening = { name: 'Evening Briefing', prompt: 'Canadian headlines, AI stories, tomorrow forecast, Outlook inbox', cadence: 'weekdays', hour: 18, minute: 0 };
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_account: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				requests.push(request);
				if (requests.length === 1) {
					return {
						kind: 'reply', text: '', toolCalls: [
							{ id: 'morning', name: 'assistant.propose_routine', arguments: JSON.stringify(morning) },
							{ id: 'evening', name: 'assistant.propose_routine', arguments: JSON.stringify(evening) },
						]
					};
				}
				return { kind: 'reply', text: 'Both briefings are drafted!' };
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, emptyMcpService(), upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		assert.strictEqual(await runs.send('account', '/briefing set up a morning briefing and an evening briefing'), true);
		assert.deepStrictEqual(new AssistantConversationStore(storage).read('account').turns[0].routineDrafts, [morning, evening]);
	});

	test('learns from a connected result and saves only the source-backed memory selected in review', async () => {
		const storage = store.add(new InMemoryStorageService());
		const conversations = new AssistantConversationStore(storage);
		const requests: ReachAssistantRequest[] = [];
		const mcpService = mcpServiceWith(fakeServer({
			name: 'Gmail work', tools: [upcastPartial<IMcpTool>({
				id: 'mcp_sent', visibility: McpToolVisibility.Model,
				definition: { name: 'list_sent', inputSchema: { type: 'object' }, annotations: { readOnlyHint: true } },
				call: async () => ({ content: [{ type: 'text', text: 'message-1234: fixture email content\nmessage-5678: second fixture email' }] }),
			})]
		}));
		const propose = (id: string, note: string, reference: string) => ({
			id, name: 'assistant.propose_memory', arguments: JSON.stringify({ note, tool: 'mcp_sent', reference }),
		});
		const responses: ReachAssistantResponse[] = [
			{ kind: 'reply', text: '', toolCalls: [{ id: 'read', name: 'mcp_sent', arguments: '{}' }] },
			{
				kind: 'reply', text: '', toolCalls: [
					propose('valid', 'Prefers short weekly updates', 'message-1234'),
					propose('other', 'Prefers to confirm dates in writing', 'message-5678'),
					propose('invalid', 'Unfounded statement', 'unverified-reference'),
				]
			},
			{ kind: 'reply', text: 'Found two candidates from sent email.' },
		];
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_account: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				requests.push(request);
				return responses.shift()!;
			}
		};
		const quickInput = new class extends mock<IQuickInputService>() {
			override async pick(items: unknown[]): Promise<never> {
				assert.strictEqual(conversations.read('account').instructions, '');
				assert.deepStrictEqual(items.map(item => (item as { label: string }).label), [
					'Prefers short weekly updates', 'Prefers to confirm dates in writing',
				]);
				return [items[0]] as never;
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, mcpService, upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInput, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		assert.strictEqual(await runs.send('account', '/learn'), true);
		assert.match(requests[0].messages[1].content, /both my sent and received email/);
		assert.match(requests[0].messages[1].content, /my commits, pull requests, reviews, and issue comments/);
		assert.match(requests[2].messages.find(message => message.toolCallId === 'invalid')?.content ?? '', /exact link or record ID/);
		assert.strictEqual(conversations.read('account').instructions, '- Prefers short weekly updates (Gmail work · list_sent, message-1234)');
		assert.strictEqual(conversations.read('other account').instructions, '');
		assert.doesNotMatch(JSON.stringify(conversations.read('account')), /fixture email content/);
	});

	test('records a provider-reported MCP failure as failed and gives its reason to the model', async () => {
		const storage = store.add(new InMemoryStorageService());
		const reason = 'This Google Cloud project requires Google Workspace Developer Preview enrollment.';
		const tool = upcastPartial<IMcpTool>({
			id: 'list_labels', visibility: McpToolVisibility.Model,
			definition: { name: 'list_labels', inputSchema: { type: 'object' }, annotations: { readOnlyHint: true } },
			call: async () => ({ isError: true, content: [{ type: 'text', text: reason }] }),
		});
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_account: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				if (request.messages.at(-1)?.role === 'tool') {
					assert.strictEqual(request.messages.at(-1)?.content, `Tool reported an error:\n${reason}`);
					return { kind: 'reply', text: 'Gmail is blocked by project enrollment.' };
				}
				return { kind: 'reply', text: '', toolCalls: [{ id: 'read-labels', name: 'list_labels', arguments: '{}' }] };
			}
		};
		const mcp = mcpServiceWith(fakeServer({ name: 'Gmail', tools: [tool] }));
		const runs = store.add(new AssistantRuns(reach, storage, mcp, upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		assert.strictEqual(await runs.send('account', 'List Gmail labels'), true);
		const [activity] = new AssistantConversationStore(storage).read('account').turns[0].activity!;
		assert.strictEqual(activity.status, 'failed');
		assert.strictEqual(activity.detail, reason);
		assert.strictEqual(await runs.sendScheduledRoutine('account', {
			id: 'gmail-review', name: 'Mail review', prompt: 'Read Gmail labels', serverIds: [JSON.stringify(['user', 'notes-server'])],
		}, 'gmail-run'), 'needsAttention');
		const lastReport = new AssistantConversationStore(storage).read('account').turns.at(-1);
		assert.strictEqual(lastReport?.routineRunId, 'gmail-run');
		assert.strictEqual(lastReport?.activity?.[0].status, 'failed');
	});

	test('routes Microsoft mail names to Outlook and keeps Work IQ explicitly available', () => {
		for (const name of ['Outlook', 'Outlook.com', 'Outlook (Direct)', 'Microsoft 365', 'Microsoft 365 Mail', 'Outlook / Microsoft 365 Mail']) {
			assert.strictEqual(findAssistantAppConnection(name)?.id, 'outlook-direct');
		}
		for (const name of ['Work IQ', 'Microsoft 365 (Work IQ)', 'Microsoft Teams']) {
			assert.strictEqual(findAssistantAppConnection(name)?.id, 'microsoft-365');
		}
	});

	test('opens only approved catalog app setup through Assistant', async () => {
		const storage = store.add(new InMemoryStorageService());
		const calls: unknown[][] = [];
		const toolResults: string[] = [];
		const approvalLabels: string[] = [];
		const setup = new DeferredPromise<unknown>();
		const commands = upcastPartial<ICommandService>({ executeCommand: async <T>(...args: unknown[]) => { calls.push(args); return await setup.p as T; } });
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_account: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				const results = request.messages.filter(message => message.role === 'tool');
				if (results.length) {
					toolResults.push(results.at(-1)!.content);
					return { kind: 'reply', text: 'OAuth setup result received.' };
				}
				const app = request.messages.at(-1)?.content;
				return { kind: 'reply', text: '', toolCalls: [{ id: 'oauth', name: 'assistant.connect_app', arguments: JSON.stringify({ app }) }] };
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, emptyMcpService(), upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, commands, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		let allow = false;
		store.add(runs.onDidRequestApproval(request => {
			approvalLabels.push(request.label);
			runs.respondApproval(request.id, allow ? 'allow' : 'reject');
		}));
		assert.strictEqual(await runs.send('account', 'Slack'), true);
		assert.strictEqual(calls.length, 0);
		allow = true;
		assert.strictEqual(await runs.send('account', 'https://untrusted.example/oauth'), true);
		assert.strictEqual(calls.length, 0);
		assert.strictEqual(await runs.send('account', 'Slack'), true);
		assert.deepStrictEqual(calls, [[McpCommandIds.ConnectApp, undefined, 'slack', 'assistant']]);
		assert.strictEqual(runs.openAppSetup('slack'), false);
		assert.strictEqual(approvalLabels.length, 3);
		assert.match(toolResults[0], /declined/);
		assert.match(toolResults[1], /No matching app/);
		assert.match(toolResults[2], /finish the setup and any browser authorization/);
		await setup.complete(undefined);
	});

	test('discovers Slack tools beyond the model limit and still asks before sending', async () => {
		const storage = store.add(new InMemoryStorageService());
		const sent: unknown[] = [];
		const slackTool = upcastPartial<IMcpTool>({
			id: 'mcp_slack_send_message', visibility: McpToolVisibility.Model,
			definition: { name: 'slack_send_message', description: 'Send a message to a Slack channel', inputSchema: { type: 'object', properties: { text: { type: 'string' } } } },
			call: async args => { sent.push(args); return { content: [{ type: 'text', text: 'Synthetic Slack delivery confirmed' }] }; },
		});
		const mcpService = upcastPartial<IMcpService>({
			servers: constObservable([
				fakeServer({ toolCount: 40, instructions: 'Fixture notes instructions' }),
				fakeServer({ name: 'Slack', tools: [slackTool], instructions: 'Fixture Slack instructions' }),
			])
		});
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_account: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				assert.ok(request.tools && request.tools.length <= 32);
				const results = request.messages.filter(message => message.role === 'tool');
				if (!results.length) {
					assert.ok(!request.tools.some(spec => spec.name === slackTool.id));
					assert.match(request.messages[0].content, /Fixture notes instructions/);
					assert.doesNotMatch(request.messages[0].content, /Fixture Slack instructions/);
					return { kind: 'reply', text: '', toolCalls: [{ id: 'find', name: 'assistant.find_app_tools', arguments: '{"app":"Slack","query":"send message"}' }] };
				}
				if (results.length === 1) {
					assert.match(results[0].content, /mcp_slack_send_message/);
					assert.ok(request.messages[0].content.includes(slackTool.id));
					assert.match(request.messages[0].content, /Fixture Slack instructions/);
					assert.doesNotMatch(request.messages[0].content, /Fixture notes instructions/);
					assert.deepStrictEqual(request.tools.find(spec => spec.name === slackTool.id)?.parameters, slackTool.definition.inputSchema);
					return { kind: 'reply', text: '', toolCalls: [{ id: 'send', name: slackTool.id, arguments: '{"text":"Test message"}' }] };
				}
				assert.match(results[1].content, /Synthetic Slack delivery confirmed/);
				return { kind: 'reply', text: 'Message sent.' };
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, mcpService, upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		let approvedTool: string | undefined;
		store.add(runs.onDidRequestApproval(request => {
			assert.deepStrictEqual(sent, []);
			approvedTool = request.tool;
			runs.respondApproval(request.id, 'allow');
		}));
		assert.strictEqual(await runs.send('account', 'Send a test message in Slack'), true);
		assert.strictEqual(approvedTool, slackTool.id);
		assert.deepStrictEqual(sent, [{ text: 'Test message' }]);
	});

	test('reaches the end of a large tool catalog and hides cached tools that need sign-in', async () => {
		const storage = store.add(new InMemoryStorageService());
		const calls: string[] = [];
		const lastId = 'last_notes_tool';
		const disconnected = fakeServer({
			name: 'Sign-in Fixture', state: { state: McpConnectionState.Kind.Stopped, reason: 'needs-user-interaction' }, tools: [
				upcastPartial<IMcpTool>({ id: 'disconnected_tool', visibility: McpToolVisibility.Model, definition: { name: 'disconnected_tool', inputSchema: { type: 'object' } } }),
			]
		});
		const fullCatalog = fakeServer({
			name: 'Notes', tools: [...fakeServer({ toolCount: 32 }).tools.get(), upcastPartial<IMcpTool>({
				id: lastId, visibility: McpToolVisibility.Model,
				definition: { name: lastId, inputSchema: { type: 'object' }, annotations: { readOnlyHint: true } },
				call: async () => { calls.push(lastId); return { content: [{ type: 'text', text: 'Last tool returned data' }] }; },
			})]
		});
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_account: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				const tools = request.tools ?? [];
				assert.ok(tools.length <= 32 && !tools.some(tool => tool.name === 'disconnected_tool'));
				const results = request.messages.filter(message => message.role === 'tool');
				if (!results.length) {
					assert.ok(!tools.some(tool => tool.name === lastId));
					return { kind: 'reply', text: '', toolCalls: [{ id: 'index', name: 'assistant.find_app_tools', arguments: '{"app":"Notes"}' }] };
				}
				if (results.length === 1) {
					assert.match(results[0].content, /33 matching tools/);
					assert.match(results[0].content, /Other matching tool IDs[^\n]+last_notes_tool/);
					assert.ok(!tools.some(tool => tool.name === lastId));
					return { kind: 'reply', text: '', toolCalls: [{ id: 'select', name: 'assistant.find_app_tools', arguments: JSON.stringify({ app: 'Notes', name: lastId }) }] };
				}
				if (results.length === 2) {
					assert.ok(tools.some(tool => tool.name === lastId));
					return { kind: 'reply', text: '', toolCalls: [{ id: 'last', name: lastId, arguments: '{}' }] };
				}
				assert.strictEqual(results[2].content, 'Last tool returned data');
				return { kind: 'reply', text: 'Notes checked.' };
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, mcpServiceWith(disconnected, fullCatalog), upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		assert.strictEqual(await runs.send('account', 'Use the last Notes tool'), true);
		assert.deepStrictEqual(calls, [lastId]);
	});

	test('offers an authorized direct-mail send with the complete message for approval', async () => {
		const storage = store.add(new InMemoryStorageService());
		const outgoing = { to: 'recipient@example.test', subject: 'Fixture subject', body: 'Long fixture body. '.repeat(60) };
		const calls: unknown[] = [];
		const mailbox = { id: 'fixture-mailbox', provider: 'outlook' as const, label: 'Outlook', needsSignIn: false, canSend: true as const };
		const gmail = { id: 'fixture-gmail', provider: 'gmail' as const, label: 'Gmail (Direct)', needsSignIn: false };
		const connections = observableValue<readonly (typeof gmail | typeof mailbox)[]>('sendTestMailboxes', [gmail, mailbox]);
		const mail = upcastPartial<ISignalCDirectMail>({
			connections,
			send: async (_connection, args) => { calls.push(args); return 'Accepted by the fixture provider.'; },
		});
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_account: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				if (request.messages.at(-1)?.role === 'tool') { return { kind: 'reply', text: 'Accepted.' }; }
				assert.ok(request.tools?.some(tool => tool.name === 'assistant.outlook_send'));
				assert.ok(!request.tools?.some(tool => tool.name === 'assistant.gmail_send'));
				assert.ok(request.messages.some(message => message.role === 'system' && message.content.includes('Gmail (Direct): read-only')));
				return {
					kind: 'reply', text: '', toolCalls: [
						{ id: 'send', name: 'assistant.outlook_send', arguments: JSON.stringify(outgoing) },
						{ id: 'duplicate', name: 'assistant.outlook_send', arguments: JSON.stringify({ body: outgoing.body, subject: outgoing.subject, to: outgoing.to }) },
					]
				};
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, emptyMcpService(), upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, mail, configurationStub, mailWatcherStub, promptsStub));
		let approvals = 0;
		store.add(runs.onDidRequestApproval(request => {
			assert.strictEqual(calls.length, approvals ? 2 : 0);
			assert.strictEqual(request.argsPreview, `To: ${outgoing.to}\nSubject: ${outgoing.subject}\n\n${outgoing.body}`);
			approvals++;
			runs.respondApproval(request.id, 'always');
		}));
		assert.strictEqual(await runs.send('account', 'Send the fixture email with Outlook'), true);
		assert.strictEqual(approvals, 1);
		assert.strictEqual(await runs.send('account', 'Send another fixture email with Outlook'), true);
		assert.strictEqual(approvals, 1);
		connections.set([gmail, { ...mailbox, id: 'replacement-mailbox' }], undefined);
		assert.strictEqual(await runs.send('account', 'Send a fixture email with the replacement Outlook mailbox'), true);
		assert.strictEqual(approvals, 2);
		assert.deepStrictEqual(calls, [outgoing, outgoing, outgoing]);
	});

	test('keeps send available after ready and reports only a verified mail outcome', async () => {
		const storage = store.add(new InMemoryStorageService());
		const conversation = new AssistantConversationStore(storage);
		const to = 'recipient@example.test';
		conversation.appendTurn('account', { prompt: `Send an email to ${to}`, reply: 'Say ready to send.' });
		const gmail = { id: 'fixture-gmail', provider: 'gmail' as const, label: 'Gmail (Direct)', needsSignIn: false, canSend: true as const };
		let sends = 0;
		let requests = 0;
		const mail = upcastPartial<ISignalCDirectMail>({
			connections: constObservable([gmail]),
			send: async (_connection, args) => {
				assert.strictEqual(args.to, to);
				sends++;
				return 'Accepted by the fixture provider.';
			},
		});
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_account: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				assert.strictEqual(request.tools?.find(tool => /gmail_/.test(tool.name))?.name, 'assistant.gmail_send');
				if (++requests === 1) {
					return { kind: 'reply', text: '**Status: Email sent**\n\nThe draft was created.' };
				}
				return { kind: 'reply', text: '', toolCalls: [{ id: 'send-mail', name: 'assistant.gmail_send', arguments: JSON.stringify({ to, subject: 'Fixture', body: 'Fixture message' }) }] };
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, emptyMcpService(), upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, mail, configurationStub, mailWatcherStub, promptsStub));
		store.add(runs.onDidRequestApproval(request => runs.respondApproval(request.id, 'allow')));
		assert.strictEqual(await runs.send('account', 'ready'), true);
		assert.match(conversation.read('account').turns.at(-1)?.reply ?? '', /No email send was confirmed/);
		assert.strictEqual(sends, 0);
		assert.strictEqual(await runs.send('account', 'ready'), true);
		assert.match(conversation.read('account').turns.at(-1)?.reply ?? '', /accepted the email for sending/);
		assert.strictEqual(await runs.send('account', 'ready'), true);
		assert.match(conversation.read('account').turns.at(-1)?.reply ?? '', /already accepted this email/);
		// Regenerating the confirmed delivery must keep its receipt for subsequent retries.
		for (let repeat = 0; repeat < 2; repeat++) {
			assert.strictEqual(await runs.send('account', 'ready', undefined, { replaceFrom: 2 }), true);
			assert.match(conversation.read('account').turns.at(-1)?.reply ?? '', /already accepted this email/);
		}
		assert.strictEqual(requests, 2);
		assert.strictEqual(sends, 1);
	});

	for (const results of [[true], [false], [true, false]]) {
		test(`reuses connected mail outcomes ${results.join(', ')} across regenerations and confirmations`, async () => {
			const storage = store.add(new InMemoryStorageService());
			const conversation = new AssistantConversationStore(storage);
			const recipient = 'fixture@example.test';
			const body = 'Private synthetic content';
			const id = 'mcp_fixture_dispatch';
			let sends = 0;
			let requests = 0;
			const mailTool = upcastPartial<IMcpTool>({
				id, visibility: McpToolVisibility.Model,
				definition: { name: 'send_email', inputSchema: { type: 'object', properties: { to: { type: 'string' }, body: { type: 'string' } } } },
				call: async () => ({ isError: !results[sends++], content: [{ type: 'text', text: `${recipient}: ${body}` }] }),
			});
			const reach = new class extends mock<IReachSubscriptionConnection>() {
				override readonly onDidChangeAccount = Event.None;
				override async sendAssistantPrompt(_account: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
					requests++;
					assert.ok(request.tools?.some(tool => tool.name === id));
					const calls = results.map((_, index) => ({ id: String(index), name: id, arguments: JSON.stringify({ to: recipient, body: body + index }) }));
					calls.splice(1, 0, { id: 'duplicate', name: id, arguments: JSON.stringify({ body: body + 0, to: recipient }) });
					return { kind: 'reply', text: '', toolCalls: calls };
				}
			};
			const runs = store.add(new AssistantRuns(reach, storage, mcpServiceWith(fakeServer({ name: 'Mailbox', tools: [mailTool] })), upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
			store.add(runs.onDidRequestApproval(request => runs.respondApproval(request.id, 'allow')));
			const prompt = 'Send the fixture email.';
			assert.strictEqual(await runs.send('account', prompt), true);
			const expected = { accepted: results.filter(Boolean).length, unconfirmed: results.includes(false) };
			const check = () => {
				const turn = conversation.read('account').turns.at(-1)!;
				assert.deepStrictEqual(turn.mailDelivery, expected);
				assert.strictEqual(/accepted/.test(turn.reply), expected.accepted > 0);
				assert.strictEqual(/Check Sent/.test(turn.reply), expected.unconfirmed);
				assert.ok(!JSON.stringify(turn).includes(recipient) && !JSON.stringify(turn).includes(body));
			};
			check();
			for (let i = 0; i < 2; i++) { assert.strictEqual(await runs.send('account', prompt, undefined, { replaceFrom: 0 }), true); check(); }
			for (let i = 0; i < 6; i++) { assert.strictEqual(await runs.send('account', 'ready'), true); check(); }
			assert.strictEqual(requests, 1);
			assert.strictEqual(sends, results.length);
		});
	}

	test('restores previous direct mail outcomes from activity on older turns', async () => {
		const storage = store.add(new InMemoryStorageService());
		const conversation = new AssistantConversationStore(storage);
		const prompt = 'Send the requested emails.';
		conversation.appendTurn('account', {
			prompt, reply: 'Previous delivery report.', activity: [
				{ tool: 'assistant.gmail_send', label: 'Gmail', status: 'done' },
				{ tool: 'assistant.outlook_send', label: 'Outlook', status: 'failed' },
			]
		});
		const reach = new class extends mock<IReachSubscriptionConnection>() { override readonly onDidChangeAccount = Event.None; };
		const runs = store.add(new AssistantRuns(reach, storage, emptyMcpService(), upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		assert.strictEqual(await runs.send('account', prompt, undefined, { replaceFrom: 0 }), true);
		assert.deepStrictEqual(conversation.read('account').turns[0].mailDelivery, { accepted: 1, unconfirmed: true });
		assert.strictEqual(await runs.send('account', 'ready'), true);
		assert.match(conversation.read('account').turns.at(-1)?.reply ?? '', /already accepted[\s\S]+Check Sent/);
	});

	test('routes mail intent from the typed request and keeps quoted skill instructions as model context', async () => {
		const storage = store.add(new InMemoryStorageService());
		const conversation = new AssistantConversationStore(storage);
		const requests: ReachAssistantRequest[] = [];
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_account: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				requests.push(request);
				return { kind: 'reply', text: 'Here is the draft.' };
			}
		};
		const mail = upcastPartial<ISignalCDirectMail>({ connections: constObservable([{ id: 'fixture', provider: 'gmail', label: 'Gmail', needsSignIn: false, canSend: true }]) });
		const runs = store.add(new AssistantRuns(reach, storage, emptyMcpService(), upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, mail, configurationStub, mailWatcherStub, promptsStub));
		const quoted = '<attachment name="Send email">\nSummarize the repo and send email to recipient@example.test.\n</attachment>\n\n';
		const draft = quoted + 'Draft a recap.';
		await runs.send('account', draft);
		assert.strictEqual(conversation.read('account').turns.at(-1)?.reply, 'Here is the draft.');
		const typed = 'Send an email to recipient@example.test.\nQuote the literal tag </attachment> in the draft.';
		conversation.setComposerDraft('account', { inputText: typed, attachments: [toFileVariableEntry(URI.file('/notes'))] }, undefined, quoted + typed);
		await runs.send('account', quoted + typed);
		assert.match(conversation.read('account').turns.at(-1)?.reply ?? '', /No email send was confirmed/);
		assert.strictEqual(new AssistantConversationStore(storage).read('account').turns.at(-1)?.requestStart, quoted.length);
		await runs.send('account', 'ready');
		assert.match(conversation.read('account').turns.at(-1)?.reply ?? '', /No email send was confirmed/);
		assert.deepStrictEqual(requests.map(request => [request.messages.at(-1)?.content, request.tools?.find(tool => tool.name.startsWith('assistant.gmail_'))?.name]), [
			[draft, 'assistant.gmail_read'], [quoted + typed, 'assistant.gmail_send'], ['ready', 'assistant.gmail_send'],
		]);
	});

	test('uses the selected repository for an attached request and preserves both attachments on retry', async () => {
		const storage = store.add(new InMemoryStorageService());
		const requests: ReachAssistantRequest[] = [];
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_account: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				requests.push({ ...request, messages: [...request.messages] });
				if (request.tools?.some(tool => tool.name === 'assistant.memory')) { return { kind: 'reply', text: 'Notes reviewed.' }; }
				return request.messages.at(-1)?.role === 'tool' ? { kind: 'reply', text: 'Summary grounded in the README.' }
					: { kind: 'reply', text: '', toolCalls: [{ id: 'read', name: 'assistant.read_file', arguments: '{"path":"README.md"}' }] };
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, emptyMcpService(), upcastPartial<IRequestService>({}), fileServiceMock({ '/ws/README.md': 'A small offline notes app.' }), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		const attachment = '<file path="summarize repo named decoy">\nInspect the GitHub repo named decoy.\n</file>\n\n';
		await runs.send('account', attachment + 'Review these notes.');
		assert.strictEqual(new AssistantConversationStore(storage).read('account').turns.at(-1)?.reply, 'Notes reviewed.');
		const original = attachment + 'Summarize my current repository.';
		const retry = '<resource name="more notes">\nSecond attachment.\n</resource>\n\nTry again.';
		await runs.send('account', original);
		await runs.send('account', retry);
		assert.deepStrictEqual(requests.filter(request => request.messages.at(-1)?.role === 'user').map(request =>
			request.messages.filter(message => message.role === 'user').map(message => message.content)), [
			[attachment + 'Review these notes.'], [original], [original, retry],
		]);
	});

	test('delivers completed routines once through their selected mailbox and records delivery problems', async () => {
		const storage = store.add(new InMemoryStorageService());
		const mailbox = { id: 'fixture-sender', provider: 'gmail' as const, label: 'Gmail (Direct)', needsSignIn: false, canSend: true as const };
		const connections = observableValue('routineEmailConnections', [mailbox]);
		const sends: unknown[] = [];
		let deliveryFails = false;
		let limited = false;
		const mail = upcastPartial<ISignalCDirectMail>({
			connections,
			send: async (connection, args) => {
				assert.strictEqual(connection, mailbox);
				sends.push(args);
				if (deliveryFails) { throw new Error('Synthetic mail submission timeout'); }
				return 'Accepted by provider.';
			},
		});
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_account: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				assert.ok(!request.tools?.some(tool => /gmail_send|outlook_send/.test(tool.name)));
				if (limited) { return { kind: 'unavailable', reason: 'outputLimit' }; }
				return { kind: 'reply', text: 'Scheduled fixture message.' };
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, emptyMcpService(), upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, mail, configurationStub, mailWatcherStub, promptsStub));
		const routine = { id: 'email-routine', name: 'Status update', prompt: 'Write an update', serverIds: [], emailDelivery: { connectionId: mailbox.id, recipient: 'recipient@example.test' } };
		const lastDelivery = () => new AssistantConversationStore(storage).read('account').turns.at(-1)?.activity?.find(item => item.tool === 'assistant.routine_email');
		const approvedDelivery = async () => true;
		assert.strictEqual(await runs.sendScheduledRoutine('account', routine, 'complete', approvedDelivery), 'completed');
		assert.deepStrictEqual(sends, [{ to: routine.emailDelivery.recipient, subject: routine.name, body: 'Scheduled fixture message.' }]);
		assert.strictEqual(lastDelivery()?.status, 'done');
		assert.strictEqual(await runs.sendScheduledRoutine('account', { ...routine, serverIds: ['missing-source'] }, 'incomplete', approvedDelivery), 'needsAttention');
		assert.strictEqual(lastDelivery()?.status, 'rejected');
		assert.strictEqual(await runs.sendScheduledRoutine('account', routine, 'revoked-delivery', async () => false), 'needsAttention');
		assert.strictEqual(lastDelivery()?.status, 'rejected');
		connections.set([{ ...mailbox, id: 'another-mailbox' }], undefined);
		assert.strictEqual(await runs.sendScheduledRoutine('account', routine, 'missing-mailbox', approvedDelivery), 'needsAttention');
		assert.strictEqual(lastDelivery()?.status, 'failed');
		assert.strictEqual(sends.length, 1);
		connections.set([mailbox], undefined);
		deliveryFails = true;
		assert.strictEqual(await runs.sendScheduledRoutine('account', routine, 'unconfirmed', approvedDelivery), 'needsAttention');
		assert.strictEqual(sends.length, 2);
		assert.match(lastDelivery()?.detail ?? '', /Check Sent/);
		const sendsBefore = sends.length;
		const turnsBefore = new AssistantConversationStore(storage).read('account').turns.length;
		limited = true;
		assert.strictEqual(await runs.sendScheduledRoutine('account', routine, 'truncated', approvedDelivery), 'failed');
		assert.strictEqual(sends.length, sendsBefore);
		assert.strictEqual(new AssistantConversationStore(storage).read('account').turns.length, turnsBefore);
		assert.strictEqual(runs.getError('account'), 'outputLimit');
	});

	test('saves exact text within one rendered reply and retains raw Markdown for a whole reply', () => {
		const root = document.body.appendChild(document.createElement('div'));
		const reply = root.appendChild(document.createElement('p'));
		reply.classList.add('reach-assistant-reply');
		const first = reply.appendChild(document.createTextNode('First '));
		const emphasized = reply.appendChild(document.createElement('em')).appendChild(document.createTextNode('idea'));
		const outside = root.appendChild(document.createElement('p')).appendChild(document.createTextNode('Another reply'));
		const selection = document.getSelection()!;
		const markdown = 'First *idea*';
		try {
			selection.removeAllRanges();
			assert.strictEqual(assistantReplyPassage(reply, markdown), markdown);
			selection.setBaseAndExtent(first, 2, emphasized, 3);
			assert.strictEqual(assistantReplyPassage(reply, markdown), 'rst ide');
			selection.setBaseAndExtent(emphasized, 3, first, 2);
			assert.strictEqual(assistantReplyPassage(reply, markdown), 'rst ide');
			selection.setBaseAndExtent(first, 2, outside, 4);
			assert.strictEqual(assistantReplyPassage(reply, markdown), undefined);
			selection.setBaseAndExtent(outside, 1, outside, 4);
			assert.strictEqual(assistantReplyPassage(reply, markdown), markdown);
		} finally {
			selection.removeAllRanges();
			root.remove();
		}
	});

	test('opens linked tasks and announces tool approvals once in the shared composer', async () => {
		const opened: string[] = [];
		const descriptors: ICustomViewDescriptor[] = [];
		const active = observableValue<ICustomViewDescriptor | undefined>('assistantTestView', undefined);
		const views = new class extends mock<ICustomViewService>() {
			override readonly activeCustomView = active;
			override readonly activeCustomViewOpen = constObservable(undefined);
			override registerCustomView(value: ICustomViewDescriptor) { descriptors.push(value); return Disposable.None; }
			override showCustomView(id: string) {
				opened.push(id);
				active.set(descriptors.find(descriptor => descriptor.id === id), undefined);
			}
		};
		const instantiation = store.add(new TestInstantiationService());
		instantiation.stub(ICustomViewService, views);
		instantiation.stub(ISignalCDirectMail, directMailStub);
		const focusedSession = observableValue<ISession | undefined>('focusedAssistantProject', undefined);
		const host = upcastPartial<ISessionHost>({ session: focusedSession, chat: constObservable(undefined) });
		const editorSelection = observableValue<ReturnType<IAssistantAgentChats['editorSelection']['get']>>('testEditorSelection', undefined);
		const agentChats = upcastPartial<IAssistantAgentChats>({ target: undefined, editorSelection });
		const project = URI.parse('file:///linked-project');
		instantiation.stub(IFocusedSessionHostService, upcastPartial<IFocusedSessionHostService>({ getFocusedHost: () => host }));
		instantiation.stub(INewSessionComposerService, upcastPartial<INewSessionComposerService>({
			getComposer: () => upcastPartial<INewSessionComposer>({ workspaceSelection: { folderUri: project, state: 'selected', origin: WorkspaceSelectionOrigin.User, historyState: 'loaded', sessionFallbackState: 'idle', registeredProviderCount: 1 } }),
		}));
		instantiation.stub(IAssistantAgentChats, agentChats);
		const storage = store.add(new InMemoryStorageService());
		const accountChanged = store.add(new Emitter<void>());
		const changedInput = store.add(new Emitter<void>());
		const submitted: { accountId: string; request: ReachAssistantRequest; conversationId?: string }[] = [];
		const signals: AccessibilitySignal[] = [];
		let account: IReachAccount | undefined;
		let inputValue = '';
		let draftFiles: IChatDraft['attachments'] = [];
		const excerpts: { name: string; content: string }[] = [];
		const taskResult = new DeferredPromise<ReachAssistantResponse>();
		const routineAccounts: string[] = [];
		const pickers: string[] = [];
		const savedPickerHidden = store.add(new Emitter<IQuickInputHideEvent>());
		let savedPickerVisible = false;
		let pendingAccountCheck: DeferredPromise<void> | undefined;
		const releasedDictationDrafts: string[] = [];
		let clearConfirmed = false;
		let clearDialogs = 0;
		instantiation.stub(ICommandService, upcastPartial<ICommandService>({
			executeCommand: async (id, ...args) => {
				await CommandsRegistry.getCommand(id)!.handler(instantiation, ...args);
				return undefined;
			},
		}));
		instantiation.stub(IClipboardService, new class extends mock<IClipboardService>() { });
		instantiation.stub(IQuickInputService, upcastPartial<IQuickInputService>({
			pick: async (_items, options) => { pickers.push(options?.title ?? ''); return undefined; },
			createQuickPick: <T extends IQuickPickItem>() => upcastPartial<IQuickPick<T, { useSeparators: boolean }>>({
				onDidChangeValue: Event.None, onDidAccept: Event.None, onDidHide: savedPickerHidden.event,
				show: () => { savedPickerVisible = true; },
				hide: () => { savedPickerVisible = false; savedPickerHidden.fire({ reason: QuickInputHideReason.Other }); },
				dispose: () => { },
			}),
		}));
		instantiation.stub(IAssistantBriefingScheduler, upcastPartial<IAssistantBriefingScheduler>({
			getRoutines: async id => { routineAccounts.push(id); return []; },
		}));
		instantiation.stub(IAssistantMailWatcher, upcastPartial<IAssistantMailWatcher>({
			start: () => undefined, isEnabled: () => true, onDidChange: Event.None,
			status: () => ({ enabled: true, connections: [], alerts: [] }),
			checkNow: async () => 0, setEnabled: async () => undefined, markSeen: () => undefined,
		}));
		instantiation.stub(IDialogService, upcastPartial<IDialogService>({
			confirm: async options => {
				assert.strictEqual(options.custom, true);
				assert.strictEqual(options.message, 'Are you sure you want to clear this chat?');
				clearDialogs++;
				return { confirmed: clearConfirmed };
			},
		}));
		instantiation.stub(IRequestService, new class extends mock<IRequestService>() { });
		instantiation.stub(IFileService, fileServiceMock());
		instantiation.stub(IWorkspaceContextService, workspaceServiceMock());
		instantiation.stub(IWorkspaceTrustRequestService, upcastPartial<IWorkspaceTrustRequestService>({}));
		instantiation.stub(IGitService, upcastPartial<IGitService>({}));
		instantiation.stub(ISearchService, searchServiceMock());
		instantiation.stub(IAccessibilitySignalService, new class extends mock<IAccessibilitySignalService>() {
			override async playSignal(signal: AccessibilitySignal) { signals.push(signal); }
		});
		instantiation.stub(IOpenerService, openerStub);
		instantiation.stub(IStorageService, storage);
		const mcpService = mcpServiceWith(fakeServer({ name: 'Notes', toolCount: 3 }));
		instantiation.stub(IMcpService, mcpService);
		instantiation.stubInstance(SharedChatInputComposer, {
			dispose() { },
			render(parent, _root, _notices, inlineControls) { if (inlineControls) { parent.appendChild(inlineControls); } },
			focus() { },
			setInputEnabled() { },
			cancelDictation() { releasedDictationDrafts.push(inputValue); },
			getInputValue: () => inputValue,
			setInputValue: value => { inputValue = value; changedInput.fire(); },
			getDraft: () => ({ inputText: inputValue, attachments: draftFiles }),
			setDraft: draft => { inputValue = draft?.inputText ?? ''; draftFiles = draft?.attachments ?? []; changedInput.fire(); },
			attachTextContext: (name, content) => { excerpts.push({ name, content }); },
			onDidChangeInput: changedInput.event,
			attachments: [],
		});
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = accountChanged.event;
			override readonly assistantSupported = true;
			override async getAccount() { if (pendingAccountCheck) { await pendingAccountCheck.p; } return account; }
			override async listAssistantModels() { return [{ id: 'reach-qwen', name: 'Qwen' }]; }
			override async sendAssistantPrompt(accountId: string, request: ReachAssistantRequest, _requestId?: string, _token?: CancellationToken, conversationId?: string): Promise<ReachAssistantResponse> {
				submitted.push({ accountId, request, conversationId });
				if (request.messages.at(-1)?.content === 'Write a fixture') {
					return { kind: 'reply', text: '', toolCalls: [{ id: 'write-fixture', name: 'assistant.write_file', arguments: '{"path":"fixture.txt","content":"hello"}' }] };
				}
				if (conversationId) { return taskResult.p; }
				return { kind: 'reply', text: '**Connected.**' };
			}
		};
		instantiation.stub(IReachSubscriptionConnection, reach);
		const runs = store.add(new AssistantRuns(reach, storage, mcpService, new class extends mock<IRequestService>() { }, fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		instantiation.stub(IAssistantRuns, runs);
		store.add(new AssistantSessionContribution(views));
		for (const menu of [AgentsWindowMenus.CustomViewAssistant, AgentsWindowMenus.CustomViewAutomations]) {
			assert.ok(MenuRegistry.getMenuItems(menu).some(item => isIMenuItem(item) && item.command.id === MANAGE_ASSISTANT_ROUTINES_ACTION_ID));
		}
		const headerActions = MenuRegistry.getMenuItems(AgentsWindowMenus.CustomViewAssistant).filter(isIMenuItem);
		assert.strictEqual(headerActions.find(item => item.command.id === CLEAR_ASSISTANT_CHAT_ACTION_ID)?.order, 2);
		const clearChat = CommandsRegistry.getCommand(CLEAR_ASSISTANT_CHAT_ACTION_ID)!;
		const manageRoutines = CommandsRegistry.getCommand(MANAGE_ASSISTANT_ROUTINES_ACTION_ID)!;
		await manageRoutines.handler(instantiation);
		assert.deepStrictEqual(opened, [ASSISTANT_CUSTOM_VIEW_ID]);
		assert.deepStrictEqual(agentChats.target?.workspace, project);
		const [descriptor] = descriptors;
		assert.strictEqual(descriptor.id, ASSISTANT_CUSTOM_VIEW_ID);
		const view = store.add(instantiation.createInstance<AbstractCustomView>(descriptor.ctor));
		const content = document.createElement('div');
		const created = spy(instantiation, 'createInstance');
		view.render(content);
		await timeout(0);
		assert.match(content.textContent ?? '', /Sign in with your SignalREACH account/);
		assert.deepStrictEqual(routineAccounts, [], 'Opening routines must wait for a signed-in account');
		const chat = content.querySelector<HTMLElement>('.reach-assistant-chat')!;
		assert.strictEqual(chat.hidden, true);
		account = { id: 'my-account', subscriptionActive: true };
		const conversation = new AssistantConversationStore(storage);
		conversation.appendTurn(account.id, { prompt: 'Saved locally', reply: 'Earlier reply' });
		conversation.setDraft(account.id, 'Saved draft');
		conversation.setModel(account.id, 'reach-qwen');
		accountChanged.fire();
		await timeout(0);
		const composerOptions = created.getCalls().find(call => call.args[0] === SharedChatInputComposer)?.args.at(1) as ConstructorParameters<typeof SharedChatInputComposer>[0];
		created.restore();
		assert.strictEqual(chat.hidden, false);
		assert.match(content.querySelector('[role=log]')?.textContent ?? '', /Earlier reply/);
		assert.strictEqual(chat.querySelector('.reach-assistant-app-name')?.textContent, 'Notes');
		assert.strictEqual(chat.querySelector('.reach-assistant-app-count')?.textContent, '3');
		assert.ok(chat.querySelector('.reach-assistant-app-dot.running'));
		assert.ok(chat.querySelector('.reach-assistant-app-add'));
		assert.strictEqual(inputValue, 'Saved draft');
		assert.deepStrictEqual(routineAccounts, ['my-account']);
		assert.deepStrictEqual(pickers, ['Assistant routines']);
		assert.ok(chat.querySelector('.reach-assistant-composer .reach-assistant-controls'));
		assert.match(chat.querySelector('.reach-assistant-model')?.textContent ?? '', /Qwen/);
		const button = (label: string) => Array.from(chat.querySelectorAll<HTMLButtonElement>('button')).find(element => element.textContent === label)!;
		assert.ok(button('Saved ideas'));
		const originalFile = toFileVariableEntry(URI.joinPath(project, 'draft.txt'));
		draftFiles = [originalFile];
		const exactPassage = '  Keep the spacing.\nSecond line.  ';
		inputValue = '/save ' + exactPassage;
		assert.strictEqual(await composerOptions.submitOverride!(), true);
		assert.strictEqual(conversation.listSavedIdeas(account.id)[0].text, exactPassage);
		assert.deepStrictEqual(conversation.readComposerDraft(account.id).attachments.map(item => item.id), [originalFile.id]);
		inputValue = '/save ' + exactPassage;
		assert.strictEqual(await composerOptions.submitOverride!(), true);
		assert.strictEqual(conversation.listSavedIdeas(account.id).length, 1);
		inputValue = '/save ' + 'x'.repeat(8001);
		assert.strictEqual(await composerOptions.submitOverride!(), false);
		assert.strictEqual(inputValue.length, 8007);
		assert.strictEqual(submitted.length, 0);
		draftFiles = [];
		inputValue = 'Saved draft';
		changedInput.fire();
		const selection = { text: 'Unsaved excerpt', startLine: 2, endLine: 4, truncated: false };
		editorSelection.set({ file: URI.file('/elsewhere/selection.ts'), selection }, undefined);
		assert.strictEqual(button('Attach IDE selection · selection.ts:2–4'), undefined);
		editorSelection.set({ file: URI.file('/ws/src/selection.ts'), selection }, undefined);
		assert.strictEqual(button('Attach IDE selection · selection.ts:2–4'), undefined, 'The selected project scopes attachments');
		editorSelection.set({ file: URI.joinPath(project, 'src/selection.ts'), selection }, undefined);
		assert.deepStrictEqual(excerpts, []);
		button('Attach IDE selection · selection.ts:2–4').click();
		assert.deepStrictEqual(excerpts, [{ name: 'src/selection.ts:2–4', content: selection.text }]);
		assert.strictEqual(inputValue, 'Saved draft');
		assert.strictEqual(submitted.length, 0);
		editorSelection.set(undefined, undefined);
		button('New task').click();
		const task = conversation.read(account.id).tasks![0];
		assert.strictEqual(releasedDictationDrafts.at(-1), 'Saved draft');
		assert.strictEqual(inputValue, '');
		assert.strictEqual(conversation.read(account.id, task.id).modelId, 'reach-qwen');
		assert.strictEqual(readAssistantTaskTarget(task)?.workspace?.toString(), project.toString());
		inputValue = 'Unsent on exit';
		changedInput.fire();
		assert.strictEqual(conversation.read(account.id, task.id).draft, undefined);
		await storage.flush(WillSaveStateReason.SHUTDOWN);
		assert.strictEqual(conversation.read(account.id, task.id).draft, 'Unsent on exit');
		const pending = runs.send(account.id, 'Research an approach', task.id);
		assert.strictEqual(chat.querySelector<HTMLElement>('.reach-assistant-pending')?.hidden, false);
		await clearChat.handler(instantiation);
		assert.strictEqual(clearDialogs, 0);
		await manageRoutines.handler(instantiation);
		await timeout(0);
		assert.strictEqual(inputValue, 'Saved draft');
		assert.strictEqual(conversation.read(account.id, task.id).draft, 'Unsent on exit');
		assert.deepStrictEqual(pickers, ['Assistant routines', 'Assistant routines']);
		assert.deepStrictEqual(opened, [ASSISTANT_CUSTOM_VIEW_ID], 'Opening the picker must reuse the visible Assistant view');
		assert.match(chat.querySelector('.reach-assistant-task-updates')?.textContent ?? '', /Research an approach.*Running/);
		assert.strictEqual(submitted.length, 1);
		assert.strictEqual(submitted[0].accountId, 'my-account');
		assert.strictEqual(submitted[0].conversationId, task.id);
		assert.strictEqual(submitted[0].request.modelId, 'reach-qwen');
		assert.deepStrictEqual(submitted[0].request.messages.at(-1), { role: 'user', content: 'Research an approach' });
		await taskResult.complete({ kind: 'reply', text: '**Findings ready.**', model: 'Qwen' });
		assert.strictEqual(await pending, true);
		assert.match(chat.querySelector('.reach-assistant-task-updates')?.textContent ?? '', /Complete.*Findings ready/);
		assert.ok(signals.includes(AccessibilitySignal.taskCompleted));
		assert.strictEqual(conversation.read(account.id).turns.length, 1);
		chat.querySelector<HTMLElement>('.reach-assistant-reply-tools .codicon-refresh')!.click();
		await timeout(0);
		assert.strictEqual(conversation.read(account.id).turns.length, 1);
		assert.strictEqual(chat.querySelector('.reach-assistant-reply strong')?.textContent, 'Connected.');
		assert.strictEqual(inputValue, 'Saved draft');
		button('Research an approach').click();
		assert.strictEqual(chat.querySelector('.reach-assistant-reply strong')?.textContent, 'Findings ready.');
		assert.strictEqual(chat.querySelector('.reach-assistant-reply-tools .action-label')?.getAttribute('role'), 'button');
		const approval = Event.toPromise(runs.onDidRequestApproval);
		const finishing = runs.send(account.id, 'Write a fixture', task.id);
		await approval;
		assert.strictEqual(chat.querySelector<HTMLElement>('.reach-assistant-approval')?.hidden, false);
		button('Reject').click();
		assert.strictEqual(await finishing, true);
		assert.strictEqual(signals.filter(signal => signal === AccessibilitySignal.chatUserActionRequired).length, 1);
		await clearChat.handler(instantiation);
		assert.strictEqual(clearDialogs, 1);
		assert.strictEqual(conversation.read(account.id, task.id).turns.length, 2);
		clearConfirmed = true;
		await clearChat.handler(instantiation);
		assert.strictEqual(clearDialogs, 2);
		assert.strictEqual(conversation.read(account.id, task.id).turns.length, 0);
		assert.strictEqual(conversation.read(account.id, task.id).draft, 'Unsent on exit');
		assert.strictEqual(conversation.read(account.id).turns.length, 1);
		assert.doesNotMatch(chat.querySelector('[role=log]')?.textContent ?? '', /Findings ready/);
		draftFiles = [toFileVariableEntry(URI.joinPath(project, 'unsent-notes.txt'))];
		conversation.appendTurn(account.id, { prompt: 'Edit the example', reply: 'Edited.', activity: [{ tool: 'assistant.edit_file', label: 'Edit', status: 'done' }] }, task.id);
		button('Summarize the changes').click();
		await timeout(0);
		assert.strictEqual(submitted.at(-1)?.request.messages.at(-1)?.content, 'Summarize the changes you just made to my files.');
		assert.strictEqual(inputValue, 'Unsent on exit');
		assert.deepStrictEqual(conversation.readComposerDraft(account.id, task.id).attachments.map(file => file.id), draftFiles.map(file => file.id));
		draftFiles = [];
		button('Saved ideas').click();
		assert.strictEqual(savedPickerVisible, true);
		pendingAccountCheck = new DeferredPromise<void>();
		account = undefined;
		accountChanged.fire();
		assert.strictEqual(savedPickerVisible, false, 'Hide saved passages immediately when the signed-in account needs to be verified again');
		button('Saved ideas').click();
		assert.strictEqual(savedPickerVisible, false, 'Wait for account verification before opening another collection');
		await pendingAccountCheck.complete();
		pendingAccountCheck = undefined;
		await timeout(0);
		assert.strictEqual(releasedDictationDrafts.at(-1), 'Unsent on exit');
		assert.strictEqual(chat.hidden, true);
		assert.strictEqual(content.querySelector('[role=log]')?.textContent, '');
		assert.strictEqual(conversation.read('my-account').draft, 'Saved draft');
		await manageRoutines.handler(instantiation);
		active.set(undefined, undefined);
		account = { id: 'my-account', subscriptionActive: true };
		accountChanged.fire();
		await timeout(0);
		assert.strictEqual(pickers.length, 2, 'Leaving Assistant must cancel a picker request waiting for sign-in');
		const fileRoots = [project, URI.parse('file:///linked-utilities')];
		focusedSession.set(upcastPartial<ISession>({
			mainChat: constObservable(undefined as never),
			workspace: constObservable(upcastPartial<ISessionWorkspace>({
				uri: project, folders: fileRoots.map(workingDirectory => upcastPartial<ISessionFolder>({ workingDirectory })), isVirtualWorkspace: false,
			})),
		}), undefined);
		await CommandsRegistry.getCommand(OPEN_ASSISTANT_ACTION_ID)!.handler(instantiation);
		assert.deepStrictEqual(agentChats.target?.fileRoots, fileRoots);
		const breadcrumb = chat.querySelector<HTMLElement>('.reach-assistant-breadcrumb')!;
		if (!breadcrumb.hidden) { breadcrumb.querySelector('button')!.click(); }
		const source = toPasteVariableEntry('Skill · meeting-follow-up', 'Record agreed owners.', { id: 'assistant-skill:fixture' });
		const taskText = 'Summarize this excerpt: Mira will send the draft Friday.';
		const prefix = `<attachment name="${source.name}">\n${source.value}\n</attachment>\n\n`;
		inputValue = '/task ' + taskText;
		draftFiles = [source];
		const parentTurns = conversation.read(account.id).turns.length;
		assert.strictEqual(await composerOptions.submitOverride!(), true);
		await timeout(0);
		const linked = conversation.read(account.id).tasks!.at(-1)!;
		const [childTurn] = conversation.read(account.id, linked.id).turns;
		assert.deepStrictEqual([submitted.at(-1)?.conversationId, linked.title, childTurn.prompt, childTurn.requestStart], [linked.id, taskText, prefix + taskText, prefix.length]);
		assert.strictEqual(conversation.read(account.id).turns.length, parentTurns);
		inputValue = '/task Nested task';
		assert.strictEqual(await composerOptions.submitOverride!(), false);
		assert.match(chat.querySelector('.reach-assistant-feedback')!.textContent ?? '', /Return to Assistant/);
	});

	test('regenerates from the selected turn only after a successful reply and protects concurrent changes', async () => {
		const storage = store.add(new InMemoryStorageService());
		const conversation = new AssistantConversationStore(storage);
		for (const prompt of ['Earlier', 'Selected', 'Later']) { conversation.appendTurn('account', { prompt, reply: prompt + ' reply' }); }
		const original = conversation.read('account').turns;
		let response = new DeferredPromise<ReachAssistantResponse>();
		const requests: ReachAssistantRequest[] = [];
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_account: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				requests.push(request);
				return response.p;
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, emptyMcpService(), upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		const failed = runs.send('account', 'Selected', undefined, { replaceFrom: 1 });
		assert.deepStrictEqual(conversation.read('account').turns, original);
		await response.complete({ kind: 'unavailable' });
		assert.strictEqual(await failed, false);
		assert.deepStrictEqual(conversation.read('account').turns, original);
		assert.deepStrictEqual(requests[0].messages.slice(1).map(message => message.content), ['Earlier', 'Earlier reply', 'Selected']);

		response = new DeferredPromise();
		const cancelled = runs.send('account', 'Selected');
		runs.stop('account');
		await response.complete({ kind: 'reply', text: 'Late answer' });
		assert.strictEqual(await cancelled, false);
		assert.deepStrictEqual(conversation.read('account').turns, original);

		response = new DeferredPromise();
		const retry = runs.send('account', 'Selected');
		await response.complete({ kind: 'reply', text: 'Replacement' });
		assert.strictEqual(await retry, true);
		assert.deepStrictEqual(conversation.read('account').turns.map(turn => turn.reply), ['Earlier reply', 'Replacement']);
		response = new DeferredPromise();
		const stale = runs.send('account', 'Selected', undefined, { replaceFrom: 1 });
		conversation.appendTurn('account', { prompt: 'From another window', reply: 'Keep me' });
		await response.complete({ kind: 'reply', text: 'Stale answer' });
		assert.strictEqual(await stale, false);
		assert.strictEqual(runs.getError('account'), 'historyChanged');
		assert.strictEqual(runs.getFailedPrompt('account'), undefined);
		assert.deepStrictEqual(conversation.read('account').turns.map(turn => turn.reply), ['Earlier reply', 'Replacement', 'Keep me']);
	});

	test('compacts twice without discarding searchable exchanges and pins the latest summary in model context', async () => {
		const storage = store.add(new InMemoryStorageService());
		const conversation = new AssistantConversationStore(storage);
		const recorded = new Date(2024, 2, 4, 12).getTime();
		for (let i = 0; i < 16; i++) { conversation.appendTurn('account', { prompt: `Original exchange ${i}`, reply: `Recorded value ${i}`, at: recorded }); }
		const requests: ReachAssistantRequest[] = [];
		let reply = 'First summary';
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_account: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				requests.push({ ...request, messages: [...request.messages] });
				if (request.tools?.length === 1 && request.tools[0].name === 'assistant.search_conversation' && request.messages.at(-1)?.role === 'user') {
					return { kind: 'reply', text: '', toolCalls: [{ id: 'recall', name: 'assistant.search_conversation', arguments: '{"query":"Original exchange 0","from_date":"2024-03-04","to_date":"2024-03-04"}' }] };
				}
				return { kind: 'reply', text: reply };
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, emptyMcpService(), upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		const label = 'Summarize the conversation';
		assert.ok(await runs.send('account', label, undefined, { compact: true }));
		assert.match(requests[0].messages[0].content, /12 of 16 retained turns/);
		assert.ok(!requests[0].messages.some(message => message.content === 'Original exchange 0'));
		assert.match(requests[1].messages.at(-1)?.content ?? '', /Recorded value 0/);
		const reopened = new AssistantConversationStore(storage);
		assert.strictEqual(assistantSummaryIndex(reopened.read('account').turns), 16);
		for (let i = 0; i < 15; i++) { conversation.appendTurn('account', { prompt: `Later exchange ${i}`, reply: `Later answer ${i}` }); }
		reply = 'Second summary';
		assert.ok(await runs.send('account', label, undefined, { compact: true }));
		assert.ok(requests[2].messages.some(message => message.content === 'First summary'));
		assert.ok(!requests[2].messages.some(message => message.content === 'Later exchange 0'));
		assert.match(requests[3].messages.at(-1)?.content ?? '', /Recorded value 0/);
		assert.strictEqual(await runs.send('account', label, undefined, { compact: true }), false);
		reply = 'Follow-up answer';
		assert.ok(await runs.send('account', 'What comes next?'));
		assert.deepStrictEqual(requests[4].messages.slice(1).map(message => message.content), [label, 'Second summary', 'What comes next?']);
		assert.strictEqual(reopened.read('account').turns.length, 34);
		assert.strictEqual(reopened.read('account').turns.filter(turn => turn.prompt === 'Original exchange 0').length, 1);
		conversation.popTurn('account');
		reply = 'Regenerated summary';
		assert.ok(await runs.send('account', label, undefined, { replaceFrom: 32 }));
		assert.strictEqual(conversation.popTurn('account')?.compacted, true);
		assert.strictEqual(assistantSummaryIndex(conversation.read('account').turns), 16);
		assert.ok(await runs.send('account', 'Edited original', undefined, { replaceFrom: 1 }));
		assert.deepStrictEqual(conversation.read('account').turns.map(turn => turn.prompt), ['Original exchange 0', 'Edited original']);
		conversation.clearTurns('account');
		assert.deepStrictEqual(reopened.read('account').turns, []);
	});

	test('rejects a stale compaction without changing a concurrent exchange or draft', async () => {
		const storage = store.add(new InMemoryStorageService());
		const conversation = new AssistantConversationStore(storage);
		conversation.appendTurn('account', { prompt: 'First', reply: 'One' });
		conversation.appendTurn('account', { prompt: 'Second', reply: 'Two' });
		const waiting = new DeferredPromise<ReachAssistantResponse>();
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(): Promise<ReachAssistantResponse> { return waiting.p; }
		};
		const runs = store.add(new AssistantRuns(reach, storage, emptyMcpService(), upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		const pending = runs.send('account', 'Summarize the conversation', undefined, { compact: true });
		conversation.appendTurn('account', { prompt: 'Concurrent', reply: 'Three' });
		conversation.setDraft('account', 'Work in progress');
		await waiting.complete({ kind: 'reply', text: 'Old snapshot' });
		assert.strictEqual(await pending, false);
		assert.deepStrictEqual(conversation.read('account').turns.map(turn => turn.reply), ['One', 'Two', 'Three']);
		assert.strictEqual(conversation.read('account').draft, 'Work in progress');
	});

	test('rebases inline attachment references when removing the typed task command', () => {
		const taskText = 'Summarize these notes: Mira will send the draft Friday.';
		const attachment = toPasteVariableEntry('Skill · meeting-follow-up', 'Record agreed owners.', { id: 'assistant-skill:fixture' });
		const command = '  /task   ';
		const taskDraft = getAssistantLinkedTaskDraft({
			inputText: command + taskText,
			attachments: [{ ...attachment, range: { start: command.length, endExclusive: command.length + 9 } }],
			references: [{ id: 'fixture', range: { startLineNumber: 1, startColumn: command.length + 1, endLineNumber: 1, endColumn: command.length + 10 }, data: 'fixture' }],
		})!;
		assert.strictEqual(taskDraft.inputText, taskText);
		assert.deepStrictEqual(taskDraft.attachments[0].range, { start: 0, endExclusive: 9 });
		assert.strictEqual(taskDraft.references?.[0].range.startColumn, 1);
		assert.strictEqual(getAssistantLinkedTaskDraft({ inputText: 'A quote about /task Work', attachments: [attachment] }), undefined);
	});

	test('restores each conversation’s editable attachments and keeps its prepared prompt for retry', () => {
		const storage = store.add(new InMemoryStorageService());
		const conversation = new AssistantConversationStore(storage);
		const originalAttachment = URI.file('/fixture/service-marker.txt');
		const draft = { inputText: 'Read this file', attachments: [toFileVariableEntry(originalAttachment)] };
		conversation.setDraft('account', 'Unsent main draft');
		conversation.setComposerDraft('account', draft, 'task', 'Prepared prompt with file content');

		const reopened = new AssistantConversationStore(storage);
		assert.deepStrictEqual(reopened.readComposerDraft('account'), { inputText: 'Unsent main draft', attachments: [] });
		const restored = reopened.readComposerDraft('account', 'task');
		assert.strictEqual(restored.inputText, draft.inputText);
		assert.deepStrictEqual(restored.attachments.map(attachment => attachment.value), [originalAttachment]);
		reopened.setComposerDraft('account', restored, 'task');
		assert.strictEqual(reopened.read('account', 'task').draft, 'Prepared prompt with file content');

		const replacementAttachment = URI.file('/fixture/another-file.txt');
		reopened.setComposerDraft('account', { ...restored, attachments: [toFileVariableEntry(replacementAttachment)] }, 'task');
		assert.strictEqual(reopened.read('account', 'task').draft, draft.inputText);
		assert.deepStrictEqual(reopened.readComposerDraft('account', 'task').attachments.map(attachment => attachment.value), [replacementAttachment]);
		assert.strictEqual(reopened.readComposerDraft('other-account', 'task').attachments.length, 0);
		reopened.setDraft('account', '', 'task');
		assert.deepStrictEqual(reopened.readComposerDraft('account', 'task'), { inputText: '', attachments: [] });
	});

	test('keeps a replacement attachment draft when an earlier request with the same text finishes', async () => {
		const storage = store.add(new InMemoryStorageService());
		const conversation = new AssistantConversationStore(storage);
		const reply = new DeferredPromise<ReachAssistantResponse>();
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt() { return reply.p; }
		};
		const runs = store.add(new AssistantRuns(reach, storage, emptyMcpService(), new class extends mock<IRequestService>() { }, fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		const text = 'Read this file';
		conversation.setComposerDraft('account', { inputText: text, attachments: [toFileVariableEntry(URI.file('/fixture/first.txt'))] });
		const sending = runs.send('account', text);
		const replacement = URI.file('/fixture/second.txt');
		conversation.setComposerDraft('account', { inputText: text, attachments: [toFileVariableEntry(replacement)] });
		await reply.complete({ kind: 'reply', text: 'Read the first file.' });
		assert.strictEqual(await sending, true);
		assert.deepStrictEqual(conversation.readComposerDraft('account').attachments.map(attachment => attachment.value), [replacement]);
	});

	test('captures IDE roots only while the current New Chat selection belongs to that switch', async () => {
		const instantiation = store.add(new TestInstantiationService());
		const views = upcastPartial<ICustomViewService>({
			activeCustomView: constObservable(undefined),
			registerCustomView: () => Disposable.None,
			showCustomView: () => { },
		});
		const root = URI.file('/ide-active');
		const sibling = URI.file('/ide-sibling');
		const activeFile = URI.joinPath(root, 'main.ts');
		const selectionVersion = observableValue('workspace selection', 1);
		const navigationVersion = observableValue('workspace navigation', 2);
		const composer = upcastPartial<INewSessionComposer>({
			workspaceSelection: { folderUri: root, state: 'selected', origin: WorkspaceSelectionOrigin.WindowContext, historyState: 'loaded', sessionFallbackState: 'idle', registeredProviderCount: 1 },
		});
		const target = upcastPartial<IAssistantAgentChats>({
			editorWorkspace: { selectedFolder: root, fileRoots: [root, sibling], fileRootNames: ['Active', 'Sibling'], activeFile, selectionVersion: 1, navigationVersion: 2 },
		});
		const host = upcastPartial<ISessionHost>({ session: constObservable(undefined), chat: constObservable(undefined) });
		instantiation.stub(IFocusedSessionHostService, upcastPartial<IFocusedSessionHostService>({ getFocusedHost: () => host }));
		instantiation.stub(INewSessionComposerService, upcastPartial<INewSessionComposerService>({
			getComposer: () => composer, userWorkspaceSelectionVersion: selectionVersion, userNavigationVersion: navigationVersion,
		}));
		instantiation.stub(IAssistantAgentChats, target);
		instantiation.stub(IAssistantMailWatcher, upcastPartial<IAssistantMailWatcher>({
			start: () => undefined, isEnabled: () => true, onDidChange: Event.None,
			status: () => ({ enabled: true, connections: [], alerts: [] }),
			checkNow: async () => 0, setEnabled: async () => undefined, markSeen: () => undefined,
		}));
		store.add(new AssistantSessionContribution(views));
		const open = () => CommandsRegistry.getCommand(OPEN_ASSISTANT_ACTION_ID)!.handler(instantiation);
		await open();
		assert.deepStrictEqual(target.target?.fileRoots, [root, sibling]);
		assert.deepStrictEqual(target.target?.fileRootNames, ['Active', 'Sibling']);
		assert.deepStrictEqual(target.target?.activeEditorFile, activeFile);
		assert.deepStrictEqual(target.target?.workspace, root);

		navigationVersion.set(3, undefined);
		await open();
		assert.strictEqual(target.target?.fileRoots, undefined);
		assert.strictEqual(target.target?.activeEditorFile, undefined);
		navigationVersion.set(2, undefined);
		selectionVersion.set(2, undefined);
		await open();
		assert.strictEqual(target.target?.fileRoots, undefined);
	});

	test('runs text-protocol tool calls, feeds results back, and records activity', async () => {
		const storage = store.add(new InMemoryStorageService());
		const sent: ReachAssistantRequest[] = [];
		const responses: ReachAssistantResponse[] = [
			{ kind: 'reply', text: 'Let me check.\n```tool_call\n{"name": "assistant.get_current_time", "arguments": {}}\n```' },
			{ kind: 'reply', text: 'It is noon, give or take.', model: 'Qwen' },
		];
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_accountId: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				sent.push(request);
				return responses.shift()!;
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, emptyMcpService(), new class extends mock<IRequestService>() { }, fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		assert.strictEqual(await runs.send('account', 'What time is it?'), true);
		assert.strictEqual(sent.length, 2);
		assert.strictEqual(sent[1].messages.at(-1)?.role, 'user');
		assert.match(String(sent[1].messages.at(-1)?.content), /<tool_result name="assistant\.get_current_time">/);
		const turns = new AssistantConversationStore(storage).read('account').turns;
		assert.strictEqual(turns.length, 1);
		assert.strictEqual(turns[0].reply, 'It is noon, give or take.');
		assert.deepStrictEqual(turns[0].activity?.map(item => [item.tool, item.status]), [['assistant.get_current_time', 'done']]);
	});

	test('returns every text tool result within the native user-message limit', async () => {
		const paths = ['one.txt', 'two.txt', 'short.txt'];
		let result = '';
		const reach = upcastPartial<IReachSubscriptionConnection>({
			onDidChangeAccount: Event.None,
			sendAssistantPrompt: async (_accountId, request) => {
				const last = request.messages.at(-1)!;
				if (last.content === 'Compare the files') {
					return { kind: 'reply', text: paths.map(path => '```tool_call\n' + JSON.stringify({ name: 'assistant.read_file', arguments: { path } }) + '\n```').join('\n') };
				}
				result = last.content;
				return result.length <= REACH_ASSISTANT_PROMPT_LIMIT ? { kind: 'reply', text: 'Compared.' } : { kind: 'unavailable' };
			},
		});
		const files = fileServiceMock({ '/ws/one.txt': 'a'.repeat(7000), '/ws/two.txt': 'b'.repeat(7000), '/ws/short.txt': 'Keep this whole answer.' });
		const runs = store.add(new AssistantRuns(reach, store.add(new InMemoryStorageService()), emptyMcpService(), upcastPartial<IRequestService>({}), files, workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		assert.strictEqual(await runs.send('account', 'Compare the files'), true);
		assert.match(result, /a{3000,}…\n<\/tool_result>/);
		assert.match(result, /b{3000,}…\n<\/tool_result>/);
		assert.ok(result.endsWith('Keep this whole answer.\n</tool_result>'));
	});

	test('answers every native call ID when a model requests more tools than a step can execute', async () => {
		const storage = store.add(new InMemoryStorageService());
		const calls = Array.from({ length: 5 }, (_, index) => ({ id: `read-${index}`, name: 'assistant.read_file', arguments: JSON.stringify({ path: `file-${index}` }) }));
		let replies: ReachAssistantRequest['messages'] = [];
		const reach = upcastPartial<IReachSubscriptionConnection>({
			onDidChangeAccount: Event.None,
			sendAssistantPrompt: async (_accountId, request) => {
				if (request.messages.at(-1)?.role === 'user') { return { kind: 'reply', text: '', toolCalls: calls }; }
				replies = request.messages.filter(message => message.role === 'tool');
				return { kind: 'reply', text: 'Reviewed.' };
			},
		});
		const runs = store.add(new AssistantRuns(reach, storage, emptyMcpService(), upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		assert.ok(await runs.send('account', 'Check these files'));
		assert.deepStrictEqual(replies.map(reply => reply.toolCallId), calls.map(call => call.id));
		assert.match(replies[4].content, /Skipped/);
		assert.strictEqual(new AssistantConversationStore(storage).read('account').turns[0].activity?.length, 4);
	});

	test('asks the user a clarifying question and feeds the answer back', async () => {
		const storage = store.add(new InMemoryStorageService());
		const sent: ReachAssistantRequest[] = [];
		const responses: ReachAssistantResponse[] = [
			{ kind: 'reply', text: 'One moment.\n```tool_call\n{"name": "assistant.ask_user", "arguments": {"question": "Which color?"}}\n```' },
			{ kind: 'reply', text: 'Blue it is.' },
		];
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_accountId: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				sent.push(request);
				return responses.shift()!;
			}
		};
		const questions: string[] = [];
		const quickInput = new class extends mock<IQuickInputService>() {
			override async input(options?: { title?: string }) {
				questions.push(options?.title ?? '');
				return 'Blue';
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, emptyMcpService(), new class extends mock<IRequestService>() { }, fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInput, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		assert.strictEqual(await runs.send('account', 'Pick a color.'), true);
		assert.deepStrictEqual(questions, ['Which color?']);
		assert.strictEqual(sent.length, 2);
		assert.match(String(sent[1].messages.at(-1)?.content), /The user answered: Blue/);
		const turns = new AssistantConversationStore(storage).read('account').turns;
		assert.strictEqual(turns[0].reply, 'Blue it is.');
		assert.deepStrictEqual(turns[0].activity?.map(item => [item.tool, item.status]), [['assistant.ask_user', 'done']]);
	});

	test('asks before write actions, runs them on Allow, and learns native tool_calls', async () => {
		const storage = store.add(new InMemoryStorageService());
		const sent: ReachAssistantRequest[] = [];
		const called: Record<string, unknown>[] = [];
		const writeTool = new class extends mock<IMcpTool>() {
			override readonly id = 'mcp_notes_write';
			override readonly visibility = McpToolVisibility.Model;
			override readonly definition = { name: 'mcp_notes_write', title: 'Write note', description: 'Writes a note', annotations: { destructiveHint: true }, inputSchema: { type: 'object' } } as never;
			override async call(params: Record<string, unknown>) {
				called.push(params);
				return { content: [{ type: 'text', text: 'written' }] } as never;
			}
		};
		const server = new class extends mock<IMcpServer>() {
			override readonly definition = { label: 'Notes' } as never;
			override readonly connectionState = constObservable<McpConnectionState>({ state: McpConnectionState.Kind.Running });
			override readonly enablement = constObservable(ContributionEnablementState.EnabledWorkspace);
			override readonly serverMetadata = constObservable({ serverName: 'Notes', icons: undefined as never });
			override readonly tools = constObservable<IMcpTool[]>([writeTool]);
			override readonly prompts = constObservable([]);
		};
		const mcpService = new class extends mock<IMcpService>() {
			override readonly servers = constObservable<IMcpServer[]>([server]);
		};
		const approvals: IAssistantApprovalRequest[] = [];
		const responses: ReachAssistantResponse[] = [
			{ kind: 'reply', text: '', toolCalls: [{ id: 'call-1', name: 'mcp_notes_write', arguments: '{"note":"hi"}' }] },
			{ kind: 'reply', text: 'Note written.' },
		];
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_accountId: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				sent.push(request);
				return responses.shift()!;
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, mcpService, new class extends mock<IRequestService>() { }, fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		store.add(runs.onDidRequestApproval(request => {
			approvals.push(request);
			assert.strictEqual(runs.getPendingApproval('account')?.id, request.id);
			runs.respondApproval(request.id, 'allow');
			assert.strictEqual(runs.getPendingApproval('account'), undefined);
		}));
		assert.strictEqual(await runs.send('account', 'Write a note'), true);
		assert.strictEqual(approvals.length, 1);
		assert.strictEqual(approvals[0].tool, 'mcp_notes_write');
		assert.deepStrictEqual(called, [{ note: 'hi' }]);
		assert.strictEqual(sent.length, 2);
		const echoed = sent[1].messages.at(-2);
		assert.strictEqual(echoed?.role, 'assistant');
		assert.deepStrictEqual(echoed?.toolCalls, [{ id: 'call-1', name: 'mcp_notes_write', arguments: '{"note":"hi"}' }]);
		assert.deepStrictEqual(sent[1].messages.at(-1), { role: 'tool', toolCallId: 'call-1', name: 'mcp_notes_write', content: 'written' });
		const activity = new AssistantConversationStore(storage).read('account').turns[0].activity;
		assert.deepStrictEqual(activity?.map(item => item.status), ['done']);
	});

	test('posts a live plan checklist while a run executes', async () => {
		const storage = store.add(new InMemoryStorageService());
		const sent: ReachAssistantRequest[] = [];
		const release = new DeferredPromise<ReachAssistantResponse>();
		const responses: ReachAssistantResponse[] = [
			{ kind: 'reply', text: 'Planning.\n```tool_call\n{"name": "assistant.update_plan", "arguments": {"items": [{"step": "Scan files", "status": "inProgress"}, {"step": "Write report", "status": "pending"}]}}\n```' },
		];
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_accountId: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				sent.push(request);
				return responses.length ? responses.shift()! : release.p;
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, emptyMcpService(), new class extends mock<IRequestService>() { }, fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		const pending = runs.send('account', 'Do the work');
		await timeout(10);
		assert.deepStrictEqual(runs.getPendingDetail('account')?.plan, [
			{ step: 'Scan files', status: 'inProgress' },
			{ step: 'Write report', status: 'pending' },
		]);
		release.complete({ kind: 'reply', text: 'All done.' });
		assert.strictEqual(await pending, true);
		assert.match(String(sent.at(-1)?.messages.at(-1)?.content), /Plan updated: 0\/2 complete/);
	});

	test('carries a rejection note back to the model', async () => {
		const storage = store.add(new InMemoryStorageService());
		const sent: ReachAssistantRequest[] = [];
		const writeTool = new class extends mock<IMcpTool>() {
			override readonly id = 'mcp_notes_write';
			override readonly visibility = McpToolVisibility.Model;
			override readonly definition = { name: 'mcp_notes_write', title: 'Write note', description: 'Writes a note', annotations: { destructiveHint: true }, inputSchema: { type: 'object' } } as never;
		};
		const server = new class extends mock<IMcpServer>() {
			override readonly definition = { label: 'Notes' } as never;
			override readonly connectionState = constObservable<McpConnectionState>({ state: McpConnectionState.Kind.Running });
			override readonly enablement = constObservable(ContributionEnablementState.EnabledWorkspace);
			override readonly serverMetadata = constObservable({ serverName: 'Notes', icons: undefined as never });
			override readonly tools = constObservable<IMcpTool[]>([writeTool]);
			override readonly prompts = constObservable([]);
		};
		const responses: ReachAssistantResponse[] = [
			{ kind: 'reply', text: '', toolCalls: [{ id: 'call-1', name: 'mcp_notes_write', arguments: '{"note":"hi"}' }] },
			{ kind: 'reply', text: 'Understood, skipped it.' },
		];
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_accountId: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				sent.push(request);
				return responses.shift()!;
			}
		};
		const mcpService = new class extends mock<IMcpService>() {
			override readonly servers = constObservable<IMcpServer[]>([server]);
		};
		const runs = store.add(new AssistantRuns(reach, storage, mcpService, new class extends mock<IRequestService>() { }, fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		store.add(runs.onDidRequestApproval(request => { runs.respondApproval(request.id, 'reject', 'Keep the note empty'); }));
		assert.strictEqual(await runs.send('account', 'Write a note'), true);
		assert.strictEqual(sent.length, 2);
		assert.match(String(sent[1].messages.at(-1)?.content), /Reason: Keep the note empty/);
		const activity = new AssistantConversationStore(storage).read('account').turns[0].activity;
		assert.deepStrictEqual(activity?.map(item => [item.status, item.detail]), [['rejected', 'Declined: Keep the note empty']]);
	});

	test('reads workspace files through the agent loop without leaving the workspace', async () => {
		const storage = store.add(new InMemoryStorageService());
		const sent: ReachAssistantRequest[] = [];
		const responses: ReachAssistantResponse[] = [
			{ kind: 'reply', text: '```tool_call\n{"name": "assistant.read_file", "arguments": {"path": "src/app.ts"}}\n```' },
			{ kind: 'reply', text: '```tool_call\n{"name": "assistant.read_file", "arguments": {"path": "../secret.txt"}}\n```' },
			{ kind: 'reply', text: 'The file exports an app.' },
		];
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_accountId: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				sent.push(request);
				return responses.shift()!;
			}
		};
		const files = fileServiceMock({ '/ws/src/app.ts': 'export const app = 1;\nexport const other = 2;' });
		const runs = store.add(new AssistantRuns(reach, storage, emptyMcpService(), new class extends mock<IRequestService>() { }, files, workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		assert.strictEqual(await runs.send('account', 'Read the app file'), true);
		const toolResults = sent.flatMap(request => request.messages.filter(m => m.role === 'user' && typeof m.content === 'string' && m.content.includes('<tool_result')).map(m => m.content));
		assert.match(toolResults[0], /1: export const app = 1;/);
		assert.match(toolResults[1], /Paths must stay inside the workspace folder/);
	});

	test('writes workspace files after approval and blocks traversal', async () => {
		const storage = store.add(new InMemoryStorageService());
		const files: Record<string, string> = {};
		const responses: ReachAssistantResponse[] = [
			{ kind: 'reply', text: '```tool_call\n{"name": "assistant.write_file", "arguments": {"path": "notes/todo.md", "content": "# todo"}}\n```' },
			{ kind: 'reply', text: '```tool_call\n{"name": "assistant.write_file", "arguments": {"path": "../escape.txt", "content": "x"}}\n```' },
			{ kind: 'reply', text: 'Created the note; the other path was refused.' },
		];
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_accountId: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				return responses.shift()!;
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, emptyMcpService(), new class extends mock<IRequestService>() { }, fileServiceMock(files), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		store.add(runs.onDidRequestApproval(request => { runs.respondApproval(request.id, 'allow'); }));
		assert.strictEqual(await runs.send('account', 'Make a todo file'), true);
		assert.strictEqual(files['/ws/notes/todo.md'], '# todo');
		assert.strictEqual(Object.keys(files).length, 1);
		const toolResults = new AssistantConversationStore(storage).read('account').turns[0].activity;
		assert.deepStrictEqual(toolResults?.map(item => item.status), ['done', 'done']);
	});

	test('keeps an active request scoped to its starting worktree through approval', async () => {
		const storage = store.add(new InMemoryStorageService());
		const checkout = URI.parse('file:///worktree-feature');
		const project = URI.parse('file:///ws');
		const agentChats = upcastPartial<IAssistantAgentChats>({
			target: { workspace: project, worktree: { uri: checkout, label: 'feature' } },
			getWorktreeStatus: async () => 'ready',
		});
		const conversation = new AssistantConversationStore(storage);
		const taskId = conversation.createTask('account', agentChats.target);
		const files: Record<string, string> = { '/worktree-feature/.git': 'gitdir: /ws/.git/worktrees/feature', '/ws/notes.md': 'main checkout' };
		const responses: ReachAssistantResponse[] = [
			{ kind: 'reply', text: '', toolCalls: [{ id: 'write', name: 'assistant.write_file', arguments: '{"path":"notes.md","content":"feature checkout"}' }] },
			{ kind: 'reply', text: 'Updated the note.' },
			{ kind: 'reply', text: '', toolCalls: [{ id: 'resume-write', name: 'assistant.write_file', arguments: '{"path":"resumed.md","content":"resumed in feature checkout"}' }] },
			{ kind: 'reply', text: 'Resumed the task.' },
		];
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt() { return responses.shift()!; }
		};
		const runs = store.add(new AssistantRuns(reach, storage, emptyMcpService(), new class extends mock<IRequestService>() { }, fileServiceMock(files), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChats, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		let argsPreview: string | undefined;
		store.add(runs.onDidRequestApproval(request => {
			argsPreview = request.argsPreview;
			agentChats.target = { workspace: project };
			runs.respondApproval(request.id, 'allow');
		}));
		assert.strictEqual(await runs.send('account', 'Update the note'), true);
		assert.ok(argsPreview?.includes(checkout.fsPath));
		assert.strictEqual(files['/ws/notes.md'], 'main checkout');
		assert.deepStrictEqual({ ...files }, { '/worktree-feature/.git': 'gitdir: /ws/.git/worktrees/feature', '/ws/notes.md': 'main checkout', '/worktree-feature/notes.md': 'feature checkout' });

		const reopened = store.add(new AssistantRuns(reach, storage, emptyMcpService(), new class extends mock<IRequestService>() { }, fileServiceMock(files), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChats, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		store.add(reopened.onDidRequestApproval(request => {
			assert.ok(request.argsPreview.includes(checkout.fsPath));
			reopened.respondApproval(request.id, 'allow');
		}));
		conversation.setDraft('account', 'Unsent follow-up', taskId);
		assert.strictEqual(await reopened.send('account', 'Resume on the original checkout', taskId, { preserveDraft: true }), true);
		assert.strictEqual(conversation.read('account', taskId).draft, 'Unsent follow-up');
		assert.strictEqual(files['/worktree-feature/resumed.md'], 'resumed in feature checkout');
		assert.strictEqual(files['/ws/resumed.md'], undefined);
	});

	test('drains explicit follow-ups in order with their captured checkout and preserves a separate draft', async () => {
		const storage = store.add(new InMemoryStorageService());
		const conversation = new AssistantConversationStore(storage);
		const original = URI.file('/original-worktree');
		const other = URI.file('/other-worktree');
		const agentChats = upcastPartial<IAssistantAgentChats>({
			target: { worktree: { uri: original, label: 'original' } },
			getWorktreeStatus: async () => 'ready',
		});
		const waiting = new DeferredPromise<ReachAssistantResponse>();
		const finished = new DeferredPromise<void>();
		const submitted: string[] = [];
		const files: Record<string, string> = {};
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_account: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				const last = request.messages.at(-1);
				if (last?.role === 'tool') { return { kind: 'reply', text: 'Saved.' }; }
				const prompt = last?.content ?? '';
				submitted.push(prompt);
				if (prompt === 'First') { return waiting.p; }
				return { kind: 'reply', text: '', toolCalls: [{ id: 'write', name: 'assistant.write_file', arguments: JSON.stringify({ path: `${prompt}.txt`, content: prompt }) }] };
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, emptyMcpService(), new class extends mock<IRequestService>() { }, fileServiceMock(files), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChats, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		store.add(runs.onDidRequestApproval(request => { runs.respondApproval(request.id, 'allow'); }));
		store.add(runs.onDidChange(account => {
			if (!runs.getPending(account) && conversation.read(account).turns.length === 3) { void finished.complete(); }
		}));
		const first = runs.send('account', 'First');
		assert.ok(runs.queuePrompt('account', 'Second'));
		assert.ok(runs.queuePrompt('account', 'Third'));
		assert.deepStrictEqual(conversation.read('account').queuedPrompts?.map(item => item.prompt), ['Second', 'Third']);
		agentChats.target = { worktree: { uri: other, label: 'other' } };
		conversation.setDraft('account', 'Draft to keep');
		await waiting.complete({ kind: 'reply', text: 'First finished.' });
		assert.ok(await first);
		await finished.p;
		assert.deepStrictEqual(submitted, ['First', 'Second', 'Third']);
		assert.deepStrictEqual(files, { '/original-worktree/Second.txt': 'Second', '/original-worktree/Third.txt': 'Third' });
		assert.strictEqual(conversation.read('account').draft, 'Draft to keep');
	});

	test('queued routine invocations keep their selected connection scope', async () => {
		const storage = store.add(new InMemoryStorageService());
		const conversation = new AssistantConversationStore(storage);
		let autostarts = 0;
		let gmailStarts = 0;
		let unrelatedStarts = 0;
		const gmail = fakeServer({ name: 'Gmail', id: 'gmail-server', running: false, tools: [], onStart: () => { gmailStarts++; } });
		const unrelated = fakeServer({ name: 'Project tracker', id: 'tracker-server', running: false, tools: [], onStart: () => { unrelatedStarts++; } });
		const mcp = new class extends mock<IMcpService>() {
			override readonly servers = observableValue<readonly IMcpServer[]>('fixtureServers', [gmail, unrelated]);
			override autostart() { autostarts++; return constObservable(IAutostartResult.Empty); }
		};
		const waiting = new DeferredPromise<ReachAssistantResponse>();
		const finished = new DeferredPromise<void>();
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_account: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				return request.messages.at(-1)?.content === 'First' ? waiting.p : { kind: 'reply', text: 'Routine complete.' };
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, mcp, upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		store.add(runs.onDidChange(account => {
			if (!runs.getPending(account) && conversation.read(account).turns.length === 2) { void finished.complete(); }
		}));
		const first = runs.send('account', 'First');
		assert.ok(runs.queuePrompt('account', '<attachment name="Routine · Weekly report">Summarize mail.</attachment>\n\nRun this routine now.', undefined, undefined, undefined, [assistantConnectionKey(gmail)]));
		assert.deepStrictEqual([gmailStarts, unrelatedStarts], [0, 0]);
		await waiting.complete({ kind: 'reply', text: 'First complete.' });
		assert.ok(await first);
		await finished.p;
		assert.deepStrictEqual([autostarts, gmailStarts, unrelatedStarts], [0, 1, 0]);
		assert.strictEqual(unrelated.connectionState.get().state, McpConnectionState.Kind.Stopped);
	});

	test('keeps queued messages paused after a stop and across a new runs service', async () => {
		const storage = store.add(new InMemoryStorageService());
		const conversation = new AssistantConversationStore(storage);
		const waiting = new DeferredPromise<ReachAssistantResponse>();
		const requests: string[] = [];
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_account: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				const prompt = request.messages.at(-1)?.content ?? '';
				requests.push(prompt);
				return prompt === 'Stop this' ? waiting.p : { kind: 'reply', text: 'Sent explicitly.' };
			}
		};
		const createRuns = () => store.add(new AssistantRuns(reach, storage, emptyMcpService(), new class extends mock<IRequestService>() { }, fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		const runs = createRuns();
		const pending = runs.send('account', 'Stop this');
		assert.ok(runs.queuePrompt('account', 'Keep this'));
		runs.stop('account');
		await waiting.complete({ kind: 'stopped' });
		assert.strictEqual(await pending, false);
		assert.deepStrictEqual(requests, ['Stop this']);
		const reopened = createRuns();
		assert.strictEqual(reopened.getPending('account'), undefined);
		assert.strictEqual(conversation.read('account').queuedPrompts?.[0].prompt, 'Keep this');
		const finished = new DeferredPromise<void>();
		store.add(reopened.onDidChange(account => {
			if (!reopened.getPending(account) && conversation.read(account).turns.length) { void finished.complete(); }
		}));
		assert.ok(reopened.sendQueuedPrompt('account'));
		await finished.p;
		assert.deepStrictEqual(requests, ['Stop this', 'Keep this']);
	});

	test('searches workspace contents through the search service', async () => {
		const storage = store.add(new InMemoryStorageService());
		const responses: ReachAssistantResponse[] = [
			{ kind: 'reply', text: '```tool_call\n{"name": "assistant.search_files", "arguments": {"pattern": "reachAssistant"}}\n```' },
			{ kind: 'reply', text: 'It lives in the platform layer.' },
		];
		const queries: unknown[] = [];
		const search = new class extends mock<ISearchService>() {
			override async textSearch(query: unknown) {
				queries.push(query);
				return {
					results: [{
						resource: URI.parse('file:///ws/src/app.ts'),
						results: [{ previewText: '  const reachAssistant = 1;  ', rangeLocations: [{ source: { startLineNumber: 4, startColumn: 0, endLineNumber: 4, endColumn: 0 }, preview: { startLineNumber: 4, startColumn: 0, endLineNumber: 4, endColumn: 0 } }] }],
					}],
					limitHit: false, messages: [], stats: undefined,
				} as never;
			}
		};
		const sent: ReachAssistantRequest[] = [];
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_accountId: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				sent.push(request);
				return responses.shift()!;
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, emptyMcpService(), new class extends mock<IRequestService>() { }, fileServiceMock(), workspaceServiceMock(), search, quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		assert.strictEqual(await runs.send('account', 'Find reachAssistant'), true);
		assert.strictEqual(queries.length, 1);
		const query = queries[0] as { contentPattern: { pattern: string }; folderQueries: { folder: URI }[] };
		assert.strictEqual(query.contentPattern.pattern, 'reachAssistant');
		assert.strictEqual(query.folderQueries[0].folder.toString(), 'file:///ws');
		const toolResults = sent.flatMap(request => request.messages.filter(m => m.role === 'user' && typeof m.content === 'string' && m.content.includes('<tool_result')).map(m => m.content));
		assert.match(toolResults[0], /src\/app\.ts:5: const reachAssistant = 1;/);
	});

	test('starts a linked task through the agent loop after approval', async () => {
		const storage = store.add(new InMemoryStorageService());
		const checkout = URI.parse('file:///research-feature');
		const agentChats = upcastPartial<IAssistantAgentChats>({ target: { workspace: URI.parse('file:///ws'), worktree: { uri: checkout, label: 'research' } } });
		const responses: ReachAssistantResponse[] = [
			{ kind: 'reply', text: '```tool_call\n{"name": "assistant.create_task", "arguments": {"title": "Research", "instructions": "Look into the search API"}}\n```' },
			{ kind: 'reply', text: 'Started the research task.' },
			{ kind: 'reply', text: 'Task working.' },
		];
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_accountId: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				return responses.shift() ?? { kind: 'reply', text: 'Done.' };
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, emptyMcpService(), new class extends mock<IRequestService>() { }, fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChats, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		store.add(runs.onDidRequestApproval(request => {
			agentChats.target = { workspace: URI.parse('file:///other-project') };
			runs.respondApproval(request.id, 'allow');
		}));
		assert.strictEqual(await runs.send('account', 'Kick off research'), true);
		const tasks = new AssistantConversationStore(storage).read('account').tasks;
		assert.strictEqual(tasks?.length, 1);
		assert.strictEqual(tasks?.[0].title, 'Research');
		assert.strictEqual(readAssistantTaskTarget(tasks?.[0])?.worktree?.uri.toString(), checkout.toString());
	});

	test('edits workspace files with unique-match guards', async () => {
		const storage = store.add(new InMemoryStorageService());
		const files: Record<string, string> = { '/ws/src/app.ts': 'let x = 1;\nlet y = 2;\nlet x = 3;' };
		const responses: ReachAssistantResponse[] = [
			{ kind: 'reply', text: '```tool_call\n{"name": "assistant.edit_file", "arguments": {"path": "src/app.ts", "old_text": "let x", "new_text": "const x"}}\n```' },
			{ kind: 'reply', text: '```tool_call\n{"name": "assistant.edit_file", "arguments": {"path": "src/app.ts", "old_text": "let y = 2;", "new_text": "const y = 4;"}}\n```' },
			{ kind: 'reply', text: 'Updated the y declaration; x was ambiguous.' },
		];
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_accountId: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				return responses.shift()!;
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, emptyMcpService(), new class extends mock<IRequestService>() { }, fileServiceMock(files), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		store.add(runs.onDidRequestApproval(request => { runs.respondApproval(request.id, 'allow'); }));
		assert.strictEqual(await runs.send('account', 'Fix the declarations'), true);
		assert.strictEqual(files['/ws/src/app.ts'], 'let x = 1;\nconst y = 4;\nlet x = 3;');
	});

	test('creates persistent assistant documents and records them as transcript artifacts', async () => {
		const storage = store.add(new InMemoryStorageService());
		let step = 0;
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_accountId: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				step++;
				if (step === 1) {
					assert.ok(!request.tools?.some(tool => tool.name === 'assistant.create_document'));
					assert.ok(request.tools?.some(tool => tool.name === 'assistant.find_app_tools'));
					assert.match(request.messages[0].content, /substantial reusable deliverables/);
					return { kind: 'reply', text: '', toolCalls: [{ id: 'discover-document', name: 'assistant.find_app_tools', arguments: '{"name":"assistant.create_document"}' }] };
				}
				if (step === 2) {
					assert.ok(request.tools?.some(tool => tool.name === 'assistant.create_document'));
					return { kind: 'reply', text: '', toolCalls: [{ id: 'create-document', name: 'assistant.create_document', arguments: JSON.stringify({ title: 'Launch Plan', content: '# Launch Plan — A complete rollout plan.' }) }] };
				}
				assert.match(request.messages.at(-1)?.content ?? '', /"status":"created"/);
				return { kind: 'reply', text: 'Created Launch Plan.' };
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, emptyMcpService(), new class extends mock<IRequestService>() { }, fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		assert.strictEqual(await runs.send('account', 'Create a launch plan document'), true);
		const conversation = new AssistantConversationStore(storage);
		const [document] = conversation.listDocuments('account');
		assert.ok(document);
		assert.strictEqual(document.title, 'Launch Plan');
		assert.match(document.content, /complete rollout plan/);
		const activity = conversation.read('account').turns[0].activity?.find(item => item.tool === 'assistant.create_document');
		assert.strictEqual(activity?.documentId, document.id);
	});

	test('updates persistent assistant documents without losing their identity', () => {
		const storage = store.add(new InMemoryStorageService());
		const conversation = new AssistantConversationStore(storage);
		const created = conversation.saveDocument('account', 'Launch Plan', '# Launch Plan\n\nDraft.');
		assert.strictEqual(created.status, 'saved');
		if (created.status !== 'saved') { return; }
		const updated = conversation.saveDocument('account', 'Launch Plan v2', '# Launch Plan\n\nFinal.', created.document.id);
		assert.strictEqual(updated.status, 'saved');
		if (updated.status !== 'saved') { return; }
		assert.strictEqual(updated.document.id, created.document.id);
		assert.strictEqual(updated.document.createdAt, created.document.createdAt);
		const reopened = new AssistantConversationStore(storage);
		assert.strictEqual(reopened.readDocument('account', created.document.id)?.title, 'Launch Plan v2');
		assert.strictEqual(reopened.listDocuments('account').length, 1);
		assert.strictEqual(reopened.removeDocument('account', created.document.id), true);
		assert.strictEqual(reopened.readDocument('account', created.document.id), undefined);
	});

	test('opens a workspace file at the requested line and column', async () => {
		const storage = store.add(new InMemoryStorageService());
		let opened: URI | string | undefined;
		const openerService = new class extends mock<IOpenerService>() {
			override async open(target: URI | string) { opened = target; return true; }
		};
		const catalog = collectAssistantTools({
			accountId: 'account', store: new AssistantConversationStore(storage), mcpService: emptyMcpService(),
			requestService: new class extends mock<IRequestService>() { }, fileService: fileServiceMock({ '/ws/src/app.ts': 'const ready = true;' }),
			workspaceService: workspaceServiceMock(), searchService: searchServiceMock(), openerService,
		});
		const tool = catalog.entries.get('assistant.open_file');
		assert.ok(tool?.kind === 'builtin');
		assert.strictEqual(await tool.run({ path: 'src/app.ts', line: 3, column: 2 }, CancellationToken.None), 'Opened app.ts at line 3 in the editor.');
		assert.deepStrictEqual(opened, URI.parse('file:///ws/src/app.ts#3,2'));
	});

	test('grounds repository requests in a selected local project in a folderless Agents window', async () => {
		const storage = store.add(new InMemoryStorageService());
		const workspaceService = upcastPartial<IWorkspaceContextService>({ getWorkspace: () => ({ folders: [] }) as never });
		const project = URI.file('/projects/vscode');
		const agentChats = upcastPartial<IAssistantAgentChats>({ target: { workspace: project } });
		let step = 0;
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_accountId: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				step++;
				assert.match(request.messages[0].content, /Workspace folder: "vscode"/);
				assert.match(request.messages[0].content, /inspect it first with assistant\.list_files or assistant\.read_file/);
				if (step === 1) {
					return { kind: 'reply', text: '<tool_call><function=assistant.read_file><parameter=path>README.md</parameter></function></tool_call>' };
				}
				assert.match(request.messages.at(-1)?.content ?? '', /1: VS Code fork marker/);
				return { kind: 'reply', text: 'The selected VS Code fork contains the marker.' };
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, emptyMcpService(), upcastPartial<IRequestService>({}), fileServiceMock({ '/projects/vscode/README.md': 'VS Code fork marker' }), workspaceService, searchServiceMock(), quickInputStub, openerStub, agentChats, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		assert.strictEqual(await runs.send('account', 'Summarize my VS Code fork, labeled vscode.'), true);
		assert.strictEqual(step, 2);
		assert.deepStrictEqual(getAssistantWorkspaceFolders(workspaceService, { workspace: project, fileRoots: [] }), []);
	});

	test('grounds bare retries and explicit corrections in the exact GitHub repo', async () => {
		const storage = store.add(new InMemoryStorageService());
		const conversation = new AssistantConversationStore(storage);
		conversation.appendTurn('account', { prompt: 'Summarize my VS Code fork repo.', reply: 'The wrong repo is VSCode-SignalR-EA-C-H.' });
		conversation.appendTurn('account', { prompt: 'ok now summarize my SignalHive repo', reply: 'The wrong repo is on main.' });
		conversation.appendTurn('account', { prompt: 'try again', reply: 'The wrong repo is still on main.' });
		const reads: Record<string, unknown>[] = [];
		const github = fakeServer({
			name: 'GitHub', tools: [
				upcastPartial<IMcpTool>({
					id: 'mcp_github_get_file_contents', visibility: McpToolVisibility.Model,
					definition: { name: 'get_file_contents', inputSchema: { type: 'object' }, annotations: { readOnlyHint: true } },
					call: async args => {
						reads.push(args);
						return { content: [{ type: 'text', text: '# SignalHive fixture' }] };
					},
				}),
			]
		});
		let step = 0;
		let input: string;
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_accountId: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				step++;
				assert.match(request.messages[0].content, /"SignalHive"/);
				assert.strictEqual(request.messages[1].content, input === 'try  again' ? 'ok now summarize my SignalHive repo' : input);
				assert.ok(!request.messages.some(message => message.content.includes('The wrong repo is')));
				assert.ok(!request.tools?.some(tool => tool.name === 'assistant.read_file'), 'The unrelated local checkout must not be offered');
				if (step === 1) { return { kind: 'reply', text: 'Invented repository summary.' }; }
				if (step <= 3) {
					const repo = step === 2 ? 'VSCode-SignalR-EA-C-H' : 'SignalHive';
					return { kind: 'reply', text: '', toolCalls: [{ id: `call-${step}`, name: 'mcp_github_get_file_contents', arguments: JSON.stringify({ owner: 'fixture', repo, path: 'README.md' }) }] };
				}
				assert.strictEqual(request.messages.at(-1)?.content, '# SignalHive fixture');
				return { kind: 'reply', text: 'SignalHive has a verified README.' };
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, mcpServiceWith(github), upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		for (input of ['try  again', 'i think you got the wrong one, try again, SignalHive repo in my Github']) {
			step = 0;
			reads.length = 0;
			assert.strictEqual(await runs.send('account', input), true);
			assert.deepStrictEqual(reads, [{ owner: 'fixture', repo: 'SignalHive', path: 'README.md' }]);
			assert.strictEqual(conversation.read('account').turns.at(-1)?.reply, 'SignalHive has a verified README.');
		}
	});

	test('reports missing repository evidence when the model keeps answering without tools', async () => {
		const storage = store.add(new InMemoryStorageService());
		let requests = 0;
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(): Promise<ReachAssistantResponse> {
				requests++;
				return { kind: 'reply', text: 'Unverified remote repository details.' };
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, emptyMcpService(), upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		assert.strictEqual(await runs.send('account', 'Summarize my SignalHive repo.'), true);
		assert.strictEqual(requests, 2);
		assert.match(new AssistantConversationStore(storage).read('account').turns[0].reply, /couldn’t read the requested repository from GitHub/);
	});

	test('grounds a qualified GitHub repo mentioned with its README', async () => {
		const storage = store.add(new InMemoryStorageService());
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_accountId: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				assert.match(request.messages[0].content, /"fixture\/SignalHive"/);
				return { kind: 'reply', text: 'Unverified details.' };
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, emptyMcpService(), upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		assert.strictEqual(await runs.send('account', 'Summarize `fixture/SignalHive` from its README.'), true);
		assert.match(new AssistantConversationStore(storage).read('account').turns[0].reply, /couldn’t read the requested repository from GitHub/);
	});

	test('restores named workspace roots and returns searchable paths that read the right file', async () => {
		const storage = store.add(new InMemoryStorageService());
		const conversation = new AssistantConversationStore(storage);
		const fileRoots = [URI.file('/api/shared'), URI.file('/web/shared')];
		const markers = fileRoots.map(root => URI.joinPath(root, 'marker.txt'));
		const taskId = conversation.createTask('account', { fileRoots, fileRootNames: ['API', 'Web UI'], activeEditorFile: markers[0] });
		const target = readAssistantTaskTarget(conversation.read('account').tasks?.find(task => task.id === taskId));
		assert.deepStrictEqual(target?.fileRootNames, ['API', 'Web UI']);
		assert.strictEqual(target?.activeEditorFile?.toString(), markers[0].toString());
		const catalog = collectAssistantTools({
			accountId: 'account', store: conversation, mcpService: emptyMcpService(),
			requestService: new class extends mock<IRequestService>() { },
			fileService: fileServiceMock({ '/api/shared/marker.txt': 'API marker', '/web/shared/marker.txt': 'Web marker' }),
			workspaceService: workspaceServiceMock(), openerService: openerStub, agentChatTarget: target,
			searchService: upcastPartial<ISearchService>({
				fileSearch: async query => {
					assert.deepStrictEqual(query.folderQueries.map(folder => folder.folder.toString()), fileRoots.map(root => root.toString()));
					return { results: markers.map(resource => ({ resource })), messages: [] };
				},
			}),
		});
		const find = catalog.entries.get('assistant.find_files');
		const read = catalog.entries.get('assistant.read_file');
		assert.ok(find?.kind === 'builtin' && read?.kind === 'builtin');
		const paths = (await find.run({ pattern: 'marker.txt' }, CancellationToken.None)).split('\n');
		assert.deepStrictEqual(paths, ['API/marker.txt', 'Web UI/marker.txt']);
		for (const [i, marker] of ['API marker', 'Web marker'].entries()) {
			assert.strictEqual(await read.run({ path: paths[i] }, CancellationToken.None), '1: ' + marker);
		}
		assert.strictEqual(await read.run({ path: 'marker.txt' }, CancellationToken.None), '1: API marker');
		assert.match(await read.run({ path: 'shared/marker.txt' }, CancellationToken.None), /ambiguous/);
		assert.match(await read.run({ path: 'Web UI/../secret.txt' }, CancellationToken.None), /stay inside/);
	});

	test('gives duplicate root names stable prefixes that avoid existing names', () => {
		const fileRoots = [URI.file('/api/shared'), URI.file('/web/shared'), URI.file('/tests/shared (1)')];
		const folders = getAssistantWorkspaceFolders(workspaceServiceMock(), { fileRoots });
		assert.deepStrictEqual(folders.map(folder => folder.name), ['shared (2)', 'shared (3)', 'shared (1)']);
		assert.match(String(resolveAssistantWorkspacePath(folders, 'shared/check.txt')), /ambiguous/);
		assert.deepStrictEqual(resolveAssistantWorkspacePath(folders, 'shared (3)/check.txt'), URI.file('/web/shared/check.txt'));
	});

	if (isWindows) {
		test('formats live Windows editor paths against restored roots with lowercase drive letters', () => {
			const roots = [
				{ uri: URI.parse('file:///d:/fixture/service'), name: 'Service', originalName: 'Service' },
				{ uri: URI.parse('file:///d:/fixture/ui'), name: 'UI', originalName: 'UI' },
			];
			const liveFile = URI.file('D:\\fixture\\service\\main.ts');
			assert.strictEqual(formatAssistantWorkspacePath(roots, liveFile), 'Service/main.ts');
			assert.strictEqual(formatAssistantWorkspacePath(roots, URI.file('E:\\fixture\\service\\main.ts')), undefined);
			assert.strictEqual(formatAssistantWorkspacePath(roots, URI.file('D:\\fixture\\Service\\main.ts')), undefined);
		});
	}

	test('lists and reads linked tasks by id', async () => {
		const storage = store.add(new InMemoryStorageService());
		const conversation = new AssistantConversationStore(storage);
		const taskId = conversation.createTask('account');
		conversation.updateTask('account', taskId, { title: 'Research', status: 'completed', update: 'Found the API.' });
		conversation.appendTurn('account', { prompt: 'Look into the search API', reply: 'ISearchService does it.' }, taskId);
		const responses: ReachAssistantResponse[] = [
			{ kind: 'reply', text: '```tool_call\n{"name": "assistant.list_tasks", "arguments": {}}\n```' },
			{ kind: 'reply', text: '```tool_call\n{"name": "assistant.read_task", "arguments": {"id": "' + taskId + '"}}\n```' },
			{ kind: 'reply', text: 'Research found ISearchService.' },
		];
		const sent: ReachAssistantRequest[] = [];
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_accountId: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				sent.push(request);
				return responses.shift()!;
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, emptyMcpService(), new class extends mock<IRequestService>() { }, fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		assert.strictEqual(await runs.send('account', 'How is research going?'), true);
		const toolResults = sent.flatMap(request => request.messages.filter(m => m.role === 'user' && typeof m.content === 'string' && m.content.includes('<tool_result')).map(m => m.content));
		assert.match(toolResults[0], /Research \[completed\] \(id: /);
		assert.match(toolResults[1], /ISearchService does it\./);
	});

	test('declines write actions without asking in safe-only mode', async () => {
		const storage = store.add(new InMemoryStorageService());
		const writeTool = new class extends mock<IMcpTool>() {
			override readonly id = 'mcp_notes_write';
			override readonly visibility = McpToolVisibility.Model;
			override readonly definition = { name: 'mcp_notes_write', annotations: { destructiveHint: true }, inputSchema: { type: 'object' } } as never;
			override call(): never { throw new Error('must not run'); }
		};
		const server = new class extends mock<IMcpServer>() {
			override readonly definition = { label: 'Notes' } as never;
			override readonly connectionState = constObservable<McpConnectionState>({ state: McpConnectionState.Kind.Running });
			override readonly enablement = constObservable(ContributionEnablementState.EnabledWorkspace);
			override readonly serverMetadata = constObservable({ serverName: 'Notes', icons: undefined as never });
			override readonly tools = constObservable<IMcpTool[]>([writeTool]);
			override readonly prompts = constObservable([]);
		};
		const mcpService = new class extends mock<IMcpService>() {
			override readonly servers = constObservable<IMcpServer[]>([server]);
		};
		const conversation = new AssistantConversationStore(storage);
		conversation.allowTool('account', 'mcp_notes_write');
		conversation.setPermissionMode('account', 'safeOnly');
		const responses: ReachAssistantResponse[] = [
			{ kind: 'reply', text: '```tool_call\n{"name": "mcp_notes_write", "arguments": {"note":"x"}}\n```' },
			{ kind: 'reply', text: 'I could not write it — writes are disabled in this thread.' },
		];
		let approvalAsked = false;
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_accountId: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				return responses.shift()!;
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, mcpService, new class extends mock<IRequestService>() { }, fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		store.add(runs.onDidRequestApproval(() => { approvalAsked = true; }));
		assert.strictEqual(await runs.send('account', 'Write a note'), true);
		assert.strictEqual(approvalAsked, false);
		const activity = new AssistantConversationStore(storage).read('account').turns[0].activity;
		assert.deepStrictEqual(activity?.map(item => item.status), ['rejected']);
		assert.deepStrictEqual(conversation.read('account').allowedTools, ['mcp_notes_write']);
	});

	test('requires approval when an agent launch grant is revoked while Assistant is waiting for a model reply', async () => {
		const storage = store.add(new InMemoryStorageService());
		const conversation = new AssistantConversationStore(storage);
		conversation.allowTool('account', 'assistant.start_agent');
		const toolReply = new DeferredPromise<ReachAssistantResponse>();
		let requests = 0;
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(): Promise<ReachAssistantResponse> {
				return requests++ === 0 ? toolReply.p : { kind: 'reply', text: 'Done.' };
			}
		};
		let launches = 0;
		const agentChats = upcastPartial<IAssistantAgentChats>({ start: async () => { launches++; return 'Started an agent'; } });
		const runs = store.add(new AssistantRuns(reach, storage, emptyMcpService(), new class extends mock<IRequestService>() { }, fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChats, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		const approvals: string[] = [];
		store.add(runs.onDidRequestApproval(request => {
			approvals.push(request.tool);
			runs.respondApproval(request.id, 'reject');
		}));
		const pending = runs.send('account', 'Start an agent');
		assert.strictEqual(requests, 1);
		conversation.setAllowedTools('account', []);
		await toolReply.complete({ kind: 'reply', text: '', toolCalls: [{ id: 'launch', name: 'assistant.start_agent', arguments: '{"prompt":"Check the build"}' }] });
		assert.strictEqual(await pending, true);
		assert.deepStrictEqual(approvals, ['assistant.start_agent']);
		assert.strictEqual(launches, 0);
	});

	test('remembers the worktree by local project and account, including when opened from its checkout', () => {
		const storage = store.add(new InMemoryStorageService());
		const project = URI.file('/project');
		const checkout = { uri: URI.file('/project-feature'), label: 'feature', identity: { gitDirectory: URI.file('/project/.git/worktrees/feature'), createdAt: 1 } };
		const first = new AssistantConversationStore(storage);
		first.setWorktreeSelection('account-a', project, checkout);
		first.setLastWorktreeProject('account-a', project);
		first.setWorktreeSelection('legacy-account', project, { uri: checkout.uri, label: checkout.label });

		const reopened = new AssistantConversationStore(storage);
		assert.strictEqual(reopened.getLastWorktreeProject('account-a')?.toString(), project.toString());
		for (const folder of [project, checkout.uri]) {
			const selection = reopened.getWorktreeSelection('account-a', folder);
			assert.ok(selection);
			assert.strictEqual(selection.project.toString(), project.toString());
			assert.strictEqual(selection.worktree.uri.toString(), checkout.uri.toString());
			assert.strictEqual(selection.worktree.label, checkout.label);
			assert.strictEqual(selection.worktree.identity?.gitDirectory.toString(), checkout.identity.gitDirectory.toString());
			assert.strictEqual(selection.worktree.identity.createdAt, checkout.identity.createdAt);
		}
		assert.strictEqual(reopened.getWorktreeSelection('account-a', URI.file('/other-project')), undefined);
		assert.strictEqual(reopened.getWorktreeSelection('account-b', project), undefined);
		assert.strictEqual(reopened.getLastWorktreeProject('account-b'), undefined);
		const legacy = reopened.getWorktreeSelection('legacy-account', project)?.worktree;
		assert.strictEqual(legacy?.uri.toString(), checkout.uri.toString());
		assert.strictEqual(legacy.identity, undefined);
		reopened.setWorktreeSelection('account-a', project, undefined);
		assert.strictEqual(reopened.getWorktreeSelection('account-a', checkout.uri), undefined);
	});

	test('restores the last open task only for its account and project', () => {
		const storage = store.add(new InMemoryStorageService());
		const first = new AssistantConversationStore(storage);
		const project = URI.file('/project');
		const other = URI.file('/other-project');
		const task = first.createTask('account-a');
		first.setLastOpenTask('account-a', project, task);
		const reopened = new AssistantConversationStore(storage);
		assert.strictEqual(reopened.getLastOpenTask('account-a', project), task);
		assert.strictEqual(reopened.getLastOpenTask('account-a', other), undefined);
		assert.strictEqual(reopened.getLastOpenTask('account-a', undefined), undefined);
		assert.strictEqual(reopened.getLastOpenTask('account-b', project), undefined);
		reopened.setLastOpenTask('account-a', project, undefined);
		assert.strictEqual(reopened.getLastOpenTask('account-a', project), undefined);

		reopened.setLastOpenTask('account-a', undefined, task);
		assert.strictEqual(reopened.getLastOpenTask('account-a', undefined), task);
		assert.strictEqual(reopened.getLastOpenTask('account-a', project), undefined);
		reopened.deleteTask('account-a', task);
		assert.strictEqual(reopened.getLastOpenTask('account-a', undefined), undefined);
	});

	test('uses the same conversation after reopening and keeps other accounts separate', () => {
		const storage = store.add(new InMemoryStorageService());
		const first = new AssistantConversationStore(storage);
		first.setInstructions('account-a', 'Keep replies short.');
		first.appendTurn('account-a', { prompt: 'Start here.', reply: 'Ready.' });
		first.appendTurn('account-b', { prompt: 'Another account.', reply: 'Ready.' });
		const reopened = new AssistantConversationStore(storage);
		reopened.appendTurn('account-a', { prompt: 'Continue.', reply: 'Continuing.' });
		const strip = (account: string) => {
			const convo = reopened.read(account);
			return { ...convo, turns: convo.turns.map(({ at: _at, ...turn }) => turn) };
		};
		assert.deepStrictEqual(strip('account-a'), {
			instructions: 'Keep replies short.',
			turns: [{ prompt: 'Start here.', reply: 'Ready.' }, { prompt: 'Continue.', reply: 'Continuing.' }],
		});
		assert.ok(reopened.read('account-a').turns.every(turn => typeof turn.at === 'number'));
		assert.deepStrictEqual(strip('account-b'), {
			instructions: '', turns: [{ prompt: 'Another account.', reply: 'Ready.' }],
		});
	});

	test('parses XML tool calls from local models and strips them from the display', () => {
		const text = 'I will pull the top-level view first.\n\n<tool_call>\n<function=mcp_github_get_file_contents>\n<parameter=path>\n\n</parameter>\n</function>\n</tool_call>';
		const parsed = parseTextToolCalls(text);
		assert.deepStrictEqual(parsed.calls, [{ name: 'mcp_github_get_file_contents', args: { path: '' } }]);
		assert.deepStrictEqual(parsed.errors, []);
		assert.strictEqual(parsed.display, 'I will pull the top-level view first.');
	});

	test('parses XML function bodies given as JSON and coerces parameters with the tool schema', () => {
		const specs = [{ name: 'mcp_github_list_issues', description: '', parameters: { type: 'object', properties: { per_page: { type: 'number' }, include_closed: { type: 'boolean' }, labels: { type: 'array' } } } }];
		const text = '<tool_call>\n<function=mcp_github_list_issues>{"owner":"acme","repo":"app"}</function>\n</tool_call>\n\n<tool_call>\n<function=mcp_github_list_issues>\n<parameter=per_page>5</parameter>\n<parameter=include_closed>true</parameter>\n<parameter=labels>["bug"]</parameter>\n<parameter=owner>acme</parameter>\n</function>\n</tool_call>';
		const parsed = parseTextToolCalls(text, specs);
		assert.deepStrictEqual(parsed.calls, [
			{ name: 'mcp_github_list_issues', args: { owner: 'acme', repo: 'app' } },
			{ name: 'mcp_github_list_issues', args: { per_page: 5, include_closed: true, labels: ['bug'], owner: 'acme' } },
		]);
		assert.strictEqual(parsed.display, '');
	});

	test('reports malformed XML tool blocks without leaving them in the display', () => {
		const parsed = parseTextToolCalls('Before.\n<tool_call>\nnot a function\n</tool_call>\nAfter.');
		assert.deepStrictEqual(parsed.calls, []);
		assert.strictEqual(parsed.errors.length, 1);
		assert.strictEqual(parsed.display, 'Before.\n\nAfter.');
		const badBody = parseTextToolCalls('<tool_call>\n<function=x>random text</function>\n</tool_call>');
		assert.deepStrictEqual(badBody.calls, []);
		assert.strictEqual(badBody.errors.length, 1);
	});

	test('keeps the fenced JSON tool-call protocol working', () => {
		const parsed = parseTextToolCalls('Checking.\n```tool_call\n{"name":"assistant.get_current_time","arguments":{}}\n```');
		assert.deepStrictEqual(parsed.calls, [{ name: 'assistant.get_current_time', args: {} }]);
		assert.strictEqual(parsed.display, 'Checking.');
	});

	test('executes XML tool calls from a local model and keeps its final reply', async () => {
		const storage = store.add(new InMemoryStorageService());
		const requests: ReachAssistantRequest[] = [];
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_account: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				requests.push(request);
				if (requests.length === 1) {
					return { kind: 'reply', text: 'Let me check.\n\n<tool_call>\n<function=assistant.get_current_time>\n</function>\n</tool_call>' };
				}
				assert.match(request.messages.at(-1)?.content ?? '', /tool_result name="assistant.get_current_time"/);
				return { kind: 'reply', text: 'Here is the summary you asked for.' };
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, emptyMcpService(), upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		assert.strictEqual(await runs.send('account', 'summarize the time'), true);
		assert.strictEqual(requests.length, 2);
		const turn = new AssistantConversationStore(storage).read('account').turns[0];
		assert.strictEqual(turn.reply, 'Here is the summary you asked for.');
		assert.ok(turn.activity?.some(item => item.tool === 'assistant.get_current_time' && item.status === 'done'));
	});

	test('offers tools from a connection that finished starting mid-turn and recovers on the next step', async () => {
		const storage = store.add(new InMemoryStorageService());
		let toolCalls = 0;
		const github = fakeServer({
			name: 'GitHub', tools: [
				upcastPartial<IMcpTool>({
					id: 'github_get_file_contents', visibility: McpToolVisibility.Model,
					definition: { name: 'get_file_contents', inputSchema: { type: 'object' }, annotations: { readOnlyHint: true } },
					call: async () => {
						toolCalls++;
						return { content: [{ type: 'text', text: 'SignalREACH readme contents' }] };
					},
				}),
			]
		});
		const mcp = new class extends mock<IMcpService>() {
			override readonly servers = observableValue<readonly IMcpServer[]>('fixtureServers', []);
			override autostart() { return constObservable(IAutostartResult.Empty); }
		};
		let step = 0;
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_account: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				step++;
				if (step === 1) {
					assert.ok(!request.tools?.some(tool => tool.name === 'github_get_file_contents'));
					// The connection finishes starting while the model is answering.
					mcp.servers.set([github], undefined);
					return { kind: 'reply', text: '', toolCalls: [{ id: 'probe', name: 'github_get_file_contents', arguments: '{}' }] };
				}
				assert.ok(request.tools?.some(tool => tool.name === 'github_get_file_contents'), 'the late connection must be offered on the next step');
				assert.match(request.messages[0].content, /github_get_file_contents/);
				if (step === 2) {
					assert.match(request.messages.at(-1)?.content ?? '', /unknown tool "github_get_file_contents"/);
					return { kind: 'reply', text: '', toolCalls: [{ id: 'probe2', name: 'github_get_file_contents', arguments: '{}' }] };
				}
				assert.strictEqual(request.messages.at(-1)?.content, 'SignalREACH readme contents');
				return { kind: 'reply', text: 'SignalREACH summarized.' };
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, mcp, upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		assert.strictEqual(await runs.send('account', 'summarize SignalREACH'), true);
		assert.strictEqual(toolCalls, 1);
		const turn = new AssistantConversationStore(storage).read('account').turns[0];
		assert.strictEqual(turn.reply, 'SignalREACH summarized.');
		assert.ok(turn.activity?.some(item => item.tool === 'github_get_file_contents' && item.status === 'done'));
	});

	test('saves Hermes-style memory batches and injects them into the next system prompt', async () => {
		const storage = store.add(new InMemoryStorageService());
		const requests: ReachAssistantRequest[] = [];
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_account: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				requests.push(request);
				if (requests.length === 1) {
					return { kind: 'reply', text: '', toolCalls: [{ id: 'mem', name: 'assistant.memory', arguments: JSON.stringify({ target: 'user', operations: [{ action: 'add', content: 'User is Michael, builds AI models' }] }) }] };
				}
				assert.match(request.messages[0].content, /USER PROFILE/);
				assert.match(request.messages[0].content, /User is Michael, builds AI models/);
				assert.match(request.messages[0].content, /USER PROFILE[^\[]*\[\d+% — \d+\/1400 chars\]/);
				assert.match(request.messages.at(-1)?.content ?? '', /Memory updated/);
				return { kind: 'reply', text: 'Saved.' };
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, emptyMcpService(), upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		assert.strictEqual(await runs.send('account', 'remember who I am'), true);
		assert.strictEqual(requests.length, 2);
		assert.strictEqual(new AssistantConversationStore(storage).readMemoryBank('account').user, 'User is Michael, builds AI models');
	});

	test('coerces stringified tool arguments against the tool schema before calling MCP tools', async () => {
		const storage = store.add(new InMemoryStorageService());
		let received: Record<string, unknown> | undefined;
		const github = fakeServer({
			name: 'GitHub', tools: [
				upcastPartial<IMcpTool>({
					id: 'github_search_repositories', visibility: McpToolVisibility.Model,
					definition: { name: 'search_repositories', inputSchema: { type: 'object', properties: { query: { type: 'string' }, minimal_output: { type: 'boolean' }, per_page: { type: 'number' } } }, annotations: { readOnlyHint: true } },
					call: async (args: Record<string, unknown>) => {
						received = args;
						return { content: [{ type: 'text', text: 'Found SignalR.E.A.C.H' }] };
					},
				}),
			]
		});
		const mcp = new class extends mock<IMcpService>() {
			override readonly servers = constObservable<readonly IMcpServer[]>([github]);
			override autostart() { return constObservable(IAutostartResult.Empty); }
		};
		let requests = 0;
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(): Promise<ReachAssistantResponse> {
				requests++;
				return requests === 1
					? { kind: 'reply', text: '', toolCalls: [{ id: 'search', name: 'github_search_repositories', arguments: JSON.stringify({ query: 'SignalREACH', minimal_output: 'true', per_page: '5' }) }] }
					: { kind: 'reply', text: 'Found the repo.' };
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, mcp, upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, mailWatcherStub, promptsStub));
		assert.strictEqual(await runs.send('account', 'find my repo'), true);
		assert.deepStrictEqual(received, { query: 'SignalREACH', minimal_output: true, per_page: 5 });
	});

	test('offers the mail-watch tool and injects reported alerts into the system prompt', async () => {
		const storage = store.add(new InMemoryStorageService());
		new AssistantConversationStore(storage).recordMailAlerts('account', [{ id: 'm1', connectionId: 'c', label: 'Outlook', subject: 'Urgent: review', from: 'boss@corp.com', at: Date.now() }]);
		const watcher = upcastPartial<IAssistantMailWatcher>({
			status: () => ({ enabled: true, connections: [{ label: 'Outlook', connected: true }], alerts: [] }),
			checkNow: async () => 0,
			setEnabled: async () => undefined,
		});
		const requests: ReachAssistantRequest[] = [];
		const reach = new class extends mock<IReachSubscriptionConnection>() {
			override readonly onDidChangeAccount = Event.None;
			override async sendAssistantPrompt(_account: string, request: ReachAssistantRequest): Promise<ReachAssistantResponse> {
				requests.push(request);
				if (requests.length === 1) {
					assert.match(request.messages[0].content, /Important-mail alerts already reported/);
					assert.match(request.messages[0].content, /Urgent: review/);
					return { kind: 'reply', text: '', toolCalls: [{ id: 'watch', name: 'assistant.watch_mail', arguments: '{"action":"status"}' }] };
				}
				assert.match(request.messages.at(-1)?.content ?? '', /Important-mail watch is on/);
				return { kind: 'reply', text: 'The watch is on.' };
			}
		};
		const runs = store.add(new AssistantRuns(reach, storage, emptyMcpService(), upcastPartial<IRequestService>({}), fileServiceMock(), workspaceServiceMock(), searchServiceMock(), quickInputStub, openerStub, agentChatsStub, assistantCommandStub, directMailStub, configurationStub, watcher, promptsStub));
		assert.strictEqual(await runs.send('account', 'is my important mail watch on?'), true);
		assert.strictEqual(requests.length, 2);
	});
});
