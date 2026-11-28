/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import assert from 'assert';
import { mainWindow } from '../../../../../../base/browser/window.js';
import { ICodeEditor } from '../../../../../../editor/browser/editorBrowser.js';
import { ensureNoDisposablesAreLeakedInTestSuite } from '../../../../../../base/test/common/utils.js';
import { mock } from '../../../../../../base/test/common/mock.js';
import { isIMenuItem, MenuId, MenuRegistry } from '../../../../../../platform/actions/common/actions.js';
import { CommandsRegistry } from '../../../../../../platform/commands/common/commands.js';
import { TestConfigurationService } from '../../../../../../platform/configuration/test/common/testConfigurationService.js';
import { ContextKeyService } from '../../../../../../platform/contextkey/browser/contextKeyService.js';
import { TestInstantiationService } from '../../../../../../platform/instantiation/test/common/instantiationServiceMock.js';
import { IKeybindingService } from '../../../../../../platform/keybinding/common/keybinding.js';
import { NullLogService } from '../../../../../../platform/log/common/log.js';
import { ChatSpeechToTextConnectingAction, ChatSpeechToTextPreparingAction, getDictationShortcutOperation, registerChatSpeechToTextActions, runDictationShortcut, ToggleChatSpeechToTextAction } from '../../../browser/actions/chatSpeechToTextActions.js';
import { IChatWidget, IChatWidgetService } from '../../../browser/chat.js';
import { ChatSpeechToTextState, IChatSpeechToTextService } from '../../../browser/speechToText/chatSpeechToTextService.js';
import { IDictationOnboardingService } from '../../../browser/speechToText/dictationOnboarding.js';
import { ChatContextKeys } from '../../../common/actions/chatContextKeys.js';

suite('Chat Speech to Text Actions', () => {
	const store = ensureNoDisposablesAreLeakedInTestSuite();

	test('offers the configured dictation control in the Assistant composer across model preparation states', () => {
		store.add(registerChatSpeechToTextActions());
		const configuration = new TestConfigurationService({ 'dictation.showButton': true });
		store.add(configuration.onDidChangeConfigurationEmitter);
		const context = store.add(new ContextKeyService(configuration));
		ChatContextKeys.enabled.bindTo(context).set(true);
		const configured = ChatContextKeys.speechToTextConfigured.bindTo(context);
		const preparing = ChatContextKeys.speechToTextPreparing.bindTo(context);
		const cloud = ChatContextKeys.speechToTextUsesMai.bindTo(context);
		const actions = [ToggleChatSpeechToTextAction.ID, ChatSpeechToTextPreparingAction.ID, ChatSpeechToTextConnectingAction.ID];
		const items = MenuRegistry.getMenuItems(MenuId.for('AgentSessionsTextComposerExecute')).filter(isIMenuItem).filter(item => actions.includes(item.command.id));
		const visible = () => [...new Set(items.filter(item => context.contextMatchesRules(item.when)).map(item => item.command.id))];
		const unavailable = visible();
		configured.set(true);
		const ready = visible();
		preparing.set(true);
		const loadingLocal = visible();
		cloud.set(true);
		assert.deepStrictEqual([unavailable, ready, loadingLocal, visible()], [[], [actions[0]], [actions[1]], [actions[2]]]);
	});

	test('disables the dictation action during session preparation', () => {
		const configuration = new TestConfigurationService();
		store.add(configuration.onDidChangeConfigurationEmitter);
		const contextKeyService = store.add(new ContextKeyService(configuration));
		const preparing = ChatContextKeys.transcriptProgressActive.bindTo(contextKeyService);
		const action = new ToggleChatSpeechToTextAction();
		preparing.set(true);
		const duringPreparation = contextKeyService.contextMatchesRules(action.desc.precondition);
		preparing.set(false);
		assert.deepStrictEqual({
			duringPreparation,
			afterPreparation: contextKeyService.contextMatchesRules(action.desc.precondition),
		}, { duringPreparation: false, afterPreparation: true });
	});

	test('dictation commands do not record into the disabled preparation input', async () => {
		store.add(registerChatSpeechToTextActions());
		const instantiationService = store.add(new TestInstantiationService());
		const widget = new class extends mock<IChatWidget>() {
			override readonly isTranscriptProgressActive = true;
		};
		instantiationService.stub(IChatWidgetService, new class extends mock<IChatWidgetService>() {
			override readonly lastFocusedWidget = widget;
		});
		instantiationService.stub(IChatSpeechToTextService, new class extends mock<IChatSpeechToTextService>() { });
		instantiationService.stub(IKeybindingService, new class extends mock<IKeybindingService>() { });

		for (const commandId of [ToggleChatSpeechToTextAction.ID, 'workbench.action.chat.holdToSpeechToText']) {
			const command = CommandsRegistry.getCommand(commandId)!;
			await command.handler(instantiationService);
			await command.handler(instantiationService, { widget });
		}
	});

	test('resolves the dictation toggle operation', () => {
		assert.deepStrictEqual([
			getDictationShortcutOperation(false, ChatSpeechToTextState.Idle, false),
			getDictationShortcutOperation(true, ChatSpeechToTextState.Recording, false),
			getDictationShortcutOperation(true, ChatSpeechToTextState.Recording, true),
			getDictationShortcutOperation(false, ChatSpeechToTextState.Transcribing, false),
		], [
			'start',
			'stop',
			'cancel',
			undefined,
		]);
	});

	test('starts dictation while first-run onboarding is shown', async () => {
		const calls: string[] = [];
		let state = ChatSpeechToTextState.Idle;
		const speechService = new class extends mock<IChatSpeechToTextService>() {
			override get state(): ChatSpeechToTextState { return state; }
			override readonly isPreparingModel = false;
			override readonly analyserNode = new class extends mock<AnalyserNode>() { };
			override async switchMicrophone(_window: Window & typeof globalThis, deviceId: string): Promise<AnalyserNode | undefined> {
				calls.push(`switchMicrophone:${deviceId}`);
				return this.analyserNode;
			}
		};
		const keybindingService = new class extends mock<IKeybindingService>() {
			override enableKeybindingHoldMode(): Promise<void> | undefined {
				return undefined;
			}
		};
		const onboardingService = new class extends mock<IDictationOnboardingService>() {
			override showIfNeeded(): boolean {
				calls.push('showIfNeeded');
				return true;
			}
			override refreshMicrophones(analyserNode?: AnalyserNode, switchMicrophone?: (deviceId: string) => Promise<AnalyserNode | undefined>): void {
				calls.push(`refreshMicrophones:${analyserNode === speechService.analyserNode}`);
				void switchMicrophone?.('mic-b');
			}
		};
		const editor = new class extends mock<ICodeEditor>() {
			override getDomNode(): HTMLElement {
				return mainWindow.document.body;
			}
		};

		await runDictationShortcut(
			{ speechService, keybindingService, logService: new NullLogService(), onboardingService },
			'test.dictation',
			editor,
			async () => {
				calls.push('startDictation');
				state = ChatSpeechToTextState.Recording;
			},
		);

		assert.deepStrictEqual(calls, ['showIfNeeded', 'startDictation', 'refreshMicrophones:true', 'switchMicrophone:mic-b']);
	});
});
