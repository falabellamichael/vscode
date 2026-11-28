/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import assert from 'assert';
import { $, append, getWindow } from '../../../base/browser/dom.js';
import { mainWindow } from '../../../base/browser/window.js';
import { toDisposable } from '../../../base/common/lifecycle.js';
import { ensureNoDisposablesAreLeakedInTestSuite } from '../../../base/test/common/utils.js';
import { ColorThemeData } from '../../../workbench/services/themes/common/colorThemeData.js';
import { generateColorThemeCSS } from '../../../workbench/services/themes/browser/colorThemeCss.js';
import { agentsActivityBarBackground, agentsActivityBarBorder, agentsChatInputBackground, agentsChatInputForeground, agentsDetailBackground, agentsPanelBackground, agentsPanelBorder, agentsPanelForeground } from '../../../workbench/common/agentsColors.js';
import { applySessionBarThemeColors, applySessionViewThemeColors } from '../../../workbench/contrib/sessions/browser/parts/sessionBarStyles.js';
import '../../../workbench/browser/parts/sessions/media/sessionsEmptyState.css';
import '../../../workbench/contrib/sessions/browser/chat/media/chatWidget.css';
import '../../../workbench/contrib/sessions/browser/parts/media/chatCompositeBar.css';
import '../../../workbench/contrib/sessions/browser/parts/media/sessionView.css';
import '../../../workbench/contrib/sessions/browser/parts/media/sessionsGridContent.css';
import '../../../workbench/contrib/userDataProfile/browser/media/userDataProfilesEditor.css';
import '../../../workbench/contrib/welcomeGettingStarted/browser/media/gettingStarted.css';
import '../../../workbench/contrib/welcomeWalkthrough/browser/media/walkThroughPart.css';
import '../../../workbench/contrib/workspace/browser/media/workspaceTrustEditor.css';
import '../../../workbench/contrib/preferences/browser/media/settingsEditor2.css';
import '../../../workbench/contrib/preferences/browser/media/keybindingsEditor.css';
import '../../../workbench/contrib/chat/browser/chatManagement/media/chatManagementEditor.css';
import '../../../workbench/contrib/mcp/browser/media/mcpServerEditor.css';
import '../../../workbench/contrib/extensions/browser/media/runtimeExtensionsEditor.css';
import '../../../workbench/contrib/processExplorer/browser/media/processExplorer.css';
import '../../browser/media/style.css';
import '../../browser/media/workbench.css';
import '../../contrib/editor/browser/media/emptyFileEditor.css';
import '../../contrib/sessions/browser/media/assistantSession.css';

suite('Sessions - Extension theme compatibility', () => {
	const store = ensureNoDisposablesAreLeakedInTestSuite();

	test('paints the entire central session canvas with editor colors even when panels use another palette', () => {
		const root = append(mainWindow.document.body, $('.monaco-workbench.agent-sessions-workbench'));
		store.add(toDisposable(() => root.remove()));
		const canvas = append(append(append(root, $('.part.sessionspart')), $('.content')), $('.session-view'));
		const content = append(canvas, $('.session-view-content'));
		const theme = ColorThemeData.createUnloadedTheme('vs-dark', {
			'agentsPanel.background': '#191a1b',
			'activeSessionView.background': '#191a1b',
			'inactiveSessionView.background': '#191a1b',
		});
		for (const palette of [
			{ 'editor.background': '#17324d', 'editor.foreground': '#ddeeff', 'panel.background': '#442233', 'sideBar.background': '#332244' },
			{ 'editor.background': '#fff4db', 'editor.foreground': '#202938', 'panel.background': '#dce8f4', 'sideBar.background': '#e5dcee' },
		]) {
			theme.setCustomColors(palette);
			for (const active of [true, false]) {
				applySessionViewThemeColors(canvas, theme, active);
				const painted = getWindow(canvas).getComputedStyle(canvas);
				assert.strictEqual(canvas.style.getPropertyValue('--session-view-background'), palette['editor.background']);
				assert.strictEqual(canvas.style.getPropertyValue('--session-view-foreground'), palette['editor.foreground']);
				const expected = $('div');
				expected.style.backgroundColor = palette['editor.background'];
				assert.strictEqual(painted.backgroundColor, expected.style.backgroundColor);
				assert.strictEqual(getWindow(content).getComputedStyle(content).backgroundColor, expected.style.backgroundColor);
			}
		}
		theme.setCustomColors({ 'editor.background': '#17324d', 'activeSessionView.background': '#112233', 'inactiveSessionView.background': '#334455' });
		applySessionViewThemeColors(canvas, theme, true);
		assert.strictEqual(canvas.style.getPropertyValue('--session-view-background'), '#112233');
		applySessionViewThemeColors(canvas, theme, false);
		assert.strictEqual(canvas.style.getPropertyValue('--session-view-background'), '#334455');
	});

	test('inherits conventional theme colors for chat cards, details, composer, frames and activity bar', () => {
		const theme = ColorThemeData.createUnloadedTheme('vs-dark', {});
		const pairs = [
			[agentsPanelBackground, 'panel.background'], [agentsPanelForeground, 'foreground'],
			[agentsDetailBackground, 'sideBar.background'], [agentsPanelBorder, 'panel.border'],
			[agentsActivityBarBackground, 'activityBar.background'], [agentsActivityBarBorder, 'activityBar.border'],
			[agentsChatInputBackground, 'input.background'], [agentsChatInputForeground, 'input.foreground'],
		] as const;
		for (const palette of [
			{ 'panel.background': '#1a1b26', foreground: '#c0caf5', 'sideBar.background': '#16161e', 'panel.border': '#1f2335', 'activityBar.background': '#13141c', 'activityBar.border': '#24283b', 'input.background': '#1f2335', 'input.foreground': '#a9b1d6' },
			{ 'panel.background': '#fdf6e3', foreground: '#384c54', 'sideBar.background': '#eee8d5', 'panel.border': '#93a1a1', 'activityBar.background': '#e4ddc8', 'activityBar.border': '#839496', 'input.background': '#e8e1cc', 'input.foreground': '#586e75' },
		]) {
			theme.setCustomColors(palette);
			for (const [agent, standard] of pairs) {
				assert.strictEqual(theme.getColor(agent)?.toString(), theme.getColor(standard)?.toString(), `${agent} should inherit ${standard}`);
			}
		}
		theme.setCustomColors({ 'panel.background': '#112233', 'agentsPanel.background': '#445566', 'sideBar.background': '#778899', 'agentsDetail.background': '#aabbcc' });
		assert.strictEqual(theme.getColor(agentsPanelBackground)?.toString(), '#445566');
		assert.strictEqual(theme.getColor(agentsDetailBackground)?.toString(), '#aabbcc');
	});

	test('applies opt-in Assistant canvas colors without recoloring light messages or controls', () => {
		const root = append(mainWindow.document.body, $('.monaco-workbench.agent-sessions-workbench'));
		const style = append(mainWindow.document.head, $('style'));
		store.add(toDisposable(() => { root.remove(); style.remove(); }));
		const assistant = append(root, $('.reach-assistant'));
		assistant.style.backgroundColor = 'var(--vscode-editor-background)';
		assistant.style.color = 'var(--vscode-editor-foreground)';
		const reply = append(assistant, $('.reach-assistant-reply'));
		const link = append(reply, $('a', { href: '#' }));
		const streaming = append(assistant, $('.reach-assistant-streaming'));
		const app = append(assistant, $('button.reach-assistant-app'));
		const count = append(app, $('span.reach-assistant-app-count'));
		const message = append(assistant, $('.reach-assistant-message'));
		const approval = append(assistant, $('.reach-assistant-approval'));
		const approvalControl = append(approval, $('button.reach-assistant-control'));
		const composer = append(assistant, $('.reach-assistant-composer'));
		const composerControl = append(append(composer, $('.reach-assistant-controls')), $('button.reach-assistant-control'));
		const theme = ColorThemeData.createUnloadedTheme('vs', {});
		const colors = {
			'editor.background': '#000000', 'editor.foreground': '#ffffff', foreground: '#17130f', descriptionForeground: '#595959',
			'icon.foreground': '#1f1f1f', 'chat.requestBackground': '#ffffff', 'editorWidget.background': '#ffffff',
			'textLink.foreground': '#007acc',
		};
		const read = (element: HTMLElement) => getWindow(element).getComputedStyle(element).color;
		const paint = () => {
			style.textContent = generateColorThemeCSS(theme, '.agent-sessions-workbench').code
				+ '.agent-sessions-workbench a { color: var(--vscode-textLink-foreground); }';
		};
		theme.setCustomColors(colors);
		paint();
		const normal = [reply, streaming, app, count, link].map(read);
		assert.deepStrictEqual(normal, ['rgb(23, 19, 15)', read(assistant), 'rgb(23, 19, 15)', 'rgb(89, 89, 89)', 'rgb(0, 122, 204)']);
		const lightControls = [message, approvalControl, composerControl].map(read);
		assert.deepStrictEqual(lightControls, ['rgb(23, 19, 15)', 'rgb(23, 19, 15)', 'rgb(31, 31, 31)']);
		assert.strictEqual(read(approval), 'rgb(255, 255, 255)');

		theme.setCustomColors({ ...colors, 'assistantView.foreground': '#ffffff', 'assistantView.descriptionForeground': '#b3b3b3', 'assistantView.linkForeground': '#69bdff' });
		paint();
		assert.strictEqual(getWindow(assistant).getComputedStyle(assistant).backgroundColor, 'rgb(0, 0, 0)');
		assert.deepStrictEqual([reply, streaming, app, count, link].map(read), ['rgb(255, 255, 255)', 'rgb(255, 255, 255)', 'rgb(255, 255, 255)', 'rgb(179, 179, 179)', 'rgb(105, 189, 255)']);
		assert.deepStrictEqual([message, approvalControl, composerControl].map(read), lightControls);
		assert.strictEqual(read(approval), 'rgb(23, 19, 15)');
		assert.strictEqual(getWindow(message).getComputedStyle(message).backgroundColor, 'rgb(255, 255, 255)');

		theme.setCustomColors(colors);
		paint();
		assert.deepStrictEqual([reply, streaming, app, count, link].map(read), normal);
	});

	test('keeps sidebar, menu and picker surfaces on their theme colors across palette changes', () => {
		const root = append(mainWindow.document.body, $('.monaco-workbench.agent-sessions-workbench.vs-dark.dock-detail-panel'));
		const style = append(mainWindow.document.head, $('style'));
		store.add(toDisposable(() => { root.remove(); style.remove(); }));
		const surfaces = [
			append(root, $('.part.sidebar')),
			append(root, $('.action-widget')),
			append(root, $('.context-view.monaco-menu-container')),
			append(root, $('.quick-input-widget')),
			append(append(append(root, $('.part.auxiliarybar')), $('.content')), $('.pane-body')),
		];
		const theme = ColorThemeData.createUnloadedTheme('vs-dark', {});
		const read = () => surfaces.map(element => getWindow(element).getComputedStyle(element).backgroundColor);
		for (const [colors, expected] of [
			[{ 'editor.background': '#1a1b26', 'panel.background': '#1a1b26', 'sideBar.background': '#16161e', 'menu.background': '#202132', 'quickInput.background': '#242538' }, ['rgb(22, 22, 30)', 'rgb(32, 33, 50)', 'rgb(32, 33, 50)', 'rgb(36, 37, 56)', 'rgb(22, 22, 30)']],
			[{ 'editor.background': '#fafafa', 'panel.background': '#fafafa', 'sideBar.background': '#eeeeff', 'menu.background': '#ddddff', 'quickInput.background': '#ccccff' }, ['rgb(238, 238, 255)', 'rgb(221, 221, 255)', 'rgb(221, 221, 255)', 'rgb(204, 204, 255)', 'rgb(238, 238, 255)']],
		] as const) {
			theme.setCustomColors(colors);
			style.textContent = generateColorThemeCSS(theme, '.agent-sessions-workbench').code;
			assert.deepStrictEqual(read(), expected);
		}
	});

	test('keeps Agents sidebar, session headers, empty File editor and composer canvas legible on mixed light themes', () => {
		const theme = ColorThemeData.createUnloadedTheme('vs', {});
		theme.setCustomColors({
			'editor.background': '#000000',
			'editor.foreground': '#ffffff',
			'sideBar.background': '#ffffff',
			'panel.background': '#ffffff',
			foreground: '#000000',
			'panelTitle.activeForeground': '#000000',
			'panelTitle.inactiveForeground': '#595959',
		});

		assert.strictEqual(theme.getColor('agents.background')?.toString(), '#ffffff');

		const root = append(mainWindow.document.body, $('.monaco-workbench.agent-sessions-workbench'));
		const style = append(mainWindow.document.head, $('style'));
		store.add(toDisposable(() => { root.remove(); style.remove(); }));
		style.textContent = generateColorThemeCSS(theme, '.agent-sessions-workbench').code;

		const bar = append(root, $('.chat-composite-bar'));
		applySessionBarThemeColors(bar, theme);
		assert.strictEqual(bar.style.getPropertyValue('--chat-tab-active-foreground'), '#ffffff');

		const emptyFileEditor = append(root, $('.empty-file-editor'));
		const fileActions = append(append(append(emptyFileEditor, $('.sessions-watermark-actions')), $('.monaco-action-bar')), $('.actions-container'));
		const fileEntry = append(fileActions, $('a.sessions-add-tab-watermark-entry'));

		const auxTitle = append(append(append(root, $('.part.auxiliarybar')), $('.sessions-empty-state')), $('p.sessions-empty-state-title'));

		const sessionView = append(root, $('.session-view'));
		applySessionViewThemeColors(sessionView, theme, true);
		const welcomeTitle = append(sessionView, $('h1.new-session-welcome-message-title'));

		const colorOf = (el: HTMLElement) => getWindow(el).getComputedStyle(el).color;
		assert.strictEqual(colorOf(fileEntry), 'rgb(255, 255, 255)');
		assert.strictEqual(colorOf(auxTitle), 'rgb(0, 0, 0)');
		assert.strictEqual(colorOf(welcomeTitle), 'rgb(255, 255, 255)');
	});

	test('keeps Settings, Profiles, Welcome, Keybindings and management editors legible on mixed light themes', () => {
		const theme = ColorThemeData.createUnloadedTheme('vs', {});
		theme.setCustomColors({
			'editor.background': '#000000',
			'editor.foreground': '#ffffff',
			'sideBar.background': '#ffffff',
			'panel.background': '#ffffff',
			foreground: '#000000',
		});

		const root = append(mainWindow.document.body, $('.monaco-workbench'));
		const style = append(mainWindow.document.head, $('style'));
		store.add(toDisposable(() => { root.remove(); style.remove(); }));
		style.textContent = generateColorThemeCSS(theme, '.monaco-workbench').code;

		const content = append(append(root, $('.part.editor')), $('.content'));
		const targetWindow = getWindow(root);

		for (const id of [
			'workbench.editor.userDataProfiles',
			'gettingStartedPage',
			'workbench.editor.walkThroughPart',
			'workbench.editor.workspaceTrust',
			'workbench.editor.settings2',
			'workbench.editor.keybindings',
			'workbench.editor.modelsManagement',
			'workbench.editor.mcpServer',
			'workbench.editor.runtimeExtensions',
			'workbench.editor.processExplorer',
		]) {
			const instance = append(content, $('.editor-instance'));
			instance.setAttribute('data-editor-id', id);
			const computed = targetWindow.getComputedStyle(instance);
			assert.strictEqual(computed.backgroundColor, 'rgb(255, 255, 255)', `${id} backgroundColor`);
			assert.strictEqual(computed.color, 'rgb(0, 0, 0)', `${id} color`);
			assert.strictEqual(computed.display, 'flow-root', `${id} display`);
		}

		const gettingStarted = append(content, $('.gettingStartedContainer'));
		assert.strictEqual(targetWindow.getComputedStyle(gettingStarted).backgroundColor, 'rgb(255, 255, 255)');
		assert.strictEqual(targetWindow.getComputedStyle(gettingStarted).color, 'rgb(0, 0, 0)');
	});
});
