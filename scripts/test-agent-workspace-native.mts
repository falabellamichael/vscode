/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import assert from 'node:assert/strict';
import { execFile, execFileSync } from 'node:child_process';
import { parseArgs, promisify } from 'node:util';
import { mkdir, readFile, readdir, realpath, rename, writeFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import { delimiter, dirname, join } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { _electron, type ElectronApplication, type Locator, type Page } from 'playwright';
import { allocateInspectorArgs } from './agent-workspace-test-utils.mts';

assert.equal(process.platform, 'win32', 'This smoke check uses the prepared Windows development runtime');
const require = createRequire(import.meta.url);
const { startServer, registerScenario } = require('./chat-simulation/common/mock-llm-server.ts') as typeof import('./chat-simulation/common/mock-llm-server.ts');
const { buildEnv } = require('./chat-simulation/common/utils.js') as { buildEnv(server: { url: string }): Record<string, string> };
const { createLaunchOptions } = require('./code-agent-workspace.ts') as { createLaunchOptions(root: string, workspace: string, ide: boolean, env: NodeJS.ProcessEnv, optimized?: boolean, optimizedBuild?: string): { executable: string; args: string[]; environment: NodeJS.ProcessEnv } };
const root = dirname(dirname(fileURLToPath(import.meta.url)));
const run = join(root, '.build', 'agent-workspace-native-smoke', `${new Date().toISOString().replace(/[:.]/g, '-')}-${process.pid}`);
const { values: flags } = parseArgs({
	options: {
		'empty-workspace': { type: 'boolean' },
		conversation: { type: 'boolean' },
		'running-conversation': { type: 'boolean' },
		protocol: { type: 'boolean' },
		'workspace-protocol': { type: 'boolean' },
		keyboard: { type: 'boolean' },
		'palette-switch': { type: 'boolean' },
		'app-connections': { type: 'boolean' },
		themes: { type: 'boolean' },
		'switch-accessibility': { type: 'boolean' },
		'switch-layout': { type: 'boolean' },
		'view-menu': { type: 'boolean' },
		'diff-workspace-root': { type: 'boolean' },
		'focused-diff-side': { type: 'boolean' },
		'untitled-workspace-root': { type: 'boolean' },
		'explorer-workspace-root': { type: 'boolean' },
		'explorer-folder-menu': { type: 'boolean' },
		'recent-workspace-root': { type: 'boolean' },
		'originating-editor': { type: 'boolean' },
		'reopen-saved-editor-file': { type: 'boolean' },
		'reopen-saved-editor-group': { type: 'boolean' },
		'reopen-saved-editor-preview': { type: 'boolean' },
		'preview-conflict': { type: 'boolean' },
		'disable-editor-restoration': { type: 'boolean' },
		'reopen-saved-diff': { type: 'boolean' },
		'focus-agents-from-other-editor': { type: 'boolean' },
		'warm-workspace-switch': { type: 'boolean' },
		'save-draft-on-workspace-switch': { type: 'boolean' },
		optimized: { type: 'boolean' },
		'optimized-build': { type: 'string' },
	},
});
const optimized = flags.optimized ?? false;
const optimizedBuild = flags['optimized-build'];
assert.ok(optimizedBuild === undefined || (optimized && optimizedBuild.trim()), 'Use --optimized --optimized-build <folder> to select a staged desktop build.');
const { keyboard, 'app-connections': appConnections, themes } = flags;
const paletteSwitch = flags['palette-switch'];
const switchAccessibility = flags['switch-accessibility'];
const switchLayout = flags['switch-layout'];
const viewMenu = flags['view-menu'];
const focusedDiffSide = flags['focused-diff-side'];
const diffWorkspaceRoot = flags['diff-workspace-root'] || focusedDiffSide;
const untitledWorkspaceRoot = flags['untitled-workspace-root'];
const explorerWorkspaceRoot = flags['explorer-workspace-root'];
const explorerFolderMenu = flags['explorer-folder-menu'];
const recentWorkspaceRoot = flags['recent-workspace-root'];
const selectsSecondRoot = diffWorkspaceRoot || untitledWorkspaceRoot || explorerWorkspaceRoot || explorerFolderMenu || recentWorkspaceRoot;
const themePalettes = {
	dark: {
		name: 'Dark 2026', themeClass: 'vscode-theme-defaults-themes-2026-dark-json',
		colors: {
			'editor.background': '#17324d', 'editor.foreground': '#ddeeff', 'panel.background': '#30243a',
			'sideBar.background': '#11202b', 'input.background': '#1e2f44',
			'titleBar.activeBackground': '#0a1118', 'titleBar.inactiveBackground': '#0a1118',
		},
	},
	light: {
		name: 'Light 2026', themeClass: 'vscode-theme-defaults-themes-2026-light-json',
		colors: {
			'editor.background': '#fff4db', 'editor.foreground': '#25364a', 'panel.background': '#e6edf7',
			'sideBar.background': '#e8dfee', 'input.background': '#f2e9f6',
			'titleBar.activeBackground': '#e0d6ee', 'titleBar.inactiveBackground': '#e0d6ee',
		},
	},
} as const;
let expectedTheme: typeof themePalettes[keyof typeof themePalettes] = themePalettes.dark;
const workspaceProtocol = flags['workspace-protocol'];
const protocol = workspaceProtocol || flags.protocol;
const launchToken = 'synthetic-test-launch-token';
const conversation = protocol || flags.conversation;
const runningConversation = flags['running-conversation'];
const emptyWorkspace = conversation || flags['empty-workspace'];
const originatingEditor = flags['originating-editor'];
const reopenSavedEditorGroup = flags['reopen-saved-editor-group'];
const reopenSavedEditorPreview = flags['reopen-saved-editor-preview'];
const previewConflict = flags['preview-conflict'];
const disableEditorRestoration = flags['disable-editor-restoration'];
const reopenSavedEditorFile = flags['reopen-saved-editor-file'] || reopenSavedEditorGroup || reopenSavedEditorPreview;
const reopenSavedDiff = flags['reopen-saved-diff'];
const focusAgentsFromOtherEditor = flags['focus-agents-from-other-editor'];
const warmWorkspaceSwitch = flags['warm-workspace-switch'];
const saveDraftOnWorkspaceSwitch = flags['save-draft-on-workspace-switch'];
const folderlessOrigin = originatingEditor && emptyWorkspace;
assert.ok(!themes || !emptyWorkspace, '--themes checks the multi-root workspace round trip');
assert.ok(!switchAccessibility || !emptyWorkspace, '--switch-accessibility checks both title bars in the workspace round trip');
assert.ok(!switchLayout || !emptyWorkspace, '--switch-layout checks the IDE title bar in the workspace round trip');
assert.ok(!viewMenu || !emptyWorkspace, '--view-menu opens the saved IDE workspace in Agents');
assert.ok(!paletteSwitch || (!keyboard && !explorerFolderMenu && !emptyWorkspace), '--palette-switch checks a project round trip through the command palette');
assert.ok(!runningConversation || (!emptyWorkspace && !originatingEditor && !warmWorkspaceSwitch), '--running-conversation checks the standard two-window project handoff');
assert.ok(!diffWorkspaceRoot || !emptyWorkspace, '--diff-workspace-root compares a file from the second workspace root');
assert.ok(!untitledWorkspaceRoot || (!emptyWorkspace && !diffWorkspaceRoot), '--untitled-workspace-root uses a separate multi-root test');
assert.ok(!explorerWorkspaceRoot || (!emptyWorkspace && !diffWorkspaceRoot && !untitledWorkspaceRoot), '--explorer-workspace-root uses a separate multi-root test');
assert.ok(!explorerFolderMenu || (!emptyWorkspace && !diffWorkspaceRoot && !untitledWorkspaceRoot && !explorerWorkspaceRoot && !viewMenu && !keyboard), '--explorer-folder-menu uses its own opening and activation steps');
assert.ok(!recentWorkspaceRoot || (!emptyWorkspace && !diffWorkspaceRoot && !untitledWorkspaceRoot && !explorerWorkspaceRoot && !explorerFolderMenu), '--recent-workspace-root starts with a scratch editor over a workspace tab');
assert.ok(!folderlessOrigin || (conversation && !protocol), 'Use --originating-editor with --conversation to check folderless editor routing');
assert.ok(!reopenSavedEditorFile || (!emptyWorkspace && !originatingEditor && !warmWorkspaceSwitch && !runningConversation), '--reopen-saved-editor-file uses the standard saved-workspace round trip');
assert.ok(!reopenSavedDiff || (!emptyWorkspace && !originatingEditor && !warmWorkspaceSwitch && !runningConversation && !reopenSavedEditorFile), '--reopen-saved-diff uses the standard saved-workspace round trip');
assert.ok(!disableEditorRestoration || reopenSavedEditorGroup, '--disable-editor-restoration requires --reopen-saved-editor-group');
assert.ok(!previewConflict || (reopenSavedEditorPreview && reopenSavedEditorGroup && !disableEditorRestoration), '--preview-conflict requires preview and split-group restoration');
assert.ok(!focusAgentsFromOtherEditor || (originatingEditor && !emptyWorkspace), 'Use --focus-agents-from-other-editor with --originating-editor in a multi-root workspace');
assert.ok(!warmWorkspaceSwitch || (!emptyWorkspace && !originatingEditor && !selectsSecondRoot && !themes && !viewMenu), '--warm-workspace-switch checks switching between separate projects');
assert.ok(!saveDraftOnWorkspaceSwitch || warmWorkspaceSwitch, '--save-draft-on-workspace-switch requires --warm-workspace-switch');
const portable = join(run, 'portable');
const profile = join(portable, 'user-data');
const fixtureHome = join(run, 'home');
await mkdir(fixtureHome, { recursive: true });
const folders = [join(run, 'one'), join(run, 'two')];
const agentHomes = ['codex', 'claude', 'copilot', 'state'];
for (const directory of [...folders, join(run, 'extensions'), ...agentHomes.map(name => join(run, name))]) {
	await mkdir(directory, { recursive: true });
}
for (const folder of folders) {
	await writeFile(join(folder, 'README.md'), '# Native continuity fixture\n');
	// Keep repository discovery inside the fixture instead of discovering the checkout's parent repository.
	execFileSync('git', ['init', '--quiet'], { cwd: folder, windowsHide: true, stdio: 'pipe' });
}
const workspace = join(run, 'continuity.code-workspace');
const savedWorkspaceTitle = 'Signal-C saved workspace fixture';
await writeFile(workspace, JSON.stringify({
	folders: folders.map(path => ({ path })),
	...(originatingEditor && !emptyWorkspace ? { settings: { 'window.title': savedWorkspaceTitle } } : {}),
}));
const runner = join(run, 'unused-extension-test-runner.cjs');
// Hidden-window mode requires a test location. No extension-development path is supplied, so no runner executes.
await writeFile(runner, 'exports.run = async () => { throw new Error("Unexpected extension-development test execution"); };\n');

const responseText = 'Signal-C native conversation retained.';
const runningPrompt = 'Finish this request while I switch modes. [scenario:mode-switch-running]';
const runningReply = 'The running agent completed after the switch.';
const runningResponse = Promise.withResolvers<void>();
if (conversation || runningConversation) { registerScenario('text-only', [{ content: responseText, delayMs: 0 }]); }
if (runningConversation) { registerScenario('mode-switch-running', [{ content: runningReply, delayMs: 0, waitFor: runningResponse.promise }]); }
const mock = await startServer(0, { logger: () => {}, captureRequests: true });
let app: ElectronApplication | undefined;
let deadline: ReturnType<typeof setTimeout> | undefined;
let stage = 'preparing';
const report = {
	optimized,
	executable: '',
	scenario: workspaceProtocol ? 'workspace-protocol-conversation' : protocol ? 'protocol-conversation' : conversation ? 'folderless-conversation' : emptyWorkspace ? 'empty-workspace' : runningConversation ? 'running-workspace-conversation' : 'multi-root',
	windowReuseVerified: false, dirtyBufferRetained: false, agentsDraftRetained: false,
	conversationCompleted: false, conversationRestoredInIDE: false, modelRequestCountBeforeHandoff: 0,
	runningConversationOpenedInIDE: false, runningConversationRoundTripVerified: false,
	warmProtocolConversationRestored: false, coldProtocolConversationRestored: false,
	recentWorkspaceProtocolVerified: false,
	originatingEditorVerified: false,
	singleFolderOriginVerified: false,
	closedSingleFolderOriginVerified: false,
	replacementFolderEditorPreserved: false,
	repurposedFolderEditorPreserved: false,
	reusedWindowDraftFilePreserved: false,
	focusOnlyPreservedOrigin: false,
	explicitSwitchUpdatedOrigin: false,
	warmWorkspaceSwitchVerified: false,
	warmWorkspaceDraftRetained: false,
	warmWorkspaceReturnRetained: false,
	savedDraftWorkspaceSwitchVerified: false,
	inlineDraftSwitchNoticeVerified: false,
	inlineDraftSwitchNoticeDismissed: false,
	closedOriginFallbackVerified: false,
	closedOriginFilePreferredOverAlternate: false,
	reopenedOriginWindowPreferredOverAlternate: false,
	missingOriginFileFallbackVerified: false,
	reopenedOriginWorkspaceVerified: false,
	repurposedSavedWorkspaceRestored: false,
	reopenedSavedEditorFileVerified: false,
	reopenedSavedEditorViewportVerified: false,
	reopenedSavedEditorGroupVerified: false,
	reopenedPreviewTabVerified: false,
	restoredDifferentPreviewVerified: false,
	editorRestorationPreferenceVerified: false,
	reopenedSavedEditorReplacementPreserved: false,
	reopenedSingleRootWorkspaceFileVerified: false,
	reopenedSavedDiffVerified: false,
	reopenedStickyDiffVerified: false,
	reopenedDiffWithMissingFileFallbackVerified: false,
	changedOriginFallbackVerified: false,
	launchEnvironmentExcludedFromTrace: false,
	coldCliFallbackVerified: false, coldClaudeCliVersion: '',
	keyboardActivations: [] as { stage: string; method: string; commandId: string }[],
	paletteActivations: [] as { label: string; commandId: string; legacyQuery: string }[],
	modelRequestCount: 0,
	appPickerVerified: false, payPalSandboxVerified: false, datadogRegionVerified: false, quickBooksRequirementsVerified: false, appSetupCancellationVerified: false,
	appHeaderVerified: false, googleOAuthSetupVerified: false, browseAppConnectionsVerified: false,
	windowChecks: [] as { stage: string; windows: { id: number; hidden: boolean; title: string }[] }[],
	themeChecks: [] as { stage: string; mode: string; theme: string; canvas: string[]; titlebar: string }[],
	switchSettingChangedWithoutReload: false,
	switchAccessibilityChecks: [] as { mode: string; checked: string; description: string; decorations: string[] }[],
	switchLayoutVerified: false,
	viewMenuHandoffVerified: false,
	diffWorkspaceRootVerified: false,
	focusedDiffSideVerified: false,
	untitledWorkspaceRootVerified: false,
	explorerWorkspaceRootVerified: false,
	explorerFolderMenuVerified: false,
	explorerFolderMenuPreservedDraft: false,
	recentWorkspaceRootVerified: false,
};
try {
	const settings = {
		'workbench.startupEditor': 'none', 'workbench.editor.showTabs': 'multiple',
		'chat.allowAnonymousAccess': true, 'chat.agent.enabled': true,
		'chat.agentHost.claudeAgent.enabled': false, 'chat.agentHost.codexAgent.enabled': false,
		'chat.agentHost.unsafeTestToken': 'smoketest-fake-agent-host-token',
		'github.copilot.advanced.debug.overrideProxyUrl': mock.url,
		'github.copilot.chat.githubMcpServer.enabled': false, 'github.copilot.chat.cli.legacyProvider.enabled': false,
		'chat.agentHost.githubMcpServer.enabled': false, 'chat.agentHost.computerUse.enabled': false,
		'chat.agentHost.browserUse.enabled': false, 'chat.agentHost.experimental.prewarmSessions': false,
		'chat.agentHost.experimental.autoAttachPullRequests': false,
		'agents.subscriptions.devinExecutable': join(run, 'not-installed-devin.exe'),
		'sessions.useWorktree': false, 'files.autoSave': 'off',
		...(disableEditorRestoration ? { 'workbench.editor.restoreEditors': false } : {}),
		...(reopenSavedEditorPreview ? { 'workbench.editor.enablePreview': true, 'workbench.list.openMode': 'singleClick' } : {}),
		...(focusedDiffSide || reopenSavedDiff ? { 'diffEditor.renderSideBySide': true, 'diffEditor.useInlineViewWhenSpaceIsLimited': false } : {}),
		'extensions.autoUpdate': false, 'extensions.autoCheckUpdates': false,
		...(saveDraftOnWorkspaceSwitch ? { 'sessions.drafts.enabled': true } : {}),
		...(themes ? { 'workbench.colorTheme': expectedTheme.name, 'workbench.colorCustomizations': expectedTheme.colors } : {}),
		...(appConnections ? { 'chat.customizations.marketplace.enabled': true } : {}),
		...(viewMenu ? { 'window.menuBarVisibility': 'compact', 'window.zoomLevel': 3, 'chat.titleBar.openInAgentsWindow.enabled': false } : {}),
		...(protocol ? { 'security.promptForLocalFileProtocolHandling': false, 'security.promptForRemoteFileProtocolHandling': false } : {}),
	};
	for (const file of [join(profile, 'User', 'settings.json'), join(profile, 'User', 'profiles', 'builtin', 'agents', 'settings.json')]) {
		await mkdir(dirname(file), { recursive: true });
		await writeFile(file, JSON.stringify(settings));
	}
	await mkdir(join(profile, 'User', 'globalStorage'), { recursive: true });
	await writeFile(join(profile, 'User', 'globalStorage', 'agent-host-config.json'), JSON.stringify({ githubMcpServerEnabled: false }));
	const env = buildEnv(mock);
	if (conversation || runningConversation) {
		// Only the synthetic token supplied by buildEnv is used; no account profile is copied.
		const token = JSON.parse(Buffer.from(env.VSCODE_COPILOT_CHAT_TOKEN, 'base64').toString()) as Record<string, unknown>;
		Object.assign(token, { copilot_plan: 'individual', sku: 'copilot_individual', isNoAuthUser: false });
		env.VSCODE_COPILOT_CHAT_TOKEN = Buffer.from(JSON.stringify(token)).toString('base64');
	}
	// Trace logging includes the environment of forwarded CLI requests. Keep only
	// OS launch variables and the mock credentials generated by buildEnv.
	const fixtureEnvironment = new Set([
		'PATH', 'PATHEXT', 'SYSTEMROOT', 'WINDIR', 'SYSTEMDRIVE', 'COMSPEC', 'OS',
		'PROGRAMDATA', 'PROGRAMFILES', 'PROGRAMFILES(X86)', 'PROGRAMW6432', 'TEMP', 'TMP', 'TMPDIR',
		'NODE_ENV', 'ELECTRON_ENABLE_LOGGING', 'IS_SCENARIO_AUTOMATION',
		'VSCODE_DEV', 'VSCODE_CLI', 'GITHUB_PAT', 'VSCODE_COPILOT_CHAT_TOKEN',
	]);
	for (const key of Object.keys(env)) {
		if (!fixtureEnvironment.has(key.toUpperCase())) {
			delete env[key];
		}
	}
	Object.assign(env, {
		// Portable mode also confines argv.json and disables OS protocol registration.
		VSCODE_PORTABLE: portable,
		// Keep Chromium diagnostics with the fixture when testing a staged release.
		ELECTRON_LOG_FILE: join(run, 'chromium.log'),
		// Workspace-less chat scratch directories use os.homedir(), independently of SDK storage.
		USERPROFILE: fixtureHome, HOME: fixtureHome, APPDATA: join(fixtureHome, 'AppData', 'Roaming'), LOCALAPPDATA: join(fixtureHome, 'AppData', 'Local'),
		CODEX_HOME: join(run, 'codex'), CLAUDE_CONFIG_DIR: join(run, 'claude'), COPILOT_HOME: join(run, 'copilot'), XDG_STATE_HOME: join(run, 'state'),
		COPILOT_API_URL: mock.url, COPILOT_DEBUG_GITHUB_API_URL: mock.url,
		GITHUB_COPILOT_API_TOKEN: 'smoketest-fake-agent-host-token', VSCODE_AGENT_HOST_CAPI_URL_OVERRIDE: mock.url,
		...(protocol ? { SIGNALC_TEST_LAUNCH_TOKEN: launchToken } : {}),
	});
	assert.equal(execFileSync(process.execPath, ['-p', 'require("node:os").homedir()'], { env, windowsHide: true, encoding: 'utf8' }).trim(), fixtureHome);
	const inspectArgs = await allocateInspectorArgs();
	stage = 'launching';
	const launch = createLaunchOptions(root, workspace, true, {}, optimized, optimizedBuild);
	const executable = launch.executable;
	if (optimized) {
		delete env.VSCODE_DEV;
		env.NODE_ENV = 'production';
		for (const key of ['VSCODE_AGENT_HOST_CODEX_SDK_ROOT', 'VSCODE_AGENT_HOST_CLAUDE_SDK_ROOT']) {
			if (launch.environment[key]) { env[key] = launch.environment[key]; }
		}
	}
	const applicationArgs = optimized ? [] : [root];
	const fallbackIndex = launch.args.indexOf('--agent-cli-fallback-dir');
	assert.ok(fallbackIndex >= 0, 'The native fixture requires the installed official Claude CLI fallback');
	const cliFallback = launch.args[fallbackIndex + 1];
	const pathKey = Object.keys(env).find(key => key.toLowerCase() === 'path') ?? 'PATH';
	// Startup must restore the fallback without help from the parent process's PATH.
	env[pathKey] = (env[pathKey] ?? '').split(delimiter).filter(entry => entry.toLowerCase() !== cliFallback.toLowerCase()).join(delimiter);
	const commonArgs = ['--hidden-window', '--extensionTestsPath', runner, '--enable-smoke-test-driver',
		'--user-data-dir', profile, '--shared-data-dir', join(run, 'shared'), '--extensions-dir', join(run, 'extensions'),
		'--agent-codex-home', join(run, 'codex'), '--agent-claude-config-dir', join(run, 'claude'),
		'--agent-cli-fallback-dir', cliFallback,
		'--use-inmemory-secretstorage', '--skip-welcome', '--skip-release-notes', '--disable-workspace-trust',
		'--disable-crash-reporter',
		'--disable-telemetry', '--disable-experiments', '--disable-updates', '--disable-extension=vscode.vscode-api-tests',
		...(protocol || keyboard || paletteSwitch || originatingEditor || reopenSavedEditorFile || reopenSavedDiff || warmWorkspaceSwitch ? ['--log', 'trace'] : [])];
	// A second instance must omit extensionTestsPath to forward to the isolated app.
	const forwardingArgs = commonArgs.filter(arg => arg !== '--extensionTestsPath' && arg !== runner);
	const forwardToRunningApp = async (args: string[]) => {
		await promisify(execFile)(executable, [...applicationArgs, ...forwardingArgs, ...args], { env, windowsHide: true, timeout: 45000 }).catch(async (error: NodeJS.ErrnoException & { signal?: string; stdout?: string; stderr?: string }) => {
			await writeFile(join(run, 'launch-helper-error.json'), JSON.stringify({ code: error.code, signal: error.signal, stdout: error.stdout?.replace(/https?:\/\/\S+/g, '[URL]'), stderr: error.stderr?.replace(/https?:\/\/\S+/g, '[URL]') }, null, 2));
			throw new Error('The isolated launch helper did not exit successfully');
		});
	};
	app = await _electron.launch({ executablePath: executable, args: [...applicationArgs, ...(emptyWorkspace ? ['--agents'] : [workspace]), ...commonArgs, ...inspectArgs], env, timeout: 60000 });
	const launched = app;
	deadline = setTimeout(() => { void app?.close().catch(() => {}); }, 180000);
	await (await launched.firstWindow({ timeout: 45000 })).waitForLoadState('domcontentloaded');
	const launchContext = await launched.evaluate(({ app }) => ({ executable: process.execPath, portable: process.env.VSCODE_PORTABLE, profile: app.getPath('userData') }));
	report.executable = await realpath(launchContext.executable);
	assert.equal(report.executable, await realpath(executable), 'The fixture must run the selected executable');
	assert.deepEqual({ portable: launchContext.portable, profile: launchContext.profile }, { portable, profile }, 'The fixture must keep portable configuration and editor storage under its own run directory');
	const capture = async (page: Page, name: string, target = launched) => {
		const window = await target.browserWindow(page);
		// A first capture wakes the hidden compositor, which may otherwise return its previous frame.
		await window.evaluate(async nativeWindow => { await nativeWindow.webContents.capturePage(undefined, { stayHidden: true, stayAwake: true }); });
		// DOM visibility can precede the hidden window's next compositor frame.
		await page.evaluate(async () => {
			await document.fonts.ready;
			await new Promise<void>(resolve => requestAnimationFrame(() => requestAnimationFrame(() => resolve())));
		});
		try {
			// CDP screenshots can stall after a hidden native window becomes inactive.
			const png = await window.evaluate(async nativeWindow => (await nativeWindow.webContents.capturePage(undefined, { stayHidden: true, stayAwake: true })).toPNG().toString('base64'));
			assert.ok(png.length, 'The hidden native capture must contain image data');
			await writeFile(join(run, name), Buffer.from(png, 'base64'));
		} finally {
			await window.dispose();
		}
	};
	const themeWindowTimes = new Map<Page, number>();
	const checkTheme = async (page: Page, label: string) => {
		await page.waitForFunction(({ colors, themeClass }) => {
			const root = document.querySelector('.monaco-workbench');
			if (!root?.classList.contains(themeClass)) { return false; }
			const style = getComputedStyle(root);
			return Object.entries(colors).every(([key, value]) => style.getPropertyValue(`--vscode-${key.replaceAll('.', '-')}`).trim() === value);
		}, expectedTheme, { timeout: 30000 });
		const painted = await page.evaluate(() => {
			const root = document.querySelector('.monaco-workbench')!;
			const agents = root.classList.contains('agent-sessions-workbench');
			const selector = agents ? '.session-view, .session-view > .session-view-content' : '.part.editor .monaco-editor .monaco-editor-background';
			const canvases = [...root.querySelectorAll<HTMLElement>(selector)]
				.filter(element => element.getBoundingClientRect().width > 0 && element.getBoundingClientRect().height > 0);
			return {
				mode: agents ? 'agents' : 'ide',
				canvas: canvases.map(element => getComputedStyle(element).backgroundColor),
				titlebar: getComputedStyle(document.querySelector('.part.titlebar')!).backgroundColor,
				timeOrigin: performance.timeOrigin,
			};
		});
		const rgb = (hex: string) => `rgb(${parseInt(hex.slice(1, 3), 16)}, ${parseInt(hex.slice(3, 5), 16)}, ${parseInt(hex.slice(5, 7), 16)})`;
		const { mode, canvas, titlebar, timeOrigin } = painted;
		assert.equal(timeOrigin, themeWindowTimes.get(page) ?? timeOrigin, `${label}: the ${mode} window should not reload`);
		themeWindowTimes.set(page, timeOrigin);
		assert.equal(titlebar, rgb(expectedTheme.colors['titleBar.activeBackground']), `${label}: title bar should retain the theme`);
		assert.ok(canvas.length, `${label}: expected a visible ${mode} canvas`);
		for (const background of canvas) {
			assert.equal(background, rgb(expectedTheme.colors['editor.background']), `${label}: the visible ${mode} canvas should retain the editor color`);
		}
		report.themeChecks.push({ stage: label, mode, theme: expectedTheme.name, canvas, titlebar });
	};
	const checkWindows = async (label: string) => {
		const windows = await launched.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows().map(window => ({ id: window.id, hidden: !window.isVisible(), title: window.getTitle() })));
		report.windowChecks.push({ stage: label, windows });
		assert.ok(windows.every(window => window.hidden), 'Every fixture window must remain hidden');
		if (themes) {
			for (const page of launched.windows()) { await checkTheme(page, label); }
		}
		return windows.map(window => window.id).sort((a, b) => a - b);
	};
	const checkSwitchAccessibility = async (page: Page, mode: 'Agents' | 'IDE') => {
		const control = page.getByRole('switch', { name: 'IDE mode', exact: true });
		const cdp = await page.context().newCDPSession(page);
		try {
			await page.emulateMedia({ forcedColors: 'active' });
			await control.waitFor({ state: 'visible' });
			const cue = await page.waitForFunction(() => {
				const element = document.querySelector('.part.titlebar .workspace-mode-switch');
				if (!element?.getBoundingClientRect().width || !matchMedia('(forced-colors: active)').matches) { return false; }
				const decorations = [...element.querySelectorAll('.workspace-mode-segment')].map(segment => getComputedStyle(segment).textDecorationLine);
				return decorations.length === 2 && decorations.every(Boolean) ? decorations : false;
			}, undefined, { timeout: 15000 });
			const decorations = await cue.jsonValue() as string[];
			await cue.dispose();
			assert.deepEqual(decorations, mode === 'Agents' ? ['underline', 'none'] : ['none', 'underline'],
				'Only the active mode needs a visible, color-independent indicator');

			const { root } = await cdp.send('DOM.getDocument');
			const { nodeId } = await cdp.send('DOM.querySelector', { nodeId: root.nodeId, selector: '.part.titlebar .workspace-mode-switch' });
			const { nodes } = await cdp.send('Accessibility.getPartialAXTree', { nodeId, fetchRelatives: false });
			const node = nodes[0];
			assert.equal(node.role?.value, 'switch');
			assert.equal(node.name?.value, 'IDE mode');
			const checked = String(node.properties?.find(property => property.name === 'checked')?.value.value);
			assert.equal(checked, String(mode === 'IDE'));
			const description = String(node.description?.value);
			assert.ok(description.startsWith(`${mode} is active. Switch to ${mode === 'Agents' ? 'IDE' : 'Agents'}`), description);
			assert.match(description, /Ctrl\+K Ctrl\+Shift\+A/);
			report.switchAccessibilityChecks.push({ mode, checked, description, decorations });
			await capture(page, `switch-${mode.toLowerCase()}-forced-colors.png`);
		} finally {
			await page.emulateMedia({ forcedColors: null });
			await cdp.detach();
		}
	};
	const simulateNativeFocus = async (page: Page) => {
		const window = await launched.browserWindow(page);
		// Update the main process's focus history without showing a fixture window.
		try { await window.evaluate(nativeWindow => { nativeWindow.emit('focus'); }); } finally { await window.dispose(); }
	};
	const destroyFixtureWindow = async (page: Page) => {
		const closed = page.waitForEvent('close', { timeout: 30000 });
		const window = await launched.browserWindow(page);
		try { await window.evaluate(nativeWindow => nativeWindow.destroy()); } finally { await window.dispose(); }
		await closed;
	};
	const runWorkbenchCommand = async (page: Page, label: string) => {
		await page.keyboard.press('Control+Shift+P');
		const picker = page.locator('.quick-input-widget');
		await picker.locator('input').first().fill('>' + label);
		await picker.getByRole('option').filter({ hasText: label }).first().click();
		await picker.waitFor({ state: 'hidden' });
	};
	const rendererLogs = async () => {
		const logs: string[] = [];
		for (const log of await readdir(join(profile, 'logs'), { withFileTypes: true })) {
			if (!log.isDirectory()) { continue; }
			for (const window of await readdir(join(profile, 'logs', log.name), { withFileTypes: true })) {
				if (!window.isDirectory() || !window.name.startsWith('window')) { continue; }
				logs.push(await readFile(join(profile, 'logs', log.name, window.name, 'renderer.log'), 'utf8').catch(() => ''));
			}
		}
		return logs;
	};
	const checkAppConnections = async (page: Page) => {
		stage = 'native app connections';
		// Initial account and provider UI can take focus while the workbench restores.
		await page.waitForFunction(() => !!Reflect.get(globalThis, 'driver'));
		await page.evaluate(async () => { await Reflect.get(globalThis, 'driver').whenWorkbenchRestored(); });
		const paths = [join(profile, 'User/mcp.json'), join(profile, 'User/profiles/builtin/agents/mcp.json'), join(run, 'copilot/mcp-config.json'), ...folders.flatMap(folder => [join(folder, '.mcp.json'), join(folder, '.vscode/mcp.json')])];
		const readConfigurations = () => Promise.all(paths.map(file => readFile(file, 'utf8').catch((error: NodeJS.ErrnoException) => { if (error.code === 'ENOENT') { return undefined; } throw error; })));
		const before = await readConfigurations();
		const input = page.locator('.quick-input-widget input').first();
		const openAppPicker = async () => {
			await page.getByRole('switch', { name: 'IDE mode', exact: true }).focus();
			await page.keyboard.press('Control+Shift+P');
			await input.fill('>Signal-C: Connect App');
			await page.getByRole('option').filter({ hasText: 'Signal-C: Connect App' }).waitFor();
			await input.press('Enter');
			await page.getByText('Signal-C App Connections', { exact: true }).waitFor();
		};
		await openAppPicker();
		await input.fill('PayPal');
		const payPalProduction = page.getByRole('option').filter({ has: page.getByText('PayPal (Production)', { exact: true }) });
		const payPalSandbox = page.getByRole('option').filter({ has: page.getByText('PayPal (Sandbox)', { exact: true }) });
		await Promise.all([payPalProduction.waitFor(), payPalSandbox.waitFor()]);
		await payPalSandbox.click();
		const sandboxReview = page.getByRole('dialog').filter({ hasText: 'Connect PayPal (Sandbox)' });
		await sandboxReview.waitFor();
		const sandboxDetails = await sandboxReview.innerText();
			assert.ok(sandboxDetails.includes('https://mcp.sandbox.paypal.com/sse'));
		assert.ok(sandboxDetails.includes('sandbox test account'));
		report.payPalSandboxVerified = true;
		await page.keyboard.press('Escape');
		await page.locator('.quick-input-widget').waitFor({ state: 'hidden' });
		await openAppPicker();
		await input.fill('Datadog');
		const datadogUs = page.getByRole('option').filter({ has: page.getByText('Datadog (US1)', { exact: true }) });
		const datadogEu = page.getByRole('option').filter({ has: page.getByText('Datadog (EU1)', { exact: true }) });
		await Promise.all([datadogUs.waitFor(), datadogEu.waitFor()]);
		await datadogEu.click();
		const datadogReview = page.getByRole('dialog').filter({ hasText: 'Connect Datadog (EU1)' });
		await datadogReview.waitFor();
		const datadogDetails = await datadogReview.innerText();
		assert.ok(datadogDetails.includes('https://mcp.datadoghq.eu/v1/mcp'));
		assert.ok(datadogDetails.includes('site that hosts your Datadog organization'));
		report.datadogRegionVerified = true;
		await page.keyboard.press('Escape');
		await page.locator('.quick-input-widget').waitFor({ state: 'hidden' });
		await openAppPicker();
		await input.fill('QuickBooks');
		await page.getByText('QuickBooks Online (Local Server)', { exact: true }).waitFor();
		const hosted = page.getByRole('option').filter({ has: page.getByText('QuickBooks Online (Hosted Pilot)', { exact: true }) });
		await hosted.waitFor();
		report.appPickerVerified = true;
		await hosted.click();
		const review = page.getByRole('dialog').filter({ hasText: 'Connect QuickBooks Online' });
		await review.waitFor();
		assert.ok((await review.innerText()).includes('invitation-only pilot'));
		assert.ok((await review.innerText()).includes('subscription alone'));
		assert.ok((await review.innerText()).includes('currently unavailable in Signal-C'));
		assert.equal(await review.getByRole('button', { name: 'Connect with OAuth', exact: true }).count(), 0);
		report.quickBooksRequirementsVerified = true;
		await capture(page, 'app-connections-quickbooks.png');
		await page.keyboard.press('Escape');
		await page.locator('.quick-input-widget').waitFor({ state: 'hidden' });
		stage = 'MCP page app connection entry point';
		await page.getByRole('switch', { name: 'IDE mode', exact: true }).focus();
		await page.keyboard.press('Control+Shift+P');
		await input.fill('>Customizations');
		await page.getByRole('option').filter({ has: page.getByText('Customizations', { exact: true }) }).first().waitFor();
		await input.press('Enter');
		await page.getByText('MCP Servers', { exact: true }).first().click();
		const editor = page.locator('.ai-customization-management-editor');
		await editor.waitFor();
		const connectApp = editor.getByRole('button', { name: 'Connect App', exact: true });
		await connectApp.waitFor();
		await capture(page, 'app-connections-mcp-page.png');
		await connectApp.focus();
		await connectApp.press('Enter');
		await page.getByText('Signal-C App Connections', { exact: true }).waitFor();
		report.appHeaderVerified = true;
		await input.fill('Gmail');
		await page.getByRole('option').filter({ has: page.getByText('Gmail', { exact: true }) }).waitFor();
		await input.press('Enter');
		const googleReview = page.getByRole('dialog').filter({ hasText: 'Connect Gmail' });
		await googleReview.waitFor();
		assert.ok((await googleReview.innerText()).includes('Google Cloud project'));
		assert.ok((await googleReview.innerText()).includes('AI-powered agent'));
		await capture(page, 'app-connections-google-oauth.png');
		await googleReview.getByRole('button', { name: 'Connect with OAuth', exact: true }).press('Enter');
		await page.getByText('Gmail OAuth App Client ID', { exact: true }).waitFor();
		await input.press('Escape');
		await page.locator('.quick-input-widget').waitFor({ state: 'hidden' });
		assert.ok(await connectApp.evaluate(element => element.ownerDocument.activeElement === element), 'Cancelled app setup must restore focus to its MCP header button');
		report.googleOAuthSetupVerified = true;
		stage = 'Browse customizations app connections';
		await editor.locator('button.sidebar-home-button').getByText('Browse', { exact: true }).click();
		const gallery = editor.locator('.customization-discovery');
		await gallery.getByText('Browse customizations', { exact: true }).waitFor();
		await gallery.getByRole('button', { name: 'View all customizations from Signal-C App Connections', exact: true }).waitFor();
		await capture(page, 'browse-customizations-app-connections.png');
		await gallery.getByRole('button', { name: /^(?:Customization source|Customization marketplace|Marketplace): All (?:Sources|Marketplaces)$/, exact: true }).click();
		await page.getByRole('menu').getByText('Signal-C App Connections', { exact: true }).click();
		await gallery.locator('.customization-discovery-search').click({ position: { x: 80, y: 12 } });
		await page.keyboard.type('Datadog EU1');
		const connectDatadog = gallery.getByRole('button', { name: 'Connect Datadog (EU1)', exact: true });
		await connectDatadog.press('Enter');
		await datadogReview.waitFor();
		assert.ok((await datadogReview.innerText()).includes('https://mcp.datadoghq.eu/v1/mcp'));
		await page.keyboard.press('Escape');
		await datadogReview.waitFor({ state: 'hidden' });
		assert.ok(await connectDatadog.evaluate(element => element.ownerDocument.activeElement === element), 'Cancelling the setup review must return focus to its gallery card');
		await gallery.locator('.customization-discovery-search').click({ position: { x: 80, y: 12 } });
		await page.keyboard.press('Control+a');
		await page.keyboard.type('Gmail');
		const galleryGmail = gallery.getByRole('button', { name: 'Connect Gmail', exact: true });
		await galleryGmail.press('Enter');
		await googleReview.waitFor();
		await googleReview.getByRole('button', { name: 'Connect with OAuth', exact: true }).press('Enter');
		await page.getByText('Gmail OAuth App Client ID', { exact: true }).waitFor();
		await input.press('Escape');
		await page.locator('.quick-input-widget').waitFor({ state: 'hidden' });
		assert.ok(await galleryGmail.evaluate(element => element.ownerDocument.activeElement === element), 'Cancelling client registration must return focus to its gallery card');
		report.browseAppConnectionsVerified = true;
		await page.keyboard.press('Control+w');
		await editor.waitFor({ state: 'hidden' });
		assert.deepEqual(await readConfigurations(), before, 'Cancelled app setup must not modify any fixture MCP configuration');
		report.appSetupCancellationVerified = true;
		await checkWindows(stage);
	};
	const activateSwitch = async (page: Page, roundTrip: number, sourceInput: Locator) => {
		const control = page.getByRole('switch', { name: 'IDE mode', exact: true });
		const usePalette = paletteSwitch && roundTrip === 1;
		if (!keyboard && !usePalette) { await control.click(); return; }
		const toAgents = await control.getAttribute('aria-checked') === 'true';
		const commandId = toAgents ? 'workbench.action.openWorkspaceInAgentsWindow' : 'agents.openSessionInVSCode';
		const countCommands = async () => (await rendererLogs()).reduce((total, log) => total + log.split(/\r?\n/).filter(line => line.endsWith(`CommandService#executeCommand ${commandId}`)).length, 0);
		const before = await countCommands();
		const paletteLabel = toAgents ? 'Switch to Agents' : 'Switch to IDE';
		const legacyQuery = toAgents ? 'Open in Agents' : 'Open in Editor';
		if (usePalette) {
			await page.keyboard.press('Control+Shift+P');
			const picker = page.locator('.quick-input-widget');
			const search = picker.locator('input').first();
			const result = picker.getByRole('option').filter({ hasText: paletteLabel });
			await search.fill('>' + legacyQuery);
			await result.waitFor({ timeout: 10000 });
			await search.fill('>' + paletteLabel);
			await result.click({ timeout: 10000 });
		} else if (roundTrip === 1) {
			await sourceInput.click();
			await page.keyboard.press('Control+k');
			await page.keyboard.press('Control+Shift+a');
		} else {
			await control.focus();
			assert.ok(await control.evaluate(element => element.ownerDocument.activeElement === element), 'Keyboard activation must focus the production switch');
			await page.keyboard.press(toAgents ? 'Enter' : 'Space');
		}
		const timeout = Date.now() + 15000;
		while (await countCommands() === before && Date.now() < timeout) { await new Promise(resolve => setTimeout(resolve, 100)); }
		assert.equal(await countCommands(), before + 1, 'Each activation must dispatch exactly one mode command');
		if (usePalette) {
			report.paletteActivations.push({ label: paletteLabel, commandId, legacyQuery });
		} else {
			report.keyboardActivations.push({ stage, method: roundTrip === 1 ? 'global-shortcut' : toAgents ? 'Enter' : 'Space', commandId });
		}
	};
	const exerciseNarrowSwitch = async (page: Page, activate: () => Promise<void>) => {
		const control = page.getByRole('switch', { name: 'IDE mode', exact: true });
		const window = await launched.browserWindow(page);
		const bounds = await window.evaluate(native => native.getBounds());
		const titleActions = page.locator('.part.titlebar .action-toolbar-container .action-item');
		const expandedActionCount = await titleActions.count();
		const checkCompactSwitch = async (width: number, zoom: number) => {
			await page.waitForFunction(expectedWidth => Math.abs(innerWidth - expectedWidth) <= 1, width / 1.2 ** zoom, { timeout: 15000 });
			await control.locator('.workspace-mode-icon').waitFor({ timeout: 5000 });
			assert.equal(await control.getAttribute('aria-checked'), 'true');
			const compact = await control.evaluate(element => {
				const { x, y, width, height } = element.getBoundingClientRect();
				const titlebar = element.closest('.titlebar-container')!;
				return {
					focused: document.activeElement === element,
					pointerTarget: element.contains(document.elementFromPoint(x + width / 2, y + height / 2)),
					labelsHidden: [...element.querySelectorAll('.workspace-mode-segment')].every(label => getComputedStyle(label).display === 'none'),
					titlebarFits: titlebar.scrollWidth <= titlebar.clientWidth,
				};
			});
			assert.deepEqual(compact, { focused: true, pointerTarget: true, labelsHidden: true, titlebarFits: true }, `The compact IDE switch must stay operable at ${width}px and zoom ${zoom}`);
		};
		try {
			await control.focus();
			await window.evaluate((native, height) => native.setSize(500, height), bounds.height);
			for (const zoom of [1, 2, 3]) {
				await page.keyboard.press('Control+=');
				await checkCompactSwitch(500, zoom);
			}
			await capture(page, 'ide-switch-narrow.png');
			const minimumWidth = await window.evaluate((native, height) => {
				const [width] = native.getMinimumSize();
				native.setSize(width, height);
				return width;
			}, bounds.height);
			await checkCompactSwitch(minimumWidth, 3);
			await capture(page, 'ide-switch-minimum-width.png');
			const overflow = page.locator('.part.titlebar .action-toolbar-container').getByRole('button', { name: 'More Actions...', exact: true });
			await overflow.click({ timeout: 5000 });
			const menu = page.locator('.context-view').getByRole('menu').filter({ hasText: 'Primary Side Bar' });
			await menu.waitFor({ timeout: 5000 });
			await page.keyboard.press('Escape');
			await menu.waitFor({ state: 'hidden' });
			await activate();
			await window.evaluate((native, originalBounds) => native.setBounds(originalBounds), bounds);
			await control.getByText('Agents', { exact: true }).waitFor();
			await control.getByText('IDE', { exact: true }).waitFor();
			assert.equal(await titleActions.count(), expandedActionCount, 'Widening restores the title actions');
			report.switchLayoutVerified = true;
		} finally {
			try {
				await page.keyboard.press('Control+Numpad0');
				await window.evaluate((native, originalBounds) => native.setBounds(originalBounds), bounds);
			} finally {
				await window.dispose();
			}
		}
	};
	const selectMockModel = async (agents: Page) => {
		const modelButton = agents.locator('.new-chat-widget-container').getByText('Auto', { exact: true });
		await modelButton.waitFor({ timeout: 60000 });
		const catalogDeadline = Date.now() + 60000;
		let catalogReady = false;
		while (!catalogReady && Date.now() < catalogDeadline) {
			// Read only this fixture's logs; the SDK publishes its catalog asynchronously.
			const dirs = await readdir(join(profile, 'logs'), { withFileTypes: true });
			for (const dir of dirs.filter(entry => entry.isDirectory())) {
				const log = await readFile(join(profile, 'logs', dir.name, 'agenthost.log'), 'utf8').catch(() => '');
				catalogReady ||= /Found \d+ models:.*GPT-5\.3-Codex \(Mock\)/.test(log);
			}
			if (!catalogReady) { await new Promise(resolve => setTimeout(resolve, 250)); }
		}
		assert.ok(catalogReady, 'The native host must discover the mock model catalog');
		await modelButton.click();
		const model = agents.locator('.chat-model-picker-widget').getByText('GPT-5.3-Codex (Mock)', { exact: true });
		await model.waitFor({ timeout: 30000 });
		assert.equal(await model.evaluate(element => !!element.closest('[aria-disabled="true"], [disabled], .disabled')), false);
		await model.click();
	};
	if (emptyWorkspace) {
		stage = 'empty Agents';
		const agents = await launched.firstWindow({ timeout: 30000 });
		await agents.locator('.agent-sessions-workbench').waitFor({ timeout: 60000 });
		await agents.getByRole('switch', { name: 'IDE mode', exact: true }).waitFor();
		if (appConnections) { await checkAppConnections(agents); }
		let input = agents.locator('.new-chat-widget-container .interactive-input-part .monaco-editor');
		const draft = 'Signal-C folderless draft preserved across mode switches';
		await input.click(); await agents.keyboard.type(draft);
		if (conversation) {
			stage = 'completing the native mock conversation';
			await selectMockModel(agents);
			stage = 'selecting a composer output type';
			const outputList = agents.locator('.sessions-output-type-picker-list');
			const assertOutputBesidePlus = async (composer: Locator, icon: string) => {
				const indicator = composer.locator(`.chat-output-type-indicator.${icon}`);
				await indicator.waitFor();
				const [plus, selected] = await Promise.all([
					composer.locator('.chat-input-toolbars .action-label.codicon-add-compact').boundingBox(),
					indicator.boundingBox(),
				]);
				assert.ok(plus && selected && selected.x >= plus.x + plus.width && selected.x - plus.x - plus.width <= 8 && Math.abs(plus.y - selected.y) <= 4,
					'The selected output icon must sit immediately to the right of + on the same row');
			};
			const newComposer = agents.locator('.new-chat-widget-container .interactive-input-part');
			await newComposer.locator('.chat-input-toolbars .action-label.codicon-add-compact').click();
			await agents.getByText('Output Type...', { exact: true }).click();
			await outputList.getByText('Website', { exact: true }).first().click();
			await assertOutputBesidePlus(newComposer, 'codicon-globe');
			await input.click(); await agents.keyboard.press('Enter');
			await agents.locator('.interactive-session').getByText(responseText, { exact: false }).last().waitFor({ timeout: 45000 });
			report.conversationCompleted = true;
			report.modelRequestCountBeforeHandoff = mock.getRequests().length;
			assert.ok(report.modelRequestCountBeforeHandoff > 0, 'The native conversation must reach the local model endpoint');
			const outputRequest = mock.getRequests().find(request => JSON.stringify(request).includes(draft));
			assert.ok(outputRequest && JSON.stringify(outputRequest).includes('website'), 'The selected output type must reach the model request');
			assert.ok(JSON.stringify(outputRequest).includes('<requested_output>'), 'The session output contribution must steer the model request through host instructions');
			const ongoingComposer = agents.locator('.chat-view-chat .interactive-input-part').last();
			const outputIndicator = ongoingComposer.locator('.chat-output-type-indicator');
			await assertOutputBesidePlus(ongoingComposer, 'codicon-globe');
			await capture(agents, 'output-type-website.png');
			await outputIndicator.click();
			await outputList.locator('.monaco-list-row', { hasText: 'Website' }).filter({ has: agents.locator('.codicon-check') }).waitFor();
			await outputList.getByText('Document', { exact: true }).first().click();
			await assertOutputBesidePlus(ongoingComposer, 'codicon-file-text');
			await outputIndicator.press('Enter');
			await outputList.getByText('Output (Auto-Select)', { exact: true }).first().click();
			await outputIndicator.waitFor({ state: 'hidden' });
			input = agents.locator('.interactive-input-part .monaco-editor').last();
			await input.click(); await agents.keyboard.type(draft);
			await capture(agents, 'native-conversation-completed.png');
			console.log('Completed the native conversation through the local mock endpoint');
		}
		stage = 'checking the Assistant view activity bar';
		const assistantEntry = agents.locator('.session-section', { hasText: 'Assistant' });
		await assistantEntry.first().click();
		const assistantBar = agents.locator('.custom-view-activity-bar');
		await assistantBar.waitFor();
		const browserTab = assistantBar.locator('.action-label.codicon-globe');
		const terminalTab = assistantBar.locator('.action-label.codicon-terminal');
		await Promise.all([browserTab.waitFor(), terminalTab.waitFor()]);
		const [gridBox, barBox] = await Promise.all([
			agents.locator('.part.customviewgridpart').boundingBox(),
			assistantBar.boundingBox(),
		]);
		assert.ok(gridBox && barBox && barBox.x >= gridBox.x && barBox.x + barBox.width >= gridBox.x + gridBox.width - 8,
			'The Assistant view activity bar must dock at the right edge of the custom view');
		await browserTab.click();
		await agents.locator('.part.editor .browser-root').waitFor();
		assert.ok(await agents.locator('.part.customviewgridpart').isVisible(),
			'The Assistant view must stay open beside the browser editor');
		await terminalTab.click();
		await agents.locator('.part.panel').waitFor();
		assert.ok(await agents.locator('.part.customviewgridpart').isVisible(),
			'The Assistant view must stay open beside the terminal panel');
		await capture(agents, 'assistant-activity-bar.png');
		if (conversation) {
			await agents.locator('.monaco-list-row', { hasText: 'Signal-C folderless draft' }).first().click();
			await agents.locator('.part.customviewgridpart').waitFor({ state: 'hidden' });
		}
		assert.equal((await checkWindows(stage)).length, 1);
		stage = 'first empty IDE';
		let [ide] = await Promise.all([
			launched.waitForEvent('window', { timeout: 45000 }),
			agents.getByRole('switch', { name: 'IDE mode', exact: true }).click(),
		]);
		await ide.getByRole('switch', { name: 'IDE mode', exact: true }).waitFor({ timeout: 60000 });
		if (conversation) {
			await ide.locator('.interactive-session').getByText(responseText, { exact: false }).last().waitFor({ timeout: 45000 });
			report.conversationRestoredInIDE = true;
			console.log('Restored the native conversation in the folderless IDE');
		}
		let protocolUri: string | undefined;
		let deliverProtocol: ((uri: string) => Promise<void>) | undefined;
		if (protocol) {
			stage = 'warm editor protocol';
			let chatResource: string | undefined;
			for (const renderer of await rendererLogs()) {
				for (const match of renderer.matchAll(/vscode:openChatSession\s+(\S+)/g)) { chatResource = match[1]; }
			}
			assert.ok(chatResource, 'The first native handoff must record its conversation identity');
			const scheme = await launched.evaluate(() => (globalThis as typeof globalThis & { _VSCODE_PRODUCT_JSON: { urlProtocol: string } })._VSCODE_PRODUCT_JSON.urlProtocol);
			const link = new URL(`${scheme}://editor-window/open`);
			link.searchParams.set('session', chatResource);
			if (workspaceProtocol) { link.searchParams.set('folders', JSON.stringify(folders.map(folder => pathToFileURL(folder).toString()))); }
			protocolUri = link.toString();
			const closed = ide.waitForEvent('close', { timeout: 30000 });
			const window = await launched.browserWindow(ide);
			try { await window.evaluate(nativeWindow => nativeWindow.close()); } finally { await window.dispose(); }
			await closed;
			// Direct CLI delivery exercises this fixture's URL receiver without OS protocol registration.
			deliverProtocol = uri => forwardToRunningApp(['--open-url', '--', uri]);
			[ide] = await Promise.all([
				launched.waitForEvent('window', { timeout: 45000 }),
				deliverProtocol(protocolUri),
			]);
			await ide.getByRole('switch', { name: 'IDE mode', exact: true }).waitFor({ timeout: 60000 });
			await ide.locator('.interactive-session').getByText(responseText, { exact: false }).last().waitFor({ timeout: 60000 });
			if (workspaceProtocol) {
				await ide.getByRole('tree').getByText('one', { exact: true }).waitFor();
				await ide.getByRole('tree').getByText('two', { exact: true }).waitFor();
			}
			report.warmProtocolConversationRestored = true;
			await capture(ide, 'protocol-warm-editor.png');
			console.log('Restored the conversation through the running editor protocol receiver');
			if (workspaceProtocol) {
				stage = 'workspace protocol with identical roots';
				const alternateWorkspace = join(run, 'Alternate.code-workspace');
				await writeFile(alternateWorkspace, JSON.stringify({
					folders: folders.map(path => ({ path })),
					settings: { 'window.title': 'Alternate workspace' },
				}));
				const [alternate] = await Promise.all([
					launched.waitForEvent('window', { timeout: 45000 }),
					forwardToRunningApp(['--new-window', alternateWorkspace]),
				]);
				await alternate.getByRole('tree').getByText('two', { exact: true }).waitFor();
				assert.ok(await alternate.evaluate(token => Reflect.get(globalThis, 'vscode').process.env.SIGNALC_TEST_LAUNCH_TOKEN === token, launchToken));
				const expectedWindows = await checkWindows(stage);
				const alternateWindow = await launched.browserWindow(alternate);
				try {
					// Simulate the alternate editor becoming active without showing a fixture window.
					await alternateWindow.evaluate(nativeWindow => { nativeWindow.emit('focus'); });
					await deliverProtocol(protocolUri);
					await alternate.locator('.interactive-session').getByText(responseText, { exact: false }).last().waitFor({ timeout: 12000 });
					assert.deepEqual(await checkWindows(stage), expectedWindows);
					report.recentWorkspaceProtocolVerified = true;
					const alternateClosed = alternate.waitForEvent('close', { timeout: 30000 });
					await alternateWindow.evaluate(nativeWindow => nativeWindow.close());
					await alternateClosed;
				} finally { await alternateWindow.dispose(); }
			}
		}
		let alternateEditor: Page | undefined;
		if (folderlessOrigin) {
			stage = 'opening a second folderless editor';
			[alternateEditor] = await Promise.all([
				launched.waitForEvent('window', { timeout: 45000 }),
				forwardToRunningApp(['--new-window']),
			]);
			await alternateEditor.getByRole('switch', { name: 'IDE mode', exact: true }).waitFor();
			for (const editor of [ide, alternateEditor]) {
				await editor.evaluate(() => {
					Reflect.set(globalThis, 'fixtureChatHandoffs', 0);
					Reflect.get(globalThis, 'vscode').ipcRenderer.on('vscode:openChatSession', () => {
						Reflect.set(globalThis, 'fixtureChatHandoffs', Reflect.get(globalThis, 'fixtureChatHandoffs') + 1);
					});
				});
			}
		}
		const expectChatHandoff = async (editors: Page[], target: Page, switchToIDE: () => Promise<void>) => {
			const counts = () => Promise.all(editors.map(editor => editor.evaluate(() => Number(Reflect.get(globalThis, 'fixtureChatHandoffs')))));
			const before = await counts();
			await switchToIDE();
			const until = Date.now() + 12000;
			let after = await counts();
			while (after.every((count, index) => count === before[index]) && Date.now() < until) {
				await new Promise(resolve => setTimeout(resolve, 100));
				after = await counts();
			}
			assert.deepEqual(after, before.map((count, index) => count + Number(editors[index] === target)), 'The conversation must go to the selected editor only');
		};
		const expectedIds = await checkWindows(stage);
		assert.equal(expectedIds.length, alternateEditor ? 3 : 2);
		const marker = '// Signal-C empty editor unsaved buffer';
		await ide.locator('body').click({ position: { x: 600, y: 400 } });
		await ide.keyboard.press('Control+n');
		await ide.locator('.monaco-editor').first().waitFor({ state: 'visible' });
		await ide.keyboard.type(marker);
		await ide.locator('.tabs-container .tab.active.dirty').waitFor();
		if (workspaceProtocol) {
			assert.ok(protocolUri && deliverProtocol);
			const reordered = new URL(protocolUri);
			reordered.searchParams.set('folders', JSON.stringify([...folders].reverse().map(folder => pathToFileURL(folder).toString())));
			await deliverProtocol(reordered.toString());
			await ide.locator('.interactive-session').getByText(responseText, { exact: false }).last().waitFor();
			assert.deepEqual(await checkWindows('reordered workspace protocol'), expectedIds, 'Reordered roots must reuse the same workspace editor');
			await ide.getByRole('code').getByText(marker, { exact: false }).waitFor();
			await ide.locator('.tabs-container .tab.active.dirty').waitFor();
		}
		for (let roundTrip = 1; roundTrip <= 2; roundTrip++) {
			stage = `empty return to Agents ${roundTrip}`;
			await activateSwitch(ide, roundTrip, ide.getByRole('code').getByText(marker, { exact: false }));
			await ide.locator('[role="switch"][aria-label="IDE mode"][aria-busy="false"]').waitFor({ timeout: 30000 });
			assert.deepEqual(await checkWindows(stage), expectedIds);
			if (conversation) { await agents.locator('.interactive-session').getByText(responseText, { exact: false }).last().waitFor(); }
			assert.ok((await input.innerText()).replaceAll('\u00a0', ' ').includes(draft));
			report.agentsDraftRetained = true;
			stage = `empty return to IDE ${roundTrip}`;
			if (alternateEditor) {
				await simulateNativeFocus(alternateEditor);
				await expectChatHandoff([ide, alternateEditor], ide, () => activateSwitch(agents, roundTrip, input));
				report.originatingEditorVerified = true;
			} else {
				await activateSwitch(agents, roundTrip, input);
			}
			await agents.locator('[role="switch"][aria-label="IDE mode"][aria-busy="false"]').waitFor({ timeout: 30000 });
			assert.deepEqual(await checkWindows(stage), expectedIds, 'A folderless handoff must reuse the existing editor');
			if (conversation) { await ide.locator('.interactive-session').getByText(responseText, { exact: false }).last().waitFor(); }
			await ide.getByRole('code').getByText(marker, { exact: false }).waitFor();
			await ide.locator('.tabs-container .tab.active.dirty').waitFor();
			report.windowReuseVerified = true; report.dirtyBufferRetained = true;
		}
		await capture(ide, 'empty-ide-returned.png');
		await capture(agents, 'empty-agents-retained.png');
		if (alternateEditor) {
			stage = 'return after the originating folderless editor closed';
			await destroyFixtureWindow(ide);
			await expectChatHandoff([alternateEditor], alternateEditor, () => agents.getByRole('switch', { name: 'IDE mode', exact: true }).click());
			assert.equal((await checkWindows(stage)).length, 2);
			report.closedOriginFallbackVerified = true;
		}
		if (conversation) { assert.equal(mock.getRequests().length, report.modelRequestCountBeforeHandoff, 'Restoring the conversation must not send another model request'); }
		console.log('Two folderless round trips retained the editor window, dirty buffer, and Agents draft');
		if (protocol) {
			stage = 'cold editor protocol';
			assert.ok(protocolUri);
			await launched.close();
			const coldEnv = { ...env };
			delete coldEnv.CODEX_HOME; delete coldEnv.CLAUDE_CONFIG_DIR;
			app = await _electron.launch({ executablePath: executable, args: [...applicationArgs, '--open-url', ...commonArgs, ...await allocateInspectorArgs(), '--', protocolUri], env: coldEnv, timeout: 60000 });
			const coldApp = app;
			await coldApp.firstWindow({ timeout: 45000 });
			assert.deepEqual(await coldApp.evaluate(() => ({ codex: process.env.CODEX_HOME, claude: process.env.CLAUDE_CONFIG_DIR })), { codex: join(run, 'codex'), claude: join(run, 'claude') }, 'Cold URL startup must restore explicit agent homes without launcher environment variables');
			assert.ok(await coldApp.evaluate((_electron, directory) => (process.env.PATH ?? '').split(';').some(entry => entry.toLowerCase() === directory.toLowerCase()), cliFallback), 'Cold URL startup must restore the explicit CLI fallback without the launcher PATH');
			report.coldClaudeCliVersion = await coldApp.evaluate(() => new Promise<string>((resolve, reject) => {
				process.getBuiltinModule('child_process').execFile('claude.exe', ['--version'], { windowsHide: true, timeout: 15000 }, (error, stdout) => error ? reject(error) : resolve(stdout.trim()));
			}));
			assert.match(report.coldClaudeCliVersion, /Claude Code/);
			report.coldCliFallbackVerified = true;
			let coldIDE: Page | undefined;
			const restored = Date.now() + 60000;
			while (!coldIDE && Date.now() < restored) {
				for (const page of coldApp.windows()) {
					if (await page.locator('.interactive-session').getByText(responseText, { exact: false }).last().isVisible()) { coldIDE = page; break; }
				}
				if (!coldIDE) { await new Promise(resolve => setTimeout(resolve, 200)); }
			}
			assert.ok(coldIDE, 'Cold protocol startup must restore the conversation in one editor');
			await coldIDE.getByRole('switch', { name: 'IDE mode', exact: true }).waitFor({ timeout: 60000 });
			assert.equal(await coldIDE.getByRole('switch', { name: 'IDE mode', exact: true }).getAttribute('aria-checked'), 'true');
			await coldIDE.locator('.interactive-session').getByText(responseText, { exact: false }).last().waitFor({ timeout: 60000 });
			const windows = await coldApp.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows().map(window => ({ id: window.id, hidden: !window.isVisible(), title: window.getTitle() })));
			report.windowChecks.push({ stage, windows });
			assert.equal(windows.length, workspaceProtocol ? 1 : 2, 'Cold startup must recover its dirty editor and reuse a matching workspace');
			assert.ok(windows.every(window => window.hidden));
			if (workspaceProtocol) {
				await coldIDE.getByRole('tree').getByText('one', { exact: true }).waitFor();
				await coldIDE.getByRole('tree').getByText('two', { exact: true }).waitFor();
			}
			let conversationWindows = 0;
			let recoveredBuffer = false;
			for (const page of coldApp.windows()) {
				if (await page.locator('.interactive-session').getByText(responseText, { exact: false }).last().isVisible()) { conversationWindows++; }
				if (await page.getByRole('code').getByText(marker, { exact: false }).isVisible()) {
					await page.locator('.tabs-container .tab.active.dirty').waitFor();
					recoveredBuffer = true;
				}
			}
			assert.equal(conversationWindows, 1, 'Only the selected editor receives the external conversation');
			assert.ok(recoveredBuffer, 'Cold startup must recover the unsaved editor buffer');
			assert.equal(mock.getRequests().length, report.modelRequestCountBeforeHandoff, 'Cold restoration must not repeat the prompt');
			report.coldProtocolConversationRestored = true;
			await capture(coldIDE, 'protocol-cold-editor.png', coldApp);
			console.log('Restored the persisted conversation through cold editor protocol startup');
		}

	} else {
		const ide = await launched.firstWindow({ timeout: 30000 });
		stage = 'native IDE';
		await ide.getByRole('tree').getByText('one', { exact: true }).waitFor({ timeout: 60000 });
		await ide.getByRole('tree').getByText('two', { exact: true }).waitFor();
		if (appConnections) { await checkAppConnections(ide); }
		const marker = '// Signal-C native unsaved buffer';
		const untitledName = 'new-note.txt';
		const untitledPath = untitledWorkspaceRoot || explorerFolderMenu ? join(folders[untitledWorkspaceRoot ? 1 : 0], untitledName) : undefined;
		const createDirtyEditor = async () => {
			if (untitledPath) {
				stage = 'opening an unsaved file with a destination in a workspace root';
				await forwardToRunningApp(['--reuse-window', untitledPath]);
				await ide.locator('.tabs-container .tab.active').getByText(untitledName, { exact: true }).waitFor();
			} else {
				await ide.locator('body').click({ position: { x: 600, y: 400 } });
				await ide.keyboard.press('Control+n');
			}
			await ide.locator('.part.editor .monaco-editor').first().waitFor({ state: 'visible' });
			await ide.keyboard.type(marker);
			await ide.getByRole('code').getByText(marker, { exact: false }).waitFor();
			await ide.locator('.tabs-container .tab.active.dirty').waitFor();
			console.log('Created a dirty buffer in the hidden native multi-root editor');
		};
		const selectedExplorerRoot = ide.locator('.explorer-folders-view [role="treeitem"][aria-selected="true"]').getByText('two', { exact: true });
		if (recentWorkspaceRoot) {
			stage = 'opening a second-root file before the untitled scratch buffer';
			await forwardToRunningApp(['--reuse-window', join(folders[1], 'README.md')]);
			await ide.locator('.tabs-container .tab.active').getByText('README.md', { exact: true }).waitFor();
		}
		if (explorerWorkspaceRoot) {
			stage = 'selecting the second root in Explorer with no active editor';
			await ide.getByRole('tree').getByText('two', { exact: true }).click();
			await selectedExplorerRoot.waitFor();
			assert.equal(await ide.locator('.tabs-container .tab.active').count(), 0);
		} else {
			await createDirtyEditor();
		}
		if (recentWorkspaceRoot) {
			await ide.locator('.tabs-container .tab:not(.active)').getByText('README.md', { exact: true }).waitFor();
		}
		await checkWindows('IDE ready');
		if (diffWorkspaceRoot) {
			stage = focusedDiffSide ? 'focusing the original side of a diff across project roots' : 'opening a diff between the second root and an external file';
			const comparisonFile = focusedDiffSide ? join(folders[0], 'comparison.txt') : join(run, 'outside.txt');
			await writeFile(comparisonFile, 'Modified file for the diff fixture\n');
			await forwardToRunningApp(['--reuse-window', '--diff', join(folders[1], 'README.md'), comparisonFile]);
			const diff = ide.locator('.part.editor .monaco-diff-editor');
			await diff.waitFor();
			await ide.locator('.tabs-container .tab.active').getByText(/(comparison|outside)\.txt/).waitFor();
			if (focusedDiffSide) {
				const original = diff.locator('.editor.original .view-line').first();
				await original.click({ position: { x: 8, y: 8 } });
				await ide.keyboard.press('Control+Home');
				await ide.keyboard.press('Shift+End');
				await diff.locator('.editor.original .monaco-editor.focused').waitFor();
			}
		}
		stage = 'native Agents';
		const openExplorerFolderInAgents = async (folderName: string) => {
			await ide.getByRole('tree').getByText(folderName, { exact: true }).click({ button: 'right' });
			await ide.getByRole('menuitem', { name: 'Open in Agents', exact: true }).click({ timeout: 10000 });
		};
		const openAgents = async () => {
			if (explorerFolderMenu) {
				stage = 'opening the second workspace root from the Explorer context menu';
				await openExplorerFolderInAgents('two');
				return;
			}
			const control = ide.getByRole('switch', { name: 'IDE mode', exact: true });
			if (!viewMenu) {
				if (focusedDiffSide) {
					await control.focus();
					await ide.locator('.part.editor .monaco-diff-editor .editor.original .monaco-editor.focused').waitFor({ state: 'hidden' });
					await control.press('Enter');
				} else {
					await control.click();
				}
				return;
			}
			stage = 'opening Agents from the compact View menu';
			await control.waitFor({ state: 'hidden' });
			const window = await launched.browserWindow(ide);
			const bounds = await window.evaluate(native => native.getBounds());
			try {
				const width = await window.evaluate((native, height) => {
					const [width] = native.getMinimumSize();
					native.setSize(width, height);
					return width;
				}, bounds.height);
				await ide.waitForFunction(expectedWidth => Math.abs(innerWidth - expectedWidth) <= 1, width / 1.2 ** 3, { timeout: 15000 });
				await ide.getByRole('menuitem', { name: 'Application Menu', exact: true }).click({ timeout: 10000 });
				const view = ide.getByRole('menuitem', { name: 'View', exact: true });
				await view.click();
				const action = ide.locator('.menubar-menu-items-holder').getByText('Open in Agents', { exact: true });
				await capture(ide, 'view-menu-open-in-agents.png');
				await action.waitFor({ timeout: 5000 });
				await action.click();
				return bounds.width;
			} finally {
				try {
					await window.evaluate((native, originalBounds) => native.setBounds(originalBounds), bounds);
				} finally {
					await window.dispose();
				}
			}
		};
		const [agents, originalIDEWidth] = await Promise.all([
			launched.waitForEvent('window', { timeout: 45000 }),
			openAgents(),
		]);
		await agents.locator('.agent-sessions-workbench').waitFor({ timeout: 60000 });
		if (selectsSecondRoot) {
			stage = 'selecting the second workspace root in Agents';
			await agents.locator('.new-chat-widget-container').getByText('two', { exact: true }).waitFor({ timeout: 15000 });
			report.recentWorkspaceRootVerified = !!recentWorkspaceRoot;
			await capture(agents, 'agents-selected-workspace-root.png');
		} else {
			await agents.getByText('one', { exact: true }).first().waitFor({ timeout: 45000 });
		}
		if (explorerFolderMenu) {
			stage = 'opening workspace roots from Explorer with Agents already open';
			const composer = agents.locator('.new-chat-widget-container');
			for (const rootName of ['one', 'two']) {
				await openExplorerFolderInAgents(rootName);
				await composer.getByText(rootName, { exact: true }).waitFor({ timeout: 15000 });
			}
			report.explorerFolderMenuVerified = true;
		}
		if (viewMenu) {
			const settingsFile = join(profile, 'User', 'settings.json');
			const settings = JSON.parse(await readFile(settingsFile, 'utf8'));
			delete settings['window.menuBarVisibility'];
			delete settings['window.zoomLevel'];
			delete settings['chat.titleBar.openInAgentsWindow.enabled'];
			await writeFile(settingsFile, JSON.stringify(settings));
			assert.ok(originalIDEWidth);
			await ide.waitForFunction(width => Math.abs(innerWidth - width) <= 1, originalIDEWidth, { timeout: 15000 });
			await ide.getByRole('switch', { name: 'IDE mode', exact: true }).waitFor({ timeout: 15000 });
			report.viewMenuHandoffVerified = true;
		}
		if (switchAccessibility) {
			stage = 'checking both mode switches in forced colors';
			await checkSwitchAccessibility(ide, 'IDE');
			await checkSwitchAccessibility(agents, 'Agents');
		}
		const alternateWorkspace = join(run, 'Alternate.code-workspace');
		const agentsWorkspace = join(profile, 'User', 'agent-sessions.code-workspace');
		let alternate: Page | undefined;
		if (originatingEditor) {
			stage = 'opening a second IDE with the same folders';
			await writeFile(alternateWorkspace, JSON.stringify({ folders: folders.map(path => ({ path })) }));
			[alternate] = await Promise.all([
				launched.waitForEvent('window', { timeout: 45000 }),
				forwardToRunningApp(['--new-window', alternateWorkspace]),
			]);
			await alternate.getByRole('tree').getByText('two', { exact: true }).waitFor();
			if (themes) {
				await alternate.locator('body').click({ position: { x: 600, y: 400 } });
				await alternate.keyboard.press('Control+n');
				await alternate.locator('.part.editor .monaco-editor-background').first().waitFor();
			}
		}
		let expectedIds = await checkWindows('Agents ready');
		assert.equal(expectedIds.length, alternate ? 3 : 2, 'Only the fixture windows should be open');
		const workspaceRoutes = async () => {
			const logs = join(profile, 'logs');
			const directories = (await readdir(logs, { withFileTypes: true })).filter(entry => entry.isDirectory()).map(entry => entry.name).sort();
			const mainLogs = await Promise.all(directories.map(directory => readFile(join(logs, directory, 'main.log'), 'utf8').catch(() => '')));
			return mainLogs.flatMap(log => [...log.matchAll(/windowsManager#open pathsToOpen (\[[^\r\n]*\])/g)]);
		};
		const expectRoutedWorkspace = async (expectedPath: string, switchToEditor: () => Promise<void>, key: 'configPath' | 'uri' = 'configPath') => {
			const count = (await workspaceRoutes()).length;
			await switchToEditor();
			const until = Date.now() + 10000;
			let route = (await workspaceRoutes())[count];
			while (!route && Date.now() < until) {
				await new Promise(resolve => setTimeout(resolve, 100));
				route = (await workspaceRoutes())[count];
			}
			assert.ok(route, 'The desktop main process must record the selected workspace');
			const paths = JSON.parse(route[1]) as {
				workspace?: Partial<Record<typeof key, { path: string }>>;
				fileUri?: { path: string };
				options?: {
					viewState?: {
						cursorState?: { selectionStart: { lineNumber: number; column: number }; position: { lineNumber: number; column: number } }[];
						viewState?: { firstPosition?: { lineNumber: number } };
					};
				};
			}[];
			assert.deepEqual(paths.filter(path => path.workspace).map(path => path.workspace?.[key]?.path.toLowerCase()), [decodeURIComponent(pathToFileURL(expectedPath).pathname).toLowerCase()]);
			return paths;
		};
		const input = agents.locator('.new-chat-widget-container .interactive-input-part .monaco-editor');
		const draft = 'Signal-C native draft preserved across mode switches';
		const inlineDraftNotice = agents.locator('.new-chat-widget-container .chat-input-notification-widget').filter({ hasText: 'Your draft stayed in this workspace.' });
		let otherIDE: Page | undefined;
		const useModeSwitch = (page: Page) => page.getByRole('switch', { name: 'IDE mode', exact: true }).click();
		const notificationsCenter = agents.locator('.notifications-center');
		const showNotifications = async () => {
			if (!await notificationsCenter.isVisible()) {
				await agents.keyboard.press('Control+Shift+P');
				const picker = agents.locator('.quick-input-widget');
				await picker.locator('input').first().fill('>Show Notifications');
				await picker.getByRole('option').filter({ hasText: 'Show Notifications' }).first().click();
				await notificationsCenter.waitFor({ timeout: 15000 });
			}
		};
		if (warmWorkspaceSwitch) {
			stage = 'opening a second project with no shared workspace roots';
			const otherFolder = join(run, 'other-project');
			await mkdir(otherFolder);
			await writeFile(join(otherFolder, 'other.txt'), 'Another workspace\n');
			[otherIDE] = await Promise.all([
				launched.waitForEvent('window', { timeout: 45000 }),
				forwardToRunningApp(['--new-window', otherFolder]),
			]);
			await otherIDE.getByRole('tree').getByText('other.txt', { exact: true }).waitFor({ timeout: 60000 });
			expectedIds = await checkWindows(stage);
			assert.equal(expectedIds.length, 3);
			const composer = agents.locator('.new-chat-widget-container');
			stage = 'switching an idle Agents composer to the other project';
			await useModeSwitch(otherIDE);
			await composer.getByText('other-project', { exact: true }).waitFor({ timeout: 15000 });
			stage = 'switching the idle composer back to the first project';
			await useModeSwitch(ide);
			await composer.getByText('one', { exact: true }).waitFor({ timeout: 15000 });
			report.warmWorkspaceSwitchVerified = true;
			stage = 'preserving the populated composer when switching from the other project';
			await input.click();
			await agents.keyboard.type(draft);
			await useModeSwitch(otherIDE);
			const noticeAction = saveDraftOnWorkspaceSwitch ? 'Save Draft and Open Workspace' : 'Open Draft Settings';
			await inlineDraftNotice.getByText('The switch requested other-project.', { exact: true }).waitFor({ timeout: 15000 });
			await inlineDraftNotice.getByRole('button', { name: `Your draft stayed in this workspace. ${noticeAction}`, exact: true }).waitFor({ timeout: 15000 });
			await capture(agents, 'agents-inline-workspace-handoff.png');
			await showNotifications();
			await notificationsCenter.getByRole('button', { name: noticeAction, exact: true }).first().waitFor({ timeout: 15000 });
			await agents.keyboard.press('Escape');
			await notificationsCenter.waitFor({ state: 'hidden' });
			report.inlineDraftSwitchNoticeVerified = true;
			if (!saveDraftOnWorkspaceSwitch) {
				await inlineDraftNotice.getByRole('button', { name: 'Dismiss notification', exact: true }).click();
				await inlineDraftNotice.waitFor({ state: 'hidden' });
				await showNotifications();
				await notificationsCenter.getByRole('button', { name: noticeAction, exact: true }).waitFor({ state: 'hidden' });
				await agents.keyboard.press('Escape');
				await notificationsCenter.waitFor({ state: 'hidden' });
				report.inlineDraftSwitchNoticeDismissed = true;
			}
			await composer.getByText('one', { exact: true }).waitFor({ timeout: 15000 });
			assert.ok((await input.innerText()).replaceAll('\u00a0', ' ').includes(draft));
			report.warmWorkspaceDraftRetained = true;
		} else {
			await input.click(); await agents.keyboard.type(draft);
		}
		await capture(agents, 'agents-draft.png');
		if (themes) {
			stage = 'live theme change in both modes';
			expectedTheme = themePalettes.light;
			const settingsFile = join(profile, 'User', 'settings.json');
			const currentSettings = JSON.parse(await readFile(settingsFile, 'utf8'));
			await writeFile(settingsFile, JSON.stringify({ ...currentSettings, 'workbench.colorTheme': expectedTheme.name, 'workbench.colorCustomizations': expectedTheme.colors }));
			assert.deepEqual(await checkWindows(stage), expectedIds);
		}
		if (keyboard) {
			stage = 'changing the title bar switch setting';
			const settingsFile = join(profile, 'User', 'settings.json');
			const currentSettings = JSON.parse(await readFile(settingsFile, 'utf8'));
			const windows = [ide, agents];
			const rendererStarts = () => Promise.all(windows.map(page => page.evaluate(() => performance.timeOrigin)));
			const starts = await rendererStarts();
			const setSwitchEnabled = (enabled: boolean) => writeFile(settingsFile, JSON.stringify({ ...currentSettings, 'workbench.titleBar.workspaceModeSwitch': enabled }));
			await setSwitchEnabled(false);
			await Promise.all([
				ide.locator('.part.titlebar .open-in-agents-titlebar-widget').waitFor({ timeout: 15000 }),
				agents.locator('.part.titlebar .open-in-vscode-titlebar-widget').waitFor({ timeout: 15000 }),
			]);
			for (const page of windows) {
				assert.equal(await page.getByRole('switch', { name: 'IDE mode', exact: true }).count(), 0);
			}
			await setSwitchEnabled(true);
			await Promise.all(windows.map(page => page.getByRole('switch', { name: 'IDE mode', exact: true }).waitFor({ timeout: 15000 })));
			assert.deepEqual(await rendererStarts(), starts);
			assert.deepEqual(await checkWindows(stage), expectedIds);
			report.switchSettingChangedWithoutReload = true;
		}
		for (let roundTrip = 1; roundTrip <= 2; roundTrip++) {
			stage = `return to IDE ${roundTrip}`;
			if (alternate) {
				await simulateNativeFocus(alternate);
				if (focusAgentsFromOtherEditor && roundTrip === 1) {
					stage = 'focusing Agents from the alternate IDE using Open Agents Window';
					await expectRoutedWorkspace(agentsWorkspace, async () => {
						await alternate.keyboard.press('Control+Shift+P');
						const picker = alternate.locator('.quick-input-widget');
						await picker.locator('input').first().fill('>Open Agents Window');
						await picker.getByRole('option').filter({ hasText: 'Open Agents Window' }).first().click();
					});
					stage = 'returning to the original IDE after the focus-only command';
				}
				await expectRoutedWorkspace(workspace, () => activateSwitch(agents, roundTrip, input));
				report.originatingEditorVerified = true;
				report.focusOnlyPreservedOrigin ||= !!focusAgentsFromOtherEditor && roundTrip === 1;
			} else {
				await activateSwitch(agents, roundTrip, input);
			}
			await agents.locator('[role="switch"][aria-label="IDE mode"][aria-busy="false"]').waitFor({ timeout: 30000 });
			assert.deepEqual(await checkWindows(stage), expectedIds, 'The handoff must not open another editor window');
			if (diffWorkspaceRoot && roundTrip === 1) {
				await ide.locator('.tabs-container .tab.active').getByText(/(comparison|outside)\.txt/).waitFor();
				await ide.locator('.tabs-container .tab.dirty').press('Enter');
				report.diffWorkspaceRootVerified = true;
				report.focusedDiffSideVerified = !!focusedDiffSide;
			}
			if (explorerWorkspaceRoot && roundTrip === 1) {
				assert.equal(await ide.locator('.tabs-container .tab.active').count(), 0);
				await selectedExplorerRoot.waitFor();
				report.explorerWorkspaceRootVerified = true;
				await createDirtyEditor();
			}
			await ide.getByRole('code').getByText(marker, { exact: false }).waitFor();
			await ide.locator('.tabs-container .tab.active.dirty').waitFor();
			if (untitledPath) {
				await ide.locator('.tabs-container .tab.active.dirty').getByText(untitledName, { exact: true }).waitFor();
				await assert.rejects(readFile(untitledPath), { code: 'ENOENT' });
				report.untitledWorkspaceRootVerified = !!untitledWorkspaceRoot;
			}
			await ide.getByRole('tree').getByText('two', { exact: true }).waitFor();
			report.windowReuseVerified = true; report.dirtyBufferRetained = true;
			await capture(ide, `ide-return-${roundTrip}.png`);
			stage = `return to Agents ${roundTrip}`;
			const switchToAgents = () => activateSwitch(ide, roundTrip, ide.getByRole('code').getByText(marker, { exact: false }));
			if (explorerFolderMenu && roundTrip === 1) {
				stage = 'opening another Explorer root with an unsent Agents draft';
				await openExplorerFolderInAgents('one');
				await agents.waitForFunction(() => ['Agents kept the current session.', 'Your draft stayed in this workspace.'].some(message => document.body.innerText.includes(message)), undefined, { timeout: 15000 });
				await agents.locator('.new-chat-widget-container').getByText('two', { exact: true }).waitFor();
				report.explorerFolderMenuPreservedDraft = true;
			} else if (switchLayout && roundTrip === 2) {
				await exerciseNarrowSwitch(ide, switchToAgents);
			} else {
				await switchToAgents();
			}
			await ide.locator('[role="switch"][aria-label="IDE mode"][aria-busy="false"]').waitFor({ timeout: 30000 });
			assert.deepEqual(await checkWindows(stage), expectedIds);
			assert.ok((await input.innerText()).replaceAll('\u00a0', ' ').includes(draft), 'The original Agents draft must remain');
			report.agentsDraftRetained = true;
			console.log(`Native round trip ${roundTrip} retained the window IDs, the dirty buffer, and the Agents draft`);
		}
		await capture(agents, 'agents-returned.png');
		if (reopenSavedDiff) {
			const selectedRoot = selectsSecondRoot ? 1 : 0;
			const original = join(folders[selectedRoot], 'compare-original.txt');
			const modified = join(folders[1 - selectedRoot], 'compare-modified.txt');
			await writeFile(original, 'The original comparison file.\n');
			await writeFile(modified, 'The modified comparison file.\n');
			stage = 'capturing both sides of a saved workspace comparison';
			await simulateNativeFocus(ide);
			await forwardToRunningApp(['--reuse-window', '--diff', original, modified]);
			const originalPane = (page: Page) => page.locator('.part.editor .monaco-diff-editor .editor.original');
			await originalPane(ide).locator('.view-line').first().click({ position: { x: 8, y: 8 } });
			await ide.keyboard.press('Control+k');
			await ide.keyboard.press('Shift+Enter');
			await ide.locator('.tabs-container .tab.active.sticky').getByText(/compare-modified\.txt/).waitFor();
			await ide.keyboard.press('Control+Home');
			await ide.keyboard.press('ArrowRight');
			await ide.keyboard.press('ArrowRight');
			for (let i = 0; i < 3; i++) { await ide.keyboard.press('Shift+ArrowRight'); }
			await originalPane(ide).locator('.monaco-editor.focused .selected-text').first().waitFor();
			await expectRoutedWorkspace(agentsWorkspace, () => useModeSwitch(ide));
			await destroyFixtureWindow(ide);

			stage = 'reopening both sides of the saved workspace comparison';
			const [reopened, paths] = await Promise.all([
				launched.waitForEvent('window', { timeout: 45000 }),
				expectRoutedWorkspace(workspace, () => useModeSwitch(agents)),
			]);
			const comparePath = (path: string) => decodeURIComponent(pathToFileURL(path).pathname).toLowerCase();
			assert.deepEqual(paths.flatMap(path => path.fileUri ? [path.fileUri.path.toLowerCase()] : []), [comparePath(original), comparePath(modified)]);
			const diff = reopened.locator('.part.editor .monaco-diff-editor');
			await diff.locator('.editor.original .view-line').getByText('The original comparison file.', { exact: true }).waitFor();
			await diff.locator('.editor.modified .view-line').getByText('The modified comparison file.', { exact: true }).waitFor();
			await originalPane(reopened).locator('.monaco-editor.focused .selected-text').first().waitFor();
			assert.ok(await reopened.locator('.tabs-container .tab.active.sticky').getByText(/compare-modified\.txt/).isVisible(), 'The comparison must reopen as a pinned tab');
			assert.ok((await input.innerText()).replaceAll('\u00a0', ' ').includes(draft));
			report.reopenedSavedDiffVerified = true;
			report.reopenedStickyDiffVerified = true;
			await capture(reopened, 'saved-workspace-restored-diff.png');

			stage = 'restoring the active file when the other side of the comparison is gone';
			await expectRoutedWorkspace(agentsWorkspace, () => useModeSwitch(reopened));
			await destroyFixtureWindow(reopened);
			await rename(modified, join(folders[1 - selectedRoot], 'moved-comparison-file.txt'));
			const [fallback, fallbackPaths] = await Promise.all([
				launched.waitForEvent('window', { timeout: 45000 }),
				expectRoutedWorkspace(workspace, () => useModeSwitch(agents)),
			]);
			assert.deepEqual(fallbackPaths.flatMap(path => path.fileUri ? [path.fileUri.path.toLowerCase()] : []), [comparePath(original)]);
			await fallback.locator('.tabs-container .tab.active').getByText('compare-original.txt', { exact: true }).waitFor();
			await fallback.locator('.part.statusbar').getByText('Ln 1, Col 6 (3 selected)', { exact: true }).waitFor();
			assert.equal(await fallback.locator('.part.editor .monaco-diff-editor').count(), 0);
			report.reopenedDiffWithMissingFileFallbackVerified = true;
		}
		if (reopenSavedEditorFile) {
			const root = folders[selectsSecondRoot ? 1 : 0];
			const file = join(root, 'saved-workspace-selection.txt');
			const firstLine = 'Keep this line and selection in the saved workspace.';
			const content = [firstLine, 'Keep a second cursor here, too.', ...Array.from({ length: 120 }, (_, index) => 'Viewport test line ' + (index + 1)), ''].join('\n');
			await writeFile(file, content);
			const differentPreviewName = 'different-preview.txt';
			if (previewConflict) {
				await writeFile(join(root, differentPreviewName), 'Keep this newer preview in its own tab.\n');
			}
			const selectionLabel = '2 selections (6 characters selected)';
			const fileName = 'saved-workspace-selection.txt';
			const activeTab = (page: Page) => page.locator('.part.editor .editor-group-container.active .tabs-container .tab.active');
			const previewLabel = (tab: Locator) => tab.locator('.monaco-icon-label.italic');
			const openExplorerPreview = async (page: Page, name: string) => {
				const explorer = page.locator('.part.sidebar').getByRole('tree').getByText(name, { exact: true });
				await explorer.click();
				await previewLabel(activeTab(page)).getByText(name, { exact: true }).waitFor({ timeout: 15000 });
			};
			const activeEditor = (page: Page) => page.locator('.part.editor .editor-group-container.active .monaco-editor').first();
			const leadingViewportLine = (page: Page) => activeEditor(page).evaluate(editor => {
				const scrollable = editor.querySelector('.editor-scrollable');
				if (!scrollable) { return ''; }
				const top = scrollable.getBoundingClientRect().top;
				const line = [...editor.querySelectorAll('.view-lines .view-line')].find(line => line.getBoundingClientRect().bottom > top + 1);
				return line?.textContent?.replaceAll('\u00a0', ' ') ?? '';
			});
			const waitForViewportLine = async (page: Page, matches: (text: string) => boolean) => {
				const until = Date.now() + 15000;
				let text = await leadingViewportLine(page);
				while (!matches(text) && Date.now() < until) {
					await new Promise(resolve => setTimeout(resolve, 100));
					text = await leadingViewportLine(page);
				}
				assert.ok(matches(text), 'Unexpected leading editor line: ' + JSON.stringify(text));
				return text;
			};
			let topLine: string;
			const captureSelection = async (page: Page) => {
				await simulateNativeFocus(page);
				if (reopenSavedEditorPreview) {
					await openExplorerPreview(page, fileName);
				} else {
					await forwardToRunningApp(['--reuse-window', file]);
					await activeTab(page).getByText(fileName, { exact: true }).waitFor();
				}
				await activeEditor(page).click({ position: { x: 120, y: 40 } });
				await page.keyboard.press('Control+Home');
				for (let i = 0; i < 5; i++) { await page.keyboard.press('ArrowRight'); }
				await page.keyboard.press('Control+Alt+ArrowDown');
				for (let i = 0; i < 3; i++) { await page.keyboard.press('Shift+ArrowLeft'); }
				await page.locator('.part.statusbar').getByText(selectionLabel, { exact: true }).waitFor({ timeout: 15000 });
				await activeEditor(page).hover();
				await page.mouse.wheel(0, 700);
				topLine = await waitForViewportLine(page, text => text.startsWith('Viewport test line '));
				await expectRoutedWorkspace(agentsWorkspace, () => useModeSwitch(page));
			};
			const reopenSavedWorkspace = async (hasDifferentPreview = false, workspacePath = workspace) => {
				const [reopened, paths] = await Promise.all([
					launched.waitForEvent('window', { timeout: 45000 }),
					expectRoutedWorkspace(workspacePath, () => useModeSwitch(agents)),
				]);
				const recovered = paths.find(path => path.fileUri?.path.toLowerCase() === decodeURIComponent(pathToFileURL(file).pathname).toLowerCase());
				assert.ok(recovered, 'The saved workspace must reopen the captured file');
				await activeTab(reopened).getByText(fileName, { exact: true }).waitFor({ timeout: 15000 });
				await reopened.locator('.part.statusbar').getByText(selectionLabel, { exact: true }).waitFor({ timeout: 15000 });
				assert.deepEqual(recovered.options?.viewState?.cursorState?.map(({ selectionStart, position }) => [selectionStart, position]), [
					[{ lineNumber: 1, column: 6 }, { lineNumber: 1, column: 3 }],
					[{ lineNumber: 2, column: 6 }, { lineNumber: 2, column: 3 }],
				]);
				assert.ok((recovered.options?.viewState?.viewState?.firstPosition?.lineNumber ?? 0) > 2);
				await waitForViewportLine(reopened, text => text === topLine);
				report.reopenedSavedEditorViewportVerified = true;
				if (reopenSavedEditorPreview) {
					assert.equal(await previewLabel(activeTab(reopened)).count(), hasDifferentPreview ? 0 : 1, 'The recovered tab should remain a preview unless that would replace a different preview');
					if (hasDifferentPreview) {
						await previewLabel(reopened.locator('.part.editor .editor-group-container.active .tabs-container .tab')).getByText(differentPreviewName, { exact: true }).waitFor();
						report.restoredDifferentPreviewVerified = true;
					} else {
						report.reopenedPreviewTabVerified = true;
					}
				}
				assert.equal(await readFile(file, 'utf8'), content);
				assert.ok((await input.innerText()).replaceAll('\u00a0', ' ').includes(draft));
				return reopened;
			};
			stage = 'capturing an active file and selection from the saved workspace';
			const editorGroups = (page: Page) => page.locator('.part.editor .editor-group-container');
			if (reopenSavedEditorGroup) {
				// Dispose of the synthetic scratch buffer so the fixture can close the IDE normally.
				await runWorkbenchCommand(ide, 'Revert and Close Editor');
				await ide.locator('.tabs-container .tab.dirty').waitFor({ state: 'hidden' });
				const leftFile = join(root, 'left-group.txt');
				await writeFile(leftFile, 'Keep this file open in the first editor group.\n');
				await simulateNativeFocus(ide);
				await forwardToRunningApp(['--reuse-window', leftFile]);
				await ide.locator('.tabs-container .tab.active').getByText('left-group.txt', { exact: true }).waitFor();
				await runWorkbenchCommand(ide, 'Split Editor Right');
				await editorGroups(ide).nth(1).waitFor();
			}
			await captureSelection(ide);
			if (reopenSavedEditorGroup) {
				await editorGroups(ide).nth(1).locator('.tabs-container .tab.active').getByText('saved-workspace-selection.txt', { exact: true }).waitFor();
				if (previewConflict) {
					await openExplorerPreview(ide, differentPreviewName);
				}
				const closed = ide.waitForEvent('close', { timeout: 30000 });
				const window = await launched.browserWindow(ide);
				try { await window.evaluate(native => native.close()); } finally { await window.dispose(); }
				await closed;
			} else {
				await destroyFixtureWindow(ide);
			}

			stage = 'reopening the saved workspace with its active file and selection';
			const reopened = await reopenSavedWorkspace(previewConflict);
			await reopened.getByRole('tree').getByText('two', { exact: true }).waitFor();
			if (disableEditorRestoration) {
				assert.equal(await editorGroups(reopened).count(), 1, 'Turning off editor restoration should still reopen the file in a single group');
				report.editorRestorationPreferenceVerified = true;
			} else if (reopenSavedEditorGroup) {
				assert.equal(await editorGroups(reopened).count(), 2, 'The ordinary IDE close should preserve both editor groups');
				await editorGroups(reopened).nth(0).locator('.tabs-container .tab.active').getByText('left-group.txt', { exact: true }).waitFor();
				const right = editorGroups(reopened).nth(1);
				await right.locator('.tabs-container .tab.active').getByText('saved-workspace-selection.txt', { exact: true }).waitFor();
				assert.ok(await right.evaluate(group => group.classList.contains('active')), 'The restored file must be active in the original group');
				report.reopenedSavedEditorGroupVerified = true;
			}
			report.reopenedSavedEditorFileVerified = true;
			await capture(reopened, 'saved-workspace-restored-file.png');

			stage = 'preserving the active tab when the saved workspace is already open';
			await expectRoutedWorkspace(agentsWorkspace, () => useModeSwitch(reopened));
			await destroyFixtureWindow(reopened);
			const [replacement] = await Promise.all([
				launched.waitForEvent('window', { timeout: 45000 }),
				forwardToRunningApp(['--new-window', workspace]),
			]);
			await replacement.getByRole('tree').getByText('two', { exact: true }).waitFor();
			await simulateNativeFocus(replacement);
			await forwardToRunningApp(['--reuse-window', join(root, 'README.md')]);
			const replacementTab = replacement.locator('.tabs-container .tab.active').getByText('README.md', { exact: true });
			await replacementTab.waitFor();
			const liveIds = await checkWindows(stage);
			const livePaths = await expectRoutedWorkspace(workspace, () => useModeSwitch(agents));
			assert.equal(livePaths.length, 1, 'Returning to an already open workspace must not request a stale file');
			await agents.locator('[role="switch"][aria-label="IDE mode"][aria-busy="false"]').waitFor();
			assert.deepEqual(await checkWindows(stage), liveIds);
			assert.equal(await replacementTab.isVisible(), true);
			report.reopenedSavedEditorReplacementPreserved = true;

			stage = 'reopening a saved workspace configured with a single root';
			await destroyFixtureWindow(replacement);
			const singleRootWorkspace = join(run, 'continuity-single-root.code-workspace');
			await writeFile(singleRootWorkspace, JSON.stringify({ folders: [{ path: root }] }));
			const [singleRoot] = await Promise.all([
				launched.waitForEvent('window', { timeout: 45000 }),
				forwardToRunningApp(['--new-window', singleRootWorkspace]),
			]);
			await captureSelection(singleRoot);
			await destroyFixtureWindow(singleRoot);
			await reopenSavedWorkspace(false, singleRootWorkspace);
			report.reopenedSingleRootWorkspaceFileVerified = true;
		}
		if (runningConversation) {
			stage = 'starting a workspace conversation in Agents';
			await selectMockModel(agents);
			await input.click();
			await agents.keyboard.press('Control+a');
			await agents.keyboard.type('Start a project conversation for the mode switch');
			await agents.keyboard.press('Enter');
			const agentsChat = agents.locator('.interactive-session').filter({ hasText: responseText }).last();
			await agentsChat.getByText(responseText, { exact: true }).first().waitFor({ timeout: 45000 });
			const liveInput = agentsChat.locator('.interactive-input-part .monaco-editor');
			await liveInput.click();
			await agents.keyboard.type(runningPrompt);
			await agents.keyboard.press('Enter');
			const runningRequests = () => mock.getRequests().filter(request => JSON.stringify(request).includes(runningPrompt));
			const requestDeadline = Date.now() + 15000;
			while (!runningRequests().length && Date.now() < requestDeadline) {
				await new Promise(resolve => setTimeout(resolve, 100));
			}
			assert.ok(runningRequests().length, 'The fixture model must receive the request before switching modes');
			stage = 'opening an in-progress conversation in the project IDE';
			await agents.getByRole('switch', { name: 'IDE mode', exact: true }).press('Space');
			const ideChat = ide.locator('.interactive-session').filter({ hasText: responseText }).last();
			await ideChat.getByText(runningPrompt, { exact: true }).first().waitFor({ timeout: 25000 });
			assert.equal(await ideChat.getByText(runningReply, { exact: true }).count(), 0);
			assert.deepEqual(await checkWindows(stage), expectedIds);
			report.runningConversationOpenedInIDE = true;
			runningResponse.resolve();
			await ideChat.getByText(runningReply, { exact: true }).first().waitFor({ timeout: 25000 });
			await ide.locator('.tabs-container .tab.active.dirty').waitFor();
			await ide.getByRole('code').getByText(marker, { exact: false }).waitFor();
			stage = 'returning to the same conversation in Agents';
			const requestsAfterReply = runningRequests().length;
			await ide.getByRole('switch', { name: 'IDE mode', exact: true }).press('Enter');
			await ide.locator('[role="switch"][aria-label="IDE mode"][aria-busy="false"]').waitFor();
			await agentsChat.getByText(runningReply, { exact: true }).first().waitFor();
			assert.deepEqual(await checkWindows(stage), expectedIds);
			assert.equal(runningRequests().length, requestsAfterReply, 'A mode switch must not resend the prompt');
			report.runningConversationRoundTripVerified = true;
			console.log('The project IDE received the in-progress chat and returned to Agents with its reply');
		}
		if (otherIDE) {
			stage = 'returning to the draft workspace after the first IDE closed';
			await expectRoutedWorkspace(agentsWorkspace, () => useModeSwitch(otherIDE));
			await destroyFixtureWindow(ide);
			const [reopened] = await Promise.all([
				launched.waitForEvent('window', { timeout: 45000 }),
				expectRoutedWorkspace(workspace, () => useModeSwitch(agents)),
			]);
			await reopened.getByRole('tree').getByText('one', { exact: true }).waitFor({ timeout: 60000 });
			await reopened.getByRole('tree').getByText('two', { exact: true }).waitFor();
			assert.ok((await input.innerText()).replaceAll('\u00a0', ' ').includes(draft));
			report.warmWorkspaceReturnRetained = true;
			if (saveDraftOnWorkspaceSwitch) {
				stage = 'saving the original draft and opening the requested project from the notice';
				const composer = agents.locator('.new-chat-widget-container');
				const action = inlineDraftNotice.getByRole('button', { name: 'Your draft stayed in this workspace. Save Draft and Open Workspace', exact: true });
				await action.focus({ timeout: 15000 });
				await action.press('Enter');
				await composer.getByText('other-project', { exact: true }).waitFor({ timeout: 15000 });
				await inlineDraftNotice.waitFor({ state: 'hidden' });
				await agents.waitForFunction(expected => {
					const editor = document.querySelector<HTMLElement>('.new-chat-widget-container .interactive-input-part .monaco-editor');
					return !!editor?.getBoundingClientRect().height && !editor.innerText.replaceAll('\u00a0', ' ').includes(expected);
				}, draft, { timeout: 15000 });
				stage = 'recovering the saved draft in its original project';
				await showNotifications();
				await notificationsCenter.getByRole('link', { name: 'Open Draft', exact: true }).first().click({ timeout: 15000 });
				await composer.getByText('one', { exact: true }).waitFor({ timeout: 15000 });
				await agents.waitForFunction(expected => document.querySelector<HTMLElement>('.new-chat-widget-container .interactive-input-part .monaco-editor')?.innerText.replaceAll('\u00a0', ' ').includes(expected), draft, { timeout: 15000 });
				report.savedDraftWorkspaceSwitchVerified = true;
			}
		}
		if (alternate) {
			const switchToEditor = () => agents.getByRole('switch', { name: 'IDE mode', exact: true }).click();
			stage = 'returning to the single-folder editor after switching from a saved workspace';
			const activeFolder = selectsSecondRoot ? folders[1] : folders[0];
			const [singleFolderEditor] = await Promise.all([
				launched.waitForEvent('window', { timeout: 45000 }),
				forwardToRunningApp(['--new-window', activeFolder]),
			]);
			await singleFolderEditor.getByRole('tree').getByText('README.md', { exact: true }).waitFor();
			await expectRoutedWorkspace(agentsWorkspace, () => useModeSwitch(singleFolderEditor));
			await expectRoutedWorkspace(activeFolder, switchToEditor, 'uri');
			await agents.locator('[role="switch"][aria-label="IDE mode"][aria-busy="false"]').waitFor({ timeout: 30000 });
			assert.equal((await checkWindows(stage)).length, expectedIds.length + 1);
			assert.ok((await input.innerText()).replaceAll('\u00a0', ' ').includes(draft), 'Switching to the single-folder editor must keep the Agents draft');
			stage = 'returning to the saved workspace after switching from the single-folder editor';
			await expectRoutedWorkspace(agentsWorkspace, () => useModeSwitch(ide));
			await expectRoutedWorkspace(workspace, switchToEditor);
			await agents.locator('[role="switch"][aria-label="IDE mode"][aria-busy="false"]').waitFor({ timeout: 30000 });
			assert.equal((await checkWindows(stage)).length, expectedIds.length + 1);
			report.singleFolderOriginVerified = true;

			stage = 'reopening the closed single-folder editor with its selected file';
			await singleFolderEditor.getByRole('tree').getByText('README.md', { exact: true }).dblclick();
			await singleFolderEditor.locator('.tabs-container .tab.active').getByText('README.md', { exact: true }).waitFor();
			await singleFolderEditor.locator('.part.editor .monaco-editor .view-lines').first().click();
			await singleFolderEditor.keyboard.press('Control+Home');
			await singleFolderEditor.keyboard.press('ArrowRight');
			await singleFolderEditor.keyboard.press('ArrowRight');
			for (let i = 0; i < 3; i++) {
				await singleFolderEditor.keyboard.press('Shift+ArrowRight');
			}
			const selectionLabel = 'Ln 1, Col 6 (3 selected)';
			await singleFolderEditor.locator('.part.statusbar').getByText(selectionLabel, { exact: true }).waitFor();
			await expectRoutedWorkspace(agentsWorkspace, () => useModeSwitch(singleFolderEditor));
			await destroyFixtureWindow(singleFolderEditor);
			const [reopenedFolderEditor] = await Promise.all([
				launched.waitForEvent('window', { timeout: 45000 }),
				expectRoutedWorkspace(activeFolder, switchToEditor, 'uri'),
			]);
			await reopenedFolderEditor.locator('.tabs-container .tab.active').getByText('README.md', { exact: true }).waitFor();
			await reopenedFolderEditor.locator('.part.statusbar').getByText(selectionLabel, { exact: true }).waitFor();
			assert.equal((await checkWindows(stage)).length, expectedIds.length + 1);
			assert.ok((await input.innerText()).replaceAll('\u00a0', ' ').includes(draft), 'Reopening the single-folder editor must keep the Agents draft');
			report.closedSingleFolderOriginVerified = true;

			stage = 'returning to a reopened folder without changing its active tab';
			await expectRoutedWorkspace(agentsWorkspace, () => useModeSwitch(reopenedFolderEditor));
			await destroyFixtureWindow(reopenedFolderEditor);
			const replacementFile = 'replacement.txt';
			await writeFile(join(activeFolder, replacementFile), 'This is the active file in the replacement editor.\n');
			const [replacementFolderEditor] = await Promise.all([
				launched.waitForEvent('window', { timeout: 45000 }),
				forwardToRunningApp(['--new-window', activeFolder]),
			]);
			await replacementFolderEditor.getByRole('tree').getByText(replacementFile, { exact: true }).dblclick();
			const replacementTab = replacementFolderEditor.locator('.tabs-container .tab.active').getByText(replacementFile, { exact: true });
			await replacementTab.waitFor();
			const paths = await expectRoutedWorkspace(activeFolder, switchToEditor, 'uri');
			assert.equal(paths.length, 1, 'Returning to an already open folder must not request a stale file');
			await agents.locator('[role="switch"][aria-label="IDE mode"][aria-busy="false"]').waitFor();
			assert.equal(await replacementTab.isVisible(), true, 'Returning to an already open folder must keep its active tab');
			assert.equal((await checkWindows(stage)).length, expectedIds.length + 1);
			report.replacementFolderEditorPreserved = true;

			stage = 'returning after the original editor switched to a different folder';
			const anotherFolder = join(run, 'other-project');
			const anotherFile = 'other.txt';
			const anotherFilePath = join(anotherFolder, anotherFile);
			const anotherFileContents = 'Keep this project open in the original window.\n';
			await mkdir(anotherFolder);
			await writeFile(anotherFilePath, anotherFileContents);
			execFileSync('git', ['init', '--quiet'], { cwd: anotherFolder, windowsHide: true, stdio: 'pipe' });
			await expectRoutedWorkspace(agentsWorkspace, () => useModeSwitch(replacementFolderEditor));
			await simulateNativeFocus(replacementFolderEditor);
			await expectRoutedWorkspace(anotherFolder, () => forwardToRunningApp(['--reuse-window', anotherFolder]), 'uri');
			await replacementFolderEditor.getByRole('tree').getByText(anotherFile, { exact: true }).dblclick();
			const anotherTab = replacementFolderEditor.locator('.tabs-container .tab.active').getByText(anotherFile, { exact: true });
			await anotherTab.waitFor();
			const anotherEditor = replacementFolderEditor.locator('.part.editor .monaco-editor .view-lines').first();
			const unsavedEdit = 'Unsaved edit in the other project.';
			await anotherEditor.click();
			await replacementFolderEditor.keyboard.press('Control+Home');
			await replacementFolderEditor.keyboard.type(unsavedEdit);
			const [restoredFolderEditor] = await Promise.all([
				launched.waitForEvent('window', { timeout: 45000 }),
				expectRoutedWorkspace(activeFolder, switchToEditor, 'uri'),
			]);
			await restoredFolderEditor.locator('.tabs-container .tab.active').getByText(replacementFile, { exact: true }).waitFor();
			assert.equal(await anotherTab.isVisible(), true, 'Returning to the old project must leave the other project open');
			assert.ok((await anotherEditor.innerText()).replaceAll('\u00a0', ' ').includes(unsavedEdit), 'Returning to the old project must preserve the other editor’s unsaved changes');
			assert.equal(await readFile(anotherFilePath, 'utf8'), anotherFileContents);
			assert.equal((await checkWindows(stage)).length, expectedIds.length + 2);
			assert.ok((await input.innerText()).replaceAll('\u00a0', ' ').includes(draft), 'Returning to the old project must keep the Agents draft');
			report.repurposedFolderEditorPreserved = true;
			await destroyFixtureWindow(restoredFolderEditor);

			stage = 'keeping the first project file when the same window switches from a second project';
			await expectRoutedWorkspace(agentsWorkspace, () => useModeSwitch(replacementFolderEditor));
			await inlineDraftNotice.getByText('The switch requested other-project.', { exact: true }).waitFor({ timeout: 15000 });
			assert.ok((await input.innerText()).replaceAll('\u00a0', ' ').includes(draft), 'The first project must keep its draft when the second project opens Agents');
			const [draftFolderEditor] = await Promise.all([
				launched.waitForEvent('window', { timeout: 45000 }),
				expectRoutedWorkspace(activeFolder, switchToEditor, 'uri'),
			]);
			await draftFolderEditor.locator('.tabs-container .tab.active').getByText(replacementFile, { exact: true }).waitFor();
			assert.equal(await anotherTab.isVisible(), true, 'Returning to the draft’s editor must keep the second project open');
			report.reusedWindowDraftFilePreserved = true;
			await destroyFixtureWindow(draftFolderEditor);
			await destroyFixtureWindow(replacementFolderEditor);
			await expectRoutedWorkspace(agentsWorkspace, () => useModeSwitch(ide));

			if (focusAgentsFromOtherEditor) {
				stage = 'switching explicitly from the alternate IDE';
				await expectRoutedWorkspace(agentsWorkspace, () => alternate.getByRole('switch', { name: 'IDE mode', exact: true }).click());
				await expectRoutedWorkspace(alternateWorkspace, switchToEditor);
				await agents.locator('[role="switch"][aria-label="IDE mode"][aria-busy="false"]').waitFor({ timeout: 30000 });
				assert.ok((await input.innerText()).replaceAll('\u00a0', ' ').includes(draft), 'Explicitly switching from another IDE must keep the Agents draft');
				assert.deepEqual(await checkWindows(stage), expectedIds);
				await expectRoutedWorkspace(agentsWorkspace, () => ide.getByRole('switch', { name: 'IDE mode', exact: true }).click());
				report.explicitSwitchUpdatedOrigin = true;
			}
			stage = 'return after the originating editor closed';
			await destroyFixtureWindow(ide);
			await expectRoutedWorkspace(alternateWorkspace, switchToEditor);
			await agents.locator('[role="switch"][aria-label="IDE mode"][aria-busy="false"]').waitFor({ timeout: 30000 });
			assert.equal((await checkWindows(stage)).length, 2);
			report.closedOriginFallbackVerified = true;

			stage = 'reopening the originating workspace with its active file while an alternate is open';
			const alternateBuffer = 'Unsaved content in the alternate workspace';
			await alternate.keyboard.press('Control+n');
			await alternate.locator('.part.editor .monaco-editor .view-lines').first().click();
			await alternate.keyboard.type(alternateBuffer);
			const alternateText = alternate.getByRole('code').getByText(alternateBuffer, { exact: true });
			await alternateText.waitFor();
			const contextFile = join(activeFolder, 'originating-editor-context.txt');
			await writeFile(contextFile, 'Restore this file and selection to its original workspace.\n');
			const selectContextText = async (page: Page) => {
				await page.locator('.part.editor .monaco-editor .view-lines').first().click();
				await page.keyboard.press('Control+Home');
				await page.keyboard.press('ArrowRight');
				await page.keyboard.press('ArrowRight');
				for (let i = 0; i < 3; i++) {
					await page.keyboard.press('Shift+ArrowRight');
				}
				await page.locator('.part.statusbar').getByText(selectionLabel, { exact: true }).waitFor();
			};
			const [source] = await Promise.all([
				launched.waitForEvent('window', { timeout: 45000 }),
				forwardToRunningApp(['--new-window', workspace]),
			]);
			await source.getByRole('tree').getByText('originating-editor-context.txt', { exact: true }).dblclick();
			await selectContextText(source);
			await expectRoutedWorkspace(agentsWorkspace, () => useModeSwitch(source));
			await destroyFixtureWindow(source);
			const [recoveredSource] = await Promise.all([
				launched.waitForEvent('window', { timeout: 45000 }),
				expectRoutedWorkspace(workspace, switchToEditor),
			]);
			await recoveredSource.locator('.tabs-container .tab.active').getByText('originating-editor-context.txt', { exact: true }).waitFor();
			await recoveredSource.locator('.part.statusbar').getByText(selectionLabel, { exact: true }).waitFor();
			assert.ok(await alternateText.isVisible(), 'Reopening the originating workspace must keep the alternate editor’s unsaved buffer');
			assert.equal((await checkWindows(stage)).length, 3);
			report.closedOriginFilePreferredOverAlternate = true;

			stage = 'returning to the exact reopened workspace without changing its active tab';
			await recoveredSource.getByRole('tree').getByText('README.md', { exact: true }).first().dblclick();
			const changedTab = recoveredSource.locator('.tabs-container .tab.active').getByText('README.md', { exact: true });
			await changedTab.waitFor();
			const movedContextFile = join(activeFolder, 'moved-originating-editor-context.txt');
			await rename(contextFile, movedContextFile);
			await simulateNativeFocus(alternate);
			const returnPaths = await expectRoutedWorkspace(workspace, switchToEditor);
			assert.equal(returnPaths.length, 1, 'Returning to a live workspace must not request its stale file');
			assert.ok(await changedTab.isVisible(), 'Returning to the reopened editor must preserve its active tab');
			assert.equal((await checkWindows(stage)).length, 3);
			assert.ok((await input.innerText()).replaceAll('\u00a0', ' ').includes(draft));
			report.reopenedOriginWindowPreferredOverAlternate = true;
			await destroyFixtureWindow(recoveredSource);

			stage = 'using the alternate when the originating file is gone';
			await expectRoutedWorkspace(alternateWorkspace, switchToEditor);
			await agents.locator('[role="switch"][aria-label="IDE mode"][aria-busy="false"]').waitFor({ timeout: 30000 });
			assert.equal((await checkWindows(stage)).length, 2);
			assert.ok(await alternateText.isVisible(), 'The fallback must preserve the alternate editor’s unsaved buffer');
			report.missingOriginFileFallbackVerified = true;
			await rename(movedContextFile, contextFile);

			stage = 'reopening the originating workspace after all editors closed';
			await destroyFixtureWindow(alternate);
			const [reopened] = await Promise.all([
				launched.waitForEvent('window', { timeout: 45000 }),
				expectRoutedWorkspace(workspace, switchToEditor),
			]);
			await reopened.getByRole('tree').getByText('one', { exact: true }).waitFor({ timeout: 60000 });
			await reopened.getByRole('tree').getByText('two', { exact: true }).waitFor();
			await reopened.waitForFunction(title => document.title.includes(title), savedWorkspaceTitle);
			assert.equal(launched.windows().length, 2);
			report.reopenedOriginWorkspaceVerified = true;

			stage = 'recovering the saved origin after its window changes projects';
			const contextTab = reopened.locator('.tabs-container .tab.active').getByText('originating-editor-context.txt', { exact: true });
			await contextTab.waitFor();
			const scratchTab = reopened.locator('.tabs-container .tab.dirty').filter({ hasText: marker });
			if (await scratchTab.count()) {
				await scratchTab.press('Enter');
				await runWorkbenchCommand(reopened, 'Revert and Close Editor');
				await scratchTab.waitFor({ state: 'hidden' });
				await contextTab.waitFor();
			}
			await selectContextText(reopened);
			await expectRoutedWorkspace(agentsWorkspace, () => useModeSwitch(reopened));
			const switchedWorkspace = join(run, 'switched-project.code-workspace');
			await writeFile(switchedWorkspace, JSON.stringify({ folders: [{ path: anotherFolder }] }));
			await simulateNativeFocus(reopened);
			await expectRoutedWorkspace(switchedWorkspace, () => forwardToRunningApp(['--reuse-window', switchedWorkspace]));
			await reopened.getByRole('tree').getByText(anotherFile, { exact: true }).dblclick();
			const switchedTab = reopened.locator('.tabs-container .tab.active').getByText(anotherFile, { exact: true });
			await switchedTab.waitFor();
			const switchedEditor = reopened.locator('.part.editor .monaco-editor .view-lines').first();
			const switchedEdit = 'Unsaved edit in the newly opened saved workspace.';
			await switchedEditor.click();
			await reopened.keyboard.press('Control+Home');
			await reopened.keyboard.type(switchedEdit);
			await reopened.locator('.tabs-container .tab.active.dirty').getByText(anotherFile, { exact: true }).waitFor();
			await expectRoutedWorkspace(agentsWorkspace, () => useModeSwitch(reopened));
			await inlineDraftNotice.getByText('The switch requested other-project.', { exact: true }).waitFor({ timeout: 15000 });
			const assertSwitchedWorkspacePreserved = async () => {
				assert.ok(await switchedTab.isVisible(), 'Returning to the saved origin must keep the other workspace in the existing window');
				assert.ok((await switchedEditor.innerText()).replaceAll('\u00a0', ' ').includes(switchedEdit), 'Returning to the saved origin must preserve the other workspace’s unsaved edit');
				assert.equal(await readFile(anotherFilePath, 'utf8'), anotherFileContents);
			};
			const [restoredSavedOrigin] = await Promise.all([
				launched.waitForEvent('window', { timeout: 45000 }),
				expectRoutedWorkspace(workspace, switchToEditor),
			]);
			await restoredSavedOrigin.locator('.tabs-container .tab.active').getByText('originating-editor-context.txt', { exact: true }).waitFor();
			await restoredSavedOrigin.locator('.part.statusbar').getByText(selectionLabel, { exact: true }).waitFor();
			await assertSwitchedWorkspacePreserved();
			assert.ok((await input.innerText()).replaceAll('\u00a0', ' ').includes(draft));
			assert.equal((await checkWindows(stage)).length, 3);
			report.repurposedSavedWorkspaceRestored = true;
			await destroyFixtureWindow(restoredSavedOrigin);

			stage = 'return after the saved workspace no longer contains the Agents folder';
			const [agentsFolder, otherFolder] = selectsSecondRoot ? [folders[1], folders[0]] : folders;
			await writeFile(workspace, JSON.stringify({ folders: [{ path: otherFolder }] }));
			const [folderEditor] = await Promise.all([
				launched.waitForEvent('window', { timeout: 45000 }),
				expectRoutedWorkspace(agentsFolder, switchToEditor, 'uri'),
			]);
			await folderEditor.getByRole('tree').getByText('README.md', { exact: true }).waitFor();
			await assertSwitchedWorkspacePreserved();
			assert.equal((await checkWindows(stage)).length, 3);
			report.changedOriginFallbackVerified = true;
		}
	}
	if (keyboard) { assert.equal(report.keyboardActivations.length, 4, 'The native fixture must exercise both modes by shortcut and focused control'); }
	if (paletteSwitch) { assert.equal(report.paletteActivations.length, 2, 'The native fixture must find both mode switches in the Command Palette'); }
	if (protocol) {
		stage = 'launch environment trace';
		const logFiles = await readdir(join(profile, 'logs'), { recursive: true, withFileTypes: true });
		const traceLogs = await Promise.all(logFiles.filter(file => file.isFile() && ['main.log', 'renderer.log', 'ptyhost.log'].includes(file.name))
			.map(file => readFile(join(file.parentPath, file.name), 'utf8')));
		assert.ok(traceLogs.some(log => log.includes('Received a launch request from another instance')));
		assert.ok(traceLogs.some(log => log.includes('workbench#open(): with configuration')));
		assert.ok(traceLogs.every(log => !log.includes(launchToken)), 'Launch and workbench traces must exclude the process environment');
		report.launchEnvironmentExcludedFromTrace = true;
	}
} catch (error) {
	const failedApp = app;
	await Promise.all((failedApp?.windows() ?? []).map(async (page, i) => {
		try {
			const window = await failedApp!.browserWindow(page);
			try {
				const png = await window.evaluate(async nativeWindow => (await nativeWindow.webContents.capturePage(undefined, { stayHidden: true, stayAwake: true })).toPNG().toString('base64'));
				await writeFile(join(run, `failure-${i}.png`), Buffer.from(png, 'base64'));
			} finally { await window.dispose(); }
			await writeFile(join(run, `failure-${i}.json`), JSON.stringify(await page.evaluate(() => ({ title: document.title, text: document.body.innerText.slice(0, 12000) })), null, 2));
		} catch { /* Retain other diagnostics if a failed window already closed. */ }
	}));
	await writeFile(join(run, 'failure.json'), JSON.stringify({ stage, message: error instanceof Error ? error.message : String(error) }, null, 2));
	throw error;
} finally {
	if (deadline) { clearTimeout(deadline); }
	runningResponse.resolve();
	try {
		await app?.close();
	} finally {
		report.modelRequestCount = mock.getRequests().length;
		try { await mock.close(); }
		finally {
			await writeFile(join(run, 'report.json'), JSON.stringify(report, null, 2));
			console.log(`Artifacts: ${run}`);
		}
	}
}
