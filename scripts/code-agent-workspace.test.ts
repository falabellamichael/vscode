/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { spawnSync } = require('node:child_process');
const { test } = require('node:test');
const { createLaunchOptions, createPrimaryLaunchOptions } = require('./code-agent-workspace.ts');

test('routes browser callbacks to the normal Signal-C profile and runtime', () => {
	const root = path.resolve('fork with spaces');
	const appData = path.resolve('user data');
	const inherited = { APPDATA: appData, VSCODE_PORTABLE: 'another-build', VSCODE_IPC_HOOK_CLI: 'another-editor' };
	const normal = createPrimaryLaunchOptions(root, undefined, false, inherited);
	const uri = 'signal-c://dynamicauthprovider/accounts.google.com/redirect?nonce=fixture&windowId=1';
	const callback = createPrimaryLaunchOptions(root, undefined, false, inherited, uri);
	assert.strictEqual(normal.profile, path.join(appData, 'Signal-C'));
	assert.strictEqual(callback.executable, normal.executable);
	assert.deepStrictEqual(callback.environment, normal.environment);
	assert.strictEqual(callback.environment.VSCODE_DEV, '1');
	assert.strictEqual(callback.environment.VSCODE_SIGNAL_C_PRIMARY, '1');
	assert.strictEqual(callback.environment.VSCODE_PORTABLE, undefined);
	assert.deepStrictEqual(callback.args, [...normal.args.slice(0, -1), '--open-url', '--', uri]);
	for (const invalid of ['', 'https://example.test', 'signal-c://test/" --flag', 'signal-c://test/\n']) {
		assert.throws(() => createPrimaryLaunchOptions(root, undefined, false, inherited, invalid), /Invalid Signal-C callback/);
	}
});

test('isolates editor IPC and agent homes without mutating the launching environment', () => {
	const inherited = { PATH: 'keep-path', SystemRoot: 'keep-system-root', VSCODE_IPC_HOOK_CLI: 'other-editor', vscode_portable: 'other-profile', ELECTRON_RUN_AS_NODE: '1', codex_home: 'other-codex', claude_config_dir: 'other-claude' };
	const original = { ...inherited };
	const root = path.resolve('fork with spaces');
	const launch = createLaunchOptions(root, root, false, inherited);
	assert.deepStrictEqual(inherited, original);
	assert.strictEqual(launch.environment.PATH, inherited.PATH);
	assert.strictEqual(launch.environment.SystemRoot, inherited.SystemRoot);
	assert.strictEqual(launch.environment.CODEX_HOME, path.join(root, '.build', 'agent-workspace-codex'));
	assert.strictEqual(launch.environment.CLAUDE_CONFIG_DIR, path.join(root, '.build', 'agent-workspace-claude'));
	assert.strictEqual(launch.args[launch.args.indexOf('--agent-codex-home') + 1], launch.environment.CODEX_HOME);
	assert.strictEqual(launch.args[launch.args.indexOf('--agent-claude-config-dir') + 1], launch.environment.CLAUDE_CONFIG_DIR);
	assert.strictEqual(launch.environment.VSCODE_IPC_HOOK_CLI, undefined);
	assert.strictEqual(launch.environment.vscode_portable, undefined);
	assert.strictEqual(launch.environment.ELECTRON_RUN_AS_NODE, undefined);
	assert.strictEqual(launch.environment.codex_home, undefined);
	assert.strictEqual(launch.environment.claude_config_dir, undefined);
});

test('makes the installed official Claude CLI a fallback without replacing PATH or copying credentials', { skip: process.platform !== 'win32' }, () => {
	const root = fs.mkdtempSync(path.join(os.tmpdir(), 'signal-c-claude-cli-'));
	const sdk = path.join(root, 'node_modules', '@anthropic-ai', 'claude-agent-sdk');
	const nativeName = `@anthropic-ai/claude-agent-sdk-${process.platform}-${process.arch}`;
	const native = path.join(root, 'node_modules', ...nativeName.split('/'));
	assert.strictEqual(path.dirname(path.resolve(root)), path.resolve(os.tmpdir()));
	try {
		fs.mkdirSync(sdk, { recursive: true });
		fs.mkdirSync(native, { recursive: true });
		fs.writeFileSync(path.join(root, 'package.json'), '{}');
		fs.writeFileSync(path.join(sdk, 'package.json'), JSON.stringify({ name: '@anthropic-ai/claude-agent-sdk', main: 'sdk.mjs', optionalDependencies: { [nativeName]: 'test' } }));
		fs.writeFileSync(path.join(sdk, 'sdk.mjs'), '');
		fs.writeFileSync(path.join(native, 'package.json'), JSON.stringify({ name: nativeName }));
		fs.writeFileSync(path.join(native, 'claude.exe'), '');
		const inherited = { Path: 'existing-client-directory', CLAUDE_CONFIG_DIR: 'another-client-home' };
		const launch = createLaunchOptions(root, root, false, inherited);
		assert.ok(launch.args.includes('--agent-cli-fallback-dir'));
		assert.strictEqual(launch.args[launch.args.indexOf('--agent-cli-fallback-dir') + 1], native);
		assert.strictEqual(launch.environment.Path, `existing-client-directory${path.delimiter}${native}`);
		assert.strictEqual(launch.environment.PATH, undefined);
		assert.strictEqual(launch.environment.CLAUDE_CONFIG_DIR, path.join(root, '.build', 'agent-workspace-claude'));
		const optimized = createLaunchOptions(root, root, false, inherited, true);
		assert.strictEqual(optimized.environment.VSCODE_AGENT_HOST_CLAUDE_SDK_ROOT, root);
		assert.strictEqual(optimized.environment.VSCODE_AGENT_HOST_CODEX_SDK_ROOT, undefined);
		assert.deepStrictEqual(inherited, { Path: 'existing-client-directory', CLAUDE_CONFIG_DIR: 'another-client-home' });
		const alreadyAvailable = createLaunchOptions(root, root, false, { Path: native });
		assert.strictEqual(alreadyAvailable.environment.Path, native);
		fs.unlinkSync(path.join(native, 'claude.exe'));
		const unavailable = createLaunchOptions(root, root, false, inherited);
		assert.strictEqual(unavailable.environment.Path, inherited.Path);
		assert.strictEqual(unavailable.args.includes('--agent-cli-fallback-dir'), false);
	} finally {
		assert.strictEqual(path.dirname(path.resolve(root)), path.resolve(os.tmpdir()));
		fs.rmSync(root, { recursive: true, force: true });
	}
});

test('keeps workspace arguments separate from the three isolated editor storage paths in both modes', () => {
	const root = path.resolve('fork with spaces');
	const workspace = path.resolve('project with spaces');
	for (const ide of [false, true]) {
		const launch = createLaunchOptions(root, workspace, ide, {});
		assert.strictEqual(launch.args.includes('--agents'), !ide);
		assert.ok(launch.args.includes('--agent-workspace-preview'));
		assert.ok(launch.args.includes(workspace));
		for (const [flag, directory] of [['--user-data-dir', 'agent-workspace-profile'], ['--shared-data-dir', 'agent-workspace-shared'], ['--extensions-dir', 'agent-workspace-extensions']]) {
			assert.strictEqual(launch.args[launch.args.indexOf(flag) + 1], path.join(root, '.build', directory));
		}
	}
});

test('optimized desktop uses bundled code while preserving the existing workspace and isolated profile destinations', () => {
	const root = path.resolve('fork with spaces');
	const workspace = path.resolve('project with spaces');
	const inherited = { VSCODE_DEV: '1', VSCODE_IPC_HOOK_CLI: 'another-editor', NODE_ENV: 'development', Path: 'keep' };
	for (const ide of [false, true]) {
		const development = createLaunchOptions(root, workspace, ide, inherited);
		const optimized = createLaunchOptions(root, workspace, ide, inherited, true);
		assert.strictEqual(optimized.executable, path.join(path.dirname(root), `VSCode-win32-${process.arch}`, 'Signal-C.exe'));
		assert.deepStrictEqual(optimized.args, development.args.slice(1));
		assert.strictEqual(optimized.environment.VSCODE_DEV, undefined);
		assert.strictEqual(optimized.environment.NODE_ENV, 'production');
		assert.strictEqual(optimized.environment.VSCODE_IPC_HOOK_CLI, undefined);
		assert.strictEqual(optimized.environment.VSCODE_AGENT_HOST_CODEX_SDK_ROOT, undefined);
		assert.strictEqual(optimized.profile, development.profile);
		assert.deepStrictEqual(optimized.agentHomes, development.agentHomes);
	}
	assert.deepStrictEqual(inherited, { VSCODE_DEV: '1', VSCODE_IPC_HOOK_CLI: 'another-editor', NODE_ENV: 'development', Path: 'keep' });
});

test('prefers the Signal-C runtime and supports an existing development runtime', () => {
	const temporary = fs.mkdtempSync(path.join(os.tmpdir(), 'signal-c-runtime-'));
	const build = path.join(temporary, '.build');
	const runtime = path.join(build, 'electron');
	fs.mkdirSync(runtime, { recursive: true });
	const branded = path.join(runtime, 'Signal-C.exe');
	const previous = path.join(runtime, 'Code - OSS.exe');
	try {
		assert.strictEqual(createLaunchOptions(temporary, temporary, false, {}).executable, branded);
		fs.writeFileSync(previous, '');
		assert.strictEqual(createLaunchOptions(temporary, temporary, false, {}).executable, previous);
		fs.writeFileSync(branded, '');
		assert.strictEqual(createLaunchOptions(temporary, temporary, false, {}).executable, branded);
	} finally {
		for (const file of [branded, previous]) { if (fs.existsSync(file)) { fs.unlinkSync(file); } }
		fs.rmdirSync(runtime);
		fs.rmdirSync(build);
		fs.rmdirSync(temporary);
	}
});

test('selects a side-by-side optimized package while preserving workspace, profiles and agent homes', () => {
	const root = path.resolve('fork with spaces');
	const workspace = path.resolve('project with spaces');
	const build = path.resolve('updated package with spaces');
	for (const ide of [false, true]) {
		const original = createLaunchOptions(root, workspace, ide, {}, true);
		const updated = createLaunchOptions(root, workspace, ide, {}, true, build);
		assert.strictEqual(updated.executable, path.join(build, 'Signal-C.exe'));
		assert.deepStrictEqual(updated.args, original.args);
		assert.strictEqual(updated.profile, original.profile);
		assert.deepStrictEqual(updated.agentHomes, original.agentHomes);
		assert.deepStrictEqual(updated.environment, original.environment);
	}
});

test('requires optimized mode when selecting a package', { skip: process.platform !== 'win32' }, () => {
	const result = spawnSync(process.execPath, [path.join(__dirname, 'code-agent-workspace.ts'), '--dry-run', '--optimized-build', 'updated package'], { encoding: 'utf8' });
	assert.notStrictEqual(result.status, 0);
	assert.match(result.stderr, /Use --optimized --optimized-build/);
});

test('dry-run accepts a workspace with spaces without launching an editor or creating a profile', { skip: process.platform !== 'win32' }, () => {
	const temporary = fs.mkdtempSync(path.join(os.tmpdir(), 'agent-workspace-launch-'));
	const workspace = path.join(temporary, 'Project with spaces');
	fs.mkdirSync(workspace);
	try {
		const result = spawnSync(process.execPath, [path.join(__dirname, 'code-agent-workspace.ts'), '--dry-run', '--ide', '--workspace', workspace], { encoding: 'utf8' });
		assert.strictEqual(result.status, 0, result.stderr);
		const launch = JSON.parse(result.stdout);
		assert.ok(launch.args.includes(workspace));
		assert.strictEqual(launch.args.includes('--agents'), false);
		assert.deepStrictEqual(fs.readdirSync(temporary), ['Project with spaces']);
	} finally {
		fs.rmdirSync(workspace);
		fs.rmdirSync(temporary);
	}
});
