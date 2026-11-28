/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

const fs = require('fs') as typeof import('fs');
const path = require('path') as typeof import('path');
const { spawn } = require('child_process') as typeof import('child_process');
const { createRequire } = require('node:module') as typeof import('node:module');
const minimist = require('minimist') as typeof import('minimist');
const product = JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'product.json'), 'utf8')) as { nameShort: string };

/** Use the installed, unmodified official CLI as a development fallback without installing or copying it. */
function bundledClaudeDirectory(repoRoot: string): string | undefined {
	if (process.platform !== 'win32' || !fs.existsSync(path.join(repoRoot, 'node_modules', '@anthropic-ai', 'claude-agent-sdk', 'package.json'))) { return undefined; }
	try {
		const fromRepo = createRequire(path.join(repoRoot, 'package.json'));
		const sdkEntry = fromRepo.resolve('@anthropic-ai/claude-agent-sdk');
		const nativePackage = `@anthropic-ai/claude-agent-sdk-${process.platform}-${process.arch}`;
		const manifest = JSON.parse(fs.readFileSync(path.join(path.dirname(sdkEntry), 'package.json'), 'utf8')) as { optionalDependencies?: Record<string, string> };
		if (!manifest.optionalDependencies?.[nativePackage]) { return undefined; }
		const nativeManifest = createRequire(sdkEntry).resolve(`${nativePackage}/package.json`);
		const directory = path.dirname(nativeManifest);
		return fs.existsSync(path.join(directory, 'claude.exe')) ? directory : undefined;
	} catch { return undefined; }
}

function createLaunchOptions(repoRoot: string, workspace: string, ide: boolean, inheritedEnvironment: NodeJS.ProcessEnv, optimized = false, optimizedBuild?: string) {
	const buildRoot = path.join(repoRoot, '.build');
	const profile = path.join(buildRoot, 'agent-workspace-profile');
	const codexHome = path.join(buildRoot, 'agent-workspace-codex');
	const claudeHome = path.join(buildRoot, 'agent-workspace-claude');
	const environment = { ...inheritedEnvironment };
	const executable = path.join(buildRoot, 'electron', `${product.nameShort}.exe`);
	const previousExecutable = path.join(buildRoot, 'electron', 'Code - OSS.exe');
	// Editor-injected IPC and SDK overrides belong to the launching editor. Keep
	// this development fork on its own runtime, storage, and agent configuration.
	for (const name of Object.keys(environment)) {
		if (/^(?:VSCODE_|ELECTRON_RUN_AS_NODE$|CODEX_HOME$|CLAUDE_CONFIG_DIR$)/i.test(name)) {
			delete environment[name];
		}
	}
	Object.assign(environment, {
		NODE_ENV: optimized ? 'production' : 'development', ...(!optimized ? { VSCODE_DEV: '1' } : {}), VSCODE_CLI: '1',
		CODEX_HOME: codexHome,
		CLAUDE_CONFIG_DIR: claudeHome,
	});
	if (optimized) {
		// Locally packaged builds have no published SDK catalog. Reuse this
		// checkout's installed SDKs through the existing explicit root overrides.
		for (const [name, variable] of [['@openai/codex', 'VSCODE_AGENT_HOST_CODEX_SDK_ROOT'], ['@anthropic-ai/claude-agent-sdk', 'VSCODE_AGENT_HOST_CLAUDE_SDK_ROOT']]) {
			if (fs.existsSync(path.join(repoRoot, 'node_modules', name, 'package.json'))) { environment[variable] = repoRoot; }
		}
	}
	const claudeDirectory = bundledClaudeDirectory(repoRoot);
	if (claudeDirectory) {
		const pathKey = Object.keys(environment).find(key => key.toLowerCase() === 'path') ?? 'PATH';
		const currentPath = environment[pathKey] ?? '';
		if (!currentPath.split(path.delimiter).some(entry => entry.toLowerCase() === claudeDirectory.toLowerCase())) {
			// Append so a separately installed CLI and explicit executable settings keep precedence.
			environment[pathKey] = currentPath ? `${currentPath}${path.delimiter}${claudeDirectory}` : claudeDirectory;
		}
	}
	return {
		// Existing development runtimes remain usable until native assets are rebuilt.
		executable: optimized ? path.join(optimizedBuild ? path.resolve(optimizedBuild) : path.join(path.dirname(repoRoot), `VSCode-win32-${process.arch}`), `${product.nameShort}.exe`) : fs.existsSync(executable) || !fs.existsSync(previousExecutable) ? executable : previousExecutable,
		args: [...(!optimized ? [repoRoot] : []), ...(ide ? [] : ['--agents']), workspace,
			'--agent-workspace-preview',
			'--user-data-dir', profile,
			'--shared-data-dir', path.join(buildRoot, 'agent-workspace-shared'),
			'--extensions-dir', path.join(buildRoot, 'agent-workspace-extensions'),
			'--agent-codex-home', codexHome, '--agent-claude-config-dir', claudeHome,
			...(claudeDirectory ? ['--agent-cli-fallback-dir', claudeDirectory] : []),
			'--disable-extension=vscode.vscode-api-tests'],
		profile,
		agentHomes: [codexHome, claudeHome],
		environment,
	};
}

function createPrimaryLaunchOptions(repoRoot: string, workspace: string | undefined, ide: boolean, inheritedEnvironment: NodeJS.ProcessEnv, callback?: string) {
	const { executable, environment } = createLaunchOptions(repoRoot, repoRoot, ide, inheritedEnvironment);
	if (!environment.APPDATA) {
		throw new Error('APPDATA is required to open your Signal-C profile.');
	}
	const profile = path.join(environment.APPDATA, 'Signal-C');
	delete environment.CODEX_HOME;
	delete environment.CLAUDE_CONFIG_DIR;
	if (inheritedEnvironment.CODEX_HOME) { environment.CODEX_HOME = inheritedEnvironment.CODEX_HOME; }
	if (inheritedEnvironment.CLAUDE_CONFIG_DIR) { environment.CLAUDE_CONFIG_DIR = inheritedEnvironment.CLAUDE_CONFIG_DIR; }
	environment.VSCODE_SIGNAL_C_PRIMARY = '1';
	const args = [repoRoot, '--user-data-dir', profile, '--disable-extension=vscode.vscode-api-tests'];
	if (callback !== undefined) {
		if (!/^signal-c:\/\//i.test(callback) || /["\u0000-\u0020]/.test(callback) || !URL.canParse(callback)) {
			throw new Error('Invalid Signal-C callback URL.');
		}
		args.push('--open-url', '--', callback);
	} else {
		if (!ide) { args.push('--agents'); }
		if (workspace) { args.push(workspace); }
	}
	return { executable, args, profile, agentHomes: [], environment };
}

function main() {
	const args = minimist(process.argv.slice(2), { boolean: ['help', 'ide', 'optimized', 'dry-run', 'primary'], string: ['workspace', 'optimized-build', 'open-url'] });
	if (args.help) {
		console.log('node scripts/code-agent-workspace.ts [--primary | --optimized [--optimized-build <folder>]] [--ide] [--workspace <folder>] [--dry-run]\nLaunch the preview by default. --primary runs the checkout with the normal Signal-C profile. --optimized uses the packaged desktop build with the preview profile.');
		return;
	}
	if (process.platform !== 'win32') {
		throw new Error('This development launcher currently supports Windows.');
	}
	const unknown = Object.keys(args).filter(key => !['_', 'help', 'ide', 'optimized', 'optimized-build', 'dry-run', 'workspace', 'primary', 'open-url'].includes(key));
	if (unknown.length || args._.length) {
		throw new Error('Use --workspace <folder> to select a workspace; additional editor arguments are not forwarded.');
	}
	if (args['optimized-build'] !== undefined && (!args.optimized || !args['optimized-build'].trim())) {
		throw new Error('Use --optimized --optimized-build <folder> to select a packaged desktop build.');
	}
	const callback = args['open-url'];
	if ((args.primary && args.optimized) || (callback !== undefined && (!args.primary || typeof callback !== 'string' || args.ide || args.workspace !== undefined))) {
		throw new Error('Use --primary --open-url <Signal-C URL> to handle a desktop callback.');
	}
	const repoRoot = path.resolve(__dirname, '..');
	const workspace = path.resolve(args.workspace ?? repoRoot);
	if (!fs.statSync(workspace).isDirectory()) {
		throw new Error('The workspace must be an existing directory.');
	}
	const options = args.primary
		? createPrimaryLaunchOptions(repoRoot, args.workspace === undefined ? undefined : workspace, args.ide, process.env, callback)
		: createLaunchOptions(repoRoot, workspace, args.ide, process.env, args.optimized, args['optimized-build']);
	if (args['dry-run']) {
		const launchArgs = callback === undefined ? options.args : [...options.args.slice(0, -1), '<Signal-C callback>'];
		console.log(JSON.stringify({ executable: options.executable, args: launchArgs }, null, 2));
		return;
	}
	if (!fs.existsSync(options.executable)) {
		throw new Error(args.optimized ? 'The optimized desktop build is missing. Run the vscode-win32 build task for this architecture before launching with --optimized.' : 'The development Electron runtime is missing. Prepare the native build before launching.');
	}
	for (const directory of [options.profile, ...options.agentHomes]) {
		fs.mkdirSync(directory, { recursive: true });
	}
	if (!args.primary) {
		// Seed only a new isolated profile. Later user choices are preserved.
		const settings = path.join(options.profile, 'User', 'settings.json');
		fs.mkdirSync(path.dirname(settings), { recursive: true });
		try {
			fs.writeFileSync(settings, JSON.stringify({
				'chat.agentHost.codexAgent.enabled': true,
				'chat.agentHost.claudeAgent.enabled': false,
			}, null, '\t') + '\n', { flag: 'wx' });
		} catch (error) {
			if ((error as NodeJS.ErrnoException).code !== 'EEXIST') { throw error; }
		}
	}
	const child = spawn(options.executable, options.args, { cwd: repoRoot, env: options.environment, detached: true, stdio: 'ignore' });
	child.on('error', error => { console.error(error.message); process.exitCode = 1; });
	child.on('spawn', () => {
		if (callback === undefined) {
			const mode = args.primary ? 'primary checkout' : args.optimized ? 'optimized desktop' : 'development';
			console.log(`Opened Signal-C ${args.ide ? 'IDE' : 'Agents'} ${mode} window (PID ${child.pid}).`);
		}
		child.unref();
	});
}

module.exports = { createLaunchOptions, createPrimaryLaunchOptions };
if (require.main === module) {
	try { main(); } catch (error) { console.error(error instanceof Error ? error.message : error); process.exitCode = 1; }
}
