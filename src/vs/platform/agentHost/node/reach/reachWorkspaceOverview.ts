/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { execFile } from 'child_process';
import { lstat, open } from 'fs/promises';
import { promisify } from 'util';
import { match } from '../../../../base/common/glob.js';
import { escapeMarkdownSyntaxTokens } from '../../../../base/common/htmlContent.js';
import { parse } from '../../../../base/common/json.js';
import { basename, dirname, join } from '../../../../base/common/path.js';
import { URI } from '../../../../base/common/uri.js';
import { rgDiskPath } from '../../../../base/node/ripgrep.js';
import { localize } from '../../../../nls.js';

const run = promisify(execFile);
const project = /\b(?:workspace|worktree|repository|repo|project|codebase)\b/i;
const overview = /\b(?:summari[sz]e|describe|explain|outline|overview of|summary of|tour of|walk me through|tell me about|what(?:'s| is) in|what(?:'s| is)|what does)\s+(?:(?:the|this|that|my|our|current|open|selected|local|whole|entire)\s+)*(?:workspace|worktree|repository|repo|project|codebase)\b/i;
const knownDirectory = /^(?:app|apps|src|lib|packages|extensions|tests?|docs?|server|client|frontend|backend|android|ios|scripts|tools)$/i;
const manifest = /^(?:package\.json|pyproject\.toml|Cargo\.toml|go\.mod|pom\.xml|build\.gradle(?:\.kts)?|CMakeLists\.txt|composer\.json|Gemfile|[^/]{1,80}\.csproj)$/i;

export function wantsReachWorkspaceOverview(prompt: string): boolean {
	return project.test(prompt) && (overview.test(prompt) || /\b(?:workspace|worktree|repository|repo|project|codebase)\s+(?:summary|overview|structure|about|do(?:es)?)\b/i.test(prompt));
}

/** Read only a bounded prefix of a regular file. Never open an enumerated symlink or special file. */
async function prefix(path: string, size: number): Promise<string | undefined> {
	try {
		const entry = await lstat(path);
		if (!entry.isFile()) { return undefined; }
		const handle = await open(path, 'r');
		try {
			const actual = await handle.stat();
			if (!actual.isFile() || actual.dev !== entry.dev || actual.ino !== entry.ino) { return undefined; }
			const bytes = Buffer.alloc(size);
			const { bytesRead } = await handle.read(bytes, 0, size, 0);
			if (bytes.subarray(0, bytesRead).includes(0)) { return undefined; }
			return new TextDecoder('utf-8').decode(bytes.subarray(0, bytesRead));
		} finally { await handle.close(); }
	} catch { return undefined; }
}

function description(value: unknown): string | undefined {
	if (typeof value !== 'string') { return undefined; }
	const line = value.replace(/!?\[([^\]]*)\]\([^)]*\)/g, '$1').replace(/\s+/g, ' ').trim();
	if (!line || /(?:https?:\/\/|\b(?:password|passwd|secret|bearer|access[_ -]?token|api[_ -]?key|private[_ -]?key)\b|[A-Za-z0-9_+/=-]{45,}|[<>\u0000-\u001f])/.test(line.toLowerCase())) { return undefined; }
	return escapeMarkdownSyntaxTokens(line.slice(0, 360));
}

async function exclusionPatterns(root: string): Promise<string[]> {
	const directory = join(root, '.vscode');
	if (!(await lstat(directory).catch(() => undefined))?.isDirectory()) { return []; }
	const settings = parse(await prefix(join(directory, 'settings.json'), 64 * 1024) ?? '{}');
	const result: string[] = [];
	for (const key of ['files.exclude', 'search.exclude']) {
		for (const [pattern, value] of Object.entries(settings?.[key] ?? {})) {
			if (value && pattern.length <= 512) { result.push(pattern); }
		}
	}
	return result;
}

async function rootOverview(root: URI, signal: AbortSignal): Promise<string | undefined> {
	if (root.scheme !== 'file' || root.authority || dirname(root.fsPath) === root.fsPath) { return undefined; }
	const rootName = description(basename(root.fsPath)) ?? localize('reach.workspace.selectedFolder', "Selected folder");
	const lines = [`**${rootName}**`];
	try {
		// The general host file search follows symlinks. For an overview, walk only two levels and leave symlinks untouched.
		const { stdout } = await run(await rgDiskPath(), ['--files', '--null', '--no-config', '--no-require-git', '--max-depth', '2', '--glob', '!node_modules/', '--glob', '!vendor/', '--glob', '!dist/', '--glob', '!build/', '--glob', '!out/', '--glob', '!target/'], {
			cwd: root.fsPath, timeout: 4000, maxBuffer: 512 * 1024, windowsHide: true, signal,
		});
		const excludes = await exclusionPatterns(root.fsPath);
		const files = stdout.split('\0').map(path => path.replace(/\\/g, '/').replace(/^\.\//, '')).filter(path => path && !excludes.some(pattern => match(pattern, path) || match(pattern, path.split('/')[0])));
		const atRoot = files.filter(path => !path.includes('/'));
		const docs = atRoot.find(name => /^readme(?:\.md|\.txt|\.rst)?$/i.test(name));
		if (docs) {
			const readme = (await prefix(join(root.fsPath, docs), 8192))?.split(/\r?\n/);
			const heading = readme?.find(line => /^#\s+/.test(line))?.replace(/^#\s+/, '');
			const intro = readme?.find(line => /^[\p{L}]/u.test(line) && line !== heading && description(line));
			const text = [description(heading), description(intro)].filter(Boolean).join(' — ');
			if (text) { lines.push(`- \`${docs}\`: ${text}`); }
		}
		if (atRoot.includes('package.json')) {
			const pkg = parse(await prefix(join(root.fsPath, 'package.json'), 64 * 1024) ?? '{}');
			const text = [description(pkg?.name), description(pkg?.description)].filter(Boolean).join(' — ');
			if (text) { lines.push(`- \`package.json\`: ${text}`); }
		}
		const manifests = atRoot.filter(name => manifest.test(name) && /^[\w.-]+$/.test(name)).sort().slice(0, 12);
		if (manifests.length) { lines.push(localize('reach.workspace.manifests', "- Project files: {0}", manifests.map(name => `\`${name}\``).join(', '))); }
		const directories = [...new Set(files.filter(path => path.includes('/')).map(path => path.split('/')[0]).filter(dir => knownDirectory.test(dir)))].sort();
		if (directories.length) { lines.push(localize('reach.workspace.directories', "- Folders: {0}", directories.map(name => `\`${name}/\``).join(', '))); }
	} catch {
		signal.throwIfAborted();
	}
	if (lines.length === 1) { lines.push(localize('reach.workspace.noProjectEvidence', "No readable top-level project description was found. Attach a README or a source file for a more detailed summary.")); }
	return lines.join('\n');
}

/** Collect source-labelled facts only from the local folders attached to this chat. */
export async function readReachWorkspaceOverview(roots: readonly URI[], signal: AbortSignal): Promise<string> {
	const sections: string[] = [];
	for (const root of roots.slice(0, 4)) {
		signal.throwIfAborted();
		const section = await rootOverview(root, signal);
		if (section) { sections.push(section); }
	}
	if (!sections.length) {
		return localize('reach.workspace.noFolder', "This chat has no local workspace folder to summarize. Choose a workspace for the chat or attach its README.");
	}
	return localize('reach.workspace.overview', "Workspace overview") + '\n\n' + sections.join('\n\n');
}

