/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { constants } from 'fs';
import { copyFile, link, lstat, mkdir, open, realpath, unlink } from 'fs/promises';
import { isValidBasename } from '../../../base/common/extpath.js';
import { dirname, join, relative } from '../../../base/common/path.js';
import { URI } from '../../../base/common/uri.js';
import { generateUuid } from '../../../base/common/uuid.js';
import { localize } from '../../../nls.js';
import type { ReachWorkspaceFileRequest, ReachWorkspaceFileResult } from '../common/native.js';

const byteLimit = 1024 * 1024;

function invalidPath(): Error {
	return new Error(localize('reach.file.invalidPath', "Use a relative file path inside the chat's local workspace folder."));
}

async function verifiedDirectory(root: string, parts: readonly string[], create: boolean): Promise<string> {
	let dir = root;
	for (const part of parts) {
		dir = join(dir, part);
		if (create) {
			try { await mkdir(dir); } catch (err) { if ((err as NodeJS.ErrnoException).code !== 'EEXIST') { throw err; } }
		}
		if (!(await lstat(dir)).isDirectory() || relative(dir, await realpath(dir)) !== '') { throw invalidPath(); }
	}
	return dir;
}

export async function accessReachWorkspaceFile(root: URI, request: ReachWorkspaceFileRequest): Promise<ReachWorkspaceFileResult> {
	if (root.scheme !== 'file' || root.authority || !root.path || !(await lstat(root.fsPath)).isDirectory()) { throw invalidPath(); }
	const parts = typeof request.path === 'string' ? request.path.replace(/^(?:\.[\\/])+/, '').split(/[\\/]/) : [];
	if (!parts.length || parts.some(part => !isValidBasename(part, true) || /[\u0000-\u001f]/.test(part))) { throw invalidPath(); }
	if (request.encoding !== undefined && request.encoding !== 'utf8' && request.encoding !== 'base64') { throw new Error(localize('reach.file.encoding', "Choose utf8 or base64 for the file encoding.")); }
	const canonicalRoot = await realpath(root.fsPath);
	const parents = parts.slice(0, -1);
	const uri = URI.joinPath(root, ...parts).toString();
	if (request.operation === 'read') {
		const path = join(await verifiedDirectory(canonicalRoot, parents, false), parts[parts.length - 1]);
		const entry = await lstat(path);
		if (!entry.isFile() || relative(path, await realpath(path)) !== '') { throw invalidPath(); }
		const file = await open(path, constants.O_RDONLY | (constants.O_NOFOLLOW ?? 0) | (constants.O_NONBLOCK ?? 0));
		try {
			const stat = await file.stat();
			if (stat.ino !== entry.ino || stat.dev !== entry.dev || !stat.isFile()) { throw invalidPath(); }
			if (stat.size > byteLimit) { throw new Error(localize('reach.file.largeRead', "This file exceeds the 1 MiB read limit.")); }
			// Read through the verified handle, bounded even if the file subsequently grows.
			const bytes = Buffer.alloc(stat.size);
			let offset = 0;
			while (offset < bytes.length) {
				const { bytesRead } = await file.read(bytes, offset, bytes.length - offset, offset);
				if (!bytesRead) { break; }
				offset += bytesRead;
			}
			const data = bytes.subarray(0, offset);
			let content: string;
			try { content = request.encoding === 'base64' ? data.toString('base64') : new TextDecoder('utf-8', { fatal: true, ignoreBOM: true }).decode(data); }
			catch { throw new Error(localize('reach.file.binary', "Use base64 to read this binary file.")); }
			return { uri, bytes: offset, content };
		} finally { await file.close(); }
	}
	if (request.operation !== 'create' || typeof request.content !== 'string') { throw new Error(localize('reach.file.missingContent', "Supply the file content to create a file.")); }
	const encoding = request.encoding ?? 'utf8';
	if (Buffer.byteLength(request.content, encoding) > byteLimit) { throw new Error(localize('reach.file.largeCreate', "This content exceeds the 1 MiB creation limit.")); }
	const data = Buffer.from(request.content, encoding);
	if (encoding === 'base64' && data.toString('base64') !== request.content) { throw new Error(localize('reach.file.badBase64', "Provide valid, padded base64 content.")); }
	const parent = await verifiedDirectory(canonicalRoot, parents, true);
	const target = join(parent, parts[parts.length - 1]);
	const temporary = join(parent, '.reach-' + generateUuid() + '.tmp');
	const file = await open(temporary, 'wx', 0o600);
	try {
		try { await file.writeFile(data); } finally { await file.close(); }
		if (relative(dirname(target), await realpath(parent)) !== '') { throw invalidPath(); }
		try {
			// Publish only complete bytes. Both paths claim the user's filename exclusively.
			await link(temporary, target);
		} catch (err) {
			if (!['ENOTSUP', 'EOPNOTSUPP', 'EPERM', 'ENOSYS'].includes((err as NodeJS.ErrnoException).code ?? '')) { throw err; }
			await copyFile(temporary, target, constants.COPYFILE_EXCL);
		}
	} catch (err) {
		if ((err as NodeJS.ErrnoException).code === 'EEXIST') { throw new Error(localize('reach.file.exists', "A file already exists at {0}.", request.path)); }
		throw err;
	} finally { await unlink(temporary).catch(() => { /* A completed file remains published if temporary-file cleanup fails. */ }); }
	return { uri, bytes: data.length };
}
