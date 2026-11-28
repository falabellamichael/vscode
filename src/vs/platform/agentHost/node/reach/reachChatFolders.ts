/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { isEqualOrParent } from '../../../../base/common/extpath.js';
import { isAbsolute, relative } from '../../../../base/common/path.js';
import { isLinux } from '../../../../base/common/platform.js';
import { basename, extUriBiasedIgnorePathCase as resources } from '../../../../base/common/resources.js';
import { URI } from '../../../../base/common/uri.js';
import { localize } from '../../../../nls.js';
import { getAttachedLocalPaths } from '../../common/attachedLocalPaths.js';
import { MessageAttachmentKind, type MessageAttachment, type MessageResourceAttachment } from '../../common/state/protocol/channels-chat/state.js';

export function isFolderAttachment(attachment: MessageAttachment): attachment is MessageResourceAttachment & { displayKind: 'directory' } {
	return attachment.type === MessageAttachmentKind.Resource && attachment.displayKind === 'directory';
}

export interface AttachedChatFolder {
	readonly path: string;
	readonly workspaceFolder?: string;
	readonly absolutePath: string;
}

/** Qualify model paths before approval. The editor still validates the resulting path and existing files. */
export function qualifyFolderCreate(folder: AttachedChatFolder, args: Record<string, unknown>): string | undefined {
	if (typeof args.path !== 'string') { return undefined; }
	if (args.workspaceFolder !== undefined && folder.workspaceFolder !== undefined && args.workspaceFolder !== folder.workspaceFolder) {
		return 'Use the attached destination workspaceFolder: ' + JSON.stringify(folder.workspaceFolder);
	}
	let path = args.path.replace(/\\/g, '/').replace(/^(?:\.\/)+/, '');
	if (!path || path.split('/').includes('..')) { return 'Use a file path inside the attached folder.'; }
	if (isAbsolute(path)) {
		path = relative(folder.absolutePath, path).replace(/\\/g, '/');
		if (!path || isAbsolute(path) || path === '..' || path.startsWith('../')) { return 'Use a file path inside the attached folder.'; }
		args.path = folder.path === '.' ? path : folder.path + '/' + path;
	} else if (folder.path !== '.' && !isEqualOrParent(path, folder.path, !isLinux, true)) {
		args.path = folder.path + '/' + path;
	}
	if (folder.workspaceFolder) { args.workspaceFolder = folder.workspaceFolder; }
	return undefined;
}

/** Translate folder chips into the same relative paths and root names the editor tools accept. */
export function describeChatFolders(attachments: readonly MessageAttachment[] | undefined, roots: readonly URI[]) {
	const folders = attachments?.filter(isFolderAttachment) ?? [];
	const targets = new Map<string, AttachedChatFolder>();
	for (const folder of folders) {
		const [local] = getAttachedLocalPaths([folder]);
		const uri = local && resources.removeTrailingPathSeparator(local);
		const root = uri && !uri.authority && roots.filter(root => root.scheme === 'file' && resources.isEqualOrParent(uri, root))
			.sort((a, b) => b.path.length - a.path.length)[0];
		if (!root) {
			throw new Error(localize('reach.folder.selectWorkspace', "Choose “{0}” as a local workspace folder for this chat to use the attached folder.", folder.label));
		}
		if (roots.length > 1 && roots.filter(other => basename(other) === basename(root)).length > 1) {
			throw new Error(localize('reach.folder.ambiguousWorkspace', "Multiple workspace folders are named “{0}”. Select the intended workspace folder on its own to use this attachment.", basename(root)));
		}
		targets.set(resources.getComparisonKey(uri), { path: resources.relativePath(root, uri) || '.', ...(roots.length > 1 && { workspaceFolder: basename(root) }), absolutePath: uri.fsPath });
	}
	const paths = [...targets.values()];
	const description = paths.length ? 'User-attached folders (workspace-relative paths): ' + JSON.stringify(paths)
		+ '. Use path as the workspace-relative prefix for files and workspaceFolder when supplied. A dot denotes the workspace root. For a new project, create the requested files there. If you need a directory listing, list the attached absolutePath once. An empty listing means you can start there. If more than one destination matches the request, ask which to use.' : '';
	return { folders, paths, description };
}
