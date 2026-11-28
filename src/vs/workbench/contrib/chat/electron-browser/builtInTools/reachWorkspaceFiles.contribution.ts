/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { CancellationToken } from '../../../../../base/common/cancellation.js';
import { CancellationError } from '../../../../../base/common/errors.js';
import { hashAsync } from '../../../../../base/common/hash.js';
import { escapeMarkdownLinkLabel, MarkdownString } from '../../../../../base/common/htmlContent.js';
import { Disposable } from '../../../../../base/common/lifecycle.js';
import { basename } from '../../../../../base/common/resources.js';
import { URI } from '../../../../../base/common/uri.js';
import { Range } from '../../../../../editor/common/core/range.js';
import { localize } from '../../../../../nls.js';
import { INativeHostService, type ReachWorkspaceFileRequest } from '../../../../../platform/native/common/native.js';
import { IWorkspaceTrustManagementService } from '../../../../../platform/workspace/common/workspaceTrust.js';
import { registerWorkbenchContribution2, WorkbenchPhase } from '../../../../common/contributions.js';
import { ITextFileService } from '../../../../services/textfile/common/textfiles.js';
import { ILanguageModelToolsService, ToolDataSource, type IToolData, type IToolInvocationPreparationContext } from '../../common/tools/languageModelToolsService.js';

const fullRevisionByteLimit = 256 * 1024;
// Ranged revisions validate by re-reading and re-hashing the section, which is cheap even
// for large sections — the cap only needs to exclude absurd spans, not useful editing ranges.
const sectionRevisionCharLimit = 256 * 1024;

function fileTool(operation: 'create' | 'read'): IToolData {
	const create = operation === 'create';
	const displayName = create ? localize('reach.tools.create', "Create workspace file") : localize('reach.tools.read', "Read workspace file");
	return {
		id: create ? 'reach_create_file' : 'reach_read_file',
		toolReferenceName: create ? 'reachCreateFile' : 'reachReadFile',
		source: ToolDataSource.Internal, displayName, runsInWorkspace: true, canBeReferencedInPrompt: true, canRequestPreApproval: create,
		modelDescription: create
			? 'Create a new file and missing parent folders in the chat workspace. Supply exact UTF-8 text for code and documents, or padded base64 for binary files. Existing files are preserved. Returns the file URI.'
			: 'Read a workspace file up to 1 MiB as UTF-8 or padded base64. For long text files, read sections using inclusive, one-based startLine and endLine.',
		inputSchema: {
			type: 'object', additionalProperties: false,
			properties: {
				path: { type: 'string', description: 'File path relative to the workspace folder.' },
				workspaceFolder: { type: 'string', description: 'Name of the workspace folder. Required only when the chat has multiple folders.' },
				encoding: { type: 'string', enum: ['utf8', 'base64'], default: 'utf8' },
				...(create ? { content: { type: 'string', description: 'Complete contents of the new file.' } } : {
					startLine: { type: 'integer', minimum: 1 }, endLine: { type: 'integer', minimum: 1 },
				}),
			},
			required: create ? ['path', 'content'] : ['path'],
		},
	};
}

type WorkspaceContext = Pick<IToolInvocationPreparationContext, 'workingDirectory' | 'workingDirectories'>;

function targetRoot(context: WorkspaceContext | undefined, folderName: unknown): URI {
	const folders = context?.workingDirectories ?? (context?.workingDirectory ? [context.workingDirectory] : []);
	const candidates = folderName === undefined ? folders : folders.filter(folder => folderName === basename(folder));
	if (!candidates.length && !folders.length) { throw new Error(localize('reach.tools.noWorkspace', "Choose a local workspace folder for this chat.")); }
	if (candidates.length !== 1) { throw new Error(localize('reach.tools.chooseWorkspace', "Choose one workspace folder by name from {0}.", folders.map(folder => basename(folder)).join(', '))); }
	if (candidates[0].scheme !== 'file' || candidates[0].authority) { throw new Error(localize('reach.tools.localFile', "Choose a local workspace folder for this file operation.")); }
	return candidates[0];
}

export class ReachWorkspaceFilesContribution extends Disposable {
	static readonly ID = 'chat.reachWorkspaceFiles';

	constructor(
		@ILanguageModelToolsService tools: ILanguageModelToolsService,
		@INativeHostService nativeHost: INativeHostService,
		@IWorkspaceTrustManagementService trust: IWorkspaceTrustManagementService,
		@ITextFileService textFiles: ITextFileService,
	) {
		super();
		const set = this._register(tools.createToolSet(ToolDataSource.Internal, 'reach-workspace-files', 'reachFiles', { description: localize('reach.tools.files', "Workspace files") }));
		for (const operation of ['create', 'read'] as const) {
			const data = fileTool(operation);
			this._register(tools.registerTool(data, {
				prepareToolInvocation: async context => {
					const root = targetRoot(context, context.parameters.workspaceFolder);
					const path = String(context.parameters.path);
					const invocationMessage = operation === 'create' ? localize('reach.tools.creating', "Creating {0}", path) : localize('reach.tools.reading', "Reading {0}", path);
					if (operation === 'read') { return { invocationMessage }; }
					const message = new MarkdownString().appendText(URI.joinPath(root, path).fsPath);
					if (context.parameters.encoding !== 'base64' && typeof context.parameters.content === 'string') {
						message.appendCodeblock('', context.parameters.content.slice(0, 4096));
					}
					return { invocationMessage, confirmationMessages: { title: localize('reach.tools.confirmCreate', "Create workspace file?"), message } };
				},
				invoke: async (invocation, _countTokens, _progress, token: CancellationToken) => {
					const root = targetRoot(invocation.context, invocation.parameters.workspaceFolder);
					if (!(await trust.getUriTrustInfo(root)).trusted) { throw new Error(localize('reach.tools.trust', "Trust the chat's workspace folder to use this tool.")); }
					if (token.isCancellationRequested) { throw new CancellationError(); }
					const { path, encoding } = invocation.parameters;
					const request: ReachWorkspaceFileRequest = operation === 'create'
						? { operation, path, encoding, content: invocation.parameters.content }
						: { operation, path, encoding };
					const result = await nativeHost.accessReachWorkspaceFile(root, request);
					const uri = URI.parse(result.uri);
					const link = new MarkdownString().appendLink(uri, escapeMarkdownLinkLabel(basename(uri))).value;
					const toolResultMessage = new MarkdownString(operation === 'create'
						? localize('reach.tools.created', "Created {0}", link)
						: localize('reach.tools.didRead', "Read {0}", link));
					let output: typeof result & { startLine?: number; endLine?: number; totalLines?: number; fileRevision?: string } = result;
					const { startLine, endLine } = invocation.parameters;
					if (operation === 'read' && (startLine !== undefined || endLine !== undefined)) {
						const lines = result.content?.split(/\r\n|\n|\r/) ?? [];
						const start = startLine ?? 1, end = endLine ?? lines.length;
						if (encoding === 'base64' || !Number.isSafeInteger(start) || !Number.isSafeInteger(end) || start < 1 || start > lines.length || end < start) {
							throw new Error(localize('reach.tools.lineRange', "Choose a valid UTF-8 text range within the file's {0} lines.", lines.length));
						}
						const last = Math.min(end, lines.length);
						const content = lines.slice(start - 1, last).join('\n');
						output = { ...result, content, startLine: start, endLine: last, totalLines: lines.length };
						if (content.length <= sectionRevisionCharLimit && !content.includes('\0')) {
							output.fileRevision = start + ':' + last + ':' + await hashAsync(content);
						}
					} else if (operation === 'read' && encoding !== 'base64' && result.content !== undefined && result.bytes <= fullRevisionByteLimit && !result.content.includes('\0')) {
						output = { ...result, fileRevision: await hashAsync(result.content) };
					}
					return { content: [{ kind: 'text', value: JSON.stringify(output) }], toolResultMessage, toolResultDetails: [uri] };
				},
			}));
			this._register(set.addTool(data));
		}
		const edit: IToolData = {
			id: 'reach_edit_file', toolReferenceName: 'reachEditFile', source: ToolDataSource.Internal,
			displayName: localize('reach.tools.edit', "Edit workspace file"), runsInWorkspace: true, canBeReferencedInPrompt: true, canRequestPreApproval: true,
			modelDescription: 'Edit an existing workspace text file. For a targeted edit, pass a short unique oldText from the source and its replacement in newText; if the passage occurs several times, set nearLine to the line of the occurrence you mean. Alternatively use fileRevision from a read to replace exactly the lines returned by that read with newText. A full-file revision targets the whole file. Use only one mode and preserve unrelated content.',
			inputSchema: {
				type: 'object', additionalProperties: false, required: ['path', 'newText'],
				properties: {
					path: { type: 'string' }, workspaceFolder: { type: 'string', description: 'Folder name for multi-root workspaces.' },
					oldText: { type: 'string', description: 'Nonempty passage occurring exactly once. If it occurs several times, set nearLine to pick the occurrence. Omit fileRevision when supplying oldText.' }, newText: { type: 'string' },
					fileRevision: { type: 'string', description: 'Revision from a read. For a line range, newText replaces only those lines. For a full-file revision, supply the complete revised file. Omit oldText.' },
					nearLine: { type: 'number', description: '1-based line number. When oldText occurs several times, the occurrence nearest this line is edited.' },
				},
			},
		};
		const readRevision = async (root: URI, path: string, revision: string) => {
			if (!(await trust.getUriTrustInfo(root)).trusted) { throw new Error(localize('reach.tools.trust', "Trust the chat's workspace folder to use this tool.")); }
			const file = await nativeHost.accessReachWorkspaceFile(root, { operation: 'read', path });
			const section = /^(\d+):(\d+):([a-f0-9]{40})$/.exec(revision);
			if (section && file.content !== undefined && !file.content.includes('\0')) {
				const start = Number(section[1]), end = Number(section[2]);
				const lines = file.content.split(/\r\n|\n|\r/);
				if (start >= 1 && end >= start && end <= lines.length) {
					const before = lines.slice(start - 1, end).join('\n');
					if (before.length <= sectionRevisionCharLimit && await hashAsync(before) === section[3]) {
						const range = new Range(start, 1, end, (end === 1 ? lines[0].replace(/^\uFEFF/, '') : lines[end - 1]).length + 1);
						return { ...file, before, range };
					}
				}
			}
			if (file.content === undefined || file.content.includes('\0') || file.bytes > fullRevisionByteLimit || await hashAsync(file.content) !== revision) {
				throw new Error(localize('reach.tools.staleRevision', "The revision does not match this file — the file changed since that read (a successful edit invalidates every earlier revision). Read the current file or target section again, or use the fileRevision returned by the latest successful edit result."));
			}
			return { ...file, before: file.content, range: undefined };
		};
		const validateEdit = (oldText: unknown, newText: unknown, fileRevision: unknown) => {
			if (typeof newText !== 'string' || (fileRevision === undefined ? typeof oldText !== 'string' || !oldText : typeof fileRevision !== 'string' || oldText !== undefined)) {
				throw new Error(localize('reach.tools.editInput', "Supply either an original passage or a file revision, plus the replacement text."));
			}
		};
		this._register(tools.registerTool(edit, {
			prepareToolInvocation: async context => {
				const { path, workspaceFolder, oldText, newText, fileRevision } = context.parameters;
				validateEdit(oldText, newText, fileRevision);
				const root = targetRoot(context, workspaceFolder);
				const before = fileRevision !== undefined ? (await readRevision(root, path, fileRevision)).before : oldText;
				const message = new MarkdownString().appendText(URI.joinPath(root, path).fsPath)
					.appendText('\n\n' + localize('reach.tools.before', "Before")).appendCodeblock('', before)
					.appendText('\n\n' + localize('reach.tools.after', "After")).appendCodeblock('', newText);
				return { invocationMessage: localize('reach.tools.editing', "Editing {0}", path), confirmationMessages: { title: localize('reach.tools.confirmEdit', "Edit workspace file?"), message } };
			},
			invoke: async (invocation, _count, _progress, token) => {
				const { path, workspaceFolder, oldText, newText, fileRevision, nearLine } = invocation.parameters;
				const root = targetRoot(invocation.context, workspaceFolder);
				if (!(await trust.getUriTrustInfo(root)).trusted) { throw new Error(localize('reach.tools.trust', "Trust the chat's workspace folder to use this tool.")); }
				validateEdit(oldText, newText, fileRevision);
				// The native reader checks containment, symlinks, size and UTF-8 before opening the editor model.
				const revision = fileRevision === undefined ? undefined : await readRevision(root, path, fileRevision);
				const file = revision ?? await nativeHost.accessReachWorkspaceFile(root, { operation: 'read', path });
				const uri = URI.parse(file.uri);
				const model = await textFiles.files.resolve(uri);
				if (token.isCancellationRequested) { throw new CancellationError(); }
				if (!model.isResolved() || textFiles.isDirty(uri) || model.isReadonly()) { throw new Error(localize('reach.tools.editUnavailable', "The file must be saved and writable before applying this edit.")); }
				const buffer = model.textEditorModel;
				const normalize = (value: string) => value.replace(/\r\n|\r|\n/g, buffer.getEOL());
				const original = buffer.getValue();
				const needle = fileRevision === undefined ? normalize(oldText) : original;
				const index = original.indexOf(needle);
				if (file.content === undefined || file.content.includes('\0') || normalize(file.content.replace(/^\uFEFF/, '')) !== original) {
					throw new Error(localize('reach.tools.editMatch', "Read the file again and supply a passage that matches exactly once. Copy a shorter unique line or tag from the latest read, including its whitespace."));
				}
				// Small models often copy passages with drifting whitespace. When the exact
				// passage is absent, locate it tolerantly: one unique whitespace-flexible match
				// is almost certainly the intended target, and the replacement text is what matters.
				let locatedIndex = index;
				let locatedLength = needle.length;
				let tolerant = false;
				let disambiguatedLine: number | undefined;
				const nearLineNumber = typeof nearLine === 'number' ? nearLine : typeof nearLine === 'string' && /^\d+$/.test(nearLine.trim()) ? Number(nearLine.trim()) : undefined;
				const near = fileRevision === undefined && nearLineNumber !== undefined && Number.isSafeInteger(nearLineNumber) && nearLineNumber >= 1 ? nearLineNumber : undefined;
				const nearest = (candidates: { index: number; length: number }[]) => {
					let best = candidates[0], bestDistance = Infinity;
					for (const candidate of candidates) {
						const distance = Math.abs(buffer.getPositionAt(candidate.index).lineNumber - (near ?? 1));
						if (distance < bestDistance) { bestDistance = distance; best = candidate; }
					}
					return best;
				};
				if (fileRevision === undefined && index < 0 && needle.length <= 4000 && needle.trim()) {
					const escaped = needle.trim().split(/\s+/).map(part => part.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('\\s+');
					if (escaped) {
						const pattern = new RegExp(escaped, 'g');
						const tolerantMatches: { index: number; length: number }[] = [];
						for (let match = pattern.exec(original); match && tolerantMatches.length < 12; match = pattern.exec(original)) {
							tolerantMatches.push({ index: match.index, length: match[0].length });
						}
						const chosen = tolerantMatches.length === 1 ? tolerantMatches[0] : tolerantMatches.length > 1 && near !== undefined ? nearest(tolerantMatches) : undefined;
						if (chosen) {
							locatedIndex = chosen.index;
							locatedLength = chosen.length;
							tolerant = true;
							if (tolerantMatches.length > 1) { disambiguatedLine = buffer.getPositionAt(chosen.index).lineNumber; }
						} else if (tolerantMatches.length > 1) {
							const positions = tolerantMatches.slice(0, 6).map(match => buffer.getPositionAt(match.index).lineNumber);
							throw new Error(localize('reach.tools.editAmbiguousTolerant', "That passage does not match exactly, but after whitespace normalization it appears about {0} times in the file (lines {1}). The edit was not applied. Re-issue the same edit with nearLine set to the line you mean — for example nearLine: {2} picks the occurrence at line {2}.", tolerantMatches.length, positions.join(', '), positions[0]));
						}
					}
				}
				if (fileRevision === undefined && index < 0 && !tolerant) {
					throw new Error(localize('reach.tools.editNotFound', "That passage does not exist in the current file. Read the file again and copy oldText verbatim from the latest read, including its exact whitespace. If the file is long, read the section you want to change with startLine and endLine first."));
				}
				if (fileRevision === undefined && index >= 0 && index !== original.lastIndexOf(needle)) {
					if (near !== undefined) {
						const exactMatches: { index: number; length: number }[] = [];
						for (let at = original.indexOf(needle); at >= 0 && exactMatches.length < 50; at = original.indexOf(needle, at + 1)) {
							exactMatches.push({ index: at, length: needle.length });
						}
						const chosen = nearest(exactMatches);
						locatedIndex = chosen.index;
						locatedLength = needle.length;
						disambiguatedLine = buffer.getPositionAt(chosen.index).lineNumber;
					} else {
						const positions: number[] = [];
						for (let at = original.indexOf(needle); at >= 0 && positions.length < 6; at = original.indexOf(needle, at + 1)) {
							positions.push(buffer.getPositionAt(at).lineNumber);
						}
						const total = original.split(needle).length - 1;
						throw new Error(localize('reach.tools.editAmbiguous', "That passage appears {0} times in the file (lines {1}). The edit was not applied. Re-issue the same edit with nearLine set to the line you mean — for example nearLine: {2} picks the occurrence at line {2}. Alternatively widen oldText with unique surrounding lines so it matches exactly once.", total, positions.join(', '), positions[0]));
					}
				}
				if (!tolerant && disambiguatedLine === undefined && (index < 0 || index !== original.lastIndexOf(needle))) {
					throw new Error(localize('reach.tools.editMatch', "Read the file again and supply a passage that matches exactly once. Copy a shorter unique line or tag from the latest read, including its whitespace."));
				}
				const range = revision?.range ?? Range.fromPositions(buffer.getPositionAt(locatedIndex), buffer.getPositionAt(locatedIndex + locatedLength));
				buffer.pushStackElement();
				buffer.pushEditOperations(null, [{ range, text: newText }], () => null);
				buffer.pushStackElement();
				if (!(await model.save()) || model.isDirty()) { throw new Error(localize('reach.tools.editSave', "The edit is in the editor but could not be saved. Check the file before retrying.")); }
				// Report the post-edit revision so follow-up edits chain without a re-read.
				let freshRevision: string | undefined;
				let freshLines: number | undefined;
				try {
					const fresh = await nativeHost.accessReachWorkspaceFile(root, { operation: 'read', path });
					if (fresh.content !== undefined && !fresh.content.includes('\0')) {
						freshLines = fresh.content.split(/\r\n|\n|\r/).length;
						if (fresh.bytes <= fullRevisionByteLimit) { freshRevision = await hashAsync(fresh.content); }
					}
				} catch { /* Revision enrichment is best-effort. */ }
				const applied = localize('reach.tools.edited', "Saved changes to {0}.", path) + (tolerant ? ' ' + localize('reach.tools.editedTolerant', "The passage matched after whitespace normalization.") : '') + (disambiguatedLine !== undefined ? ' ' + localize('reach.tools.editedDisambiguated', "The occurrence nearest nearLine was applied at line {0}.", disambiguatedLine) : '');
				return { content: [{ kind: 'text', value: JSON.stringify({
					uri: uri.toString(), message: applied,
					...(disambiguatedLine !== undefined && { appliedAtLine: disambiguatedLine }),
					...(freshRevision && { fileRevision: freshRevision }),
					...(freshLines && { totalLines: freshLines }),
					revisionNote: 'This edit changed the file. Every earlier fileRevision for it is now invalid. For the next change use the fileRevision from this result or a fresh read.',
				}) }], toolResultDetails: [uri] };
			},
		}));
		this._register(set.addTool(edit));
	}
}

registerWorkbenchContribution2(ReachWorkspaceFilesContribution.ID, ReachWorkspaceFilesContribution, WorkbenchPhase.BlockRestore);
