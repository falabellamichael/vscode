/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import assert from 'assert';
import { CancellationToken } from '../../../../../base/common/cancellation.js';
import { isCancellationError } from '../../../../../base/common/errors.js';
import { Disposable } from '../../../../../base/common/lifecycle.js';
import { URI } from '../../../../../base/common/uri.js';
import { mock } from '../../../../../base/test/common/mock.js';
import { ensureNoDisposablesAreLeakedInTestSuite } from '../../../../../base/test/common/utils.js';
import { createTextModel } from '../../../../../editor/test/common/testTextModel.js';
import { INativeHostService, type ReachWorkspaceFileRequest } from '../../../../../platform/native/common/native.js';
import { IWorkspaceTrustManagementService } from '../../../../../platform/workspace/common/workspaceTrust.js';
import { IResolvedTextFileEditorModel, ITextFileEditorModelManager, ITextFileService } from '../../../../services/textfile/common/textfiles.js';
import { ILanguageModelToolsService, ToolSet, type IToolData, type IToolImpl } from '../../common/tools/languageModelToolsService.js';
import { ReachWorkspaceFilesContribution } from '../../electron-browser/builtInTools/reachWorkspaceFiles.contribution.js';

suite('REACH workbench file tools', () => {
	const store = ensureNoDisposablesAreLeakedInTestSuite();
	const editor = URI.file('/editor'), agent = URI.file('/agent-project');
	const session = URI.parse('agent-host-reach:/session');
	function fixture(textFiles = new class extends mock<ITextFileService>() { }(), content = 'ok') {
		const registered = new Map<string, IToolImpl>();
		const calls: URI[] = [];
		let trusted = true;
		store.add(new ReachWorkspaceFilesContribution(new class extends mock<ILanguageModelToolsService>() {
			override createToolSet() {
				return new class extends mock<ToolSet & { dispose(): void }>() {
					override dispose() { }
					override addTool() { return Disposable.None; }
				};
			}
			override registerTool(data: IToolData, impl: IToolImpl) { registered.set(data.id, impl); return Disposable.None; }
		}, new class extends mock<INativeHostService>() {
			override async accessReachWorkspaceFile(root: URI, request: ReachWorkspaceFileRequest) {
				calls.push(root);
				return { uri: URI.joinPath(root, request.path).toString(), bytes: content.length, content };
			}
		}, new class extends mock<IWorkspaceTrustManagementService>() {
			override async getUriTrustInfo(uri: URI) { assert.strictEqual(uri, agent); return { uri, trusted }; }
		}, textFiles));
		return { registered, calls, revokeTrust: () => { trusted = false; } };
	}

	test('exact replacement previews both passages, preserves other text and refuses stale, ambiguous and dirty edits', async () => {
		const original = 'heading\r\nbody\r\nfooter\r\nfooter';
		const buffer = store.add(createTextModel(original));
		let dirty = false, saves = 0;
		const model = new class extends mock<IResolvedTextFileEditorModel>() {
			override readonly textEditorModel = buffer;
			override isResolved(): this is IResolvedTextFileEditorModel { return true; }
			override isReadonly() { return false; }
			override isDirty(): this is IResolvedTextFileEditorModel { return dirty; }
			override async save() { saves++; return true; }
		}();
		const { registered } = fixture(new class extends mock<ITextFileService>() {
			override isDirty() { return dirty; }
			override readonly files = new class extends mock<ITextFileEditorModelManager>() { override async resolve() { return model; } }();
		}(), original);
		const edit = registered.get('reach_edit_file')!;
		const [part] = (await registered.get('reach_read_file')!.invoke({ callId: 'read', toolId: 'reach_read_file', parameters: { path: 'index.html' }, context: { sessionResource: session, workingDirectories: [agent] } }, async () => 0, { report() { } }, CancellationToken.None)).content;
		assert.ok(part.kind === 'text');
		const revision = JSON.parse(part.value).fileRevision;
		assert.match(revision, /^[0-9a-f]{40}$/);
		const replaceFile = (fileRevision: string) => edit.invoke({ callId: 'edit', toolId: 'reach_edit_file', parameters: { path: 'index.html', fileRevision, newText: 'revised' }, context: { sessionResource: session, workingDirectories: [agent] } }, async () => 0, { report() { } }, CancellationToken.None);
		await assert.rejects(replaceFile('stale'), /revision does not match/);
		const fullPreview = await edit.prepareToolInvocation!({ toolCallId: 'edit', chatSessionResource: session, workingDirectories: [agent], parameters: { path: 'index.html', fileRevision: revision, newText: 'revised' } }, CancellationToken.None);
		assert.match(JSON.stringify(fullPreview?.confirmationMessages?.message), /Before.*heading.*footer.*After.*revised/);
		const parameters = { path: 'index.html', oldText: 'heading\nbody', newText: 'heading\nnew body' };
		const prepared = await edit.prepareToolInvocation!({ toolCallId: 'edit', chatSessionResource: session, workingDirectories: [agent], parameters }, CancellationToken.None);
		assert.strictEqual(prepared?.confirmationMessages?.title, 'Edit workspace file?');
		assert.match(JSON.stringify(prepared?.confirmationMessages?.message), /Before.*heading.*After.*new body/);
		const invoke = (oldText: string) => edit.invoke({ callId: 'edit', toolId: 'reach_edit_file', parameters: { ...parameters, oldText }, context: { sessionResource: session, workingDirectories: [agent] } }, async () => 0, { report() { } }, CancellationToken.None);
		await assert.rejects(invoke('footer'), /appears 2 times in the file \(lines 3, 4\)/);
		dirty = true;
		await assert.rejects(invoke(parameters.oldText), /saved and writable/);
		dirty = false;
		await assert.rejects(invoke('missing passage'), /does not exist in the current file/);
		await invoke(parameters.oldText);
		await assert.rejects(invoke(parameters.oldText), /matches exactly once/);
		assert.deepStrictEqual([saves, buffer.getValue()], [1, 'heading\r\nnew body\r\nfooter\r\nfooter']);
		await assert.rejects(replaceFile(revision), /matches exactly once/);
		buffer.setValue(original);
		const rangeRead = await registered.get('reach_read_file')!.invoke({ callId: 'read', toolId: 'reach_read_file', parameters: { path: 'index.html', startLine: 4, endLine: 4 }, context: { sessionResource: session, workingDirectories: [agent] } }, async () => 0, { report() { } }, CancellationToken.None);
		const rangePart = rangeRead.content[0];
		assert.ok(rangePart.kind === 'text');
		const rangeRevision = JSON.parse(rangePart.value).fileRevision;
		assert.match(rangeRevision, /^4:4:[0-9a-f]{40}$/);
		await assert.rejects(replaceFile(rangeRevision.replace('4:4:', '1:1:')), /revision does not match/);
		await replaceFile(rangeRevision);
		assert.strictEqual(buffer.getValue(), 'heading\r\nbody\r\nfooter\r\nrevised');
		buffer.setValue(original);
		await replaceFile(revision);
		assert.deepStrictEqual([saves, buffer.getValue()], [3, 'revised']);
	});

	test('nearLine disambiguates a repeated passage and reports the applied line', async () => {
		const original = 'heading\r\nbody\r\nfooter\r\nfooter';
		const buffer = store.add(createTextModel(original));
		const model = new class extends mock<IResolvedTextFileEditorModel>() {
			override readonly textEditorModel = buffer;
			override isResolved(): this is IResolvedTextFileEditorModel { return true; }
			override isReadonly() { return false; }
			override isDirty(): this is IResolvedTextFileEditorModel { return false; }
			override async save() { return true; }
		}();
		const { registered } = fixture(new class extends mock<ITextFileService>() {
			override isDirty() { return false; }
			override readonly files = new class extends mock<ITextFileEditorModelManager>() { override async resolve() { return model; } }();
		}(), original);
		const edit = registered.get('reach_edit_file')!;
		const invokeNear = (nearLine: number) => edit.invoke({ callId: 'edit', toolId: 'reach_edit_file', parameters: { path: 'index.html', oldText: 'footer', newText: 'end', nearLine }, context: { sessionResource: session, workingDirectories: [agent] } }, async () => 0, { report() { } }, CancellationToken.None);
		const result = await invokeNear(4);
		const part = result.content[0];
		assert.ok(part.kind === 'text');
		assert.strictEqual(JSON.parse(part.value).appliedAtLine, 4);
		assert.strictEqual(buffer.getValue(), 'heading\r\nbody\r\nfooter\r\nend');
		buffer.setValue(original);
		await invokeNear(3);
		assert.strictEqual(buffer.getValue(), 'heading\r\nbody\r\nend\r\nfooter');
		buffer.setValue(original);
		await invokeNear(100);
		assert.strictEqual(buffer.getValue(), 'heading\r\nbody\r\nfooter\r\nend');
	});

	test('a whitespace-drifted passage locates its unique tolerant match and chains a fresh revision', async () => {
		const original = 'heading\r\n\r\nbody text\r\nfooter';
		const buffer = store.add(createTextModel(original));
		const model = new class extends mock<IResolvedTextFileEditorModel>() {
			override readonly textEditorModel = buffer;
			override isResolved(): this is IResolvedTextFileEditorModel { return true; }
			override isReadonly() { return false; }
			override isDirty(): this is IResolvedTextFileEditorModel { return false; }
			override async save() { return true; }
		}();
		const { registered } = fixture(new class extends mock<ITextFileService>() {
			override isDirty() { return false; }
			override readonly files = new class extends mock<ITextFileEditorModelManager>() { override async resolve() { return model; } }();
		}(), original);
		const edit = registered.get('reach_edit_file')!;
		// The model copied the passage with single spaces where the file has a blank line.
		const result = await edit.invoke({ callId: 'edit', toolId: 'reach_edit_file', parameters: { path: 'index.html', oldText: 'heading body text', newText: 'replaced' }, context: { sessionResource: session, workingDirectories: [agent] } }, async () => 0, { report() { } }, CancellationToken.None);
		assert.strictEqual(buffer.getValue(), 'replaced\r\nfooter');
		const [part] = result.content;
		assert.ok(part.kind === 'text');
		const payload = JSON.parse(part.value);
		assert.match(payload.message, /whitespace normalization/);
		assert.match(payload.fileRevision, /^[0-9a-f]{40}$/, 'success results carry the fresh revision for chained edits');
		assert.strictEqual(typeof payload.totalLines, 'number');
		assert.match(payload.revisionNote, /invalid/);
	});

	test('creates only under the selected chat root and returns a link after native creation', async () => {
		const { registered, calls } = fixture();
		const create = registered.get('reach_create_file')!;
		const parameters = { path: 'report.csv', content: 'ok', workspaceFolder: 'agent-project' };
		const workingDirectories = [editor, agent];
		const preparation = await create.prepareToolInvocation!({ toolCallId: 'tool', chatSessionResource: session, parameters, workingDirectories }, CancellationToken.None);
		assert.strictEqual(preparation?.confirmationMessages?.title, 'Create workspace file?');
		assert.strictEqual(calls.length, 0);
		const result = await create.invoke({ callId: 'tool', toolId: 'reach_create_file', parameters, context: { sessionResource: session, workingDirectories } }, async () => 0, { report() { } }, CancellationToken.None);
		assert.deepStrictEqual(calls, [agent]);
		assert(Array.isArray(result.toolResultDetails));
		assert.deepStrictEqual(result.toolResultDetails.map(resource => resource.toString()), [URI.joinPath(agent, 'report.csv').toString()]);
		assert(result.toolResultMessage && typeof result.toolResultMessage !== 'string' && result.toolResultMessage.value.includes('report.csv]('));
	});

	test('provides and verifies full-file revisions up to 64 KiB for an attached rebuild', async () => {
		for (const size of [64 * 1024, 64 * 1024 + 1]) {
			const { registered } = fixture(undefined, 'x'.repeat(size));
			const context = { sessionResource: session, workingDirectories: [agent] };
			const read = registered.get('reach_read_file')!;
			const [part] = (await read.invoke({ callId: 'read', toolId: 'reach_read_file', parameters: { path: 'game.html' }, context }, async () => 0, { report() { } }, CancellationToken.None)).content;
			assert.ok(part.kind === 'text');
			const { fileRevision } = JSON.parse(part.value);
			if (size > 64 * 1024) {
				assert.strictEqual(fileRevision, undefined);
			} else {
				assert.match(fileRevision, /^[0-9a-f]{40}$/);
				const edit = registered.get('reach_edit_file')!;
				const prepare = (revision: string) => edit.prepareToolInvocation!({ toolCallId: 'edit', chatSessionResource: session, workingDirectories: [agent], parameters: { path: 'game.html', fileRevision: revision, newText: '<main>Sudoku</main>' } }, CancellationToken.None);
				assert.ok((await prepare(fileRevision))?.confirmationMessages);
				await assert.rejects(prepare('0'.repeat(40)), /revision does not match/);
			}
		}
	});

	test('reads an inclusive range from a larger text file without changing the full-read result', async () => {
		const { registered } = fixture(undefined, 'alpha\r\nbravo\r\ncharlie');
		const read = registered.get('reach_read_file')!;
		const invoke = (range: Record<string, unknown>) => read.invoke({
			callId: 'read', toolId: 'reach_read_file', parameters: { path: 'index.html', ...range },
			context: { sessionResource: session, workingDirectories: [agent] },
		}, async () => 0, { report() { } }, CancellationToken.None);
		for (const [range, expected] of [[{}, 'alpha\r\nbravo\r\ncharlie'], [{ startLine: 2, endLine: 3 }, 'bravo\ncharlie']] as const) {
			const [part] = (await invoke(range)).content;
			assert.ok(part.kind === 'text');
			assert.strictEqual(JSON.parse(part.value).content, expected);
		}
		await assert.rejects(invoke({ startLine: 4 }), /within the file/);
	});

	test('refuses missing, ambiguous, untrusted and cancelled calls before dispatch', async () => {
		const { registered, calls, revokeTrust } = fixture();
		const read = registered.get('reach_read_file')!;
		const invoke = (roots: URI[], token: CancellationToken = CancellationToken.None) => read.invoke({
			callId: 'tool', toolId: 'reach_read_file', parameters: { path: 'project.txt' },
			context: { sessionResource: session, workingDirectory: editor, workingDirectories: roots },
		}, async () => 0, { report() { } }, token);
		await assert.rejects(invoke([]), /Choose a local workspace/);
		await assert.rejects(invoke([editor, agent]), /Choose one workspace/);
		await assert.rejects(invoke([agent], CancellationToken.Cancelled), isCancellationError);
		revokeTrust();
		await assert.rejects(invoke([agent]), /Trust the chat/);
		assert.strictEqual(calls.length, 0);
	});
});
