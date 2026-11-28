/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import assert from 'assert';
import { mkdir, mkdtemp, readFile, readdir, rm, symlink, writeFile } from 'fs/promises';
import { tmpdir } from 'os';
import { join } from '../../../../base/common/path.js';
import { URI } from '../../../../base/common/uri.js';
import { ensureNoDisposablesAreLeakedInTestSuite } from '../../../../base/test/common/utils.js';
import { accessReachWorkspaceFile } from '../../node/reachWorkspaceFile.js';

suite('REACH workspace file tools', () => {
	ensureNoDisposablesAreLeakedInTestSuite();
	let temporary: string, root: URI;
	setup(async () => {
		temporary = await mkdtemp(join(tmpdir(), 'reach-file-tools-'));
		root = URI.file(join(temporary, 'workspace'));
		await mkdir(root.fsPath);
	});
	teardown(async () => { await rm(temporary, { recursive: true, force: true, maxRetries: 3 }); });

	test('creates code and binary files, returns their real URI and reads the exact bytes', async () => {
		const content = '\ufeffimport { π } from "./math";\r\n';
		const created = await accessReachWorkspaceFile(root, { operation: 'create', path: 'source/code.ts', content });
		assert.strictEqual(created.uri, URI.joinPath(root, 'source', 'code.ts').toString());
		assert.strictEqual(await readFile(URI.parse(created.uri).fsPath, 'utf8'), content);
		assert.strictEqual((await accessReachWorkspaceFile(root, { operation: 'read', path: 'source/code.ts' })).content, content);
		const bytes = Buffer.from([0x00, 0xff, 0x80, 0x0a]);
		await accessReachWorkspaceFile(root, { operation: 'create', path: 'asset.bin', content: bytes.toString('base64'), encoding: 'base64' });
		assert.deepStrictEqual(await readFile(join(root.fsPath, 'asset.bin')), bytes);
		assert.strictEqual((await accessReachWorkspaceFile(root, { operation: 'read', path: 'asset.bin', encoding: 'base64' })).content, bytes.toString('base64'));
		await assert.rejects(accessReachWorkspaceFile(root, { operation: 'read', path: 'asset.bin' }), /base64/);
		await assert.rejects(accessReachWorkspaceFile(root, { operation: 'create', path: 'bad.bin', content: 'bad data', encoding: 'base64' }), /base64/);
	});

	test('preserves an existing destination even when several callers publish concurrently', async () => {
		const values = ['first', 'second', 'third'];
		const concurrent = await Promise.allSettled(values.map(content => accessReachWorkspaceFile(root, { operation: 'create', path: 'same.txt', content })));
		const winner = concurrent.findIndex(result => result.status === 'fulfilled');
		assert(winner >= 0);
		assert.strictEqual(concurrent.filter(result => result.status === 'fulfilled').length, 1);
		await assert.rejects(accessReachWorkspaceFile(root, { operation: 'create', path: 'same.txt', content: 'replacement' }), /already exists/);
		assert.strictEqual(await readFile(join(root.fsPath, 'same.txt'), 'utf8'), values[winner]);
		assert.deepStrictEqual(await readdir(root.fsPath), ['same.txt']);
	});

	test('rejects absolute paths, traversal, device names, a symlink root and a subfolder junction', async () => {
		const outside = join(temporary, 'outside');
		await mkdir(outside);
		await writeFile(join(outside, 'private.txt'), 'outside data');
		await symlink(outside, join(root.fsPath, 'linked'), 'junction');
		for (const path of ['../outside/private.txt', 'linked/private.txt', '/outside.txt', 'C:\\outside.txt', '\\\\host\\share\\file', 'file.md:ads', 'NUL.txt']) {
			await assert.rejects(accessReachWorkspaceFile(root, { operation: 'read', path }));
			await assert.rejects(accessReachWorkspaceFile(root, { operation: 'create', path, content: 'changed' }));
		}
		await assert.rejects(accessReachWorkspaceFile(URI.file(join(root.fsPath, 'linked')), { operation: 'create', path: 'new.txt', content: 'changed' }));
		assert.deepStrictEqual(await readdir(outside), ['private.txt']);
		assert.strictEqual(await readFile(join(outside, 'private.txt'), 'utf8'), 'outside data');
	});
});
