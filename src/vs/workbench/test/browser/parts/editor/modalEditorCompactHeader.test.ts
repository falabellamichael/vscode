/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import assert from 'assert';
import { DisposableStore, toDisposable } from '../../../../../base/common/lifecycle.js';
import { ensureNoDisposablesAreLeakedInTestSuite } from '../../../../../base/test/common/utils.js';
import { canMoveModalEditorToEditorArea, IModalEditorOptions, IModalEditorOptionsProvider, isModalEditorOptionsProvider } from '../../../../../platform/editor/common/editor.js';
import '../../../../browser/parts/editor/media/modalEditorPart.css';

const MODAL_HEADER_HEIGHT = 33; // matches fallback in modalEditorPart.ts

/**
 * Mirrors the compact-header reconciliation in `ModalEditorPart.create`:
 *
 *   const editorModalOptions = isModalEditorOptionsProvider(activeEditor)
 *       ? activeEditor.getModalEditorOptions()
 *       : undefined;
 *   modalElement.classList.toggle('compact-header', !!editorModalOptions?.compactHeader);
 *
 * and the sash height resolution which prefers the measured header height
 * over the fallback constant.
 */
function applyCompactHeaderClass(modalElement: { classList: DOMTokenList }, activeEditor: unknown): void {
	const editorModalOptions = isModalEditorOptionsProvider(activeEditor) ? activeEditor.getModalEditorOptions() : undefined;
	modalElement.classList.toggle('compact-header', !!editorModalOptions?.compactHeader);
}

function resolveHeaderHeight(headerElement: { offsetHeight: number }): number {
	return headerElement.offsetHeight || MODAL_HEADER_HEIGHT;
}

class TestEditor implements IModalEditorOptionsProvider {
	constructor(private readonly options: IModalEditorOptions | undefined) { }
	getModalEditorOptions(): IModalEditorOptions | undefined { return this.options; }
}

suite('Modal Editor Compact Header', () => {

	const disposables = new DisposableStore();

	teardown(() => disposables.clear());

	ensureNoDisposablesAreLeakedInTestSuite();

	test('isModalEditorOptionsProvider type guard', () => {
		assert.strictEqual(isModalEditorOptionsProvider(undefined), false);
		assert.strictEqual(isModalEditorOptionsProvider(null), false);
		assert.strictEqual(isModalEditorOptionsProvider({}), false);
		assert.strictEqual(isModalEditorOptionsProvider({ getModalEditorOptions: 'nope' }), false);
		assert.strictEqual(isModalEditorOptionsProvider(new TestEditor({ compactHeader: true })), true);
	});

	test('compact-header class toggles based on active editor options', () => {
		const modalElement = document.createElement('div');

		// No active editor → class not set
		applyCompactHeaderClass(modalElement, undefined);
		assert.strictEqual(modalElement.classList.contains('compact-header'), false);

		// Editor that doesn't implement the provider → class not set
		applyCompactHeaderClass(modalElement, { /* not a provider */ });
		assert.strictEqual(modalElement.classList.contains('compact-header'), false);

		// Provider returning undefined → class not set
		applyCompactHeaderClass(modalElement, new TestEditor(undefined));
		assert.strictEqual(modalElement.classList.contains('compact-header'), false);

		// Provider opting out → class not set
		applyCompactHeaderClass(modalElement, new TestEditor({ compactHeader: false }));
		assert.strictEqual(modalElement.classList.contains('compact-header'), false);

		// Provider opting in → class set
		applyCompactHeaderClass(modalElement, new TestEditor({ compactHeader: true }));
		assert.strictEqual(modalElement.classList.contains('compact-header'), true);

		// Switching back to a non-compact editor clears the class
		applyCompactHeaderClass(modalElement, new TestEditor({ compactHeader: false }));
		assert.strictEqual(modalElement.classList.contains('compact-header'), false);
	});

	test('moving to an editor area defaults to enabled and supports opt out', () => {
		assert.deepStrictEqual([
			canMoveModalEditorToEditorArea(undefined),
			canMoveModalEditorToEditorArea({}),
			canMoveModalEditorToEditorArea(new TestEditor(undefined)),
			canMoveModalEditorToEditorArea(new TestEditor({})),
			canMoveModalEditorToEditorArea(new TestEditor({ canMoveToEditorArea: true })),
			canMoveModalEditorToEditorArea(new TestEditor({ canMoveToEditorArea: false })),
		], [true, true, true, true, true, false]);
	});

	test('header height prefers measured offsetHeight over fallback constant', () => {
		// Before layout: offsetHeight === 0 → fallback constant is used
		assert.strictEqual(resolveHeaderHeight({ offsetHeight: 0 }), MODAL_HEADER_HEIGHT);

		// Default header measured at 33px
		assert.strictEqual(resolveHeaderHeight({ offsetHeight: 33 }), 33);

		// Compact-header variant measured at 40px → measured value wins
		assert.strictEqual(resolveHeaderHeight({ offsetHeight: 40 }), 40);
	});

	test('compact-header uses workbench foreground instead of titleBar foreground for title, icons and separator', () => {
		const style = document.createElement('style');
		style.textContent = '.monaco-modal-editor-block { --vscode-titleBar-activeForeground: rgb(255, 255, 255); --vscode-editor-background: rgb(0, 0, 0); --vscode-sideBar-background: rgb(255, 255, 255); --vscode-agentsPanel-foreground: rgb(24, 45, 79); }';
		document.head.appendChild(style);

		const block = document.createElement('div');
		block.className = 'monaco-modal-editor-block';
		const part = block.appendChild(document.createElement('div'));
		part.className = 'modal-editor-part';
		const header = part.appendChild(document.createElement('div'));
		header.className = 'modal-editor-header';
		const title = header.appendChild(document.createElement('div'));
		title.className = 'modal-editor-title';
		const actions = header.appendChild(document.createElement('div'));
		actions.className = 'modal-editor-action-container';
		const actionsContainer = actions.appendChild(document.createElement('div'));
		actionsContainer.className = 'actions-container';
		const icon = actionsContainer.appendChild(document.createElement('a'));
		icon.className = 'action-label codicon';
		const separator = actions.appendChild(document.createElement('div'));
		separator.className = 'modal-editor-action-separator';
		document.body.appendChild(block);

		disposables.add(toDisposable(() => {
			block.remove();
			style.remove();
		}));

		assert.strictEqual(getComputedStyle(part).backgroundColor, 'rgb(0, 0, 0)');
		assert.strictEqual(getComputedStyle(header).color, 'rgb(255, 255, 255)');
		assert.strictEqual(getComputedStyle(title).color, 'rgb(255, 255, 255)');
		assert.strictEqual(getComputedStyle(icon).color, 'rgb(255, 255, 255)');
		assert.strictEqual(getComputedStyle(separator).backgroundColor, 'rgb(255, 255, 255)');

		block.classList.add('compact-header');
		assert.strictEqual(getComputedStyle(part).backgroundColor, 'rgb(255, 255, 255)');
		assert.strictEqual(getComputedStyle(header).backgroundColor, 'rgb(255, 255, 255)');
		assert.strictEqual(getComputedStyle(header).color, 'rgb(24, 45, 79)');
		assert.strictEqual(getComputedStyle(title).color, 'rgb(24, 45, 79)');
		assert.strictEqual(getComputedStyle(icon).color, 'rgb(24, 45, 79)');
		assert.strictEqual(getComputedStyle(separator).backgroundColor, 'rgb(24, 45, 79)');
	});

	test('ai-customization management sidebar, content and group headers use workbench colors instead of editor background', () => {
		const style = document.createElement('style');
		style.textContent = '.ai-customization-management-editor { --vscode-editor-background: rgb(0, 0, 0); --vscode-sideBar-background: rgb(255, 255, 255); --vscode-sideBar-foreground: rgb(0, 0, 0); --vscode-agentsPanel-background: rgb(242, 238, 227); --vscode-agentsPanel-foreground: rgb(23, 19, 15); --vscode-sideBarSectionHeader-foreground: rgb(255, 255, 255); --vscode-foreground: rgb(0, 0, 0); }';
		document.head.appendChild(style);

		const editor = document.createElement('div');
		editor.className = 'ai-customization-management-editor';
		const sidebarContent = editor.appendChild(document.createElement('div'));
		sidebarContent.className = 'sidebar-content';
		const managementContent = editor.appendChild(document.createElement('div'));
		managementContent.className = 'management-content';
		const groupHeader = managementContent.appendChild(document.createElement('div'));
		groupHeader.className = 'ai-customization-group-header';
		const groupLabel = groupHeader.appendChild(document.createElement('span'));
		groupLabel.className = 'group-label';
		document.body.appendChild(editor);

		disposables.add(toDisposable(() => {
			editor.remove();
			style.remove();
		}));

		assert.strictEqual(getComputedStyle(sidebarContent).backgroundColor, 'rgb(255, 255, 255)');
		assert.strictEqual(getComputedStyle(sidebarContent).color, 'rgb(0, 0, 0)');
		assert.strictEqual(getComputedStyle(managementContent).backgroundColor, 'rgb(242, 238, 227)');
		assert.strictEqual(getComputedStyle(managementContent).color, 'rgb(23, 19, 15)');
		assert.strictEqual(getComputedStyle(groupLabel).color, 'rgb(0, 0, 0)');
	});

	test('non-compact modal editor uses panel background and workbench foreground on mixed-polarity themes', () => {
		const style = document.createElement('style');
		style.textContent = '.monaco-modal-editor-block { --vscode-editor-background: rgb(0, 0, 0); --vscode-panel-background: rgb(255, 255, 255); --vscode-foreground: rgb(0, 0, 0); }';
		document.head.appendChild(style);

		const block = document.createElement('div');
		block.className = 'monaco-modal-editor-block';
		const part = block.appendChild(document.createElement('div'));
		part.className = 'modal-editor-part';
		document.body.appendChild(block);

		disposables.add(toDisposable(() => {
			block.remove();
			style.remove();
		}));

		assert.strictEqual(getComputedStyle(part).backgroundColor, 'rgb(255, 255, 255)');
		assert.strictEqual(getComputedStyle(part).color, 'rgb(0, 0, 0)');
	});
});
