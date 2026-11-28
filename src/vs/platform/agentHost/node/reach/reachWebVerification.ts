/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { extUriBiasedIgnorePathCase as resources } from '../../../../base/common/resources.js';
import { URI } from '../../../../base/common/uri.js';
import { localize } from '../../../../nls.js';
import { ToolResultContentType, type ToolCallResult } from '../../common/state/sessionState.js';

/** Track browser evidence for the actual file saved by the editor during a web build. */
export class ReachWebVerification {
	private entry: URI | undefined;
	private pageId: string | undefined;
	private tested = false;
	private prompted = false;
	private issue: string | undefined;
	private checkFailure: string | undefined;
	private readonly deferredActions = new Set<string>();
	private readEntry: URI | undefined;
	private readonly enabled: boolean;

	constructor(prompt: string) {
		this.enabled = !/\b(?:don['’]?t|do not|skip|without|no)\s+(?:any\s+)?(?:browser|preview|test|verify)|\b(?:i['’]?ll|i will|let me)\s+(?:handle\s+)?(?:test|verify)|\b(?:code|source) only\b/i.test(prompt);
	}

	get pending(): boolean { return !!this.entry && !this.tested; }

	observe(name: string, args: Record<string, unknown>, result: ToolCallResult): void {
		if (!this.enabled) { return; }
		const text = result.content?.flatMap(item => item.type === ToolResultContentType.Text ? [item.text] : []).join('\n') ?? '';
		if (name === 'reach_create_file' || name === 'reach_edit_file' || name === 'reach_read_file') {
			if (!result.success) { return; }
			if (typeof args.path === 'string' && /\.html?$/i.test(args.path)) {
				// Only an editor receipt identifies an addressable preview.
				try {
					const file = JSON.parse(text);
					if (typeof file.uri === 'string') {
						const uri = URI.parse(file.uri);
						if (uri.scheme === 'file' && !uri.authority) {
							if (name === 'reach_read_file') { this.readEntry = uri; }
							else if (!this.entry || /\/index\.html?$/i.test(uri.path)) { this.entry = uri; }
						}
					}
				} catch { /* No addressable file receipt. */ }
			}
			if (name === 'reach_read_file') { return; }
			if (!this.entry && typeof args.path === 'string' && /\.(?:[cm]?js|css)$/i.test(args.path)) { this.entry = this.readEntry; }
			this.tested = false;
			this.pageId = undefined;
			this.issue = undefined;
			this.checkFailure = undefined;
			this.deferredActions.clear();
			this.prompted = false;
			return;
		}
		if (!this.entry) { return; }
		if (name === 'open_browser_page' && typeof args.url === 'string') {
			try { if (!resources.isEqual(URI.parse(args.url), this.entry)) { return; } } catch { return; }
			this.pageId = result.success ? /^Page ID:\s*(\S+)/m.exec(text)?.[1] : undefined;
		} else if (!this.pageId || args.pageId !== this.pageId) {
			return;
		}
		const failure = /(?:\(pageError\)|\(requestFailed\)|\(console\) \[error\])[^\n]*/i.exec(text)?.[0];
		if (failure || !result.success) {
			if (!this.issue && !this.checkFailure) { this.prompted = false; }
			if (failure) { this.issue = failure.slice(0, 400); }
			else { this.checkFailure = (result.error?.message || text || 'The browser check failed.').slice(0, 400); }
			this.tested = false;
		} else if (this.pageId && !this.issue) {
			const resumedAction = typeof args.deferredResultId === 'string' && this.deferredActions.delete(args.deferredResultId);
			const scriptedAction = name === 'run_playwright_code' && (resumedAction || (typeof args.code === 'string' && /\.(?:click|dblclick|fill|press|type|check|uncheck|selectOption|dragTo|tap|dispatchEvent|insertText)\s*\(/.test(args.code)));
			const deferredId = /^\[deferredResultId=([^\]]+)\]/m.exec(text)?.[1];
			if (scriptedAction && deferredId) { this.deferredActions.add(deferredId); }
			else if (name === 'click_element' || name === 'type_in_page' || scriptedAction) {
				this.checkFailure = undefined;
				this.tested = true;
			}
		}
	}

	/** Give the agent one chance to verify or repair each saved revision before reporting its outcome. */
	beforeFinish(canOpen: boolean): { instruction?: string; reply?: string } | undefined {
		if (!this.pending || !this.entry) { return undefined; }
		if (canOpen && !this.prompted) {
			this.prompted = true;
			const problem = this.issue ?? this.checkFailure;
			return { instruction: problem
				? 'The browser reported this problem: ' + problem + '. Check the target against the current snapshot. If the app is broken, inspect and repair the saved source, then reopen the updated page and exercise its main interaction.'
				: 'Finish verifying the saved application before answering. Open the exact editor receipt URL ' + JSON.stringify(this.entry.toString()) + ' with open_browser_page and forceNew=true if the source changed. Inspect its startup snapshot and runtime errors. Exercise a meaningful user action and check its effect with the enabled browser tools. If it fails, read and repair the actual files and repeat the check.' };
		}
		const problem = this.issue ?? this.checkFailure;
		return { reply: problem
			? localize('reach.web.failed', "Saved {0}. Browser verification found a problem that still needs repair: {1}", this.entry.fsPath, problem)
			: localize('reach.web.unverified', "Saved {0}. The main interaction still needs browser verification.", this.entry.fsPath) };
	}
}
