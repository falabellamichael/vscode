/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { generateUuid } from '../../../../base/common/uuid.js';
import { hashAsync } from '../../../../base/common/hash.js';
import { stableStringify } from '../../../../base/common/objects.js';
import { isAbsolute, relative, sep } from '../../../../base/common/path.js';
import { isEqualOrParent } from '../../../../base/common/extpath.js';
import { isLinux } from '../../../../base/common/platform.js';
import { localize } from '../../../../nls.js';
import type { ReachAssistantChatMessage, ReachAssistantRequest, ReachAssistantResponse } from '../../../native/common/native.js';
import { ResponsePartKind, ToolCallStatus, ToolResultContentType, type ToolCallResult, type Turn } from '../../common/state/sessionState.js';
import { findReachTools, isWorkspaceBuildRequest, ReachChatToolCatalog, type ReachChatToolSelection } from './reachChatToolCatalog.js';
import { qualifyFolderCreate } from './reachChatFolders.js';
import { ReachWebVerification } from './reachWebVerification.js';
import type { ReachClientTools } from './reachClientTools.js';

function addRecoveryInstruction(messages: ReachAssistantChatMessage[], instruction: string): void {
	if (messages[0]?.role === 'system') {
		messages[0] = { role: 'system', content: messages[0].content + '\n\n' + instruction };
	} else {
		messages.unshift({ role: 'system', content: instruction });
	}
}

function dedupeConsecutiveLines(text: string): string {
	const lines = text.split('\n');
	if (lines.length < 4) { return text; }
	const out: string[] = [];
	for (let i = 0; i < lines.length; i++) {
		let run = 1;
		while (i + run < lines.length && lines[i + run] === lines[i]) { run++; }
		if (run >= 3) {
			out.push(lines[i], `(the previous line repeated ${run} times)`);
		} else {
			for (let k = 0; k < run; k++) { out.push(lines[i]); }
		}
		i += run - 1;
	}
	return out.join('\n');
}

async function modelResult(result: ToolCallResult, isTextFileRead: boolean, textReadBudgetBytes: number, earlierResult?: string): Promise<string> {
	const envelope = {
		success: result.success, message: result.pastTenseMessage, error: result.error,
		content: result.content?.map(item => {
			if (item.type === ToolResultContentType.EmbeddedResource) { return { type: item.type, contentType: item.contentType }; }
			if (!isTextFileRead && item.type === ToolResultContentType.Text && typeof item.text === 'string' && item.text.length > 512) {
				return { ...item, text: dedupeConsecutiveLines(item.text) };
			}
			return item;
		}),
		structuredContent: result.structuredContent,
	};
	const output = JSON.stringify(envelope);
	if (output === earlierResult || output.length <= 24_000) { return output; }
	const part = result.content?.[0];
	if (isTextFileRead && result.success && part?.type === ToolResultContentType.Text) {
		try {
			const file = JSON.parse(part.text);
			if (typeof file.content === 'string' && !file.content.includes('\0')) {
				// Native Cohere unwraps the envelope before sending it. Budget the actual UTF-8
				// source so ordinary files are not reduced to 60 lines by double JSON escaping.
				if (Buffer.byteLength(file.content, 'utf8') <= textReadBudgetBytes) { return output; }
				const lines = file.content.split(/\r\n|\r|\n/);
				const start = file.startLine ?? 1;
				let count = 0, length = 0;
				while (count < Math.min(60, lines.length) && length + lines[count].length + 1 <= 6000) {
					length += lines[count++].length + 1;
				}
				if (count) {
					const end = start + count - 1;
					const content = lines.slice(0, count).join('\n');
					const totalLines = file.totalLines ?? lines.length;
					const excerpt: Record<string, unknown> = { uri: file.uri, bytes: file.bytes, startLine: start, endLine: end, totalLines, content, fileRevision: start + ':' + end + ':' + await hashAsync(content) };
					if (end < totalLines) {
						excerpt.readGuide = `This excerpt shows only lines ${start}-${end} of ${totalLines}. The rest of the file exists but is not shown here. Before editing, read the exact section you want to change with startLine and endLine, then edit using the fileRevision returned by that read — revision edits replace exactly those lines and never fail on matching.`;
					}
					const bounded = JSON.stringify({ success: true, content: [{ type: ToolResultContentType.Text, text: JSON.stringify(excerpt) }] });
					if (bounded.length <= 24_000) { return bounded; }
				}
			}
		} catch { /* Other tool formats use the bounded preview below. */ }
	}
	return JSON.stringify({ success: result.success, truncated: true, preview: output.slice(0, 24_000) });
}

/** Restore completed call/result pairs when the editor resumes an interrupted task. */
export async function resumedReachToolMessages(turns: readonly Turn[], task: string): Promise<ReachAssistantChatMessage[]> {
	const start = turns.findLastIndex(turn => turn.message.text === task);
	if (start < 0) { return []; }
	const history: ReachAssistantChatMessage[] = [];
	let remaining = 48 * 1024;
	for (let i = turns.length - 1; i >= start; i--) {
		for (const part of [...turns[i].responseParts].reverse()) {
			if (part.kind !== ResponsePartKind.ToolCall) { continue; }
			const call = part.toolCall;
			if (call.status !== ToolCallStatus.Completed || typeof call.toolInput !== 'string' || call.error?.code === 'tool-result-denied') { continue; }
			const name = call.toolName.replace(/[^a-zA-Z0-9_-]/g, '_');
			const content = await modelResult(call, name === 'reach_read_file', 0);
			const size = Buffer.byteLength(call.toolInput) + Buffer.byteLength(content);
			if (size > remaining) { continue; }
			remaining -= size;
			history.unshift(
				{ role: 'assistant', content: '', toolCalls: [{ id: call.toolCallId, name, arguments: call.toolInput }] },
				{ role: 'tool', name, toolCallId: call.toolCallId, content },
			);
		}
	}
	return history;
}

export async function runReachChatToolTurn(
	messages: ReachAssistantChatMessage[], prompt: string, client: ReachClientTools, sender: string | undefined, turnId: string, signal: AbortSignal,
	request: (request: ReachAssistantRequest, id: string, holdFinalText: boolean, buffered: boolean) => Promise<ReachAssistantResponse>, nextPart: (text: string) => void,
	selection?: ReachChatToolSelection,
): Promise<ReachAssistantResponse> {
	const catalog = new ReachChatToolCatalog(client.definitions(sender), prompt, selection);
	const web = new ReachWebVerification(prompt);
	const destination = selection?.attachedFolders?.length === 1 && /\b(?:in|into|inside|under)\s+(?:(?:this|that|the attached|the selected)\s+(?:folder|directory))\b/i.test(prompt)
		? selection.attachedFolders[0] : undefined;
	const completed = new Map<string, { name: string; args: string; result: string }>();
	const textReads = new Map<string, { id: string; result: string }>();
	const lastResults = new Map<string, string>();
	const openedPages = new Map<string, string>();
	const savedEdits = new Map<string, { id: string; oldText: string; newText: string; observed: boolean }>();
	const outcomes: string[] = [];
const writeLikeCopilotTools = new Set(['copilot_createFile', 'copilot_applyPatch', 'copilot_editFiles', 'copilot_insertEdit', 'copilot_replaceString', 'copilot_multiReplaceString', 'copilot_editNotebook', 'copilot_createNewJupyterNotebook', 'copilot_createDirectory']);
	let wroteFile = false;
	let remainingTextReadBytes = selection?.textReadBudgetBytes ?? 0;
	let unchangedReads = 0;
	let omitRepeatedReader = false;
	let omitRepeatedOpener = false;
	let consecutiveUnavailableRounds = 0;
	const unavailableToolNames: string[] = [];
	let retriedRejectedToolCall = false;
	let retriedShorterOutput = false;
	let retriedShorterOutputTwice = false;
	let retriedAfterTools = false;
	let retryWithoutStreaming = false;
	let incompleteWrite: { key: string; round: number } | undefined;
	let retriedWritePromise = false;
	let retriedCompletionClaim = false;
	let consecutiveEditFailures = 0;
	let totalEditFailures = 0;
	let editSteeringAdded = false;
	let emptyAttachedFolder = false;
	let guidedBrowserRetry = false;
	let acknowledgedBrowserCheck = false;
	const modificationTask = selection?.editingAttachedFiles || isWorkspaceBuildRequest(prompt)
		|| /\b(?:fix|edit|improve|update|add|change|create|build|write|repair|implement|redesign|modify|rebuild|rewrite)\b/i.test(prompt);
	let roundsSinceWriteAttempt = 0;
	let writeFreeRounds = 0;
	let paralysisWarnings = 0;
	let hitRoundCeiling = false;
	for (let round = 0, stalledRounds = 0; stalledRounds < 32; round++) {
		signal.throwIfAborted();
		stalledRounds++;
		if (round === 20) {
			addRecoveryInstruction(messages, 'You have used many tool rounds in this turn. Finish the current step and give the final answer soon instead of starting new exploration.');
		} else if (round === 35 || (round > 35 && round < 80 && round % 15 === 5)) {
			addRecoveryInstruction(messages, 'This turn is spending its tool rounds. Give the final answer in one of your next replies, reporting what is complete and what remains.');
		} else if (writeFreeRounds >= 80) {
			// Only write-free rounds spend the budget: a turn that keeps successfully
			// changing files is doing real work and is never cut off.
			hitRoundCeiling = true;
			break;
		}
		catalog.refresh(client.definitions(sender));
		const tools = catalog.specs(wroteFile).filter(tool => !(omitRepeatedReader && catalog.originalName(tool.name) === 'reach_read_file') && !(omitRepeatedOpener && catalog.originalName(tool.name) === 'open_browser_page') && !(emptyAttachedFolder && tool.name === 'copilot_listDirectory'));
		omitRepeatedReader = false;
		omitRepeatedOpener = false;
		const roundRequest = { messages, tools };
		let response = await request(roundRequest, round === 0 ? turnId : generateUuid(), web.pending, retryWithoutStreaming);
		if (response.kind === 'unavailable' && response.reason === 'providerStreamError') {
			// The subscription gateway released this failed inference before settlement. Its partial calls never ran.
			signal.throwIfAborted();
			nextPart(localize('reach.chat.retryModelStream', "Retrying the interrupted response from {0}…", response.model));
			response = await request(roundRequest, generateUuid(), web.pending, retryWithoutStreaming);
			if (response.kind === 'unavailable' && response.reason === 'providerStreamError') {
				throw new Error(localize('reach.chat.modelStreamFailed', "{0} stopped generating during both attempts at this step.", response.model));
			}
		}
		if (response.kind === 'unavailable' && response.reason === 'responseTooLarge') {
			if (retryWithoutStreaming) { return response; }
			retryWithoutStreaming = true;
			nextPart('');
			addRecoveryInstruction(messages, 'The gateway rejected the previous generation because its response was too large. No action from that generation ran. Continue from the confirmed tool results. Make the next change a smaller complete file or one targeted edit, then continue with the remaining work.');
			continue;
		}
		if (response.kind === 'unavailable' && !response.reason && (response.status === undefined || [502, 503, 504].includes(response.status))
			&& outcomes.length && !retriedAfterTools) {
			// Retry inference with the existing call/result pairs. No editor action is pending here.
			retriedAfterTools = true;
			nextPart('');
			continue;
		}
		if (response.kind === 'unavailable' && response.reason === 'outputLimit') {
			nextPart('');
			if (tools.length && (selection?.editingAttachedFiles || messages.at(-1)?.role === 'tool')) {
				if (!retriedShorterOutput) {
					retriedShorterOutput = true;
					addRecoveryInstruction(messages, 'Your last generation reached the output limit. No tool from that generation ran. Think briefly — long reasoning consumes the same output budget. Use a smaller complete tool call for the next change, then continue from its result. Keep the final answer brief.');
					continue;
				}
				if (!retriedShorterOutputTwice) {
					retriedShorterOutputTwice = true;
					addRecoveryInstruction(messages, 'The output limit stopped your generation again. Work in smaller pieces: write or edit one short section per tool call, think in one or two sentences, and answer in one short paragraph. Do not attempt a large single change.');
					continue;
				}
			}
			return response;
		}
		if (response.kind === 'unavailable' && response.reason === 'inputTooLarge' && catalog.reducePageSize()) {
			continue; // The gateway rejected the schemas before inference. Retry a smaller page.
		}
		if (response.kind === 'unavailable' && response.reason === 'rejectedToolCall') {
			nextPart(''); // Separate any streamed preamble from the retry's response.
			if (retriedRejectedToolCall) { return response; }
			retriedRejectedToolCall = true;
			// The provider returned an unrecognized call as text. It is not safe to dispatch or replay it.
			addRecoveryInstruction(messages, 'The previous tool request was printed as XML and did not run. Call only functions listed in the current tool schemas. Use the provided tool protocol without introductory text or XML. Use the tool catalog search if available, or give the final answer if you already have enough information.');
				nextPart(localize('reach.chat.correctingFormat', "The response used a format that does not run. Correcting it now…"));
			continue;
		}
		if (outcomes.length && (response.kind === 'unavailable' || response.kind === 'allowanceExhausted')) {
			const message = response.kind === 'allowanceExhausted'
				? localize('reach.chat.allowanceAfterTools', "The REACH request allowance ran out before the final reply.")
				: response.kind === 'unavailable' && response.status !== undefined ? localize('reach.chat.httpErrorAfterTools', "The REACH request for {0} failed after running tools (HTTP {1}).", response.model, response.status)
				: localize('reach.chat.unavailableAfterTools', "REACH could not finish the reply after running tools.");
			throw new Error(message + '\n\n' + outcomes.join('\n'));
		}
		if (response.kind !== 'reply' || !response.toolCalls?.length) {
			if (response.kind === 'reply' && web.pending) {
				const verification = web.beforeFinish(client.definitions(sender).some(tool => tool.name === 'open_browser_page' && (!selection?.includeTool || selection.includeTool(tool.name))));
				if (verification?.instruction) {
					addRecoveryInstruction(messages, verification.instruction);
					nextPart('');
					continue;
				}
				if (verification?.reply) { return { kind: 'reply', text: verification.reply }; }
			}
			if (response.kind === 'reply' && (incompleteWrite || (selection?.editingAttachedFiles && !wroteFile))
				&& /\b(?:let me|I(?:['’]ll| will| am going to| need to))\s+(?:now\s+)?(?:write|edit|create|finish|update|fix|apply|replace)\b[^.!?\n]{0,120}[.!\s]*$/i.test(response.text)) {
				nextPart(response.text);
				if (retriedWritePromise) {
					throw new Error(localize('reach.chat.writeNotConfirmed', "The model stopped after describing its next change. The editor did not confirm that write."));
				}
				retriedWritePromise = true;
				addRecoveryInstruction(messages, 'Your last answer only announced a file change. To carry it out, send the advertised editor function with the replacement text now. If you cannot proceed, explain the blocker.');
				continue;
			}
			if (response.kind === 'reply' && !wroteFile && !retriedCompletionClaim && (selection?.editingAttachedFiles || isWorkspaceBuildRequest(prompt))
				&& /(?:have|has) been (?:successfully )?(?:created|saved|written|generated)|successfully (?:created|saved|written)|(?:files?|project|game|app) (?:are|is) (?:now )?ready|you(?:'re| are) all set|can now open/i.test(response.text)) {
				nextPart(response.text);
				retriedCompletionClaim = true;
				addRecoveryInstruction(messages, 'Your last answer claimed files were created or saved, but no file tool ran in this turn, so nothing exists on disk. Do not report success for actions that did not run. Carry out the change now with the advertised file function, or explain the blocker honestly.');
				nextPart(localize('reach.chat.correctingClaim', "That report claimed success before anything was saved. Correcting it now…"));
				continue;
			}
			return response.kind === 'reply' && outcomes.length ? { ...response, actions: outcomes } : response;
		}
		signal.throwIfAborted();
		nextPart(response.text);
		messages.push({ role: 'assistant', content: response.text, toolCalls: response.toolCalls, toolPlan: response.toolPlan });
		let roundHadRealCall = false;
		let roundHadWriteAttempt = false;
		let roundHadWriteSuccess = false;
		for (const call of response.toolCalls) {
			signal.throwIfAborted();
			let result: string;
			const originalName = catalog.originalName(call.name);
			if (originalName === 'reach_create_file' || originalName === 'reach_edit_file' || writeLikeCopilotTools.has(originalName ?? '')) { roundHadWriteAttempt = true; }
			const previous = completed.get(call.id);
			if (previous) {
				result = previous.name === call.name && previous.args === call.arguments ? previous.result : 'Rejected: a tool call ID was reused with different input.';
			} else if (!originalName && call.name !== findReachTools) {
				if (!unavailableToolNames.includes(call.name)) { unavailableToolNames.push(call.name); }
				const advertised = tools.map(tool => tool.name).join(', ');
				const workspaceFiles = ['reach_create_file', 'reach_read_file', 'reach_edit_file'].filter(name => catalog.originalName(name)).join(', ');
				result = `The function "${call.name}" does not exist in this chat and the call did not run. The exact names in the current tool schemas are: ${advertised}.${workspaceFiles ? ` Enabled workspace file functions are: ${workspaceFiles}.` : ''} Retry the intended action with one exact function name; do not guess names by pattern.`;
			} else {
				roundHadRealCall = true;
				if (originalName) { catalog.pinWorkspaceFileTool(call.name); }
				let args: Record<string, unknown> | undefined;
				try {
					const parsed: unknown = JSON.parse(call.arguments || '{}');
					if (parsed !== null && typeof parsed === 'object' && !Array.isArray(parsed)) { args = parsed as Record<string, unknown>; }
				} catch { /* Send a tool error back so the model can correct its arguments. */ }
				// Some models return paths relative to the folder chip. The editor resolves them
				// against the workspace. Qualify new files before the editor displays its approval.
				const folderError = destination && args && originalName === 'reach_create_file' ? qualifyFolderCreate(destination, args) : undefined;
				// Directory listings may give the model an absolute path. The editor file
				// tools take a relative path and still enforce containment and link checks.
				if (args && selection?.localRoot && (originalName === 'reach_create_file' || originalName === 'reach_edit_file' || originalName === 'reach_read_file')) {
					const path = args.path;
					if (typeof path === 'string' && isAbsolute(path) && !/^[\\/]{2}/.test(path) && !path.split(/[\\/]/).includes('..')) {
						const local = relative(selection.localRoot, path);
						if (local && !isAbsolute(local) && local !== '..' && !local.startsWith('..' + sep)) { args.path = local; }
					}
				}
				const fileKey = JSON.stringify([args?.workspaceFolder, args?.path]);
				const savedEdit = savedEdits.get(fileKey);
				const writeField = originalName === 'reach_create_file' ? 'content' : originalName === 'reach_edit_file' ? 'newText' : undefined;
				if (!args) {
					result = 'Invalid tool arguments. Supply a JSON object.';
				} else if (folderError) {
					result = JSON.stringify({ success: false, message: folderError });
				} else if (destination && originalName === 'copilot_listDirectory' && typeof args.path === 'string' && isAbsolute(args.path)
					&& !isEqualOrParent(args.path.replace(/\\/g, '/'), destination.absolutePath.replace(/\\/g, '/'), !isLinux, true)) {
					result = JSON.stringify({ success: false, message: 'The requested project destination is ' + destination.absolutePath + '. List that folder directly or start creating the requested files there.' });
				} else if (writeField && typeof args[writeField] !== 'string' && client.definitions(sender).find(t => t.name === originalName)?.inputSchema?.required?.includes(writeField)) {
					const key = writeField + fileKey;
					if (incompleteWrite?.key === key && incompleteWrite.round !== round) {
						throw new Error(localize('reach.chat.missingWriteArgument', "The model twice omitted the required {0} argument to {1}. That file operation did not run.", writeField, call.name));
					}
					incompleteWrite = { key, round };
					result = JSON.stringify({ success: false, message: `Missing required string "${writeField}". The file operation did not run. Include the actual text in this argument. Use reach_edit_file for an existing file.` });
				} else if (call.name === findReachTools) {
					catalog.refresh(client.definitions(sender));
					const offset = args.offset ?? 0;
					result = typeof args.query === 'string' && typeof offset === 'number' && Number.isSafeInteger(offset) && offset >= 0
						? JSON.stringify(catalog.selectPage(args.query, offset)) : 'Supply a query string and an optional nonnegative integer offset.';
				} else if (originalName === 'reach_edit_file' && savedEdit?.observed && savedEdit.oldText === args.oldText && savedEdit.newText === args.newText && args.fileRevision === undefined) {
					result = JSON.stringify({ success: true, alreadyAppliedToolCall: savedEdit.id, path: args.path, message: 'This exact edit already succeeded. The latest full read contains the replacement. Continue with a different change if work remains.' });
				} else if (originalName === 'open_browser_page' && typeof args.url === 'string' && openedPages.has(args.url)) {
					// The earlier editor receipt contains the live page ID needed by read_page.
					result = openedPages.get(args.url)!;
					omitRepeatedOpener = true;
				} else {
					const execution = await client.invoke(turnId, originalName!, args, signal, sender);
					const wasPending = web.pending;
					web.observe(originalName!, args, execution);
					if (wasPending && !web.pending && !acknowledgedBrowserCheck) {
						acknowledgedBrowserCheck = true;
						addRecoveryInstruction(messages, 'The editor observed a successful interaction on the saved page. If the requested behavior is verified, finish with the result. If a required interaction is still broken, inspect the source and fix it.');
					} else if (web.pending && !execution.success && !guidedBrowserRetry && ['click_element', 'type_in_page', 'run_playwright_code'].includes(originalName!)) {
						guidedBrowserRetry = true;
						addRecoveryInstruction(messages, 'For a failed browser action, use the actual ref from the snapshot or a selector from the saved source. If the intended control is hidden, unresponsive or still fails, read the relevant HTML, CSS or script and fix the cause before repeating the check.');
					}
					if (execution.error?.code === 'denied') {
						return { kind: 'reply', text: localize('reach.chat.toolDenied', "Permission to run {0} was declined. I stopped this turn.", call.name) };
					}
					if (execution.error?.code === 'tool-result-denied') {
						return { kind: 'reply', text: localize('reach.chat.resultDenied', "{0} ran. Sharing its result was declined, so I stopped this turn.", call.name) };
					}
					if (execution.error?.code === 'clientDisconnected') {
						return { kind: 'reply', text: localize('reach.chat.clientDisconnected', "The editor disconnected while running {0}. Check its result before repeating the action.", call.name) };
					}
					const textRead = originalName === 'reach_read_file' && args.encoding !== 'base64';
					const key = JSON.stringify([args.workspaceFolder, args.path, args.startLine, args.endLine]);
					const earlier = textRead ? textReads.get(key) : undefined;
					result = await modelResult(execution, textRead, remainingTextReadBytes, earlier?.result);
					if (execution.success) {
						const operation = stableStringify([originalName, args]);
						if (lastResults.get(operation) !== result) {
							lastResults.set(operation, result);
							stalledRounds = 0;
						} else if (stalledRounds === 2) {
							addRecoveryInstruction(messages, 'The editor returned the same result for the same tool and arguments again. Use the existing result, change the arguments or choose a different action to make progress.');
						}
					}
					if (destination && originalName === 'copilot_listDirectory' && execution.success && typeof args.path === 'string'
						&& !relative(destination.absolutePath, args.path) && /folder is empty/i.test(result)) {
						emptyAttachedFolder = true;
						addRecoveryInstruction(messages, 'The editor confirmed the attached destination folder is empty. Start creating the requested project there. No further directory discovery is needed.');
					}
					if (textRead && savedEdit && args.startLine === undefined && args.endLine === undefined) {
						const part = execution.content?.[0];
						savedEdit.observed = false;
						if (execution.success && part?.type === ToolResultContentType.Text) {
							try {
								const source = JSON.parse(part.text).content;
								savedEdit.observed = typeof source === 'string' && !source.includes(savedEdit.oldText) && source.includes(savedEdit.newText);
							} catch { /* Only a complete editor text read can confirm the change. */ }
						}
					}
					if (execution.success && originalName === 'open_browser_page' && typeof args.url === 'string') {
						openedPages.set(args.url, result);
					}
					if (execution.success && (originalName === 'close_page' || originalName === 'navigate_page')) { openedPages.clear(); }
					if (textRead && execution.success) {
						// Compare fresh editor results so changes made outside this turn are visible too.
						if (earlier?.result === result) {
							unchangedReads++;
							// While repairing an attachment, the same reader is needed to obtain a narrower revision.
							omitRepeatedReader = wroteFile || (!selection?.editingAttachedFiles && unchangedReads >= 2);
							result = JSON.stringify({ success: true, unchangedFromToolCall: earlier.id, path: args.path, message: 'The editor returned the same source as the earlier read. Use that source for the edit. To inspect another section, supply a different startLine and endLine.' });
						} else {
							unchangedReads = 0;
							remainingTextReadBytes = Math.max(0, remainingTextReadBytes - Buffer.byteLength(result));
							textReads.set(key, { id: call.id, result });
						}
					} else if (execution.success && (originalName === 'reach_create_file' || originalName === 'reach_edit_file' || writeLikeCopilotTools.has(originalName ?? ''))) {
						wroteFile = true;
						roundHadWriteSuccess = true;
						if (originalName === 'reach_edit_file') { consecutiveEditFailures = 0; }
						incompleteWrite = undefined;
						retriedShorterOutput = false;
						retriedShorterOutputTwice = false;
						retriedWritePromise = false;
						openedPages.clear();
						unchangedReads = 0;
						if (originalName === 'reach_edit_file' && typeof args.oldText === 'string' && typeof args.newText === 'string') {
							savedEdits.set(fileKey, { id: call.id, oldText: args.oldText, newText: args.newText, observed: false });
						} else { savedEdits.delete(fileKey); }
					}
					if (!execution.success && originalName === 'reach_edit_file') {
						consecutiveEditFailures++;
						totalEditFailures++;
						// Two failures in a row, or four across the turn even with successes in between:
						// the model is flailing against exact matching and needs the positional escape.
						if (!editSteeringAdded && (consecutiveEditFailures === 2 || totalEditFailures === 4)) {
							editSteeringAdded = true;
							addRecoveryInstruction(messages, 'Edits keep failing. Stop retrying the same edit unchanged. If the error says the passage appears several times, re-issue the same edit with nearLine set to the line of the occurrence you mean — nearLine picks it, no other change needed. If the passage is not found, read the exact section with startLine and endLine, copy oldText verbatim from that read, and try once more. Do not repeat an edit that already failed without changing oldText or adding nearLine.');
						}
					}
					if (execution.success) {
						outcomes.push(localize('reach.chat.actionCompleted', "- {0} completed.", call.name));
					} else {
						const detail = execution.error?.message?.replace(/\s+/g, ' ').trim() ?? '';
						const firstSentence = detail.slice(0, detail.indexOf('. ') < 0 ? 160 : Math.min(detail.indexOf('. ') + 1, 160));
						outcomes.push(firstSentence
							? localize('reach.chat.actionFailedReason', "- {0} failed: {1}", call.name, firstSentence)
							: localize('reach.chat.actionFailed', "- {0} reported a failure.", call.name));
					}
				}
				completed.set(call.id, { name: call.name, args: call.arguments, result });
			}
			messages.push({ role: 'tool', name: call.name, toolCallId: call.id, content: result });
		}
		if (roundHadRealCall) {
			consecutiveUnavailableRounds = 0;
			unavailableToolNames.length = 0;
		} else {
			consecutiveUnavailableRounds++;
		}
		if (roundHadWriteAttempt) { roundsSinceWriteAttempt = 0; } else { roundsSinceWriteAttempt++; }
		if (roundHadWriteSuccess) { writeFreeRounds = 0; } else { writeFreeRounds++; }
		if (modificationTask && roundsSinceWriteAttempt >= 8) {
			// Exploration paralysis: rounds of reading and searching without a single write
			// attempt. Small models wait for perfect understanding that never arrives.
			roundsSinceWriteAttempt = 0;
			paralysisWarnings++;
			addRecoveryInstruction(messages, paralysisWarnings === 1
				? 'You have spent many rounds reading and searching without changing anything. You have enough information to act: apply the most important change now with reach_edit_file or reach_create_file, even if it is partial — a small improvement is progress and you can refine it after. Do not read or search again before making an edit. If the requested change is genuinely impossible to locate, finish now and explain what you found.'
				: 'You still have not changed anything. Stop exploring. Make one concrete edit with reach_edit_file or reach_create_file in your next reply, or finish with an explanation of the blocker.');
		}
		if (consecutiveUnavailableRounds >= 3) {
			const loopMessage = outcomes.length
				? localize('reach.chat.unavailableLoop', "The model kept calling functions that do not exist instead of the advertised tools, so the turn stopped. The completed work is preserved below.")
				: localize('reach.chat.unavailableLoopNoWork', "The model repeatedly called functions that do not exist instead of the advertised tools. The turn stopped before any work completed.");
			const names = unavailableToolNames.length ? localize('reach.chat.unavailableLoopNames', "Unavailable function names: {0}.", unavailableToolNames.join(', ')) : '';
			throw new Error(loopMessage + (names ? '\n' + names : '') + (outcomes.length ? '\n\n' + outcomes.join('\n') : ''));
		}
	}
	if (hitRoundCeiling) {
		// Never die on the round ceiling with an error. Take one final tool-free inference so
		// the model itself summarizes what it did and what remains, then end the turn normally.
		addRecoveryInstruction(messages, 'Tool rounds are exhausted. Do not call any function. Write the final answer now in a few sentences: what you accomplished or changed, and what remains unfinished.');
		try {
			const final = await request({ messages, tools: [] }, generateUuid(), false, retryWithoutStreaming);
			if (final.kind === 'reply' && final.text.trim()) {
				const reply: ReachAssistantResponse = { kind: 'reply', text: final.text };
				return outcomes.length ? { ...reply, actions: outcomes } : reply;
			}
		} catch { /* Fall through to the preserved-work summary. */ }
		const capMessage = outcomes.length
			? localize('reach.chat.roundCap', 'The turn used its tool round limit before finishing. The completed work is preserved below.')
			: localize('reach.chat.roundCapNoWork', 'The turn used its tool round limit before finishing any work.');
		return { kind: 'reply', text: capMessage + (outcomes.length ? '\n\n' + outcomes.join('\n') : '') };
	}
	return { kind: 'reply', text: localize('reach.chat.toolsStalled', "The last 32 rounds returned no new successful editor result. Check the tool results above before continuing.") };
}
