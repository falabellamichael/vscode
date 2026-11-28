/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { RunOnceScheduler, timeout } from '../../../../base/common/async.js';
import { CancellationToken } from '../../../../base/common/cancellation.js';
import { DisposableStore } from '../../../../base/common/lifecycle.js';
import { SSEParser } from '../../../../base/common/sseParser.js';
import { localize } from '../../../../nls.js';
import type { ReachAssistantChatMessage, ReachAssistantToolCall, ReachAssistantToolSpec } from '../../../native/common/native.js';
import { toUpstreamMessage } from '../../../native/common/reachChatTransport.js';

/** Project an editor schema onto the JSON Schema subset accepted by native strict tool use. */
export function toCohereStrictTools(tools: readonly ReachAssistantToolSpec[]): ReachAssistantToolSpec[] {
	let fields = 0;
	function project(value: unknown): Record<string, unknown> | undefined {
		if (!value || typeof value !== 'object') { return undefined; }
		const schema = value as Record<string, unknown>;
		const { type, description, properties, required, items } = schema;
		if (typeof type !== 'string' || schema.$ref || schema.anyOf || schema.oneOf || schema.allOf) { return undefined; }
		const result: Record<string, unknown> = { type };
		if (typeof description === 'string') { result.description = description; }
		if (Array.isArray(schema.enum)) { result.enum = schema.enum; }
		if (type === 'array') {
			result.items = project(items);
			return result.items ? result : undefined;
		}
		if (type !== 'object') { return ['string', 'integer', 'number', 'boolean'].includes(type) ? result : undefined; }
		if (!properties || typeof properties !== 'object' || Array.isArray(properties)) { return undefined; }
		const children: Record<string, unknown> = Object.create(null);
		for (const [key, child] of Object.entries(properties)) {
			if (++fields > 200) { return undefined; }
			children[key] = project(child);
			if (!children[key]) { return undefined; }
		}
		result.properties = children;
		if (Array.isArray(required) && required.length) { result.required = required; }
		return result;
	}
	return tools.flatMap(tool => {
		const before = fields;
		// The editor accepts either a snapshot ref or a selector. Native strict mode
		// can advertise the snapshot-ref form without weakening the required target.
		const input = tool.parameters as Record<string, unknown> | undefined;
		const schema = tool.name === 'click_element' && Array.isArray(input?.anyOf) && Array.isArray(input.required)
			? project({ ...input, anyOf: undefined, required: [...input.required, 'ref'] })
			: project(input);
		// The ordinary opener also supports asking the user to share a tab. Its URL-taking
		// form is the one strict mode can express.
		if (schema && !schema.required && tool.name === 'open_browser_page' && (schema.properties as Record<string, unknown>)?.url) { schema.required = ['url']; }
		if (schema?.required) { return [{ ...tool, parameters: schema }]; }
		fields = before;
		return [];
	});
}

export function toCohereMessage(message: ReachAssistantChatMessage, supportsToolPlan = true): Record<string, unknown> {
	if (message.role === 'user' && message.imageUrls?.length) {
		return { role: 'user', content: [
			...(message.content ? [{ type: 'text', text: message.content }] : []),
			...message.imageUrls.map(url => ({ type: 'image_url', image_url: { url } })),
		] };
	}
	if (message.role === 'tool') {
		let data = message.content;
		if (message.name === 'reach_read_file') {
			// The editor and the agent each serialize tool results. Give the model the actual source text.
			try {
				const result = JSON.parse(data);
				if (result.truncated) {
					data = 'This file result exceeded the limit. Read smaller sections using startLine and endLine. Edit a short unique passage from a section using oldText and newText.';
				} else if (result.success && result.content?.length === 1) {
					const file = JSON.parse(result.content[0].text);
					if (typeof file.content === 'string') {
						const section = file.startLine ? 'Lines ' + file.startLine + '–' + file.endLine + ' of ' + file.totalLines + '. ' : '';
						const revision = typeof file.fileRevision === 'string' ? 'fileRevision: ' + file.fileRevision + '. ' + (section
							? 'A revision edit replaces only the lines below.'
							: 'Full file. For a requested rewrite, send one reach_edit_file call with this fileRevision and the complete revised file in newText. For a localized change, use a unique passage or a line range.') : '';
						// The loop's excerpt guidance carries the exact recovery path. Never drop it.
						const guide = typeof file.readGuide === 'string' ? file.readGuide + ' ' : '';
						data = section || revision || guide ? guide + section + revision + '\n\n' + file.content : file.content;
					}
				}
			} catch { /* Preserve other result formats, including editor errors. */ }
		}
		return { role: 'tool', tool_call_id: message.toolCallId, content: [{ type: 'document', document: { data } }] };
	}
	const wire = toUpstreamMessage(message);
	if (message.toolCalls?.length) {
		// Chat V2 rejects assistant text content alongside tool_calls when replaying a turn.
		delete wire.content;
		if (supportsToolPlan && message.toolPlan) { wire.tool_plan = message.toolPlan; }
	}
	return wire;
}

interface CohereStreamEvent {
	type: string;
	index?: number;
	delta?: {
		finish_reason?: string;
		error?: string;
		message?: { content?: { type?: string; text?: string }; tool_plan?: string; tool_calls?: { id?: string; function?: { name?: string; arguments?: string } } };
	};
	error?: { code?: number | string; message: string };
}

export class InterruptedCohereStreamError extends Error {
	constructor(message: string, readonly hasVisibleText: boolean, readonly toolName?: string, readonly retryWithoutStreaming = false, readonly outputLimited = false) { super(message); }
}

export class InvalidCohereToolArgumentsError extends InterruptedCohereStreamError {
	constructor(hasVisibleText: boolean, toolName?: string, message = 'Cohere returned malformed tool arguments. No call from that batch was run.') {
		super(message, hasVisibleText, toolName);
	}
}

/** Command A+ sends reasoning and visible text in separate blocks and ends each tool call explicitly. */
export async function readCohereStream(body: ReadableStream<Uint8Array>, onDelta: (text: string) => void, onActivity?: (tool?: string) => void) {
	let text = '';
	let toolPlan = '';
	let finishReason: string | undefined;
	let finishDetail: string | undefined;
	let providerError: CohereStreamEvent['error'];
	const calls = new Map<number, ReachAssistantToolCall>();
	const ended = new Set<number>();
	const textBlocks = new Set<number>();
	const parser = new SSEParser(({ data }) => {
		if (data.trim() === '[DONE]' || finishReason || providerError) { return; }
		const event: CohereStreamEvent = JSON.parse(data);
		if (event.error) { providerError = event.error; return; }
		const message = event.delta?.message;
		onActivity?.(event.type === 'tool-call-start' ? message?.tool_calls?.function?.name : undefined);
		const index = event.index ?? 0;
		switch (event.type) {
			case 'tool-plan-delta':
				toolPlan += message?.tool_plan ?? '';
				break;
			case 'content-start':
				if (message?.content?.type === 'text') { textBlocks.add(index); } else { textBlocks.delete(index); }
				// The initial block may already contain visible text.
				break;
			case 'content-end':
				textBlocks.delete(index);
				break;
			case 'tool-call-start': {
				const call = message?.tool_calls;
				if (!call?.id || !call.function?.name || calls.has(index)) { throw new Error('Invalid Cohere tool-call start.'); }
				calls.set(index, { id: call.id, name: call.function.name, arguments: call.function.arguments ?? '' });
				break;
			}
			case 'tool-call-delta': {
				const call = calls.get(index);
				if (!call || ended.has(index)) { throw new Error('Unexpected Cohere tool-call delta.'); }
				calls.set(index, { ...call, arguments: call.arguments + (message?.tool_calls?.function?.arguments ?? '') });
				break;
			}
			case 'tool-call-end':
				if (!calls.has(index)) { throw new Error('Unexpected Cohere tool-call end.'); }
				ended.add(index);
				break;
			case 'message-end':
				finishReason = event.delta?.finish_reason?.toUpperCase();
				finishDetail = event.delta?.error;
				break;
		}
		if ((event.type === 'content-start' || event.type === 'content-delta') && textBlocks.has(index) && message?.content?.text) {
			text += message.content.text;
			onDelta(message.content.text);
		}
	});
	const reader = body.getReader();
	try {
		while (!finishReason && !providerError) {
			const { done, value } = await reader.read();
			if (done) { break; }
			parser.feed(value);
		}
	} finally {
		try { await reader.cancel(); } finally { reader.releaseLock(); }
	}
	return finishCohereGeneration(text, toolPlan, [...calls.values()], finishReason, !!text, ended.size === calls.size, finishDetail, providerError);
}

function finishCohereGeneration(text: string, toolPlan: string, calls: ReachAssistantToolCall[], finishReason: string | undefined, visible: boolean, completeCalls: boolean, detail?: string, providerError?: CohereStreamEvent['error']) {
	const singleTool = calls.every(call => call.name === calls[0].name) ? calls[0]?.name : undefined;
	if (finishReason === 'ERROR' || finishReason === 'TIMEOUT') {
		const message = typeof detail === 'string' && detail.trim() ? detail.trim().replace(/\s+/g, ' ').slice(0, 300) : localize('cohere.stream.missingDetail', "The provider supplied no further details.");
		throw new InterruptedCohereStreamError(localize('cohere.stream.interrupted', "Cohere interrupted generation ({0}). {1}", finishReason, message), visible, singleTool, finishReason === 'ERROR' && !calls.length && !visible);
	}
	if (!providerError && finishReason === 'MAX_TOKENS' && !completeCalls) {
		// The token budget cut a tool batch off mid-arguments. Retrying the same generation
		// truncates again — the turn loop escalates with smaller-piece guidance instead.
		throw new InterruptedCohereStreamError('Cohere reached its output limit during a tool request. No pending tool request was run.', visible, singleTool, false, true);
	}
	if (!providerError && (!finishReason || !completeCalls || (calls.length > 0 && finishReason !== 'TOOL_CALL' && finishReason !== 'MAX_TOKENS'))) {
		throw new InterruptedCohereStreamError('Cohere returned an incomplete response. No pending tool request was run.', visible, singleTool);
	}
	// The native API validates every assistant call in history, even if the editor rejected its arguments.
	if (!providerError) {
		const ids = new Set<string>();
		const createTargets = new Set<string>();
		for (const call of calls) {
			if (ids.has(call.id)) { throw new InterruptedCohereStreamError('Cohere returned duplicate tool call IDs. No call from that batch was run.', visible); }
			ids.add(call.id);
			let args: unknown;
			try { args = JSON.parse(call.arguments); } catch { /* Invalid JSON is rejected below. */ }
			if (!args || typeof args !== 'object' || Array.isArray(args)) {
				throw new InvalidCohereToolArgumentsError(visible, singleTool);
			}
			if (call.name === 'reach_create_file') {
				const { workspaceFolder, path } = args as { workspaceFolder?: string; path?: string };
				if (typeof path === 'string') {
					const target = JSON.stringify([workspaceFolder, path]);
					if (createTargets.has(target)) { throw new InvalidCohereToolArgumentsError(visible, singleTool, 'Cohere requested multiple creations of the same file in one batch.'); }
					createTargets.add(target);
				}
			}
		}
	}
	return {
		text, toolPlan, toolCalls: calls, requestedTools: calls.length > 0 || finishReason === 'TOOL_CALL', providerError,
		// A batch whose calls all completed before the budget ran out is valid: the
		// truncation only cut what followed them, so the round proceeds normally.
		finishReason: (finishReason === 'TOOL_CALL' || (finishReason === 'MAX_TOKENS' && calls.length > 0)) ? 'tool_calls' : finishReason === 'COMPLETE' ? 'stop' : finishReason,
	};
}

interface CohereBufferedMessage {
	content?: { type?: string; text?: string }[];
	tool_plan?: string;
	tool_calls?: { id?: string; type?: string; function?: { name?: string; arguments?: string } }[];
}

async function readCohereBufferedResponse(response: Response, onDelta: (text: string) => void) {
	const body: { finish_reason?: string; error?: string; message?: CohereBufferedMessage } = await response.json();
	const message = body.message;
	const text = Array.isArray(message?.content) ? message.content.filter(part => part?.type === 'text' && typeof part.text === 'string').map(part => part.text).join('') : '';
	const wireCalls = message?.tool_calls;
	if (wireCalls !== undefined && wireCalls !== null && !Array.isArray(wireCalls)) { throw new InterruptedCohereStreamError('Cohere returned an invalid tool call list.', false); }
	const calls: ReachAssistantToolCall[] = [];
	for (const call of wireCalls ?? []) {
		if (typeof call?.id !== 'string' || !call.id || call.type !== 'function' || typeof call.function?.name !== 'string' || !call.function.name || typeof call.function.arguments !== 'string') { break; }
		calls.push({ id: call.id, name: call.function.name, arguments: call.function.arguments });
	}
	const result = finishCohereGeneration(text, message?.tool_plan ?? '', calls, body.finish_reason?.toUpperCase(), false, calls.length === (wireCalls?.length ?? 0), body.error);
	if (text) { onDelta(text); }
	return result;
}

export function isCohereMonthlyLimit(message: string): boolean {
	return /\b(?:per[- ]month|monthly) request limit\b/i.test(message);
}

/** Both timers belong to a single inference request, so editor approvals have no deadline. */
export async function fetchCohereStream(fetchResponse: (signal: AbortSignal) => Promise<Response>, signal: AbortSignal, onDelta: (text: string) => void, onStatus: (message: string) => void, buffered = false) {
	const idle = new AbortController();
	const networkSignal = AbortSignal.any([signal, idle.signal]);
	const watchdog = new RunOnceScheduler(() => idle.abort(new Error(localize('cohere.stream.idle', "Cohere sent no response events for 90 seconds. Retry this turn to continue."))), 90_000);
	let tool: string | undefined;
	const status = new RunOnceScheduler(() => {
		if (!networkSignal.aborted) {
			onStatus(tool ? localize('cohere.stream.preparing', "Cohere is preparing {0}.", tool) : localize('cohere.stream.waiting', "Waiting for Cohere to generate its response."));
		}
	}, 8_000);
	watchdog.schedule();
	status.schedule();
	try {
		let response = await fetchResponse(networkSignal);
		// Transient failures retry in place with the results of completed tools preserved: a
		// rate limit or a brief service outage must not kill a working turn. One shared budget
		// covers both kinds so a flapping service cannot loop the round forever.
		let transientRetries = 0;
		for (;;) {
			if (response.status === 429 && transientRetries < 4) {
				const detail = await response.text();
				const header = response.headers.get('Retry-After')?.trim();
				const waitMs = header ? (/^\d+(\.\d+)?$/.test(header) ? Number(header) * 1000 : Date.parse(header) - Date.now()) : /per.minute request limit/i.test(detail) ? 60_000 : NaN;
				if (isCohereMonthlyLimit(detail) || !(waitMs <= 60_000 && Number.isFinite(waitMs))) {
					return { text: '', toolCalls: undefined, requestedTools: false, finishReason: undefined, providerError: { code: 429, message: detail } };
				}
				transientRetries++;
				watchdog.cancel();
				status.cancel();
				const seconds = Math.max(1, Math.ceil(waitMs / 1000));
				onStatus(localize('cohere.stream.rateLimit', "Cohere reached its request limit. Retrying this model request in {0} seconds. Completed tool results are preserved.", seconds));
				const wait = new DisposableStore();
				try { await timeout(seconds * 1000, CancellationToken.fromAbortSignal(networkSignal, wait)); } finally { wait.dispose(); }
				watchdog.schedule();
				status.schedule();
				response = await fetchResponse(networkSignal);
				continue;
			}
			if (response.status >= 500 && response.status < 600 && transientRetries < 4) {
				transientRetries++;
				watchdog.cancel();
				status.cancel();
				const seconds = Math.min(2 * 2 ** (transientRetries - 1), 8);
				onStatus(localize('cohere.stream.transientOutage', "The Cohere service is temporarily unavailable (HTTP {0}). Retrying in {1} seconds. Completed tool results are preserved.", response.status, seconds));
				response.body?.cancel().catch(() => { });
				const wait = new DisposableStore();
				try { await timeout(seconds * 1000, CancellationToken.fromAbortSignal(networkSignal, wait)); } finally { wait.dispose(); }
				watchdog.schedule();
				status.schedule();
				response = await fetchResponse(networkSignal);
				continue;
			}
			break;
		}
		if (!response.ok || !response.body) {
			return { text: '', toolCalls: undefined, requestedTools: false, finishReason: undefined, providerError: { code: response.status, message: await response.text() } };
		}
		if (buffered) { return await readCohereBufferedResponse(response, onDelta); }
		return await readCohereStream(response.body, text => { status.cancel(); onDelta(text); }, name => {
			watchdog.schedule();
			if (name && tool !== name) { tool = name; if (!status.isScheduled()) { status.schedule(); } }
		});
	} catch (error) {
		networkSignal.throwIfAborted();
		throw error;
	} finally {
		status.dispose();
		watchdog.dispose();
	}
}
