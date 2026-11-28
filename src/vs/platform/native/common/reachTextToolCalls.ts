/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import type { ReachAssistantToolSpec } from './native.js';

const TOOL_CALL_BLOCK = /```tool_call\s*\r?\n([\s\S]*?)```/g;
const XML_TOOL_CALL_BLOCK = /<tool_call>\s*([\s\S]*?)<\/tool_call>/g;
const XML_FUNCTION_BLOCK = /<function=([^<>\s]+)>([\s\S]*?)<\/function>/g;
const XML_PARAMETER_BLOCK = /<parameter=([^<>\s]+)>([\s\S]*?)<\/parameter>/g;

export interface ITextToolCall {
	readonly name: string;
	readonly args: Record<string, unknown>;
}

/** Local models often stringify scalars; coerce them back using the tool's JSON schema. */
export function coerceAssistantToolArguments(args: Record<string, unknown>, spec: ReachAssistantToolSpec | undefined): Record<string, unknown> {
	const properties = (spec?.parameters as { properties?: Record<string, { type?: string }> } | undefined)?.properties;
	let changed = false;
	const next: Record<string, unknown> = { ...args };
	for (const [key, value] of Object.entries(args)) {
		if (typeof value !== 'string') { continue; }
		const trimmed = value.trim();
		const type = properties?.[key]?.type;
		if (type === 'boolean') {
			if (trimmed === 'true' || trimmed === 'false') { next[key] = trimmed === 'true'; changed = true; }
		} else if (type === 'number' || type === 'integer') {
			const number = Number(trimmed);
			if (trimmed !== '' && Number.isFinite(number)) { next[key] = number; changed = true; }
		} else if (type === 'object' || type === 'array' || (!type && (trimmed.startsWith('{') || trimmed.startsWith('[')))) {
			try { next[key] = JSON.parse(trimmed); changed = true; } catch { /* Keep the original string. */ }
		}
	}
	return changed ? next : args;
}

function parseXmlToolArguments(body: string, spec: ReachAssistantToolSpec | undefined): Record<string, unknown> | undefined {
	const trimmed = body.trim();
	if (!trimmed) { return {}; }
	if (trimmed.startsWith('{')) {
		try {
			const parsed: unknown = JSON.parse(trimmed);
			if (typeof parsed === 'object' && parsed !== null && !Array.isArray(parsed)) { return parsed as Record<string, unknown>; }
		} catch { /* Fall back to parameter tags below. */ }
	}
	const args: Record<string, unknown> = {};
	let found = false;
	for (const match of trimmed.matchAll(XML_PARAMETER_BLOCK)) {
		found = true;
		args[match[1]] = match[2].trim();
	}
	return found ? coerceAssistantToolArguments(args, spec) : undefined;
}

/** Parses the fenced fallback protocol and the XML tool calls local models emit when native function calls are unavailable. */
export function parseTextToolCalls(text: string, specs: readonly ReachAssistantToolSpec[] = []): { calls: ITextToolCall[]; errors: string[]; display: string } {
	const calls: ITextToolCall[] = [];
	const errors: string[] = [];
	let display = text.replace(TOOL_CALL_BLOCK, (_block, json: string) => {
		try {
			const parsed = JSON.parse(json) as { name?: unknown; arguments?: unknown; args?: unknown };
			const args = parsed.arguments ?? parsed.args ?? {};
			if (typeof parsed.name !== 'string' || !parsed.name) {
				errors.push('a tool_call block is missing the "name" field');
			} else if (typeof args !== 'object' || args === null || Array.isArray(args)) {
				errors.push(`tool_call for "${parsed.name}" must pass "arguments" as an object`);
			} else {
				calls.push({ name: parsed.name, args: args as Record<string, unknown> });
			}
		} catch {
			errors.push('a tool_call block contained invalid JSON');
		}
		return '';
	});
	display = display.replace(XML_TOOL_CALL_BLOCK, (_block, inner: string) => {
		let matched = false;
		for (const fn of inner.matchAll(XML_FUNCTION_BLOCK)) {
			matched = true;
			const name = fn[1].trim();
			const args = parseXmlToolArguments(fn[2], specs.find(spec => spec.name === name));
			if (!args) {
				errors.push(`tool_call for "${name}" must pass arguments as JSON or <parameter=...> tags`);
			} else {
				calls.push({ name, args });
			}
		}
		if (!matched) { errors.push('a tool_call block contained no <function=...> call'); }
		return '';
	}).trim();
	return { calls, errors, display };
}
