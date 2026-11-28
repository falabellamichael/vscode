/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { IChatRequestVariableEntry } from './attachments/chatVariableEntries.js';
import { isSessionOutputKind, SESSION_OUTPUT_TYPE_META_KEY, type SessionOutputKind } from '../../../../platform/agentHost/common/meta/agentSessionOutputMeta.js';

export type { SessionOutputKind } from '../../../../platform/agentHost/common/meta/agentSessionOutputMeta.js';
export { SESSION_OUTPUT_TYPE_META_KEY, isSessionOutputKind, getSessionOutputKindFromMetadata } from '../../../../platform/agentHost/common/meta/agentSessionOutputMeta.js';

export const SESSION_OUTPUT_CONTEXT_ID_PREFIX = 'session-output:';

/** The explicit output choice of hidden send context, or `undefined` for Auto-Select/no context. */
export function getSessionOutputKindFromSendContext(sendContext: readonly IChatRequestVariableEntry[] | undefined): SessionOutputKind | undefined {
	const kind = sendContext?.find(entry => entry.kind === 'promptText' && entry.id.startsWith(SESSION_OUTPUT_CONTEXT_ID_PREFIX))?.id.slice(SESSION_OUTPUT_CONTEXT_ID_PREFIX.length);
	return isSessionOutputKind(kind) ? kind : undefined;
}

/** The output choice of hidden send context, defaulting older drafts to Auto-Select. */
export function readSessionOutputTypeFromSendContext(sendContext: readonly IChatRequestVariableEntry[] | undefined): SessionOutputKind {
	return getSessionOutputKindFromSendContext(sendContext) ?? 'auto';
}

/** Request metadata for an explicit output choice, preserving unrelated metadata. */
export function withSessionOutputTypeMetadata(metadata: Record<string, unknown> | undefined, sendContext: readonly IChatRequestVariableEntry[] | undefined): Record<string, unknown> | undefined {
	const kind = getSessionOutputKindFromSendContext(sendContext);
	const result = { ...metadata };
	if (kind) {
		result[SESSION_OUTPUT_TYPE_META_KEY] = kind;
	} else {
		delete result[SESSION_OUTPUT_TYPE_META_KEY];
	}
	return Object.keys(result).length ? result : undefined;
}
