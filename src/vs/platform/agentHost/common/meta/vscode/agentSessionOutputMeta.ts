/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

/** The deliverable a user picked in the chat output picker, recorded on the outgoing message. */
export type SessionOutputKind =
	| 'auto'
	| 'website'
	| 'mobile'
	| 'design'
	| 'slides'
	| 'animation'
	| 'dataVisualization'
	| 'image'
	| 'document'
	| 'spreadsheet'
	| 'mission';

const SESSION_OUTPUT_KINDS = new Set<SessionOutputKind>(['auto', 'website', 'mobile', 'design', 'slides', 'animation', 'dataVisualization', 'image', 'document', 'spreadsheet', 'mission']);

export const SESSION_OUTPUT_TYPE_META_KEY = 'vscode.sessionOutputType';

export function isSessionOutputKind(value: unknown): value is SessionOutputKind {
	return typeof value === 'string' && SESSION_OUTPUT_KINDS.has(value as SessionOutputKind);
}

/** The explicit output choice recorded in message or request metadata, or `undefined`. */
export function getSessionOutputKindFromMetadata(metadata: Record<string, unknown> | undefined): SessionOutputKind | undefined {
	const kind = metadata?.[SESSION_OUTPUT_TYPE_META_KEY];
	return isSessionOutputKind(kind) ? kind : undefined;
}
