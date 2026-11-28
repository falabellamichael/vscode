/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { isHighSurrogate, isLowSurrogate } from '../../../../base/common/strings.js';
import type { IAssistantSavedIdea } from './assistantConversation.js';

export const ASSISTANT_READ_SAVED_IDEAS_TOOL = 'assistant.read_saved_ideas';
const ideasPerPage = 5;
const charactersPerPage = 2_000;

function ideaTextOffset(text: string, offset: number): number {
	return isHighSurrogate(text.charCodeAt(offset - 1)) && isLowSurrogate(text.charCodeAt(offset)) ? offset - 1 : offset;
}

function savedIdeaSourceTitle(idea: IAssistantSavedIdea): string {
	return idea.sourceTitle.slice(0, ideaTextOffset(idea.sourceTitle, 100));
}

/** Match direct, standalone requests for the user's collection. An empty query lists recent ideas. */
export function readSavedIdeaQuery(request: string): string | undefined {
	const text = request.trim().replace(/[.?!]+$/, '').replace(/^(?:(?:could|can|would) you\s+(?:please\s+)?|please\s+)/i, '');
	const direct = /^(?:(?:show(?: me)?|list|find|search|recall|what (?:is|was|are|were)) my (?:saved|bookmarked) ideas?|(?:show me the|list the) ideas? I (?:saved|bookmarked)|what ideas? (?:have I (?:saved|bookmarked)|did I (?:save|bookmark)))(?: (?:about|on|for|from|matching) ([^\r\n?!]{1,120}))?$/i;
	// Separate the topic from a bounded creative request in "using my saved ideas about X, suggest ..." or "use my saved ideas about X to draft ...".
	const creative = /^(?:use|using|based on) my (?:saved|bookmarked) ideas? (?:about|on|for) ([^,\r\n?!]{1,120}?)(?:,\s*|\s+to\s+)(?:please\s+)?(?:suggest|brainstorm|compare|summari[sz]e|draft|sketch|outline|develop|propose)\b/i;
	const topic = direct.exec(text) ?? creative.exec(text);
	return topic ? (topic[1] ?? '').trim().replace(/^(?:the|a|an)\s+/i, '') : undefined;
}

/** Only ideas matching the user's original query are readable, including subsequent pages. */
export function readSavedIdeaPage(ideas: readonly IAssistantSavedIdea[], query: string, args: Record<string, unknown>, consultedSources?: Set<string>): string {
	const page = args.page ?? 1;
	if (typeof page !== 'number' || !Number.isSafeInteger(page) || page < 1) {
		return 'Use a positive integer page number.';
	}
	if (args.id !== undefined) {
		const idea = ideas.find(idea => idea.id === args.id);
		if (!idea) { return 'No matching saved idea has that ID.'; }
		const pages = Math.ceil(idea.text.length / charactersPerPage);
		if (page > pages) { return 'Page exceeds the length of this idea.'; }
		const start = ideaTextOffset(idea.text, (page - 1) * charactersPerPage);
		const end = ideaTextOffset(idea.text, page * charactersPerPage);
		const title = savedIdeaSourceTitle(idea);
		consultedSources?.add(title);
		return `Saved idea ${idea.id} from ${JSON.stringify(title)}, page ${page} of ${pages}:\n${idea.text.slice(start, end)}`;
	}
	const pages = Math.ceil(ideas.length / ideasPerPage);
	if (!pages) { return 'No matching saved ideas.'; }
	if (page > pages) { return 'Page exceeds the number of matches.'; }
	const terms = query.toLowerCase().split(/\s+/).filter(Boolean);
	const pageIdeas = ideas.slice((page - 1) * ideasPerPage, page * ideasPerPage);
	const matches = pageIdeas.map(idea => {
		const body = idea.text.toLowerCase();
		const positions = terms.map(term => body.indexOf(term)).filter(index => index >= 0);
		const start = ideaTextOffset(idea.text, positions.length ? Math.max(0, Math.min(...positions) - 70) : 0);
		const end = ideaTextOffset(idea.text, Math.min(idea.text.length, start + 360));
		const kind = start === 0 && end === idea.text.length ? 'complete' : 'excerpt';
		return `${idea.id} · ${JSON.stringify(savedIdeaSourceTitle(idea))} · ${new Date(idea.createdAt).toISOString()} · ${kind}\n${idea.text.slice(start, end)}`;
	});
	for (const idea of pageIdeas) { consultedSources?.add(savedIdeaSourceTitle(idea)); }
	return `${ideas.length} matching saved ideas, page ${page} of ${pages}\n\n${matches.join('\n\n')}`;
}
