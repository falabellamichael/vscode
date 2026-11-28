/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { CancellationToken } from '../../../../base/common/cancellation.js';
import { CancellationError } from '../../../../base/common/errors.js';
import { escapeUntrustedMarkdown } from '../../../../base/common/htmlContent.js';
import { Lazy } from '../../../../base/common/lazy.js';
import { createTrustedTypesPolicy } from '../../../../base/browser/trustedTypes.js';
import { localize } from '../../../../nls.js';
import type { AssistantNewsFeed, AssistantNewsScope } from '../../../../platform/native/common/native.js';

export const ASSISTANT_NEWS_TOOL = 'assistant.get_news_headlines';

const HOUR_MS = 60 * 60 * 1000;
const AI_FRESHNESS_HOURS = 7 * 24;
const MAX_HEADLINES = 6;
// Chromium requires TrustedHTML even when DOMParser is creating a detached XML document.
// This policy is private to RSS parsing. Only text, timestamps, and HTTPS links leave the parser.
const rssPolicy = new Lazy(() => createTrustedTypesPolicy('assistantNewsXml', { createHTML: value => value }));

interface Headline {
	readonly title: string;
	readonly url: string;
	readonly publisher: string;
	readonly publishedAt: number;
}

interface HeadlineResult {
	readonly headlines: readonly (Omit<Headline, 'publishedAt'> & { readonly publishedAt: string })[];
	readonly sources: readonly { readonly publisher: string; readonly status: string }[];
}

export function formatAssistantAINews(result: string, succeeded: boolean): string {
	const heading = `## ${localize('assistantNews.aiHeading', "AI news")}`;
	if (!succeeded) {
		return `${heading}\n\n${escapeUntrustedMarkdown(result)}`;
	}
	const { headlines, sources } = JSON.parse(result) as HeadlineResult;
	const items = headlines.slice(0, 3).map(({ title, url, publisher, publishedAt }) =>
		`- [${escapeUntrustedMarkdown(title)}](<${url}>) — ${escapeUntrustedMarkdown(publisher)}, ${publishedAt.slice(0, 10)}`);
	const gaps = sources.filter(source => source.status !== 'Available')
		.map(({ publisher, status }) => `${escapeUntrustedMarkdown(publisher)} (${escapeUntrustedMarkdown(status)})`);
	return [heading, '', ...items, ...(gaps.length ? ['', localize('assistantNews.coverageGaps', "Coverage gaps — {0}", gaps.join(', '))] : [])].join('\n');
}

function child(element: Element, name: string): Element | undefined {
	return Array.from(element.children).find(node => node.tagName === name);
}

function readHeadlines(feed: AssistantNewsFeed, now: number, freshnessHours: number): Headline[] | undefined {
	if (!feed.xml || /<!DOCTYPE\s/i.test(feed.xml)) { return undefined; }
	const xml = rssPolicy.value?.createHTML(feed.xml) ?? feed.xml;
	const feedDocument = new DOMParser().parseFromString(xml as string, 'application/xml');
	const root = feedDocument.firstElementChild;
	const channel = root?.tagName === 'rss' && child(root, 'channel');
	if (!channel) { return undefined; }
	const elements = feedDocument.createTreeWalker(root, NodeFilter.SHOW_ELEMENT);
	for (let node = elements.nextNode(); node; node = elements.nextNode()) {
		if (node.nodeName === 'parsererror') { return undefined; }
	}
	const headlines: Headline[] = [];
	for (const item of channel.children) {
		if (item.tagName !== 'item') { continue; }
		const title = child(item, 'title')?.textContent?.replace(/\s+/g, ' ').trim().slice(0, 240);
		const link = child(item, 'link')?.textContent?.trim();
		const publishedAt = Date.parse(child(item, 'pubDate')?.textContent ?? '');
		if (!title || !link || link.length > 512 || !Number.isFinite(publishedAt) || publishedAt < now - freshnessHours * HOUR_MS || publishedAt > now) { continue; }
		let url: URL;
		try { url = new URL(link); } catch { continue; }
		if (url.protocol !== 'https:' || url.username || url.password) { continue; }
		url.hash = '';
		headlines.push({ title, url: url.href, publisher: feed.publisher, publishedAt });
	}
	return headlines;
}

export async function getAssistantNewsHeadlines(
	readFeeds: (scope: AssistantNewsScope) => Promise<readonly AssistantNewsFeed[]>,
	scope: unknown,
	token: CancellationToken,
): Promise<string> {
	if (scope === undefined) { scope = 'canada'; }
	if (scope !== 'toronto' && scope !== 'canada' && scope !== 'world' && scope !== 'ai') { throw new Error('Choose Toronto, Canada, world, or AI headlines.'); }
	if (token.isCancellationRequested) { throw new CancellationError(); }
	const feeds = await readFeeds(scope);
	if (token.isCancellationRequested) { throw new CancellationError(); }
	const now = Date.now();
	const freshnessHours = scope === 'ai' ? AI_FRESHNESS_HOURS : 48;
	const freshnessLabel = scope === 'ai' ? '7 days' : '48 hours';
	const candidates: Headline[] = [];
	const sources = feeds.map(feed => {
		const headlines = readHeadlines(feed, now, freshnessHours);
		if (headlines?.length) { candidates.push(...headlines); }
		return { publisher: feed.publisher, status: feed.error ?? (headlines ? (headlines.length ? 'Available' : `No headlines in the last ${freshnessLabel}`) : 'Invalid RSS response') };
	});
	if (!candidates.length) {
		throw new Error(`No current ${scope} headlines. ${sources.map(source => `${source.publisher}: ${source.status}`).join('. ')}`);
	}
	const seen = new Set<string>();
	const publishers = new Set<string>();
	const headlines: Headline[] = [];
	const extra: Headline[] = [];
	for (const headline of candidates.sort((a, b) => b.publishedAt - a.publishedAt)) {
		const title = headline.title.toLowerCase();
		if (seen.has(headline.url) || seen.has(title)) { continue; }
		seen.add(headline.url);
		seen.add(title);
		// Give each AI publisher a slot before filling the list from a busy feed.
		(scope === 'ai' && publishers.has(headline.publisher) ? extra : headlines).push(headline);
		publishers.add(headline.publisher);
		if (headlines.length === MAX_HEADLINES) { break; }
	}
	return JSON.stringify({ scope, headlines: [...headlines, ...extra].slice(0, MAX_HEADLINES).map(headline => ({ ...headline, publishedAt: new Date(headline.publishedAt).toISOString() })), sources });
}
