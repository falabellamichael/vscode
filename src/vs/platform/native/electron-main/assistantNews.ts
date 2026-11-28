/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import type { AssistantNewsFeed, AssistantNewsScope } from '../common/native.js';

const sources = {
	toronto: [['CBC News', 'https://www.cbc.ca/webfeed/rss/rss-canada-toronto'], ['Global News', 'https://globalnews.ca/toronto/feed/']],
	canada: [['CBC News', 'https://www.cbc.ca/webfeed/rss/rss-canada'], ['Global News', 'https://globalnews.ca/canada/feed/']],
	world: [['CBC News', 'https://www.cbc.ca/webfeed/rss/rss-world'], ['Global News', 'https://globalnews.ca/world/feed/']],
	ai: [
		['OpenAI', 'https://openai.com/news/rss.xml'],
		['Google DeepMind', 'https://deepmind.google/blog/rss.xml'],
		['Hugging Face', 'https://huggingface.co/blog/feed.xml'],
	],
} satisfies Record<AssistantNewsScope, readonly (readonly [string, string])[]>;
const MAX_FEED_BYTES = 512 * 1024;
// OpenAI publishes its full archive in the feed, which currently exceeds 512 KiB.
const MAX_OPENAI_FEED_BYTES = 2 * 1024 * 1024;

export async function fetchAssistantNewsFeeds(scope: AssistantNewsScope, fetcher: (url: string, options: RequestInit) => Promise<Response>): Promise<readonly AssistantNewsFeed[]> {
	if (!Object.hasOwn(sources, scope)) { throw new Error('Unknown headline scope.'); }
	return Promise.all(sources[scope].map(async ([publisher, url]): Promise<AssistantNewsFeed> => {
		const maxBytes = publisher === 'OpenAI' ? MAX_OPENAI_FEED_BYTES : MAX_FEED_BYTES;
		try {
			const response = await fetcher(url, {
				headers: { Accept: 'application/rss+xml, application/xml, text/xml' },
				credentials: 'omit',
				redirect: 'error',
				signal: AbortSignal.timeout(8_000),
			});
			if (!response.ok || !response.body) {
				await response.body?.cancel();
				return { publisher, error: response.ok ? 'Empty response' : `HTTP ${response.status}` };
			}
			const reader = response.body.getReader();
			const decoder = new TextDecoder();
			let size = 0;
			let xml = '';
			try {
				for (;;) {
					const { done, value } = await reader.read();
					if (done) { break; }
					size += value.byteLength;
					if (size > maxBytes) {
						await reader.cancel();
						return { publisher, error: 'Feed exceeds the size limit' };
					}
					xml += decoder.decode(value, { stream: true });
				}
				return { publisher, xml: xml + decoder.decode() };
			} finally {
				reader.releaseLock();
			}
		} catch {
			return { publisher, error: 'Feed request failed or timed out' };
		}
	}));
}
