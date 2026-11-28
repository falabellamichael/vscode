/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import assert from 'assert';
import { ensureNoDisposablesAreLeakedInTestSuite } from '../../../../base/test/common/utils.js';
import { REACH_ACCOUNT_ORIGIN, type ReachAssistantChatMessage } from '../../common/native.js';
import { fetchAssistantNewsFeeds } from '../../electron-main/assistantNews.js';
import { fetchAssistantWeatherForecast } from '../../electron-main/assistantWeather.js';
import { requestReachAssistant } from '../../common/reachChatTransport.js';
import { reachSubscriptionGuide } from '../../common/reachSubscriptionGuide.js';
import { ReachAssistantTurns } from '../../electron-main/reachAssistantTurns.js';

suite('Assistant news feeds', () => {
	ensureNoDisposablesAreLeakedInTestSuite();

	test('keeps AI sources separate and accommodates the larger OpenAI feed', async () => {
		const urls: string[] = [];
		const largeFeed = '<rss>' + ' '.repeat(900_000) + '</rss>';
		const feeds = await fetchAssistantNewsFeeds('ai', async (url, options) => {
			urls.push(url);
			assert.strictEqual(options.redirect, 'error');
			assert.strictEqual(options.credentials, 'omit');
			return new Response(url.includes('openai.com') ? largeFeed : '<rss/>');
		});
		assert.deepStrictEqual(urls, ['https://openai.com/news/rss.xml', 'https://deepmind.google/blog/rss.xml', 'https://huggingface.co/blog/feed.xml']);
		assert.deepStrictEqual(feeds.map(feed => feed.publisher), ['OpenAI', 'Google DeepMind', 'Hugging Face']);
		assert.strictEqual(feeds[0].xml, largeFeed);
	});
});

suite('The Weather Network forecast', () => {
	ensureNoDisposablesAreLeakedInTestSuite();

	const key = 'synthetic-weather-fixture';
	const vancouver = { city: 'Vancouver', subdivision: 'bc', country: 'ca' };

	test('uses the requested city and its provider-local day and sends the key only as a Pelmorex header', async () => {
		const result = await fetchAssistantWeatherForecast('tomorrow', vancouver, key, async (url, options) => {
			assert.strictEqual(url, 'https://api.data.pelmorex.com/weather/LongTermForecasts/CA/BC/Vancouver?locale=en-CA');
			assert.strictEqual(new Headers(options.headers).get('x-api-key'), key);
			assert.strictEqual(options.redirect, 'error');
			assert.strictEqual(options.credentials, 'omit');
			return Response.json({
				Location: { Name: 'Vancouver', SubDivision: 'British Columbia', SubDivisionCode: 'BC' },
				LongTermPeriod: [
					{ Period: '0', TimestampLocal: '2026-11-18T00:00:00', TemperatureCMax: '-5' },
					{ Period: '1', TimestampLocal: '2026-11-19T00:00:00', TemperatureCMax: '8', TemperatureCMin: '-2', POPPercentDay: '0', FxConditionDay: 'Sunny' },
				],
			});
		});
		assert.deepStrictEqual(result, {
			source: 'The Weather Network', sourceUrl: 'https://www.theweathernetwork.com/en/city/ca/british-columbia/vancouver/7-days',
			location: 'Vancouver, BC, CA', date: '2026-11-19', highC: 8, lowC: -2, precipitationChancePercent: 0, conditions: 'Sunny',
		});
	});

	test('supports the country and city route with no subdivision and encodes the city name', async () => {
		const result = await fetchAssistantWeatherForecast('today', { city: 'Aix en Provence', country: 'FR' }, key, async url => {
			assert.strictEqual(url, 'https://api.data.pelmorex.com/weather/LongTermForecasts/FR/Aix%20en%20Provence?locale=en-CA');
			return Response.json({ LongTermPeriod: [{ Period: 0, TimestampLocal: '2026-11-20T00:00:00', TemperatureCMax: '13' }] });
		});
		assert.strictEqual(result.date, '2026-11-20');
		assert.strictEqual(result.location, 'Aix en Provence, FR');
		assert.strictEqual(result.sourceUrl, 'https://www.theweathernetwork.com/en/city');
	});

	test('reports missing dates and rejected keys without returning raw provider content', async () => {
		await assert.rejects(() => fetchAssistantWeatherForecast('tomorrow', vancouver, key, async () =>
			Response.json({ LongTermPeriod: [{ Period: '0', TimestampLocal: '2026-11-18T00:00:00', TemperatureCMax: '7' }] })), /requested forecast day for Vancouver/);
		await assert.rejects(() => fetchAssistantWeatherForecast('today', vancouver, key, async () =>
			Response.json({ detail: key }, { status: 403 })),
		(error: Error) => error.message.includes('/weather') && !error.message.includes(key));
		await assert.rejects(() => fetchAssistantWeatherForecast('today', { city: '../test', country: 'FR' }, key, async () => {
			assert.fail('Invalid city reached the provider');
		}), /Enter a city/);
	});
});

suite('Open-Meteo public weather fallback', () => {
	ensureNoDisposablesAreLeakedInTestSuite();

	const toronto = { city: 'Toronto', country: 'CA', subdivision: 'ON' };
	const geocodingUrl = 'https://geocoding-api.open-meteo.com/v1/search?name=Toronto&count=10&language=en&format=json';
	const forecastUrl = 'https://api.open-meteo.com/v1/forecast?latitude=43.6532&longitude=-79.3832&daily=temperature_2m_max,temperature_2m_min,precipitation_probability_max,weather_code&timezone=auto&forecast_days=2';
	const geoResult = { name: 'Toronto', country_code: 'CA', admin1: 'Ontario', latitude: 43.6532, longitude: -79.3832 };
	const daily = {
		time: ['2026-11-18', '2026-11-19'],
		temperature_2m_max: [4, 7],
		temperature_2m_min: [-1, 2],
		precipitation_probability_max: [20, 60],
		weather_code: [2, 61],
	};
	const openMeteoFetcher = async (url: string, options: RequestInit) => {
		assert.strictEqual(options.credentials, 'omit');
		assert.strictEqual(options.redirect, 'error');
		assert.strictEqual(new Headers(options.headers).get('x-api-key'), null);
		if (url === geocodingUrl) { return Response.json({ results: [geoResult] }); }
		if (url === forecastUrl) { return Response.json({ daily }); }
		assert.fail(`Unexpected fallback URL: ${url}`);
	};

	test('uses Open-Meteo without a key and maps the provider-local day', async () => {
		const result = await fetchAssistantWeatherForecast('tomorrow', toronto, '', openMeteoFetcher, true);
		assert.deepStrictEqual(result, {
			source: 'Open-Meteo',
			sourceUrl: 'https://open-meteo.com/en/docs#latitude=43.6532&longitude=-79.3832',
			location: 'Toronto, Ontario, CA',
			date: '2026-11-19', highC: 7, lowC: 2, precipitationChancePercent: 60, conditions: 'Slight rain',
		});
	});

	test('falls back when Pelmorex rejects the key, but only with the fallback enabled', async () => {
		let fallbackCalls = 0;
		const rejecting = async (url: string, options: RequestInit) => {
			if (url.startsWith('https://api.data.pelmorex.com')) { return Response.json({ detail: 'rejected' }, { status: 403 }); }
			fallbackCalls++;
			return openMeteoFetcher(url, options);
		};
		const result = await fetchAssistantWeatherForecast('today', toronto, 'synthetic-weather-fixture', rejecting, true);
		assert.strictEqual(result.source, 'Open-Meteo');
		assert.strictEqual(result.date, '2026-11-18');
		assert.strictEqual(fallbackCalls, 2);
		await assert.rejects(() => fetchAssistantWeatherForecast('today', toronto, 'synthetic-weather-fixture', rejecting, false), /Pelmorex rejected the API key/);
		await assert.rejects(() => fetchAssistantWeatherForecast('today', toronto, '', openMeteoFetcher, false), /valid Pelmorex API key/);
	});

	test('falls back when the Pelmorex request fails or times out', async () => {
		const result = await fetchAssistantWeatherForecast('today', toronto, 'synthetic-weather-fixture', async (url, options) => {
			if (url.startsWith('https://api.data.pelmorex.com')) { throw new Error('network down'); }
			return openMeteoFetcher(url, options);
		}, true);
		assert.strictEqual(result.source, 'Open-Meteo');
	});

	test('keeps The Weather Network primary when it succeeds', async () => {
		const result = await fetchAssistantWeatherForecast('today', toronto, 'synthetic-weather-fixture', async url => {
			if (url.startsWith('https://api.data.pelmorex.com')) {
				return Response.json({ Location: { Name: 'Toronto', SubDivisionCode: 'ON' }, LongTermPeriod: [{ Period: 0, TimestampLocal: '2026-11-18T00:00:00', TemperatureCMax: '5' }] });
			}
			assert.fail('The public fallback must not run while Pelmorex succeeds');
		}, true);
		assert.strictEqual(result.source, 'The Weather Network');
	});

	test('filters geocoding by country and prefers the requested province', async () => {
		const result = await fetchAssistantWeatherForecast('today', toronto, '', async (url, options) => {
			if (url === geocodingUrl) {
				return Response.json({ results: [
					{ name: 'Toronto', country_code: 'US', admin1: 'Ohio', latitude: 40.46, longitude: -80.6 },
					{ name: 'Toronto', country_code: 'CA', admin1: 'Saskatchewan', latitude: 52.0, longitude: -105.0 },
					geoResult,
				] });
			}
			return openMeteoFetcher(url, options);
		}, true);
		assert.strictEqual(result.location, 'Toronto, Ontario, CA');
	});

	test('reports an unknown city and a missing forecast day from the fallback', async () => {
		await assert.rejects(() => fetchAssistantWeatherForecast('today', toronto, '', async (url, options) => {
			if (url === geocodingUrl) { return Response.json({}); }
			return openMeteoFetcher(url, options);
		}, true), /Open-Meteo could not find Toronto, ON, CA/);
		await assert.rejects(() => fetchAssistantWeatherForecast('tomorrow', toronto, '', async (url, options) => {
			if (url === forecastUrl) { return Response.json({ daily: { ...daily, time: ['2026-11-18'] } }); }
			return openMeteoFetcher(url, options);
		}, true), /requested forecast day/);
	});
});

suite('REACH Assistant transport', () => {
	ensureNoDisposablesAreLeakedInTestSuite();

	const account = { id: 'account', access: { activeSubscription: true }, requestAllowance: { remaining: 2 }, allowedModels: [] };
	const requestId = '7b0fd1ed-357e-4117-b6f7-c05dcd891992';

	test('uses the authenticated catalog even when the account has an empty model list', async () => {
		const calls: { path: string; options: RequestInit }[] = [];
		const fetcher = async (url: string, options: RequestInit) => {
			assert.strictEqual(new URL(url).origin, REACH_ACCOUNT_ORIGIN);
			const path = new URL(url).pathname;
			calls.push({ path, options });
			switch (path) {
				case '/v1/account': return Response.json(account);
				case '/v1/models': return Response.json({ data: [
					{ id: 'embedding', access: 'requests', capabilities: { chat: false, endpoint: '/v1/embeddings' } },
					{ id: 'decision', access: 'requests', capabilities: { endpoint: '/v1/decisions' } },
					{ id: 'api-model', access: 'tokens' },
					{ id: 'qwen', access: 'requests', name: 'Qwen', capabilities: { outputTokenLimit: true } },
				] });
				case '/v1/chat/completions': return Response.json({ choices: [{ message: { content: 'Hello!' } }] });
				default: throw new Error('Unexpected request');
			}
		};
		const result = await requestReachAssistant(fetcher, account.id, 'test-token', { messages: [{ role: 'user', content: 'Hello' }] }, requestId, new AbortController().signal);
		assert.deepStrictEqual(result, { kind: 'reply', text: 'Hello!', model: 'Qwen' });
		assert.deepStrictEqual(calls.map(call => call.path), ['/v1/account', '/v1/models', '/v1/chat/completions']);
		const completion = calls[2].options;
		assert.strictEqual(new Headers(completion.headers).get('Idempotency-Key'), requestId);
		assert.strictEqual(completion.redirect, 'error');
		const payload = JSON.parse(String(completion.body));
		assert.strictEqual(payload.model, 'qwen');
		assert.deepStrictEqual(payload.messages.map((message: { role: string }) => message.role), ['system', 'user']);
		assert.deepStrictEqual(payload.messages.at(-1), { role: 'user', content: 'Hello' });
	});

	test('grounds regular REACH Subscription answers above earlier chat text without adding the guide to unrelated turns or Assistant', async () => {
		const systems: string[] = [];
		const fetcher = async (url: string, options: RequestInit) => {
			switch (new URL(url).pathname) {
				case '/v1/account': return Response.json(account);
				case '/v1/models': return Response.json({ data: [{ id: 'chat', access: 'requests' }] });
				case '/v1/chat/completions':
					systems.push(JSON.parse(String(options.body)).messages[0].content);
					return Response.json({ choices: [{ message: { content: 'answer' } }] });
				default: throw new Error('Unexpected request');
			}
		};
		const query = 'write a .md for vs code agents to understand how REACH Subscription works. call it Subscription.md';
		for (const [prompt, surface] of [[query, 'chat'], ['Explain a VS Code event subscription.', 'chat'], [query, 'assistant']] as const) {
			await requestReachAssistant(fetcher, account.id, 'test-token', { messages: [
				{ role: 'user', content: 'What is a subscription?' },
				{ role: 'assistant', content: 'It is an event stream.' },
				{ role: 'user', content: prompt },
			] }, requestId, new AbortController().signal, undefined, surface);
		}
		assert.deepStrictEqual(systems.map(system => system.includes(reachSubscriptionGuide)), [true, false, false]);
		for (const fact of ['SignalREACH account plan', 'GET /v1/account', 'GET /v1/models', 'POST /v1/chat/completions']) {
			assert.ok(systems[0].includes(fact), fact);
		}
	});

	test('forwards tool specs and parses tool_calls replies', async () => {
		const calls: { path: string; options: RequestInit }[] = [];
		const fetcher = async (url: string, options: RequestInit) => {
			const path = new URL(url).pathname;
			calls.push({ path, options });
			switch (path) {
				case '/v1/account': return Response.json(account);
				case '/v1/models': return Response.json({ data: [{ id: 'qwen', access: 'requests', name: 'Qwen' }] });
				case '/v1/chat/completions': return Response.json({ choices: [{ message: { content: null, tool_calls: [
					{ id: 'call-1', type: 'function', function: { name: 'mcp_mail_search', arguments: '{"q":"hi"}' } },
					{ id: 'call-2', type: 'function', function: { name: 'broken' } },
				] } }] });
				default: throw new Error('Unexpected request');
			}
		};
		const request = {
			messages: [{ role: 'system' as const, content: 'Extra instructions' }, { role: 'user' as const, content: 'Search mail' }],
			tools: [{ name: 'mcp_mail_search', description: 'Search mail', parameters: { type: 'object' } }],
			textToolIndex: 'Only needed by the text tool protocol.',
		};
		const result = await requestReachAssistant(fetcher, account.id, 'test-token', request, requestId, new AbortController().signal);
		assert.deepStrictEqual(result, {
			kind: 'reply', text: '', model: 'Qwen',
			toolCalls: [
				{ id: 'call-1', name: 'mcp_mail_search', arguments: '{"q":"hi"}' },
				{ id: 'call-2', name: 'broken', arguments: '' },
			],
		});
		const payload = JSON.parse(String(calls[2].options.body!));
		assert.strictEqual(payload.tool_choice, 'auto');
		assert.deepStrictEqual(payload.tools, [{ type: 'function', function: { name: 'mcp_mail_search', description: 'Search mail', parameters: { type: 'object' } } }]);
		assert.strictEqual(payload.messages[0].role, 'system');
		assert.match(payload.messages[0].content, /Extra instructions/);
		assert.ok(!payload.messages[0].content.includes(request.textToolIndex));
	});

	test('preserves text tool calls and their results on models that reject native tools', async () => {
		const system = 'For list_notes, emit a fenced tool_call and wait for the result.';
		const messages: ReachAssistantChatMessage[] = [{ role: 'system', content: system }, { role: 'user', content: 'List my notes.' }];
		const toolCall = '```tool_call\n{"name":"list_notes","arguments":{}}\n```';
		const toolResult = '<tool_result name="list_notes">\nOne meeting note.\n</tool_result>';
		const send = (step: number) => requestReachAssistant(async (url, options) => {
			switch (new URL(url).pathname) {
				case '/v1/account': return Response.json(account);
				case '/v1/models': return Response.json({ data: [{ id: 'text-only', access: 'requests', allow_tools: false }] });
				case '/v1/chat/completions': {
					const payload = JSON.parse(String(options.body));
					if (Object.hasOwn(payload, 'tools') || Object.hasOwn(payload, 'tool_choice')) { return new Response(null, { status: 400 }); }
					assert.ok(payload.messages[0].content.includes(system));
					assert.ok(payload.messages[0].content.endsWith('list_notes: List notes'));
					assert.deepStrictEqual(payload.messages.slice(1), messages.slice(1));
					return Response.json({ choices: [{ message: { content: step === 0 ? toolCall : 'You have one meeting note.' } }] });
				}
				default: throw new Error('Unexpected request');
			}
		}, account.id, 'test-token', {
			modelId: 'text-only', messages, tools: [{ name: 'list_notes', description: 'List notes', parameters: { type: 'object' } }], textToolIndex: 'list_notes: List notes',
		}, `${requestId}.${step}`, new AbortController().signal);

		assert.deepStrictEqual(await send(0), { kind: 'reply', text: toolCall, model: 'text-only' });
		messages.push({ role: 'assistant', content: toolCall }, { role: 'user', content: toolResult });
		assert.deepStrictEqual(await send(1), { kind: 'reply', text: 'You have one meeting note.', model: 'text-only' });
	});

	test('echoes prior tool calls and tool results back to the model', async () => {
		const calls: { path: string; options: RequestInit }[] = [];
		const fetcher = async (url: string, options: RequestInit) => {
			const path = new URL(url).pathname;
			calls.push({ path, options });
			switch (path) {
				case '/v1/account': return Response.json(account);
				case '/v1/models': return Response.json({ data: [{ id: 'qwen', access: 'requests' }] });
				case '/v1/chat/completions': return Response.json({ choices: [{ message: { content: 'Done.' } }] });
				default: throw new Error('Unexpected request');
			}
		};
		const request = {
			messages: [
				{ role: 'user' as const, content: 'Search mail' },
				{ role: 'assistant' as const, content: '', toolCalls: [{ id: 'call-1', name: 'mcp_mail_search', arguments: '{"q":"hi"}' }] },
				{ role: 'tool' as const, toolCallId: 'call-1', name: 'mcp_mail_search', content: 'found 2' },
			],
		};
		const result = await requestReachAssistant(fetcher, account.id, 'test-token', request, requestId, new AbortController().signal);
		assert.deepStrictEqual(result, { kind: 'reply', text: 'Done.', model: 'qwen' });
		const payload = JSON.parse(String(calls[2].options.body!));
		assert.deepStrictEqual(payload.messages.at(-3), { role: 'user', content: 'Search mail' });
		assert.deepStrictEqual(payload.messages.at(-2).tool_calls, [{ id: 'call-1', type: 'function', function: { name: 'mcp_mail_search', arguments: '{"q":"hi"}' } }]);
		assert.deepStrictEqual(payload.messages.at(-1), { role: 'tool', content: 'found 2', name: 'mcp_mail_search', tool_call_id: 'call-1' });
	});

	test('streams reply deltas and assembles streamed tool_calls', async () => {
		const events = [
			`data: ${JSON.stringify({ choices: [{ delta: { content: 'Hello ' } }] })}\n\n`,
			`data: ${JSON.stringify({ choices: [{ delta: { content: 'there' } }] })}\n\n`,
			`data: ${JSON.stringify({ choices: [{ delta: { tool_calls: [{ index: 0, id: 'call-1', function: { name: 'mcp_mail', arguments: '{"q":' } }] } }] })}\n\n`,
			`data: ${JSON.stringify({ choices: [{ delta: { tool_calls: [{ index: 0, function: { arguments: '"hi"}' } }] } }] })}\n\n`,
			'data: [DONE]\n\n',
		];
		const stream = new ReadableStream<Uint8Array>({
			start(controller) {
				for (const event of events) { controller.enqueue(new TextEncoder().encode(event)); }
				controller.close();
			},
		});
		const calls: { path: string; options: RequestInit }[] = [];
		const fetcher = async (url: string, options: RequestInit) => {
			const path = new URL(url).pathname;
			calls.push({ path, options });
			switch (path) {
				case '/v1/account': return Response.json(account);
				case '/v1/models': return Response.json({ data: [{ id: 'qwen', access: 'requests' }] });
				case '/v1/chat/completions': return new Response(stream, { headers: { 'Content-Type': 'text/event-stream' } });
				default: throw new Error('Unexpected request');
			}
		};
		const deltas: string[] = [];
		const result = await requestReachAssistant(fetcher, account.id, 'test-token', { messages: [{ role: 'user', content: 'Hi' }] }, requestId, new AbortController().signal, delta => deltas.push(delta));
		assert.deepStrictEqual(result, { kind: 'reply', text: 'Hello there', model: 'qwen', toolCalls: [{ id: 'call-1', name: 'mcp_mail', arguments: '{"q":"hi"}' }] });
		assert.deepStrictEqual(deltas, ['Hello ', 'there']);
		assert.strictEqual(JSON.parse(String(calls[2].options.body)).stream, true);
	});

	test('reads fragmented CRLF events and completes at the end marker', async () => {
		const wire = [
			': keep-alive',
			'data: {"choices":',
			'data: [{"delta":{"content":"Café"}}]}',
			'',
			'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"id":"call-1","function":{"name":"list_notes","arguments":"{}"}}]}}]}',
			'',
			'data: [DONE]', '', '',
		].join('\r\n');
		let cancelled = false;
		const stream = new ReadableStream<Uint8Array>({
			start(controller) {
				const bytes = new TextEncoder().encode(wire);
				for (let index = 0; index < bytes.length; index++) {
					controller.enqueue(bytes.subarray(index, index + 1));
				}
			},
			pull(controller) { controller.close(); },
			cancel() { cancelled = true; },
		}, { highWaterMark: 0 });
		const deltas: string[] = [];
		const result = await requestReachAssistant(async url => {
			switch (new URL(url).pathname) {
				case '/v1/account': return Response.json(account);
				case '/v1/models': return Response.json({ data: [{ id: 'qwen', access: 'requests' }] });
				case '/v1/chat/completions': return new Response(stream, { headers: { 'Content-Type': 'text/event-stream' } });
				default: throw new Error('Unexpected request');
			}
		}, account.id, 'test-token', { messages: [{ role: 'user', content: 'Find my notes about Café' }] }, requestId, new AbortController().signal, delta => deltas.push(delta));

		assert.deepStrictEqual({ result, deltas, cancelled }, {
			result: { kind: 'reply', text: 'Café', model: 'qwen', toolCalls: [{ id: 'call-1', name: 'list_notes', arguments: '{}' }] },
			deltas: ['Café'], cancelled: true,
		});
	});

	test('regular chat treats a tool-only completion as unanswered without streaming its marker', async () => {
		const complete = (response: Response, chunks: string[]) => requestReachAssistant(async url => {
			switch (new URL(url).pathname) {
				case '/v1/account': return Response.json(account);
				case '/v1/models': return Response.json({ data: [{ id: 'chat-model', access: 'requests' }] });
				case '/v1/chat/completions': return response;
				default: throw new Error('Unexpected request');
			}
		}, account.id, 'test-token', { messages: [{ role: 'user', content: 'Summarize the workspace' }] }, requestId, new AbortController().signal, delta => chunks.push(delta), 'chat');
		const stream = (parts: string[]) => new Response(parts.map(content =>
			'data: ' + JSON.stringify({ choices: [{ delta: { content } }] }) + '\n\n').join('') + 'data: [DONE]\n\n',
		{ headers: { 'Content-Type': 'text/event-stream' } });
		for (const response of [stream(['tool', '_calls']), Response.json({ choices: [{ message: { content: 'tool_calls' } }] }),
			Response.json({ choices: [{ finish_reason: 'tool_calls', message: { content: null, tool_calls: [{ id: 'call', function: { name: 'read' } }] } }] })]) {
			const deltas: string[] = [];
			const result = await complete(response, deltas);
			assert.deepStrictEqual({ result, deltas }, { result: { kind: 'unavailable', reason: 'toolOnly' }, deltas: [] });
		}
		const chunks: string[] = [];
		const result = await complete(stream(['tool', '_calls', ' is the field name.']), chunks);
		assert.deepStrictEqual([result, chunks.join('')], [{ kind: 'reply', text: 'tool_calls is the field name.', model: 'chat-model' }, 'tool_calls is the field name.']);
	});

	test('stops before submitting a prompt when cancelled during model lookup', async () => {
		const control = new AbortController();
		const calls: string[] = [];
		const result = await requestReachAssistant(async url => {
			const path = new URL(url).pathname;
			calls.push(path);
			if (path === '/v1/account') { return Response.json(account); }
			control.abort();
			return Response.json({ data: [{ id: 'qwen', access: 'requests' }] });
		}, account.id, 'test-token', { messages: [{ role: 'user', content: 'Hello' }] }, requestId, control.signal);
		assert.deepStrictEqual(result, { kind: 'stopped' });
		assert.deepStrictEqual(calls, ['/v1/account', '/v1/models']);
	});

	test('reports outputLimit for truncated completions on every surface and transport', async () => {
		for (const surface of ['assistant', 'chat'] as const) {
			for (const stream of [false, true]) {
				const fetcher = async (url: string, options: RequestInit) => {
					switch (new URL(url).pathname) {
						case '/v1/account': return Response.json(account);
						case '/v1/models': return Response.json({ data: [{ id: 'qwen', access: 'requests' }] });
						case '/v1/chat/completions': {
							const choice = { finish_reason: 'length', message: { content: 'Partial mock summary', tool_calls: [{ id: 'call-1', type: 'function', function: { name: 'read_file', arguments: '{}' } }] } };
							return JSON.parse(String(options.body)).stream
								? new Response(`data: ${JSON.stringify({ choices: [{ finish_reason: choice.finish_reason, delta: { content: choice.message.content, tool_calls: [{ index: 0, id: 'call-1', function: { name: 'read_file', arguments: '{}' } }] } }] })}\n\ndata: [DONE]\n\n`, { headers: { 'Content-Type': 'text/event-stream' } })
								: Response.json({ choices: [choice] });
						}
						default: throw new Error('Unexpected request');
					}
				};
				const result = await requestReachAssistant(fetcher, account.id, 'test-token',
					{ messages: [{ role: 'user', content: 'Summarize the mock note' }], tools: [{ name: 'read_file', description: 'Read', parameters: { type: 'object' } }] },
					requestId, new AbortController().signal, stream ? () => { } : undefined, surface);
				assert.deepStrictEqual(result, { kind: 'unavailable', reason: 'outputLimit', model: 'qwen' });
			}
		}
	});

	test('reuses the verified turn across rounds and releases it on finish or window close', async () => {
		const paths: string[] = [];
		const completions: { authorization: string | null; idempotency: string | null }[] = [];
		let unauthorizedAt = 0;
		const fetcher = async (url: string, options: RequestInit) => {
			const path = new URL(url).pathname;
			paths.push(path);
			switch (path) {
				case '/v1/account': return Response.json(account);
				case '/v1/models': return Response.json({ data: [{ id: 'qwen', access: 'requests' }] });
				case '/v1/chat/completions': {
					const headers = new Headers(options.headers);
					completions.push({ authorization: headers.get('Authorization'), idempotency: headers.get('Idempotency-Key') });
					if (completions.length === unauthorizedAt) { return new Response(null, { status: 401 }); }
					return Response.json({ choices: [{ message: { content: 'Round answer' } }] });
				}
				default: throw new Error('Unexpected request');
			}
		};
		const turns = new ReachAssistantTurns(fetcher);
		const request = { messages: [{ role: 'user' as const, content: 'Summarize the mock note' }] };
		const send = (requestId: string, turnId = 'turn-1', windowId: number | undefined = 1, token = 'test-token', conversationId = 'conversation') =>
			turns.send(windowId, account.id, token, request, requestId, conversationId, turnId, () => { });
		unauthorizedAt = 3;
		assert.deepStrictEqual(await send('r1'), { kind: 'reply', text: 'Round answer', model: 'qwen' });
		assert.deepStrictEqual(await send('r2'), { kind: 'reply', text: 'Round answer', model: 'qwen' });
		assert.deepStrictEqual(await send('r3'), { kind: 'signedOut' });
		assert.deepStrictEqual(paths, ['/v1/account', '/v1/models', '/v1/chat/completions', '/v1/chat/completions', '/v1/chat/completions']);
		assert.deepStrictEqual(completions.map(call => call.authorization), ['Bearer test-token', 'Bearer test-token', 'Bearer test-token']);
		assert.deepStrictEqual(completions.map(call => call.idempotency), ['r1', 'r2', 'r3']);
		const before = paths.length;
		assert.deepStrictEqual(await send('r4', 'turn-1', 1, 'other-token'), { kind: 'signedOut' });
		assert.deepStrictEqual(await send('r4b', 'turn-1', 1, 'test-token', 'other-conversation'), { kind: 'unavailable' });
		assert.strictEqual(paths.length, before);
		turns.finish(2, 'turn-1');
		assert.deepStrictEqual(await send('r5'), { kind: 'reply', text: 'Round answer', model: 'qwen' });
		assert.deepStrictEqual(paths.slice(before), ['/v1/chat/completions']);
		turns.finish(1, 'turn-1');
		assert.deepStrictEqual(await send('r6'), { kind: 'reply', text: 'Round answer', model: 'qwen' });
		assert.deepStrictEqual(paths.slice(-3), ['/v1/account', '/v1/models', '/v1/chat/completions']);
		assert.deepStrictEqual(await send('r7', 'turn-1', 2), { kind: 'reply', text: 'Round answer', model: 'qwen' });
		assert.deepStrictEqual(paths.slice(-3), ['/v1/account', '/v1/models', '/v1/chat/completions']);
		assert.deepStrictEqual(await send('r8', 'turn-2'), { kind: 'reply', text: 'Round answer', model: 'qwen' });
		assert.deepStrictEqual(paths.slice(-3), ['/v1/account', '/v1/models', '/v1/chat/completions']);
	});

	test('finish during a completion stops the round without leaving the conversation busy', async () => {
		const paths: string[] = [];
		let hangNext = false;
		let markStarted: (() => void) | undefined;
		const started = new Promise<void>(resolve => { markStarted = resolve; });
		const fetcher = async (url: string, options: RequestInit) => {
			const path = new URL(url).pathname;
			paths.push(path);
			switch (path) {
				case '/v1/account': return Response.json(account);
				case '/v1/models': return Response.json({ data: [{ id: 'qwen', access: 'requests' }] });
				case '/v1/chat/completions': {
					if (hangNext) {
						hangNext = false;
						markStarted?.();
						if (options.signal?.aborted) { throw new DOMException('Aborted', 'AbortError'); }
						await new Promise<void>((_resolve, reject) => {
							options.signal?.addEventListener('abort', () => reject(new DOMException('Aborted', 'AbortError')), { once: true });
						});
					}
					return Response.json({ choices: [{ message: { content: 'Round answer' } }] });
				}
				default: throw new Error('Unexpected request');
			}
		};
		const turns = new ReachAssistantTurns(fetcher);
		const request = { messages: [{ role: 'user' as const, content: 'Summarize the mock note' }] };
		assert.deepStrictEqual(await turns.send(1, account.id, 'test-token', request, 'r1', 'conversation', 'turn-1', () => { }), { kind: 'reply', text: 'Round answer', model: 'qwen' });
		hangNext = true;
		const pending = turns.send(1, account.id, 'test-token', request, 'r2', 'conversation', 'turn-1', () => { });
		await started;
		turns.finish(1, 'turn-1');
		assert.deepStrictEqual(await pending, { kind: 'stopped' });
		assert.deepStrictEqual(await turns.send(1, account.id, 'test-token', request, 'r3', 'conversation', 'turn-1', () => { }), { kind: 'reply', text: 'Round answer', model: 'qwen' });
		assert.deepStrictEqual(paths.slice(-3), ['/v1/account', '/v1/models', '/v1/chat/completions']);
	});

	test('classifies transient preflight failures separately from completion failures', async () => {
		const attempt = async (accountStatus: number, completionStatus: number) => {
			const paths: string[] = [];
			const result = await requestReachAssistant(async url => {
				const path = new URL(url).pathname;
				paths.push(path);
				switch (path) {
					case '/v1/account': return accountStatus === 200 ? Response.json(account) : new Response(null, { status: accountStatus });
					case '/v1/models': return Response.json({ data: [{ id: 'qwen', access: 'requests' }] });
					case '/v1/chat/completions': return new Response(null, { status: completionStatus });
					default: throw new Error('Unexpected request');
				}
			}, account.id, 'test-token', { messages: [{ role: 'user', content: 'Summarize the mock note' }] }, requestId, new AbortController().signal);
			return { result, paths };
		};
		assert.deepStrictEqual(await attempt(503, 200), { result: { kind: 'unavailable', reason: 'preflightUnavailable' }, paths: ['/v1/account'] });
		assert.deepStrictEqual(await attempt(404, 200).then(value => value.result), { kind: 'unavailable' });
		assert.deepStrictEqual(await attempt(200, 503), { result: { kind: 'unavailable' }, paths: ['/v1/account', '/v1/models', '/v1/chat/completions'] });
	});
});
