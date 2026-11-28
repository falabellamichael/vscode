/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import assert from 'assert';
import sinon from 'sinon';
import { VSBuffer } from '../../../../base/common/buffer.js';
import { hashAsync } from '../../../../base/common/hash.js';
import { URI } from '../../../../base/common/uri.js';
import { ensureNoDisposablesAreLeakedInTestSuite } from '../../../../base/test/common/utils.js';
import type { IFileService } from '../../../files/common/files.js';
import { resolveAgentChatContext } from '../../common/agent.js';
import { ActionType } from '../../common/state/sessionActions.js';
import { MessageAttachmentKind, MessageKind, ResponsePartKind, ToolResultContentType, TurnState } from '../../common/state/sessionState.js';
import { CohereChatAgent } from '../../node/cohere/cohereChatAgent.js';
import { fetchCohereStream, toCohereMessage } from '../../node/cohere/cohereChatStream.js';
import { ReachWebVerification } from '../../node/reach/reachWebVerification.js';

suite('Cohere regular chat provider', () => {
	const store = ensureNoDisposablesAreLeakedInTestSuite();
	const chat = URI.parse('ahp-chat:/cohere-tools');
	const session = URI.parse('cohere:/fixture');
	const root = URI.file('/project');
	const prompt = 'summarize the workspace please';
	const tool = { name: 'workspace.read', inputSchema: { type: 'object', properties: { path: { type: 'string' } } } } as const;

	function stream(...choices: unknown[]): Response {
		return new Response(choices.map(choice => 'data: ' + JSON.stringify({ choices: [choice] }) + '\n\n').join('') + 'data: [DONE]\n\n');
	}

	async function fixture(completion: (init: RequestInit) => Response, workingDirectories: URI[] = [root], model?: string, files: Pick<IFileService, 'readFile'> = { readFile: () => { throw new Error('unexpected attachment'); } }) {
		const agent = store.add(new CohereChatAgent(files, async (url, init) => {
			if (String(url).endsWith('/models')) { return Response.json({ data: [] }); }
			assert.strictEqual(String(url), model ? 'https://api.cohere.com/v2/chat' : 'https://api.cohere.com/compatibility/v1/chat/completions');
			return completion(init!);
		}));
		await agent.authenticate('https://api.cohere.com', 'fixture-key');
		await agent.chats.createChat(chat, session, { workingDirectories, model: model ? { id: model } : undefined });
		agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = [tool];
		return agent;
	}

	const png = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+j7xkAAAAASUVORK5CYII=';
	test('a web build needs a working interaction on the saved revision and page', () => {
		const url = URI.joinPath(root, 'index.html').toString();
		const result = (text: string) => ({ success: true, pastTenseMessage: 'Done', content: [{ type: ToolResultContentType.Text as const, text }] });
		const web = new ReachWebVerification('create the app');
		web.observe('reach_create_file', { path: 'index.html' }, result(JSON.stringify({ uri: url })));
		assert.ok(web.beforeFinish(true)?.instruction?.includes(url));
		web.observe('open_browser_page', { url }, result('Page ID: first\n(pageError) RangeError'));
		web.observe('click_element', { pageId: 'first' }, result('Clicked'));
		assert.match(web.beforeFinish(true)!.instruction!, /RangeError/);
		web.observe('reach_edit_file', { path: 'game.js' }, result('Saved'));
		web.observe('click_element', { pageId: 'first' }, result('Clicked'));
		assert.ok(web.pending);
		web.observe('open_browser_page', { url }, result('Page ID: second\nSnapshot: Ready'));
		web.observe('click_element', { pageId: 'second', element: 'Play' }, { ...result('Missing ref'), success: false });
		web.observe('run_playwright_code', { pageId: 'second', code: 'return page.title()' }, result('Result: "Game"'));
		assert.ok(web.pending);
		web.observe('run_playwright_code', { pageId: 'second', code: 'await page.getByText("Play").click()' }, result('[deferredResultId=action] Still running'));
		assert.ok(web.pending);
		web.observe('run_playwright_code', { pageId: 'second', deferredResultId: 'other' }, result('Done'));
		assert.ok(web.pending);
		web.observe('run_playwright_code', { pageId: 'second', deferredResultId: 'action' }, result('Snapshot: Playing'));
		assert.ok(!web.pending);
		const manual = new ReachWebVerification("Create it. I'll test it myself.");
		manual.observe('reach_create_file', { path: 'index.html' }, result(JSON.stringify({ uri: url })));
		assert.ok(!manual.pending);
	});

	test('North Mini retries rejected native tool names with strict schemas and no incomplete call', async () => {
		let requests = 0;
		const agent = await fixture(init => {
			const body = JSON.parse(String(init.body));
			if (++requests === 1) {
				return nativeStream({ type: 'message-end', delta: { finish_reason: 'ERROR', error: 'all generated tool calls were hallucinated' } });
			}
			assert.strictEqual(body.strict_tools, true);
			assert.strictEqual(body.stream, false);
			assert.deepStrictEqual(body.tools.map((t: { function: { name: string } }) => t.function.name), ['reach_read_file']);
			assert.deepStrictEqual(body.tools[0].function.parameters.properties.startLine, { type: 'integer' });
			return Response.json({ finish_reason: 'COMPLETE', message: { content: [{ type: 'text', text: 'Ready.' }] } });
		}, [root], 'north-mini-code-1.0');
		agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = [
			{ name: 'reach_read_file', inputSchema: { type: 'object', properties: { path: { type: 'string' }, startLine: { type: 'integer', minimum: 1 } }, required: ['path'] } }, { name: 'optional_tool' },
		];
		await agent.chats.sendMessage(chat, 'look at the project', undefined, [], 'strict', 'editor');
		assert.strictEqual(requests, 2);
	});
	for (const model of ['command-a-vision-07-2025', 'command-a-plus-05-2026']) {
		test(model + ' sends pasted and embedded PNG pixels and restores the original image context', async () => {
			const uri = URI.joinPath(root, 'pasted.png').toString();
			const readFile = sinon.stub().resolves({ value: VSBuffer.wrap(Buffer.from(png, 'base64')) });
			let requests = 0, saved = '';
			const agent = await fixture(init => {
				const body = JSON.parse(String(init.body));
				assert.strictEqual(body.model, model);
				assert.strictEqual(!!body.tools?.length, model.startsWith('command-a-plus-'));
				const images = body.messages.find((m: { role: string }) => m.role === 'user').content;
				assert.deepStrictEqual(images.map((p: { type: string }) => p.type), ['text', 'image_url', 'image_url'].slice(0, requests === 2 ? 2 : 3));
				assert.match(images[0].text, /What is in the picture/);
				assert.ok(images.slice(1).every((p: { image_url: { url: string } }) => p.image_url.url === 'data:image/png;base64,' + png));
				requests++;
				return nativeStream({ type: 'content-start', delta: { message: { content: { type: 'text', text: 'I see it.' } } } }, { type: 'message-end', delta: { finish_reason: 'COMPLETE' } });
			}, [root], model, { readFile });
			store.add(agent.onDidChangeChatData(event => { saved = event.providerData; }));
			await agent.chats.sendMessage(chat, 'What is in the picture?', undefined, [
				{ type: MessageAttachmentKind.Resource, displayKind: 'image', label: 'pasted.png', uri },
				{ type: MessageAttachmentKind.EmbeddedResource, label: 'second', contentType: 'image/png', data: png },
				{ type: MessageAttachmentKind.Simple, label: 'Browser Pages', modelRepresentation: 'Browser context' },
			], 'image', 'editor');
			const original = JSON.parse(saved).turns[0].images;
			assert.ok(original[0].attachment.uri === uri && !JSON.stringify(original[0]).includes(png));
			await agent.chats.releaseChat(chat, session);
			await agent.materializeChat(chat, session, saved);
			agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = [tool];
			await agent.chats.sendMessage(chat, 'Which color?', undefined, [], 'follow', 'editor');
			readFile.resolves({ value: VSBuffer.wrap(Buffer.concat([Buffer.from(png, 'base64'), Buffer.from('changed')])) });
			await agent.chats.sendMessage(chat, 'Again?', undefined, [], 'changed', 'editor');
			assert.strictEqual(requests, 3);
		});
	}

	for (const firstTurnImage of [true, false]) {
		test('Auto routes ' + (firstTurnImage ? 'a first-turn image to A+' : 'an in-session image to Vision') + ' and replays it after restoring the chat', async () => {
			const vision = 'command-a-vision-07-2025';
			const plus = 'command-a-plus-05-2026';
			const text = 'command-a-03-2025';
			const available = firstTurnImage ? [vision, text, plus] : [text, vision];
			const visual = firstTurnImage ? plus : vision;
			const requests: string[] = [];
			const agent = store.add(new CohereChatAgent({ readFile: () => { throw new Error('unexpected file'); } }, async (url, init) => {
				if (String(url).endsWith('/models')) { return Response.json({ data: available.map(id => ({ id })) }); }
				const body = JSON.parse(String(init?.body));
				requests.push(body.model);
				if (body.model === text) {
					assert.strictEqual(String(url), 'https://api.cohere.com/compatibility/v1/chat/completions');
					return stream({ delta: { content: 'Hello.' }, finish_reason: 'stop' });
				}
				assert.strictEqual(String(url), 'https://api.cohere.com/v2/chat');
				assert.strictEqual(body.model, visual);
				assert.strictEqual(!!body.tools?.length, visual === plus);
				const imageTurn = body.messages.find((m: { role: string; content: unknown }) => m.role === 'user' && Array.isArray(m.content) && m.content.some(part => part.type === 'image_url'));
				assert.strictEqual(imageTurn.content[0].text, 'Describe this.');
				assert.strictEqual(imageTurn.content[1].image_url.url, 'data:image/png;base64,' + png);
				return nativeStream({ type: 'content-start', delta: { message: { content: { type: 'text', text: 'I see it.' } } } }, { type: 'message-end', delta: { finish_reason: 'COMPLETE' } });
			}));
			await agent.authenticate('https://api.cohere.com', 'fixture-key');
			const model = firstTurnImage ? undefined : { id: 'auto' };
			await agent.chats.createChat(chat, session, { model, workingDirectories: [root] });
			agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = [tool];
			if (!firstTurnImage) { await agent.chats.sendMessage(chat, 'Hello.', undefined, [], 'plain', 'editor'); }
			let saved = '';
			store.add(agent.onDidChangeChatData(event => { saved = event.providerData; }));
			await agent.chats.sendMessage(chat, 'Describe this.', undefined, [{ type: MessageAttachmentKind.EmbeddedResource, label: 'pasted.png', contentType: 'image/png', data: png }], 'image', 'editor');
			assert.strictEqual(agent.models.get().find(entry => entry.id === 'auto')?.supportsVision, true);
			await agent.chats.releaseChat(chat, session);
			await agent.materializeChat(chat, session, saved);
			agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = [tool];
			await agent.chats.sendMessage(chat, 'What color?', undefined, [], 'follow', 'editor');
			assert.deepStrictEqual(agent.chats.getModel?.(chat, session), model);
			assert.deepStrictEqual(requests, firstTurnImage ? [plus, plus] : [text, vision, vision]);
		});
	}

	test('North Mini recovers an interrupted file create with a native call and sends its real editor receipt back', async () => {
		let count = 0;
		const agent = await fixture(init => {
			const body = JSON.parse(String(init.body));
			assert.strictEqual(body.response_format, undefined);
			if (++count === 1) {
				return nativeStream({ type: 'tool-call-start', delta: { message: { tool_calls: { id: 'cut', function: { name: 'reach_create_file', arguments: '{"path":"game.js","content":' } } } } }, { type: 'message-end', delta: { finish_reason: 'ERROR', error: 'invalid tool generation' } });
			}
			if (count === 2) {
				assert.strictEqual(body.strict_tools, true);
				assert.strictEqual(body.stream, false);
				return Response.json({ finish_reason: 'TOOL_CALL', message: { content: [{ type: 'text', text: 'Saving the file.' }], tool_plan: 'Write the file.', tool_calls: [{ id: 'complete', type: 'function', function: { name: 'reach_create_file', arguments: '{"path":"game.js","content":"init();"}' } }] } });
			}
			assert.strictEqual(body.messages.at(-1).tool_call_id, 'complete');
			assert.match(body.messages.at(-1).content[0].document.data, /saved fixture/);
			const assistant = body.messages.at(-2);
			assert.strictEqual(assistant.tool_calls[0].id, 'complete');
			assert.ok(!('content' in assistant) && !('tool_plan' in assistant));
			return nativeStream({ type: 'content-start', delta: { message: { content: { type: 'text', text: 'File saved.' } } } }, { type: 'message-end', delta: { finish_reason: 'COMPLETE' } });
		}, [root], 'north-mini-code-1.0');
		agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = [{ name: 'reach_create_file', inputSchema: { type: 'object', properties: { path: { type: 'string' }, content: { type: 'string' } }, required: ['path', 'content'] } }];
		let writes = 0;
		store.add(agent.onDidChatProgress(e => {
			if (e.kind === 'action' && e.action.type === ActionType.ChatToolCallReady) {
				writes++;
				assert.deepStrictEqual(JSON.parse(String(e.action.toolInput)), { path: 'game.js', content: 'init();' });
				agent.onClientToolCallComplete(chat, e.action.toolCallId, { success: true, pastTenseMessage: 'saved fixture' });
			}
		}));
		await agent.chats.sendMessage(chat, 'Create game.js', [root], [], 'turn', 'editor');
		assert.strictEqual(writes, 1);
		assert.strictEqual(count, 3);
	});

	test('North keeps working after a successful write followed by an empty generation', async () => {
		let requests = 0;
		const invoked: unknown[] = [];
		const agent = await fixture(init => {
			const body = JSON.parse(String(init.body));
			requests++;
			if (requests === 2) { return nativeStream({ type: 'message-end', delta: { finish_reason: 'COMPLETE' } }); }
			if (requests === 4) {
				assert.strictEqual(body.messages.at(-1).tool_call_id, 'step-3');
				return nativeStream({ type: 'content-start', delta: { message: { content: { type: 'text', text: 'Verified.' } } } }, { type: 'message-end', delta: { finish_reason: 'COMPLETE' } });
			}
			assert.ok(requests === 1 || requests === 3);
			const name = requests === 1 ? 'reach_create_file' : 'reach_read_file';
			assert.ok(body.tools.some((tool: { function: { name: string } }) => tool.function.name === name));
			const args = requests === 1 ? { path: 'value.txt', content: '42' } : { path: 'value.txt' };
			return nativeStream({ type: 'tool-call-start', delta: { message: { tool_calls: { id: 'step-' + requests, function: { name, arguments: JSON.stringify(args) } } } } }, { type: 'tool-call-end' }, { type: 'message-end', delta: { finish_reason: 'TOOL_CALL' } });
		}, [root], 'north-mini-code-1.0');
		agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = ['reach_create_file', 'reach_read_file'].map(name => ({ name }));
		store.add(agent.onDidChatProgress(e => {
			if (e.kind === 'action' && e.action.type === ActionType.ChatToolCallReady) {
				invoked.push(JSON.parse(String(e.action.toolInput)));
				agent.onClientToolCallComplete(chat, e.action.toolCallId, { success: true, pastTenseMessage: 'Done' });
			}
		}));
		await agent.chats.sendMessage(chat, 'Create value.txt and verify its contents', [root], [], 'turn', 'editor');
		assert.deepStrictEqual(invoked, [{ path: 'value.txt', content: '42' }, { path: 'value.txt' }]);
	});

	for (const scenario of ['progress', 'failure', 'unchanged'] as const) {
		test('North tool round limit with ' + scenario, async () => {
			let requests = 0;
			let completedReads = 0;
			const reader = scenario === 'unchanged' ? 'copilot_readFile' : 'reach_read_file';
			const agent = await fixture(init => {
				const body = JSON.parse(String(init.body));
				requests++;
				assert.ok(requests <= 35);
				if (requests > 1) {
					assert.strictEqual(body.messages.at(-1).tool_call_id, 'read-' + completedReads);
				}
				if (requests === 35) {
					return nativeStream({ type: 'content-start', delta: { message: { content: { type: 'text', text: 'Read all sections.' } } } }, { type: 'message-end', delta: { finish_reason: 'COMPLETE' } });
				}
				const line = scenario === 'unchanged' ? 1 : requests;
				const args = { path: 'notes.txt', startLine: line, endLine: line };
				return nativeStream({ type: 'tool-call-start', delta: { message: { tool_calls: { id: 'read-' + requests, function: { name: reader, arguments: JSON.stringify(args) } } } } }, { type: 'tool-call-end' }, { type: 'message-end', delta: { finish_reason: 'TOOL_CALL' } });
			}, [root], 'north-mini-code-1.0');
			agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = [{ name: reader }];
			store.add(agent.onDidChatProgress(e => {
				if (e.kind === 'action' && e.action.type === ActionType.ChatToolCallReady) {
					completedReads++;
					agent.onClientToolCallComplete(chat, e.action.toolCallId, { success: scenario !== 'failure', pastTenseMessage: 'Read line' });
				}
			}));
			await agent.chats.sendMessage(chat, 'Read the sections of notes.txt in order', [root], [], 'turn', 'editor');
			assert.strictEqual(completedReads, { progress: 34, failure: 32, unchanged: 33 }[scenario]);
			const [turn] = await agent.chats.getMessages(chat, session);
			assert.match(JSON.stringify(turn.responseParts), scenario === 'progress' ? /Read all sections/ : /no new successful editor result/);
		});
	}

	test('rejects unsupported image inputs before a provider request', async () => {
		const never = () => { throw new Error('unexpected inference'); };
		const attachment = { type: MessageAttachmentKind.EmbeddedResource as const, label: 'test.png', contentType: 'image/png', data: png };
		const auto = await fixture(never);
		await assert.rejects(auto.chats.sendMessage(chat, 'Describe it', undefined, [attachment], 'auto', 'editor'), /could not find an image-capable model/);
		assert.strictEqual(auto.models.get().find(entry => entry.id === 'auto')?.supportsVision, false);
		const textOnly = await fixture(never, [root], 'command-a-03-2025');
		await assert.rejects(textOnly.chats.sendMessage(chat, 'Describe it', undefined, [attachment], 'text', 'editor'), /Choose Command A Vision or Command A\+/);
		const vision = await fixture(never, [root], 'command-a-vision-07-2025');
		await assert.rejects(vision.chats.sendMessage(chat, 'Describe it', undefined, [{ ...attachment, data: 'bm90IGFuIGltYWdl' }], 'bad', 'editor'), /Check the format/);
		await assert.rejects(vision.chats.sendMessage(chat, 'Describe these', undefined, Array(21).fill(attachment), 'many', 'editor'), /20 images/);
	});

	test('dispatches fragmented calls to the sending editor and sends the correlated result back to Cohere', async () => {
		let requests = 0;
		const agent = await fixture(init => {
			const body = JSON.parse(String(init.body));
			assert.deepStrictEqual(body.tools.map((item: { function: { name: string } }) => item.function.name), ['workspace_read']);
			if (++requests === 1) {
				assert.ok(body.messages[0].content.includes(root.fsPath));
				return stream({ delta: { tool_calls: [{ index: 0, id: 'call-1', function: { name: 'workspace_read', arguments: '{"path":' } }] } },
					{ delta: { tool_calls: [{ index: 0, function: { arguments: '"README.md"}' } }] }, finish_reason: 'tool_calls' });
			}
			assert.strictEqual(body.messages.at(-2).tool_calls[0].id, 'call-1');
			const result = body.messages.at(-1);
			assert.deepStrictEqual([result.role, result.tool_call_id, result.name], ['tool', 'call-1', 'workspace_read']);
			assert.match(result.content, /Read README.md/);
			return stream({ delta: { content: 'A garden log.' }, finish_reason: 'stop' });
		});
		agent.getOrCreateActiveClient(chat, session, { clientId: 'other' }).tools = [{ name: 'other_tool' }];
		const calls: string[] = [];
		store.add(agent.onDidChatProgress(event => {
			if (event.kind !== 'action') { return; }
			if (event.action.type === ActionType.ChatToolCallStart) { calls.push(event.action.toolName); }
			if (event.action.type === ActionType.ChatToolCallReady) {
				assert.strictEqual(event.action.toolInput, '{"path":"README.md"}');
				agent.onClientToolCallComplete(chat, event.action.toolCallId, { success: true, pastTenseMessage: 'Read README.md' });
			}
		}));
		await agent.chats.sendMessage(chat, prompt, undefined, [], 'turn', 'editor');
		assert.deepStrictEqual(calls, ['workspace.read']);
		assert.strictEqual(requests, 2);
		const [turn] = await agent.chats.getMessages(chat, session);
		assert.match(JSON.stringify(turn.responseParts), /A garden log/);
	});

	test('repeated file reads reference unchanged source and deliver fresh edits made outside the chat', async () => {
		let requests = 0, reads = 0;
		const agent = await fixture(init => {
			const body = JSON.parse(String(init.body));
			requests++;
			const last = body.messages.at(-1);
			if (requests === 2) { assert.match(last.content, /first version/); }
			if (requests === 3) {
				assert.strictEqual(JSON.parse(last.content).unchangedFromToolCall, 'read-1');
				assert.doesNotMatch(last.content, /first version/);
			}
			if (requests === 4) {
				assert.match(last.content, /updated version/);
				return stream({ delta: { content: 'Read updated source.' }, finish_reason: 'stop' });
			}
			return stream({ delta: { tool_calls: [{ id: 'read-' + requests, function: { name: 'reach_read_file', arguments: '{"path":"index.html"}' } }] }, finish_reason: 'tool_calls' });
		});
		agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = [{ name: 'reach_read_file' }];
		store.add(agent.onDidChatProgress(e => {
			if (e.kind === 'action' && e.action.type === ActionType.ChatToolCallReady) {
				const content = ++reads < 3 ? 'first version' : 'updated version';
				agent.onClientToolCallComplete(chat, e.action.toolCallId, { success: true, pastTenseMessage: 'Read file', content: [{ type: ToolResultContentType.Text, text: JSON.stringify({ content }) }] });
			}
		}));
		await agent.chats.sendMessage(chat, 'Read index.html', undefined, [], 'turn', 'editor');
		assert.deepStrictEqual([requests, reads], [4, 3]);
	});

	test('reuses a confirmed edit only while a full editor read still shows that replacement', async () => {
		const edit = 'reach_edit_file', read = 'reach_read_file';
		const steps = [edit, read, edit, read, edit];
		let requests = 0, edits = 0, reads = 0;
		const agent = await fixture(init => {
			const body = JSON.parse(String(init.body));
			const name = steps[requests++];
			if (requests === 4) { assert.strictEqual(JSON.parse(body.messages.at(-1).content).alreadyAppliedToolCall, '1'); }
			if (!name) { return stream({ delta: { content: 'Done.' }, finish_reason: 'stop' }); }
			const args = name === edit ? { path: 'game.html', oldText: 'before', newText: 'after' } : { path: 'game.html' };
			return stream({ delta: { tool_calls: [{ id: String(requests), function: { name, arguments: JSON.stringify(args) } }] }, finish_reason: 'tool_calls' });
		});
		agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = [edit, read].map(name => ({ name }));
		store.add(agent.onDidChatProgress(e => {
			if (e.kind !== 'action' || e.action.type !== ActionType.ChatToolCallReady) { return; }
			const isRead = steps[requests - 1] === read;
			if (!isRead) { edits++; }
			agent.onClientToolCallComplete(chat, e.action.toolCallId, { success: true, pastTenseMessage: 'Done', content: isRead ? [{ type: ToolResultContentType.Text, text: JSON.stringify({ content: ++reads === 1 ? 'after' : 'before' }) }] : [] });
		}));
		await agent.chats.sendMessage(chat, 'Edit game.html', undefined, [], 'turn', 'editor');
		assert.deepStrictEqual([edits, reads, requests], [2, 2, 6]);
	});

	test('leaves tools out of ordinary chat when the sender has none selected', async () => {
		const agent = await fixture(init => {
			const body = JSON.parse(String(init.body));
			assert.strictEqual(body.tools, undefined);
			assert.ok(body.messages[0].content.includes(root.fsPath));
			return stream({ delta: { content: 'Hello' }, finish_reason: 'stop' });
		});
		agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = [];
		await agent.chats.sendMessage(chat, 'Hello', undefined, [], 'turn', 'editor');
	});

	test('folderless file requests keep unrelated tools and exclude workspace writers even from catalog search', async () => {
		const extras = Array.from({ length: 18 }, (_, index) => ({ name: 'calendar_' + index }));
		let requests = 0;
		const agent = await fixture(init => {
			const body = JSON.parse(String(init.body));
			const names = body.tools.map((item: { function: { name: string } }) => item.function.name);
			assert.match(body.messages[0].content, /provide the requested contents inline/);
			if (++requests === 1) {
				assert.ok(names.includes('calendar_0'));
				assert.ok(names.every((name: string) => name === 'reach_find_tools' || name.startsWith('calendar_')));
				return stream({ delta: { tool_calls: [{ id: 'search', function: { name: 'reach_find_tools', arguments: '{"query":"create"}' } }] }, finish_reason: 'tool_calls' });
			}
			assert.deepStrictEqual(JSON.parse(body.messages.at(-1).content), { tools: [], total: 0 });
			return stream({ delta: { content: '<html>Sudoku</html>' }, finish_reason: 'stop' });
		}, []);
		agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = [
			...['copilot_createFile', 'reach_create_file', 'copilot_createDirectory', 'reach_read_file', 'copilot_findFiles', 'run_in_terminal'].map(name => ({ name })), ...extras,
		];
		await agent.chats.sendMessage(chat, 'create a sudoku html file', undefined, [], 'turn', 'editor');
		assert.strictEqual(requests, 2);
	});

	for (const suppliedPath of ['index.html', 'NMC-Sudoku/index.html', URI.joinPath(root, 'NMC-Sudoku/index.html').fsPath]) {
		test('creates inside the attached directory with model path ' + suppliedPath, async () => {
			const folder = { type: MessageAttachmentKind.Resource as const, displayKind: 'directory', label: 'NMC-Sudoku', uri: URI.joinPath(root, 'NMC-Sudoku/').toString() };
			const roots = [root, URI.file('/other')];
			let requests = 0;
			const agent = await fixture(init => {
				const body = JSON.parse(String(init.body));
				if (++requests === 1) {
					assert.match(body.messages.at(-1).content, /"path":"NMC-Sudoku","workspaceFolder":"project"/);
					assert.match(body.messages.at(-1).content, /Browser context/);
					return stream({ delta: { tool_calls: [{ id: 'write', function: { name: 'reach_create_file', arguments: JSON.stringify({ path: suppliedPath, content: 'fixture' }) } }] }, finish_reason: 'tool_calls' });
				}
				return stream({ delta: { content: 'Created.' }, finish_reason: 'stop' });
			}, roots);
			agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = [{ name: 'reach_create_file' }];
			const writes: unknown[] = [];
			store.add(agent.onDidChatProgress(e => {
				if (e.kind === 'action' && e.action.type === ActionType.ChatToolCallReady) {
					writes.push(JSON.parse(String(e.action.toolInput)));
					agent.onClientToolCallComplete(chat, e.action.toolCallId, { success: true, pastTenseMessage: 'Created.' });
				}
			}));
			await agent.chats.sendMessage(chat, 'create your own sudoku version in this folder', roots, [folder, { type: MessageAttachmentKind.Simple, label: 'Browser', modelRepresentation: 'Browser context' }], 'turn', 'editor');
			assert.deepStrictEqual(writes, [{ path: 'NMC-Sudoku/index.html', workspaceFolder: 'project', content: 'fixture' }]);
			assert.deepStrictEqual((await agent.chats.getMessages(chat, session))[0].message.attachments, [folder]);
		});
	}

	test('distinguishes a folder named like an image, the root, and folders outside the selected workspace', async () => {
		let requests = 0;
		const agent = await fixture(init => {
			const content = JSON.parse(String(init.body)).messages.at(-1).content;
			assert.match(content, /"path":"design.png"/);
			assert.match(content, /"path":"\."/);
			requests++;
			return stream({ delta: { content: 'Received both folders.' }, finish_reason: 'stop' });
		});
		const folder = (uri: URI) => ({ type: MessageAttachmentKind.Resource as const, displayKind: 'directory', label: 'design.png', uri: uri.toString() });
		await agent.chats.sendMessage(chat, 'compare', [root], [folder(root), folder(URI.joinPath(root, 'design.png'))], 'turn', 'editor');
		for (const roots of [[], [URI.file('/project-other')]]) {
			await assert.rejects(agent.chats.sendMessage(chat, 'create here', roots, [folder(root)], 'outside', 'editor'), /as a local workspace folder/);
		}
		assert.strictEqual(requests, 1);
	});

	test('creates through the enabled workspace writer and returns its result to Cohere', async () => {
		const otherTools = Array.from({ length: 20 }, (_, index) => ({ name: 'other_' + index, description: 'create a sudoku html file' }));
		let requests = 0;
		const agent = await fixture(init => {
			const body = JSON.parse(String(init.body));
			const names = body.tools.map((item: { function: { name: string } }) => item.function.name);
			assert.strictEqual(names[0], 'reach_create_file');
			assert.ok(names.includes('reach_find_tools') && !names.includes('copilot_createFile'));
			if (++requests === 1) {
				return stream({ delta: { tool_calls: [{ id: 'write', function: { name: 'reach_create_file', arguments: '{"path":"sudoku/index.html","content":"<html>Sudoku</html>"}' } }] }, finish_reason: 'tool_calls' });
			}
			assert.strictEqual(JSON.parse(body.messages.at(-1).content).success, true);
			return stream({ delta: { content: 'Created sudoku/index.html' }, finish_reason: 'stop' });
		});
		agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = [...otherTools, { name: 'copilot_createFile' }, { name: 'reach_create_file' }];
		const paths: string[] = [];
		store.add(agent.onDidChatProgress(event => {
			if (event.kind === 'action' && event.action.type === ActionType.ChatToolCallReady) {
				paths.push(JSON.parse(String(event.action.toolInput)).path);
				agent.onClientToolCallComplete(chat, event.action.toolCallId, { success: true, pastTenseMessage: 'Created sudoku/index.html' });
			}
		}));
		await agent.chats.sendMessage(chat, 'create a sudoku html file', undefined, [], 'turn', 'editor');
		assert.deepStrictEqual(paths, ['sudoku/index.html']);
		assert.strictEqual(requests, 2);
	});

	function rejectedToolStream(): Response {
		return new Response('data: ' + JSON.stringify({ choices: [], error: { code: 422, message: 'all generated tool calls were hallucinated' } }) + '\n\ndata: [DONE]\n\n');
	}

	test('recovers an invalid generation after directory creation without dispatching the directory again', async () => {
		let requests = 0;
		const agent = await fixture(init => {
			const body = JSON.parse(String(init.body));
			switch (++requests) {
				case 1: return stream({ delta: { tool_calls: [{ id: 'dir', function: { name: 'copilot_createDirectory', arguments: '{"dirPath":"sudoku"}' } }] }, finish_reason: 'tool_calls' });
				case 2: return Response.json({ message: 'your request resulted in an invalid tool generation' }, { status: 422 });
				case 3:
					assert.deepStrictEqual(body.tools.map((item: { function: { name: string } }) => item.function.name), ['reach_create_file']);
					assert.strictEqual(body.messages.at(-1).tool_call_id, 'dir');
					return stream({ delta: { tool_calls: [{ id: 'file', function: { name: 'reach_create_file', arguments: '{"path":"sudoku/index.html","content":"Sudoku"}' } }] }, finish_reason: 'tool_calls' });
				case 4:
					assert.strictEqual(body.messages.at(-1).tool_call_id, 'file');
					return stream({ delta: { content: 'Saved the file.' }, finish_reason: 'stop' });
				default: throw new Error('unexpected retry');
			}
		});
		agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = [{ name: 'copilot_createDirectory' }, { name: 'reach_create_file' }];
		const calls: string[] = [];
		store.add(agent.onDidChatProgress(event => {
			if (event.kind !== 'action') { return; }
			if (event.action.type === ActionType.ChatToolCallStart) { calls.push(event.action.toolName); }
			if (event.action.type === ActionType.ChatToolCallReady) {
				agent.onClientToolCallComplete(chat, event.action.toolCallId, { success: true, pastTenseMessage: 'Created' });
			}
		}));
		await agent.chats.sendMessage(chat, 'create sudoku/index.html', undefined, [], 'turn', 'editor');
		assert.deepStrictEqual(calls, ['copilot_createDirectory', 'reach_create_file']);
		assert.strictEqual(requests, 4);
	});

	test('recovers a Cohere in-stream 422 using the advertised wire names', async () => {
		let requests = 0;
		const agent = await fixture(init => {
			if (++requests === 1) { return rejectedToolStream(); }
			const body = JSON.parse(String(init.body));
			assert.match(body.messages[0].content, /exact function names: workspace_read/);
			assert.ok(body.messages[0].content.includes(root.fsPath));
			return stream({ delta: { content: 'Recovered' }, finish_reason: 'stop' });
		});
		await agent.chats.sendMessage(chat, prompt, undefined, [], 'turn', 'editor');
		assert.strictEqual(requests, 2);
		assert.match(JSON.stringify(await agent.chats.getMessages(chat, session)), /Recovered/);
	});

	test('reports the provider error after one rejected retry without dispatching a tool', async () => {
		let requests = 0;
		let invocations = 0;
		const agent = await fixture(() => { requests++; return rejectedToolStream(); });
		store.add(agent.onDidChatProgress(event => {
			if (event.kind === 'action' && event.action.type === ActionType.ChatToolCallStart) { invocations++; }
		}));
		await assert.rejects(agent.chats.sendMessage(chat, prompt, undefined, [], 'turn', 'editor'), /422.*all generated tool calls were hallucinated/);
		assert.deepStrictEqual([requests, invocations], [2, 0]);
	});

	for (const [failure, fail] of [
		['in-stream 422', rejectedToolStream],
		['HTTP 422', () => Response.json({ error: { message: 'all generated tool calls were hallucinated' } }, { status: 422 })],
		['empty final answer', () => stream({ delta: {}, finish_reason: 'stop' })],
	] as const) {
		test(`finishes from two completed reads after ${failure} without offering or rerunning tools`, async () => {
			let requests = 0;
			const paths: string[] = [];
			const agent = await fixture(init => {
				const body = JSON.parse(String(init.body));
				switch (++requests) {
					case 1: return stream({ delta: { tool_calls: ['README.md', 'package.json'].map((path, index) => ({
						index, id: 'read-' + index, function: { name: 'workspace_read', arguments: JSON.stringify({ path }) },
					})) }, finish_reason: 'tool_calls' });
					case 2:
						assert.strictEqual(body.messages.filter((message: { role: string }) => message.role === 'tool').length, 2);
						return fail();
					case 3:
						assert.strictEqual(body.tools, undefined);
						assert.deepStrictEqual(body.messages.map((message: { role: string }) => message.role), ['system', 'user']);
						assert.match(body.messages[1].content, /Observed README.md/);
						assert.match(body.messages[1].content, /Observed package.json/);
						assert.ok(body.messages[1].content.startsWith(prompt));
						return stream({ delta: { content: 'Summary from both sources.' }, finish_reason: 'stop' });
					default: throw new Error('unexpected model request');
				}
			});
			store.add(agent.onDidChatProgress(event => {
				if (event.kind === 'action' && event.action.type === ActionType.ChatToolCallReady) {
					const path = JSON.parse(String(event.action.toolInput)).path;
					paths.push(path);
					agent.onClientToolCallComplete(chat, event.action.toolCallId, { success: true, pastTenseMessage: 'Observed ' + path });
				}
			}));
			await agent.chats.sendMessage(chat, prompt, undefined, [], 'turn', 'editor');
			assert.deepStrictEqual(paths, ['README.md', 'package.json']);
			assert.strictEqual(requests, 3);
			assert.match(JSON.stringify(await agent.chats.getMessages(chat, session)), /Summary from both sources/);
		});
	}

	for (const finalReply of [true, false]) {
		test(`keeps a completed terminal call after "No valid response generated" when the tool-free answer ${finalReply ? 'succeeds' : 'also fails'}`, async () => {
			const banner = 'Serving HTTP on :: port 8000';
			const error = { code: 422, message: 'No valid response generated. Try updating messages' };
			let requests = 0;
			const agent = await fixture(init => {
				const body = JSON.parse(String(init.body));
				switch (++requests) {
					case 1: return stream({ delta: { tool_calls: [{ id: 'terminal', function: { name: 'run_in_terminal', arguments: '{}' } }] }, finish_reason: 'tool_calls' });
					case 2:
						assert.match(body.messages.at(-1).content, /Serving HTTP/);
						return new Response('data: ' + JSON.stringify({ choices: [], error }) + '\n\ndata: [DONE]\n\n');
					case 3:
						assert.strictEqual(body.tools, undefined);
						assert.deepStrictEqual(body.messages.map((message: { role: string }) => message.role), ['system', 'user']);
						assert.ok(body.messages.at(-1).content.includes(banner));
						return finalReply ? stream({ delta: { content: 'The server started. The link still needs a check.' }, finish_reason: 'stop' })
							: Response.json({ error }, { status: 422 });
					default: throw new Error('unexpected extra request');
				}
			});
			agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = [{ name: 'run_in_terminal' }];
			let calls = 0;
			store.add(agent.onDidChatProgress(event => {
				if (event.kind === 'action' && event.action.type === ActionType.ChatToolCallReady) {
					calls++;
					agent.onClientToolCallComplete(chat, event.action.toolCallId, { success: true, pastTenseMessage: banner });
				}
			}));
			await agent.chats.sendMessage(chat, 'The link is broken', undefined, [], 'turn', 'editor');
			const [turn] = await agent.chats.getMessages(chat, session);
			assert.match(JSON.stringify(turn.responseParts), finalReply ? /link still needs a check/ : /run_in_terminal.*completed/);
			assert.deepStrictEqual([requests, calls], [3, 1]);
		});
	}

	test('never dispatches a truncated tool call', async () => {
		const agent = await fixture(() => stream({ delta: { tool_calls: [{ id: 'partial', function: { name: 'workspace_read', arguments: '{}' } }] }, finish_reason: 'length' }));
		let calls = 0;
		store.add(agent.onDidChatProgress(event => { if (event.kind === 'action' && event.action.type === ActionType.ChatToolCallStart) { calls++; } }));
		await assert.rejects(agent.chats.sendMessage(chat, prompt, undefined, [], 'turn', 'editor'), /incomplete tool request/);
		assert.strictEqual(calls, 0);
	});

	function nativeStream(...events: unknown[]): Response {
		return new Response(events.map(event => 'data: ' + JSON.stringify(event) + '\n\n').join('') + 'data: [DONE]\n\n');
	}

	test('makes a path from the selected folder relative and leaves an outside path for the native guard', async () => {
		const inside = URI.joinPath(root, 'sudoku.html').fsPath;
		const outside = URI.joinPath(root, '..', 'other', 'sudoku.html').fsPath;
		let requests = 0;
		const agent = await fixture(() => {
			if (++requests <= 2) {
				const call = { id: String(requests), function: { name: 'reach_read_file', arguments: JSON.stringify({ path: requests === 1 ? inside : outside }) } };
				return nativeStream({ type: 'tool-call-start', delta: { message: { tool_calls: call } } }, { type: 'tool-call-end' }, { type: 'message-end', delta: { finish_reason: 'TOOL_CALL' } });
			}
			return nativeStream({ type: 'content-start', delta: { message: { content: { type: 'text', text: 'Done.' } } } }, { type: 'message-end', delta: { finish_reason: 'COMPLETE' } });
		}, [root], 'command-a-plus-05-2026');
		agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = [{ name: 'reach_read_file' }];
		const paths: string[] = [];
		store.add(agent.onDidChatProgress(e => {
			if (e.kind === 'action' && e.action.type === ActionType.ChatToolCallReady) {
				paths.push(JSON.parse(String(e.action.toolInput)).path);
				agent.onClientToolCallComplete(chat, e.action.toolCallId, { success: requests === 1, pastTenseMessage: 'Read result' });
			}
		}));
		await agent.chats.sendMessage(chat, 'Read files', undefined, [], 'turn', 'editor');
		assert.deepStrictEqual(paths, ['sudoku.html', outside]);
	});

	test('reuses an open preview until a source change allows a fresh preview', async () => {
		const openPage = 'open_browser_page', edit = 'reach_edit_file';
		const invoke = (name: string, id: number) => nativeStream({ type: 'tool-call-start', delta: { message: { tool_calls: { id: String(id), function: { name, arguments: JSON.stringify(name === openPage ? { url: 'file:///project/game.html' } : { path: 'game.html', oldText: 'one', newText: 'two' }) } } } } }, { type: 'tool-call-end' }, { type: 'message-end', delta: { finish_reason: 'TOOL_CALL' } });
		let requests = 0;
		const agent = await fixture(init => {
			const body = JSON.parse(String(init.body));
			switch (++requests) {
				case 1: case 2: case 4: return invoke(openPage, requests);
				case 3:
					assert.match(body.messages.at(-1).content[0].document.data, /page-1/);
					assert.ok(!body.tools.some((item: { function: { name: string } }) => item.function.name === openPage));
					return invoke(edit, requests);
				case 5: return nativeStream({ type: 'content-start', delta: { message: { content: { type: 'text', text: 'Done.' } } } }, { type: 'message-end', delta: { finish_reason: 'COMPLETE' } });
				default: throw new Error('Unexpected request');
			}
		}, [root], 'command-a-plus-05-2026');
		agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = [openPage, edit].map(name => ({ name }));
		const dispatched: string[] = [];
		store.add(agent.onDidChatProgress(e => {
			if (e.kind !== 'action') { return; }
			if (e.action.type === ActionType.ChatToolCallStart) { dispatched.push(e.action.toolName); }
			if (e.action.type === ActionType.ChatToolCallReady) { agent.onClientToolCallComplete(chat, e.action.toolCallId, { success: true, pastTenseMessage: 'page-' + requests }); }
		}));
		await agent.chats.sendMessage(chat, 'Preview the changed file', undefined, [], 'turn', 'editor');
		assert.deepStrictEqual(dispatched, [openPage, edit, openPage]);
	});

	test('continues past a server rejection of its own tool history without rerunning the editor action', async () => {
		let requests = 0, executions = 0;
		const call = (id: string) => nativeStream({ type: 'tool-call-start', delta: { message: { tool_calls: { id, function: { name: 'workspace_read', arguments: '{}' } } } } }, { type: 'tool-call-end' }, { type: 'message-end', delta: { finish_reason: 'TOOL_CALL' } });
		const agent = await fixture(init => {
			const body = JSON.parse(String(init.body));
			switch (++requests) {
				case 1: return call('first');
				case 2: return Response.json({ message: 'invalid tool call provided in messages[2].tool_calls[0]: tool arguments must be a stringified JSON object' }, { status: 400 });
				case 3:
					assert.strictEqual(executions, 1);
					assert.strictEqual(body.messages[2].tool_calls, undefined);
					assert.match(body.messages[3].content, /Already returned editor result.*workspace_read/);
					return call('second');
				case 4:
					assert.strictEqual(body.messages[2].tool_calls, undefined);
					assert.strictEqual(body.messages.at(-1).tool_call_id, 'second');
					return nativeStream({ type: 'content-start', delta: { message: { content: { type: 'text', text: 'Finished.' } } } }, { type: 'message-end', delta: { finish_reason: 'COMPLETE' } });
				default: throw new Error('Unexpected request');
			}
		}, [root], 'command-a-plus-05-2026');
		store.add(agent.onDidChatProgress(e => {
			if (e.kind === 'action' && e.action.type === ActionType.ChatToolCallReady) {
				executions++;
				agent.onClientToolCallComplete(chat, e.action.toolCallId, { success: true, pastTenseMessage: 'Read' });
			}
		}));
		await agent.chats.sendMessage(chat, prompt, undefined, [], 'turn', 'editor');
		assert.deepStrictEqual([requests, executions], [4, 2]);
	});

	for (const scenario of ['generation error', 'repeated destination', 'truncated draft']) {
		test('requests one structured file draft after ' + scenario, async () => {
			const complete = scenario !== 'truncated draft';
			const parameters = { type: 'object' as const, properties: { path: { type: 'string' as const }, content: { type: 'string' as const } }, required: ['path', 'content'] };
			const input = { path: 'sudoku.html', content: '<main>Board</main>' };
			let requests = 0;
			const agent = await fixture(init => {
				const body = JSON.parse(String(init.body));
				switch (++requests) {
					case 1:
						if (scenario === 'repeated destination') {
							const batch = [0, 1].flatMap(index => [
								{ type: 'tool-call-start', index, delta: { message: { tool_calls: { id: 'duplicate-' + index, function: { name: 'reach_create_file', arguments: JSON.stringify({ ...input, content: 'version ' + index }) } } } } },
								{ type: 'tool-call-end', index },
							]);
							return nativeStream(...batch, { type: 'message-end', delta: { finish_reason: 'TOOL_CALL' } });
						}
						return nativeStream(
							{ type: 'tool-call-start', delta: { message: { tool_calls: { id: 'partial', function: { name: 'reach_create_file', arguments: '{"path":' } } } } },
							{ type: 'message-end', delta: { finish_reason: 'ERROR', error: 'your request resulted in an invalid tool generation' } });
					case 2:
						assert.deepStrictEqual(body.response_format, { type: 'json_object', schema: parameters });
						assert.strictEqual(body.thinking, undefined);
						assert.strictEqual(body.stream, false);
						assert.strictEqual(body.tools, undefined);
						assert.match(body.messages.at(-1).content, /Generate one JSON object/);
						assert.ok(!JSON.stringify(body.messages).includes('"partial"'));
						return Response.json({ message: { content: [{ type: 'thinking', thinking: 'Preparing the file' }, { type: 'text', text: JSON.stringify(input) }] }, finish_reason: complete ? 'COMPLETE' : 'MAX_TOKENS' });
					case 3:
						assert.strictEqual(body.response_format, undefined);
						assert.strictEqual(body.messages.at(-2).tool_calls.length, 1);
						assert.strictEqual(body.messages.at(-2).tool_calls[0].id, body.messages.at(-1).tool_call_id);
						return nativeStream({ type: 'content-start', delta: { message: { content: { type: 'text', text: 'Saved.' } } } }, { type: 'message-end', delta: { finish_reason: 'COMPLETE' } });
					default: throw new Error('Unexpected request');
				}
			}, [root], 'command-a-plus-05-2026');
			agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = [{ name: 'reach_create_file', inputSchema: parameters }];
			const calls: string[] = [], visible: string[] = [];
			store.add(agent.onDidChatProgress(e => {
				if (e.kind !== 'action') { return; }
				if (e.action.type === ActionType.ChatDelta) { visible.push(e.action.content); }
				if (e.action.type === ActionType.ChatToolCallReady) {
					calls.push(String(e.action.toolInput));
					agent.onClientToolCallComplete(chat, e.action.toolCallId, { success: true, pastTenseMessage: 'Saved' });
				}
			}));
			const turn = agent.chats.sendMessage(chat, 'Create sudoku.html', undefined, [], 'turn', 'editor');
			if (complete) { await turn; } else { await assert.rejects(turn, /output limit/); }
			assert.deepStrictEqual(calls, complete ? [JSON.stringify(input)] : []);
			assert.deepStrictEqual(visible, complete ? ['Saved.'] : []);
			assert.strictEqual(requests, complete ? 3 : 2);
		});
	}

	for (const [model, extraTools] of [['command-a-plus-05-2026', 0], ['command-a-plus-05-2026', 20], ['north-mini-code-1.0', 20]] as const) {
		test(model + ' edits an attached file in place and can discover a separately requested writer with ' + extraTools + ' extra tools', async () => {
			const path = 'boards/game.html';
			const readFile = sinon.stub().resolves({ value: VSBuffer.fromString('<main>board</main>') });
			const edit = 'reach_edit_file', create = 'reach_create_file', search = 'reach_find_tools', read = 'reach_read_file';
			let requests = 0;
			const agent = await fixture(init => {
				const body = JSON.parse(String(init.body));
				const names = body.tools.map((t: { function: { name: string } }) => t.function.name);
				const call = (name: string, args: object) => nativeStream({ type: 'tool-call-start', delta: { message: { tool_calls: { id: 'call-' + requests, function: { name, arguments: JSON.stringify(args) } } } } }, { type: 'tool-call-end' }, { type: 'message-end', delta: { finish_reason: 'TOOL_CALL' } });
				switch (++requests) {
					case 1:
						assert.ok(names.includes(edit) && names.includes(search) && !names.includes(create));
						assert.match(body.messages[0].content, /boards\/game\.html/);
						return call(read, { path });
					case 2:
					case 3: return call(read, { path });
					case 4:
						assert.ok(names.includes(read), 'an unchanged read must still allow a narrower range before the edit');
						return call(edit, { path, oldText: 'board', newText: 'Sudoku' });
					case 5: return call(search, { query: create });
					case 6:
						assert.ok(names.includes(create));
						return call(create, { path: 'notes.txt', content: 'Changed the board.' });
					case 7: return nativeStream({ type: 'content-start', delta: { message: { content: { type: 'text', text: 'Updated the board and wrote the notes.' } } } }, { type: 'message-end', delta: { finish_reason: 'COMPLETE' } });
					default: throw new Error('Unexpected request');
				}
			}, [root], model, { readFile });
			agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = [read, edit, create, ...Array.from({ length: extraTools }, (_, i) => 'other_' + i)].map(name => ({ name }));
			const paths: string[] = [];
			store.add(agent.onDidChatProgress(e => {
				if (e.kind === 'action' && e.action.type === ActionType.ChatToolCallReady) {
					const args = JSON.parse(String(e.action.toolInput));
					if (args.oldText || args.content) { paths.push(args.path); }
					agent.onClientToolCallComplete(chat, e.action.toolCallId, { success: true, pastTenseMessage: 'Done' });
				}
			}));
			const uri = URI.joinPath(root, path);
			await agent.chats.sendMessage(chat, 'Rebuild this to be better than New_Sudoku.html and save the release notes in notes.txt', undefined, [{ type: MessageAttachmentKind.Resource, uri: uri.toString(), label: 'game.html', displayKind: 'document' }], 'turn', 'editor');
			assert.deepStrictEqual(paths, [path, 'notes.txt']);
			assert.strictEqual(readFile.firstCall.args[0].toString(), uri.toString());
		});
	}

	test('keeps confirmed changes in the saved reply when a later file draft is incomplete', async () => {
		let requests = 0, writes = 0;
		const agent = await fixture(init => {
			const body = JSON.parse(String(init.body));
			switch (++requests) {
				case 1: return nativeStream({ type: 'tool-call-start', delta: { message: { tool_calls: { id: 'saved', function: { name: 'reach_edit_file', arguments: '{"path":"game.html","oldText":"old","newText":"new"}' } } } } }, { type: 'tool-call-end' }, { type: 'message-end', delta: { finish_reason: 'TOOL_CALL' } });
				case 2: return nativeStream({ type: 'tool-call-start', delta: { message: { tool_calls: { id: 'partial', function: { name: 'reach_create_file', arguments: '{"path":' } } } } }, { type: 'message-end', delta: { finish_reason: 'ERROR', error: 'invalid tool generation' } });
				case 3:
					assert.ok(body.response_format.schema);
					return Response.json({ message: { content: [{ type: 'text', text: '{"path":"another.html"}' }] }, finish_reason: 'COMPLETE' });
				default: throw new Error('Unexpected request');
			}
		}, [root], 'command-a-plus-05-2026');
		agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = ['reach_edit_file', 'reach_create_file'].map(name => ({ name }));
		store.add(agent.onDidChatProgress(e => {
			if (e.kind === 'action' && e.action.type === ActionType.ChatToolCallReady) {
				writes++;
				agent.onClientToolCallComplete(chat, e.action.toolCallId, { success: true, pastTenseMessage: 'Saved' });
			}
		}));
		await agent.chats.sendMessage(chat, 'Edit game.html and create another.html', undefined, [], 'turn', 'editor');
		assert.strictEqual(writes, 1);
		const [turn] = await agent.chats.getMessages(chat, session);
		assert.match(JSON.stringify(turn.responseParts), /This draft was not saved.*editor confirmed changes to game\.html/);
	});

	test('continues a failed attached request from host history and obtains a real final reply after a tool stream error', async () => {
		let requests = 0, edits = 0;
		const agent = await fixture(init => {
			const body = JSON.parse(String(init.body));
			switch (++requests) {
				case 1:
					assert.match(body.messages[0].content, /game\.html/);
					assert.match(body.messages.at(-1).content, /Rebuild this to be better than reference\.html/);
					return nativeStream({ type: 'tool-call-start', delta: { message: { tool_calls: { id: 'edit', function: { name: 'reach_edit_file', arguments: '{"path":"game.html","oldText":"old","newText":"new"}' } } } } }, { type: 'tool-call-end' }, { type: 'message-end', delta: { finish_reason: 'TOOL_CALL' } });
				case 2: return nativeStream({ type: 'message-end', delta: { finish_reason: 'ERROR', error: 'all generated tool calls were hallucinated' } });
				case 3:
					assert.strictEqual(body.stream, false);
					assert.strictEqual(body.tools, undefined);
					assert.ok(body.messages.every((item: { role: string }) => item.role !== 'tool'));
					assert.match(body.messages.at(-1).content, /Saved game\.html/);
					assert.deepStrictEqual(body.response_format.schema.required, ['reply']);
					return Response.json({ message: { content: [{ type: 'text', text: '{"reply":"I updated game.html."}' }] }, finish_reason: 'COMPLETE' });
				default: throw new Error('Unexpected request');
			}
		}, [root], 'command-a-plus-05-2026');
		agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = [{ name: 'reach_edit_file' }, { name: 'reach_create_file' }];
		store.add(agent.onDidChatProgress(e => {
			if (e.kind === 'action' && e.action.type === ActionType.ChatToolCallReady) {
				edits++;
				agent.onClientToolCallComplete(chat, e.action.toolCallId, { success: true, pastTenseMessage: 'Saved game.html' });
			}
		}));
		await agent.chats.sendMessage(chat, 'finish what you were doing', undefined, [], 'resume', 'editor', undefined, {
			...resolveAgentChatContext(session, chat),
			previousTurns: [{ id: 'first', state: TurnState.Error, startedAt: new Date().toISOString(), usage: undefined, responseParts: [], message: { text: 'Rebuild this to be better than reference.html', origin: { kind: MessageKind.User }, attachments: [{ type: MessageAttachmentKind.Resource, label: 'game.html', uri: URI.joinPath(root, 'game.html').toString() }] } }],
		});
		assert.deepStrictEqual([requests, edits], [3, 1]);
		assert.match(JSON.stringify((await agent.chats.getMessages(chat, session))[0].responseParts), /I updated game\.html/);
	});

	test('a failed post-read recovery never records an interruption notice as a finished answer', async () => {
		let requests = 0;
		const agent = await fixture(() => {
			if (++requests === 1) { return nativeStream({ type: 'tool-call-start', delta: { message: { tool_calls: { id: 'read', function: { name: 'workspace_read', arguments: '{}' } } } } }, { type: 'tool-call-end' }, { type: 'message-end', delta: { finish_reason: 'TOOL_CALL' } }); }
			if (requests === 2) { return nativeStream({ type: 'message-end', delta: { finish_reason: 'ERROR' } }); }
			return Response.json({ finish_reason: 'ERROR', error: 'Generation unavailable' });
		}, [root], 'command-a-plus-05-2026');
		store.add(agent.onDidChatProgress(e => {
			if (e.kind === 'action' && e.action.type === ActionType.ChatToolCallReady) { agent.onClientToolCallComplete(chat, e.action.toolCallId, { success: true, pastTenseMessage: 'Read file' }); }
		}));
		await assert.rejects(agent.chats.sendMessage(chat, prompt, undefined, [], 'turn', 'editor'), /Generation unavailable/);
		assert.strictEqual(requests, 3);
		assert.deepStrictEqual(await agent.chats.getMessages(chat, session), []);
	});

	for (const outcome of ['saved', 'denied', 'stale', 'truncated']) {
		test('a single-file rebuild uses approved sources and releases its model summary only after save: ' + outcome, async () => {
			let requests = 0;
			const calls: string[] = [], deltas: string[] = [];
			const revision = 'a'.repeat(40);
			const agent = await fixture(init => {
				const body = JSON.parse(String(init.body));
				if (++requests === 2 && outcome === 'stale') {
					assert.match(body.messages.at(-1).content[0].document.data, /revision changed/);
					return nativeStream({ type: 'content-start', delta: { message: { content: { type: 'text', text: 'The file changed on disk.' } } } }, { type: 'message-end', delta: { finish_reason: 'COMPLETE' } });
				}
				assert.strictEqual(requests, 1);
				assert.deepStrictEqual(body.response_format.schema.required, ['content', 'summary']);
				assert.strictEqual(body.tools, undefined);
				assert.strictEqual(body.thinking, undefined);
				assert.strictEqual(body.stream, true);
				assert.match(body.messages.at(-1).content, /original game\.html.*original reference\.html/s);
				return nativeStream(
					{ type: 'content-start', index: 0, delta: { message: { content: { type: 'thinking', thinking: 'Internal' } } } }, { type: 'content-end', index: 0 },
					{ type: 'content-start', index: 1, delta: { message: { content: { type: 'text', text: '{"content":"<main>New grid</main>",' } } } },
					{ type: 'content-delta', index: 1, delta: { message: { content: { text: '"summary":"Rebuilt game.html with a new grid."}' } } } },
					{ type: 'message-end', delta: { finish_reason: outcome === 'truncated' ? 'MAX_TOKENS' : 'COMPLETE' } });
			}, [root], 'command-a-plus-05-2026', { readFile: sinon.stub().resolves({ value: VSBuffer.fromString('original game.html') }) });
			agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = ['reach_read_file', 'reach_edit_file', 'reach_create_file'].map(name => ({ name }));
			store.add(agent.onDidChatProgress(e => {
				if (e.kind !== 'action') { return; }
				if (e.action.type === ActionType.ChatDelta) { deltas.push(e.action.content); }
				if (e.action.type !== ActionType.ChatToolCallReady) { return; }
				const args = JSON.parse(String(e.action.toolInput));
				assert.deepStrictEqual(deltas, []);
				calls.push(args.path);
				if (args.newText === undefined) {
					agent.onClientToolCallComplete(chat, e.action.toolCallId, { success: true, pastTenseMessage: 'Read', content: [{ type: ToolResultContentType.Text, text: JSON.stringify({ fileRevision: revision, content: 'original ' + args.path }) }] });
				} else {
					assert.deepStrictEqual(args, { path: 'game.html', fileRevision: revision, newText: '<main>New grid</main>' });
					agent.onClientToolCallComplete(chat, e.action.toolCallId, outcome === 'saved' ? { success: true, pastTenseMessage: 'Saved' } : { success: false, pastTenseMessage: 'Not saved', error: { code: outcome === 'denied' ? 'denied' : 'stale', message: 'revision changed' } });
				}
			}));
			const turn = agent.chats.sendMessage(chat, 'Rebuild this to be better than reference.html', undefined, [{ type: MessageAttachmentKind.Resource, label: 'game.html', uri: URI.joinPath(root, 'game.html').toString() }], 'turn', 'editor');
			if (outcome === 'truncated') { await assert.rejects(turn, /original file was preserved/); } else { await turn; }
			assert.deepStrictEqual(calls, outcome === 'truncated' ? ['game.html', 'reference.html'] : ['game.html', 'reference.html', 'game.html']);
			assert.strictEqual(deltas.join('').includes('Rebuilt game.html with a new grid.'), outcome === 'saved');
		});
	}

	for (const invalidRange of [false, true]) {
		test('an explicit localized edit uses the revision of its selected lines: ' + (invalidRange ? 'outside the file' : 'in bounds'), async () => {
			const calls: string[] = [];
			let requests = 0;
			const source = '<main>\r\n<span>Theme</span>\r\n</main>';
			const expectedRevision = '2:2:' + await hashAsync('<span>Theme</span>');
			const agent = await fixture(init => {
				const body = JSON.parse(String(init.body));
				const draft = { startLine: 2, endLine: invalidRange ? 4 : 2, newText: '<button>Theme</button>', summary: 'Updated game.html.' };
				const response = ++requests === 1 ? JSON.stringify(draft) : 'The requested range was invalid.';
				if (requests === 1) {
					assert.deepStrictEqual(body.response_format.schema.required, ['startLine', 'endLine', 'newText', 'summary']);
					assert.ok(body.messages.at(-1).content.includes('1\t<main>\n2\t<span>Theme</span>\n3\t</main>'));
				} else {
					assert.ok(invalidRange && body.tools.length);
				}
				return nativeStream({ type: 'content-start', delta: { message: { content: { type: 'text', text: response } } } }, { type: 'message-end', delta: { finish_reason: 'COMPLETE' } });
			}, [root], 'command-a-plus-05-2026');
			agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = [{ name: 'reach_read_file' }, { name: 'reach_edit_file' }];
			store.add(agent.onDidChatProgress(e => {
				if (e.kind !== 'action' || e.action.type !== ActionType.ChatToolCallReady) { return; }
				const args = JSON.parse(String(e.action.toolInput));
				assert.strictEqual(args.path, 'game.html');
				if (args.newText === undefined) {
					calls.push('read');
					agent.onClientToolCallComplete(chat, e.action.toolCallId, { success: true, pastTenseMessage: 'Read', content: [{ type: ToolResultContentType.Text, text: JSON.stringify({ content: source }) }] });
				} else {
					assert.deepStrictEqual(args, { path: 'game.html', fileRevision: expectedRevision, newText: '<button>Theme</button>' });
					calls.push('edit');
					agent.onClientToolCallComplete(chat, e.action.toolCallId, { success: true, pastTenseMessage: 'Saved' });
				}
			}));
			await agent.chats.sendMessage(chat, 'Repair the theme control in game.html. Rename the element. Make this one localized edit and report the result.', undefined, [], 'turn', 'editor');
			assert.deepStrictEqual(calls, invalidRange ? ['read'] : ['read', 'edit']);
			assert.strictEqual(requests, invalidRange ? 2 : 1);
		});
	}

	test('groups related edits from a full small-file read and preserves the boundary of a range read', () => {
		const read = (file: object) => toCohereMessage({
			role: 'tool', name: 'reach_read_file', toolCallId: 'read',
			content: JSON.stringify({ success: true, content: [{ text: JSON.stringify({ content: '<main></main>', fileRevision: 'revision', ...file }) }] }),
		}).content;
		const [whole, range] = [read({}), read({ startLine: 3, endLine: 3, totalLines: 12 })];
		assert.match(JSON.stringify(whole), /one reach_edit_file call.*complete revised file/);
		assert.match(JSON.stringify(range), /Lines 3–3 of 12.*only the lines below/);
		assert.doesNotMatch(JSON.stringify(range), /complete revised file/);
	});

	test('uses the preceding file turn in a restored chat to select tools for an unnamed edit', async () => {
		const model = 'command-a-plus-05-2026';
		const extras = Array.from({ length: 20 }, (_, i) => ({ name: 'notebook_' + i, description: 'more improvements and style' }));
		const fileTools = ['reach_create_file', 'reach_read_file', 'reach_edit_file'];
		let requests = 0;
		const agent = await fixture(init => {
			const body = JSON.parse(String(init.body));
			const names = body.tools.map((item: { function: { name: string } }) => item.function.name);
			if (++requests === 1) {
				assert.match(body.messages.at(-2).content, /New_Sudoku_Game\.html/);
				assert.deepStrictEqual(names, [...fileTools, 'reach_find_tools']);
			} else {
				assert.ok(names.includes('calendar_events'));
			}
			return nativeStream({ type: 'content-start', delta: { message: { content: { type: 'text', text: 'Updated New_Sudoku_Game.html.' } } } }, { type: 'message-end', delta: { finish_reason: 'COMPLETE' } });
		}, [root], model);
		await agent.chats.releaseChat(chat, session);
		await agent.materializeChat(chat, session, JSON.stringify({
			version: 1, cwd: root.toString(), model, createdAt: 1, modifiedAt: 2,
			turns: [{ id: 'previous', startedAt: new Date(0).toISOString(), prompt: 'continue', reply: 'The editor confirmed changes to New_Sudoku_Game.html.' }],
		}));
		agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = [...extras, ...fileTools.map(name => ({ name })), { name: 'calendar_events' }];
		await agent.chats.sendMessage(chat, 'there\'s no grid, make more improvements and improve the style, add a dark mode', undefined, [], 'edit', 'editor');
		await agent.chats.sendMessage(chat, 'Show my calendar events', undefined, [], 'calendar', 'editor');
		assert.strictEqual(requests, 2);
	});

	test('follow-up history replays completed tool actions with the assistant reply', async () => {
		const model = 'command-a-plus-05-2026';
		let followUpBody: { messages: unknown } | undefined;
		const agent = await fixture(init => {
			followUpBody = JSON.parse(String(init.body));
			return nativeStream({ type: 'content-start', delta: { message: { content: { type: 'text', text: 'Updated.' } } } }, { type: 'message-end', delta: { finish_reason: 'COMPLETE' } });
		}, [root], model);
		await agent.chats.releaseChat(chat, session);
		await agent.materializeChat(chat, session, JSON.stringify({
			version: 1, cwd: root.toString(), model, createdAt: 1, modifiedAt: 2,
			turns: [{ id: 'previous', startedAt: new Date(0).toISOString(), prompt: 'build it', reply: 'Created the thing.', actions: ['- reach_create_file completed.'] }],
		}));
		await agent.chats.sendMessage(chat, 'now update it', undefined, [], 'follow', 'editor');
		const wire = JSON.stringify(followUpBody!.messages);
		assert.match(wire, /Created the thing/);
		assert.match(wire, /Actions completed in that turn:\\n- reach_create_file completed\./);
	});

	test('recovers an empty first stream with one native JSON response, dispatches its complete tool once, and finishes the turn', async () => {
		let requests = 0;
		let firstBody: unknown;
		const agent = await fixture(init => {
			const body = JSON.parse(String(init.body));
			switch (++requests) {
				case 1:
					assert.strictEqual(body.stream, true);
					firstBody = body;
					return nativeStream({ type: 'message-end', delta: { finish_reason: 'ERROR', error: 'Original generation detail' } });
				case 2:
					assert.strictEqual(body.stream, false);
					assert.deepStrictEqual({ ...body, stream: true }, firstBody);
					return Response.json({ finish_reason: 'TOOL_CALL', message: { role: 'assistant', tool_plan: 'Read the file.', tool_calls: [
						{ type: 'function', id: 'complete', function: { name: 'workspace_read', arguments: '{"path":"index.html"}' } },
					] } });
				case 3:
					assert.strictEqual(body.stream, true);
					assert.strictEqual(body.messages.at(-2).tool_plan, 'Read the file.');
					assert.strictEqual(body.messages.at(-1).tool_call_id, 'complete');
					return nativeStream({ type: 'content-start', delta: { message: { content: { type: 'text', text: 'Finished.' } } } }, { type: 'message-end', delta: { finish_reason: 'COMPLETE' } });
				default: throw new Error('Unexpected retry');
			}
		}, [root], 'command-a-plus-05-2026');
		const calls: string[] = [], deltas: string[] = [], notices: string[] = [];
		store.add(agent.onDidChatProgress(event => {
			if (event.kind !== 'action') { return; }
			if (event.action.type === ActionType.ChatResponsePart && event.action.part.kind === ResponsePartKind.SystemNotification) {
				const content = event.action.part.content;
				notices.push(typeof content === 'string' ? content : content.markdown);
			}
			if (event.action.type === ActionType.ChatDelta) { deltas.push(event.action.content); }
			if (event.action.type === ActionType.ChatToolCallReady) {
				calls.push(String(event.action.toolInput));
				agent.onClientToolCallComplete(chat, event.action.toolCallId, { success: true, pastTenseMessage: 'Read file' });
			}
		}));
		await agent.chats.sendMessage(chat, prompt, undefined, [], 'turn', 'editor');
		assert.ok(notices.some(notice => notice.includes('Original generation detail')));
		assert.deepStrictEqual([requests, calls, deltas], [3, ['{"path":"index.html"}'], ['Finished.']]);
	});

	for (const failure of ['provider error', 'duplicate IDs', 'invalid arguments']) {
		test('stops after one native JSON retry without dispatching a tool for ' + failure, async () => {
			let requests = 0, dispatches = 0;
			const agent = await fixture(init => {
				if (++requests === 1) { return nativeStream({ type: 'message-end', delta: { finish_reason: 'ERROR' } }); }
				assert.strictEqual(requests, 2);
				assert.strictEqual(JSON.parse(String(init.body)).stream, false);
				if (failure === 'provider error') { return Response.json({ finish_reason: 'ERROR', error: 'Fixture upstream detail' }); }
				const call = { type: 'function', id: 'read', function: { name: 'workspace_read', arguments: '{}' } };
				const second = failure === 'duplicate IDs' ? call : { ...call, id: 'broken', function: { ...call.function, arguments: '{"path":' } };
				return Response.json({ finish_reason: 'TOOL_CALL', message: { tool_calls: [call, second] } });
			}, [root], 'command-a-plus-05-2026');
			store.add(agent.onDidChatProgress(event => { if (event.kind === 'action' && event.action.type === ActionType.ChatToolCallStart) { dispatches++; } }));
			const expected = failure === 'provider error' ? /Fixture upstream detail/ : failure === 'duplicate IDs' ? /duplicate tool call IDs/ : /malformed tool arguments/;
			await assert.rejects(agent.chats.sendMessage(chat, prompt, undefined, [], 'turn', 'editor'), expected);
			assert.deepStrictEqual([requests, dispatches], [2, 0]);
		});
	}

	test('preserves the native tool plan and breaks unchanged verification reads after a confirmed write', async () => {
		const model = 'command-a-plus-05-2026';
		let round = 0;
		const agent = await fixture(init => {
			const body = JSON.parse(String(init.body));
			round++;
			const read = round === 2 || round === 3;
			if (round === 2) {
				assert.strictEqual(body.messages.at(-2).tool_plan, 'Save the file.');
				assert.ok(!('content' in body.messages.at(-2)));
				assert.strictEqual(body.messages.at(-1).tool_call_id, 'call-1');
			}
			if (round === 4) {
				assert.ok(!body.tools.some((t: { function: { name: string } }) => t.function.name === 'reach_read_file'));
				assert.ok(body.tools.some((t: { function: { name: string } }) => t.function.name === 'reach_edit_file'));
				assert.match(body.messages.at(-1).content[0].document.data, /unchangedFromToolCall/);
				return nativeStream({ type: 'content-start', delta: { message: { content: { type: 'text', text: 'File saved.' } } } }, { type: 'message-end', delta: { finish_reason: 'COMPLETE' } });
			}
			assert.ok(round <= 3);
			return nativeStream(
				{ type: 'content-start', delta: { message: { content: { type: 'text', text: 'Working on the file.' } } } },
				{ type: 'content-end' },
				{ type: 'tool-plan-delta', delta: { message: { tool_plan: 'Save ' } } }, { type: 'tool-plan-delta', delta: { message: { tool_plan: 'the file.' } } },
				{ type: 'tool-call-start', delta: { message: { tool_calls: { id: 'call-' + round, function: { name: read ? 'reach_read_file' : 'reach_create_file', arguments: JSON.stringify(read ? { path: 'game.html' } : { path: 'game.html', content: 'Game' }) } } } } },
				{ type: 'tool-call-end' }, { type: 'message-end', delta: { finish_reason: 'TOOL_CALL' } });
		}, [root], model);
		agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = ['reach_create_file', 'reach_read_file', 'reach_edit_file'].map(name => ({ name }));
		const calls: string[] = [];
		store.add(agent.onDidChatProgress(e => {
			if (e.kind === 'action' && e.action.type === ActionType.ChatToolCallReady) {
				calls.push(e.action.toolCallId);
				agent.onClientToolCallComplete(chat, e.action.toolCallId, { success: true, pastTenseMessage: 'Done', content: [{ type: ToolResultContentType.Text, text: JSON.stringify({ content: 'Game' }) }] });
			}
		}));
		await agent.chats.sendMessage(chat, 'Create game.html', undefined, [], 'turn', 'editor');
		assert.strictEqual(calls.length, 3);
		assert.strictEqual(round, 4);
	});

	test('regenerates an interrupted workspace write with only its schema, then retains confirmed work on a second interruption', async () => {
		const write = { name: 'reach_create_file', inputSchema: { type: 'object' as const, properties: { path: { type: 'string' as const }, content: { type: 'string' as const } }, required: ['path', 'content'] } };
		const call = (id: string, args: string, complete: boolean) => nativeStream(
			{ type: 'tool-call-start', delta: { message: { tool_calls: { id, function: { name: write.name, arguments: args } } } } },
			...(complete ? [{ type: 'tool-call-end' }] : []), { type: 'message-end', delta: { finish_reason: complete ? 'TOOL_CALL' : 'ERROR' } });
		let requests = 0;
		const agent = await fixture(init => {
			const body = JSON.parse(String(init.body));
			switch (++requests) {
				case 1: return call('unfinished', '{"path":"game.html","content":', false);
				case 2:
					assert.strictEqual(body.strict_tools, true);
					assert.deepStrictEqual(body.tools.map((t: { function: { name: string } }) => t.function.name), [write.name]);
					assert.ok(!JSON.stringify(body.messages).includes('unfinished'));
					return call('complete', '{"path":"game.html","content":"Game"}', true);
				case 3:
					assert.strictEqual(body.messages.at(-1).tool_call_id, 'complete');
					return call('unfinished-second-file', '{"path":"style.css"', false);
				case 4:
					assert.strictEqual(body.strict_tools, true);
					return call('unfinished-again', '{}', false);
				case 5:
					assert.match(body.messages.at(-2).content, /editor confirmed changes to game.html/);
					assert.doesNotMatch(body.messages.at(-2).content, /style.css/);
					return nativeStream({ type: 'content-start', delta: { message: { content: { type: 'text', text: 'Resumed.' } } } }, { type: 'message-end', delta: { finish_reason: 'COMPLETE' } });
				default: throw new Error('Unexpected request');
			}
		}, [root], 'command-a-plus-05-2026');
		agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = [tool, write];
		const calls: string[] = [];
		store.add(agent.onDidChatProgress(e => {
			if (e.kind === 'action' && e.action.type === ActionType.ChatToolCallReady) {
				calls.push(String(e.action.toolInput));
				agent.onClientToolCallComplete(chat, e.action.toolCallId, { success: true, pastTenseMessage: 'Saved file' });
			}
		}));
		await agent.chats.sendMessage(chat, 'Create a game', undefined, [], 'first', 'editor');
		await agent.chats.sendMessage(chat, 'Continue', undefined, [], 'second', 'editor');
		assert.deepStrictEqual(calls, ['{"path":"game.html","content":"Game"}']);
		assert.strictEqual(requests, 5);
	});

	test('Command A+ receives the entire ordinary source file including errors beyond the first 60 lines', async () => {
		const content = Array.from({ length: 680 }, (_, i) => i === 309 ? 'for (let num = 1; num <= .model;' : '// source line ' + i + ' "a quoted value"').join('\n');
		assert.ok(Buffer.byteLength(JSON.stringify(JSON.stringify({ content }))) > 24_000);
		let requests = 0;
		const agent = await fixture(init => {
			if (++requests === 1) { return nativeStream(
				{ type: 'tool-call-start', delta: { message: { tool_calls: { id: 'source', function: { name: 'reach_read_file', arguments: '{"path":"game.html"}' } } } } },
				{ type: 'tool-call-end' }, { type: 'message-end', delta: { finish_reason: 'TOOL_CALL' } }); }
			const body = JSON.parse(String(init.body));
			assert.strictEqual(body.messages.at(-1).content[0].document.data, content);
			return nativeStream({ type: 'content-start', delta: { message: { content: { type: 'text', text: 'Read line 310.' } } } }, { type: 'message-end', delta: { finish_reason: 'COMPLETE' } });
		}, [root], 'command-a-plus-05-2026');
		agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = [{ name: 'reach_read_file' }];
		store.add(agent.onDidChatProgress(e => {
			if (e.kind === 'action' && e.action.type === ActionType.ChatToolCallReady) {
				agent.onClientToolCallComplete(chat, e.action.toolCallId, { success: true, pastTenseMessage: 'Read file', content: [{ type: ToolResultContentType.Text, text: JSON.stringify({ content }) }] });
			}
		}));
		await agent.chats.sendMessage(chat, 'Fix game.html', undefined, [], 'turn', 'editor');
		assert.strictEqual(requests, 2);
	});

	test('large file reads return an editable section, and rate limiting reuses the read without rerunning it', async () => {
		const content = Array.from({ length: 3500 }, (_, i) => '<p>saved puzzle line ' + i + '</p>').join('\n');
		let requests = 0, reads = 0, previous = '';
		const agent = await fixture(init => {
			const body = JSON.parse(String(init.body));
			switch (++requests) {
				case 1: return nativeStream(
					{ type: 'tool-call-start', index: 0, delta: { message: { tool_calls: { id: 'source', function: { name: 'reach_read_file', arguments: '{"path":"game.html"}' } } } } },
					{ type: 'tool-call-end', index: 0 }, { type: 'message-end', delta: { finish_reason: 'TOOL_CALL' } });
				case 2: {
					const data = body.messages.at(-1).content[0].document.data;
					assert.match(data, /Lines 1–60 of 3500/);
					assert.match(data, /fileRevision: 1:60:[0-9a-f]{40}/);
					assert.match(data, /<p>saved puzzle line 59<\/p>/);
					assert.ok(!data.includes('line 60<'));
					const spec = body.tools.find((t: { function: { name: string } }) => t.function.name === 'reach_edit_file').function.parameters;
					assert.ok(spec.properties.oldText && spec.properties.fileRevision, 'both targeted and section edits stay available');
					previous = String(init.body);
					return new Response('per-minute request limit', { status: 429, headers: { 'Retry-After': '0' } });
				}
				case 3:
					assert.strictEqual(String(init.body), previous);
					return nativeStream({ type: 'content-start', delta: { message: { content: { type: 'text', text: 'Read the source.' } } } }, { type: 'message-end', delta: { finish_reason: 'COMPLETE' } });
				default: throw new Error('Unexpected request');
			}
		}, [root], 'command-a-plus-05-2026');
		agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = [{ name: 'reach_read_file' }, { name: 'reach_edit_file', inputSchema: { type: 'object', properties: { oldText: { type: 'string' }, fileRevision: { type: 'string' } } } }];
		const status: string[] = [];
		store.add(agent.onDidChatProgress(e => {
			if (e.kind !== 'action') { return; }
			if (e.action.type === ActionType.ChatResponsePart && e.action.part.kind === 'systemNotification') { status.push(String(e.action.part.content)); }
			if (e.action.type === ActionType.ChatToolCallReady) {
				reads++;
				agent.onClientToolCallComplete(chat, e.action.toolCallId, { success: true, pastTenseMessage: 'Read source', content: [{ type: ToolResultContentType.Text, text: JSON.stringify({ bytes: content.length, content }) }] });
			}
		}));
		await agent.chats.sendMessage(chat, prompt, undefined, [], 'turn', 'editor');
		assert.deepStrictEqual([reads, requests], [1, 3]);
		assert.match(status.join(' '), /request limit/);
	});

	test('a transient 503 retries the round in place with an identical body and announced pause', async () => {
		let requests = 0, reads = 0;
		let retriedBody = '';
		const agent = await fixture(init => {
			switch (++requests) {
				case 1: return nativeStream(
					{ type: 'tool-call-start', index: 0, delta: { message: { tool_calls: { id: 'source', function: { name: 'reach_read_file', arguments: '{"path":"game.html"}' } } } } },
					{ type: 'tool-call-end', index: 0 }, { type: 'message-end', delta: { finish_reason: 'TOOL_CALL' } });
				case 2:
					retriedBody = String(init.body);
					return new Response('upstream service restarting', { status: 503 });
				case 3:
					assert.strictEqual(String(init.body), retriedBody);
					return nativeStream({ type: 'content-start', delta: { message: { content: { type: 'text', text: 'Recovered.' } } } }, { type: 'message-end', delta: { finish_reason: 'COMPLETE' } });
				default: throw new Error('Unexpected request');
			}
		}, [root], 'command-a-plus-05-2026');
		agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = [{ name: 'reach_read_file' }];
		const status: string[] = [];
		store.add(agent.onDidChatProgress(e => {
			if (e.kind !== 'action') { return; }
			if (e.action.type === ActionType.ChatResponsePart && e.action.part.kind === 'systemNotification') { status.push(String(e.action.part.content)); }
			if (e.action.type === ActionType.ChatToolCallReady) {
				reads++;
				agent.onClientToolCallComplete(chat, e.action.toolCallId, { success: true, pastTenseMessage: 'Read file', content: [{ type: ToolResultContentType.Text, text: JSON.stringify({ content: 'body' }) }] });
			}
		}));
		await agent.chats.sendMessage(chat, 'Fix game.html', undefined, [], 'turn', 'editor');
		assert.deepStrictEqual([reads, requests], [1, 3]);
		assert.match(status.join(' '), /temporarily unavailable \(HTTP 503\).*Completed tool results are preserved/);
		assert.match(JSON.stringify(await agent.chats.getMessages(chat, session)), /Recovered\./);
	});

	for (const canRecover of [true, false]) {
		test('North Mini Code keeps page diagnostics available after a failed browser tool and 422: ' + canRecover, async () => {
			const calls: string[] = [];
			let requests = 0;
			const invalid = () => Response.json({ message: 'No valid response generated. Try updating messages' }, { status: 422 });
			const agent = await fixture(init => {
				const body = JSON.parse(String(init.body));
				requests++;
				if (requests === 1 || requests === 3) {
					const tools = body.tools.map((t: { function: { name: string } }) => t.function.name);
					assert.ok(tools.includes('open_browser_page') && tools.includes('reach_read_file') && tools.includes('reach_find_tools'));
					assert.ok(!tools.includes('copilot_getNotebookSummary'));
				}
				if (requests === 1) {
					return nativeStream({ type: 'tool-call-start', delta: { message: { tool_calls: { id: 'page', function: { name: 'open_browser_page', arguments: '{}' } } } } }, { type: 'tool-call-end' }, { type: 'message-end', delta: { finish_reason: 'TOOL_CALL' } });
				}
				if (requests === 2) { return invalid(); }
				if (requests === 3) {
					assert.strictEqual(body.stream, true);
					assert.match(body.messages.at(-1).content, /no shared page/);
					return canRecover ? nativeStream({ type: 'tool-call-start', delta: { message: { tool_calls: { id: 'source', function: { name: 'reach_read_file', arguments: '{"path":"index.html"}' } } } } }, { type: 'tool-call-end' }, { type: 'message-end', delta: { finish_reason: 'TOOL_CALL' } }) : invalid();
				}
				assert.strictEqual(requests, 4);
				assert.match(body.messages.at(-1).content[0].document.data, /missing-script.js/);
				return nativeStream({ type: 'content-start', delta: { message: { content: { type: 'text', text: 'index.html references missing-script.js.' } } } }, { type: 'message-end', delta: { finish_reason: 'COMPLETE' } });
			}, [root], 'north-mini-code-1.0');
			agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = ['copilot_getNotebookSummary', 'open_browser_page', 'reach_read_file'].map(name => ({ name }));
			store.add(agent.onDidChatProgress(e => {
				if (e.kind !== 'action' || e.action.type !== ActionType.ChatToolCallReady) { return; }
				const readFile = JSON.parse(String(e.action.toolInput)).path === 'index.html';
				calls.push(readFile ? 'source' : 'page');
				agent.onClientToolCallComplete(chat, e.action.toolCallId, !readFile
					? { success: false, pastTenseMessage: 'Unavailable', error: { code: 'unavailable', message: 'no shared page' } }
					: { success: true, pastTenseMessage: 'Read', content: [{ type: ToolResultContentType.Text, text: JSON.stringify({ content: '<script src="missing-script.js"></script>' }) }] });
			}));
			const turn = agent.chats.sendMessage(chat, 'the page doesn\'t load', undefined, [], 'turn', 'editor');
			if (canRecover) { await turn; } else {
				await assert.rejects(turn, /No valid response generated/);
				assert.deepStrictEqual(await agent.chats.getMessages(chat, session), []);
			}
			assert.deepStrictEqual(calls, canRecover ? ['page', 'source'] : ['page']);
			assert.strictEqual(requests, canRecover ? 4 : 3);
		});
	}

	test('a monthly model limit is reported immediately even with a short retry header', async () => {
		let requests = 0;
		const agent = await fixture(() => {
			requests++;
			return Response.json({ message: 'You are past the per-month request limit for this model, please wait and try again later.' }, { status: 429, headers: { 'Retry-After': '0' } });
		}, [root], 'command-a-plus-05-2026');
		await assert.rejects(agent.chats.sendMessage(chat, 'Hello', undefined, [], 'turn', 'editor'), /monthly request limit for command-a-plus-05-2026/);
		assert.strictEqual(requests, 1);
		assert.deepStrictEqual(await agent.chats.getMessages(chat, session), []);
	});

	test('a continuing per-minute limit preserves the confirmed edit for the next user turn', async () => {
		let requests = 0, edits = 0;
		const agent = await fixture(init => {
			if (++requests === 1) {
				return nativeStream({ type: 'tool-call-start', delta: { message: { tool_calls: { id: 'saved', function: { name: 'reach_edit_file', arguments: '{"path":"game.html","oldText":"day","newText":"night"}' } } } } },
					{ type: 'tool-call-end' }, { type: 'message-end', delta: { finish_reason: 'TOOL_CALL' } });
			}
			// The round retries transient failures four times before failing, so a fatal run of
			// rate limiting is the initial response plus all four retries.
			if (requests < 7) { return new Response('per-minute request limit', { status: 429, headers: { 'Retry-After': '0' } }); }
			const body = JSON.parse(String(init.body));
			assert.match(body.messages.find((m: { role: string }) => m.role === 'assistant').content, /editor confirmed changes to game.html/);
			return nativeStream({ type: 'content-start', delta: { message: { content: { type: 'text', text: 'Resumed.' } } } }, { type: 'message-end', delta: { finish_reason: 'COMPLETE' } });
		}, [root], 'command-a-plus-05-2026');
		agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = [{ name: 'reach_edit_file' }];
		store.add(agent.onDidChatProgress(e => {
			if (e.kind === 'action' && e.action.type === ActionType.ChatToolCallReady) {
				edits++;
				agent.onClientToolCallComplete(chat, e.action.toolCallId, { success: true, pastTenseMessage: 'Saved file' });
			}
		}));
		await agent.chats.sendMessage(chat, 'Change game colors', undefined, [], 'first', 'editor');
		await agent.chats.sendMessage(chat, 'Continue', undefined, [], 'second', 'editor');
		assert.deepStrictEqual([edits, requests], [1, 7]);
	});

	test('slow tool generation reports progress and only times out after 90 seconds without events', async () => {
		const clock = sinon.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout'] });
		try {
			let stream!: ReadableStreamDefaultController<Uint8Array>;
			const statuses: string[] = [];
			const result = fetchCohereStream(async signal => new Response(new ReadableStream({
				start(controller) {
					stream = controller;
					signal.addEventListener('abort', () => controller.error(signal.reason), { once: true });
				},
			})), new AbortController().signal, () => assert.fail('unexpected text'), msg => statuses.push(msg));
			const failure = assert.rejects(result, /90 seconds/);
			const send = (event: unknown) => stream.enqueue(new TextEncoder().encode('data: ' + JSON.stringify(event) + '\n\n'));
			send({ type: 'tool-call-start', delta: { message: { tool_calls: { id: 'edit', function: { name: 'reach_edit_file' } } } } });
			await clock.tickAsync(80_000);
			assert.deepStrictEqual(statuses, ['Cohere is preparing reach_edit_file.']);
			send({ type: 'tool-call-delta', delta: { message: { tool_calls: { function: { arguments: '{' } } } } });
			await clock.tickAsync(90_000);
			await failure;
		} finally { clock.restore(); }
	});

	test('Command A+ continues from inspection to editing after a rejected generation and streams its final answer', async () => {
		const call = (name: string, id: string) => nativeStream(
			{ type: 'tool-call-start', index: 0, delta: { message: { tool_calls: { id, function: { name, arguments: '' } } } } },
			{ type: 'tool-call-delta', index: 0, delta: { message: { tool_calls: { function: { arguments: '{"path":' } } } } },
			{ type: 'tool-call-delta', index: 0, delta: { message: { tool_calls: { function: { arguments: '"index.html"}' } } } } },
			{ type: 'tool-call-end', index: 0 }, { type: 'message-end', delta: { finish_reason: 'TOOL_CALL' } },
		);
		let requests = 0;
		const agent = await fixture(init => {
			const body = JSON.parse(String(init.body));
			switch (++requests) {
				case 1: return call('reach_read_file', 'read');
				case 2: {
					const result = body.messages.at(-1);
					assert.strictEqual(result.tool_call_id, 'read');
					assert.strictEqual(result.content[0].document.data, '<p id="demo">\nIndex source\n</p>');
					return Response.json({ message: 'No valid response generated' }, { status: 422 });
				}
				case 3:
					assert.ok(body.tools.some((t: { function: { name: string } }) => t.function.name === 'reach_edit_file'));
					assert.match(body.messages.at(-1).content, /Index source/);
					return call('reach_edit_file', 'edit');
				case 4:
					assert.strictEqual(body.messages.at(-1).tool_call_id, 'edit');
					return nativeStream(
						{ type: 'content-start', index: 0, delta: { message: { content: { type: 'thinking', thinking: 'private block' } } } },
						{ type: 'content-start', index: 1, delta: { message: { content: { type: 'text', text: '' } } } },
						{ type: 'content-delta', index: 1, delta: { message: { content: { text: 'Saved dark mode.' } } } },
						{ type: 'message-end', delta: { finish_reason: 'COMPLETE' } },
					);
				default: throw new Error('Unexpected request');
			}
		}, [root], 'command-a-plus-05-2026');
		agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = [{ name: 'reach_read_file' }, { name: 'reach_edit_file' }];
		const calls: string[] = [], deltas: string[] = [];
		store.add(agent.onDidChatProgress(event => {
			if (event.kind !== 'action') { return; }
			if (event.action.type === ActionType.ChatDelta) { deltas.push(event.action.content); }
			if (event.action.type === ActionType.ChatToolCallStart) { calls.push(event.action.toolName); }
			if (event.action.type === ActionType.ChatToolCallReady) {
				assert.deepStrictEqual(JSON.parse(String(event.action.toolInput)), { path: 'index.html' });
				agent.onClientToolCallComplete(chat, event.action.toolCallId, calls.length === 1
					? { success: true, pastTenseMessage: 'Read index.html', content: [{ type: ToolResultContentType.Text, text: JSON.stringify({ uri: 'file:///index.html', content: '<p id="demo">\nIndex source\n</p>' }) }] }
					: { success: true, pastTenseMessage: 'Saved edit' });
			}
		}));
		await agent.chats.sendMessage(chat, 'Add dark mode', undefined, [], 'turn', 'editor');
		assert.deepStrictEqual({ calls, deltas, requests }, { calls: ['reach_read_file', 'reach_edit_file'], deltas: ['Saved dark mode.'], requests: 4 });
	});

	test('Command A+ retries generation with the editor receipt when a later tool stream ends in ERROR', async () => {
		let requests = 0;
		const start = { type: 'tool-call-start', index: 0, delta: { message: { tool_calls: { id: 'read', function: { name: 'workspace_read', arguments: '{}' } } } } };
		const agent = await fixture(init => {
			switch (++requests) {
				case 1: return nativeStream(start, { type: 'tool-call-end', index: 0 }, { type: 'message-end', delta: { finish_reason: 'TOOL_CALL' } });
				case 2:
					return nativeStream(start, { type: 'message-end', delta: { finish_reason: 'ERROR' } });
				case 3: {
					const body = JSON.parse(String(init.body));
					assert.strictEqual(body.stream, false);
					assert.deepStrictEqual(body.messages.map((m: { role: string }) => m.role), ['system', 'user']);
					assert.match(body.messages.at(-1).content, /Read file/);
					return Response.json({ message: { content: [{ type: 'text', text: '{"reply":"Finished."}' }] }, finish_reason: 'COMPLETE' });
				}
				default: throw new Error('Unexpected request');
			}
		}, [root], 'command-a-plus-05-2026');
		let executed = 0;
		store.add(agent.onDidChatProgress(event => {
			if (event.kind === 'action' && event.action.type === ActionType.ChatToolCallReady) {
				executed++;
				agent.onClientToolCallComplete(chat, event.action.toolCallId, { success: true, pastTenseMessage: 'Read file' });
			}
		}));
		await agent.chats.sendMessage(chat, prompt, undefined, [], 'turn', 'editor');
		assert.deepStrictEqual([requests, executed], [3, 1]);
	});

	for (const badArgs of ['{"path":', '[]', '"{}"']) {
		test('Command A+ regenerates a mixed batch before dispatch or history replay for ' + badArgs, async () => {
			const call = (id: string, index: number, args: string) => [
				{ type: 'tool-call-start', index, delta: { message: { tool_calls: { id, function: { name: 'workspace_read', arguments: args } } } } },
				{ type: 'tool-call-end', index },
			];
			const end = { type: 'message-end', delta: { finish_reason: 'TOOL_CALL' } };
			let requests = 0, executed = 0;
			const agent = await fixture(init => {
				const body = JSON.parse(String(init.body));
				switch (++requests) {
					case 1: return nativeStream(...call('first', 0, '{}'), ...call('malformed', 1, badArgs), end);
					case 2:
						assert.strictEqual(executed, 0);
						assert.ok(body.messages.every((m: { role: string }) => m.role === 'system' || m.role === 'user'));
						assert.match(body.messages[0].content, /one tool call at a time/);
						return nativeStream(...call('corrected', 0, '{"path":"game.html"}'), end);
					case 3:
						assert.deepStrictEqual(body.messages.at(-2).tool_calls.map((c: { id: string }) => c.id), ['corrected']);
						assert.strictEqual(body.messages.at(-1).tool_call_id, 'corrected');
						return nativeStream({ type: 'content-start', delta: { message: { content: { type: 'text', text: 'Finished.' } } } }, { type: 'message-end', delta: { finish_reason: 'COMPLETE' } });
					default: throw new Error('Unexpected retry');
				}
			}, [root], 'command-a-plus-05-2026');
			store.add(agent.onDidChatProgress(event => {
				if (event.kind === 'action' && event.action.type === ActionType.ChatToolCallReady) {
					executed++;
					agent.onClientToolCallComplete(chat, event.action.toolCallId, { success: true, pastTenseMessage: 'Read file' });
				}
			}));
			await agent.chats.sendMessage(chat, prompt, undefined, [], 'turn', 'editor');
			assert.deepStrictEqual([requests, executed], [3, 1]);
		});
	}

	test('Command A+ does not dispatch a streamed tool without its end event', async () => {
		const agent = await fixture(() => nativeStream(
			{ type: 'tool-call-start', index: 0, delta: { message: { tool_calls: { id: 'partial', function: { name: 'workspace_read', arguments: '{}' } } } } },
			{ type: 'message-end', delta: { finish_reason: 'TOOL_CALL' } },
		), [root], 'command-a-plus-05-2026');
		const calls: string[] = [];
		store.add(agent.onDidChatProgress(e => { if (e.kind === 'action' && e.action.type === ActionType.ChatToolCallStart) { calls.push(e.action.toolName); } }));
		await assert.rejects(agent.chats.sendMessage(chat, prompt, undefined, [], 'turn', 'editor'), /incomplete response/);
		assert.deepStrictEqual(calls, []);
	});
});
