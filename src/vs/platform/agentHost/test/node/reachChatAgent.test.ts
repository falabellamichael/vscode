/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import assert from 'assert';
import * as sinon from 'sinon';
import { mkdir, mkdtemp, rm, writeFile } from 'fs/promises';
import { tmpdir } from 'os';
import { DeferredPromise } from '../../../../base/common/async.js';
import { VSBuffer } from '../../../../base/common/buffer.js';
import { isCancellationError } from '../../../../base/common/errors.js';
import { join } from '../../../../base/common/path.js';
import { URI } from '../../../../base/common/uri.js';
import { ensureNoDisposablesAreLeakedInTestSuite } from '../../../../base/test/common/utils.js';
import { REACH_ACCOUNT_ORIGIN } from '../../../native/common/native.js';
import type { IFileService } from '../../../files/common/files.js';
import { reachCompletionRetryDelaysMs, requestReachAssistant } from '../../../native/common/reachChatTransport.js';
import { resolveAgentChatContext, type AgentSignal } from '../../common/agent.js';
import { ActionType } from '../../common/state/sessionActions.js';
import { MessageAttachmentKind, MessageKind, ResponsePartKind, ToolCallConfirmationReason, ToolCallStatus, ToolResultContentType, TurnState } from '../../common/state/sessionState.js';
import { ReachChatAgent } from '../../node/reach/reachChatAgent.js';
import { findReachTools, ReachChatToolCatalog } from '../../node/reach/reachChatToolCatalog.js';

suite('REACH regular chat provider', () => {
	const store = ensureNoDisposablesAreLeakedInTestSuite();
	const chat = URI.parse('ahp-chat:/reach-regular-test');
	const session = URI.parse('reach:/test');
	// Retries must be near-instant in tests: production delays are tuned to outlast gateway blips.
	const productionRetryDelays = [...reachCompletionRetryDelaysMs];
	setup(() => { reachCompletionRetryDelaysMs.splice(0, reachCompletionRetryDelaysMs.length, 1); });
	teardown(() => { reachCompletionRetryDelaysMs.splice(0, reachCompletionRetryDelaysMs.length, ...productionRetryDelays); });

	function fixture(supportsTools = true, files: Pick<IFileService, 'readFile'> = { readFile: () => { throw new Error('unexpected file read'); } }) {
		let active = true, remaining = 5;
		let completion: ((options: RequestInit) => Promise<Response>) | undefined;
		let accountGate: Promise<void> | undefined;
		let catalogGate: Promise<void> | undefined;
		let accountReads = 0;
		let eligible = true;
		let vision = false;
		const requests: { model: string; messages: { role: string; content: string; tool_calls?: unknown[]; tool_call_id?: string }[]; tools?: unknown[] }[] = [];
		const changes: string[] = [];
		const events: AgentSignal[] = [];
		const agent = store.add(new ReachChatAgent(files, async (resource, options = {}) => {
			const token = new Headers(options.headers).get('Authorization');
			const path = new URL(resource.toString()).pathname;
			assert.equal(new URL(resource.toString()).origin, REACH_ACCOUNT_ORIGIN);
			assert.equal(options.redirect, 'error');
			if (path === '/v1/account') {
				accountReads++;
				await accountGate;
				return Response.json({ id: token === 'Bearer token2' ? 'second-account' : 'first-account', access: { activeSubscription: active }, requestAllowance: { remaining } });
			}
			if (path === '/v1/models') {
				await catalogGate;
				return Response.json({ data: [
					...(eligible ? [{ id: 'chat1', name: 'First', access: 'requests', allow_tools: supportsTools, capabilities: { context_window: 65536 } }, { id: 'chat2', access: 'requests', capabilities: { vision } }] : []),
					{ id: 'wallet', access: 'paid' }, { id: 'vectors', access: 'requests', capabilities: { chat: false, endpoint: '/v1/embeddings' } },
				] });
			}
			assert.equal(path, '/v1/chat/completions');
			assert.equal(options.method, 'POST');
			requests.push(JSON.parse(String(options.body)));
			return completion ? completion(options) : Response.json({ choices: [{ message: { role: 'assistant', content: 'answer' } }] });
		}));
		store.add(agent.onDidChatProgress(e => events.push(e)));
		store.add(agent.onDidChangeChatData(e => changes.push(e.providerData)));
		return { agent, requests, events, changes, setActive: (value: boolean) => active = value,
			setVision: (value: boolean) => vision = value,
			setRemaining: (value: number) => remaining = value, setEligible: (value: boolean) => eligible = value,
			setCompletion: (value: typeof completion) => completion = value,
			holdAccount: (gate: Promise<void>) => accountGate = gate, holdCatalog: (gate: Promise<void>) => catalogGate = gate, accountReads: () => accountReads };
	}

	const png = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+kWQAAAABJRU5ErkJggg==';
	const embedded = { type: MessageAttachmentKind.EmbeddedResource as const, contentType: 'image/png', label: 'clipboard.png', data: png };

	test('Auto sends subscription vision inputs through tool rounds and restores the original image context', async () => {
		const readFile = sinon.stub().resolves({ value: VSBuffer.wrap(Buffer.from(png, 'base64')) });
		const { agent, changes, requests, setVision, setCompletion } = fixture(true, { readFile });
		setVision(true);
		await agent.authenticate(REACH_ACCOUNT_ORIGIN, 'token1');
		assert.deepStrictEqual(agent.models.get().map(model => [model.id, model.supportsVision]), [['auto', true], ['chat1', false], ['chat2', true]]);
		await agent.chats.createChat(chat, session, { model: { id: 'auto' } });
		agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = [{ name: 'extra_3' }];
		store.add(agent.onDidChatProgress(e => {
			if (e.kind === 'action' && e.action.type === ActionType.ChatToolCallReady) {
				agent.onClientToolCallComplete(chat, e.action.toolCallId, { success: true, pastTenseMessage: 'Done' });
			}
		}));
		setCompletion(async () => requests.length === 1
			? Response.json({ choices: [{ message: { tool_calls: [{ id: 'probe', function: { name: 'extra_3', arguments: '{}' } }] } }] })
			: Response.json({ choices: [{ message: { content: 'answer' } }] }));
		const local = { type: MessageAttachmentKind.Resource as const, label: 'screen.png', uri: URI.file(join(tmpdir(), 'screen.png')).toString() };
		await agent.chats.sendMessage(chat, 'Compare them.', undefined, [local, embedded], 'vision', 'editor');
		const pixel = { type: 'image_url', image_url: { url: 'data:image/png;base64,' + png } };
		const original = [{ type: 'text', text: 'Compare them.' }, pixel, pixel];
		assert.strictEqual(requests.length, 2);
		for (const request of requests) {
			assert.strictEqual(request.model, 'chat2');
			assert.deepStrictEqual(request.messages.find(message => message.role === 'user')?.content, original);
		}
		const data = changes.at(-1)!;
		await agent.chats.releaseChat(chat, session);
		await agent.materializeChat(chat, session, data);
		assert.deepStrictEqual((await agent.chats.getMessages(chat, session))[0].message.attachments, [local, embedded]);
		await agent.chats.sendMessage(chat, 'What color?', undefined, [], 'followup');
		assert.deepStrictEqual(requests.at(-1)?.messages[1].content, original);
		readFile.resolves({ value: VSBuffer.wrap(Buffer.concat([Buffer.from(png, 'base64'), Buffer.from('changed')])) });
		await agent.chats.sendMessage(chat, 'Look again.', undefined, [], 'changed');
		assert.deepStrictEqual(requests.at(-1)?.messages[1].content, original.slice(0, 2));
		await agent.chats.changeModel(chat, { id: 'chat1' }, session);
		await agent.chats.sendMessage(chat, 'Use text.', undefined, [], 'text');
		assert.strictEqual(requests.at(-1)?.model, 'chat1');
		assert.strictEqual(requests.at(-1)?.messages[1].content, 'Compare them.');
	});

	test('rejects text-only subscription models, unsupported image bytes and oversized attachment counts before inference', async () => {
		const { agent, requests, setVision } = fixture();
		setVision(true);
		await agent.authenticate(REACH_ACCOUNT_ORIGIN, 'token1');
		await agent.chats.createChat(chat, session, { model: { id: 'chat1' } });
		await assert.rejects(agent.chats.sendMessage(chat, 'Look', undefined, [embedded], 'text'), /model that accepts images/);
		await agent.chats.changeModel(chat, { id: 'chat2' }, session);
		await assert.rejects(agent.chats.sendMessage(chat, 'Look', undefined, [{ ...embedded, data: 'bm90IHBpeGVscw==' }], 'invalid'), /Check the format/);
		await assert.rejects(agent.chats.sendMessage(chat, 'Look', undefined, Array(5).fill(embedded), 'many'), /4 images/);
		setVision(false);
		await assert.rejects(agent.chats.sendMessage(chat, 'Look', undefined, [embedded], 'revoked'), /model that accepts images/);
		assert.strictEqual(requests.length, 0);
	});

	test('keeps attached folders in REACH history and resolves their paths against the correct workspace', async () => {
		const { agent, requests, changes } = fixture();
		await agent.authenticate(REACH_ACCOUNT_ORIGIN, 'token1');
		const root = URI.file(join(tmpdir(), 'puzzle'));
		const nestedRoot = URI.joinPath(root, 'versions');
		const folder = { type: MessageAttachmentKind.Resource as const, displayKind: 'directory', label: 'NMC-Sudoku', uri: URI.joinPath(nestedRoot, 'NMC-Sudoku').toString() };
		await agent.chats.createChat(chat, session, { workingDirectories: [root, nestedRoot] });
		await agent.chats.sendMessage(chat, 'create your own sudoku version in this folder', [root, nestedRoot], [folder], 'turn', 'editor');
		assert.match(requests[0].messages.at(-1)!.content, /"path":"NMC-Sudoku","workspaceFolder":"versions"/);
		await agent.chats.releaseChat(chat, session);
		await agent.materializeChat(chat, session, changes.at(-1)!);
		assert.deepStrictEqual((await agent.chats.getMessages(chat, session))[0].message.attachments, [folder]);
		await agent.chats.sendMessage(chat, 'continue', [root, nestedRoot], [], 'resume', 'editor');
		assert.match(requests[1].messages.at(-1)!.content, /"path":"NMC-Sudoku","workspaceFolder":"versions"/);
	});

	test('offers All Access per chat and clamps unknown or policy-restricted approvals to default', async () => {
		const policy = sinon.stub().returns(false);
		const agent = store.add(new ReachChatAgent({ readFile: () => { throw new Error('unexpected file read'); } }, undefined, { getRootValue: policy }));
		const resolve = (value?: string) => agent.resolveChatConfig({ config: value ? { autoApprove: value } : {} });
		const initial = await resolve();
		assert.deepStrictEqual(initial.schema.properties.autoApprove.enumLabels, ['Default Approvals', 'All Access']);
		assert.strictEqual(initial.schema.properties.autoApprove.chatMutable, true);
		assert.strictEqual(initial.values.autoApprove, 'default');
		assert.strictEqual((await resolve('autoApprove')).values.autoApprove, 'autoApprove');
		assert.strictEqual((await resolve('unknown')).values.autoApprove, 'default');
		policy.returns(true);
		const restricted = await resolve('autoApprove');
		assert.deepStrictEqual([restricted.values.autoApprove, restricted.schema.properties.autoApprove.enum], ['default', ['default']]);
	});

	test('all enabled tools are discoverable and colliding wire names dispatch to the correct registry ID', () => {
		const all = Array.from({ length: 47 }, (_, i) => ({ name: 'provider_' + i, description: 'fixture' }));
		all.push({ name: 'calendar/lookup', description: 'calendar' }, { name: 'calendar.lookup', description: 'calendar' });
		const catalog = new ReachChatToolCatalog(all, 'show appointments');
		assert.ok(catalog.specs().some(tool => tool.name === findReachTools));
		const found = new Set<string>();
		let offset: number | undefined = 0;
		do {
			const page = catalog.selectPage('', offset);
			assert.ok(catalog.specs().length <= 32);
			for (const name of page.tools) { found.add(catalog.originalName(name)!); }
			offset = page.nextOffset;
		} while (offset !== undefined);
		assert.deepStrictEqual(found, new Set(all.map(tool => tool.name)));
		assert.equal(catalog.selectPage('calendar', 0).total, 2);
	});

	test('a workspace file request starts with local tools and can search for other enabled tools', () => {
		const other = Array.from({ length: 20 }, (_, i) => ({ name: 'github_' + i, description: 'create a new file in this folder' }));
		const catalog = new ReachChatToolCatalog([...other, { name: 'reach_read_file' }, { name: 'reach_create_file' }], 'create a new game in this folder', { hasWorkspace: true });
		assert.deepStrictEqual(catalog.specs().map(t => t.name), ['reach_create_file', 'reach_read_file', findReachTools]);
		catalog.selectPage('github_19', 0);
		assert.ok(catalog.specs().some(t => t.name === 'github_19'));
		const changed = catalog.selectPage('reach_create_file write new file create_file', 64);
		assert.strictEqual(changed.tools[0], 'reach_create_file');
		assert.strictEqual(catalog.specs()[0].name, 'reach_create_file');
		const build = new ReachChatToolCatalog([{ name: 'run_in_terminal' }, { name: 'reach_create_file' }, { name: 'reach_read_file' }], 'create a new game', { hasWorkspace: true, preferEditorForNewProject: true });
		assert.deepStrictEqual(build.selectPage('run_in_terminal', 16).tools, ['run_in_terminal']);
		assert.deepStrictEqual(build.specs().map(t => t.name), ['run_in_terminal', 'reach_create_file', 'reach_read_file']);
	});

	test('reports the model and HTTP status when its gateway rejects a fresh chat request', async () => {
		const { agent, requests, setCompletion } = fixture();
		await agent.authenticate(REACH_ACCOUNT_ORIGIN, 'token1');
		await agent.chats.createChat(chat, session, { model: { id: 'chat2' } });
		setCompletion(async () => Response.json({ error: { code: 'gateway_error', message: 'opaque upstream detail' } }, { status: 502 }));
		await assert.rejects(agent.chats.sendMessage(chat, 'Hello', undefined, [], 'fresh'), /chat2.*HTTP 502/);
		// The test-suite retry delays allow exactly one retry before the failure surfaces.
		assert.strictEqual(requests.length, 2);
	});

	test('survives a transient 503 with one idempotent retry and announces the pause', async () => {
		const { agent, requests, setCompletion } = fixture();
		await agent.authenticate(REACH_ACCOUNT_ORIGIN, 'token1');
		await agent.chats.createChat(chat, session, { model: { id: 'chat2' } });
		const keys: (string | null)[] = [];
		setCompletion(async init => {
			keys.push(new Headers(init.headers).get('Idempotency-Key'));
			return keys.length === 1
				? Response.json({ error: { code: 'gateway_error', message: 'upstream restarting' } }, { status: 503 })
				: Response.json({ choices: [{ message: { content: 'Recovered after the blip.' } }] });
		});
		const notifications: string[] = [];
		store.add(agent.onDidChatProgress(e => {
			if (e.kind === 'action' && e.action.type === ActionType.ChatResponsePart && e.action.part.kind === ResponsePartKind.SystemNotification) {
				notifications.push(String(e.action.part.content));
			}
		}));
		await agent.chats.sendMessage(chat, 'Hello', undefined, [], 'blip', 'editor');
		assert.strictEqual(requests.length, 2);
		assert.strictEqual(keys[0], keys[1]);
		assert.deepStrictEqual(requests[1], requests[0]);
		assert.ok(notifications.some(content => /temporarily unavailable \(HTTP 503\).*continues automatically/.test(content)), 'retry pause is announced: ' + JSON.stringify(notifications));
		assert.match(JSON.stringify(await agent.chats.getMessages(chat, session)), /Recovered after the blip\./);
	});

	for (const recovers of [true, false]) {
		test(`retries a failed first generation exactly once without executing its partial calls (recovers=${recovers})`, async () => {
			const { agent, requests, events, setCompletion } = fixture();
			await agent.authenticate(REACH_ACCOUNT_ORIGIN, 'token1');
			await agent.chats.createChat(chat, session, { model: { id: 'chat2' } });
			agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = [{ name: 'write_note' }];
			const keys: (string | null)[] = [];
			const error = { error: { code: 'provider_stream_error', message: 'Model stream interrupted' } };
			setCompletion(async init => {
				keys.push(new Headers(init.headers).get('Idempotency-Key'));
				if (keys.length === 1) {
					const partial = { choices: [{ delta: { tool_calls: [{ index: 0, id: 'partial', function: { name: 'write_note', arguments: '{' } }] } }] };
					return new Response(`data: ${JSON.stringify(partial)}\n\ndata: ${JSON.stringify(error)}\n\ndata: [DONE]\n\n`, { headers: { 'content-type': 'text/event-stream' } });
				}
				return recovers ? Response.json({ choices: [{ message: { content: 'Recovered.' } }] }) : Response.json(error, { status: 502 });
			});
			const send = agent.chats.sendMessage(chat, 'Hello', undefined, [], 'recover-first', 'editor');
			if (recovers) {
				await send;
				assert.match(JSON.stringify(await agent.chats.getMessages(chat, session)), /Recovered\./);
				assert.strictEqual(requests.length, 2);
			} else {
				await assert.rejects(send, /chat2 stopped generating during both attempts/);
				// The generation-level retry's 502 is itself retried once by the transport
				// before the failure surfaces, so the failing path makes one extra request.
				assert.strictEqual(requests.length, 3);
				assert.strictEqual(keys[2], keys[1]);
			}
			assert.deepStrictEqual(requests[1], requests[0]);
			assert.strictEqual(keys[0], 'recover-first');
			assert.ok(keys[1] && keys[1] !== keys[0]);
			assert.ok(!events.some(e => e.kind === 'action' && e.action.type === ActionType.ChatToolCallStart));
		});
	}

	test('a named project streams the workspace writer and reports its generation before executing it', async () => {
		const { agent, setCompletion, accountReads } = fixture();
		const root = URI.file(join(tmpdir(), 'reach-project'));
		await agent.authenticate(REACH_ACCOUNT_ORIGIN, 'token1');
		await agent.chats.createChat(chat, session, { workingDirectories: [root] });
		agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = ['copilot_memory', 'run_in_terminal', 'reach_create_file', 'reach_read_file'].map(name => ({ name }));
		const payloads: string[][] = [];
		setCompletion(async init => {
			const body = JSON.parse(String(init.body));
			payloads.push(body.tools.map((tool: { function: { name: string } }) => tool.function.name));
			assert.strictEqual(body.stream, true);
			if (payloads.length > 1) { return Response.json({ choices: [{ message: { content: 'Saved puzzle.py.' } }] }); }
			const deltas = [
				{ delta: { reasoning_content: 'fixture visible reasoning' } },
				{ delta: { reasoning_content: 'x'.repeat(250) } },
				{ delta: { reasoning_content: ' reasoning tail' } },
				{ delta: { tool_calls: [{ index: 0, id: 'save', function: { name: 'reach_create_file', arguments: '{"path":"puzzle.py",' } }] } },
				{ delta: { tool_calls: [{ index: 0, function: { arguments: '"content":"print(42)"}' } }] }, finish_reason: 'tool_calls' },
			];
			return new Response(deltas.map(choice => 'data: ' + JSON.stringify({ choices: [choice] }) + '\n\n').join(''), { headers: { 'Content-Type': 'text/event-stream' } });
		});
		const progress: string[] = [];
		let reasoningText = '';
		store.add(agent.onDidChatProgress(e => {
			if (e.kind === 'action' && e.action.type === ActionType.ChatResponsePart && e.action.part.kind === ResponsePartKind.Reasoning) {
				assert.strictEqual(e.action.part.content, 'The model is thinking.');
				progress.push('thinking');
			}
			if (e.kind === 'action' && e.action.type === ActionType.ChatReasoning) {
				reasoningText += e.action.content;
				if (typeof e.action.duration === 'number') { progress.push('finished thinking'); }
			}
			if (e.kind === 'action' && e.action.type === ActionType.ChatResponsePart && e.action.part.kind === ResponsePartKind.SystemNotification) {
				assert.strictEqual(typeof e.action.part.content, 'string');
				progress.push(String(e.action.part.content));
			}
			if (e.kind === 'action' && e.action.type === ActionType.ChatToolCallReady) {
				progress.push('executed');
				agent.onClientToolCallComplete(chat, e.action.toolCallId, { success: true, pastTenseMessage: 'Saved puzzle.py' });
			}
		}));
		await agent.chats.sendMessage(chat, 'Create your own version of sudoku named "Sudokus-Edge"', undefined, [], 'build', 'editor');
		assert.strictEqual(payloads.length, 2);
		assert.deepStrictEqual(progress, ['thinking', 'finished thinking', 'Preparing reach_create_file…', 'executed']);
		assert.strictEqual(reasoningText, '\n\n' + 'fixture visible reasoning' + 'x'.repeat(250) + ' reasoning tail');
		assert.deepStrictEqual(payloads[0], ['reach_create_file', 'reach_read_file', findReachTools]);
		assert.deepStrictEqual(payloads[1], ['reach_create_file', 'reach_read_file', 'run_in_terminal', findReachTools]);
		assert.strictEqual(accountReads(), 2); // One sign-in, one preflight for the entire two-round turn.
	});

	test('keeps workspace file tools available after a catalog search while finishing a project', async () => {
		const { agent, requests, setCompletion } = fixture();
		await agent.authenticate(REACH_ACCOUNT_ORIGIN, 'token1');
		const root = URI.file(join(tmpdir(), 'sudoku'));
		await agent.chats.createChat(chat, session, { workingDirectories: [root] });
		agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = [
			...Array.from({ length: 20 }, (_, i) => ({ name: 'extra_' + i })),
			...['reach_create_file', 'reach_read_file', 'reach_edit_file', 'run_in_terminal'].map(name => ({ name })),
		];
		const writes: unknown[] = [];
		store.add(agent.onDidChatProgress(e => {
			if (e.kind === 'action' && e.action.type === ActionType.ChatToolCallReady) {
				writes.push(JSON.parse(String(e.action.toolInput)));
				agent.onClientToolCallComplete(chat, e.action.toolCallId, { success: true, pastTenseMessage: 'Saved index.html' });
			}
		}));
		let searchedTools: string[] = [];
		setCompletion(async options => {
			const body = JSON.parse(String(options.body));
			const names = body.tools.map((tool: { function: { name: string } }) => tool.function.name);
			if (requests.length === 1) {
				return Response.json({ choices: [{ message: { tool_calls: [{ id: 'search', function: { name: findReachTools, arguments: '{"query":"terminal"}' } }] } }] });
			}
			if (requests.length === 2) {
				searchedTools = names;
				return Response.json({ choices: [{ message: { tool_calls: [{ id: 'write', function: { name: 'reach_create_file', arguments: '{"path":"index.html","content":"<html></html>"}' } }] } }] });
			}
			return Response.json({ choices: [{ message: { content: 'Finished index.html.' } }] });
		});
		const folder = { type: MessageAttachmentKind.Resource as const, displayKind: 'directory', label: 'sudoku', uri: root.toString() };
		await agent.chats.sendMessage(chat, 'finish this sudoku', root, [folder], 'finish-project', 'editor');
		assert.ok(searchedTools.includes('run_in_terminal'));
		assert.ok(searchedTools.includes('reach_create_file'));
		assert.ok(searchedTools.includes('reach_edit_file'));
		assert.ok(searchedTools.includes('reach_read_file'));
		assert.deepStrictEqual(writes, [{ path: 'index.html', content: '<html></html>' }]);
	});

	test('executes an enabled exact tool name omitted from the current schema page', async () => {
		const { agent, setCompletion } = fixture();
		await agent.authenticate(REACH_ACCOUNT_ORIGIN, 'token1');
		await agent.chats.createChat(chat, session);
		agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = Array.from({ length: 20 }, (_, i) => ({ name: 'extra_' + i }));
		const calls: unknown[] = [];
		store.add(agent.onDidChatProgress(e => {
			if (e.kind === 'action' && e.action.type === ActionType.ChatToolCallReady) {
				calls.push(JSON.parse(String(e.action.toolInput)));
				agent.onClientToolCallComplete(chat, e.action.toolCallId, { success: true, pastTenseMessage: 'Done' });
			}
		}));
		setCompletion(async () => calls.length === 0
			? Response.json({ choices: [{ message: { tool_calls: [{ id: 'hidden', function: { name: 'extra_19', arguments: '{"receipt":true}' } }] } }] })
			: Response.json({ choices: [{ message: { content: 'Done.' } }] }));
		await agent.chats.sendMessage(chat, 'Use the hidden enabled tool', undefined, [], 'hidden-enabled', 'editor');
		assert.deepStrictEqual(calls, [{ receipt: true }]);
	});

	test('continue carries the previous editor result into inference without executing its command again', async () => {
		const { agent, requests } = fixture();
		const root = URI.file(join(tmpdir(), 'reach-project'));
		await agent.authenticate(REACH_ACCOUNT_ORIGIN, 'token1');
		await agent.chats.createChat(chat, session, { workingDirectories: [root] });
		await agent.chats.sendMessage(chat, 'continue', [root], [], 'resume', 'editor', undefined, {
			...resolveAgentChatContext(session, chat),
			previousTurns: [{ id: 'build', state: TurnState.Cancelled, usage: undefined, message: { text: 'Build a game', origin: { kind: MessageKind.User } }, responseParts: [{
				kind: ResponsePartKind.ToolCall, toolCall: { toolCallId: 'checked', toolName: 'run_in_terminal', displayName: 'Terminal', invocationMessage: 'List files',
					toolInput: '{"command":"dir"}', status: ToolCallStatus.Completed, confirmed: ToolCallConfirmationReason.UserAction, success: true,
					pastTenseMessage: 'Listed files', content: [{ type: ToolResultContentType.Text, text: 'existing.html' }] },
			}] }],
		});
		const history = requests.at(-1)!.messages;
		const call = history.findIndex((message: { role: string }) => message.role === 'assistant');
		assert.ok(call >= 0);
		assert.deepStrictEqual(history[call].tool_calls, [{ id: 'checked', type: 'function', function: { name: 'run_in_terminal', arguments: '{"command":"dir"}' } }]);
		assert.strictEqual(history[call + 1].tool_call_id, 'checked');
		assert.match(history[call + 1].content, /existing\.html/);
	});

	test('regular chat generation follows the turn cancellation signal through a long request', async () => {
		const controller = new AbortController();
		const started = new DeferredPromise<AbortSignal | null | undefined>();
		const result = requestReachAssistant(async (url, options) => {
			switch (new URL(url).pathname) {
				case '/v1/account': return Response.json({ id: 'owner', access: { activeSubscription: true }, requestAllowance: { unlimited: true } });
				case '/v1/models': return Response.json({ data: [{ id: 'chat', access: 'requests' }] });
				default:
					started.complete(options.signal);
					return new Promise((_resolve, reject) => options.signal?.addEventListener('abort', () => reject(options.signal?.reason), { once: true }));
			}
		}, 'owner', 'fixture', { messages: [{ role: 'user', content: 'Finish this file' }] }, 'long-turn', controller.signal, undefined, 'chat');
		const generationSignal = await started.p;
		controller.abort();
		assert.deepStrictEqual(await result, { kind: 'stopped' });
		assert.strictEqual(generationSignal, controller.signal);
	});

	test('requests room for Inkling output on every chat round of routes with automatic output', async () => {
		const budgets: (number | undefined)[] = [];
		for (const [automaticOutput, write, id] of [[true, true, 'inkling-2-mini'], [false, true, 'inkling-2-mini'], [true, false, 'inkling-2-mini'], [true, true, 'inkling-2-mini-hermes']]) {
			await requestReachAssistant(async (url, options) => {
				switch (new URL(url).pathname) {
					case '/v1/account': return Response.json({ id: 'owner', access: { activeSubscription: true }, requestAllowance: { unlimited: true } });
					case '/v1/models': return Response.json({ data: [{ id, access: 'requests', capabilities: { context_window: 131072, automaticOutput } }] });
					default:
						budgets.push(JSON.parse(String(options.body)).max_tokens);
						return Response.json({ choices: [{ message: { content: 'Done.' } }] });
				}
			}, 'owner', 'fixture', { messages: [{ role: 'user', content: 'Continue' }], tools: [{ name: write ? 'reach_edit_file' : 'reach_read_file', description: '' }] },
			'test', new AbortController().signal, undefined, 'chat');
		}
		assert.deepStrictEqual(budgets, [16384, undefined, 16384, 16384]);
	});

	for (const streaming of [false, true]) {
		test(`discards an output-limited ${streaming ? 'stream' : 'JSON completion'} and resumes an attached edit without repeating tools`, async () => {
			const { agent, requests, setCompletion, changes } = fixture();
			await agent.authenticate(REACH_ACCOUNT_ORIGIN, 'token1');
			const root = URI.file(join(tmpdir(), 'puzzle'));
			await agent.chats.createChat(chat, session, { workingDirectories: [root] });
			agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = ['reach_read_file', 'reach_edit_file', 'reach_create_file'].map(name => ({ name }));
			const calls: string[] = [];
			store.add(agent.onDidChatProgress(e => {
				if (e.kind === 'action' && e.action.type === ActionType.ChatToolCallReady) {
					assert.ok(typeof e.action.toolInput === 'string');
					calls.push(e.action.toolInput);
					agent.onClientToolCallComplete(chat, e.action.toolCallId, { success: true, pastTenseMessage: 'Completed' });
				}
			}));
			const call = (id: string, name: string, args: object) => ({ id, function: { name, arguments: JSON.stringify(args) } });
			setCompletion(async options => {
				const body = JSON.parse(String(options.body));
				assert.strictEqual(body.max_tokens, undefined);
				const message = (tool: ReturnType<typeof call>) => Response.json({ choices: [{ message: { tool_calls: [tool] } }] });
				switch (requests.length) {
					case 1:
						assert.ok(body.messages.some((m: { content: string }) => m.content === 'finish this file'));
						assert.match(body.messages[0].content, /puzzle\.html/);
						assert.deepStrictEqual(body.tools.map((t: { function: { name: string } }) => t.function.name), ['reach_read_file', 'reach_edit_file', findReachTools]);
						return message(call('read', 'reach_read_file', { path: 'puzzle.html' }));
					case 2: {
						const truncated = { content: 'Let me write the full file now.', tool_calls: [call('partial', 'reach_edit_file', { path: 'puzzle.html', oldText: 'end', newText: 'unfinished' })] };
						return streaming
							? new Response('data: ' + JSON.stringify({ choices: [{ delta: truncated, finish_reason: 'length' }] }) + '\n\ndata: [DONE]\n\n', { headers: { 'Content-Type': 'text/event-stream' } })
							: Response.json({ choices: [{ message: truncated, finish_reason: 'length' }] });
					}
					case 3:
						assert.match(body.messages[0].content, /output limit/);
						assert.strictEqual(body.messages.at(-1).tool_call_id, 'read');
						return message(call('valid', 'reach_edit_file', { path: 'puzzle.html', oldText: 'end', newText: 'complete' }));
					case 4:
						return Response.json({ choices: [{ message: { content: 'Saved puzzle.html.' } }] });
					default: throw new Error('Unexpected request');
				}
			});
			await agent.chats.sendMessage(chat, 'ok continue', root, [{ type: MessageAttachmentKind.Simple, label: 'browser', modelRepresentation: 'No shared browser page.' }], 'resume', 'editor', undefined, {
				...resolveAgentChatContext(session, chat),
				previousTurns: [{ id: 'unfinished', state: TurnState.Cancelled, usage: undefined, responseParts: [], message: { text: 'finish this file', origin: { kind: MessageKind.User }, attachments: [{ type: MessageAttachmentKind.Resource, label: 'puzzle.html', uri: URI.joinPath(root, 'puzzle.html').toString() }] } }],
			});
			assert.deepStrictEqual(calls.map(input => JSON.parse(input)), [{ path: 'puzzle.html' }, { path: 'puzzle.html', oldText: 'end', newText: 'complete' }]);
			assert.match(JSON.stringify((await agent.chats.getMessages(chat, session))[0].responseParts), /Saved puzzle\.html/);
			const saved = changes.at(-1)!;
			await agent.chats.releaseChat(chat, session);
			await agent.materializeChat(chat, session, saved);
			agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = [{ name: 'reach_edit_file' }];
			setCompletion(async () => Response.json({ choices: [{ message: { content: 'Ready.' } }] }));
			await agent.chats.sendMessage(chat, 'please complete the file in browser', root, [], 'after-restart', 'editor');
			assert.match(requests.at(-1)!.messages[0].content, /puzzle\.html/);
		});
	}

	test('a second output-limit strike retries with smaller-pieces guidance before recovering', async () => {
		const { agent, requests, setCompletion } = fixture();
		await agent.authenticate(REACH_ACCOUNT_ORIGIN, 'token1');
		const root = URI.file(join(tmpdir(), 'strikes'));
		await agent.chats.createChat(chat, session, { workingDirectories: [root] });
		agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = ['reach_read_file', 'reach_edit_file'].map(name => ({ name }));
		store.add(agent.onDidChatProgress(e => {
			if (e.kind === 'action' && e.action.type === ActionType.ChatToolCallReady) {
				agent.onClientToolCallComplete(chat, e.action.toolCallId, { success: true, pastTenseMessage: 'Saved' });
			}
		}));
		setCompletion(async () => {
			switch (requests.length) {
				case 1:
					return Response.json({ choices: [{ message: { tool_calls: [{ id: 'read', function: { name: 'reach_read_file', arguments: '{"path":"game.js"}' } }] } }] });
				case 2:
				case 3:
					return Response.json({ choices: [{ message: { content: 'truncated' }, finish_reason: 'length' }] });
				case 4:
					return Response.json({ choices: [{ message: { tool_calls: [{ id: 'edit', function: { name: 'reach_edit_file', arguments: '{"path":"game.js","oldText":"a","newText":"b"}' } }] } }] });
				default:
					return Response.json({ choices: [{ message: { content: 'Fixed in pieces.' } }] });
			}
		});
		await agent.chats.sendMessage(chat, 'fix the game', root, [], 'strikes', 'editor');
		const text = (i: number) => requests[i].messages.map(m => typeof m.content === 'string' ? m.content : '').join('\n');
		assert.match(text(2), /output limit/, 'first strike guidance reaches the model');
		assert.match(text(3), /smaller pieces/, 'second strike guidance reaches the model');
	});

	test('a missing content argument and a promise cannot finish a file creation', async () => {
		const { agent, requests, setCompletion, events } = fixture();
		await agent.authenticate(REACH_ACCOUNT_ORIGIN, 'token1');
		await agent.chats.createChat(chat, session);
		agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = [{ name: 'reach_create_file', inputSchema: { type: 'object', required: ['path', 'content'] } }];
		setCompletion(async () => {
			if (requests.length === 1) {
				return Response.json({ choices: [{ message: { tool_calls: [{ id: 'empty', function: { name: 'reach_create_file', arguments: '{"path":"new.html"}' } }] } }] });
			}
			assert.match(requests.at(-1)!.messages.at(-1)!.content, /Missing required string/);
			return Response.json({ choices: [{ message: { content: 'I keep forgetting to include content. Let me write the full file now.' } }] });
		});
		await assert.rejects(agent.chats.sendMessage(chat, 'Create a file', undefined, [], 'create', 'editor'), /did not confirm that write/);
		assert.strictEqual(requests.length, 3);
		assert.ok(!events.some(e => e.kind === 'action' && e.action.type === ActionType.ChatToolCallReady));
		assert.deepStrictEqual(await agent.chats.getMessages(chat, session), []);
	});

	test('quoted examples on a text model never dispatch a workbench action', async () => {
		const { agent, events, setCompletion } = fixture(false);
		await agent.authenticate(REACH_ACCOUNT_ORIGIN, 'token1');
		await agent.chats.createChat(chat, session);
		agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = [{ name: 'sample' }];
		setCompletion(async () => Response.json({ choices: [{ message: { content: 'An example:\n```tool_call\n{"name":"sample","arguments":{}}\n```' } }] }));
		await agent.chats.sendMessage(chat, 'show an example', undefined, [], 'example', 'editor');
		assert.ok(!events.some(e => e.kind === 'action' && e.action.type === ActionType.ChatToolCallStart));
	});

	test('reduces schemas on an input limit while keeping the rest discoverable', async () => {
		const { agent, setCompletion } = fixture();
		await agent.authenticate(REACH_ACCOUNT_ORIGIN, 'token1');
		await agent.chats.createChat(chat, session);
		agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = Array.from({ length: 10 }, (_, i) => ({ name: 'tool_' + i }));
		const requests: string[][] = [];
		setCompletion(async options => {
			const body = JSON.parse(String(options.body));
			const names = body.tools.map((tool: { function: { name: string } }) => tool.function.name);
			requests.push(names);
			return names.length > 6 ? Response.json({ error: { code: 'input_too_large' } }, { status: 400 }) : Response.json({ choices: [{ message: { content: 'Ready' } }] });
		});
		await agent.chats.sendMessage(chat, 'Hello', undefined, [], 'small-model', 'editor');
		assert.deepStrictEqual(requests.map(tools => tools.length), [10, 6]);
		assert.ok(requests[1].includes(findReachTools));
	});

	test('discovers an MCP function registered while the first model request is in flight', async () => {
		const { agent, setCompletion } = fixture();
		await agent.authenticate(REACH_ACCOUNT_ORIGIN, 'token1');
		await agent.chats.createChat(chat, session);
		const client = agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' });
		client.tools = Array.from({ length: 18 }, (_, i) => ({ name: 'initial_' + i }));
		let requestCount = 0;
		setCompletion(async options => {
			requestCount++;
			if (requestCount === 1) {
				client.tools = [...client.tools, { name: 'mcp_receipt', description: 'Look up a receipt' }];
				return Response.json({ choices: [{ message: { tool_calls: [{ id: 'find', function: { name: findReachTools, arguments: '{"query":"receipt"}' } }] } }] });
			}
			const body = JSON.parse(String(options.body));
			assert.ok(body.tools.some((tool: { function: { name: string } }) => tool.function.name === 'mcp_receipt'));
			assert.match(body.messages.at(-1).content, /mcp_receipt/);
			return Response.json({ choices: [{ message: { content: 'Receipt tool is available.' } }] });
		});
		await agent.chats.sendMessage(chat, 'Look up my receipt', undefined, [], 'catalog-change', 'editor');
		assert.equal(requestCount, 2);
	});

	test('a declined action ends the turn before another tool in the same batch can run', async () => {
		const { agent, events, requests, setCompletion } = fixture();
		await agent.authenticate(REACH_ACCOUNT_ORIGIN, 'token1');
		await agent.chats.createChat(chat, session);
		agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = [{ name: 'create' }, { name: 'terminal' }];
		store.add(agent.onDidChatProgress(e => {
			if (e.kind === 'action' && e.action.type === ActionType.ChatToolCallReady) {
				agent.respondToPermissionRequest(e.action.toolCallId, false);
			}
		}));
		setCompletion(async () => Response.json({ choices: [{ message: { tool_calls: ['create', 'terminal'].map(name => ({ id: name, function: { name, arguments: '{}' } })) } }] }));
		await agent.chats.sendMessage(chat, 'Create a file', undefined, [], 'deny', 'editor');
		assert.equal(requests.length, 1);
		assert.equal(events.filter(e => e.kind === 'action' && e.action.type === ActionType.ChatToolCallStart).length, 1);
		assert.ok(events.some(e => e.kind === 'action' && e.action.type === ActionType.ChatDelta && /declined/.test(e.action.content)));
	});

	test('shows the completed action if inference fails after a workspace change', async () => {
		const { agent, events, requests, setCompletion } = fixture();
		await agent.authenticate(REACH_ACCOUNT_ORIGIN, 'token1');
		await agent.chats.createChat(chat, session);
		agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = [{ name: 'create' }];
		store.add(agent.onDidChatProgress(e => {
			if (e.kind === 'action' && e.action.type === ActionType.ChatToolCallReady) {
				agent.onClientToolCallComplete(chat, e.action.toolCallId, { success: true, pastTenseMessage: 'Created note.txt' });
			}
		}));
		setCompletion(async () => requests.length === 1
			? Response.json({ choices: [{ message: { tool_calls: [{ id: 'write1', function: { name: 'create', arguments: '{}' } }] } }] })
			: new Response(null, { status: 503 }));
		await assert.rejects(agent.chats.sendMessage(chat, 'Create a note', undefined, [], 'outage', 'editor'), /HTTP 503[\s\S]*create completed/);
		// Round two 503s and is retried once by the transport (2 requests); the loop-level
		// after-tools retry regenerates the round once more (2 more) after the first success.
		assert.strictEqual(requests.length, 5);
		assert.equal(events.filter(e => e.kind === 'action' && e.action.type === ActionType.ChatToolCallStart).length, 1);
		assert.ok(!events.some(e => e.kind === 'action' && e.action.type === ActionType.ChatTurnComplete));
	});

	test('an oversized streamed tool payload is discarded and retried as JSON, and a second size rejection fails the turn', async () => {
		const { agent, events, setCompletion } = fixture();
		await agent.authenticate(REACH_ACCOUNT_ORIGIN, 'token1');
		await agent.chats.createChat(chat, session);
		agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = [{ name: 'write_note' }];
		const streams: boolean[] = [];
		setCompletion(async init => {
			const body = JSON.parse(String(init.body));
			streams.push(body.stream);
			const error = { error: { code: 'response_too_large', message: 'The model response exceeded the service size limit.' } };
			if (!body.stream) { return Response.json(error, { status: 502 }); }
			const frame = { choices: [{ delta: { tool_calls: [{ index: 0, id: 'partial', function: { name: 'write_note', arguments: '{}' } }] } }] };
			return new Response([frame, error].map(data => 'data: ' + JSON.stringify(data) + '\n\n').join('') + 'data: [DONE]\n\n', { headers: { 'Content-Type': 'text/event-stream' } });
		});
		await assert.rejects(agent.chats.sendMessage(chat, 'Write a note', undefined, [], 'oversize', 'editor'), /response size limit/);
		assert.deepStrictEqual(streams, [true, false]);
		assert.ok(!events.some(e => e.kind === 'action' && (e.action.type === ActionType.ChatToolCallStart || e.action.type === ActionType.ChatTurnComplete)));
	});

	test('resumes after a cut-off stream without running its partial tool or replaying a finished call', async () => {
		const { agent, events, requests, setCompletion } = fixture();
		await agent.authenticate(REACH_ACCOUNT_ORIGIN, 'token1');
		await agent.chats.createChat(chat, session);
		agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = [{ name: 'write_note' }];
		store.add(agent.onDidChatProgress(e => {
			if (e.kind === 'action' && e.action.type === ActionType.ChatToolCallReady) {
				agent.onClientToolCallComplete(chat, e.action.toolCallId, { success: true, pastTenseMessage: 'Saved' });
			}
		}));
		const tool = (id: string) => ({ tool_calls: [{ id, function: { name: 'write_note', arguments: '{}' } }] });
		let cancelled = false;
		setCompletion(async () => {
			if (requests.length === 1) { return Response.json({ choices: [{ message: tool('saved') }] }); }
			const interrupted = requests.length === 2;
			const choice = interrupted ? { delta: tool('partial') } : { delta: { content: 'Saved once.' }, finish_reason: 'stop' };
			return new Response(new ReadableStream<Uint8Array>({
				start(controller) {
					controller.enqueue(new TextEncoder().encode('data: ' + JSON.stringify({ choices: [choice] }) + '\n\n'));
					if (interrupted) { controller.close(); }
				},
				cancel() { cancelled = true; },
			}), { headers: { 'Content-Type': 'text/event-stream' } });
		});
		await agent.chats.sendMessage(chat, 'Write a note', undefined, [], 'recover', 'editor');
		assert.deepStrictEqual(requests[2].messages, requests[1].messages);
		assert.strictEqual(events.filter(e => e.kind === 'action' && e.action.type === ActionType.ChatToolCallStart).length, 1);
		assert.ok(cancelled);
		assert.match(JSON.stringify(await agent.chats.getMessages(chat, session)), /Saved once/);
	});

	for (const native of [true, false]) {
		test(`recovers a prose-prefixed Edge0 read through the ${native ? 'native' : 'text'} tool protocol`, async () => {
			const { agent, events, requests, setCompletion } = fixture(native);
			await agent.authenticate(REACH_ACCOUNT_ORIGIN, 'token1');
			await agent.chats.createChat(chat, session);
			agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = [{ name: 'vscode_readFile' }];
			const xml = '<tool_call>\n<function=vscode_readFile><parameter=startLine>1</parameter></function>\n</tool_call>';
			store.add(agent.onDidChatProgress(event => {
				if (event.kind === 'action' && event.action.type === ActionType.ChatToolCallReady) {
					assert.equal(requests.length, 1);
					agent.onClientToolCallComplete(chat, event.action.toolCallId, { success: true, pastTenseMessage: 'Read diagram', structuredContent: { rows: 68 } });
				}
			}));
			setCompletion(async () => {
				if (requests.length === 1) {
					const chunks = ['Let me read the diagram.\n\n<to', 'ol_', xml.slice('<tool_'.length)];
					return new Response(chunks.map(content => 'data: ' + JSON.stringify({ choices: [{ delta: { content } }] }) + '\n\n').join('') + 'data: [DONE]\n\n',
						{ headers: { 'Content-Type': 'text/event-stream' } });
				}
				return Response.json({ choices: [{ message: { content: 'Read all 68 lines.' } }] });
			});
			await agent.chats.sendMessage(chat, 'Read the diagram', undefined, [], 'edge0-read', 'editor');
			assert.equal(requests.length, 2);
			assert.ok(!JSON.stringify(requests[1].messages).includes('<function=vscode_readFile>'));
			const receipt = requests[1].messages.at(-1)!.content;
			assert.match(native ? receipt : JSON.parse(receipt.slice(receipt.indexOf('\n') + 1)).result, /"rows":68/);
			assert.equal(events.filter(e => e.kind === 'action' && e.action.type === ActionType.ChatToolCallReady).length, 1);
			const output = events.flatMap(e => e.kind === 'action' && e.action.type === ActionType.ChatDelta ? [e.action.content] : []).join('');
			assert.doesNotMatch(output, /<to|vscode_readFile/);
			assert.match(output, /Read all 68 lines\./);
		});

		test(`executes a ${native ? 'native' : 'text-protocol'} function and feeds the result into the next model request`, async () => {
			const { agent, events, setCompletion } = fixture(native);
			await agent.authenticate(REACH_ACCOUNT_ORIGIN, 'token1');
			await agent.chats.createChat(chat, session);
			const client = agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' });
			client.tools = [{ name: 'workspace.read', description: 'Read the requested file' }];
			const ids = new Set<string>();
			store.add(agent.onDidChatProgress(e => {
				if (e.kind === 'action' && e.action.type === ActionType.ChatToolCallReady) {
					assert.strictEqual(typeof e.action.toolInput, 'string');
					assert.deepStrictEqual(JSON.parse(String(e.action.toolInput)), { path: 'test.json' });
					agent.onClientToolCallComplete(chat, e.action.toolCallId, { success: true, pastTenseMessage: 'Read file', structuredContent: { receipt: 'actual-output-791' } });
				}
			}));
			setCompletion(async options => {
				const id = new Headers(options.headers).get('Idempotency-Key')!;
				assert.ok(!ids.has(id)); ids.add(id);
				const body = JSON.parse(String(options.body));
				if (ids.size === 1) {
					if (native) {
						assert.equal(body.tools[0].function.name, 'workspace_read');
						return Response.json({ choices: [{ message: { content: 'tool_calls', tool_calls: [{ id: 'read1', function: { name: 'workspace_read', arguments: '{"path":"test.json"}' } }] } }] });
					}
					assert.equal(body.tools, undefined);
					assert.match(body.messages[0].content, /workspace_read/);
					return Response.json({ choices: [{ message: { content: '```tool_call\n{"name":"workspace_read","arguments":{"path":"test.json"}}\n```' } }] });
				}
				assert.equal(body.messages.at(-1).role, native ? 'tool' : 'user');
				assert.match(body.messages.at(-1).content, /actual-output-791/);
				return Response.json({ choices: [{ message: { content: 'The receipt was actual-output-791.' } }] });
			});
			await agent.chats.sendMessage(chat, 'Read the test file', undefined, [], 'turn1', 'editor');
			assert.equal(ids.size, 2);
			assert.equal(events.filter(e => e.kind === 'action' && e.action.type === ActionType.ChatToolCallReady).length, 1);
			assert.ok(events.some(e => e.kind === 'action' && e.action.type === ActionType.ChatDelta && e.action.content.includes('actual-output-791')));
		});
	}

	for (const retrySucceeds of [true, false]) {
		test(`handles an undeclared Edge0 tool call after a workspace tool when the retry ${retrySucceeds ? 'answers' : 'also fails'}`, async () => {
			const dir = await mkdtemp(join(tmpdir(), 'reach-edge0-overview-'));
			try {
				await writeFile(join(dir, 'README.md'), '# Kite Lab\n\nNotes for copper kite makers.');
				const { agent, events, requests, setCompletion } = fixture();
				await agent.authenticate(REACH_ACCOUNT_ORIGIN, 'token1');
				const root = URI.file(dir);
				await agent.chats.createChat(chat, session, { workingDirectories: [root] });
				agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = [{ name: 'copilot_findFiles' }];
				store.add(agent.onDidChatProgress(e => {
					if (e.kind === 'action' && e.action.type === ActionType.ChatToolCallReady) {
						agent.onClientToolCallComplete(chat, e.action.toolCallId, { success: true, pastTenseMessage: 'Found README.md' });
					}
				}));
				const parts = [retrySucceeds ? '\n<tool_' : 'I will look for the files.\n\n<tool_', 'call> <function=reach_findFiles> <parameter=maxResults> 200 </parameter>', ' <parameter=query> **/*.md </parameter> </function> </tool_call>'];
				const ids = new Set<string | null>();
				setCompletion(async options => {
					ids.add(new Headers(options.headers).get('Idempotency-Key'));
					if (requests.length === 1) {
						return Response.json({ choices: [{ message: { tool_calls: [{ id: 'find', function: { name: 'copilot_findFiles', arguments: '{"query":"**/*.md"}' } }] } }] });
					}
					if (requests.length === 3 && retrySucceeds) {
						return Response.json({ choices: [{ message: { content: 'Kite Lab documents copper kite making.' } }] });
					}
					const wire = parts.map(content => 'data: ' + JSON.stringify({ choices: [{ delta: { content } }] }) + '\n\n').join('');
					return new Response(wire + 'data: [DONE]\n\n', { headers: { 'Content-Type': 'text/event-stream' } });
				});
				if (retrySucceeds) {
					await agent.chats.sendMessage(chat, 'summarize the workspace', root, [], 'edge0-xml', 'editor');
					assert.equal(requests.length, 3);
					assert.equal(ids.size, 3);
					// The recovered undeclared call is answered in-band, naming the offender, and the turn continues.
					assert.match(requests[2].messages.at(-1)!.content, /does not exist/);
					assert.match(requests[2].messages.at(-1)!.content, /reach_findFiles/);
					assert.ok(JSON.stringify(requests[2].messages).includes('Found README.md'));
				} else {
					// A model that keeps inventing the same undeclared function stops the turn,
					// preserving the completed work instead of dying silently.
					await assert.rejects(
						agent.chats.sendMessage(chat, 'summarize the workspace', root, [], 'edge0-xml', 'editor'),
						/kept calling functions that do not exist[\s\S]*reach_findFiles[\s\S]*copilot_findFiles completed/
					);
				}
				assert.equal(events.filter(e => e.kind === 'action' && e.action.type === ActionType.ChatToolCallStart).length, 1);
				const reply = events.flatMap(e => e.kind === 'action' && e.action.type === ActionType.ChatDelta ? [e.action.content] : []).join('');
				assert.doesNotMatch(reply, /<tool_call>|reach_findFiles/);
				if (retrySucceeds) { assert.match(reply, /Kite Lab documents copper kite making/); }
			} finally { await rm(dir, { recursive: true, force: true, maxRetries: 3 }); }
		});
	}

	test('streams a fenced XML example verbatim without retrying or starting a tool', async () => {
		const { agent, events, requests, setCompletion } = fixture();
		await agent.authenticate(REACH_ACCOUNT_ORIGIN, 'token1');
		await agent.chats.createChat(chat, session);
		agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = [{ name: 'vscode_readFile' }];
		const chunks = ['Example:\n```xml\n<tool_', 'call><function=vscode_readFile>{}</function></tool_call>\n```'];
		setCompletion(async () => new Response(chunks.map(content =>
			'data: ' + JSON.stringify({ choices: [{ delta: { content } }] }) + '\n\n').join('') + 'data: [DONE]\n\n', { headers: { 'Content-Type': 'text/event-stream' } }));
		await agent.chats.sendMessage(chat, 'Explain the XML syntax', undefined, [], 'xml-example', 'editor');
		assert.equal(requests.length, 1);
		assert.ok(!events.some(e => e.kind === 'action' && e.action.type === ActionType.ChatToolCallStart));
		assert.equal(events.flatMap(e => e.kind === 'action' && e.action.type === ActionType.ChatDelta ? [e.action.content] : []).join(''), chunks.join(''));
	});

	test('restores an existing chat across a host restart while its saved credential is being verified', async () => {
		const first = fixture();
		await first.agent.authenticate(REACH_ACCOUNT_ORIGIN, 'token1');
		await first.agent.chats.createChat(chat, session);
		await first.agent.chats.sendMessage(chat, 'saved question', undefined, [], 'a');
		const data = first.changes.at(-1)!;
		await first.agent.shutdown();

		const restored = fixture();
		await restored.agent.materializeChat(chat, session, data);
		await assert.rejects(restored.agent.chats.getMessages(chat, session), /Sign in to REACH to use this chat/);
		const account = new DeferredPromise<void>(), catalog = new DeferredPromise<void>();
		restored.holdAccount(account.p);
		restored.holdCatalog(catalog.p);
		const authenticating = restored.agent.authenticate(REACH_ACCOUNT_ORIGIN, 'token1');
		const history = restored.agent.chats.getMessages(chat, session);
		const selecting = restored.agent.chats.changeModel(chat, { id: 'chat2' }, session);
		assert.equal(restored.accountReads(), 1);
		await account.complete();
		// Loading a saved transcript and selecting a model depend on the owner identity, not model discovery.
		assert.equal((await history)[0].message.text, 'saved question');
		await selecting;
		await catalog.error(new Error('catalog unavailable'));
		assert.equal(await authenticating, true);
		assert.equal(restored.requests.length, 0);
		assert.equal((await restored.agent.chats.getMessages(chat, session))[0].message.text, 'saved question');

		await restored.agent.authenticate(REACH_ACCOUNT_ORIGIN, '');
		await assert.rejects(restored.agent.chats.getMessages(chat, session), /Sign in to REACH to use this chat/);
		await restored.agent.authenticate(REACH_ACCOUNT_ORIGIN, 'token2');
		await assert.rejects(restored.agent.chats.getMessages(chat, session), /account that started/);
		await assert.rejects(restored.agent.chats.sendMessage(chat, 'follow up', undefined, [], 'b'), /account that started/);
		assert.equal(restored.requests.length, 0);
	});

	test('revocation during an initial history read never releases a saved turn to the revoked credential', async () => {
		const restored = fixture();
		const data = JSON.stringify({ version: 1, accountId: 'first-account', turns: [{ id: 'turn', prompt: 'private question', reply: 'private reply' }] });
		await restored.agent.materializeChat(chat, session, data);
		const gate = new DeferredPromise<void>();
		restored.holdAccount(gate.p);
		const signingIn = restored.agent.authenticate(REACH_ACCOUNT_ORIGIN, 'token1');
		const reading = restored.agent.chats.getMessages(chat, session);
		const sending = restored.agent.chats.sendMessage(chat, 'new question', undefined, [], 'new-turn');
		await restored.agent.authenticate(REACH_ACCOUNT_ORIGIN, '');
		await gate.complete();
		assert.equal(await signingIn, false);
		await assert.rejects(reading, /Sign in to REACH to use this chat/);
		await assert.rejects(sending, isCancellationError);
		assert.equal(await restored.agent.authenticate(REACH_ACCOUNT_ORIGIN, 'token1', 0), false);
		await assert.rejects(restored.agent.chats.getMessages(chat, session), /Sign in to REACH to use this chat/);
		assert.equal(restored.requests.length, 0);
	});

	test('a replacement credential for the same owner wins a racing restore', async () => {
		const restored = fixture();
		await restored.agent.materializeChat(chat, session, JSON.stringify({ version: 1, accountId: 'first-account', turns: [{ id: 'turn', prompt: 'saved', reply: 'answer' }] }));
		const old = new DeferredPromise<void>(), current = new DeferredPromise<void>();
		restored.holdAccount(old.p);
		const oldSignIn = restored.agent.authenticate(REACH_ACCOUNT_ORIGIN, 'token1');
		const reading = restored.agent.chats.getMessages(chat, session);
		restored.holdAccount(current.p);
		const newSignIn = restored.agent.authenticate(REACH_ACCOUNT_ORIGIN, 'token1-renewed');
		await old.complete();
		assert.equal(await oldSignIn, false);
		await current.complete();
		assert.equal(await newSignIn, true);
		assert.equal((await reading)[0].message.text, 'saved');
		assert.equal(restored.requests.length, 0);
	});

	test('account catalog, explicit model, default route, persisted follow-up, and regular chat persona', async () => {
		const { agent, requests, changes, events } = fixture();
		await agent.chats.createChat(chat, session);
		assert.equal(agent.getDescriptor().displayName, 'REACH');
		await assert.rejects(agent.chats.sendMessage(chat, 'first', undefined, [], 'a'), /Sign in/);
		assert.equal(await agent.authenticate(REACH_ACCOUNT_ORIGIN, 'token1'), true);
		assert.deepStrictEqual(agent.models.get().map(model => [model.id, model.maxContextWindow]), [['chat1', 65536], ['chat2', undefined]]);
		await agent.chats.changeModel(chat, { id: 'chat2' }, session);
		await agent.chats.sendMessage(chat, 'first', undefined, [{ type: MessageAttachmentKind.Simple, label: 'context', modelRepresentation: 'selected passage' }], 'a');
		assert.equal(requests[0].model, 'chat2');
		assert.match(requests[0].messages[0].content, /REACH, a chat provider/);
		assert.doesNotMatch(requests[0].messages[0].content, /Signal-C Assistant/);
		assert.match(requests[0].messages[1].content, /selected passage/);
		assert.ok(events.some(e => e.kind === 'action' && e.action.type === ActionType.ChatDelta));
		assert.ok(events.some(e => e.kind === 'action' && e.action.type === ActionType.ChatTurnComplete));
		const data = changes.at(-1)!;
		assert.ok(!data.includes('token1'));
		await agent.chats.releaseChat(chat, session);
		await agent.materializeChat(chat, session, data);
		const history = await agent.chats.getMessages(chat, session);
		assert.equal(history[0].message.text, 'first');
		assert.equal(history[0].responseParts[0].kind, ResponsePartKind.Markdown);
		await agent.chats.changeModel(chat, { id: 'auto' }, session);
		await agent.chats.sendMessage(chat, 'follow-up', undefined, [], 'b');
		assert.equal(requests[1].model, 'chat1');
		assert.deepStrictEqual(requests[1].messages.slice(1).map(message => message.content), ['first\n\nselected passage', 'answer', 'follow-up']);
		await agent.authenticate(REACH_ACCOUNT_ORIGIN, 'token2');
		await assert.rejects(agent.chats.sendMessage(chat, 'other account', undefined, [], 'c'), /account that started/);
		await assert.rejects(agent.chats.getMessages(chat, session), /account that started/);
		assert.equal(requests.length, 2);
	});

	test('follow-up history replays completed tool actions with the assistant reply', async () => {
		const { agent, requests, changes, setCompletion } = fixture();
		await agent.authenticate(REACH_ACCOUNT_ORIGIN, 'token1');
		await agent.chats.createChat(chat, session);
		agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = Array.from({ length: 6 }, (_, i) => ({ name: 'extra_' + i }));
		store.add(agent.onDidChatProgress(e => {
			if (e.kind === 'action' && e.action.type === ActionType.ChatToolCallReady) {
				agent.onClientToolCallComplete(chat, e.action.toolCallId, { success: true, pastTenseMessage: 'Done' });
			}
		}));
		setCompletion(async () => requests.length === 1
			? Response.json({ choices: [{ message: { tool_calls: [{ id: 'write', function: { name: 'extra_3', arguments: '{}' } }] } }] })
			: Response.json({ choices: [{ message: { content: 'Created the thing.' } }] }));
		await agent.chats.sendMessage(chat, 'build it', undefined, [], 'turn-one', 'editor');
		const data = changes.at(-1)!;
		assert.ok(data.includes('extra_3 completed'), 'persisted record carries the completed action');
		await agent.chats.releaseChat(chat, session);
		await agent.materializeChat(chat, session, data);
		await agent.chats.sendMessage(chat, 'now update it', undefined, [], 'turn-two', 'editor');
		const followUp = requests.at(-1)!;
		const assistantReplay = followUp.messages.map(message => message.content).find(content => content.includes('Created the thing.'));
		assert.ok(assistantReplay, 'follow-up replays the earlier reply');
		assert.match(assistantReplay!, /Actions completed in that turn:\n- extra_3 completed\./);
	});

	test('a turn that never finishes ends gracefully at the round ceiling with completed work preserved', async () => {
		const { agent, requests, setCompletion } = fixture();
		await agent.authenticate(REACH_ACCOUNT_ORIGIN, 'token1');
		await agent.chats.createChat(chat, session);
		agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = [{ name: 'extra_probe' }];
		store.add(agent.onDidChatProgress(e => {
			if (e.kind === 'action' && e.action.type === ActionType.ChatToolCallReady) {
				agent.onClientToolCallComplete(chat, e.action.toolCallId, { success: true, pastTenseMessage: 'Probed' });
			}
		}));
		setCompletion(async () => Response.json({ choices: [{ message: { tool_calls: [{ id: 'probe' + requests.length, function: { name: 'extra_probe', arguments: `{"step":${requests.length}}` } }] } }] }));
		// The ceiling never throws: the turn completes with a preserved-work summary.
		await agent.chats.sendMessage(chat, 'probe forever', undefined, [], 'endless', 'editor');
		assert.equal(requests.length, 81, 'eighty tool rounds plus one final tool-free inference');
		assert.ok(requests[21].messages.some(message => message.content.includes('many tool rounds')), 'round 20 recovery pressure reaches the model');
		assert.ok(requests[36].messages.some(message => message.content.includes('spending its tool rounds')), 'round 35 recovery pressure reaches the model');
		assert.ok(!requests[80].tools?.length, 'the final inference is tool-free');
	});

	test('the round ceiling ends with the model summary when the final inference answers', async () => {
		const { agent, requests, changes, setCompletion } = fixture();
		await agent.authenticate(REACH_ACCOUNT_ORIGIN, 'token1');
		await agent.chats.createChat(chat, session);
		agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = [{ name: 'extra_probe' }];
		store.add(agent.onDidChatProgress(e => {
			if (e.kind === 'action' && e.action.type === ActionType.ChatToolCallReady) {
				agent.onClientToolCallComplete(chat, e.action.toolCallId, { success: true, pastTenseMessage: 'Probed' });
			}
		}));
		setCompletion(async init => {
			const body = JSON.parse(String(init?.body ?? '{}'));
			if (!body.tools?.length) {
				return Response.json({ choices: [{ message: { content: 'I probed every corner but could not finish the task.' } }] });
			}
			return Response.json({ choices: [{ message: { tool_calls: [{ id: 'probe' + requests.length, function: { name: 'extra_probe', arguments: `{"step":${requests.length}}` } }] } }] });
		});
		await agent.chats.sendMessage(chat, 'probe forever', undefined, [], 'endless', 'editor');
		assert.equal(requests.length, 81, 'eighty tool rounds plus the final summary inference');
		assert.ok(changes.at(-1)!.includes('I probed every corner but could not finish the task.'), 'the model-written summary becomes the turn reply');
	});

	test('exploration without a write attempt steers a fix task toward editing', async () => {
		const { agent, requests, setCompletion } = fixture();
		await agent.authenticate(REACH_ACCOUNT_ORIGIN, 'token1');
		const root = URI.file(join(tmpdir(), 'paralysis'));
		await agent.chats.createChat(chat, session, { workingDirectories: [root] });
		agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = ['reach_read_file', 'reach_edit_file'].map(name => ({ name }));
		store.add(agent.onDidChatProgress(e => {
			if (e.kind === 'action' && e.action.type === ActionType.ChatToolCallReady) {
				agent.onClientToolCallComplete(chat, e.action.toolCallId, { success: true, pastTenseMessage: 'Read', content: [{ type: ToolResultContentType.Text, text: 'data' + requests.length }] });
			}
		}));
		setCompletion(async () => requests.length >= 10
			? Response.json({ choices: [{ message: { content: 'I could not locate the problem.' } }] })
			: Response.json({ choices: [{ message: { tool_calls: [{ id: 'read' + requests.length, function: { name: 'reach_read_file', arguments: `{"path":"file${requests.length}.js"}` } }] } }] }));
		await agent.chats.sendMessage(chat, 'fix the scoring bug in the game', root, [], 'paralysis', 'editor');
		assert.ok(requests.slice(8).some(request => request.messages.some(message => typeof message.content === 'string' && message.content.includes('without changing anything'))), 'anti-paralysis steering reaches the model after eight write-free rounds');
	});

	test('repeated edit failures steer the model toward revision edits', async () => {
		const { agent, requests, setCompletion } = fixture();
		await agent.authenticate(REACH_ACCOUNT_ORIGIN, 'token1');
		const root = URI.file(join(tmpdir(), 'editloop'));
		await agent.chats.createChat(chat, session, { workingDirectories: [root] });
		agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = ['reach_create_file', 'reach_read_file', 'reach_edit_file'].map(name => ({ name }));
		store.add(agent.onDidChatProgress(e => {
			if (e.kind === 'action' && e.action.type === ActionType.ChatToolCallReady) {
				agent.onClientToolCallComplete(chat, e.action.toolCallId, { success: false, pastTenseMessage: 'Edit failed' });
			}
		}));
		setCompletion(async () => Response.json({ choices: [{ message: { tool_calls: [{ id: 'edit' + requests.length, function: { name: 'reach_edit_file', arguments: `{"path":"game.js","oldText":"x${requests.length}","newText":"y"}` } }] } }] }));
		await agent.chats.sendMessage(chat, 'fix the game', root, [], 'editloop', 'editor');
		assert.ok(requests.slice(2).some(request => request.messages.some(message => typeof message.content === 'string' && message.content.includes('Edits keep failing'))), 'circuit-breaker steering reaches the model after two edit failures');
	});

	test('summarizes the selected workspace on a tool-only completion using source-labelled project facts', async () => {
		const dir = await mkdtemp(join(tmpdir(), 'reach-overview-'));
		try {
			await mkdir(join(dir, '.vscode'));
			await mkdir(join(dir, 'src'));
			await Promise.all([
				writeFile(join(dir, 'package.json'), JSON.stringify({ name: 'kite-lab', description: 'Tracks dye experiments for copper kites.', privateToken: 'do-not-send-token' })),
				writeFile(join(dir, 'README.md'), '# Kite Lab\n\nNotes for copper kite makers.'),
				writeFile(join(dir, 'Cargo.toml'), 'sensitive-manifest'),
				writeFile(join(dir, 'go.mod'), 'private-module'),
				writeFile(join(dir, '.env'), 'credential=value'),
				writeFile(join(dir, '.gitignore'), 'Cargo.toml\n'),
				writeFile(join(dir, '.vscode', 'settings.json'), '{"files.exclude":{"go.mod":true}}'),
				writeFile(join(dir, 'src', 'index.ts'), 'unrequested source code'),
			]);
			const { agent, requests, changes, setCompletion } = fixture();
			await agent.authenticate(REACH_ACCOUNT_ORIGIN, 'token1');
			const root = URI.file(dir);
			await agent.chats.createChat(chat, session, { workingDirectories: [root] });
			setCompletion(async () => Response.json({ choices: [{ message: { role: 'assistant', content: 'tool_calls' } }] }));
			await agent.chats.sendMessage(chat, 'can you summarize the workspace', root, [], 'a');
			assert.equal(requests.length, 1);
			const facts = requests[0].messages.at(-1)!.content;
			for (const text of ['package.json', 'Tracks dye experiments for copper kites.', 'README.md', 'Notes for copper kite makers.', 'src/']) {
				assert.ok(facts.includes(text), text);
			}
			assert.doesNotMatch(facts, /do-not-send-token|Cargo\.toml|go\.mod|private-module|credential=value|unrequested source code/);
			const saved = JSON.parse(changes.at(-1)!);
			assert.match(saved.turns[0].reply, /Workspace overview[\s\S]*Tracks dye experiments/);
			assert.equal(saved.turns[0].input, undefined);
			setCompletion(undefined);
			await agent.chats.sendMessage(chat, 'Thank you', root, [], 'b');
			assert.equal(requests[1].messages.at(-1)?.content, 'Thank you');
		} finally { await rm(dir, { recursive: true, force: true, maxRetries: 3 }); }
	});

	test('a folderless workspace question asks for a folder without sending an inference request', async () => {
		const { agent, requests, changes } = fixture();
		await agent.authenticate(REACH_ACCOUNT_ORIGIN, 'token1');
		await agent.chats.createChat(chat, session);
		await agent.chats.sendMessage(chat, 'what does this repo do?', undefined, [], 'a');
		assert.equal(requests.length, 0);
		assert.match(JSON.parse(changes.at(-1)!).turns[0].reply, /Choose a workspace/);
	});

	test('legacy tool-only turns display recovery guidance and are excluded from model history', async () => {
		const { agent, requests, events, changes } = fixture();
		await agent.authenticate(REACH_ACCOUNT_ORIGIN, 'token1');
		await agent.materializeChat(chat, session, JSON.stringify({ version: 1, accountId: 'first-account', createdAt: 1, modifiedAt: 2, turns: [
			{ id: 'legacy', startedAt: new Date(1).toISOString(), prompt: 'old question', reply: 'tool_calls' },
		] }));
		const [turn] = await agent.chats.getMessages(chat, session);
		assert.ok(turn.responseParts.some(part => part.kind === ResponsePartKind.Markdown && part.content.includes('without a text answer')));
		assert.deepStrictEqual(changes, []);
		await agent.chats.sendMessage(chat, 'new question', undefined, [], 'new');
		assert.deepStrictEqual(requests[0].messages.slice(1), [{ role: 'user', content: 'new question' }]);
		assert.ok(events.length > 0);
	});

	test('denied, exhausted, removed and empty model catalog never dispatch inference', async () => {
		const { agent, requests, setActive, setEligible, setRemaining } = fixture();
		await agent.chats.createChat(chat, session);
		assert.equal(await agent.authenticate('https://other.example', 'token1'), false);
		setActive(false);
		assert.equal(await agent.authenticate(REACH_ACCOUNT_ORIGIN, 'token1'), false);
		await assert.rejects(agent.chats.sendMessage(chat, 'test', undefined, [], 'a'), /active REACH subscription/);
		setActive(true);
		assert.equal(await agent.authenticate(REACH_ACCOUNT_ORIGIN, 'token1'), true);
		setRemaining(0);
		await assert.rejects(agent.chats.sendMessage(chat, 'test', undefined, [], 'b'), /allowance/);
		setRemaining(1);
		await agent.chats.changeModel(chat, { id: 'removed' }, session);
		await assert.rejects(agent.chats.sendMessage(chat, 'test', undefined, [], 'c'), /no longer available/);
		await agent.chats.changeModel(chat, { id: 'auto' }, session);
		setEligible(false);
		assert.equal(await agent.authenticate(REACH_ACCOUNT_ORIGIN, 'token1'), true);
		await assert.rejects(agent.chats.sendMessage(chat, 'test', undefined, [], 'd'), /no subscription chat models/);
		assert.equal(requests.length, 0);
	});

	test('sign-out cancels a streaming turn and clears the account model catalog', async () => {
		const { agent, setCompletion, events, changes } = fixture();
		await agent.authenticate(REACH_ACCOUNT_ORIGIN, 'token1');
		await agent.chats.createChat(chat, session);
		const started = new DeferredPromise<void>();
		setCompletion(options => new Promise((_resolve, reject) => {
			options.signal!.addEventListener('abort', () => reject(options.signal!.reason), { once: true });
			void started.complete();
		}));
		const pending = agent.chats.sendMessage(chat, 'test', undefined, [], 'a');
		await started.p;
		await agent.authenticate(REACH_ACCOUNT_ORIGIN, '');
		await assert.rejects(pending, isCancellationError);
		assert.deepStrictEqual(agent.models.get(), []);
		assert.equal(changes.length, 1);
		assert.equal(JSON.parse(changes[0]).accountId, 'first-account');
		assert.deepStrictEqual(JSON.parse(changes[0]).turns, []);
		assert.equal(events.length, 0);
	});

	for (const [scenario, token] of [['same account', 'token1-renewed'], ['different account', 'token2'], ['sign-out', '']]) {
		test('a tool turn at credential expiry resumes only for its owner: ' + scenario, async () => {
			const { agent, setCompletion, requests } = fixture();
			const clock = sinon.useFakeTimers({ toFake: ['Date'] });
			try {
				await agent.authenticate(REACH_ACCOUNT_ORIGIN, 'token1', 60);
				await agent.chats.createChat(chat, session);
				agent.getOrCreateActiveClient(chat, session, { clientId: 'editor' }).tools = [{ name: 'list_files' }];
				setCompletion(async () => Response.json({ choices: [{ message: requests.length === 1
					? { tool_calls: [{ id: 'list', function: { name: 'list_files', arguments: '{}' } }] }
					: { content: 'Finished' } }] }));
				const paused = new DeferredPromise<void>();
				let executions = 0;
				store.add(agent.onDidChatProgress(e => {
					if (e.kind !== 'action') { return; }
					if (e.action.type === ActionType.ChatToolCallReady) {
						executions++;
						clock.setSystemTime(Date.now() + 61_000);
						agent.onClientToolCallComplete(chat, e.action.toolCallId, { success: true, pastTenseMessage: 'Listed workspace files' });
					}
					if (e.action.type === ActionType.ChatResponsePart && e.action.part.kind === ResponsePartKind.SystemNotification && String(e.action.part.content).includes('sign-in expired')) { void paused.complete(); }
				}));
				const sending = agent.chats.sendMessage(chat, 'List files', undefined, [], 'renew', 'editor');
				const finished = token === 'token1-renewed' ? sending : assert.rejects(sending, isCancellationError);
				await paused.p;
				assert.strictEqual(requests.length, 1);
				await agent.authenticate(REACH_ACCOUNT_ORIGIN, token, 60);
				await finished;
				assert.strictEqual(executions, 1);
				assert.strictEqual(requests.length, token === 'token1-renewed' ? 2 : 1);
				if (token === 'token1-renewed') { assert.strictEqual(requests[1].messages.at(-1)?.tool_call_id, 'list'); }
			} finally { clock.restore(); }
		});
	}
});
