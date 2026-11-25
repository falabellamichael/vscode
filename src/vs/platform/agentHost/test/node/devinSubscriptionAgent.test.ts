/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import assert from 'assert';
import { chmod, mkdtemp, rm, writeFile } from 'fs/promises';
import { tmpdir } from 'os';
import { join } from '../../../../base/common/path.js';
import { isWindows } from '../../../../base/common/platform.js';
import sinon from 'sinon';
import { DeferredPromise } from '../../../../base/common/async.js';
import { isCancellationError } from '../../../../base/common/errors.js';
import { URI } from '../../../../base/common/uri.js';
import { ensureNoDisposablesAreLeakedInTestSuite } from '../../../../base/test/common/utils.js';
import { type ILogService, NullLogService } from '../../../log/common/log.js';
import { AcpClient, isAcpRecord } from '../../node/acp/acpClient.js';
import { DevinSubscriptionAgent, DEVIN_SUBSCRIPTION_RESOURCE, devinSubscriptionEnvironment, parseDevinModels } from '../../node/acp/devinSubscriptionAgent.js';
import { ActionType, isChatAction } from '../../common/state/sessionActions.js';
import { type AgentSignal, type IAgentChatOwnerChange, fixedChatOwnerRef, resolveAgentChatContext } from '../../common/agent.js';
import { Emitter, Event } from '../../../../base/common/event.js';
import { TestAcpTransport } from './acpTestUtils.js';
import { CustomizationType, type ClientPluginCustomization, MessageAttachmentKind, type MessageAttachment, type ChatState, MessageKind, ResponsePartKind, SessionStatus, ToolCallStatus, ToolCallConfirmationReason, ToolCallCancellationReason, ToolResultContentType, TurnState } from '../../common/state/sessionState.js';
import { CustomizationEnablementKind } from '../../common/state/protocol/state.js';
import { chatReducer } from '../../common/state/protocol/reducers.js';
import { resolveAcpPrompt } from '../../node/acp/acpPromptResolver.js';
import { type AcpMcpServer, type IDevinMcpConfiguration } from '../../node/acp/devinMcpConfiguration.js';
import { SessionConfigKey } from '../../common/sessionConfigKeys.js';
import type { IAgentChatConfigurationChangeEvent } from '../../node/agentConfigurationService.js';

const catalog = { families: [{ variants: [{ model_uid: 'swe-2-high', label: 'SWE-2 High', max_context_tokens: 262000 }] }] };
const chat = URI.parse('agent-host-chat://devin-subscription/chat');
const owner = URI.parse('agent-host-devin-subscription:/session');
const cwd = URI.file('/project with spaces');

function toolUpdate(transport: TestAcpTransport, update: Record<string, unknown>): void {
	transport.receive({ jsonrpc: '2.0', method: 'session/update', params: { sessionId: 'native', update } });
}

function nativeApprovalMode(currentValue: string) {
	return { id: 'native-mode', category: 'mode', type: 'select', currentValue, options: [{ value: 'accept-edits', name: 'Code' }, { value: 'bypass', name: 'Bypass Permissions' }] };
}

suite('Devin Subscription Agent', () => {
	const store = ensureNoDisposablesAreLeakedInTestSuite();

	function fixture(modelCatalog = catalog, mcpConfiguration?: IDevinMcpConfiguration, logService?: ILogService, configuration?: ConstructorParameters<typeof DevinSubscriptionAgent>[4]) {
		const transports: TestAcpTransport[] = [];
		const launches: { args: readonly string[]; env: NodeJS.ProcessEnv }[] = [];
		const events: AgentSignal[] = [];
		let onPrompt: ((transport: TestAcpTransport, request: Record<string, unknown>) => void) | undefined;
		let onNew: typeof onPrompt;
		let onLoad: typeof onPrompt;
		let onConfig: typeof onPrompt;
		let onInitialize: typeof onPrompt;
		const agent = store.add(new DevinSubscriptionAgent('devin.exe', {
			listModels: async () => modelCatalog,
			createClient: (_cwd, args, env) => {
				const configOptions = [{ id: 'model-choice', category: 'model', type: 'select', currentValue: 'swe-2-high', options: [{ value: 'swe-2-high', name: 'SWE-2' }] }];
				launches.push({ args, env });
				const transport = new TestAcpTransport();
				transports.push(transport);
				transport.handle = request => {
					switch (request.method) {
						case 'initialize':
							if (onInitialize) { onInitialize(transport, request); } else { transport.reply(request, { protocolVersion: 1, agentCapabilities: { loadSession: true } }); }
							break;
						case 'session/new': if (onNew) { onNew(transport, request); } else { transport.reply(request, { sessionId: 'native', configOptions }); } break;
						case 'session/load':
							if (onLoad) { onLoad(transport, request); break; }
							transport.receive({ jsonrpc: '2.0', method: 'session/update', params: { sessionId: 'native', update: { sessionUpdate: 'user_message_chunk', content: { type: 'text', text: 'original' } } } });
							transport.receive({ jsonrpc: '2.0', method: 'session/update', params: { sessionId: 'native', update: { sessionUpdate: 'agent_message_chunk', content: { type: 'text', text: 'restored' } } } });
							transport.reply(request, { configOptions }); break;
						case 'session/set_config_option':
							if (onConfig) { onConfig(transport, request); } else { transport.reply(request, { configOptions }); }
							break;
						case 'session/prompt': if (onPrompt) { onPrompt(transport, request); } else { transport.reply(request, { stopReason: 'end_turn' }); } break;
					}
				};
				return new AcpClient(transport);
			},
		}, mcpConfiguration, logService, configuration));
		store.add(agent.onDidChatProgress(event => events.push(event)));
		return { agent, transports, launches, events, set onPrompt(value: typeof onPrompt) { onPrompt = value; }, set onNew(value: typeof onNew) { onNew = value; }, set onLoad(value: typeof onLoad) { onLoad = value; }, set onConfig(value: typeof onConfig) { onConfig = value; }, set onInitialize(value: typeof onInitialize) { onInitialize = value; } };
	}

	function approvalFixture() {
		const changed = store.add(new Emitter<IAgentChatConfigurationChangeEvent>());
		const rootChanged = store.add(new Emitter<void>());
		const configured = sinon.stub().returns('default');
		const restricted = sinon.stub().returns(false);
		const f = fixture(catalog, undefined, undefined, { getChatEffectiveValue: configured, getRootValue: restricted, onDidChatConfigChange: changed.event, onDidRootConfigChange: rootChanged.event });
		f.onNew = (transport, request) => transport.reply(request, { sessionId: 'native', configOptions: [nativeApprovalMode('accept-edits')] });
		return { f, configured, restricted, changed, rootChanged };
	}

	(isWindows ? test.skip : test)('reports a non-JSON model catalog from the Devin executable as a catalog failure', async () => {
		const directory = await mkdtemp(join(tmpdir(), 'devin-catalog-'));
		try {
			const executable = join(directory, 'devin');
			await writeFile(executable, '#!/bin/sh\necho "Please sign in to continue"\n');
			await chmod(executable, 0o755);
			const agent = store.add(new DevinSubscriptionAgent(executable));
			await assert.rejects(agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token'), /Could not load the Devin model catalog/);
			assert.deepStrictEqual(agent.models.get(), []);
		} finally {
			await rm(directory, { recursive: true, force: true });
		}
	});

	test('applies native bypass before a prompt and restores the previous mode when switched back', async () => {
		const { f, configured } = approvalFixture();
		const config = await f.agent.resolveChatConfig({ config: { autoApprove: 'autoApprove' } });
		assert.deepStrictEqual(config.schema.properties.autoApprove.enum, ['default', 'autoApprove']);
		assert.strictEqual(config.values.autoApprove, 'autoApprove');
		await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
		await f.agent.chats.createChat(chat, owner, { workingDirectories: [cwd] });
		const observed: string[] = [];
		f.onConfig = (transport, request) => {
			assert.ok(isAcpRecord(request.params));
			assert.strictEqual(request.params.configId, 'native-mode');
			const mode = String(request.params.value);
			observed.push(mode);
			transport.reply(request, { configOptions: [nativeApprovalMode(mode)] });
		};
		f.onPrompt = (transport, request) => { observed.push('prompt'); transport.reply(request, { stopReason: 'end_turn' }); };
		configured.returns('autoApprove');
		await f.agent.chats.sendMessage(chat, 'first', [cwd], undefined, 'bypass-turn');
		configured.returns('default');
		await f.agent.chats.sendMessage(chat, 'second', [cwd], undefined, 'default-turn');
		assert.deepStrictEqual(observed, ['bypass', 'prompt', 'accept-edits', 'prompt']);
		assert.deepStrictEqual(configured.firstCall.args[0], { session: owner.toString(), chat: chat.toString() });
	});

	test('does not send a prompt when Devin fails to confirm bypass', async () => {
		const { f, configured } = approvalFixture();
		configured.returns('autoApprove');
		await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
		await f.agent.chats.createChat(chat, owner, { workingDirectories: [cwd] });
		f.onConfig = (transport, request) => transport.reply(request, { configOptions: [nativeApprovalMode('accept-edits')] });
		await assert.rejects(f.agent.chats.sendMessage(chat, 'first', [cwd], undefined, 'bypass-turn'), /Devin did not confirm/);
		assert.ok(!f.transports[0].sent.some(request => request.method === 'session/prompt'));
	});

	test('terminates a bypass turn on policy revocation and restores permissions before retry', async () => {
		const { f, configured, restricted, rootChanged } = approvalFixture();
		configured.returns('autoApprove');
		await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
		await f.agent.chats.createChat(chat, owner, { workingDirectories: [cwd] });
		const prompted = new DeferredPromise<void>();
		f.onPrompt = () => { void prompted.complete(); };
		const modes: string[] = [];
		f.onConfig = (transport, request) => {
			assert.ok(isAcpRecord(request.params));
			const mode = String(request.params.value);
			modes.push(mode);
			transport.reply(request, { configOptions: [nativeApprovalMode(mode)] });
		};
		const rejected = assert.rejects(f.agent.chats.sendMessage(chat, 'first', [cwd], undefined, 'bypass-turn'));
		await prompted.p;
		restricted.returns(true);
		rootChanged.fire();
		await rejected;
		const config = await f.agent.resolveChatConfig({ config: { autoApprove: 'autoApprove' } });
		assert.deepStrictEqual(config.schema.properties[SessionConfigKey.AutoApprove].enum, ['default']);
		assert.strictEqual(config.values.autoApprove, 'default');
		f.onLoad = (transport, request) => transport.reply(request, { configOptions: [nativeApprovalMode('bypass')] });
		f.onPrompt = undefined;
		await f.agent.chats.sendMessage(chat, 'second', [cwd], undefined, 'default-turn');
		assert.deepStrictEqual(modes, ['bypass', 'accept-edits']);
	});

	function clientBundle(name: string, nonce = 'revision'): ClientPluginCustomization {
		return { type: CustomizationType.Plugin, id: name, name, uri: `client-files://window/${name}`, nonce, childEnablement: { tools: [{ kind: CustomizationEnablementKind.Global, enabled: true }] } };
	}

	test('correlates a cancelled Devin send with a local report explaining the MCP change', async () => {
		const reports: string[] = [];
		const logService = new class extends NullLogService {
			override error(message: string | Error): void { reports.push(String(message)); }
		}();
		const f = fixture(catalog, { onDidChange: Event.None, getServers: async (_owner, clients = []) => clients.flatMap(client => client.customizations.map(ref => ({ name: ref.name, command: 'private-command', args: [], env: [] }))) }, logService);
		await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'private-token');
		await f.agent.chats.createChat(chat, owner, { workingDirectories: [cwd] });
		const handle = f.agent.getOrCreateActiveClient(chat, owner, { clientId: 'editor' });
		handle.customizations = [clientBundle('private-tool')];
		const prompted = new DeferredPromise<void>();
		f.onPrompt = () => { void prompted.complete(); };
		const rejected = assert.rejects(f.agent.chats.sendMessage(chat, 'private-prompt', [cwd], undefined, 'turn-for-report'), error => {
			assert.ok(error instanceof Error && isCancellationError(error));
			const reportId = /Devin report ID: (DEVIN-[\da-f-]+)/.exec(error.message)?.[1];
			assert.ok(reportId);
			assert.strictEqual(reports.length, 1);
			const { elapsedMs, ...report } = JSON.parse(reports[0].slice('[Devin Send Report] '.length));
			assert.ok(elapsedMs >= 0);
			assert.deepStrictEqual(report, {
				reportId, turnId: 'turn-for-report', stage: 'prompt', failure: 'cancelled', cancellationReason: 'mcpChanged',
				mcpChange: { trigger: 'clientUpdated', before: 1, after: 0 }, connectionStage: 'ready', mcpServerCount: 1, agentHostPid: process.pid,
			});
			return true;
		});
		await prompted.p;
		handle.customizations = [];
		await rejected;
	});

	for (const { materialized, invalidKind } of [
		{ materialized: false, invalidKind: 'multi-folder' }, { materialized: true, invalidKind: 'multi-folder' },
		{ materialized: false, invalidKind: 'remote-folder' }, { materialized: true, invalidKind: 'remote-folder' },
		{ materialized: true, invalidKind: 'different-folder' },
	] as const) {
		test(`a rejected ${invalidKind} send preserves the ${materialized ? 'running' : 'provisional'} workspace and MCP owner`, async () => {
			const reads: { owner: string; cwd: string | undefined }[] = [];
			const f = fixture(catalog, { onDidChange: Event.None, getServers: async (resource, _clients, directory) => {
				reads.push({ owner: resource.toString(), cwd: directory?.toString() });
				return [{ name: resource.toString() === owner.toString() ? 'original-tools' : 'other-tools', command: 'test-mcp', args: [], env: [] }];
			} });
			await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
			await f.agent.chats.createChat(chat, owner, { workingDirectories: [cwd] });
			if (materialized) { await f.agent.chats.sendMessage(chat, 'first', [cwd], undefined, 'first-valid'); }
			reads.length = 0;
			const otherDirectory = URI.file('/another project');
			const remoteDirectory = URI.parse('vscode-remote://ssh-remote+other/project');
			const rejectedDirectories = invalidKind === 'multi-folder' ? materialized ? [cwd, otherDirectory] : [otherDirectory, cwd]
				: invalidKind === 'remote-folder' ? materialized ? [remoteDirectory] : remoteDirectory : [otherDirectory];
			const otherContext = resolveAgentChatContext(URI.parse('agent-host-devin-subscription:/other-session'), chat);
			await assert.rejects(f.agent.chats.sendMessage(chat, 'invalid', rejectedDirectories, undefined, 'invalid-folders', undefined, otherContext), invalidKind === 'multi-folder' ? /requires one local workspace folder/ : /Cannot move a running Devin chat/);
			const afterRejection = {
				folders: (await f.agent.getChatMetadata(chat, owner))?.workingDirectories?.map(folder => folder.toString()),
				launches: f.launches.length,
				reads: [...reads],
			};
			await f.agent.chats.sendMessage(chat, 'retry', undefined, undefined, 'valid-retry');
			assert.deepStrictEqual({ afterRejection, reads, launches: f.launches.length, prompts: f.transports.flatMap(transport => transport.sent.filter(request => request.method === 'session/prompt')).length }, {
				afterRejection: { folders: [cwd.toString()], launches: materialized ? 1 : 0, reads: [] },
				reads: reads.map(() => ({ owner: owner.toString(), cwd: cwd.toString() })),
				launches: 1,
				prompts: materialized ? 2 : 1,
			});
		});
	}

	test('a valid provisional workspace change applies its new MCP owner and directory before startup', async () => {
		const nextDirectory = URI.file('/another project');
		const nextOwner = URI.parse('agent-host-devin-subscription:/other-session');
		const reads: { owner: string; cwd: string | undefined }[] = [];
		const f = fixture(catalog, { onDidChange: Event.None, getServers: async (resource, _clients, directory) => {
			reads.push({ owner: resource.toString(), cwd: directory?.toString() });
			return [];
		} });
		await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
		await f.agent.chats.createChat(chat, owner, { workingDirectories: [cwd] });
		await f.agent.chats.sendMessage(chat, 'first', [nextDirectory], undefined, 'valid-move', undefined, resolveAgentChatContext(nextOwner, chat));
		const created = f.transports[0].sent.find(request => request.method === 'session/new')!;
		assert.ok(isAcpRecord(created.params));
		assert.ok(reads.length > 0);
		assert.deepStrictEqual({ reads, nativeDirectory: created.params.cwd, folders: (await f.agent.getChatMetadata(chat, owner))?.workingDirectories?.map(folder => folder.toString()), launches: f.launches.length }, {
			reads: reads.map(() => ({ owner: nextOwner.toString(), cwd: nextDirectory.toString() })), nativeDirectory: nextDirectory.fsPath, folders: [nextDirectory.toString()], launches: 1,
		});
	});

	test('keeps an active client handle and forwards its immutable customization snapshot on the first send', async () => {
		const seen: string[] = [];
		const f = fixture(catalog, { onDidChange: Event.None, getServers: async (resource, clients = [], workingDirectory) => {
			assert.strictEqual(resource.toString(), owner.toString());
			assert.strictEqual(workingDirectory?.toString(), cwd.toString());
			return clients.flatMap(client => client.customizations.map(ref => { seen.push(`${client.clientId}:${ref.nonce}`); return { name: ref.name, command: 'test-mcp', args: [ref.nonce!], env: [] }; }));
		} });
		await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
		await f.agent.chats.createChat(chat, owner, { workingDirectories: [cwd] });
		const handle = f.agent.getOrCreateActiveClient(chat, owner, { clientId: 'editor', displayName: 'Editor' });
		assert.strictEqual(f.agent.getOrCreateActiveClient(chat, owner, { clientId: 'editor' }), handle);
		const refs = [clientBundle('client-tools')];
		handle.customizations = refs;
		refs[0].nonce = 'mutated-after-assignment';
		handle.customizations[0].nonce = 'mutated-through-getter';
		await f.agent.chats.sendMessage(chat, 'hello', [cwd], undefined, 'client-first');
		const request = f.transports[0].sent.find(item => item.method === 'session/new')!;
		assert.ok(isAcpRecord(request.params));
		assert.deepStrictEqual(request.params.mcpServers, [{ name: 'client-tools', command: 'test-mcp', args: ['revision'], env: [] }]);
		assert.ok(seen.length > 0 && seen.every(value => value === 'editor:revision'));
	});

	test('retains client contributions before materialization and isolates exact chats', async () => {
		const secondChat = URI.parse('different-chat:/opaque-chat');
		const f = fixture(catalog, { onDidChange: Event.None, getServers: async (_owner, clients = []) => clients.flatMap(client => client.customizations.map(ref => ({ name: ref.name, command: 'test-mcp', args: [], env: [] }))) });
		await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
		f.agent.getOrCreateActiveClient(chat, owner, { clientId: 'editor' }).customizations = [clientBundle('first-tools')];
		f.agent.getOrCreateActiveClient(secondChat, owner, { clientId: 'editor' }).customizations = [clientBundle('second-tools')];
		await f.agent.chats.createChat(chat, owner, { workingDirectories: [cwd] });
		await f.agent.chats.createChat(secondChat, owner, { workingDirectories: [cwd] });
		await f.agent.chats.sendMessage(chat, 'first', [cwd], undefined, 'first-client-chat');
		await f.agent.chats.sendMessage(secondChat, 'second', [cwd], undefined, 'second-client-chat');
		for (const [index, name] of ['first-tools', 'second-tools'].entries()) {
			const request = f.transports[index].sent.find(item => item.method === 'session/new')!;
			assert.ok(isAcpRecord(request.params));
			assert.deepStrictEqual(request.params.mcpServers, [{ name, command: 'test-mcp', args: [], env: [] }]);
		}
	});

	test('revokes a removed client during a turn and ignores later updates from its detached handle', async () => {
		const f = fixture(catalog, { onDidChange: Event.None, getServers: async (_owner, clients = []) => clients.flatMap(client => client.customizations.map(ref => ({ name: ref.name, command: 'test-mcp', args: [], env: [] }))) });
		await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
		await f.agent.chats.createChat(chat, owner, { workingDirectories: [cwd] });
		const handle = f.agent.getOrCreateActiveClient(chat, owner, { clientId: 'editor' });
		handle.customizations = [clientBundle('old-tools')];
		const started = new DeferredPromise<void>();
		f.onPrompt = () => { void started.complete(); };
		const firstSettled = f.agent.chats.sendMessage(chat, 'waiting', [cwd], undefined, 'client-before-remove').catch(error => error);
		await started.p;
		f.agent.removeActiveClient(chat, owner, 'editor');
		handle.customizations = [clientBundle('must-not-return')];
		await firstSettled;
		f.onPrompt = (transport, request) => transport.reply(request, { stopReason: 'end_turn' });
		await f.agent.chats.sendMessage(chat, 'follow-up', [cwd], undefined, 'client-after-remove');
		const restored = f.transports[1].sent.find(item => item.method === 'session/load')!;
		assert.ok(isAcpRecord(restored.params));
		assert.deepStrictEqual(restored.params.mcpServers, []);
		assert.strictEqual(restored.params.sessionId, 'native');
	});

	for (const failPrevious of [false, true]) {
		test(`refreshes a stale client snapshot before opening the native session without dropping the first send (old read fails=${failPrevious})`, async () => {
			const loading = new DeferredPromise<void>();
			const previous = new DeferredPromise<readonly AcpMcpServer[]>();
			const f = fixture(catalog, { onDidChange: Event.None, getServers: async (_owner, clients = []) => {
				if (clients.some(client => client.customizations.some(ref => ref.nonce === 'old'))) {
					void loading.complete();
					const result = await previous.p;
					if (failPrevious) { throw new Error('Old bundle is unavailable'); }
					return result;
				}
				return [{ name: 'current-tools', command: 'test-mcp', args: [], env: [] }];
			} });
			await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
			await f.agent.chats.createChat(chat, owner, { workingDirectories: [cwd] });
			const handle = f.agent.getOrCreateActiveClient(chat, owner, { clientId: 'editor' });
			handle.customizations = [clientBundle('tools', 'old')];
			const send = f.agent.chats.sendMessage(chat, 'first', [cwd], undefined, 'client-stale');
			await loading.p;
			handle.customizations = [clientBundle('tools', 'new')];
			void previous.complete([]);
			await send;
			assert.strictEqual(f.launches.length, 1);
			const created = f.transports[0].sent.find(item => item.method === 'session/new')!;
			assert.ok(isAcpRecord(created.params));
			assert.deepStrictEqual(created.params.mcpServers, [{ name: 'current-tools', command: 'test-mcp', args: [], env: [] }]);
		});
	}

	test('forwards enabled host MCP definitions to both new and restored native sessions', async () => {
		const server: AcpMcpServer = { name: 'editor-tools', type: 'http', url: 'http://127.0.0.1:43123/mcp', headers: [{ name: 'Authorization', value: 'test-only-header' }] };
		const owners: string[] = [];
		const f = fixture(catalog, { onDidChange: Event.None, getServers: async resource => { owners.push(resource.toString()); return [server]; } });
		f.onInitialize = (transport, request) => transport.reply(request, { protocolVersion: 1, agentCapabilities: { loadSession: true, mcpCapabilities: { http: true } } });
		await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
		await f.agent.chats.createChat(chat, owner, { workingDirectories: [cwd] });
		await f.agent.chats.sendMessage(chat, 'first', [cwd], undefined, 'mcp-first');
		await f.agent.chats.releaseChat(chat, owner);
		await f.agent.chats.sendMessage(chat, 'second', [cwd], undefined, 'mcp-second');
		for (const [index, method] of ['session/new', 'session/load'].entries()) {
			const request = f.transports[index].sent.find(item => item.method === method)!;
			assert.ok(isAcpRecord(request.params));
			assert.deepStrictEqual(request.params.mcpServers, [server]);
		}
		assert.ok(owners.every(resource => resource === owner.toString()));
		assert.ok(!f.launches.some(launch => launch.args.some(argument => argument.includes('test-only-header'))));
	});

	test('refuses an unadvertised MCP transport before creating or prompting a native session', async () => {
		const f = fixture(catalog, { onDidChange: Event.None, getServers: async () => [{ name: 'tools', type: 'http', url: 'https://tools.example/mcp', headers: [] }] });
		await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
		await f.agent.chats.createChat(chat, owner, { workingDirectories: [cwd] });
		await assert.rejects(f.agent.chats.sendMessage(chat, 'message', [cwd], undefined, 'mcp-unavailable'), /did not advertise/);
		assert.ok(!f.transports[0].sent.some(item => item.method === 'session/new' || item.method === 'session/prompt'));
	});

	test('removes changed MCP access from the active chat and reconnects its native session for a follow-up', async () => {
		const changed = store.add(new Emitter<readonly string[] | undefined>());
		let servers: readonly AcpMcpServer[] = [{ name: 'tools', command: 'test-mcp', args: [], env: [] }];
		const f = fixture(catalog, { onDidChange: changed.event, getServers: async () => servers });
		await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
		await f.agent.chats.createChat(chat, owner, { workingDirectories: [cwd] });
		const started = new DeferredPromise<void>();
		f.onPrompt = () => { void started.complete(); };
		const first = f.agent.chats.sendMessage(chat, 'waiting', [cwd], undefined, 'mcp-old');
		const firstSettled = first.catch(error => error);
		await started.p;
		servers = [];
		changed.fire([owner.toString()]);
		await firstSettled;
		f.onPrompt = (transport, request) => transport.reply(request, { stopReason: 'end_turn' });
		await f.agent.chats.sendMessage(chat, 'follow-up', [cwd], undefined, 'mcp-new');
		const restored = f.transports[1].sent.find(item => item.method === 'session/load')!;
		assert.ok(isAcpRecord(restored.params));
		assert.deepStrictEqual(restored.params.mcpServers, []);
		assert.strictEqual(restored.params.sessionId, 'native');
	});

	test('does not disturb an active turn for an unchanged MCP list or a different owner', async () => {
		const changed = store.add(new Emitter<readonly string[] | undefined>());
		const f = fixture(catalog, { onDidChange: changed.event, getServers: async () => [] });
		await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
		await f.agent.chats.createChat(chat, owner, { workingDirectories: [cwd] });
		const started = new DeferredPromise<void>();
		let pending: Record<string, unknown> | undefined;
		f.onPrompt = (_transport, request) => { pending = request; void started.complete(); };
		const send = f.agent.chats.sendMessage(chat, 'waiting', [cwd], undefined, 'mcp-unchanged');
		await started.p;
		changed.fire(['other-host:/different-owner']);
		changed.fire([owner.toString()]);
		await Promise.resolve();
		await Promise.resolve();
		f.transports[0].reply(pending!, { stopReason: 'end_turn' });
		await send;
		assert.strictEqual(f.launches.length, 1);
	});

	test('rechecks MCP definitions after initialization rather than sending a stale startup snapshot', async () => {
		let servers: readonly AcpMcpServer[] = [{ name: 'tools', command: 'test-mcp', args: [], env: [] }];
		const f = fixture(catalog, { onDidChange: Event.None, getServers: async () => servers });
		f.onInitialize = (transport, request) => { servers = []; transport.reply(request, { protocolVersion: 1, agentCapabilities: { loadSession: true } }); };
		await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
		await f.agent.chats.createChat(chat, owner, { workingDirectories: [cwd] });
		await assert.rejects(f.agent.chats.sendMessage(chat, 'message', [cwd], undefined, 'mcp-stale'), isCancellationError);
		assert.ok(!f.transports[0].sent.some(item => item.method === 'session/new' || item.method === 'session/prompt'));
		await f.agent.chats.sendMessage(chat, 'retry', [cwd], undefined, 'mcp-fresh');
		const request = f.transports[1].sent.find(item => item.method === 'session/new')!;
		assert.ok(isAcpRecord(request.params));
		assert.deepStrictEqual(request.params.mcpServers, []);
	});

	test('reconciles MCP access when a live owner reference changes without a configuration event', async () => {
		const ownerChanged = store.add(new Emitter<IAgentChatOwnerChange>());
		let currentOwner = owner;
		const movedOwner = URI.parse('different-host:/opaque-owner');
		const context = { ...resolveAgentChatContext(owner, chat), owner: { ...fixedChatOwnerRef(owner), get: () => currentOwner, onDidChange: ownerChanged.event } };
		const f = fixture(catalog, { onDidChange: Event.None, getServers: async resource => resource.toString() === owner.toString() ? [{ name: 'tools', command: 'test-mcp', args: [], env: [] }] : [] });
		await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
		await f.agent.chats.createChat(chat, context, { workingDirectories: [cwd] });
		const started = new DeferredPromise<void>();
		f.onPrompt = () => { void started.complete(); };
		const firstSettled = f.agent.chats.sendMessage(chat, 'waiting', [cwd], undefined, 'mcp-before-move').catch(error => error);
		await started.p;
		currentOwner = movedOwner;
		ownerChanged.fire({ previous: owner, current: movedOwner });
		await firstSettled;
		f.onPrompt = (transport, request) => transport.reply(request, { stopReason: 'end_turn' });
		await f.agent.chats.sendMessage(chat, 'after move', [cwd], undefined, 'mcp-after-move');
		const restored = f.transports[1].sent.find(item => item.method === 'session/load')!;
		assert.ok(isAcpRecord(restored.params));
		assert.deepStrictEqual(restored.params.mcpServers, []);
	});

	test('does not spawn a native client after its owner changes while MCP decisions load', async () => {
		let currentOwner = owner;
		const context = { ...resolveAgentChatContext(owner, chat), owner: { ...fixedChatOwnerRef(owner), get: () => currentOwner } };
		const loading = new DeferredPromise<void>();
		const decisions = new DeferredPromise<readonly AcpMcpServer[]>();
		const f = fixture(catalog, { onDidChange: Event.None, getServers: async () => { void loading.complete(); return decisions.p; } });
		await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
		await f.agent.chats.createChat(chat, context, { workingDirectories: [cwd] });
		const rejected = assert.rejects(f.agent.chats.sendMessage(chat, 'message', [cwd], undefined, 'mcp-loading'), isCancellationError);
		await loading.p;
		currentOwner = URI.parse('different-host:/opaque-owner');
		void decisions.complete([]);
		await rejected;
		assert.strictEqual(f.launches.length, 0);
	});

	test('Stop cancels MCP preparation before spawning a native client', async () => {
		const loading = new DeferredPromise<void>();
		const decisions = new DeferredPromise<readonly AcpMcpServer[]>();
		const f = fixture(catalog, { onDidChange: Event.None, getServers: async () => { void loading.complete(); return decisions.p; } });
		await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
		await f.agent.chats.createChat(chat, owner, { workingDirectories: [cwd] });
		const rejected = assert.rejects(f.agent.chats.sendMessage(chat, 'message', [cwd], undefined, 'mcp-stopped'), isCancellationError);
		await loading.p;
		await f.agent.chats.abort(chat, owner);
		void decisions.complete([]);
		await rejected;
		assert.strictEqual(f.launches.length, 0);
	});

	test('keeps subscription credentials separate from inherited CLI credentials and parses a live catalog shape', () => {
		const inherited = {
			WINDSURF_API_KEY: 'other-account',
			windsurf_api_key: 'case-variant',
			WINDSURF_EXT_HOST_PID: '79108',
			windsurf_ext_host_pid: '16684',
			WINDSURF_API_SERVER_URL: 'https://windsurf.example',
			DEVIN_API_KEY: 'foreign-key',
			devin_model: 'swe-2-high',
			DEVIN_API_URL: 'https://override.example',
			ACP_BACKEND: 'openai',
			PATH: 'keep',
		};
		assert.deepStrictEqual(devinSubscriptionEnvironment('browser-token', inherited), { PATH: 'keep', WINDSURF_API_KEY: 'devin-session-token$browser-token' });
		assert.deepStrictEqual(devinSubscriptionEnvironment('browser-token', inherited, 4242), { PATH: 'keep', WINDSURF_API_KEY: 'devin-session-token$browser-token', WINDSURF_EXT_HOST_PID: '4242' });
		assert.strictEqual(inherited.WINDSURF_API_KEY, 'other-account');
		assert.strictEqual(inherited.WINDSURF_EXT_HOST_PID, '79108');
		assert.strictEqual(parseDevinModels(catalog)[0].id, 'swe-2-high');
		assert.strictEqual(parseDevinModels(catalog)[0].maxContextWindow, 262000);
	});

	test('isolates ACP child processes from inherited extension-host PIDs and foreign Devin env overrides', async () => {
		const savedExtHostPid = process.env.WINDSURF_EXT_HOST_PID;
		const savedDevinModel = process.env.DEVIN_MODEL;
		const savedAcpBackend = process.env.ACP_BACKEND;
		process.env.WINDSURF_EXT_HOST_PID = '79108';
		process.env.DEVIN_MODEL = 'override-model';
		process.env.ACP_BACKEND = 'openai';
		try {
			const f = fixture();
			await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
			await f.agent.chats.createChat(chat, owner, { workingDirectories: [cwd] });
			await f.agent.chats.sendMessage(chat, 'prompt', [cwd], undefined, 'isolated-env-turn');
			assert.strictEqual(f.launches[0].env.WINDSURF_EXT_HOST_PID, String(process.pid));
			assert.strictEqual(f.launches[0].env.DEVIN_MODEL, undefined);
			assert.strictEqual(f.launches[0].env.ACP_BACKEND, undefined);
		} finally {
			if (savedExtHostPid === undefined) { delete process.env.WINDSURF_EXT_HOST_PID; } else { process.env.WINDSURF_EXT_HOST_PID = savedExtHostPid; }
			if (savedDevinModel === undefined) { delete process.env.DEVIN_MODEL; } else { process.env.DEVIN_MODEL = savedDevinModel; }
			if (savedAcpBackend === undefined) { delete process.env.ACP_BACKEND; } else { process.env.ACP_BACKEND = savedAcpBackend; }
		}
	});

	test('loads models only after the editor forwards a subscription account and clears them on sign-out', async () => {
		const { agent, launches } = fixture();
		assert.deepStrictEqual(agent.models.get(), []);
		await agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
		assert.strictEqual(agent.models.get().length, 1);
		assert.strictEqual(launches.length, 0);
		await agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, '');
		assert.deepStrictEqual(agent.models.get(), []);
	});

	test('defers the client, uses the selected model, and streams to the exact host turn', async () => {
		const f = fixture();
		await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
		await f.agent.chats.createChat(chat, owner, { workingDirectories: [cwd], model: { id: 'swe-2-high' } });
		assert.strictEqual(f.launches.length, 0);
		f.onPrompt = (transport, request) => {
			transport.receive({ jsonrpc: '2.0', method: 'session/update', params: { sessionId: 'different-session', update: { sessionUpdate: 'agent_message_chunk', content: { type: 'text', text: 'wrong' } } } });
			for (const text of ['hello ', 'world']) { transport.receive({ jsonrpc: '2.0', method: 'session/update', params: { sessionId: 'native', update: { sessionUpdate: 'agent_message_chunk', content: { type: 'text', text } } } }); }
			transport.reply(request, { stopReason: 'end_turn' });
		};
		await f.agent.chats.sendMessage(chat, 'prompt', [cwd], undefined, 'host-turn');
		assert.deepStrictEqual(f.launches[0].args, ['acp', '--model', 'swe-2-high']);
		assert.strictEqual(f.launches[0].env.WINDSURF_API_KEY, 'devin-session-token$browser-token');
		const deltas = f.events.filter(event => event.kind === 'action' && event.action.type === ActionType.ChatDelta);
		assert.strictEqual(deltas.length, 2);
		assert.strictEqual(f.events.at(-1)?.kind, 'action');
		const turns = await f.agent.chats.getMessages(chat, owner);
		assert.strictEqual(turns[0].id, 'host-turn');
		assert.strictEqual(turns[0].responseParts.length, 1);
	});

	test('waits for a host permission decision and never selects an always-allow option implicitly', async () => {
		const f = fixture();
		await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
		await f.agent.chats.createChat(chat, owner, { workingDirectories: [cwd] });
		let promptRequest: Record<string, unknown>;
		f.onPrompt = (transport, request) => {
			promptRequest = request;
			transport.receive({ jsonrpc: '2.0', id: request.id, method: 'session/request_permission', params: { sessionId: 'native', toolCall: { toolCallId: 'tool', title: 'Run command', rawInput: { command: 'read', path: 'README.md' } }, options: [{ optionId: 'forever', kind: 'allow_always' }, { optionId: 'once', kind: 'allow_once' }, { optionId: 'deny', kind: 'reject_once' }] } });
		};
		const send = f.agent.chats.sendMessage(chat, 'prompt', [cwd], undefined, 'host-turn');
		await new Promise<void>(resolve => {
			const listener = f.agent.onDidChatProgress(event => { if (event.kind === 'pending_confirmation') { listener.dispose(); resolve(); } });
			store.add(listener);
		});
		const pending = f.events.find(event => event.kind === 'pending_confirmation');
		assert.ok(pending?.kind === 'pending_confirmation');
		assert.strictEqual(pending.state.confirmationTitle, 'Allow Devin to run this tool?');
		assert.deepStrictEqual(pending.state.options?.map(option => option.id), ['allow-once', 'skip']);
		assert.ok(!f.events.some(event => event.kind === 'action' && event.action.type === ActionType.ChatToolCallReady), 'The host publishes the ready action after processing the confirmation');
		const transport = f.transports[0];
		toolUpdate(transport, { sessionUpdate: 'tool_call_update', toolCallId: 'tool', rawInput: { path: 'README.md', command: 'read' } });
		assert.ok(!transport.sent.some(message => Object.hasOwn(message, 'result')));
		const replied = Event.toPromise(Event.filter(transport.onSend, message => Object.hasOwn(message, 'result')));
		f.agent.respondToPermissionRequest('host-turn:tool', true);
		await replied;
		const response = transport.sent.find(message => Object.hasOwn(message, 'result'));
		assert.deepStrictEqual(response?.result, { outcome: { outcome: 'selected', optionId: 'once' } });
		toolUpdate(transport, { sessionUpdate: 'tool_call_update', toolCallId: 'tool', status: 'in_progress' });
		toolUpdate(transport, { sessionUpdate: 'tool_call_update', toolCallId: 'tool', status: 'completed', rawOutput: 'done' });
		transport.reply(promptRequest!, { stopReason: 'end_turn' });
		await send;
		const approved = (await f.agent.chats.getMessages(chat, owner))[0].responseParts[0];
		assert.ok(approved.kind === ResponsePartKind.ToolCall && approved.toolCall.status === ToolCallStatus.Completed);
		assert.strictEqual(approved.toolCall.confirmed, ToolCallConfirmationReason.UserAction);
		assert.strictEqual(approved.toolCall.success, true);
	});

	test('retains intermediate tool output and full input when the completion omits them', async () => {
		const f = fixture();
		await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
		await f.agent.chats.createChat(chat, owner, { workingDirectories: [cwd] });
		f.onPrompt = (transport, request) => {
			toolUpdate(transport, { sessionUpdate: 'tool_call', toolCallId: 'read', name: 'read_file', title: 'Read file', status: 'pending', rawInput: { path: 'D:/project with spaces/README.md' } });
			toolUpdate(transport, { sessionUpdate: 'tool_call_update', toolCallId: 'read', status: 'in_progress', content: [{ type: 'content', content: { type: 'text', text: 'file content 🌎' } }] });
			toolUpdate(transport, { sessionUpdate: 'tool_call_update', toolCallId: 'read', status: 'completed' });
			transport.reply(request, { stopReason: 'end_turn' });
		};
		await f.agent.chats.sendMessage(chat, 'read', [cwd], undefined, 'host-turn');
		const part = (await f.agent.chats.getMessages(chat, owner))[0].responseParts[0];
		assert.ok(part.kind === ResponsePartKind.ToolCall && part.toolCall.status === ToolCallStatus.Completed);
		assert.strictEqual(part.toolCall.toolInput, JSON.stringify({ path: 'D:/project with spaces/README.md' }, undefined, 2));
		assert.deepStrictEqual(part.toolCall.content, [{ type: ToolResultContentType.Text, text: 'file content 🌎' }]);
		assert.strictEqual(f.events.filter(event => event.kind === 'action' && event.action.type === ActionType.ChatToolCallContentChanged).length, 1);
	});

	for (const { label, name, metadata, expected } of [
		{ label: 'standard field takes precedence', name: 'standard-tool', metadata: { 'cognition.ai/toolName': 'native-tool' }, expected: 'standard-tool' },
		{ label: 'native MCP identity', name: undefined, metadata: { 'cognition.ai/toolName': 'mcp__probe-server__echo_probe', 'cognition.ai/inferenceToolName': 'mcp_call_tool' }, expected: 'mcp__probe-server__echo_probe' },
		{ label: 'built-in inference identity', name: undefined, metadata: { 'cognition.ai/inferenceToolName': 'mcp_list_tools' }, expected: 'mcp_list_tools' },
		{ label: 'missing extension', name: undefined, metadata: undefined, expected: 'devin-tool' },
		{ label: 'malformed extension', name: undefined, metadata: { 'cognition.ai/toolName': { name: 'invalid' }, 'cognition.ai/inferenceToolName': 42 }, expected: 'devin-tool' },
		{ label: 'blank names use the next valid field', name: '', metadata: { 'cognition.ai/toolName': '  ', 'cognition.ai/inferenceToolName': 'read_file' }, expected: 'read_file' },
		{ label: 'unrelated metadata is ignored', name: undefined, metadata: { toolName: 'untrusted-generic-key' }, expected: 'devin-tool' },
	]) {
		test(`preserves Devin tool identity in live actions and history: ${label}`, async () => {
			const f = fixture();
			await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
			await f.agent.chats.createChat(chat, owner, { workingDirectories: [cwd] });
			f.onPrompt = (transport, request) => {
				toolUpdate(transport, { sessionUpdate: 'tool_call', toolCallId: 'probe', title: 'Echo probe', name, _meta: metadata, rawInput: { nonce: 'test-only' } });
				toolUpdate(transport, { sessionUpdate: 'tool_call_update', toolCallId: 'probe', status: 'in_progress' });
				toolUpdate(transport, { sessionUpdate: 'tool_call_update', toolCallId: 'probe', status: 'completed', content: [{ type: 'content', content: { type: 'text', text: 'confirmed' } }] });
				transport.reply(request, { stopReason: 'end_turn' });
			};
			await f.agent.chats.sendMessage(chat, 'echo', [cwd], undefined, 'identity-turn');
			const start = f.events.find(event => event.kind === 'action' && event.action.type === ActionType.ChatToolCallStart);
			assert.ok(start?.kind === 'action' && start.action.type === ActionType.ChatToolCallStart);
			assert.strictEqual(start.action.toolName, expected);
			const part = (await f.agent.chats.getMessages(chat, owner))[0].responseParts[0];
			assert.ok(part.kind === ResponsePartKind.ToolCall && part.toolCall.status === ToolCallStatus.Completed);
			assert.strictEqual(part.toolCall.toolName, expected);
			assert.strictEqual(part.toolCall.displayName, 'Echo probe');
			assert.deepStrictEqual(part.toolCall.content, [{ type: ToolResultContentType.Text, text: 'confirmed' }]);
			assert.strictEqual(part.toolCall.contributor, undefined);
		});
	}

	test('retains the observed MCP identity, title, and arguments when a native permission request carries only the tool id', async () => {
		const f = fixture();
		await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
		await f.agent.chats.createChat(chat, owner, { workingDirectories: [cwd] });
		let promptRequest: Record<string, unknown>;
		f.onPrompt = (transport, request) => {
			promptRequest = request;
			toolUpdate(transport, { sessionUpdate: 'tool_call', toolCallId: 'probe', title: 'Calling echo_probe from probe-server', rawInput: { nonce: 'test-only' }, _meta: { 'cognition.ai/toolName': 'mcp__probe-server__echo_probe' } });
			transport.receive({ jsonrpc: '2.0', id: 'permission', method: 'session/request_permission', params: { sessionId: 'native', toolCall: { toolCallId: 'probe', _meta: { 'cognition.ai/inferenceToolName': 'mcp__probe-server__echo_probe' } }, options: [{ optionId: 'once', kind: 'allow_once' }] } });
		};
		const waiting = Event.toPromise(Event.filter(f.agent.onDidChatProgress, event => event.kind === 'pending_confirmation'));
		const send = f.agent.chats.sendMessage(chat, 'echo', [cwd], undefined, 'host-turn');
		const pending = await waiting;
		assert.ok(pending.kind === 'pending_confirmation');
		assert.strictEqual(pending.state.toolName, 'mcp__probe-server__echo_probe');
		assert.strictEqual(pending.state.displayName, 'Calling echo_probe from probe-server');
		assert.strictEqual(pending.state.toolInput, JSON.stringify({ nonce: 'test-only' }, undefined, 2));
		const transport = f.transports[0];
		const replied = Event.toPromise(Event.filter(transport.onSend, message => Object.hasOwn(message, 'result')));
		f.agent.respondToPermissionRequest('host-turn:probe', true);
		await replied;
		assert.deepStrictEqual(transport.sent.find(message => Object.hasOwn(message, 'result'))?.result, { outcome: { outcome: 'selected', optionId: 'once' } });
		toolUpdate(transport, { sessionUpdate: 'tool_call_update', toolCallId: 'probe', status: 'completed', content: [{ type: 'content', content: { type: 'text', text: 'confirmed' } }] });
		transport.reply(promptRequest!, { stopReason: 'end_turn' });
		await send;
		const part = (await f.agent.chats.getMessages(chat, owner))[0].responseParts[0];
		assert.ok(part.kind === ResponsePartKind.ToolCall && part.toolCall.status === ToolCallStatus.Completed);
		assert.strictEqual(part.toolCall.confirmed, ToolCallConfirmationReason.UserAction);
	});

	test('restores native Devin MCP tool identity and result from ACP history replay without emitting live tool actions', async () => {
		const f = fixture();
		await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
		await f.agent.materializeChat(chat, owner, JSON.stringify({ version: 1, sessionId: 'native', cwd: cwd.toString(), createdAt: 1 }));
		f.onLoad = (transport, request) => {
			toolUpdate(transport, { sessionUpdate: 'user_message_chunk', content: { type: 'text', text: 'echo' } });
			toolUpdate(transport, { sessionUpdate: 'tool_call', toolCallId: 'probe', title: 'Calling echo_probe from probe-server', rawInput: { nonce: 'test-only' }, _meta: { 'cognition.ai/toolName': 'mcp__probe-server__echo_probe' } });
			toolUpdate(transport, { sessionUpdate: 'tool_call_update', toolCallId: 'probe', status: 'completed', content: [{ type: 'content', content: { type: 'text', text: 'confirmed' } }] });
			transport.reply(request, {});
		};
		const part = (await f.agent.chats.getMessages(chat, owner))[0].responseParts[0];
		assert.ok(part.kind === ResponsePartKind.ToolCall && part.toolCall.status === ToolCallStatus.Completed);
		assert.strictEqual(part.toolCall.toolName, 'mcp__probe-server__echo_probe');
		assert.deepStrictEqual(part.toolCall.content, [{ type: ToolResultContentType.Text, text: 'confirmed' }]);
		assert.ok(!f.events.some(event => event.kind === 'action' && event.action.type === ActionType.ChatToolCallStart));
	});

	test('keeps the latest raw output and updates running metadata without losing the result', async () => {
		const f = fixture();
		await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
		await f.agent.chats.createChat(chat, owner, { workingDirectories: [cwd] });
		f.onPrompt = (transport, request) => {
			toolUpdate(transport, { sessionUpdate: 'tool_call', toolCallId: 'run', title: 'Run', status: 'in_progress', rawInput: 'old command', rawOutput: 'first output' });
			toolUpdate(transport, { sessionUpdate: 'tool_call_update', toolCallId: 'run', title: 'Read public file', rawInput: 'new command', rawOutput: 'latest output' });
			toolUpdate(transport, { sessionUpdate: 'tool_call_update', toolCallId: 'run', status: 'completed', rawInput: null, rawOutput: null });
			transport.reply(request, { stopReason: 'end_turn' });
		};
		await f.agent.chats.sendMessage(chat, 'read', [cwd], undefined, 'host-turn');
		const part = (await f.agent.chats.getMessages(chat, owner))[0].responseParts[0];
		assert.ok(part.kind === ResponsePartKind.ToolCall && part.toolCall.status === ToolCallStatus.Completed);
		assert.strictEqual(part.toolCall.toolInput, 'new command');
		assert.strictEqual(part.toolCall.invocationMessage, 'Read public file');
		assert.deepStrictEqual(part.toolCall.content, [{ type: ToolResultContentType.Text, text: 'latest output' }]);
	});

	test('finalizes a running tool with its partial result when the prompt transport fails', async () => {
		const f = fixture();
		await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
		await f.agent.chats.createChat(chat, owner, { workingDirectories: [cwd] });
		f.onPrompt = (transport, _request) => {
			toolUpdate(transport, { sessionUpdate: 'tool_call', toolCallId: 'read', title: 'Read file', status: 'in_progress', rawOutput: 'partial content' });
			transport.close();
		};
		await assert.rejects(f.agent.chats.sendMessage(chat, 'read', [cwd], undefined, 'host-turn'), /peer exited/);
		const complete = f.events.flatMap(event => event.kind === 'action' && event.action.type === ActionType.ChatToolCallComplete ? [event.action] : []);
		assert.strictEqual(complete.length, 1);
		assert.strictEqual(complete[0].result.success, false);
		assert.deepStrictEqual(complete[0].result.content, [{ type: ToolResultContentType.Text, text: 'partial content' }]);
	});

	test('preserves a denial as cancelled in history and ignores a late successful tool update', async () => {
		const f = fixture();
		await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
		await f.agent.chats.createChat(chat, owner, { workingDirectories: [cwd] });
		let promptRequest: Record<string, unknown>;
		f.onPrompt = (transport, request) => {
			promptRequest = request;
			transport.receive({ jsonrpc: '2.0', id: 'permission', method: 'session/request_permission', params: { sessionId: 'native', toolCall: { toolCallId: 'read', title: 'Read file', rawInput: { path: 'README.md' } }, options: [{ optionId: 'deny', kind: 'reject_once' }] } });
		};
		const pending = Event.toPromise(Event.filter(f.agent.onDidChatProgress, event => event.kind === 'pending_confirmation'));
		const send = f.agent.chats.sendMessage(chat, 'read', [cwd], undefined, 'host-turn');
		await pending;
		f.agent.respondToPermissionRequest('host-turn:read', false);
		const transport = f.transports[0];
		toolUpdate(transport, { sessionUpdate: 'tool_call_update', toolCallId: 'read', status: 'completed', rawOutput: 'late output' });
		transport.reply(promptRequest!, { stopReason: 'end_turn' });
		await send;
		const part = (await f.agent.chats.getMessages(chat, owner))[0].responseParts[0];
		assert.ok(part.kind === ResponsePartKind.ToolCall && part.toolCall.status === ToolCallStatus.Cancelled);
		assert.strictEqual(part.toolCall.reason, ToolCallCancellationReason.Denied);
		assert.strictEqual(part.toolCall.toolInput, JSON.stringify({ path: 'README.md' }, undefined, 2));
	});

	test('cancels a permission when its input changes and does not accept a stale approval', async () => {
		const f = fixture();
		let host: ChatState = { resource: chat.toString(), title: 'Test', status: SessionStatus.Idle, modifiedAt: new Date(0).toISOString(), turns: [], activeTurn: { id: 'host-turn', startedAt: new Date(0).toISOString(), message: { text: 'read', origin: { kind: MessageKind.User } }, responseParts: [], usage: undefined } };
		store.add(f.agent.onDidChatProgress(event => { if (event.kind === 'action' && isChatAction(event.action)) { host = chatReducer(host, event.action); } }));
		await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
		await f.agent.chats.createChat(chat, owner, { workingDirectories: [cwd] });
		let promptRequest: Record<string, unknown>;
		f.onPrompt = (transport, request) => {
			promptRequest = request;
			transport.receive({ jsonrpc: '2.0', id: 'permission', method: 'session/request_permission', params: { sessionId: 'native', toolCall: { toolCallId: 'run', title: 'Read file', rawInput: { command: 'read README.md' } }, options: [{ optionId: 'once', kind: 'allow_once' }] } });
		};
		const pending = Event.toPromise(Event.filter(f.agent.onDidChatProgress, event => event.kind === 'pending_confirmation'));
		const send = f.agent.chats.sendMessage(chat, 'read', [cwd], undefined, 'host-turn');
		await pending;
		const transport = f.transports[0];
		const response = Event.toPromise(Event.filter(transport.onSend, message => message.id === 'permission' && Object.hasOwn(message, 'result')));
		toolUpdate(transport, { sessionUpdate: 'tool_call_update', toolCallId: 'run', rawInput: { command: 'different command' } });
		f.agent.respondToPermissionRequest('host-turn:run', true);
		assert.deepStrictEqual((await response).result, { outcome: { outcome: 'cancelled' } });
		assert.ok(f.events.some(event => event.kind === 'action' && event.action.type === ActionType.ChatToolCallConfirmed && !event.action.approved));
		const cancelled = host.activeTurn?.responseParts[0];
		assert.ok(cancelled?.kind === ResponsePartKind.ToolCall && cancelled.toolCall.status === ToolCallStatus.Cancelled);
		assert.strictEqual(cancelled.toolCall.reason, ToolCallCancellationReason.Skipped);
		transport.reply(promptRequest!, { stopReason: 'end_turn' });
		await send;
		assert.strictEqual(host.activeTurn, undefined);
		assert.strictEqual(host.turns.length, 1);
	});

	test('keeps an aborted tool turn cancelled even when the process closes instead of replying', async () => {
		const f = fixture();
		await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
		await f.agent.chats.createChat(chat, owner, { workingDirectories: [cwd] });
		f.onPrompt = transport => {
			toolUpdate(transport, { sessionUpdate: 'tool_call', toolCallId: 'read', title: 'Read file', status: 'in_progress' });
			store.add(transport.onSend(message => { if (message.method === 'session/cancel') { transport.close(); } }));
		};
		const started = Event.toPromise(Event.filter(f.agent.onDidChatProgress, event => event.kind === 'action' && event.action.type === ActionType.ChatToolCallStart));
		const send = f.agent.chats.sendMessage(chat, 'read', [cwd], undefined, 'host-turn');
		const rejected = assert.rejects(send, /peer exited/);
		await started;
		const history = await f.agent.chats.getMessages(chat, owner);
		await f.agent.chats.abort(chat, owner);
		await rejected;
		assert.strictEqual(history[0].state, TurnState.Cancelled);
		const completed = f.events.flatMap(event => event.kind === 'action' && event.action.type === ActionType.ChatToolCallComplete ? [event.action] : []);
		assert.strictEqual(completed.length, 1);
		assert.strictEqual(completed[0].result.success, false);
	});

	test('repeated Stop does not leave a deadline that kills the next turn', async () => {
		const f = fixture();
		await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
		await f.agent.chats.createChat(chat, owner, { workingDirectories: [cwd] });
		const clock = sinon.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout'] });
		try {
			const firstPrompt = new DeferredPromise<Record<string, unknown>>();
			f.onPrompt = (_transport, request) => { void firstPrompt.complete(request); };
			const first = f.agent.chats.sendMessage(chat, 'first', [cwd], undefined, 'first-turn');
			const firstRequest = await firstPrompt.p;
			await f.agent.chats.abort(chat, owner);
			await f.agent.chats.abort(chat, owner);
			f.transports[0].reply(firstRequest, { stopReason: 'cancelled' });
			await first;
			const secondPrompt = new DeferredPromise<Record<string, unknown>>();
			f.onPrompt = (_transport, request) => { void secondPrompt.complete(request); };
			const second = f.agent.chats.sendMessage(chat, 'second', [cwd], undefined, 'second-turn').then(() => undefined, error => error);
			const secondRequest = await secondPrompt.p;
			await clock.tickAsync(2100);
			f.transports[0].reply(secondRequest, { stopReason: 'end_turn' });
			assert.strictEqual(await second, undefined);
			assert.strictEqual(f.transports[0].sent.filter(message => message.method === 'session/cancel').length, 1);
			assert.strictEqual(f.launches.length, 1);
		} finally { clock.restore(); }
	});

	test('Stop before process startup prevents the prompt and permits a later turn', async () => {
		const f = fixture();
		await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
		await f.agent.chats.createChat(chat, owner, { workingDirectories: [cwd] });
		const send = f.agent.chats.sendMessage(chat, 'stopped', [cwd], undefined, 'stopped-turn');
		const rejected = assert.rejects(send, error => isCancellationError(error));
		await f.agent.chats.abort(chat, owner);
		await rejected;
		assert.strictEqual(f.launches.length, 0);
		await f.agent.chats.sendMessage(chat, 'retry', [cwd], undefined, 'retry-turn');
		assert.strictEqual(f.launches.length, 1);
	});

	test('a follow-up submitted during force-stop waits and resumes the same conversation', async () => {
		const f = fixture();
		await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
		await f.agent.chats.createChat(chat, owner, { workingDirectories: [cwd], model: { id: 'swe-2-high' } });
		const clock = sinon.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout'] });
		try {
			const started = new DeferredPromise<void>();
			f.onPrompt = transport => {
				toolUpdate(transport, { sessionUpdate: 'tool_call', toolCallId: 'read', title: 'Read file', status: 'in_progress', rawOutput: 'partial content' });
				void started.complete();
			};
			const stopped = f.agent.chats.sendMessage(chat, 'stopped', [cwd], undefined, 'stopped-turn').then(() => undefined, error => error);
			await started.p;
			const history = await f.agent.chats.getMessages(chat, owner);
			const stoppedTurn = history[0];
			await f.agent.chats.abort(chat, owner);
			f.onPrompt = undefined;
			const followUp = f.agent.chats.sendMessage(chat, 'retry', [cwd], undefined, 'retry-turn').then(() => undefined, error => error);
			await clock.tickAsync(2001);
			assert.ok(isCancellationError(await stopped));
			assert.strictEqual(stoppedTurn.state, TurnState.Cancelled);
			const tool = stoppedTurn.responseParts[0];
			assert.ok(tool.kind === ResponsePartKind.ToolCall && tool.toolCall.status === ToolCallStatus.Completed);
			assert.strictEqual(tool.toolCall.success, false);
			assert.deepStrictEqual(tool.toolCall.content, [{ type: ToolResultContentType.Text, text: 'partial content' }]);
			assert.strictEqual(await followUp, undefined);
			assert.strictEqual(f.launches.length, 2);
			assert.deepStrictEqual(f.launches[1].args, ['acp', '--model', 'swe-2-high']);
			const requests = f.transports[1].sent;
			assert.ok(!requests.some(message => message.method === 'session/new'));
			assert.deepStrictEqual(requests.find(message => message.method === 'session/load')?.params, { sessionId: 'native', cwd: cwd.fsPath, mcpServers: [] });
			assert.strictEqual(requests.filter(message => message.method === 'session/prompt').length, 1);
		} finally { clock.restore(); }
	});

	test('a follow-up waits for graceful cancellation and is sent exactly once', async () => {
		const f = fixture();
		await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
		await f.agent.chats.createChat(chat, owner, { workingDirectories: [cwd] });
		const started = new DeferredPromise<{ transport: TestAcpTransport; request: Record<string, unknown> }>();
		f.onPrompt = (transport, request) => { void started.complete({ transport, request }); };
		const first = f.agent.chats.sendMessage(chat, 'first', [cwd], undefined, 'first-turn');
		const { transport, request } = await started.p;
		await f.agent.chats.abort(chat, owner);
		f.onPrompt = undefined;
		const second = f.agent.chats.sendMessage(chat, 'second', [cwd], undefined, 'second-turn').then(() => undefined, error => error);
		await assert.rejects(f.agent.chats.sendMessage(chat, 'competing', [cwd], undefined, 'competing-turn'), /turn is already running/);
		assert.strictEqual(transport.sent.filter(message => message.method === 'session/prompt').length, 1);
		transport.reply(request, { stopReason: 'cancelled' });
		await first;
		assert.strictEqual(await second, undefined);
		const prompts = transport.sent.filter(message => message.method === 'session/prompt');
		assert.strictEqual(prompts.length, 2);
		assert.deepStrictEqual(prompts[1].params, { sessionId: 'native', prompt: [{ type: 'text', text: 'second' }] });
		assert.strictEqual(f.launches.length, 1);
	});

	test('Stop can cancel a follow-up waiting behind the previous cancellation', async () => {
		const f = fixture();
		await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
		await f.agent.chats.createChat(chat, owner, { workingDirectories: [cwd] });
		const clock = sinon.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout'] });
		try {
			const started = new DeferredPromise<void>();
			f.onPrompt = () => { void started.complete(); };
			const first = f.agent.chats.sendMessage(chat, 'first', [cwd], undefined, 'first-turn').then(() => undefined, error => error);
			await started.p;
			await f.agent.chats.abort(chat, owner);
			f.onPrompt = undefined;
			const second = f.agent.chats.sendMessage(chat, 'cancel this too', [cwd], undefined, 'second-turn').then(() => undefined, error => error);
			await f.agent.chats.abort(chat, owner);
			await clock.tickAsync(2001);
			assert.ok(isCancellationError(await first));
			assert.ok(isCancellationError(await second));
			assert.strictEqual(f.launches.length, 1);
			await f.agent.chats.sendMessage(chat, 'third', [cwd], undefined, 'third-turn');
			assert.strictEqual(f.launches.length, 2);
			const prompts = f.transports.flatMap(transport => transport.sent.filter(message => message.method === 'session/prompt'));
			assert.strictEqual(prompts.length, 2);
			assert.deepStrictEqual(prompts[1].params, { sessionId: 'native', prompt: [{ type: 'text', text: 'third' }] });
		} finally { clock.restore(); }
	});

	test('a third explicit send stays behind both cancelled operations until cleanup finishes', async () => {
		const f = fixture();
		await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
		await f.agent.chats.createChat(chat, owner, { workingDirectories: [cwd] });
		const clock = sinon.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout'] });
		try {
			const started = new DeferredPromise<void>();
			f.onPrompt = () => { void started.complete(); };
			const first = f.agent.chats.sendMessage(chat, 'first', [cwd], undefined, 'first-turn').then(() => undefined, error => error);
			await started.p;
			await f.agent.chats.abort(chat, owner);
			const second = f.agent.chats.sendMessage(chat, 'cancel this too', [cwd], undefined, 'second-turn').then(() => undefined, error => error);
			await f.agent.chats.abort(chat, owner);
			f.onPrompt = undefined;
			const third = f.agent.chats.sendMessage(chat, 'third', [cwd], undefined, 'third-turn').then(() => undefined, error => error);
			await clock.tickAsync(1999);
			assert.strictEqual(f.launches.length, 1);
			assert.strictEqual(f.transports[0].sent.filter(message => message.method === 'session/prompt').length, 1);
			await clock.tickAsync(2);
			assert.ok(isCancellationError(await first));
			assert.ok(isCancellationError(await second));
			assert.strictEqual(await third, undefined);
			assert.strictEqual(f.launches.length, 2);
			assert.deepStrictEqual(f.transports[1].sent.find(message => message.method === 'session/prompt')?.params, { sessionId: 'native', prompt: [{ type: 'text', text: 'third' }] });
		} finally { clock.restore(); }
	});

	test('a history read during failed-turn cleanup does not start a competing replay', async () => {
		const f = fixture();
		await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
		await f.agent.chats.createChat(chat, owner, { workingDirectories: [cwd] });
		const started = new DeferredPromise<void>();
		f.onPrompt = () => { void started.complete(); };
		const send = f.agent.chats.sendMessage(chat, 'first', [cwd], undefined, 'first-turn').then(() => undefined, error => error);
		await started.p;
		f.transports[0].close();
		const history = f.agent.chats.getMessages(chat, owner);
		assert.strictEqual(f.launches.length, 1);
		assert.strictEqual((await history)[0].id, 'first-turn');
		assert.ok((await send) instanceof Error);
	});

	for (const invalidation of ['release', 'dispose', 'signOut', 'accountChanged', 'shutdown'] as const) {
		test(`${invalidation} prevents a waiting follow-up from starting`, async () => {
			const f = fixture();
			await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
			await f.agent.chats.createChat(chat, owner, { workingDirectories: [cwd] });
			const started = new DeferredPromise<void>();
			f.onPrompt = () => { void started.complete(); };
			const first = f.agent.chats.sendMessage(chat, 'first', [cwd], undefined, 'first-turn').then(() => undefined, error => error);
			await started.p;
			await f.agent.chats.abort(chat, owner);
			f.onPrompt = undefined;
			const second = f.agent.chats.sendMessage(chat, 'stale follow-up', [cwd], undefined, 'second-turn').then(() => undefined, error => error);
			if (invalidation === 'release') { await f.agent.chats.releaseChat(chat, owner); }
			else if (invalidation === 'dispose') { await f.agent.chats.disposeChat(chat, owner); }
			else if (invalidation === 'signOut') { await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, ''); }
			else if (invalidation === 'accountChanged') { await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'replacement-account-token'); }
			else { await f.agent.shutdown(); }
			assert.ok(isCancellationError(await first));
			assert.ok(isCancellationError(await second));
			assert.strictEqual(f.launches.length, 1);
			assert.strictEqual(f.transports[0].sent.filter(message => message.method === 'session/prompt').length, 1);
		});
	}

	test('a competing send or model change cannot change the workspace while a turn prepares', async () => {
		const f = fixture();
		await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
		await f.agent.chats.createChat(chat, owner, { workingDirectories: [cwd], model: { id: 'swe-2-high' } });
		const started = new DeferredPromise<{ transport: TestAcpTransport; request: Record<string, unknown> }>();
		f.onInitialize = (transport, request) => { void started.complete({ transport, request }); };
		const send = f.agent.chats.sendMessage(chat, 'first', [cwd], undefined, 'first-turn');
		const { transport, request } = await started.p;
		await assert.rejects(f.agent.chats.sendMessage(chat, 'competing', [URI.file('/another project')], undefined, 'other-turn'), /turn is already running/);
		await assert.rejects(f.agent.chats.changeModel(chat, { id: 'another-model' }, owner), /Wait for the current Devin turn/);
		transport.reply(request, { protocolVersion: 1, agentCapabilities: { loadSession: true } });
		await send;
		assert.strictEqual(f.launches.length, 1);
		assert.deepStrictEqual(transport.sent.find(message => message.method === 'session/new')?.params, { cwd: cwd.fsPath, mcpServers: [] });
		assert.strictEqual(transport.sent.filter(message => message.method === 'session/prompt').length, 1);
	});

	for (const phase of ['initialize', 'load'] as const) {
		test(`Stop cancels pending ${phase} and ignores its late response without sending a prompt`, async () => {
			const f = fixture();
			await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
			if (phase === 'load') {
				await f.agent.materializeChat(chat, owner, JSON.stringify({ version: 1, sessionId: 'native', cwd: cwd.toString(), model: 'swe-2-high' }));
			} else {
				await f.agent.chats.createChat(chat, owner, { workingDirectories: [cwd] });
			}
			const clock = sinon.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout'] });
			try {
				const started = new DeferredPromise<{ transport: TestAcpTransport; request: Record<string, unknown> }>();
				const onSetup = (transport: TestAcpTransport, request: Record<string, unknown>) => { void started.complete({ transport, request }); };
				if (phase === 'initialize') { f.onInitialize = onSetup; } else { f.onLoad = onSetup; }
				const result = f.agent.chats.sendMessage(chat, 'stopped', [cwd], undefined, 'stopped-turn').then(() => undefined, error => error);
				const { transport, request } = await started.p;
				await f.agent.chats.abort(chat, owner);
				transport.reply(request, phase === 'initialize' ? { protocolVersion: 1, agentCapabilities: { loadSession: true } } : { configOptions: [] });
				await clock.tickAsync(1);
				assert.ok(isCancellationError(await result));
				assert.ok(!transport.sent.some(message => message.method === 'session/prompt'));
				assert.ok(!transport.sent.some(message => message.method === 'session/cancel'));
				f.onInitialize = undefined;
				f.onLoad = undefined;
				await f.agent.chats.sendMessage(chat, 'retry', [cwd], undefined, 'retry-turn');
				assert.strictEqual(f.launches.length, 2);
				assert.strictEqual(f.transports[1].sent.filter(message => message.method === (phase === 'load' ? 'session/load' : 'session/new')).length, 1);
			} finally { clock.restore(); }
		});
	}

	test('declines new permission requests after Stop without asking the user to approve', async () => {
		const f = fixture();
		await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
		await f.agent.chats.createChat(chat, owner, { workingDirectories: [cwd] });
		const prompt = new DeferredPromise<{ transport: TestAcpTransport; request: Record<string, unknown> }>();
		f.onPrompt = (transport, request) => { void prompt.complete({ transport, request }); };
		const send = f.agent.chats.sendMessage(chat, 'prompt', [cwd], undefined, 'stopped-turn');
		const { transport, request } = await prompt.p;
		await f.agent.chats.abort(chat, owner);
		const response = Event.toPromise(Event.filter(transport.onSend, message => message.id === 'late-permission'));
		transport.receive({ jsonrpc: '2.0', id: 'late-permission', method: 'session/request_permission', params: { sessionId: 'native', toolCall: { toolCallId: 'late-tool', title: 'Late tool' }, options: [{ kind: 'allow_once', optionId: 'once' }] } });
		assert.deepStrictEqual((await response).result, { outcome: { outcome: 'cancelled' } });
		assert.ok(!f.events.some(event => event.kind === 'pending_confirmation'));
		// A normal completion can race with Stop; the user's cancellation still wins.
		transport.reply(request, { stopReason: 'end_turn' });
		await send;
		assert.strictEqual((await f.agent.chats.getMessages(chat, owner))[0].state, TurnState.Cancelled);
	});

	test('settles a superseded permission RPC and requires a decision for the replacement request', async () => {
		const f = fixture();
		await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
		await f.agent.chats.createChat(chat, owner, { workingDirectories: [cwd] });
		let promptRequest: Record<string, unknown>;
		f.onPrompt = (transport, request) => {
			promptRequest = request;
			transport.receive({ jsonrpc: '2.0', id: 'first-permission', method: 'session/request_permission', params: { sessionId: 'native', toolCall: { toolCallId: 'read', title: 'Read file' }, options: [{ optionId: 'first-once', kind: 'allow_once' }] } });
		};
		const pending = Event.toPromise(Event.filter(f.agent.onDidChatProgress, event => event.kind === 'pending_confirmation'));
		const send = f.agent.chats.sendMessage(chat, 'read', [cwd], undefined, 'host-turn');
		await pending;
		const transport = f.transports[0];
		const first = Event.toPromise(Event.filter(transport.onSend, message => message.id === 'first-permission'));
		transport.receive({ jsonrpc: '2.0', id: 'second-permission', method: 'session/request_permission', params: { sessionId: 'native', toolCall: { toolCallId: 'read', title: 'Read another file', rawInput: 'another.md' }, options: [{ optionId: 'second-once', kind: 'allow_once' }] } });
		assert.deepStrictEqual((await first).result, { outcome: { outcome: 'cancelled' } });
		const second = Event.toPromise(Event.filter(transport.onSend, message => message.id === 'second-permission'));
		f.agent.respondToPermissionRequest('host-turn:read', true);
		assert.deepStrictEqual((await second).result, { outcome: { outcome: 'selected', optionId: 'second-once' } });
		transport.reply(promptRequest!, { stopReason: 'end_turn' });
		await send;
	});

	test('does not turn an ordinary approval into an always-allow choice when no one-time option exists', async () => {
		const f = fixture();
		await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
		await f.agent.chats.createChat(chat, owner, { workingDirectories: [cwd] });
		let promptRequest: Record<string, unknown>;
		f.onPrompt = (transport, request) => {
			promptRequest = request;
			transport.receive({ jsonrpc: '2.0', id: 'permission', method: 'session/request_permission', params: { sessionId: 'native', toolCall: { toolCallId: 'read', title: 'Read file' }, options: [{ optionId: 'forever', kind: 'allow_always' }] } });
		};
		const pending = Event.toPromise(Event.filter(f.agent.onDidChatProgress, event => event.kind === 'pending_confirmation'));
		const send = f.agent.chats.sendMessage(chat, 'read', [cwd], undefined, 'host-turn');
		await pending;
		const transport = f.transports[0];
		const response = Event.toPromise(Event.filter(transport.onSend, message => message.id === 'permission' && Object.hasOwn(message, 'result')));
		f.agent.respondToPermissionRequest('host-turn:read', true);
		assert.deepStrictEqual((await response).result, { outcome: { outcome: 'cancelled' } });
		transport.reply(promptRequest!, { stopReason: 'end_turn' });
		await send;
	});

	test('sends a plain greeting with automatic context and preserves the attachments in history', async () => {
		const f = fixture();
		await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
		await f.agent.chats.createChat(chat, owner, { workingDirectories: [cwd] });
		const attachments: MessageAttachment[] = [
			{ type: MessageAttachmentKind.Simple, label: 'Workspace context', modelRepresentation: 'Current workspace: example' },
			{ type: MessageAttachmentKind.Simple, label: 'Display-only completion' },
			{ type: MessageAttachmentKind.EmbeddedResource, label: 'Unsaved editor', contentType: 'text/plain', data: Buffer.from('live editor text 🌎').toString('base64') },
		];
		await f.agent.chats.sendMessage(chat, 'hey', [cwd], attachments, 'greeting');
		const request = f.transports[0].sent.find(message => message.method === 'session/prompt');
		const params = request?.params as { prompt: { text: string }[] };
		assert.deepStrictEqual(params.prompt.map(block => block.text), ['hey', 'Current workspace: example', 'Unsaved editor (read-only attached content):\nlive editor text 🌎']);
		assert.deepStrictEqual((await f.agent.chats.getMessages(chat, owner))[0].message.attachments, attachments);
	});

	test('preserves file selection references and only sends binary content with an advertised capability', () => {
		const resource: MessageAttachment = { type: MessageAttachmentKind.Resource, label: 'source.ts', uri: URI.file('/project/source.ts').toString(), selection: { range: { start: { line: 2, character: 0 }, end: { line: 5, character: 0 } } } };
		const blocks = resolveAcpPrompt('inspect', [resource], {});
		assert.strictEqual(blocks[1].type, 'resource_link');
		assert.strictEqual(blocks[1].uri, resource.uri);
		assert.match(String(blocks[2].text), /lines 3-5/);
		const image: MessageAttachment = { type: MessageAttachmentKind.EmbeddedResource, label: 'screen.png', contentType: 'image/png', data: 'AA==' };
		assert.throws(() => resolveAcpPrompt('inspect', [image], {}), /screen\.png/);
		assert.deepStrictEqual(resolveAcpPrompt('inspect', [image], { image: true })[1], { type: 'image', mimeType: 'image/png', data: 'AA==' });
		const document: MessageAttachment = { ...image, label: 'document.pdf', contentType: 'application/pdf' };
		const embedded = resolveAcpPrompt('inspect', [document], { embeddedContext: true })[1];
		assert.strictEqual(embedded.type, 'resource');
		assert.deepStrictEqual(embedded.resource, { uri: 'attachment:/document.pdf', mimeType: 'application/pdf', blob: 'AA==' });
	});

	test('restores the provider session with replay instead of creating another conversation', async () => {
		const f = fixture();
		await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
		await f.agent.materializeChat(chat, owner, JSON.stringify({ version: 1, sessionId: 'native', cwd: cwd.toString() }));
		const turns = await f.agent.chats.getMessages(chat, owner);
		assert.strictEqual(turns[0].message.text, 'original');
		assert.ok(f.transports[0].sent.some(message => message.method === 'session/load'));
		assert.ok(!f.transports[0].sent.some(message => message.method === 'session/new'));
		assert.deepStrictEqual(f.events, []);
	});

	test('exposes persisted metadata before materialization and restores model, folder, and timestamps without starting a client', async () => {
		const f = fixture();
		const data = JSON.stringify({ version: 1, sessionId: 'native', cwd: cwd.toString(), model: 'swe-2-high', createdAt: 1000, modifiedAt: 2000 });
		const metadata = await f.agent.getChatMetadata(chat, owner, data);
		assert.strictEqual(metadata?.startTime, 1000);
		assert.strictEqual(metadata?.modifiedTime, 2000);
		assert.deepStrictEqual(metadata?.model, { id: 'swe-2-high' });
		assert.deepStrictEqual(metadata?.workingDirectories?.map(uri => uri.toString()), [cwd.toString()]);
		assert.strictEqual(f.launches.length, 0);
		await f.agent.materializeChat(chat, owner, data);
		const materialized = await f.agent.getChatMetadata(chat, owner);
		assert.strictEqual(materialized?.startTime, metadata?.startTime);
		assert.strictEqual(materialized?.modifiedTime, metadata?.modifiedTime);
		assert.deepStrictEqual(materialized?.model, metadata?.model);
		assert.strictEqual(materialized?.project?.uri.toString(), cwd.toString());
		assert.strictEqual(f.launches.length, 0);
		await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
		await f.agent.chats.getMessages(chat, owner);
		assert.deepStrictEqual(f.launches[0].args, ['acp', '--model', 'swe-2-high']);
		assert.ok(!f.transports[0].sent.some(message => message.method === 'session/set_config_option'));
		assert.strictEqual((await f.agent.getChatMetadata(chat, owner))?.startTime, 1000);
	});

	test('rejects corrupt or nonlocal backings without starting a client and accepts the older timestamp format', async () => {
		const f = fixture();
		for (const data of ['{', JSON.stringify({ version: 1, sessionId: '', cwd: cwd.toString() }), JSON.stringify({ version: 1, sessionId: 'native', cwd: 'https://example.com' })]) {
			assert.strictEqual(await f.agent.getChatMetadata(chat, owner, data), undefined);
			await assert.rejects(f.agent.materializeChat(chat, owner, data), /Invalid persisted Devin chat/);
		}
		const old = await f.agent.getChatMetadata(chat, owner, JSON.stringify({ version: 1, sessionId: 'native', cwd: cwd.toString(), createdAt: 1000 }));
		assert.strictEqual(old?.modifiedTime, 1000);
		assert.deepStrictEqual(f.launches, []);
	});

	test('preserves a restored variant represented by separate native model and thinking selectors', async () => {
		const f = fixture();
		await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
		await f.agent.materializeChat(chat, owner, JSON.stringify({ version: 1, sessionId: 'native', cwd: cwd.toString(), model: 'swe-2-max', createdAt: 1000 }));
		f.onLoad = (transport, request) => transport.reply(request, { configOptions: [
			{ id: 'model', category: 'model', type: 'select', currentValue: 'swe-2-high', options: [{ value: 'swe-2-high', name: 'SWE-2' }] },
			{ id: 'thought_level', category: 'thought_level', type: 'select', currentValue: 'max', options: [{ value: 'high' }, { value: 'max' }] },
		] });
		await f.agent.chats.getMessages(chat, owner);
		assert.deepStrictEqual((await f.agent.getChatMetadata(chat, owner))?.model, { id: 'swe-2-max' });
		// The host reconciles its model selection before every outgoing turn.
		await f.agent.chats.changeModel(chat, { id: 'swe-2-max', config: {} }, owner);
		assert.ok(!f.transports[0].sent.some(message => message.method === 'session/set_config_option'));
	});

	test('keeps the previous transcript when a replay fails partway, then allows a clean retry', async () => {
		const f = fixture();
		await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
		await f.agent.materializeChat(chat, owner, JSON.stringify({ version: 1, sessionId: 'native', cwd: cwd.toString() }));
		const turns = await f.agent.chats.getMessages(chat, owner);
		const previous = structuredClone(turns);
		await f.agent.chats.releaseChat(chat, owner);
		f.onLoad = (transport, request) => {
			transport.receive({ jsonrpc: '2.0', method: 'session/update', params: { sessionId: 'native', update: { sessionUpdate: 'user_message_chunk', content: { type: 'text', text: 'incomplete replay' } } } });
			transport.receive({ jsonrpc: '2.0', id: request.id, error: { code: -32603, message: 'Replay interrupted' } });
		};
		await assert.rejects(f.agent.chats.getMessages(chat, owner), /Replay interrupted/);
		assert.deepStrictEqual(turns, previous);
		f.onLoad = undefined;
		const retried = await f.agent.chats.getMessages(chat, owner);
		assert.strictEqual(retried.length, 1);
		assert.strictEqual(retried[0].message.text, 'original');
		assert.strictEqual(f.transports.length, 3);
		assert.ok(f.transports.every(transport => transport.sent.every(message => message.method !== 'session/new')));
	});

	test('reopens a released native session before applying an explicit model change', async () => {
		const f = fixture();
		await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
		await f.agent.materializeChat(chat, owner, JSON.stringify({ version: 1, sessionId: 'native', cwd: cwd.toString(), model: 'swe-2-max' }));
		await f.agent.chats.changeModel(chat, { id: 'swe-2-high' }, owner);
		const methods = f.transports[0].sent.map(request => request.method);
		assert.ok(methods.indexOf('session/load') < methods.indexOf('session/set_config_option'));
		assert.ok(!methods.includes('session/new'));
		assert.deepStrictEqual((await f.agent.getChatMetadata(chat, owner))?.model, { id: 'swe-2-high' });
	});

	test('persists the latest activity timestamp when a turn completes', async () => {
		const f = fixture();
		await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
		await f.agent.chats.createChat(chat, owner, { workingDirectories: [cwd] });
		let saved: string | undefined;
		store.add(f.agent.onDidChangeChatData(event => { saved = event.providerData; }));
		await f.agent.chats.sendMessage(chat, 'hey', [cwd], undefined, 'host-turn');
		assert.ok(saved);
		const savedMetadata = await fixture().agent.getChatMetadata(chat, owner, saved);
		const live = await f.agent.getChatMetadata(chat, owner);
		assert.strictEqual(savedMetadata?.startTime, live?.startTime);
		assert.strictEqual(savedMetadata?.modifiedTime, live?.modifiedTime);
		assert.strictEqual(savedMetadata?.project?.uri.toString(), cwd.toString());
	});

	test('persists an unconfirmed model change and reconciles the previous model after a new host loads it', async () => {
		const modelCatalog = { families: [{ variants: [...catalog.families[0].variants, { model_uid: 'swe-2-max', label: 'SWE-2 Max', max_context_tokens: 262000 }] }] };
		const f = fixture(modelCatalog);
		await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
		await f.agent.materializeChat(chat, owner, JSON.stringify({ version: 1, sessionId: 'native', cwd: cwd.toString(), model: 'swe-2-high' }));
		let saved: string | undefined;
		store.add(f.agent.onDidChangeChatData(event => { saved = event.providerData; }));
		f.onConfig = transport => transport.close();
		await assert.rejects(f.agent.chats.changeModel(chat, { id: 'swe-2-max' }, owner), /peer exited/);
		assert.ok(saved);
		assert.strictEqual(JSON.parse(saved).modelNeedsReconciliation, true);
		assert.strictEqual(JSON.parse(saved).model, 'swe-2-high');
		const resumed = fixture(modelCatalog);
		await resumed.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
		await resumed.agent.materializeChat(chat, owner, saved);
		let recovered: string | undefined;
		store.add(resumed.agent.onDidMaterializeChat(event => { recovered = event.result?.providerData; }));
		await resumed.agent.chats.getMessages(chat, owner);
		assert.ok(resumed.transports[0].sent.some(request => request.method === 'session/set_config_option'));
		assert.ok(recovered);
		assert.strictEqual(JSON.parse(recovered).modelNeedsReconciliation, undefined);
		assert.strictEqual(JSON.parse(recovered).model, 'swe-2-high');
	});

	test('allows selecting an available model when the previous unconfirmed model is no longer in the catalog', async () => {
		const f = fixture();
		await f.agent.authenticate(DEVIN_SUBSCRIPTION_RESOURCE, 'browser-token');
		await f.agent.materializeChat(chat, owner, JSON.stringify({ version: 1, sessionId: 'native', cwd: cwd.toString(), model: 'retired-model', modelNeedsReconciliation: true }));
		await f.agent.chats.changeModel(chat, { id: 'swe-2-high' }, owner);
		assert.deepStrictEqual((await f.agent.getChatMetadata(chat, owner))?.model, { id: 'swe-2-high' });
		assert.ok(f.transports[0].sent.some(request => request.method === 'session/set_config_option'));
	});
});
