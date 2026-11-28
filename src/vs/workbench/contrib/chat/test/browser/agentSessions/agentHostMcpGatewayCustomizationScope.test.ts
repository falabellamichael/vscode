/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import assert from 'assert';
import { DeferredPromise } from '../../../../../../base/common/async.js';
import { Emitter, Event } from '../../../../../../base/common/event.js';
import { constObservable, observableValue, waitForState } from '../../../../../../base/common/observable.js';
import { URI } from '../../../../../../base/common/uri.js';
import { upcastPartial } from '../../../../../../base/test/common/mock.js';
import { ensureNoDisposablesAreLeakedInTestSuite } from '../../../../../../base/test/common/utils.js';
import { createAgentHostResourceUriMapper } from '../../../../../../platform/agentHost/common/agentHostUri.js';
import { CustomizationType, type ClientPluginCustomization } from '../../../../../../platform/agentHost/common/state/sessionState.js';
import { FileService } from '../../../../../../platform/files/common/fileService.js';
import { IFileService } from '../../../../../../platform/files/common/files.js';
import { InMemoryFileSystemProvider } from '../../../../../../platform/files/common/inMemoryFilesystemProvider.js';
import { TestInstantiationService } from '../../../../../../platform/instantiation/test/common/instantiationServiceMock.js';
import { ILogService, NullLogService } from '../../../../../../platform/log/common/log.js';
import { IAgentCustomizationScope } from '../../../browser/agentSessions/agentHost/agentHostActiveClientService.js';
import { AgentHostMcpGatewayCustomizationScope } from '../../../browser/agentSessions/agentHost/agentHostMcpGatewayCustomizationScope.js';
import { ISignalCDirectMail, ISignalCDirectMailConnection } from '../../../../mcp/browser/signalCDirectMail.js';
import { directMailGatewayId } from '../../../../mcp/browser/signalCDirectMailGateway.js';
import { IMcpGatewayResult, IMcpGatewayResultServer, IWorkbenchMcpGatewayService } from '../../../../mcp/common/mcpGatewayService.js';
import { IMcpServer, IMcpService, McpCollectionDefinition, McpCollectionProvenance } from '../../../../mcp/common/mcpTypes.js';
import { IAgentHostFileSystemService, SYNCED_CUSTOMIZATION_SCHEME } from '../../../../../services/agentHost/common/agentHostFileSystemService.js';

suite('Agent host MCP gateway customization scope', () => {
	const store = ensureNoDisposablesAreLeakedInTestSuite();
	const chat = URI.parse('agent-host-devin-subscription:/opaque-chat');
	const root = URI.file('/workspace');
	const endpoint = (label = 'editor-tools'): IMcpGatewayResultServer => ({ label, address: URI.parse('http://127.0.0.1:43123/gateway/test-only-route') });
	const server = (id: string): IMcpServer => upcastPartial<IMcpServer>({
		definition: { id, label: id },
		readDefinitions: () => constObservable({ server: undefined, collection: upcastPartial<McpCollectionDefinition>({ provenance: McpCollectionProvenance.UserProfile }) }),
	});

	function fixture(initialServers: readonly IMcpServer[] = [server('selected')], initialDelay?: Promise<void>, initialMail: readonly ISignalCDirectMailConnection[] = []) {
		const fileService = store.add(new FileService(new NullLogService()));
		store.add(fileService.registerProvider(SYNCED_CUSTOMIZATION_SCHEME, store.add(new InMemoryFileSystemProvider())));
		const instantiation = store.add(new TestInstantiationService());
		instantiation.stub(IFileService, fileService);
		instantiation.stub(ILogService, new NullLogService());
		instantiation.stub(IAgentHostFileSystemService, { ensureSyncedCustomizationProvider() { } });
		const servers = observableValue<readonly IMcpServer[]>('testServers', initialServers);
		const customizations = observableValue<readonly ClientPluginCustomization[]>('baseRefs', [{ type: CustomizationType.Plugin, id: 'existing', uri: 'file:///existing', name: 'Existing' }]);
		let baseDisposed = false;
		const base: IAgentCustomizationScope = {
			customizations, customAgents: constObservable([]), tools: constObservable([]), isResolved: constObservable(true),
			activeClient: clientId => constObservable({ clientId, tools: [], customizations: [...customizations.get()] }),
			whenResolved: async () => { }, getSyncedUri: uri => uri, dispose: () => { baseDisposed = true; },
		};
		const changed = store.add(new Emitter<readonly IMcpGatewayResultServer[]>());
		const disposed = store.add(new Emitter<void>());
		let gatewayDisposed = 0;
		let gatewayServers: readonly IMcpGatewayResultServer[] = [endpoint()];
		let delay: Promise<void> | undefined = initialDelay;
		const calls: { resource: string | undefined; ids: readonly string[] | undefined }[] = [];
		const called = store.add(new Emitter<void>());
		const service = upcastPartial<IWorkbenchMcpGatewayService>({ createGateway: async (remote, resource, ids) => {
			assert.strictEqual(remote, false);
			calls.push({ resource: resource?.toString(), ids });
			called.fire();
			await delay;
			return { servers: gatewayServers, onDidChangeServers: changed.event, dispose: () => { gatewayDisposed++; disposed.fire(); } } satisfies IMcpGatewayResult;
		} });
		const mail = observableValue('testMail', initialMail);
		const scope = store.add(new AgentHostMcpGatewayCustomizationScope(base, chat, [root], upcastPartial<IMcpService>({ servers }), service, instantiation, upcastPartial<ISignalCDirectMail>({ connections: mail })));
		return { scope, fileService, customizations, servers, mail, changed, calls,
			whenCalled: async (count: number) => { if (calls.length < count) { await Event.toPromise(Event.filter(called.event, () => calls.length >= count)); } },
			whenDisposed: async (count: number) => { if (gatewayDisposed < count) { await Event.toPromise(Event.filter(disposed.event, () => gatewayDisposed >= count)); } },
			get gatewayDisposed() { return gatewayDisposed; }, get baseDisposed() { return baseDisposed; }, set gatewayServers(value: readonly IMcpGatewayResultServer[]) { gatewayServers = value; }, set delay(value: Promise<void>) { delay = value; } };
	}

	test('publishes a canonical per-chat bundle alongside existing refs and maps it to the connection URI space', async () => {
		const f = fixture();
		await f.scope.whenResolved();
		assert.deepStrictEqual(f.calls, [{ resource: chat.toString(), ids: ['selected'] }]);
		const refs = f.scope.customizations.get();
		assert.strictEqual(refs[0].id, 'existing');
		const bundle = refs[1];
		const content = await f.fileService.readFile(URI.joinPath(URI.parse(bundle.uri), '.mcp.json'));
		assert.deepStrictEqual(JSON.parse(content.value.toString()), { mcpServers: { 'editor-tools': { type: 'http', url: endpoint().address.toString() } } });
		assert.ok(bundle.childEnablement?.['editor-tools']);
		const client = f.scope.activeClient('exact-client', createAgentHostResourceUriMapper('local')).get();
		assert.strictEqual(client.clientId, 'exact-client');
		assert.strictEqual(client.customizations?.length, 2);
		assert.strictEqual(client.customizations?.[1].uri, bundle.uri);
	});

	test('shares only connected mail identities and revokes the gateway when the grant expires', async () => {
		const gmail = { id: 'gmail-account', provider: 'gmail', label: 'Gmail', needsSignIn: false } as const;
		const outlook = { id: 'outlook-account', provider: 'outlook', label: 'Outlook', needsSignIn: true } as const;
		const f = fixture([], undefined, [gmail, outlook]);
		await f.scope.whenResolved();
		assert.deepStrictEqual(f.calls.map(call => call.ids), [[directMailGatewayId(gmail)]]);
		f.mail.set([{ ...gmail, needsSignIn: true }, outlook], undefined);
		await f.scope.whenResolved();
		assert.strictEqual(f.gatewayDisposed, 1);
		assert.strictEqual(f.scope.customizations.get().length, 1);
	});

	test('removes the published bundle when the gateway revokes all server routes', async () => {
		const f = fixture();
		await f.scope.whenResolved();
		f.changed.fire([]);
		await waitForState(f.scope.customizations, refs => refs.length === 1);
		assert.strictEqual(f.scope.customizations.get()[0].id, 'existing');
	});

	test('retains gateway routes and the published bundle when discovery republishes the same selected servers', async () => {
		const f = fixture([server('second'), server('first')]);
		await f.scope.whenResolved();
		const refs = f.scope.customizations.get();
		f.servers.set([server('first'), server('second')], undefined);
		assert.strictEqual(f.gatewayDisposed, 0, 'An unchanged selection must not revoke an active Devin tool connection');
		assert.strictEqual(f.scope.isResolved.get(), true);
		assert.deepStrictEqual(f.scope.customizations.get(), refs);
		assert.strictEqual(f.calls.length, 1);
	});

	test('does not reopen the gateway for a definition update that leaves eligibility unchanged', async () => {
		const definitions = observableValue('definitions', { server: undefined, collection: upcastPartial<McpCollectionDefinition>({ provenance: McpCollectionProvenance.UserProfile }) });
		const selected = upcastPartial<IMcpServer>({ definition: { id: 'selected', label: 'selected' }, readDefinitions: () => definitions });
		const f = fixture([selected]);
		await f.scope.whenResolved();
		const refs = f.scope.customizations.get();
		definitions.set({ server: undefined, collection: upcastPartial<McpCollectionDefinition>({ provenance: McpCollectionProvenance.UserProfile }) }, undefined);
		assert.strictEqual(f.gatewayDisposed, 0);
		assert.strictEqual(f.scope.isResolved.get(), true);
		assert.deepStrictEqual(f.scope.customizations.get(), refs);
		assert.strictEqual(f.calls.length, 1);
	});

	test('still revokes a selected server when its definition is removed', async () => {
		const definitions = observableValue<{ server: undefined; collection: McpCollectionDefinition | undefined }>('definitions', { server: undefined, collection: upcastPartial<McpCollectionDefinition>({ provenance: McpCollectionProvenance.UserProfile }) });
		const selected = upcastPartial<IMcpServer>({ definition: { id: 'selected', label: 'selected' }, readDefinitions: () => definitions });
		const f = fixture([selected]);
		await f.scope.whenResolved();
		definitions.set({ server: undefined, collection: undefined }, undefined);
		assert.strictEqual(f.gatewayDisposed, 1);
		await f.scope.whenResolved();
		assert.strictEqual(f.scope.customizations.get().length, 1);
		assert.strictEqual(f.calls.length, 1);
	});

	test('does not restart a pending gateway for an unchanged discovery update', async () => {
		const blocked = new DeferredPromise<void>();
		const f = fixture([server('selected')], blocked.p);
		f.servers.set([server('selected')], undefined);
		assert.strictEqual(f.calls.length, 1);
		void blocked.complete();
		await f.scope.whenResolved();
		assert.strictEqual(f.gatewayDisposed, 0);
		assert.strictEqual(f.scope.customizations.get().length, 2);
	});

	test('revokes the old gateway immediately and does not publish a late replacement after disposal', async () => {
		const f = fixture();
		await f.scope.whenResolved();
		const blocked = new DeferredPromise<void>();
		f.delay = blocked.p;
		f.servers.set([server('replacement')], undefined);
		assert.strictEqual(f.gatewayDisposed, 1);
		assert.strictEqual(f.scope.isResolved.get(), false);
		await f.whenCalled(2);
		f.scope.dispose();
		void blocked.complete();
		await f.whenDisposed(2);
		assert.strictEqual(f.gatewayDisposed, 2);
		assert.strictEqual(f.baseDisposed, true);
		assert.strictEqual(f.scope.customizations.get().length, 1);
	});

	test('does not create a gateway when the exact chat has no eligible servers', async () => {
		const f = fixture([]);
		await f.scope.whenResolved();
		assert.deepStrictEqual(f.calls, []);
		assert.strictEqual(f.scope.customizations.get().length, 1);
	});

	test('a pending old gateway cannot hold up the latest selection or overwrite its bundle', async () => {
		const blocked = new DeferredPromise<void>();
		const f = fixture([server('old')], blocked.p);
		f.delay = Promise.resolve();
		f.servers.set([server('current')], undefined);
		await f.scope.whenResolved();
		assert.deepStrictEqual(f.calls.map(call => call.ids), [['old'], ['current']]);
		const current = f.scope.customizations.get();
		void blocked.complete();
		await f.whenDisposed(1);
		assert.deepStrictEqual(f.scope.customizations.get(), current);
	});

	test('clearing the selection settles without waiting for an obsolete gateway to open', async () => {
		const blocked = new DeferredPromise<void>();
		const f = fixture([server('old')], blocked.p);
		f.servers.set([], undefined);
		await f.scope.whenResolved();
		assert.strictEqual(f.scope.customizations.get().length, 1);
		void blocked.complete();
		await f.whenDisposed(1);
	});

	test('fails visibly on duplicate server names, disposes access, and retries on the next message', async () => {
		const f = fixture();
		f.gatewayServers = [endpoint(), endpoint()];
		await assert.rejects(f.scope.whenResolved(), /distinct names/);
		assert.strictEqual(f.gatewayDisposed, 1);
		assert.strictEqual(f.scope.customizations.get().length, 1);
		f.gatewayServers = [endpoint('unique')];
		await f.scope.whenResolved();
		assert.strictEqual(f.calls.length, 2);
		assert.ok(f.scope.customizations.get()[1].childEnablement?.unique);
	});
});
