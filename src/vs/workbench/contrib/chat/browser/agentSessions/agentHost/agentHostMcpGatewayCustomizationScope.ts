/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { Sequencer } from '../../../../../../base/common/async.js';
import { equals } from '../../../../../../base/common/arrays.js';
import { CancellationTokenSource } from '../../../../../../base/common/cancellation.js';
import { Disposable, DisposableStore, MutableDisposable, toDisposable } from '../../../../../../base/common/lifecycle.js';
import { autorun, derived, IObservable, observableValue, waitForState } from '../../../../../../base/common/observable.js';
import { URI } from '../../../../../../base/common/uri.js';
import { localize } from '../../../../../../nls.js';
import { IAgentHostResourceUriMapper } from '../../../../../../platform/agentHost/common/agentHostUri.js';
import { CustomizationEnablementKind, type SessionActiveClient } from '../../../../../../platform/agentHost/common/state/protocol/state.js';
import { type ClientPluginCustomization } from '../../../../../../platform/agentHost/common/state/sessionState.js';
import { IInstantiationService } from '../../../../../../platform/instantiation/common/instantiation.js';
import { McpServerType } from '../../../../../../platform/mcp/common/mcpPlatformTypes.js';
import { IMcpGatewayResultServer, IWorkbenchMcpGatewayService } from '../../../../mcp/common/mcpGatewayService.js';
import { IMcpService } from '../../../../mcp/common/mcpTypes.js';
import { ISignalCDirectMail } from '../../../../mcp/browser/signalCDirectMail.js';
import { directMailGatewayId } from '../../../../mcp/browser/signalCDirectMailGateway.js';
import { toAgentHostClientCustomization } from '../../../common/agentHostSessionCustomizations.js';
import { IAgentCustomizationScope } from './agentHostActiveClientService.js';
import { isAgentHostMcpGatewayServer } from './agentHostMcpServerSupport.js';
import { SyncedCustomizationBundler } from './syncedCustomizationBundler.js';

/** Adds a per-chat, explicitly selected editor gateway bundle to a shared customization scope. */
export class AgentHostMcpGatewayCustomizationScope extends Disposable implements IAgentCustomizationScope {

	private readonly _gateway = this._register(new MutableDisposable<DisposableStore>());
	private readonly _ready = observableValue(this, false);
	private readonly _refs = observableValue<readonly ClientPluginCustomization[]>(this, []);
	private readonly _queue = new Sequencer();
	private readonly _cancellation = new CancellationTokenSource();
	private readonly _bundler: SyncedCustomizationBundler;
	private _generation = 0;
	private _selectedIds: readonly string[] = [];
	private _error: Error | undefined;
	readonly customizations;
	readonly customAgents;
	readonly tools;
	readonly isResolved;

	constructor(
		private readonly _scope: IAgentCustomizationScope,
		private readonly _sessionResource: URI,
		roots: readonly URI[] | undefined,
		@IMcpService mcpService: IMcpService,
		@IWorkbenchMcpGatewayService private readonly _gatewayService: IWorkbenchMcpGatewayService,
		@IInstantiationService instantiationService: IInstantiationService,
		@ISignalCDirectMail directMail: ISignalCDirectMail,
	) {
		super();
		this._register(_scope);
		this._register(toDisposable(() => this._cancellation.dispose(true)));
		this._bundler = this._register(instantiationService.createInstance(SyncedCustomizationBundler, `mcp-gateway-${encodeURIComponent(_sessionResource.toString())}`, { standalone: true }));
		this.customizations = derived(this, reader => [..._scope.customizations.read(reader), ...this._refs.read(reader)]);
		this.customAgents = _scope.customAgents;
		this.tools = _scope.tools;
		this.isResolved = derived(this, reader => _scope.isResolved.read(reader) && this._ready.read(reader));
		this._register(autorun(reader => {
			const ids = mcpService.servers.read(reader).filter(server => isAgentHostMcpGatewayServer(server, server.readDefinitions().read(reader).collection, roots)).map(server => server.definition.id);
			const mail = directMail.connections.read(reader).filter(connection => !connection.needsSignIn).map(directMailGatewayId);
			this._refresh([...ids, ...mail].sort());
		}));
	}

	activeClient(clientId: string, resourceUris: IAgentHostResourceUriMapper): IObservable<SessionActiveClient> {
		const base = this._scope.activeClient(clientId, resourceUris);
		return derived(this, reader => {
			const client = base.read(reader);
			return { ...client, customizations: [...client.customizations ?? [], ...this._refs.read(reader).map(ref => toAgentHostClientCustomization(ref, resourceUris))] };
		});
	}

	async whenResolved(): Promise<void> {
		if (this._error) { this._refresh(this._selectedIds, true); }
		await Promise.all([this._scope.whenResolved(), waitForState(this._ready, ready => ready, undefined, this._cancellation.token)]);
		if (this._error) { throw this._error; }
	}

	getSyncedUri(sourceUri: URI): URI | undefined { return this._scope.getSyncedUri(sourceUri); }

	private _refresh(ids: readonly string[], retry = false): void {
		// Discovery and definition updates can republish the same selection. The
		// gateway already observes its servers; reopening it revokes active routes.
		if (!retry && this._generation > 0 && equals(ids, this._selectedIds)) { return; }
		this._selectedIds = ids;
		const generation = ++this._generation;
		this._gateway.clear();
		this._refs.set([], undefined);
		this._ready.set(false, undefined);
		this._error = undefined;
		if (ids.length === 0) { this._ready.set(true, undefined); return; }
		void this._open(generation, ids);
	}

	private async _open(generation: number, ids: readonly string[]): Promise<void> {
		try {
			const gateway = await this._gatewayService.createGateway(false, this._sessionResource, ids);
			if (!gateway) { throw new Error('Local gateway unavailable'); }
			if (!this._isCurrent(generation)) { gateway.dispose(); return; }
			const store = new DisposableStore();
			this._gateway.value = store;
			store.add(gateway);
			store.add(gateway.onDidChangeServers(servers => {
				void this._queue.queue(() => this._publish(generation, servers));
			}));
			await this._queue.queue(() => this._publish(generation, gateway.servers));
		} catch { this._failed(generation); }
	}

	private async _publish(generation: number, servers: readonly IMcpGatewayResultServer[]): Promise<void> {
		if (!this._isCurrent(generation)) { return; }
		try {
			if (new Set(servers.map(server => server.label)).size !== servers.length) { throw new Error('Ambiguous MCP server names'); }
			const bundle = await this._bundler.bundle([], servers.map(server => ({
				name: server.label,
				configuration: { type: McpServerType.REMOTE, url: server.address.toString() },
				enablement: [{ kind: CustomizationEnablementKind.Global, enabled: true }],
			})));
			if (this._isCurrent(generation)) {
				this._refs.set(bundle ? [bundle.ref] : [], undefined);
				this._ready.set(true, undefined);
			}
		} catch { this._failed(generation); }
	}

	private _isCurrent(generation: number): boolean {
		return !this._store.isDisposed && generation === this._generation;
	}

	private _failed(generation: number): void {
		if (!this._isCurrent(generation)) { return; }
		this._gateway.clear();
		this._refs.set([], undefined);
		this._error = new Error(localize('agentHost.mcp.gatewayUnavailable', "Could not connect this chat to its editor MCP servers. Check that enabled servers have distinct names, then retry the message."));
		this._ready.set(true, undefined);
	}
}
