/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { CancellationError } from '../../../../base/common/errors.js';
import { generateUuid } from '../../../../base/common/uuid.js';
import { localize } from '../../../../nls.js';
import type { IActiveClient } from '../../common/agent.js';
import { PendingRequestRegistry } from '../../common/pendingRequestRegistry.js';
import { ActionType, type ChatAction } from '../../common/state/sessionActions.js';
import { ToolCallContributorKind, type ToolCallResult, type ToolDefinition } from '../../common/state/protocol/state.js';

function failure(message: string, code: string): ToolCallResult {
	return { success: false, pastTenseMessage: message, error: { message, code } };
}

/** Contributions and outstanding work belong to one exact chat on the agent host. */
export class ReachClientTools {
	private readonly clients = new Map<string, IActiveClient>();
	private readonly pending = new PendingRequestRegistry<ToolCallResult, { clientId: string; turnId: string }>();

	constructor(private readonly emit: (action: ChatAction) => void) { }

	getOrCreateClient(clientId: string, displayName: string | undefined): IActiveClient {
		const existing = this.clients.get(clientId);
		if (existing) { return existing; }
		const client: IActiveClient = { clientId, displayName, customizations: [], tools: [] };
		this.clients.set(clientId, client);
		return client;
	}

	private owner(senderClientId?: string): IActiveClient | undefined {
		if (senderClientId !== undefined) { return this.clients.get(senderClientId); }
		// Older senders omit the id. With multiple windows there is no unambiguous target.
		return this.clients.size === 1 ? this.clients.values().next().value : undefined;
	}

	definitions(senderClientId?: string): readonly ToolDefinition[] {
		return this.owner(senderClientId)?.tools ?? [];
	}

	async invoke(turnId: string, name: string, args: Record<string, unknown>, signal: AbortSignal, senderClientId?: string): Promise<ToolCallResult> {
		if (signal.aborted) { throw new CancellationError(); }
		const owner = this.owner(senderClientId);
		const definition = owner?.tools.find(tool => tool.name === name);
		if (!owner || !definition) {
			return failure(localize('reach.tool.unavailable', "The tool {0} is unavailable in this chat.", name), 'toolUnavailable');
		}

		const { clientId } = owner;
		const toolCallId = generateUuid();
		const displayName = definition.title || name;
		const toolInput = JSON.stringify(args);
		const onAbort = () => this.finish(toolCallId, failure(localize('reach.tool.cancelled', "The tool call was interrupted."), 'cancelled'));
		signal.addEventListener('abort', onAbort, { once: true });
		try {
			const result = await this.pending.registerAndFire(toolCallId, () => {
				this.emit({ type: ActionType.ChatToolCallStart, turnId, toolCallId, toolName: name, displayName, contributor: { kind: ToolCallContributorKind.Client, clientId } });
				// The owning workbench evaluates the real tool's approvals before running it.
				if (this.pending.has(toolCallId)) {
					this.emit({ type: ActionType.ChatToolCallReady, turnId, toolCallId, toolInput, invocationMessage: localize('reach.tool.calling', "Calling {0}", displayName) });
				}
			}, { clientId, turnId });
			if (signal.aborted) { throw new CancellationError(); }
			return result;
		} finally {
			signal.removeEventListener('abort', onAbort);
		}
	}

	complete(toolCallId: string, result: ToolCallResult): void {
		this.pending.respond(toolCallId, result);
	}

	deny(toolCallId: string): void {
		// The host already marked this call cancelled upon the user's denial.
		this.complete(toolCallId, failure(localize('reach.tool.denied', "Permission to run the tool was declined."), 'denied'));
	}

	removeClient(clientId: string): void {
		this.clients.delete(clientId);
		for (const [id, owner] of this.pending.entries()) {
			if (owner.clientId === clientId) {
				this.finish(id, failure(localize('reach.tool.disconnected', "The client disconnected before reporting the tool result."), 'clientDisconnected'));
			}
		}
	}

	private finish(toolCallId: string, result: ToolCallResult): void {
		const owner = this.pending.getMetadata(toolCallId);
		if (owner && this.pending.respond(toolCallId, result)) {
			this.emit({ type: ActionType.ChatToolCallComplete, turnId: owner.turnId, toolCallId, result });
		}
	}
}
