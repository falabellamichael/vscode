/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { IDisposable } from '../../../base/common/lifecycle.js';
import { ReachAssistantRequest, ReachAssistantResponse } from '../common/native.js';
import { createReachTurnRequest, requestReachAssistant } from '../common/reachChatTransport.js';

interface ReachAssistantTurn {
	readonly windowId: number | undefined;
	readonly accountId: string;
	readonly conversationId: string;
	readonly accessToken: string;
	readonly modelId: string | undefined;
	readonly controller: AbortController;
	readonly send: ReturnType<typeof createReachTurnRequest>;
}

interface ReachAssistantActiveRequest {
	readonly windowId: number | undefined;
	readonly controller: AbortController;
}

/** Reuses one verified account and model preflight across the rounds of a single assistant turn. */
export class ReachAssistantTurns implements IDisposable {
	private readonly turns = new Map<string, ReachAssistantTurn>();
	private readonly active = new Map<string, ReachAssistantActiveRequest>();

	constructor(private readonly fetcher: Parameters<typeof createReachTurnRequest>[0]) { }

	async send(windowId: number | undefined, accountId: string, accessToken: string, request: ReachAssistantRequest, requestId: string, conversationId: string, turnId: string | undefined, onDelta: (delta: string) => void): Promise<ReachAssistantResponse> {
		const conversationKey = JSON.stringify([accountId, conversationId]);
		if (this.active.has(conversationKey)) { return { kind: 'busy' }; }
		const turnKey = turnId === undefined ? undefined : JSON.stringify([windowId, turnId]);
		let turn = turnKey === undefined ? undefined : this.turns.get(turnKey);
		if (turn) {
			if (turn.accountId !== accountId || turn.accessToken !== accessToken) { return { kind: 'signedOut' }; }
			if (turn.conversationId !== conversationId || turn.modelId !== request.modelId) { return { kind: 'unavailable' }; }
		} else if (turnKey !== undefined) {
			const controller = new AbortController();
			turn = { windowId, accountId, accessToken, conversationId, modelId: request.modelId, controller,
				send: createReachTurnRequest(this.fetcher, accountId, accessToken, request.modelId, controller.signal, 'assistant') };
			this.turns.set(turnKey, turn);
		}
		const controller = turn?.controller ?? new AbortController();
		this.active.set(conversationKey, { windowId, controller });
		try {
			return turn ? await turn.send(request, requestId, onDelta)
				: await requestReachAssistant(this.fetcher, accountId, accessToken, request, requestId, controller.signal, onDelta);
		} finally {
			this.active.delete(conversationKey);
		}
	}

	finish(windowId: number | undefined, turnId: string): void {
		this.release(JSON.stringify([windowId, turnId]));
	}

	finishWindow(windowId: number): void {
		for (const [key, turn] of this.turns) {
			if (turn.windowId === windowId) { this.release(key); }
		}
		for (const request of this.active.values()) {
			if (request.windowId === windowId) { request.controller.abort(); }
		}
	}

	dispose(): void {
		for (const key of this.turns.keys()) { this.release(key); }
		for (const request of this.active.values()) { request.controller.abort(); }
	}

	private release(key: string): void {
		const turn = this.turns.get(key);
		this.turns.delete(key);
		turn?.controller.abort();
	}
}
