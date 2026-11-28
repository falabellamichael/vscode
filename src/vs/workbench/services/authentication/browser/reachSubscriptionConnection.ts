/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { timeout } from '../../../../base/common/async.js';
import { encodeBase64, VSBuffer } from '../../../../base/common/buffer.js';
import { CancellationToken, CancellationTokenSource } from '../../../../base/common/cancellation.js';
import { CancellationError, isCancellationError } from '../../../../base/common/errors.js';
import { Event } from '../../../../base/common/event.js';
import { isWeb } from '../../../../base/common/platform.js';
import { isObject } from '../../../../base/common/types.js';
import { generateUuid } from '../../../../base/common/uuid.js';
import { localize } from '../../../../nls.js';
import { InstantiationType, registerSingleton } from '../../../../platform/instantiation/common/extensions.js';
import { createDecorator } from '../../../../platform/instantiation/common/instantiation.js';
import { AssistantNewsFeed, AssistantNewsScope, AssistantWeatherDay, AssistantWeatherForecast, AssistantWeatherLocation, INativeHostService, REACH_ACCOUNT_ORIGIN, ReachAssistantModel, ReachAssistantRequest, ReachAssistantResponse, ReachAuthenticationRequest } from '../../../../platform/native/common/native.js';
import { IProgressService, ProgressLocation } from '../../../../platform/progress/common/progress.js';
import { ISecretStorageService } from '../../../../platform/secrets/common/secrets.js';

const SESSION_SECRET = 'sessions.reachSubscription.session';
const WEATHER_NETWORK_KEY = 'sessions.assistant.weatherNetworkKey';

export interface IReachAccount {
	readonly id: string;
	readonly subscriptionActive: boolean;
}

interface IReachSession {
	readonly account: IReachAccount;
	readonly accessToken: string;
	readonly expiresAt: string;
}

function isRecord(value: unknown): value is Record<string, unknown> { return isObject(value); }

function validSession(value: unknown): value is Omit<IReachSession, 'account'> & { account: Record<string, unknown> & { id: string } } {
	return isRecord(value) && typeof value.accessToken === 'string' && value.accessToken.length >= 20 && value.accessToken.length <= 8192
		&& typeof value.expiresAt === 'string' && Date.parse(value.expiresAt) > Date.now()
		&& isRecord(value.account) && typeof value.account.id === 'string' && value.account.id.length > 0 && value.account.id.length <= 128;
}

export const IReachSubscriptionConnection = createDecorator<IReachSubscriptionConnection>('reachSubscriptionConnection');

export interface IReachSubscriptionConnection {
	readonly _serviceBrand: undefined;
	readonly onDidChangeAccount: Event<void>;
	/** REACH assistant requests are fulfilled by the desktop's Electron main process. */
	readonly assistantSupported: boolean;
	getAccount(): Promise<IReachAccount | undefined>;
	getLocalChatCredential(): Promise<{ readonly accessToken: string; readonly expiresAt: string } | undefined>;
	signIn(): Promise<IReachAccount | undefined>;
	signOut(): Promise<void>;
	listAssistantModels(accountId: string): Promise<ReachAssistantModel[]>;
	readAssistantNewsFeeds(scope: AssistantNewsScope): Promise<readonly AssistantNewsFeed[]>;
	readAssistantWeatherForecast(day: AssistantWeatherDay, location: AssistantWeatherLocation): Promise<AssistantWeatherForecast>;
	hasWeatherNetworkKey(): Promise<boolean>;
	saveWeatherNetworkKey(apiKey: string): Promise<void>;
	removeWeatherNetworkKey(): Promise<void>;
	/** Reuse the same cancellable token for all rounds of one turn; cancel it to release the turn. */
	sendAssistantPrompt(accountId: string, request: ReachAssistantRequest, requestId?: string, token?: CancellationToken, conversationId?: string, onDelta?: (delta: string) => void): Promise<ReachAssistantResponse>;
}

export class ReachSubscriptionConnection implements IReachSubscriptionConnection {
	declare readonly _serviceBrand: undefined;
	readonly assistantSupported = !isWeb;
	private signingIn: Promise<IReachAccount | undefined> | undefined;
	private readonly assistantRequests = new Set<string>();
	private readonly assistantTurns = new WeakMap<CancellationToken, string>();
	readonly onDidChangeAccount: Event<void>;

	constructor(
		@INativeHostService private readonly nativeHost: INativeHostService,
		@ISecretStorageService private readonly secrets: ISecretStorageService,
		@IProgressService private readonly progress: IProgressService,
	) {
		this.onDidChangeAccount = Event.map(Event.filter(secrets.onDidChangeSecret, key => key === SESSION_SECRET), () => undefined);
	}

	async getAccount(): Promise<IReachAccount | undefined> {
		return (await this.readSession())?.account;
	}

	async getLocalChatCredential(): Promise<{ readonly accessToken: string; readonly expiresAt: string } | undefined> {
		const session = await this.readSession();
		return session && { accessToken: session.accessToken, expiresAt: session.expiresAt };
	}

	signIn(): Promise<IReachAccount | undefined> {
		return this.signingIn ??= this.openSignIn().finally(() => { this.signingIn = undefined; });
	}

	async signOut(): Promise<void> {
		const session = await this.readSession();
		await this.secrets.delete(SESSION_SECRET);
		if (session) {
			await this.revoke(session.accessToken);
		}
	}

	async listAssistantModels(accountId: string): Promise<ReachAssistantModel[]> {
		const session = await this.readSession();
		if (session?.account.id !== accountId || !session.account.subscriptionActive) { return []; }
		const models = await this.nativeHost.listReachAssistantModels(session.accessToken);
		return (await this.readSession())?.accessToken === session.accessToken ? models : [];
	}

	readAssistantNewsFeeds(scope: AssistantNewsScope): Promise<readonly AssistantNewsFeed[]> {
		return this.nativeHost.readAssistantNewsFeeds(scope);
	}

	async readAssistantWeatherForecast(day: AssistantWeatherDay, location: AssistantWeatherLocation): Promise<AssistantWeatherForecast> {
		const key = await this.secrets.get(WEATHER_NETWORK_KEY);
		// Without a key, or when The Weather Network is unavailable, the native host
		// falls back to the public Open-Meteo service. Key verification stays strict.
		return this.nativeHost.readAssistantWeatherForecast(day, location, key ?? '', true);
	}

	async hasWeatherNetworkKey(): Promise<boolean> {
		return !!(await this.secrets.get(WEATHER_NETWORK_KEY));
	}

	async saveWeatherNetworkKey(apiKey: string): Promise<void> {
		const key = apiKey.trim();
		// Verify forecast access before saving the key with a known provider location.
		await this.nativeHost.readAssistantWeatherForecast('tomorrow', { city: 'Toronto', country: 'CA', subdivision: 'ON' }, key);
		await this.secrets.set(WEATHER_NETWORK_KEY, key);
	}

	removeWeatherNetworkKey(): Promise<void> {
		return this.secrets.delete(WEATHER_NETWORK_KEY);
	}

	async sendAssistantPrompt(accountId: string, request: ReachAssistantRequest, requestId = generateUuid(), token = CancellationToken.None, conversationId = '', onDelta?: (delta: string) => void): Promise<ReachAssistantResponse> {
		const key = JSON.stringify([accountId, conversationId]);
		if (this.assistantRequests.has(key)) { return { kind: 'busy' }; }
		this.assistantRequests.add(key);
		const chunks = onDelta ? this.nativeHost.onDidReceiveReachAssistantChunk(chunk => {
			if (chunk.requestId === requestId) { onDelta(chunk.delta); }
		}) : undefined;
		try {
			const session = await this.readSession();
			if (token.isCancellationRequested) { return { kind: 'stopped' }; }
			if (!session || session.account.id !== accountId) { return { kind: 'signedOut' }; }
			if (!session.account.subscriptionActive) { return { kind: 'subscriptionRequired' }; }
			let turnId = this.assistantTurns.get(token);
			if (!turnId && token !== CancellationToken.None) {
				const id = generateUuid();
				this.assistantTurns.set(token, id);
				const release = token.onCancellationRequested(() => {
					release.dispose();
					this.assistantTurns.delete(token);
					void this.nativeHost.finishReachAssistantTurn(id).catch(() => { });
				});
				turnId = id;
			}
			const pending = this.nativeHost.sendReachAssistantPrompt(accountId, session.accessToken, request, requestId, conversationId, turnId);
			const response = await pending;
			const current = await this.readSession();
			if (token.isCancellationRequested) { return { kind: 'stopped' }; }
			return current?.accessToken === session.accessToken ? response : { kind: 'signedOut' };
		} catch {
			return { kind: 'unavailable' };
		} finally {
			chunks?.dispose();
			this.assistantRequests.delete(key);
		}
	}

	private async readSession(): Promise<IReachSession | undefined> {
		try {
			const raw = await this.secrets.get(SESSION_SECRET);
			const session: unknown = raw ? JSON.parse(raw) : undefined;
			return validSession(session) ? { accessToken: session.accessToken, expiresAt: session.expiresAt,
				account: { id: session.account.id, subscriptionActive: session.account.subscriptionActive === true } } : undefined;
		} catch {
			return undefined;
		}
	}

	private async openSignIn(): Promise<IReachAccount | undefined> {
		const saved = await this.getAccount();
		if (saved) { return saved; }
		const cancellation = new CancellationTokenSource();
		try {
			return await this.progress.withProgress({
				location: ProgressLocation.Notification, cancellable: true,
				title: localize('subscriptions.reachWaiting', "Finish signing in to REACH in your browser"),
			}, () => this.finishBrowserSignIn(cancellation.token), () => cancellation.cancel());
		} catch (error) {
			if (!isCancellationError(error)) { throw error; }
			return undefined;
		} finally {
			cancellation.dispose();
		}
	}

	private async finishBrowserSignIn(token: CancellationToken, retryUnrecognizedFlow = true): Promise<IReachAccount> {
		const nonce = () => encodeBase64(VSBuffer.wrap(crypto.getRandomValues(new Uint8Array(32))), false, true);
		const state = nonce();
		const codeVerifier = nonce();
		const digest = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(codeVerifier));
		const codeChallenge = encodeBase64(VSBuffer.wrap(new Uint8Array(digest)), false, true);
		const { data } = await this.post({ action: 'start', state, codeChallenge }, token);
		if (!isRecord(data) || typeof data.flowId !== 'string' || !/^[a-f\d]{64}$/i.test(data.flowId) || typeof data.expiresAt !== 'string' || typeof data.loginUrl !== 'string') {
			throw this.invalidResponse();
		}
		const expiry = Date.parse(data.expiresAt);
		const expectedServiceUrl = `${REACH_ACCOUNT_ORIGIN}/wallet/connect#flow=${data.flowId}`;
		if (data.loginUrl !== expectedServiceUrl || !(expiry > Date.now() && expiry <= Date.now() + 15 * 60_000)) {
			throw this.invalidResponse();
		}
		if (token.isCancellationRequested) { throw new CancellationError(); }
		// Use the same hosted sign-in handoff as REACH Studio to avoid the tunnel's browser warning.
		const browserUrl = `https://signalreach.ca/login.html#client=studio&flow=${data.flowId}`;
		if (!await this.nativeHost.openExternal(browserUrl)) {
			throw new Error(localize('subscriptions.reachBrowserUnavailable', "The browser could not be opened for REACH sign-in."));
		}
		let confirmedPending = false;
		while (Date.now() < expiry) {
			const exchange = await this.post({ action: 'exchange', flowId: data.flowId, state, codeVerifier }, token);
			if (exchange.status === 202) {
				confirmedPending = true;
				await timeout(2_000, token);
				continue;
			}
			if (exchange.status === 410) {
				if (retryUnrecognizedFlow && !confirmedPending) {
					return this.finishBrowserSignIn(token, false);
				}
				break;
			}
			const result = exchange.data;
			if (!validSession(result)) { throw this.invalidResponse(); }
			const account = { id: result.account.id, subscriptionActive: isRecord(result.account.access) && result.account.access.activeSubscription === true };
			try {
				if (token.isCancellationRequested) { throw new CancellationError(); }
				await this.secrets.set(SESSION_SECRET, JSON.stringify({ accessToken: result.accessToken, expiresAt: result.expiresAt, account }));
			} catch (error) {
				await this.revoke(result.accessToken);
				throw error;
			}
			return account;
		}
		throw new Error(localize('subscriptions.reachTimeout', "REACH sign-in timed out. Open Agent Subscriptions to try again."));
	}

	private async revoke(accessToken: string): Promise<void> {
		try { await this.post({ action: 'logout', accessToken }, CancellationToken.None); } catch { /* The local sign-in remains cleared. */ }
	}

	private async post(request: ReachAuthenticationRequest, token: CancellationToken) {
		if (token.isCancellationRequested) { throw new CancellationError(); }
		let response;
		try {
			response = await this.nativeHost.requestReachAuthentication(request);
		} catch {
			if (token.isCancellationRequested) { throw new CancellationError(); }
			throw new Error(localize('subscriptions.reachNetworkError', "Could not reach the REACH sign-in service. Open Agent Subscriptions to try again."));
		}
		if ((response.status < 200 || response.status >= 300) && !(request.action === 'exchange' && response.status === 410)) {
			if (token.isCancellationRequested) { throw new CancellationError(); }
			throw new Error(localize('subscriptions.reachRequestFailed', "REACH sign-in request failed (HTTP {0}).", response.status));
		}
		return response;
	}

	private invalidResponse(): Error {
		return new Error(localize('subscriptions.reachInvalidResponse', "REACH returned an invalid sign-in response."));
	}
}

registerSingleton(IReachSubscriptionConnection, ReachSubscriptionConnection, InstantiationType.Delayed);
