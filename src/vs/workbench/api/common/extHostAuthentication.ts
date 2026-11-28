/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import type * as vscode from 'vscode';
import * as nls from '../../../nls.js';
import { Emitter, Event } from '../../../base/common/event.js';
import { MainContext, MainThreadAuthenticationShape, ExtHostAuthenticationShape, ISignalCOAuthCallback } from './extHost.protocol.js';
import { Proxied } from '../../services/extensions/common/proxyIdentifier.js';
import { Disposable, ProgressLocation } from './extHostTypes.js';
import { IExtensionDescription, ExtensionIdentifier } from '../../../platform/extensions/common/extensions.js';
import { getAuthenticationSessionRequestKey, getOAuthTokenFingerprint, IAuthenticationGetSessionsOptions, IAuthenticationProviderSessionOptions, INTERNAL_AUTH_PROVIDER_PREFIX, isGoogleWorkspaceOAuth, isSignalCMailOAuth, MCP_ACCEPT_PARTIAL_OAUTH_SCOPES, MCP_REFRESH_REJECTED_OAUTH_TOKEN, MCP_USE_HTTPS_OAUTH_CALLBACK, signalCMailOAuth } from '../../services/authentication/common/authentication.js';
import { createDecorator } from '../../../platform/instantiation/common/instantiation.js';
import { IExtHostRpcService } from './extHostRpcService.js';
import { URI, UriComponents } from '../../../base/common/uri.js';
import { AuthorizationErrorType, fetchDynamicRegistration, getClaimsFromJWT, IAuthorizationJWTClaims, IAuthorizationProtectedResourceMetadata, IAuthorizationServerMetadata, IAuthorizationTokenResponse, isAuthorizationErrorResponse, isAuthorizationTokenResponse, isSignalCOAuthScheme, scopesMatch } from '../../../base/common/oauth.js';
import { IExtHostWindow } from './extHostWindow.js';
import { IExtHostInitDataService } from './extHostInitDataService.js';
import { ILogger, ILoggerService, ILogService } from '../../../platform/log/common/log.js';
import { autorun, derivedOpts, IObservable, ISettableObservable, observableValue } from '../../../base/common/observable.js';
import { stringHash } from '../../../base/common/hash.js';
import { DisposableStore, IDisposable } from '../../../base/common/lifecycle.js';
import { IExtHostUrlsService } from './extHostUrls.js';
import { decodeBase64, encodeBase64, VSBuffer } from '../../../base/common/buffer.js';
import { equals as arraysEqual } from '../../../base/common/arrays.js';
import { IExtHostProgress } from './extHostProgress.js';
import { NotificationTelemetryId } from '../../../platform/notification/common/notificationTelemetry.js';
import { IProgressStep } from '../../../platform/progress/common/progress.js';
import { CancellationError, isCancellationError } from '../../../base/common/errors.js';
import { Sequencer, SequencerByKey } from '../../../base/common/async.js';
import { CancellationTokenSource } from '../../../base/common/cancellation.js';
import { XaaifyAuthProvider } from './extHostXaaAuthProvider.js';

export interface IExtHostAuthentication extends ExtHostAuthentication { }
export const IExtHostAuthentication = createDecorator<IExtHostAuthentication>('IExtHostAuthentication');

interface ProviderWithMetadata {
	label: string;
	provider: vscode.AuthenticationProvider;
	disposable?: vscode.Disposable;
	options: vscode.AuthenticationProviderOptions;
}

/**
 * The account icon is a {@link vscode.Uri} that does not survive being sent over the RPC boundary,
 * so it needs to be revived when an account is received from the main thread.
 */
export function reviveAccountIcon<T extends { readonly icon?: vscode.Uri | UriComponents }>(account: T): T & { readonly icon?: vscode.Uri } {
	return { ...account, icon: URI.revive(account.icon) };
}

function getInteractiveOptionsForRequestKey(options: boolean | vscode.AuthenticationGetSessionPresentationOptions | undefined) {
	return typeof options === 'object'
		? { detail: options.detail, learnMore: options.learnMore?.toString() }
		: options;
}

export class ExtHostAuthentication implements ExtHostAuthenticationShape {

	declare _serviceBrand: undefined;

	protected readonly _dynamicAuthProviderCtor = DynamicAuthProvider;
	protected readonly _xaaAuthProviderCtor = XaaifyAuthProvider(DynamicAuthProvider);

	private _proxy: Proxied<MainThreadAuthenticationShape>;
	private _authenticationProviders: Map<string, ProviderWithMetadata> = new Map<string, ProviderWithMetadata>();
	private _providerOperations = new SequencerByKey<string>();

	private _onDidChangeSessions = new Emitter<vscode.AuthenticationSessionsChangeEvent & { extensionIdFilter?: string[] }>();
	private _getSessionTaskSingler = new TaskSingler<vscode.AuthenticationSession | undefined>();

	private _onDidDynamicAuthProviderTokensChange = new Emitter<{ authProviderId: string; clientId: string; tokens: IAuthorizationToken[] }>();

	constructor(
		@IExtHostRpcService extHostRpc: IExtHostRpcService,
		@IExtHostInitDataService private readonly _initData: IExtHostInitDataService,
		@IExtHostWindow private readonly _extHostWindow: IExtHostWindow,
		@IExtHostUrlsService private readonly _extHostUrls: IExtHostUrlsService,
		@IExtHostProgress private readonly _extHostProgress: IExtHostProgress,
		@ILoggerService private readonly _extHostLoggerService: ILoggerService,
		@ILogService private readonly _logService: ILogService,
	) {
		this._proxy = extHostRpc.getProxy(MainContext.MainThreadAuthentication);
	}

	/**
	 * This sets up an event that will fire when the auth sessions change with a built-in filter for the extensionId
	 * if a session change only affects a specific extension.
	 * @param extensionId The extension that is interested in the event.
	 * @returns An event with a built-in filter for the extensionId
	 */
	getExtensionScopedSessionsEvent(extensionId: string): Event<vscode.AuthenticationSessionsChangeEvent> {
		const normalizedExtensionId = extensionId.toLowerCase();
		return Event.chain(this._onDidChangeSessions.event, ($) => $
			.filter(e => !e.extensionIdFilter || e.extensionIdFilter.includes(normalizedExtensionId))
			.map(e => ({ provider: e.provider }))
		);
	}

	async getSession(requestingExtension: IExtensionDescription, providerId: string, scopesOrRequest: readonly string[] | vscode.AuthenticationWwwAuthenticateRequest, options: vscode.AuthenticationGetSessionOptions & ({ createIfNone: true } | { forceNewSession: true } | { forceNewSession: vscode.AuthenticationForceNewSessionOptions })): Promise<vscode.AuthenticationSession>;
	async getSession(requestingExtension: IExtensionDescription, providerId: string, scopesOrRequest: readonly string[] | vscode.AuthenticationWwwAuthenticateRequest, options: vscode.AuthenticationGetSessionOptions & { forceNewSession: true }): Promise<vscode.AuthenticationSession>;
	async getSession(requestingExtension: IExtensionDescription, providerId: string, scopesOrRequest: readonly string[] | vscode.AuthenticationWwwAuthenticateRequest, options: vscode.AuthenticationGetSessionOptions & { forceNewSession: vscode.AuthenticationForceNewSessionOptions }): Promise<vscode.AuthenticationSession>;
	async getSession(requestingExtension: IExtensionDescription, providerId: string, scopesOrRequest: readonly string[] | vscode.AuthenticationWwwAuthenticateRequest, options: vscode.AuthenticationGetSessionOptions): Promise<vscode.AuthenticationSession | undefined>;
	async getSession(requestingExtension: IExtensionDescription, providerId: string, scopesOrRequest: readonly string[] | vscode.AuthenticationWwwAuthenticateRequest, options: vscode.AuthenticationGetSessionOptions = {}): Promise<vscode.AuthenticationSession | undefined> {
		const extensionId = ExtensionIdentifier.toKey(requestingExtension.identifier);
		const singlerKey = JSON.stringify([extensionId, providerId, getAuthenticationSessionRequestKey(scopesOrRequest, {
			...options,
			account: options.account && { id: options.account.id, label: options.account.label },
			authorizationServer: URI.revive(options.authorizationServer),
			createIfNone: getInteractiveOptionsForRequestKey(options.createIfNone),
			forceNewSession: getInteractiveOptionsForRequestKey(options.forceNewSession)
		})]);

		return await this._getSessionTaskSingler.getOrCreate(singlerKey, async () => {
			await this._proxy.$ensureProvider(providerId);
			const extensionName = requestingExtension.displayName || requestingExtension.name;
			const session = await this._proxy.$getSession(providerId, scopesOrRequest, extensionId, extensionName, options);
			return session && {
				...session,
				account: reviveAccountIcon(session.account),
				authorizationServer: URI.revive(session.authorizationServer)
			};
		});
	}

	async getAccounts(providerId: string) {
		await this._proxy.$ensureProvider(providerId);
		const accounts = await this._proxy.$getAccounts(providerId);
		return accounts.map(account => reviveAccountIcon(account));
	}

	registerAuthenticationProvider(id: string, label: string, provider: vscode.AuthenticationProvider, options?: vscode.AuthenticationProviderOptions): vscode.Disposable {
		const disposables = new DisposableStore();
		// Capture changes before queueing, but forward them only after the main thread acknowledges registration.
		const bufferedSessionChanges = Event.buffer<vscode.AuthenticationProviderAuthenticationSessionsChangeEvent>(
			listener => provider.onDidChangeSessions(listener), 'authentication provider registration', false, [], disposables);
		const providerData: ProviderWithMetadata = { label, provider, disposable: disposables, options: options ?? { supportsMultipleAccounts: false } };
		// register
		void this._providerOperations.queue(id, async () => {
			// This use to be synchronous, but that wasn't an accurate representation because the main thread
			// may have unregistered the provider in the meantime. I don't see how this could really be done
			// synchronously, so we just say first one wins.
			if (this._authenticationProviders.get(id)) {
				disposables.dispose();
				this._logService.error(`An authentication provider with id '${id}' is already registered. The existing provider will not be replaced.`);
				return;
			}
			this._authenticationProviders.set(id, providerData);
			try {
				await this._proxy.$registerAuthenticationProvider({
					id,
					label,
					supportsMultipleAccounts: options?.supportsMultipleAccounts ?? false,
					supportedAuthorizationServers: options?.supportedAuthorizationServers,
					supportsChallenges: options?.supportsChallenges
				});
				disposables.add(bufferedSessionChanges(e => this._proxy.$sendDidChangeSessions(id, e)));
			} catch (error) {
				disposables.dispose();
				this._authenticationProviders.delete(id);
				this._logService.error(`Failed to register authentication provider '${id}'.`, error);
			}
		});

		// unregister
		return new Disposable(() => {
			void this._providerOperations.queue(id, async () => {
				if (this._authenticationProviders.get(id) === providerData) {
					disposables.dispose();
					this._authenticationProviders.delete(id);
					await this._proxy.$unregisterAuthenticationProvider(id);
				}
			});
		});
	}

	$createSession(providerId: string, scopes: string[], options: vscode.AuthenticationProviderSessionOptions): Promise<vscode.AuthenticationSession> {
		return this._providerOperations.queue(providerId, async () => {
			const providerData = this._authenticationProviders.get(providerId);
			if (providerData) {
				options.authorizationServer = URI.revive(options.authorizationServer);
				if (options.account) {
					options.account = reviveAccountIcon(options.account);
				}
				return await providerData.provider.createSession(scopes, options);
			}

			throw new Error(`Unable to find authentication provider with handle: ${providerId}`);
		});
	}

	$removeSession(providerId: string, sessionId: string): Promise<void> {
		return this._providerOperations.queue(providerId, async () => {
			const providerData = this._authenticationProviders.get(providerId);
			if (providerData) {
				return await providerData.provider.removeSession(sessionId);
			}

			throw new Error(`Unable to find authentication provider with handle: ${providerId}`);
		});
	}

	$getSessions(providerId: string, scopes: ReadonlyArray<string> | undefined, options: IAuthenticationGetSessionsOptions): Promise<ReadonlyArray<vscode.AuthenticationSession>> {
		return this._providerOperations.queue(providerId, async () => {
			const providerData = this._authenticationProviders.get(providerId);
			if (providerData) {
				options.authorizationServer = URI.revive(options.authorizationServer);
				if (options.account) {
					options.account = reviveAccountIcon(options.account);
				}
				return await providerData.provider.getSessions(scopes, options);
			}

			throw new Error(`Unable to find authentication provider with handle: ${providerId}`);
		});
	}

	$getSessionsFromChallenges(providerId: string, constraint: vscode.AuthenticationConstraint, options: vscode.AuthenticationProviderSessionOptions): Promise<ReadonlyArray<vscode.AuthenticationSession>> {
		return this._providerOperations.queue(providerId, async () => {
			const providerData = this._authenticationProviders.get(providerId);
			if (providerData) {
				const provider = providerData.provider;
				// Check if provider supports challenges
				if (typeof provider.getSessionsFromChallenges === 'function') {
					options.authorizationServer = URI.revive(options.authorizationServer);
					if (options.account) {
						options.account = reviveAccountIcon(options.account);
					}
					return await provider.getSessionsFromChallenges(constraint, options);
				}
				throw new Error(`Authentication provider with handle: ${providerId} does not support getSessionsFromChallenges`);
			}

			throw new Error(`Unable to find authentication provider with handle: ${providerId}`);
		});
	}

	$createSessionFromChallenges(providerId: string, constraint: vscode.AuthenticationConstraint, options: vscode.AuthenticationProviderSessionOptions): Promise<vscode.AuthenticationSession> {
		return this._providerOperations.queue(providerId, async () => {
			const providerData = this._authenticationProviders.get(providerId);
			if (providerData) {
				const provider = providerData.provider;
				// Check if provider supports challenges
				if (typeof provider.createSessionFromChallenges === 'function') {
					options.authorizationServer = URI.revive(options.authorizationServer);
					if (options.account) {
						options.account = reviveAccountIcon(options.account);
					}
					return await provider.createSessionFromChallenges(constraint, options);
				}
				throw new Error(`Authentication provider with handle: ${providerId} does not support createSessionFromChallenges`);
			}

			throw new Error(`Unable to find authentication provider with handle: ${providerId}`);
		});
	}

	$onDidChangeAuthenticationSessions(id: string, label: string, extensionIdFilter?: string[]) {
		// Don't fire events for the internal auth providers
		if (!id.startsWith(INTERNAL_AUTH_PROVIDER_PREFIX)) {
			this._onDidChangeSessions.fire({ provider: { id, label }, extensionIdFilter });
		}
		return Promise.resolve();
	}

	$onDidUnregisterAuthenticationProvider(id: string): Promise<void> {
		return this._providerOperations.queue(id, async () => {
			const providerData = this._authenticationProviders.get(id);
			if (providerData) {
				providerData.disposable?.dispose();
				this._authenticationProviders.delete(id);
			}
		});
	}

	async $registerDynamicAuthProvider(
		authorizationServerComponents: UriComponents,
		serverMetadata: IAuthorizationServerMetadata,
		resourceMetadata: IAuthorizationProtectedResourceMetadata | undefined,
		clientId: string | undefined,
		clientSecret: string | undefined,
		initialTokens: IAuthorizationToken[] | undefined,
		providerId?: string,
		tokenEndpointAuthMethod?: string,
		signalCOAuthCallback?: ISignalCOAuthCallback,
	): Promise<string> {
		if (!clientId) {
			const authorizationServer = URI.revive(authorizationServerComponents);
			if (serverMetadata.registration_endpoint) {
				try {
					const registration = await fetchDynamicRegistration(serverMetadata, this._initData.environment.appName, resourceMetadata?.scopes_supported, this._initData.environment.appUriScheme);
					clientId = registration.client_id;
					clientSecret = registration.client_secret;
					tokenEndpointAuthMethod = registration.token_endpoint_auth_method;
				} catch (err) {
					this._logService.warn(`Dynamic registration failed for ${authorizationServer.toString()}: ${err.message}. Prompting user for client ID and client secret...`);
				}
			}
			// Still no client id so dynamic client registration was either not supported or failed
			if (!clientId) {
				this._logService.info(`Prompting user for client registration details for ${authorizationServer.toString()}`);
				const clientDetails = await this._proxy.$promptForClientRegistration(authorizationServer.toString());
				if (!clientDetails) {
					throw new Error('User did not provide client details');
				}
				clientId = clientDetails.clientId;
				clientSecret = clientDetails.clientSecret;
				this._logService.info(`User provided client registration for ${authorizationServer.toString()}`);
				if (clientSecret) {
					this._logService.trace(`User provided client secret for ${authorizationServer.toString()}`);
				} else {
					this._logService.trace(`User did not provide client secret for ${authorizationServer.toString()}`);
				}
			}
		}
		const provider = new this._dynamicAuthProviderCtor(
			this._extHostWindow,
			this._extHostUrls,
			this._initData,
			this._extHostProgress,
			this._extHostLoggerService,
			this._proxy,
			URI.revive(authorizationServerComponents),
			serverMetadata,
			resourceMetadata,
			clientId,
			clientSecret,
			this._onDidDynamicAuthProviderTokensChange,
			initialTokens || [],
			providerId,
			tokenEndpointAuthMethod,
			signalCOAuthCallback,
		);

		// Use the sequencer to ensure dynamic provider registration is serialized
		await this._providerOperations.queue(provider.id, async () => {
			this._authenticationProviders.set(
				provider.id,
				{
					label: provider.label,
					provider,
					disposable: Disposable.from(
						provider,
						provider.onDidChangeSessions(e => this._proxy.$sendDidChangeSessions(provider.id, e)),
						provider.onDidChangeClientId(() => this._proxy.$sendDidChangeDynamicProviderInfo({
							providerId: provider.id,
							clientId: provider.clientId,
							clientSecret: provider.clientSecret,
							tokenEndpointAuthMethod: provider.tokenEndpointAuthMethod,
						}))
					),
					options: { supportsMultipleAccounts: true }
				}
			);

			await this._proxy.$registerDynamicAuthenticationProvider({
				id: provider.id,
				label: provider.label,
				supportsMultipleAccounts: true,
				authorizationServer: authorizationServerComponents,
				resourceServer: resourceMetadata ? URI.parse(resourceMetadata.resource) : undefined,
				clientId: provider.clientId,
				clientSecret: provider.clientSecret,
				tokenEndpointAuthMethod: provider.tokenEndpointAuthMethod,
			});
		});




		return provider.id;
	}

	async $registerXaaAuthProvider(
		issuerComponents: UriComponents,
		serverMetadata: IAuthorizationServerMetadata,
		clientId: string | undefined,
		clientSecret: string | undefined,
		initialTokens: IAuthorizationToken[] | undefined
	): Promise<string> {
		const issuer = URI.revive(issuerComponents);
		// XAA does not use Dynamic Client Registration — the IdP must already trust the requesting
		// app for the target audience(s). Always require an admin-provisioned client_id (and
		// typically client_secret).
		if (!clientId) {
			this._logService.info(`Prompting user for client registration details for XAA issuer ${issuer.toString()}`);
			const clientDetails = await this._proxy.$promptForClientRegistration(issuer.toString());
			if (!clientDetails) {
				throw new Error('User did not provide client details');
			}
			clientId = clientDetails.clientId;
			clientSecret = clientDetails.clientSecret;
		}
		const provider = new this._xaaAuthProviderCtor(
			this._extHostWindow,
			this._extHostUrls,
			this._initData,
			this._extHostProgress,
			this._extHostLoggerService,
			this._proxy,
			issuer,
			serverMetadata,
			/* resourceMetadata */ undefined,
			clientId,
			clientSecret,
			this._onDidDynamicAuthProviderTokensChange,
			initialTokens || []
		);

		await this._providerOperations.queue(provider.id, async () => {
			this._authenticationProviders.set(
				provider.id,
				{
					label: provider.label,
					provider,
					disposable: Disposable.from(
						provider,
						provider.onDidChangeSessions(e => this._proxy.$sendDidChangeSessions(provider.id, e)),
						provider.onDidChangeClientId(() => this._proxy.$sendDidChangeDynamicProviderInfo({
							providerId: provider.id,
							clientId: provider.clientId,
							clientSecret: provider.clientSecret
						}))
					),
					options: { supportsMultipleAccounts: true }
				}
			);

			await this._proxy.$registerDynamicAuthenticationProvider({
				id: provider.id,
				label: provider.label,
				supportsMultipleAccounts: true,
				authorizationServer: issuerComponents,
				resourceServer: undefined,
				clientId: provider.clientId,
				clientSecret: provider.clientSecret
			});
		});

		return provider.id;
	}

	async $onDidChangeDynamicAuthProviderTokens(authProviderId: string, clientId: string, tokens?: IAuthorizationToken[]): Promise<void> {
		// Deleting the storage entry clears the provider's token array.
		this._onDidDynamicAuthProviderTokensChange.fire({ authProviderId, clientId, tokens: tokens ?? [] });
	}
}

class TaskSingler<T> {
	private _inFlightPromises = new Map<string, Promise<T>>();
	getOrCreate(key: string, promiseFactory: () => Promise<T>) {
		const inFlight = this._inFlightPromises.get(key);
		if (inFlight) {
			return inFlight;
		}

		const promise = promiseFactory().finally(() => this._inFlightPromises.delete(key));
		this._inFlightPromises.set(key, promise);

		return promise;
	}
}

// Mirrors `MICROSOFT_AUTH_HOSTNAMES` and the `is_microsoft_authorization_server` / `is_microsoft_auth_endpoint_aware`
// matching in github/copilot-agent-runtime `src/runtime/src/auth_base/microsoft.rs`; keep them in sync.
const MICROSOFT_AUTH_HOSTS = [
	'login.microsoftonline.com',
	'login.microsoftonline.de',
	'login.microsoftonline.us',
	'login.partner.microsoftonline.cn',
	'login.microsoft.com',
	'login.windows.net',
	'sts.windows.net',
];

function isMicrosoftAuthUrl(url: string | undefined): boolean {
	if (!url) {
		return false;
	}
	let hostname: string;
	try {
		hostname = new URL(url).hostname;
	} catch {
		return false;
	}
	return MICROSOFT_AUTH_HOSTS.some(host => hostname === host || hostname.endsWith(`.${host}`));
}

export class DynamicAuthProvider implements vscode.AuthenticationProvider {
	id: string;
	readonly label: string;

	private _onDidChangeSessions = new Emitter<vscode.AuthenticationProviderAuthenticationSessionsChangeEvent>();
	readonly onDidChangeSessions = this._onDidChangeSessions.event;

	private readonly _onDidChangeClientId = new Emitter<void>();
	readonly onDidChangeClientId = this._onDidChangeClientId.event;

	private readonly _tokenStore: TokenStore;
	private readonly _refreshSequencer = new Sequencer();

	protected readonly _createFlows: Array<{
		label: string;
		handler: (scopes: string[], progress: vscode.Progress<{ message: string }>, token: vscode.CancellationToken) => Promise<IAuthorizationTokenResponse>;
	}>;

	protected readonly _logger: ILogger;
	protected readonly _isMicrosoftAuth: boolean;
	protected readonly _isSignalCMailOAuth: boolean;
	protected readonly _isGoogleLoopbackOAuth: boolean;
	private readonly _disposable: DisposableStore;

	constructor(
		@IExtHostWindow protected readonly _extHostWindow: IExtHostWindow,
		@IExtHostUrlsService protected readonly _extHostUrls: IExtHostUrlsService,
		@IExtHostInitDataService protected readonly _initData: IExtHostInitDataService,
		@IExtHostProgress private readonly _extHostProgress: IExtHostProgress,
		@ILoggerService loggerService: ILoggerService,
		protected readonly _proxy: Proxied<MainThreadAuthenticationShape>,
		readonly authorizationServer: URI,
		protected readonly _serverMetadata: IAuthorizationServerMetadata,
		protected readonly _resourceMetadata: IAuthorizationProtectedResourceMetadata | undefined,
		protected _clientId: string,
		protected _clientSecret: string | undefined,
		onDidDynamicAuthProviderTokensChange: Emitter<{ authProviderId: string; clientId: string; tokens: IAuthorizationToken[] }>,
		initialTokens: IAuthorizationToken[],
		providerId?: string,
		private _tokenEndpointAuthMethod?: string,
		private readonly _signalCOAuthCallback?: ISignalCOAuthCallback,
		private readonly _fetch: typeof fetch = fetch,
	) {
		const stringifiedServer = authorizationServer.toString(true);
		// The default provider ID combines the authorization server and the resource, if provided.
		this.id = providerId ?? (_resourceMetadata?.resource
			? stringifiedServer + ' ' + _resourceMetadata?.resource
			: stringifiedServer);
		// Auth Provider label is just the resource name if provided, otherwise the authority of the authorization server.
		this.label = _resourceMetadata?.resource_name ?? this.authorizationServer.authority;
		// Metadata is untrusted, so the token endpoint must also be Microsoft to earn Microsoft-specific treatment.
		this._isMicrosoftAuth = (isMicrosoftAuthUrl(stringifiedServer) || isMicrosoftAuthUrl(_serverMetadata.authorization_endpoint))
			&& isMicrosoftAuthUrl(_serverMetadata.token_endpoint);
		this._isSignalCMailOAuth = isSignalCMailOAuth(this.id, authorizationServer, _serverMetadata, _resourceMetadata?.resource);
		this._isGoogleLoopbackOAuth = (this._isSignalCMailOAuth && !this._isMicrosoftAuth)
			|| isGoogleWorkspaceOAuth(authorizationServer, _serverMetadata, _resourceMetadata?.resource ? URI.parse(_resourceMetadata.resource) : undefined);

		this._logger = loggerService.createLogger(this.id, { name: `Auth: ${this.label}` });
		this._disposable = new DisposableStore();
		this._disposable.add(this._onDidChangeSessions);
		this._disposable.add(this._onDidChangeClientId);
		const scopedEvent = Event.chain(onDidDynamicAuthProviderTokensChange.event, $ => $
			.filter(e => e.authProviderId === this.id && e.clientId === _clientId)
			.map(e => e.tokens)
		);
		this._tokenStore = this._disposable.add(new TokenStore(
			{
				onDidChange: scopedEvent,
				set: (tokens) => _proxy.$setSessionsForDynamicAuthProvider(this.id, this.clientId, tokens),
			},
			initialTokens,
			this._logger
		));
		this._disposable.add(this._tokenStore.onDidChangeSessions(e => this._onDidChangeSessions.fire(e)));
		// Will be extended later to support other flows
		this._createFlows = [];
		if (_serverMetadata.authorization_endpoint && !this._isSignalCMailOAuth && !isSignalCOAuthScheme(_initData.environment.appUriScheme)) {
			this._createFlows.push({
				label: nls.localize('url handler', "URL Handler"),
				handler: (scopes, progress, token) => this._createWithUrlHandler(scopes, progress, token)
			});
		}
	}

	get clientId(): string {
		return this._clientId;
	}

	get clientSecret(): string | undefined {
		return this._clientSecret;
	}

	get tokenEndpointAuthMethod(): string | undefined {
		return this._tokenEndpointAuthMethod;
	}

	/**
	 * The RFC 8707 resource indicator to send in requests. Omitted for Microsoft auth.
	 */
	protected get _resourceIndicator(): string | undefined {
		return this._isMicrosoftAuth || this._isSignalCMailOAuth ? undefined : this._resourceMetadata?.resource;
	}

	/**
	 * The scopes to send to the server, adding `offline_access` when the server advertises it.
	 * Empty scopes are left empty so servers can still apply their default scopes.
	 */
	protected _getRequestScopes(scopes: readonly string[]): string[] {
		const requested = [...scopes];
		if (scopes.length && this._isGoogleLoopbackOAuth && !this._isSignalCMailOAuth && !scopes.includes('openid')) {
			// Google Workspace access tokens are opaque. The ID token identifies the
			// account if a later tool needs additional consent for the same account.
			requested.push('openid');
		}
		if (scopes.length && this._serverMetadata.scopes_supported?.includes('offline_access') && !scopes.includes('offline_access')) {
			requested.push('offline_access');
		}
		return requested;
	}

	async getSessions(scopes: readonly string[] | undefined, options: IAuthenticationProviderSessionOptions): Promise<vscode.AuthenticationSession[]> {
		this._logger.info(`Getting sessions for scopes: ${scopes?.join(' ') ?? 'all'}`);
		if (!scopes) {
			return this._tokenStore.sessions;
		}
		// Several MCP requests can ask for a session concurrently. Refresh-token rotation
		// makes the token from the first exchange unusable by the next request.
		return this._refreshSequencer.queue(() => this._getSessionsForScopes(scopes, options));
	}

	private async _getSessionsForScopes(scopes: readonly string[], options: IAuthenticationProviderSessionOptions): Promise<vscode.AuthenticationSession[]> {
		const scopeStr = scopes.join(' ');
		const acceptPartialScopes = options[MCP_ACCEPT_PARTIAL_OAUTH_SCOPES] === true;
		const findSessions = () => {
			const tokensByAccessToken = new Map(this._tokenStore.tokens.map(token => [token.access_token, token]));
			const complete: vscode.AuthenticationSession[] = [];
			const partial: vscode.AuthenticationSession[] = [];
			for (const session of this._tokenStore.sessions) {
				const request = tokensByAccessToken.get(session.accessToken)?.requested_scope;
				if (scopesMatch(session.scopes, scopes)) {
					complete.push(session);
				} else if (request !== undefined && scopesMatch(request.split(' ').filter(Boolean), scopes)) {
					if (this._isUsableScopeGrant(scopes, session.scopes, false)) {
						complete.push(session);
					} else if (acceptPartialScopes && this._isUsableScopeGrant(scopes, session.scopes, true)) {
						partial.push(session);
					}
				}
			}
			const newestFirst = (a: vscode.AuthenticationSession, b: vscode.AuthenticationSession) =>
				(tokensByAccessToken.get(b.accessToken)?.created_at ?? 0) - (tokensByAccessToken.get(a.accessToken)?.created_at ?? 0);
			return [...complete.sort(newestFirst), ...partial.sort(newestFirst)];
		};
		let sessions = findSessions();
		const rejectedFingerprint = options[MCP_REFRESH_REJECTED_OAUTH_TOKEN];
		if (typeof rejectedFingerprint === 'string') {
			for (const session of sessions) {
				if (await getOAuthTokenFingerprint(session.accessToken) === rejectedFingerprint) {
					return this._refreshRejectedSession(session, scopes, acceptPartialScopes);
				}
			}
			return [];
		}
		this._logger.info(`Found ${sessions.length} sessions for scopes: ${scopeStr}`);
		if (sessions.length) {
			const newTokens: IAuthorizationToken[] = [];
			const removedTokens: IAuthorizationToken[] = [];
			const tokenMap = new Map<string, IAuthorizationToken>(this._tokenStore.tokens.map(token => [token.access_token, token]));
			for (const session of sessions) {
				const token = tokenMap.get(session.accessToken);
				if (token && token.expires_in) {
					const now = Date.now();
					const expiresInMS = token.expires_in * 1000;
					// Check if the token is about to expire in 5 minutes or if it is expired
					if (now > token.created_at + expiresInMS - (5 * 60 * 1000)) {
						this._logger.info(`Token for session ${session.id} is about to expire, refreshing...`);
						removedTokens.push(token);
						if (!token.refresh_token) {
							// No refresh token available, cannot refresh
							this._logger.warn(`No refresh token available for scopes ${session.scopes.join(' ')}. Throwing away token.`);
							continue;
						}
						try {
							const newToken = await this._refreshSessionToken(token, token.refresh_token, session.scopes, options.silent !== true);
							this._logger.info(`Successfully created a new token for scopes ${session.scopes.join(' ')}.`);
							newTokens.push(newToken);
						} catch (err) {
							this._logger.error(`Failed to refresh token: ${err}`);
						}

					}
				}
			}
			if (newTokens.length || removedTokens.length) {
				await this._tokenStore.update({ added: newTokens, removed: removedTokens });
				// Since we updated the tokens, we need to re-filter the sessions
				// to get the latest state
				sessions = findSessions();
			}
			this._logger.info(`Found ${sessions.length} sessions for scopes: ${scopeStr}`);
			return sessions;
		}
		return [];
	}

	private async _refreshSessionToken(token: IAuthorizationToken, refreshToken: string, scopes: readonly string[], allowClientRegistration: boolean): Promise<IAuthorizationToken> {
		const refreshed = await this.exchangeRefreshTokenForToken(refreshToken, scopes, allowClientRegistration);
		refreshed.id_token ??= token.id_token;
		// If scope is omitted, inherit the grant being refreshed, even if it was partial.
		refreshed.scope = this._getGrantedScopes(scopes, refreshed).join(' ');
		refreshed.requested_scope = token.requested_scope ?? scopes.join(' ');
		return refreshed;
	}

	private async _refreshRejectedSession(session: vscode.AuthenticationSession, scopes: readonly string[], acceptPartialScopes: boolean): Promise<vscode.AuthenticationSession[]> {
		const token = this._tokenStore.tokens.find(token => token.access_token === session.accessToken);
		if (!token?.refresh_token) {
			return [];
		}
		try {
			const refreshed = await this._refreshSessionToken(token, token.refresh_token, session.scopes, false);
			const replacement = this._tokenStore.getSessionFromToken(refreshed);
			if (replacement.account.id !== session.account.id || replacement.account.label !== session.account.label
				|| !replacement.scopes.every(scope => session.scopes.includes(scope))) {
				this._logger.warn('The OAuth refresh response changed the account or expanded the rejected grant.');
				return [];
			}
			// Another window may have removed or rotated this grant while the exchange was pending.
			if (!this._tokenStore.tokens.some(current => current.access_token === token.access_token && current.refresh_token === token.refresh_token && current.scope === token.scope)) {
				return [];
			}
			await this._tokenStore.update({ added: [refreshed], removed: [token] });
			return replacement.accessToken !== session.accessToken && this._isUsableScopeGrant(scopes, replacement.scopes, acceptPartialScopes) ? [replacement] : [];
		} catch {
			this._logger.warn('Could not refresh the rejected OAuth grant.');
			return [];
		}
	}

	async createSession(scopes: string[], options: IAuthenticationProviderSessionOptions): Promise<vscode.AuthenticationSession> {
		if (this._isSignalCMailOAuth && !this._createFlows.length) {
			throw new Error(nls.localize('directMailOAuth.localRequired', "Direct mail sign-in requires a local Signal-C desktop extension host."));
		}
		if (this._isGoogleLoopbackOAuth && !this._clientSecret) {
			throw new Error(this._isSignalCMailOAuth
				? nls.localize('gmailDirectOAuth.secretRequired', "Enter the secret from your Google Desktop OAuth client in the Gmail connection setup.")
				: nls.localize('googleWorkspaceOAuth.secretRequired', "Google Workspace requires an OAuth client secret. Use Set OAuth Client Secret on the MCP server, then connect again."));
		}
		this._logger.info(`Creating session for scopes: ${scopes.join(' ')}`);
		let flows = this._createFlows;
		if (options[MCP_USE_HTTPS_OAUTH_CALLBACK] === true) {
			const httpsCallback = this._signalCOAuthCallback;
			if (!httpsCallback || !/^[1-9]\d{0,9}$/.test(String(httpsCallback.windowId)) || this._isGoogleLoopbackOAuth || this._isSignalCMailOAuth) {
				throw new Error(nls.localize('oauth.httpsCallbackUnavailable', "This connection cannot use the Signal-C HTTPS OAuth callback."));
			}
			flows = [{
				label: nls.localize('oauth.httpsCallback', "Browser Sign-in"),
				handler: (scopes, progress, token) => this._createWithUrlHandler(scopes, progress, token, httpsCallback),
			}];
		}
		let token: IAuthorizationTokenResponse | undefined;
		for (let i = 0; i < flows.length; i++) {
			const { handler } = flows[i];
			try {
				token = await this._extHostProgress.withProgressFromSource(
					{ label: this.label, id: this.id },
					{
						location: ProgressLocation.Notification,
						title: nls.localize('authenticatingTo', "Authenticating to '{0}'", this.label),
						cancellable: true
					},
					(progress, token) => handler(scopes, progress, token),
					NotificationTelemetryId.AuthenticationSignIn);
				if (token) {
					break;
				}
			} catch (err) {
				const nextMode = flows[i + 1]?.label;
				if (!nextMode) {
					throw err;
				}
				const message = isCancellationError(err)
					? nls.localize('userCanceledContinue', "Having trouble authenticating to '{0}'? Would you like to try a different way? ({1})", this.label, nextMode)
					: nls.localize('continueWith', "You have not yet finished authenticating to '{0}'. Would you like to try a different way? ({1})", this.label, nextMode);

				const result = await this._proxy.$showContinueNotification(message);
				if (!result) {
					throw new CancellationError();
				}
				this._logger.error(`Failed to create token via flow '${nextMode}': ${err}`);
			}
		}
		if (!token) {
			throw new Error('Failed to create authentication token');
		}
		const grantedScopes = this._getGrantedScopes(scopes, token);
		const acceptPartialScopes = options[MCP_ACCEPT_PARTIAL_OAUTH_SCOPES] === true;
		if (!this._isUsableScopeGrant(scopes, grantedScopes, acceptPartialScopes)) {
			const missing = scopes.filter(scope => !grantedScopes.includes(scope));
			throw new Error(nls.localize('oauth.grantMissingScopes', "The provider did not grant the required permissions ({0}).", missing.join(', ')));
		}

		// Store session for later retrieval
		const stored: IAuthorizationToken = { ...token, scope: grantedScopes.join(' '), requested_scope: scopes.join(' '), created_at: Date.now() };
		const persisted = this._tokenStore.update({ added: [stored], removed: [] });
		if (this._isSignalCMailOAuth) {
			await persisted;
		}
		const session = this._tokenStore.sessions.find(t => t.accessToken === token.access_token)!;
		this._logger.info(`Created ${token.refresh_token ? 'refreshable' : 'non-refreshable'} session for granted scopes: ${stored.scope}${token.expires_in ? ` that expires in ${token.expires_in} seconds` : ''}`);
		return session;
	}

	async removeSession(sessionId: string): Promise<void> {
		this._logger.info(`Removing session with id: ${sessionId}`);
		const session = this._tokenStore.sessions.find(session => session.id === sessionId);
		if (!session) {
			this._logger.error(`Session with id ${sessionId} not found`);
			return;
		}
		const token = this._tokenStore.tokens.find(token => token.access_token === session.accessToken);
		if (!token) {
			this._logger.error(`Failed to retrieve token for removed session: ${session.id}`);
			return;
		}
		this._tokenStore.update({ added: [], removed: [token] });
		this._logger.info(`Removed token for session: ${session.id} with scopes: ${session.scopes.join(' ')}`);
	}

	dispose(): void {
		this._disposable.dispose();
	}

	private async _createWithUrlHandler(scopes: string[], progress: vscode.Progress<IProgressStep>, token: vscode.CancellationToken, httpsCallback?: ISignalCOAuthCallback): Promise<IAuthorizationTokenResponse> {
		if (!this._serverMetadata.authorization_endpoint) {
			throw new Error('Authorization Endpoint required');
		}
		if (!this._serverMetadata.token_endpoint) {
			throw new Error('Token endpoint not available in server metadata');
		}
		if (httpsCallback && (new URL(this._serverMetadata.authorization_endpoint).protocol !== 'https:' || new URL(this._serverMetadata.token_endpoint).protocol !== 'https:')) {
			throw new Error(nls.localize('oauth.httpsCallbackRequiresHttps', "Browser sign-in requires HTTPS authorization and token endpoints."));
		}

		// Generate PKCE code verifier (random string) and code challenge (SHA-256 hash of verifier)
		const codeVerifier = this.generateRandomString(64);
		const codeChallenge = await this.generateCodeChallenge(codeVerifier);

		// Generate a random state value to prevent CSRF
		const nonce = this.generateRandomString(32);
		const callbackUri = URI.from({
			scheme: this._initData.environment.appUriScheme,
			authority: 'dynamicauthprovider',
			path: httpsCallback ? '/https/authorize' : `/${this.authorizationServer.authority}/authorize`,
			query: httpsCallback ? `nonce=${nonce}&windowId=${httpsCallback.windowId}` : `nonce=${nonce}`,
		});
		const state = httpsCallback ? `${nonce}.${httpsCallback.windowId}` : (await this._extHostUrls.createAppUri(callbackUri)).toString();

		// Prepare the authorization request URL
		const authorizationUrl = new URL(this._serverMetadata.authorization_endpoint);
		authorizationUrl.searchParams.append('client_id', this._clientId);
		authorizationUrl.searchParams.append('response_type', 'code');
		authorizationUrl.searchParams.append('state', state);
		authorizationUrl.searchParams.append('code_challenge', codeChallenge);
		authorizationUrl.searchParams.append('code_challenge_method', 'S256');
		this._configureAuthorizationUrl(authorizationUrl);
		const scopeString = scopes.join(' ');
		const requestScopeString = this._getRequestScopes(scopes).join(' ');
		if (requestScopeString) {
			// If non-empty scopes are provided, include scope parameter in the request
			authorizationUrl.searchParams.append('scope', requestScopeString);
		}
		const resource = this._resourceIndicator;
		if (resource) {
			// If a resource is specified, include it in the request
			authorizationUrl.searchParams.append('resource', resource);
		}

		const redirectUri = httpsCallback?.redirectUri ?? 'https://vscode.dev/redirect';
		authorizationUrl.searchParams.append('redirect_uri', redirectUri);

		const callback = new CancellationTokenSource(token);
		try {
			const authorizationCode = this.waitForAuthorizationCode(callbackUri, callback.token, !!httpsCallback);
			// The listener can be canceled while the external opener is still pending.
			void authorizationCode.catch(() => { });
			this._logger.info(`Opening authorization URL for scopes: ${scopeString}`);
			const opened = await this._extHostWindow.openUri(authorizationUrl.toString(), {});
			if (!opened) {
				throw new CancellationError();
			}
			progress.report({
				message: nls.localize('completeAuth', "Complete the authentication in the browser window that has opened."),
			});
			const { code } = await authorizationCode;
			this._logger.info(`Authorization code received for scopes: ${scopeString}`);
			return await this.exchangeCodeForToken(code, codeVerifier, redirectUri);
		} finally {
			callback.dispose(true);
		}
	}

	protected generateRandomString(length: number): string {
		const array = new Uint8Array(length);
		crypto.getRandomValues(array);
		return Array.from(array)
			.map(b => b.toString(16).padStart(2, '0'))
			.join('')
			.substring(0, length);
	}

	protected async generateCodeChallenge(codeVerifier: string): Promise<string> {
		const encoder = new TextEncoder();
		const data = encoder.encode(codeVerifier);
		const digest = await crypto.subtle.digest('SHA-256', data);

		// Base64url encode the digest
		return encodeBase64(VSBuffer.wrap(new Uint8Array(digest)), false, false)
			.replace(/\+/g, '-')
			.replace(/\//g, '_')
			.replace(/=+$/, '');
	}

	private async waitForAuthorizationCode(expectedState: URI, token: vscode.CancellationToken, signalCCallback: boolean): Promise<{ code: string }> {
		const result = await this._proxy.$waitForUriHandler(expectedState, token);
		if (signalCCallback) {
			const params = new URLSearchParams(result.query);
			const errors = params.getAll('error');
			const codes = params.getAll('code64');
			if (errors.length === 1 && errors[0] === 'access_denied' && !codes.length) {
				throw new CancellationError();
			}
			if (!errors.length && codes.length === 1 && /^[A-Za-z0-9_-]{2,10923}$/.test(codes[0]) && codes[0].length % 4 !== 1) {
				return { code: decodeBase64(codes[0]).toString() };
			}
			throw new Error(nls.localize('oauth.httpsCallbackMissingCode', "The provider did not return an authorization code to Signal-C."));
		}
		// Extract the code parameter directly from the query string. NOTE, URLSearchParams does not work here because
		// it will decode the query string and we need to keep it encoded.
		const codeMatch = /(?:^|[?&])code=([^&]+)/.exec(result.query || '');
		if (!codeMatch || codeMatch.length < 2) {
			// No code parameter found in the query string
			throw new Error('Authentication failed: No authorization code received');
		}
		return { code: codeMatch[1] };
	}

	protected _configureAuthorizationUrl(url: URL): void {
		if (this._isGoogleLoopbackOAuth) {
			url.searchParams.set('access_type', 'offline');
			// A previous online-only grant may otherwise keep Google from issuing a refresh token.
			url.searchParams.set('prompt', 'consent');
		} else if (this._isSignalCMailOAuth) {
			url.searchParams.set('prompt', 'select_account');
		}
	}

	private _tokenRequestInit(tokenEndpoint: string, body: URLSearchParams): RequestInit {
		const headers: Record<string, string> = {
			'Content-Type': 'application/x-www-form-urlencoded',
			'Accept': 'application/json',
		};
		const registeredMethod = this._tokenEndpointAuthMethod;
		if (registeredMethod !== undefined && registeredMethod !== 'none' && registeredMethod !== 'client_secret_basic' && registeredMethod !== 'client_secret_post') {
			throw new Error(nls.localize('unsupportedTokenEndpointAuthMethod', "The OAuth provider registered this client with an unsupported token authentication method."));
		}
		const clientSecret = registeredMethod === 'none' ? undefined : this._clientSecret;
		if ((registeredMethod === 'client_secret_basic' || registeredMethod === 'client_secret_post') && !clientSecret) {
			throw new Error(nls.localize('missingTokenEndpointClientSecret', "The registered OAuth client requires a client secret."));
		}
		if (clientSecret && new URL(tokenEndpoint).protocol !== 'https:') {
			throw new Error('OAuth client secret authentication requires an HTTPS token endpoint');
		}
		// For older or manually configured registrations, retain the server-wide fallback.
		const methods = this._serverMetadata.token_endpoint_auth_methods_supported;
		const useBasic = registeredMethod === 'client_secret_basic'
			|| (registeredMethod === undefined && Array.isArray(methods) && methods.includes('client_secret_basic') && !methods.includes('client_secret_post'));
		if (clientSecret && useBasic) {
			// OAuth form-encodes each credential before joining them for HTTP Basic (RFC 6749 §2.3.1).
			const formEncode = (value: string) => new URLSearchParams([['', value]]).toString().slice(1);
			headers.Authorization = `Basic ${encodeBase64(VSBuffer.fromString(`${formEncode(this._clientId)}:${formEncode(clientSecret)}`))}`;
			return { method: 'POST', headers, body: body.toString(), redirect: 'error' };
		}
		body.append('client_id', this._clientId);
		if (clientSecret) {
			body.append('client_secret', clientSecret);
		}
		return { method: 'POST', headers, body: body.toString(), ...(clientSecret || this._isSignalCMailOAuth || this._isGoogleLoopbackOAuth ? { redirect: 'error' } : {}) };
	}

	protected async exchangeCodeForToken(code: string, codeVerifier: string, redirectUri: string): Promise<IAuthorizationTokenResponse> {
		if (!this._serverMetadata.token_endpoint) {
			throw new Error('Token endpoint not available in server metadata');
		}

		const tokenRequest = new URLSearchParams();
		tokenRequest.append('grant_type', 'authorization_code');
		tokenRequest.append('code', code);
		tokenRequest.append('redirect_uri', redirectUri);
		tokenRequest.append('code_verifier', codeVerifier);

		// Add resource indicator if available (RFC 8707)
		const resource = this._resourceIndicator;
		if (resource) {
			tokenRequest.append('resource', resource);
		}

		this._logger.info('Exchanging authorization code for token...');
		this._logger.trace(`Url: ${this._serverMetadata.token_endpoint}`);
		const request = this._tokenRequestInit(this._serverMetadata.token_endpoint, tokenRequest);
		if (redirectUri === this._signalCOAuthCallback?.redirectUri) {
			request.redirect = 'error';
		}
		let response: Response;
		try {
			response = await this._fetch(this._serverMetadata.token_endpoint, request);
		} catch (err) {
			this._logger.error(`Failed to exchange authorization code for token: ${err}`);
			throw new Error(`Failed to exchange authorization code for token: ${err}`);
		}

		if (!response.ok) {
			const text = await response.text();
			throw new Error(`Token exchange failed: ${response.status} ${response.statusText} - ${text}`);
		}

		const result = await response.json();
		if (isAuthorizationTokenResponse(result)) {
			this._logger.info(`Successfully exchanged authorization code for token.`);
			return result;
		} else if (isAuthorizationErrorResponse(result) && result.error === AuthorizationErrorType.InvalidClient) {
			this._logger.warn(`Client ID (${this._clientId}) was invalid, generated a new one.`);
			await this._generateNewClientId();
			throw new Error(`Client ID was invalid, generated a new one. Please try again.`);
		}
		throw new Error(`Invalid authorization token response: ${JSON.stringify(result)}`);
	}

	protected async exchangeRefreshTokenForToken(refreshToken: string, scopes: readonly string[], allowClientRegistration: boolean): Promise<IAuthorizationToken> {
		if (!this._serverMetadata.token_endpoint) {
			throw new Error('Token endpoint not available in server metadata');
		}

		const tokenRequest = new URLSearchParams();
		tokenRequest.append('grant_type', 'refresh_token');
		tokenRequest.append('refresh_token', refreshToken);

		if (this._isMicrosoftAuth) {
			const requestScopeString = this._getRequestScopes(scopes).join(' ');
			if (requestScopeString) {
				tokenRequest.append('scope', requestScopeString);
			}
		}

		// Add resource indicator if available (RFC 8707)
		const resource = this._resourceIndicator;
		if (resource) {
			tokenRequest.append('resource', resource);
		}

		const response = await this._fetch(this._serverMetadata.token_endpoint, this._tokenRequestInit(this._serverMetadata.token_endpoint, tokenRequest));

		const result = await response.json();
		if (response.ok && isAuthorizationTokenResponse(result)) {
			return {
				...result,
				// OAuth servers may issue a replacement refresh token or leave the current one in effect.
				refresh_token: result.refresh_token ?? refreshToken,
				created_at: Date.now(),
			};
		} else if (isAuthorizationErrorResponse(result) && result.error === AuthorizationErrorType.InvalidClient) {
			if (!allowClientRegistration) {
				this._logger.warn(`Client ID (${this._clientId}) was invalid while silently refreshing the token.`);
				throw new Error(`Client ID was invalid while silently refreshing the token.`);
			}
			this._logger.warn(`Client ID (${this._clientId}) was invalid, generated a new one.`);
			await this._generateNewClientId();
			throw new Error(`Client ID was invalid, generated a new one. Please try again.`);
		}
		throw new Error(`Invalid authorization token response: ${JSON.stringify(result)}`);
	}

	private _isUsableScopeGrant(requested: readonly string[], granted: readonly string[], acceptPartial: boolean): boolean {
		return requested.every(scope => granted.includes(scope))
			|| (acceptPartial && requested.some(scope => !identityOAuthScopes.has(scope) && granted.includes(scope)));
	}

	private _getGrantedScopes(requested: readonly string[], token: IAuthorizationTokenResponse): string[] {
		// OAuth permits omitting scope from a successful response when it is unchanged.
		if (token.scope === undefined) {
			return [...requested];
		}
		const granted = token.scope.split(' ').filter(Boolean);
		if (this._isGoogleLoopbackOAuth) {
			for (const scope of ['email', 'profile']) {
				const index = granted.indexOf(`https://www.googleapis.com/auth/userinfo.${scope}`);
				if (index !== -1 && requested.includes(scope)) {
					granted[index] = scope;
				}
			}
		}
		if (this._isSignalCMailOAuth && this._isMicrosoftAuth) {
			// Graph may use unqualified permission names and reports access-token scopes
			// separately from the identity and refresh tokens issued by the same exchange.
			const graph = signalCMailOAuth.outlook.resource;
			for (const scope of [signalCMailOAuth.outlook.scope, signalCMailOAuth.outlook.sendScope, signalCMailOAuth.outlook.calendarScope]) {
				const index = granted.findIndex(value => [scope.toLowerCase(), scope.slice(graph.length).toLowerCase()].includes(value.toLowerCase()));
				if (index !== -1) {
					granted[index] = scope;
				}
			}
			if (token.id_token) {
				granted.push(...requested.filter(scope => scope === 'openid' || scope === 'profile'));
			}
			if (token.refresh_token && requested.includes('offline_access')) {
				granted.push('offline_access');
			}
		}
		return [...new Set(granted)];
	}

	protected async _generateNewClientId(): Promise<void> {
		try {
			const registration = await fetchDynamicRegistration(this._serverMetadata, this._initData.environment.appName, this._resourceMetadata?.scopes_supported, this._initData.environment.appUriScheme);
			this._clientId = registration.client_id;
			this._clientSecret = registration.client_secret;
			this._tokenEndpointAuthMethod = registration.token_endpoint_auth_method;
			this._onDidChangeClientId.fire();
		} catch (err) {
			// When DCR fails, try to prompt the user for a client ID and client secret
			this._logger.info(`Dynamic registration failed for ${this.authorizationServer.toString()}: ${err}. Prompting user for client ID and client secret.`);

			try {
				const clientDetails = await this._proxy.$promptForClientRegistration(this.authorizationServer.toString());
				if (!clientDetails) {
					throw new Error('User did not provide client details');
				}
				this._clientId = clientDetails.clientId;
				this._clientSecret = clientDetails.clientSecret;
				this._tokenEndpointAuthMethod = undefined;
				this._logger.info(`User provided client ID for ${this.authorizationServer.toString()}`);
				if (clientDetails.clientSecret) {
					this._logger.info(`User provided client secret for ${this.authorizationServer.toString()}`);
				} else {
					this._logger.info(`User did not provide client secret for ${this.authorizationServer.toString()} (optional)`);
				}

				this._onDidChangeClientId.fire();
			} catch (promptErr) {
				this._logger.error(`Failed to fetch new client ID and user did not provide one: ${err}`);
				throw new Error(`Failed to fetch new client ID and user did not provide one: ${err}`);
			}
		}
	}
}

const identityOAuthScopes = new Set(['openid', 'profile', 'email', 'offline_access']);

export type IAuthorizationToken = IAuthorizationTokenResponse & {
	/**
	 * The time when the token was created, in milliseconds since the epoch.
	 */
	created_at: number;
	/** The lookup scopes used at authorization time, which can differ from the actual OAuth grant. */
	requested_scope?: string;
};

export class TokenStore implements Disposable {
	private readonly _tokensObservable: ISettableObservable<IAuthorizationToken[]>;
	private readonly _sessionsObservable: IObservable<vscode.AuthenticationSession[]>;

	private readonly _onDidChangeSessions = new Emitter<vscode.AuthenticationProviderAuthenticationSessionsChangeEvent>();
	readonly onDidChangeSessions = this._onDidChangeSessions.event;

	private readonly _disposable: DisposableStore;

	constructor(
		private readonly _persistence: { onDidChange: Event<IAuthorizationToken[]>; set: (tokens: IAuthorizationToken[]) => void | Promise<void> },
		initialTokens: IAuthorizationToken[],
		private readonly _logger: ILogger
	) {
		this._disposable = new DisposableStore();
		this._tokensObservable = observableValue<IAuthorizationToken[]>('tokens', initialTokens);
		this._sessionsObservable = derivedOpts(
			{ equalsFn: (a, b) => arraysEqual(a, b, (a, b) => a.accessToken === b.accessToken && scopesMatch(a.scopes, b.scopes)) },
			(reader) => this._tokensObservable.read(reader).map(t => this.getSessionFromToken(t))
		);
		this._disposable.add(this._registerChangeEventAutorun());
		this._disposable.add(this._persistence.onDidChange((tokens) => this._tokensObservable.set(tokens, undefined)));
	}

	get tokens(): IAuthorizationToken[] {
		return this._tokensObservable.get();
	}

	get sessions(): vscode.AuthenticationSession[] {
		return this._sessionsObservable.get();
	}

	dispose() {
		this._disposable.dispose();
	}

	update({ added, removed }: { added: IAuthorizationToken[]; removed: IAuthorizationToken[] }): void | Promise<void> {
		this._logger.trace(`Updating tokens: added ${added.length}, removed ${removed.length}`);
		const currentTokens = [...this._tokensObservable.get()];
		for (const token of removed) {
			const index = currentTokens.findIndex(t => t.access_token === token.access_token);
			if (index !== -1) {
				currentTokens.splice(index, 1);
			}
		}
		for (const token of added) {
			const index = currentTokens.findIndex(t => t.access_token === token.access_token);
			if (index === -1) {
				currentTokens.push(token);
			} else {
				currentTokens[index] = token;
			}
		}
		let persisted: void | Promise<void> = undefined;
		if (added.length || removed.length) {
			this._tokensObservable.set(currentTokens, undefined);
			persisted = this._persistence.set(currentTokens);
		}
		this._logger.trace(`Tokens updated: ${currentTokens.length} tokens stored.`);
		return persisted;
	}

	private _registerChangeEventAutorun(): IDisposable {
		let previousSessions: vscode.AuthenticationSession[] = [];
		return autorun((reader) => {
			this._logger.trace('Checking for session changes...');
			const currentSessions = this._sessionsObservable.read(reader);
			if (previousSessions === currentSessions) {
				this._logger.trace('No session changes detected.');
				return;
			}

			if (!currentSessions || currentSessions.length === 0) {
				// If currentSessions is undefined, all previous sessions are considered removed
				this._logger.trace('All sessions removed.');
				if (previousSessions.length > 0) {
					this._onDidChangeSessions.fire({
						added: [],
						removed: previousSessions,
						changed: []
					});
					previousSessions = [];
				}
				return;
			}

			const added: vscode.AuthenticationSession[] = [];
			const removed: vscode.AuthenticationSession[] = [];
			const changed: vscode.AuthenticationSession[] = [];

			// The provider can reuse a bearer token while changing its granted scopes.
			for (const current of currentSessions) {
				const previous = previousSessions.find(prev => prev.accessToken === current.accessToken);
				if (!previous) {
					added.push(current);
				} else if (!scopesMatch(previous.scopes, current.scopes)) {
					changed.push(current);
				}
			}

			// Find removed sessions
			for (const prev of previousSessions) {
				const exists = currentSessions.some(current => current.accessToken === prev.accessToken);
				if (!exists) {
					removed.push(prev);
				}
			}

			// Fire the event if there are any changes
			if (added.length > 0 || removed.length > 0 || changed.length > 0) {
				this._logger.trace(`Sessions changed: added ${added.length}, removed ${removed.length}`);
				this._onDidChangeSessions.fire({ added, removed, changed });
			}

			// Update previous sessions reference
			previousSessions = currentSessions;
		});
	}

	getSessionFromToken(token: IAuthorizationTokenResponse): vscode.AuthenticationSession {
		let claims: IAuthorizationJWTClaims | undefined;
		if (token.id_token) {
			try {
				claims = getClaimsFromJWT(token.id_token);
			} catch (e) {
				// log
			}
		}
		if (!claims) {
			try {
				claims = getClaimsFromJWT(token.access_token);
			} catch (e) {
				// log
			}
		}
		// An explicit empty scope is authoritative. Fall back to JWT claims only for older records that omitted scope.
		const scopes = token.scope !== undefined
			? (token.scope ? token.scope.split(' ') : [])
			: (claims?.scope ? claims.scope.split(' ') : []);
		return {
			id: stringHash(token.access_token, 0).toString(),
			accessToken: token.access_token,
			account: {
				id: claims?.sub || 'unknown',
				// TODO: Don't say MCP...
				label: claims?.preferred_username || claims?.name || claims?.email || 'MCP',
			},
			scopes: scopes,
			idToken: token.id_token
		};
	}
}
