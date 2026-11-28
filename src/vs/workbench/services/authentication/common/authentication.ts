/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/
import { encodeBase64, VSBuffer } from '../../../../base/common/buffer.js';
import { Event } from '../../../../base/common/event.js';
import { IDisposable } from '../../../../base/common/lifecycle.js';
import { IAuthenticationChallenge, IAuthorizationProtectedResourceMetadata, IAuthorizationServerMetadata } from '../../../../base/common/oauth.js';
import { filter } from '../../../../base/common/objects.js';
import { compare } from '../../../../base/common/strings.js';
import { URI } from '../../../../base/common/uri.js';
import { createDecorator } from '../../../../platform/instantiation/common/instantiation.js';

/**
 * Use this if you don't want the onDidChangeSessions event to fire in the extension host
 */
export const INTERNAL_AUTH_PROVIDER_PREFIX = '__';

/** Identifies an OAuth bearer token in a recovery request without including the credential. */
export async function getOAuthTokenFingerprint(accessToken: string): Promise<string> {
	const digest = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(accessToken));
	return encodeBase64(VSBuffer.wrap(new Uint8Array(digest)), false, true);
}

export function isGoogleWorkspaceOAuth(authorizationServer: URI, metadata: IAuthorizationServerMetadata, resource: URI | undefined): boolean {
	return authorizationServer.scheme === 'https'
		&& authorizationServer.authority === 'accounts.google.com'
		&& metadata.authorization_endpoint === 'https://accounts.google.com/o/oauth2/v2/auth'
		&& metadata.token_endpoint === 'https://oauth2.googleapis.com/token'
		&& resource?.scheme === 'https'
		&& /^[a-z0-9]+\.googleapis\.com$/.test(resource.authority)
		&& /^\/mcp(?:\/v1)?\/?$/.test(resource.path);
}

/** Signal-C mailbox grants use registered desktop callbacks and their own storage namespace. */
export const signalCMailOAuthPrefix = 'signal-c-mail:';

export const signalCMailOAuth = {
	gmail: {
		issuer: 'https://accounts.google.com',
		authorize: 'https://accounts.google.com/o/oauth2/v2/auth',
		token: 'https://oauth2.googleapis.com/token',
		resource: 'https://gmail.googleapis.com/',
		scope: 'https://www.googleapis.com/auth/gmail.readonly',
		sendScope: 'https://www.googleapis.com/auth/gmail.send',
	},
	outlook: {
		issuer: 'https://login.microsoftonline.com/common/v2.0',
		authorize: 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize',
		token: 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
		resource: 'https://graph.microsoft.com/',
		scope: 'https://graph.microsoft.com/Mail.Read',
		sendScope: 'https://graph.microsoft.com/Mail.Send',
		calendarScope: 'https://graph.microsoft.com/Calendars.Read',
	},
} as const;

export type SignalCMailProvider = keyof typeof signalCMailOAuth;

export function isSignalCMailOAuth(providerId: string, authorizationServer: URI, metadata: IAuthorizationServerMetadata, resource: string | undefined): boolean {
	const provider: SignalCMailProvider = providerId.startsWith(`${signalCMailOAuthPrefix}gmail:`) ? 'gmail' : 'outlook';
	const config = signalCMailOAuth[provider];
	return providerId.startsWith(`${signalCMailOAuthPrefix}${provider}:`)
		&& authorizationServer.toString(true) === URI.parse(config.issuer).toString(true)
		&& metadata.authorization_endpoint === config.authorize
		&& metadata.token_endpoint === config.token
		&& resource === config.resource;
}

/**
 * Setting that controls whether the profile image (avatar) of a signed-in account
 * is shown in account related UI.
 */
export const ACCOUNTS_AVATAR_SETTING = 'workbench.accounts.showAvatar';

export interface AuthenticationSessionAccount {
	label: string;
	id: string;
	/**
	 * An optional icon for the account. This is typically a URI to a profile image/avatar.
	 */
	icon?: URI;
}

export interface AuthenticationSession {
	id: string;
	accessToken: string;
	account: AuthenticationSessionAccount;
	scopes: ReadonlyArray<string>;
	idToken?: string;
	expiresAfter?: number;
	/** The issuing OAuth server, not a REST API endpoint or resource audience. */
	readonly authorizationServer?: URI;
}

export interface AuthenticationSessionsChangeEvent {
	added: ReadonlyArray<AuthenticationSession> | undefined;
	removed: ReadonlyArray<AuthenticationSession> | undefined;
	changed: ReadonlyArray<AuthenticationSession> | undefined;
}

export interface AuthenticationProviderInformation {
	id: string;
	label: string;
	authorizationServerGlobs?: ReadonlyArray<string>;
}

/**
 * Options for creating an authentication session via the service.
 */
export interface IAuthenticationCreateSessionOptions extends IAuthenticationProviderSessionOptions {
	activateImmediate?: boolean;
}

export interface IAuthenticationWwwAuthenticateRequest {
	/**
	 * The raw WWW-Authenticate header value that triggered this challenge.
	 * This will be parsed by the authentication provider to extract the necessary
	 * challenge information.
	 */
	readonly wwwAuthenticate: string;

	/**
	 * Optional scopes for the session. If not provided, the authentication provider
	 * may use default scopes or extract them from the challenge.
	 */
	readonly fallbackScopes?: readonly string[];
}

export function isAuthenticationWwwAuthenticateRequest(obj: unknown): obj is IAuthenticationWwwAuthenticateRequest {
	return typeof obj === 'object'
		&& obj !== null
		&& 'wwwAuthenticate' in obj
		&& (typeof obj.wwwAuthenticate === 'string');
}

/**
 * Represents constraints for authentication, including challenges and optional scopes.
 * This is used when creating or retrieving sessions that must satisfy specific authentication
 * requirements from WWW-Authenticate headers.
 */
export interface IAuthenticationConstraint {
	/**
	 * Array of authentication challenges parsed from WWW-Authenticate headers.
	 */
	readonly challenges: readonly IAuthenticationChallenge[];

	/**
	 * Optional scopes for the session. If not provided, the authentication provider
	 * may extract scopes from the challenges or use default scopes.
	 */
	readonly fallbackScopes?: readonly string[];
}

/**
 * Options for getting authentication sessions via the service.
 */
export type IAuthenticationGetSessionsOptions = IAuthenticationProviderSessionOptions;

export interface AllowedExtension {
	id: string;
	name: string;
	/**
	 * If true or undefined, the extension is allowed to use the account
	 * If false, the extension is not allowed to use the account
	 * TODO: undefined shouldn't be a valid value, but it is for now
	 */
	allowed?: boolean;
	lastUsed?: number;
	// If true, this comes from the product.json
	trusted?: boolean;
}

export interface IAuthenticationProviderHostDelegate {
	/** Priority for this delegate, delegates are tested in descending priority order */
	readonly priority: number;
	create(authorizationServer: URI, serverMetadata: IAuthorizationServerMetadata, resource: IAuthorizationProtectedResourceMetadata | undefined, clientId?: string, clientSecret?: string, providerId?: string): Promise<string>;
	/**
	 * Creates an XAA (enterprise-managed, ID-JAG) authentication provider for the given SSO issuer.
	 * The returned string is the provider id.
	 */
	createXaa?(issuer: URI): Promise<string>;
}

const dynamicProviderClientSeparator = ' client:';

export function getDynamicAuthenticationProviderId(authorizationServer: URI, resource: IAuthorizationProtectedResourceMetadata | undefined, clientId?: string): string {
	const providerId = resource ? `${authorizationServer.toString(true)} ${resource.resource}` : authorizationServer.toString(true);
	return clientId ? `${providerId}${dynamicProviderClientSeparator}${encodeURIComponent(clientId)}` : providerId;
}

export function isClientScopedDynamicAuthenticationProviderId(providerId: string): boolean {
	return providerId.includes(dynamicProviderClientSeparator);
}

export const IAuthenticationService = createDecorator<IAuthenticationService>('IAuthenticationService');

export interface IAuthenticationService {
	readonly _serviceBrand: undefined;

	/**
	 * Fires when an authentication provider has been registered
	 */
	readonly onDidRegisterAuthenticationProvider: Event<AuthenticationProviderInformation>;
	/**
	 * Fires when an authentication provider has been unregistered
	 */
	readonly onDidUnregisterAuthenticationProvider: Event<AuthenticationProviderInformation>;

	/**
	 * Fires when the list of sessions for a provider has been added, removed or changed
	 */
	readonly onDidChangeSessions: Event<{ providerId: string; label: string; event: AuthenticationSessionsChangeEvent }>;

	/**
	 * Fires when the list of declaredProviders has changed
	 */
	readonly onDidChangeDeclaredProviders: Event<void>;

	/**
	 * All providers that have been statically declared by extensions. These may not actually be registered or active yet.
	 */
	readonly declaredProviders: AuthenticationProviderInformation[];

	/**
	 * Registers that an extension has declared an authentication provider in their package.json
	 * @param provider The provider information to register
	 */
	registerDeclaredAuthenticationProvider(provider: AuthenticationProviderInformation): void;

	/**
	 * Unregisters a declared authentication provider
	 * @param id The id of the provider to unregister
	 */
	unregisterDeclaredAuthenticationProvider(id: string): void;

	/**
	 * Checks if an authentication provider has been registered
	 * @param id The id of the provider to check
	 */
	isAuthenticationProviderRegistered(id: string): boolean;

	/**
	 * Checks if an authentication provider is dynamic
	 * @param id The id of the provider to check
	 */
	isDynamicAuthenticationProvider(id: string): boolean;

	/**
	 * Registers an authentication provider
	 * @param id The id of the provider
	 * @param provider The implementation of the provider
	 */
	registerAuthenticationProvider(id: string, provider: IAuthenticationProvider): void;

	/**
	 * Unregisters an authentication provider
	 * @param id The id of the provider to unregister
	 */
	unregisterAuthenticationProvider(id: string): void;

	/**
	 * Gets the provider ids of all registered authentication providers
	 */
	getProviderIds(): string[];

	/**
	 * Gets the provider with the given id.
	 * @param id The id of the provider to get
	 * @throws if the provider is not registered
	 */
	getProvider(id: string): IAuthenticationProvider;

	/**
	 * Gets all accounts that are currently logged in across all sessions
	 * @param id The id of the provider to ask for accounts
	 * @returns A promise that resolves to an array of accounts
	 */
	getAccounts(id: string): Promise<ReadonlyArray<AuthenticationSessionAccount>>;

	/**
	 * Gets all sessions that satisfy the given scopes from the provider with the given id
	 * @param id The id of the provider to ask for a session
	 * @param scopes The scopes for the session
	 * @param options Additional options for getting sessions
	 * @param activateImmediate If true, the provider should activate immediately if it is not already
	 */
	getSessions(id: string, scopeListOrRequest?: ReadonlyArray<string> | IAuthenticationWwwAuthenticateRequest, options?: IAuthenticationGetSessionsOptions, activateImmediate?: boolean): Promise<ReadonlyArray<AuthenticationSession>>;

	/**
	 * Creates an AuthenticationSession with the given provider and scopes
	 * @param providerId The id of the provider
	 * @param scopes The scopes to request
	 * @param options Additional options for creating the session
	 */
	createSession(providerId: string, scopeListOrRequest: ReadonlyArray<string> | IAuthenticationWwwAuthenticateRequest, options?: IAuthenticationCreateSessionOptions): Promise<AuthenticationSession>;

	/**
	 * Removes the session with the given id from the provider with the given id
	 * @param providerId The id of the provider
	 * @param sessionId The id of the session to remove
	 */
	removeSession(providerId: string, sessionId: string): Promise<void>;

	/**
	 * Gets an unscoped provider id for a specified authorization server. Callers
	 * with an explicitly configured OAuth client can look up its provider by ID.
	 * @param authorizationServer The authorization server url that this provider is responsible for
	 * @param resourceServer The resource server URI that should match the provider's resourceServer (if defined)
	 */
	getOrActivateProviderIdForServer(authorizationServer: URI, resourceServer?: URI): Promise<string | undefined>;

	/**
	 * Allows the ability register a delegate that will be used to start authentication providers
	 * @param delegate The delegate to register
	 */
	registerAuthenticationProviderHostDelegate(delegate: IAuthenticationProviderHostDelegate): IDisposable;

	/**
	 * Creates a dynamic authentication provider. An explicit provider ID scopes its stored registration and tokens.
	 */
	createDynamicAuthenticationProvider(authorizationServer: URI, serverMetadata: IAuthorizationServerMetadata, resourceMetadata: IAuthorizationProtectedResourceMetadata | undefined, clientId?: string, clientSecret?: string, providerId?: string): Promise<IAuthenticationProvider | undefined>;

	/**
	 * Gets or creates a built-in XAA (enterprise-managed, ID-JAG) authentication provider for the given
	 * SSO issuer. Subsequent calls with the same issuer return the existing provider. The returned id
	 * can be used with {@link getSessions}/{@link createSession} just like any other provider.
	 *
	 * @param issuer The OAuth/OIDC issuer URL (typically read from `mcp.enterpriseManagedAuth.idp`).
	 */
	createOrGetXaaProvider(issuer: URI): Promise<string | undefined>;
}

export function isAuthenticationSession(thing: unknown): thing is AuthenticationSession {
	if (typeof thing !== 'object' || !thing) {
		return false;
	}
	const maybe = thing as AuthenticationSession;
	if (typeof maybe.id !== 'string') {
		return false;
	}
	if (typeof maybe.accessToken !== 'string') {
		return false;
	}
	if (typeof maybe.account !== 'object' || !maybe.account) {
		return false;
	}
	if (typeof maybe.account.label !== 'string') {
		return false;
	}
	if (typeof maybe.account.id !== 'string') {
		return false;
	}
	if (!Array.isArray(maybe.scopes)) {
		return false;
	}
	if (maybe.idToken && typeof maybe.idToken !== 'string') {
		return false;
	}
	return true;
}

// TODO: Move this into MainThreadAuthentication
export const IAuthenticationExtensionsService = createDecorator<IAuthenticationExtensionsService>('IAuthenticationExtensionsService');
export interface IAuthenticationExtensionsService {
	readonly _serviceBrand: undefined;

	/**
	 * Fires when an account preference for a specific provider has changed for the specified extensions. Does not fire when:
	 * * An account preference is removed
	 * * A session preference is changed (because it's deprecated)
	 * * A session preference is removed (because it's deprecated)
	 */
	readonly onDidChangeAccountPreference: Event<{ extensionIds: string[]; providerId: string }>;
	/**
	 * Returns the accountName (also known as account.label) to pair with `IAuthenticationAccessService` to get the account preference
	 * @param providerId The authentication provider id
	 * @param extensionId The extension id to get the preference for
	 * @returns The accountName of the preference, or undefined if there is no preference set
	 */
	getAccountPreference(extensionId: string, providerId: string): string | undefined;
	/**
	 * Sets the account preference for the given provider and extension
	 * @param providerId The authentication provider id
	 * @param extensionId The extension id to set the preference for
	 * @param account The account to set the preference to
	 */
	updateAccountPreference(extensionId: string, providerId: string, account: AuthenticationSessionAccount): void;
	/**
	 * Removes the account preference for the given provider and extension
	 * @param providerId The authentication provider id
	 * @param extensionId The extension id to remove the preference for
	 */
	removeAccountPreference(extensionId: string, providerId: string): void;
	/**
	 * @deprecated Sets the session preference for the given provider and extension
	 * @param providerId
	 * @param extensionId
	 * @param session
	 */
	updateSessionPreference(providerId: string, extensionId: string, session: AuthenticationSession): void;
	/**
	 * @deprecated Gets the session preference for the given provider and extension
	 * @param providerId
	 * @param extensionId
	 * @param scopes
	 */
	getSessionPreference(providerId: string, extensionId: string, scopes: string[]): string | undefined;
	/**
	 * @deprecated Removes the session preference for the given provider and extension
	 * @param providerId
	 * @param extensionId
	 * @param scopes
	 */
	removeSessionPreference(providerId: string, extensionId: string, scopes: string[]): void;
	selectSession(providerId: string, extensionId: string, extensionName: string, scopeListOrRequest: ReadonlyArray<string> | IAuthenticationWwwAuthenticateRequest, possibleSessions: readonly AuthenticationSession[], options?: IAuthenticationProviderSessionOptions): Promise<AuthenticationSession>;
	requestSessionAccess(providerId: string, extensionId: string, extensionName: string, scopeListOrRequest: ReadonlyArray<string> | IAuthenticationWwwAuthenticateRequest, possibleSessions: readonly AuthenticationSession[], options?: IAuthenticationProviderSessionOptions): void;
	requestNewSession(providerId: string, scopeListOrRequest: ReadonlyArray<string> | IAuthenticationWwwAuthenticateRequest, extensionId: string, extensionName: string, options?: IAuthenticationProviderSessionOptions): Promise<void>;
	updateNewSessionRequests(providerId: string, addedSessions: readonly AuthenticationSession[]): Promise<void>;
}

/**
 * Options passed to the authentication provider when asking for sessions.
 */
export interface IAuthenticationProviderSessionOptions {
	/**
	 * Whether the provider must avoid user interaction while resolving existing sessions.
	 */
	silent?: boolean;
	/**
	 * The account that is being asked about. If this is passed in, the provider should
	 * attempt to return the sessions that are only related to this account.
	 */
	account?: AuthenticationSessionAccount;
	/**
	 * The authorization server that is being asked about. If this is passed in, the provider should
	 * attempt to return sessions that are only related to this authorization server.
	 */
	authorizationServer?: URI;
	/**
	 * The OAuth client identifier to use instead of the provider's default client.
	 */
	clientId?: string;
	/**
	 * When specified, the authentication provider will request a token bound to this resource URI
	 * (RFC 8707 resource indicator).
	 */
	resource?: string;
	/**
	 * The audience for the requested access token. Primarily used for OAuth Identity Assertion
	 * Authorization Grant (ID-JAG, defined in `draft-ietf-oauth-identity-assertion-authz-grant` using RFC 8693 token-exchange semantics) flows where the audience identifies the authorization server of the resource that
	 * will redeem the assertion (typically the resource's authorization server URL). Providers that do not understand audience-bound tokens should
	 * ignore this option.
	 */
	audience?: string;
	/**
	 * Allows the authentication provider to take in additional parameters.
	 * It is up to the provider to define what these parameters are and handle them.
	 * This is useful for passing in additional information that is specific to the provider
	 * and not part of the standard authentication flow.
	 */
	[key: string]: any;
}

/**
 * Prefix for {@link IAuthenticationProviderSessionOptions} keys that only the workbench may set.
 * Keys with this prefix are removed from every extension's request before it reaches a provider,
 * so a provider that receives one knows the workbench itself made the request.
 */
export const WORKBENCH_ONLY_SESSION_OPTION_PREFIX = '_workbench';

/** A resource-bound MCP connection may keep working with a partially granted OAuth session. */
export const MCP_ACCEPT_PARTIAL_OAUTH_SCOPES = '_workbenchMcpAcceptPartialOAuthScopes';

/** A registered MCP client explicitly selected the Signal-C desktop HTTPS callback. */
export const MCP_USE_HTTPS_OAUTH_CALLBACK = '_workbenchMcpUseHttpsOAuthCallback';

/** The access token fingerprint of an approved, resource-bound MCP grant to refresh silently. */
export const MCP_REFRESH_REJECTED_OAUTH_TOKEN = '_workbenchMcpRefreshRejectedOAuthToken';

/** Returns a copy of `options` without the keys reserved for the workbench. */
export function withoutWorkbenchOnlySessionOptions<T extends object>(options: T): T {
	// Reserved keys are never declared on `T` (they only exist through its index signature), so
	// every property `T` does declare survives and the result is still a `T`.
	return filter(options, key => !key.startsWith(WORKBENCH_ONLY_SESSION_OPTION_PREFIX)) as T;
}

/** Builds a value-sensitive request key without splitting opaque provider scope strings. */
export function getAuthenticationSessionRequestKey(scopeListOrRequest: readonly string[] | IAuthenticationWwwAuthenticateRequest, options: IAuthenticationProviderSessionOptions = {}): string {
	const request = isAuthenticationWwwAuthenticateRequest(scopeListOrRequest)
		? { wwwAuthenticate: scopeListOrRequest.wwwAuthenticate, fallbackScopes: scopeListOrRequest.fallbackScopes ? [...scopeListOrRequest.fallbackScopes].sort() : undefined }
		: [...scopeListOrRequest].sort();
	const keyOptions = {
		...options,
		account: options.account?.id,
		authorizationServer: options.authorizationServer?.toString(true)
	};
	return JSON.stringify([
		request,
		Object.entries(keyOptions)
			.filter(([, value]) => value !== undefined)
			.sort(([first], [second]) => compare(first, second))
	]);
}

/**
 * Represents an authentication provider.
 */
export interface IAuthenticationProvider {
	/**
	 * The unique identifier of the authentication provider.
	 */
	readonly id: string;

	/**
	 * The display label of the authentication provider.
	 */
	readonly label: string;

	/**
	 * The resource server URI that this provider is responsible for, if any.
	 * TODO@TylerLeonhardt: Rather than this being added to the provider, it should be passed in to
	 * getSessions/createSession/etc... this way we can have providers that handle multiple resource servers.
	 */
	readonly resourceServer?: URI;

	/**
	 * The resolved authorization servers. These can still contain globs, but should be concrete URIs
	 */
	readonly authorizationServers?: ReadonlyArray<URI>;

	/**
	 * Indicates whether the authentication provider supports multiple accounts.
	 */
	readonly supportsMultipleAccounts: boolean;

	/**
	 * Optional function to provide a custom confirmation message for authentication prompts.
	 * If not implemented, the default confirmation messages will be used.
	 * @param extensionName - The name of the extension requesting authentication.
	 * @param recreatingSession - Whether this is recreating an existing session.
	 * @returns A custom confirmation message or undefined to use the default message.
	 */
	readonly confirmation?: (extensionName: string, recreatingSession: boolean) => string | undefined;

	/**
	 * An {@link Event} which fires when the array of sessions has changed, or data
	 * within a session has changed.
	 */
	readonly onDidChangeSessions: Event<AuthenticationSessionsChangeEvent>;

	/**
	 * Retrieves a list of authentication sessions.
	 * @param scopes - An optional list of scopes. If provided, the sessions returned should match these permissions, otherwise all sessions should be returned.
	 * @param options - Additional options for getting sessions.
	 * @returns A promise that resolves to an array of authentication sessions.
	 */
	getSessions(scopes: string[] | undefined, options: IAuthenticationProviderSessionOptions): Promise<readonly AuthenticationSession[]>;

	/**
	 * Prompts the user to log in.
	 * If login is successful, the `onDidChangeSessions` event should be fired.
	 * If login fails, a rejected promise should be returned.
	 * If the provider does not support multiple accounts, this method should not be called if there is already an existing session matching the provided scopes.
	 * @param scopes - A list of scopes that the new session should be created with.
	 * @param options - Additional options for creating the session.
	 * @returns A promise that resolves to an authentication session.
	 */
	createSession(scopes: string[], options: IAuthenticationProviderSessionOptions): Promise<AuthenticationSession>;

	/**
	 * Get existing sessions that match the given authentication constraints.
	 *
	 * @param constraint The authentication constraint containing challenges and optional scopes
	 * @param options Options for the session request
	 * @returns A thenable that resolves to an array of existing authentication sessions
	 */
	getSessionsFromChallenges?(constraint: IAuthenticationConstraint, options: IAuthenticationProviderSessionOptions): Promise<readonly AuthenticationSession[]>;

	/**
	 * Create a new session based on authentication constraints.
	 * This is called when no existing session matches the constraint requirements.
	 *
	 * @param constraint The authentication constraint containing challenges and optional scopes
	 * @param options Options for the session creation
	 * @returns A thenable that resolves to a new authentication session
	 */
	createSessionFromChallenges?(constraint: IAuthenticationConstraint, options: IAuthenticationProviderSessionOptions): Promise<AuthenticationSession>;

	/**
	 * Removes the session corresponding to the specified session ID.
	 * If the removal is successful, the `onDidChangeSessions` event should be fired.
	 * If a session cannot be removed, the provider should reject with an error message.
	 * @param sessionId - The ID of the session to remove.
	 */
	removeSession(sessionId: string): Promise<void>;
}
