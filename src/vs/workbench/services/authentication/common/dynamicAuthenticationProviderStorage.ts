/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { createDecorator } from '../../../../platform/instantiation/common/instantiation.js';
import { SequencerByKey } from '../../../../base/common/async.js';
import { IAuthorizationProtectedResourceMetadata, IAuthorizationTokenResponse } from '../../../../base/common/oauth.js';
import { Event } from '../../../../base/common/event.js';
import { URI } from '../../../../base/common/uri.js';
import { getDynamicAuthenticationProviderId, IAuthenticationService } from './authentication.js';

export const IDynamicAuthenticationProviderStorageService = createDecorator<IDynamicAuthenticationProviderStorageService>('dynamicAuthenticationProviderStorageService');

export interface DynamicAuthenticationProviderInfo {
	readonly providerId: string;
	readonly label: string;
	/**
	 * @deprecated in favor of authorizationServer
	 */
	readonly issuer?: string;
	readonly authorizationServer: string;
	readonly clientId: string;
}

export interface DynamicAuthenticationProviderTokensChangeEvent {
	readonly authProviderId: string;
	readonly clientId: string;
	readonly tokens: (IAuthorizationTokenResponse & { created_at: number })[] | undefined;
}

/**
 * Service for managing storage of dynamic authentication provider data.
 */
export interface IDynamicAuthenticationProviderStorageService {
	readonly _serviceBrand: undefined;

	/**
	 * Event fired when tokens for a dynamic authentication provider change.
	 */
	readonly onDidChangeTokens: Event<DynamicAuthenticationProviderTokensChangeEvent>;

	/**
	 * Get the client registration for a dynamic authentication provider.
	 * @param providerId The provider ID or authorization server URL.
	 * @returns The client details if they exist, undefined otherwise.
	 */
	getClientRegistration(providerId: string): Promise<{ clientId?: string; clientSecret?: string; tokenEndpointAuthMethod?: string } | undefined>;

	/**
	 * Store the client registration for a dynamic authentication provider.
	 * @param providerId The provider ID or authorization server URL.
	 * @param authorizationServer The authorization server URL for the provider.
	 * @param clientId The client ID to store.
	 * @param clientSecret Optional client secret to store.
	 * @param label Optional label for the provider.
	 * @param tokenEndpointAuthMethod Optional method returned by dynamic client registration.
	 */
	storeClientRegistration(providerId: string, authorizationServer: string, clientId: string, clientSecret?: string, label?: string, tokenEndpointAuthMethod?: string): Promise<void>;

	/**
	 * Get all dynamic authentication providers that have been interacted with.
	 * @returns Array of provider information.
	 */
	getInteractedProviders(): ReadonlyArray<DynamicAuthenticationProviderInfo>;

	/**
	 * Remove a dynamic authentication provider and its stored data.
	 * @param providerId The provider ID to remove.
	 */
	removeDynamicProvider(providerId: string): Promise<void>;

	/**
	 * Get sessions for a dynamic authentication provider from secret storage.
	 * @param authProviderId The authentication provider ID.
	 * @param clientId The client ID.
	 * @returns Array of authorization tokens with creation timestamps, or undefined if none exist.
	 */
	getSessionsForDynamicAuthProvider(authProviderId: string, clientId: string): Promise<(IAuthorizationTokenResponse & { created_at: number })[] | undefined>;

	/**
	 * Set sessions for a dynamic authentication provider in secret storage.
	 * @param authProviderId The authentication provider ID.
	 * @param clientId The client ID.
	 * @param sessions Array of authorization tokens with creation timestamps.
	 */
	setSessionsForDynamicAuthProvider(authProviderId: string, clientId: string, sessions: (IAuthorizationTokenResponse & { created_at: number })[]): Promise<void>;
}

/** Retain previously stored grants for a matching client and separate other explicit clients. */
export async function resolveDynamicAuthenticationProviderId(
	authorizationServer: URI,
	resource: IAuthorizationProtectedResourceMetadata | undefined,
	client: { readonly clientId: string; readonly clientSecret?: string } | undefined,
	storage: IDynamicAuthenticationProviderStorageService,
): Promise<string> {
	const legacyId = getDynamicAuthenticationProviderId(authorizationServer, resource);
	if (!client) {
		return legacyId;
	}
	const registration = await storage.getClientRegistration(legacyId);
	return registration?.clientId === client.clientId && registration.clientSecret === client.clientSecret
		? legacyId : getDynamicAuthenticationProviderId(authorizationServer, resource, client.clientId);
}

const providerOperationsByService = new WeakMap<IAuthenticationService, SequencerByKey<string>>();

/** Serializes registration, credential rotation and rejected session recovery across MCP hosts. */
export function getMcpAuthenticationOperations(authenticationService: IAuthenticationService): SequencerByKey<string> {
	let operations = providerOperationsByService.get(authenticationService);
	if (!operations) {
		operations = new SequencerByKey();
		providerOperationsByService.set(authenticationService, operations);
	}
	return operations;
}
