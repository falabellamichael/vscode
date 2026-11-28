/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { SequencerByKey } from '../../../../base/common/async.js';
import { decodeBase64, encodeBase64, VSBuffer } from '../../../../base/common/buffer.js';
import { CancellationToken } from '../../../../base/common/cancellation.js';
import { CancellationError, isCancellationError } from '../../../../base/common/errors.js';
import { Disposable } from '../../../../base/common/lifecycle.js';
import { IObservable, observableFromEvent } from '../../../../base/common/observable.js';
import { isAbsolute } from '../../../../base/common/path.js';
import { isNative } from '../../../../base/common/platform.js';
import { URI } from '../../../../base/common/uri.js';
import { generateUuid, isUUID } from '../../../../base/common/uuid.js';
import { localize } from '../../../../nls.js';
import { IDialogService } from '../../../../platform/dialogs/common/dialogs.js';
import { IFileService } from '../../../../platform/files/common/files.js';
import { InstantiationType, registerSingleton } from '../../../../platform/instantiation/common/extensions.js';
import { createDecorator } from '../../../../platform/instantiation/common/instantiation.js';
import { INotificationService } from '../../../../platform/notification/common/notification.js';
import { IOpenerService } from '../../../../platform/opener/common/opener.js';
import { IProductService } from '../../../../platform/product/common/productService.js';
import { IQuickInputService } from '../../../../platform/quickinput/common/quickInput.js';
import { asJson, IRequestService, isSuccess, NO_FETCH_TELEMETRY } from '../../../../platform/request/common/request.js';
import { IStorageService, StorageScope, StorageTarget } from '../../../../platform/storage/common/storage.js';
import { AuthenticationSession, IAuthenticationService, signalCMailOAuth, signalCMailOAuthPrefix, SignalCMailProvider } from '../../../services/authentication/common/authentication.js';
import { IDynamicAuthenticationProviderStorageService } from '../../../services/authentication/common/dynamicAuthenticationProviderStorage.js';
import { IUserDataProfileService } from '../../../services/userDataProfile/common/userDataProfile.js';
import { signalCDirectMailConnections } from '../common/signalCAppConnections.js';

const MAIL_STORAGE_PREFIX = 'signalC.directMail.';
const mailProviders = ['gmail', 'outlook'] as const;

interface IMailbox {
	readonly account: string;
	readonly registration?: string;
	readonly canSend?: true;
	readonly canReadCalendar?: true;
}

interface IGooglePart {
	readonly mimeType?: string;
	readonly body?: { readonly data?: string };
	readonly filename?: string;
	readonly parts?: readonly IGooglePart[];
	readonly headers?: readonly { readonly name: string; readonly value: string }[];
}

interface IGoogleMessage {
	readonly id: string;
	readonly snippet?: string;
	readonly payload?: IGooglePart;
}

interface IGraphMessage {
	readonly id: string;
	readonly subject?: string;
	readonly from?: { readonly emailAddress?: { readonly address?: string } };
	readonly toRecipients?: readonly { readonly emailAddress?: { readonly address?: string } }[];
	readonly receivedDateTime?: string;
	readonly sentDateTime?: string;
	readonly importance?: string;
	readonly bodyPreview?: string;
	readonly body?: { readonly contentType?: string; readonly content?: string };
	readonly webLink?: string;
}

/** Combines a user query with Gmail's unread/important markers. */
export function directMailGmailQuery(query: string, unreadOnly: boolean, importance: string): string {
	const parts = [query.trim()];
	if (unreadOnly) { parts.push('is:unread'); }
	if (importance === 'high') { parts.push('is:important'); }
	return parts.filter(Boolean).join(' ');
}

/** Builds a Graph $filter for unread/importance constraints (cannot combine with $search). */
export function directMailGraphFilter(unreadOnly: boolean, importance: string): string {
	const parts: string[] = [];
	if (unreadOnly) { parts.push('isRead eq false'); }
	if (importance === 'high' || importance === 'normal' || importance === 'low') { parts.push(`importance eq '${importance}'`); }
	return parts.join(' and ');
}

export interface ISignalCDirectMailConnection {
	readonly id: string;
	readonly provider: SignalCMailProvider;
	readonly label: string;
	readonly needsSignIn: boolean;
	readonly canSend?: true;
	readonly canReadCalendar?: true;
}

export const ISignalCDirectMail = createDecorator<ISignalCDirectMail>('signalCDirectMail');

export interface ISignalCDirectMail {
	readonly _serviceBrand: undefined;
	readonly connections: IObservable<readonly ISignalCDirectMailConnection[]>;
	manage(provider: SignalCMailProvider): Promise<void>;
	read(connection: ISignalCDirectMailConnection, args: Record<string, unknown>, token: CancellationToken): Promise<string>;
	readCalendar(connection: ISignalCDirectMailConnection, args: Record<string, unknown>, token: CancellationToken): Promise<string>;
	send(connection: ISignalCDirectMailConnection, args: Record<string, unknown>, token: CancellationToken): Promise<string>;
}

function googlePlainText(part: IGooglePart | undefined, depth = 0): string | undefined {
	if (!part || part.filename || depth > 5) { return undefined; }
	if (part.mimeType === 'text/plain' && part.body?.data) {
		try { return decodeBase64(part.body.data.slice(0, 32000)).toString().slice(0, 6000); } catch { return undefined; }
	}
	for (const child of part.parts?.slice(0, 20) ?? []) {
		const text = googlePlainText(child, depth + 1);
		if (text) { return text; }
	}
	return undefined;
}

function googleResult(message: IGoogleMessage, full: boolean): object {
	const header = (name: string) => message.payload?.headers?.find(item => item.name.toLowerCase() === name)?.value.slice(0, 400);
	return {
		id: message.id, subject: header('subject'), from: header('from'), to: header('to'), date: header('date'),
		preview: message.snippet?.slice(0, 400), body: full ? googlePlainText(message.payload) : undefined,
		url: `https://mail.google.com/mail/u/0/#all/${encodeURIComponent(message.id)}`,
	};
}

function graphResult(message: IGraphMessage, full: boolean): object {
	return {
		id: message.id, subject: message.subject?.slice(0, 400), from: message.from?.emailAddress?.address,
		to: message.toRecipients?.slice(0, 8).map(recipient => recipient.emailAddress?.address),
		received: message.receivedDateTime, sent: message.sentDateTime, importance: message.importance,
		preview: message.bodyPreview?.slice(0, 400),
		body: full && message.body?.contentType === 'text' ? message.body.content?.slice(0, 6000) : undefined,
		url: message.webLink?.startsWith('https://') ? message.webLink : undefined,
	};
}

export function isSignalCMailAddress(value: string): boolean {
	return value.length <= 254 && /^[\w.!#$%&'*+/=?^`{|}~-]+@[a-z\d](?:[a-z\d.-]*[a-z\d])?\.[a-z]{2,}$/i.test(value);
}

function gmailMessage(from: string, to: string, subject: string, body: string): string {
	const encodedSubject = (subject.match(/.{1,11}/gu) ?? []).map(part => `=?UTF-8?B?${encodeBase64(VSBuffer.fromString(part))}?=`).join('\r\n ');
	const encodedBody = encodeBase64(VSBuffer.fromString(body.replace(/\r\n|\r|\n/g, '\r\n'))).match(/.{1,76}/g)?.join('\r\n');
	return encodeBase64(VSBuffer.fromString([
		`From: ${from}`, `To: ${to}`, `Subject: ${encodedSubject}`,
		'MIME-Version: 1.0', 'Content-Type: text/plain; charset=UTF-8', 'Content-Transfer-Encoding: base64', '', encodedBody,
	].join('\r\n')), false, true);
}

class MailUnauthorizedError extends Error {}

export class SignalCDirectMail extends Disposable implements ISignalCDirectMail {
	declare readonly _serviceBrand: undefined;
	private readonly operations = new SequencerByKey<SignalCMailProvider>();
	private readonly tokenRequests = new SequencerByKey<SignalCMailProvider>();
	readonly connections: IObservable<readonly ISignalCDirectMailConnection[]>;

	constructor(
		@IAuthenticationService private readonly authentication: IAuthenticationService,
		@IDynamicAuthenticationProviderStorageService private readonly grants: IDynamicAuthenticationProviderStorageService,
		@IStorageService private readonly storage: IStorageService,
		@IUserDataProfileService private readonly profiles: IUserDataProfileService,
		@IQuickInputService private readonly input: IQuickInputService,
		@IDialogService private readonly dialogs: IDialogService,
		@IOpenerService private readonly opener: IOpenerService,
		@INotificationService private readonly notifications: INotificationService,
		@IRequestService private readonly requests: IRequestService,
		@IProductService private readonly product: IProductService,
		@IFileService private readonly files: IFileService,
	) {
		super();
		this.connections = observableFromEvent(this, storage.onDidChangeValue(StorageScope.PROFILE, undefined, this._store), () => mailProviders.flatMap(provider => {
			const mailbox = this.mailbox(provider);
			return mailbox ? [{ id: this.connectionId(provider, mailbox.account), provider, label: signalCDirectMailConnections[provider].name, needsSignIn: this.needsSignIn(provider), ...(mailbox.canSend ? { canSend: true as const } : {}), ...(mailbox.canReadCalendar ? { canReadCalendar: true as const } : {}) }] : [];
		}));
	}

	manage(provider: SignalCMailProvider): Promise<void> {
		return this.operations.queue(provider, async () => {
			if (!isNative) { throw new Error(localize('directMail.desktopOnly', "Direct mail sign-in is available in the Signal-C desktop app.")); }
			const app = signalCDirectMailConnections[provider];
			const key = MAIL_STORAGE_PREFIX + provider;
			const before = this.storage.get(key, StorageScope.PROFILE);
			const mailbox = this.mailbox(provider);
			const baseId = this.baseProviderId(provider);
			const id = this.providerId(provider, mailbox?.registration);
			const registration = await this.grants.getClientRegistration(id);
			const bundled = await this.bundledRegistration(provider);
			const defaultRegistration = mailbox && registration?.clientId ? registration : bundled ?? registration;
			const { result } = await this.dialogs.prompt({
				type: 'info', custom: true,
				message: localize('directMail.title', "Connect {0}", app.name),
				detail: mailbox && !this.needsSignIn(provider)
					? mailbox.canSend ? localize('directMail.connectedSendDetail', "This mailbox is authorized for reading and sending. You can sign in again, switch to read-only access, or disconnect.")
						: localize('directMail.connectedDetail', "This mailbox is read-only. Sign in through your browser to let Assistant send email from this account.")
					: bundled ? localize('directMail.bundledApp', "Sign in with your {0} account in the browser to let Assistant read and send email.", provider === 'gmail' ? 'Google' : 'Microsoft') : app.setup,
				buttons: [
					mailbox?.canSend ? { label: localize('directMail.reconnect', "Sign In Again"), run: () => 'connect' as const }
						: { label: localize('directMail.enableSending', "Sign In to Read and Send"), run: () => 'enableSend' as const },
					{ label: mailbox?.canSend ? localize('directMail.readOnly', "Switch to Read Only")
						: localize('directMail.signInReadOnly', "Sign In to Read Only"), run: () => 'readOnly' as const },
					...(provider === 'outlook' && mailbox ? [mailbox.canReadCalendar
						? { label: localize('directMail.removeCalendar', "Remove Calendar Access"), run: () => 'removeCalendar' as const }
						: { label: localize('directMail.enableCalendar', "Enable Calendar Access"), run: () => 'enableCalendar' as const }] : []),
					...(mailbox && bundled && (defaultRegistration?.clientId !== bundled.clientId || defaultRegistration?.clientSecret !== bundled.clientSecret) ? [{
						label: provider === 'gmail' ? localize('directMail.useBundledGmailApp', "Use Built-In Google Sign-In") : localize('directMail.useBundledOutlookApp', "Use Built-In Microsoft Sign-In"), run: () => 'bundled' as const,
					}] : []),
					...(registration?.clientId || bundled ? [{
						label: localize('directMail.changeApp', "Use Different OAuth App"), run: () => 'replace' as const,
					}] : []),
					...(mailbox || registration?.clientId ? [{
						label: mailbox ? localize('directMail.disconnect', "Disconnect") : localize('directMail.reset', "Reset OAuth Setup"), run: () => 'disconnect' as const,
					}] : []),
					{ label: localize('directMail.guide', "Open Setup Guide"), run: () => 'guide' as const },
				],
				cancelButton: true,
			});
			if (result === 'guide') { await this.opener.open(URI.parse(app.documentation)); return; }
			const stillCurrent = () => baseId === this.baseProviderId(provider) && this.storage.get(key, StorageScope.PROFILE) === before;
			if (!stillCurrent()) { return; }
			if (result === 'disconnect') {
				await this.tokenRequests.queue(provider, async () => {
					if (!stillCurrent()) { return; }
					this.storage.remove(key, StorageScope.PROFILE);
					this.setSignInRequired(provider, false);
					await this.removeProvider(id);
				});
				return;
			}
			if (!result) { return; }
			const selectedRegistration = result === 'replace' ? undefined : result === 'bundled' ? bundled : defaultRegistration;
			const useNewProvider = !!mailbox || result === 'replace' || !!registration?.clientId
				&& (selectedRegistration?.clientId !== registration.clientId || selectedRegistration?.clientSecret !== registration.clientSecret);
			const enableSend = result === 'enableSend' || (result !== 'readOnly' && mailbox?.canSend === true);
			const enableCalendar = provider === 'outlook' && (result === 'enableCalendar' || (result !== 'removeCalendar' && mailbox?.canReadCalendar === true));
			const expectedAccount = result === 'enableSend' || result === 'readOnly' || result === 'enableCalendar' || result === 'removeCalendar' ? mailbox?.account : undefined;
			await this.connect(provider, id, { registration: selectedRegistration, useNewProvider, stillCurrent, enableSend, enableCalendar, expectedAccount });
		});
	}

	private async bundledRegistration(provider: SignalCMailProvider): Promise<{ clientId: string; clientSecret?: string } | undefined> {
		if (provider === 'outlook') {
			const clientId = this.product.signalCOutlookOAuthClientId;
			return clientId && isUUID(clientId) ? { clientId } : undefined;
		}
		const path = this.product.signalCGmailOAuthClientFile;
		if (!path || !isAbsolute(path)) { return undefined; }
		try {
			const { value } = await this.files.readFile(URI.file(path), { limits: { size: 32 * 1024 } });
			const google = JSON.parse(value.toString())?.google;
			if (typeof google?.clientId === 'string' && /^[a-z0-9.-]{1,230}\.apps\.googleusercontent\.com$/i.test(google.clientId)
				&& typeof google.clientSecret === 'string' && /^[a-z0-9._-]{1,256}$/i.test(google.clientSecret)) {
				return { clientId: google.clientId, clientSecret: google.clientSecret };
			}
		} catch {
			// Unprovisioned builds can accept a user-registered Desktop OAuth app.
		}
		return undefined;
	}

	private async connect(provider: SignalCMailProvider, currentId: string, { registration: saved, useNewProvider, stillCurrent, enableSend, enableCalendar, expectedAccount }: {
		readonly registration?: { clientId?: string; clientSecret?: string };
		readonly useNewProvider: boolean;
		readonly stillCurrent: () => boolean;
		readonly enableSend: boolean;
		readonly enableCalendar: boolean;
		readonly expectedAccount?: string;
	}): Promise<void> {
		const name = signalCDirectMailConnections[provider].name;
		const candidate = useNewProvider ? generateUuid() : undefined;
		const id = this.providerId(provider, candidate);
		const clientId = saved?.clientId ?? (await this.input.input({
			title: localize('directMail.clientTitle', "{0} OAuth Client ID", name),
			prompt: localize('directMail.clientPrompt', "Enter the OAuth client ID from your app registration."),
			ignoreFocusLost: true,
			validateInput: async value => (provider === 'gmail' ? /^[a-z0-9.-]+\.apps\.googleusercontent\.com$/i : /^[0-9a-f]{8}(?:-[0-9a-f]{4}){3}-[0-9a-f]{12}$/i).test(value.trim())
				? undefined : localize('directMail.clientInvalid', "Enter the public OAuth client ID."),
		}))?.trim();
		if (!clientId) { return; }
		const secret = provider === 'gmail' ? saved?.clientSecret ?? (await this.input.input({
			title: localize('directMail.secretTitle', "Gmail Desktop Client Secret"),
			password: true, ignoreFocusLost: true,
			validateInput: async value => value.trim() ? undefined : localize('directMail.secretRequired', "Enter the Google Desktop client secret."),
		}))?.trim() : undefined;
		if ((provider === 'gmail' && !secret) || !stillCurrent()) { return; }
		let session: AuthenticationSession | undefined;
		let previous: readonly AuthenticationSession[] = [];
		try {
			await this.ensureProvider(provider, id, { clientId, clientSecret: secret });
			if (!candidate) { previous = await this.authentication.getSessions(id); }
			session = await this.authentication.createSession(id, this.scopes(provider, enableSend, enableCalendar));
			if (session.account.id === 'unknown') {
				throw new Error(localize('directMail.noAccount', "Sign-in did not identify a mailbox account. Check that the app supports OpenID Connect."));
			}
			if (expectedAccount && session.account.id !== expectedAccount) {
				throw new Error(localize('directMail.wrongAccount', "Sign in to the same mailbox to change what Assistant can do with it."));
			}
			if (enableSend && !session.scopes.includes(signalCMailOAuth[provider].sendScope)) {
				throw new Error(localize('directMail.sendConsentMissing', "The provider did not grant mail sending permission. Sign in and grant the requested access."));
			}
			const url = new URL(provider === 'gmail' ? 'gmail/v1/users/me/profile' : 'v1.0/me/mailFolders/inbox/messages?$top=1&$select=id', signalCMailOAuth[provider].resource);
			await this.get(url, session.accessToken, CancellationToken.None);
			if (enableCalendar) {
				if (!session.scopes.includes(signalCMailOAuth.outlook.calendarScope)) {
					throw new Error(localize('directMail.calendarConsentMissing', "Microsoft did not grant calendar read permission."));
				}
				await this.get(new URL('v1.0/me/calendar?$select=id', signalCMailOAuth.outlook.resource), session.accessToken, CancellationToken.None);
			}
			const account = session.account.id;
			await this.tokenRequests.queue(provider, async () => {
				if (!stillCurrent()) { throw new CancellationError(); }
				const value = candidate || enableSend || enableCalendar ? JSON.stringify({ account, registration: candidate, ...(enableSend ? { canSend: true } : {}), ...(enableCalendar ? { canReadCalendar: true } : {}) } satisfies IMailbox) : account;
				this.storage.store(MAIL_STORAGE_PREFIX + provider, value, StorageScope.PROFILE, StorageTarget.MACHINE);
				this.setSignInRequired(provider, false);
			});
		} catch (error) {
			if (candidate) {
				await this.removeProvider(id).catch(() => this.notifyCleanupFailure(name));
			} else if (session) {
				await this.authentication.removeSession(id, session.id).catch(() => this.notifyCleanupFailure(name));
			}
			throw error;
		}
		try {
			if (candidate) {
				await this.removeProvider(currentId);
			} else {
				for (const old of previous) {
					if (old.id !== session.id) { await this.authentication.removeSession(id, old.id); }
				}
			}
		} catch {
			this.notifyCleanupFailure(name);
		}
		this.notifications.info(enableCalendar ? localize('directMail.calendarEnabled', "Microsoft 365 calendar access verified.") : enableSend ? localize('directMail.sendEnabled', "{0} is connected and mail sending is authorized.", name) : localize('directMail.verified', "{0} mailbox access verified.", name));
	}

	async read(connection: ISignalCDirectMailConnection, args: Record<string, unknown>, token: CancellationToken): Promise<string> {
		return this.useMailbox(connection, token, false, session => this.readMailbox(connection.provider, session, args, token));
	}

	async readCalendar(connection: ISignalCDirectMailConnection, args: Record<string, unknown>, token: CancellationToken): Promise<string> {
		if (connection.provider !== 'outlook' || this.mailbox('outlook')?.canReadCalendar !== true) {
			throw new Error(localize('directMail.calendarRequired', "Open the Outlook connection and choose Enable Calendar Access."));
		}
		const start = typeof args.start === 'string' ? Date.parse(args.start) : NaN;
		const end = typeof args.end === 'string' ? Date.parse(args.end) : NaN;
		if (!Number.isFinite(start) || !Number.isFinite(end) || end <= start || end - start > 93 * 86400_000) {
			throw new Error(localize('directMail.calendarRange', "Specify a valid start and end time within a 93-day range."));
		}
		return this.useMailbox(connection, token, false, async session => {
			if (!session.scopes.includes(signalCMailOAuth.outlook.calendarScope)) { throw this.signInRequired('outlook'); }
			const url = new URL('v1.0/me/calendar/calendarView', signalCMailOAuth.outlook.resource);
			url.searchParams.set('startDateTime', new Date(start).toISOString());
			url.searchParams.set('endDateTime', new Date(end).toISOString());
			url.searchParams.set('$select', 'id,subject,start,end,isAllDay,isCancelled,webLink');
			url.searchParams.set('$top', '50');
			return JSON.stringify(await this.get(url, session.accessToken, token));
		});
	}

	async send(connection: ISignalCDirectMailConnection, args: Record<string, unknown>, token: CancellationToken): Promise<string> {
		const to = typeof args.to === 'string' ? args.to.trim() : '';
		const subject = typeof args.subject === 'string' ? args.subject.trim() : '';
		const body = typeof args.body === 'string' ? args.body : '';
		if (!isSignalCMailAddress(to)) {
			throw new Error(localize('directMail.invalidRecipient', "Enter one complete recipient email address."));
		}
		if (!subject || subject.length > 250 || Array.from(subject).some(char => char.charCodeAt(0) < 32 || char.charCodeAt(0) === 127) || !body.trim() || body.length > 20_000) {
			throw new Error(localize('directMail.invalidMessageFields', "Enter a single-line subject of up to 250 characters and a message of up to 20,000 characters."));
		}
		return this.useMailbox(connection, token, true, (session, isCurrent) => this.sendMailbox(connection.provider, session.accessToken, to, subject, body, token, isCurrent));
	}

	private async useMailbox(connection: ISignalCDirectMailConnection, token: CancellationToken, sending: boolean, action: (session: AuthenticationSession, isCurrent: () => boolean) => Promise<string>): Promise<string> {
		const { provider } = connection;
		const mailbox = this.mailbox(provider);
		const account = mailbox?.account;
		if (!account || connection.id !== this.connectionId(provider, account) || this.needsSignIn(provider)) { throw this.signInRequired(provider); }
		if (sending && !mailbox.canSend) {
			throw new Error(localize('directMail.sendPermissionRequired', "Open this mailbox connection and choose Sign In to Read and Send."));
		}
		const id = this.providerId(provider, mailbox.registration);
		const isCurrent = () => {
			const current = this.mailbox(provider);
			return current?.account === account && this.providerId(provider, current.registration) === id;
		};
		const session = await this.tokenRequests.queue(provider, async () => {
			if (token.isCancellationRequested || !isCurrent()) { throw new CancellationError(); }
			const saved = await this.grants.getClientRegistration(id);
			if (!saved?.clientId) { return undefined; }
			await this.ensureProvider(provider, id, { clientId: saved.clientId, clientSecret: saved.clientSecret });
			return (await this.authentication.getSessions(id, this.scopes(provider, mailbox.canSend, mailbox.canReadCalendar), { silent: true })).find(current => current.account.id === account);
		});
		if (!isCurrent() || token.isCancellationRequested) { throw new CancellationError(); }
		if (!session) {
			this.setSignInRequired(provider, true);
			throw this.signInRequired(provider);
		}
		try {
			const result = await action(session, isCurrent);
			if (!sending && (!isCurrent() || token.isCancellationRequested)) { throw new CancellationError(); }
			return result;
		} catch (error) {
			if (!(error instanceof MailUnauthorizedError)) { throw error; }
			if (!isCurrent()) { throw new CancellationError(); }
			this.setSignInRequired(provider, true);
			throw this.signInRequired(provider);
		}
	}

	private async sendMailbox(provider: SignalCMailProvider, accessToken: string, to: string, subject: string, body: string, token: CancellationToken, isCurrent: () => boolean): Promise<string> {
		const base = signalCMailOAuth[provider].resource;
		const url = new URL(provider === 'gmail' ? 'gmail/v1/users/me/messages/send' : 'v1.0/me/sendMail', base);
		let data = JSON.stringify({
			message: { subject, body: { contentType: 'Text', content: body }, toRecipients: [{ emailAddress: { address: to } }] },
		});
		if (provider === 'gmail') {
			const profile = await this.get<{ emailAddress?: string }>(new URL('gmail/v1/users/me/profile', base), accessToken, token);
			if (typeof profile.emailAddress !== 'string' || !isSignalCMailAddress(profile.emailAddress)) {
				throw new Error(localize('directMail.noSender', "Gmail did not return a valid sending address for this mailbox."));
			}
			data = JSON.stringify({ raw: gmailMessage(profile.emailAddress, to, subject, body) });
		}
		if (token.isCancellationRequested || !isCurrent()) { throw new CancellationError(); }
		let response;
		try {
			response = await this.requests.request({
				url: url.href, type: 'POST', data, headers: { Authorization: `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
				followRedirects: 0, timeout: 20_000, disableCache: true, logUrl: false, callSite: NO_FETCH_TELEMETRY,
			}, token);
		} catch {
			throw new Error(localize('directMail.sendUnconfirmed', "The provider did not confirm submission. Check Sent before trying again."));
		}
		if (response.res.statusCode === 401) { throw new MailUnauthorizedError(localize('directMail.unauthorized', "The mailbox authorization was rejected.")); }
		if (!isSuccess(response)) { throw new Error(localize('directMail.sendRejected', "The mail provider returned HTTP {0}. Check sending permissions and your Sent folder before trying again.", response.res.statusCode)); }
		return localize('directMail.sendAccepted', "The mail provider accepted the message for sending.");
	}

	private async readMailbox(provider: SignalCMailProvider, session: AuthenticationSession, args: Record<string, unknown>, token: CancellationToken): Promise<string> {
		const messageId = typeof args.messageId === 'string' ? args.messageId.trim() : '';
		if (messageId && !/^[\w+/=-]{1,1024}$/.test(messageId)) { throw new Error(localize('directMail.invalidMessage', "Use a message ID returned by the mail tool.")); }
		const query = typeof args.query === 'string' ? args.query.trim().slice(0, 400) : '';
		const limit = typeof args.limit === 'number' && Number.isFinite(args.limit) ? Math.min(10, Math.max(1, Math.floor(args.limit))) : 5;
		const folder = args.folder === 'sent' ? 'sent' : args.folder === 'all' ? 'all' : 'inbox';
		const unreadOnly = args.unreadOnly === true;
		const importance = typeof args.importance === 'string' ? args.importance.trim().toLowerCase() : '';
		const get = <T>(url: URL) => this.get<T>(url, session.accessToken, token);
		const base = signalCMailOAuth[provider].resource;
		if (provider === 'gmail') {
			const path = 'gmail/v1/users/me/messages';
			const url = new URL(path, base);
			const getMessage = (id: string, full: boolean) => {
				const messageUrl = new URL(path + '/' + encodeURIComponent(id), base);
				if (!full) { messageUrl.searchParams.set('fields', 'id,snippet,payload(headers)'); }
				messageUrl.searchParams.set('format', 'full');
				return get<IGoogleMessage>(messageUrl);
			};
			if (messageId) { return JSON.stringify(googleResult(await getMessage(messageId, true), true)); }
			url.searchParams.set('maxResults', String(limit));
			if (folder !== 'all') { url.searchParams.set('labelIds', folder === 'sent' ? 'SENT' : 'INBOX'); }
			const combinedQuery = directMailGmailQuery(query, unreadOnly, importance);
			if (combinedQuery) { url.searchParams.set('q', combinedQuery); }
			const listed = await get<{ messages?: readonly { id: string }[] }>(url);
			const messages = await Promise.all((listed.messages ?? []).slice(0, limit).map(({ id }) => getMessage(id, false)));
			return JSON.stringify({ messages: messages.map(message => googleResult(message, false)) });
		}
		const graphFields = 'id,subject,from,toRecipients,receivedDateTime,sentDateTime,importance,bodyPreview,webLink';
		const url = new URL(messageId ? `v1.0/me/messages/${encodeURIComponent(messageId)}`
			: folder === 'all' ? 'v1.0/me/messages' : `v1.0/me/mailFolders/${folder === 'sent' ? 'sentitems' : 'inbox'}/messages`, base);
		url.searchParams.set('$select', messageId ? graphFields + ',body' : graphFields);
		if (messageId) { return JSON.stringify(graphResult(await get<IGraphMessage>(url), true)); }
		url.searchParams.set('$top', String(limit));
		if (query) {
			url.searchParams.set('$search', `"${query.replace(/["\\]/g, ' ')}"`);
		} else {
			const filter = directMailGraphFilter(unreadOnly, importance);
			if (filter) { url.searchParams.set('$filter', filter); }
			url.searchParams.set('$orderby', folder === 'sent' ? 'sentDateTime desc' : 'receivedDateTime desc');
		}
		const response = await get<{ value: readonly IGraphMessage[] }>(url);
		return JSON.stringify({ messages: response.value.slice(0, limit).map(message => graphResult(message, false)) });
	}

	private async get<T>(url: URL, accessToken: string, token: CancellationToken): Promise<T> {
		const headers: Record<string, string> = { Authorization: `Bearer ${accessToken}`, Accept: 'application/json' };
		if (url.origin === 'https://graph.microsoft.com') { headers.Prefer = 'outlook.body-content-type="text"'; }
		let response;
		try {
			response = await this.requests.request({
				url: url.href, type: 'GET', headers, followRedirects: 0, timeout: 20_000,
				disableCache: true, logUrl: false, callSite: NO_FETCH_TELEMETRY,
			}, token);
		} catch (error) {
			if (isCancellationError(error)) { throw error; }
			if (token.isCancellationRequested) { throw new CancellationError(); }
			throw new Error(localize('directMail.requestFailed', "The mailbox read failed. Check the connection and mailbox permissions."));
		}
		if (response.res.statusCode === 401) { throw new MailUnauthorizedError(localize('directMail.unauthorized', "The mailbox authorization was rejected.")); }
		if (!isSuccess(response)) { throw new Error(localize('directMail.httpStatus', "The mail provider returned HTTP {0}. Check mailbox permissions and setup.", response.res.statusCode)); }
		const data = await asJson<T>(response).catch(() => undefined);
		if (token.isCancellationRequested) { throw new CancellationError(); }
		if (data) { return data; }
		throw new Error(localize('directMail.invalidResponse', "The mail provider returned an invalid response."));
	}

	private async ensureProvider(provider: SignalCMailProvider, id: string, registration: { clientId: string; clientSecret?: string }): Promise<void> {
		if (this.authentication.isAuthenticationProviderRegistered(id)) { return; }
		const config = signalCMailOAuth[provider];
		const registrationResult = await this.authentication.createDynamicAuthenticationProvider(
			URI.parse(config.issuer),
			{ issuer: config.issuer, authorization_endpoint: config.authorize, token_endpoint: config.token, response_types_supported: ['code'], code_challenge_methods_supported: ['S256'] },
			{ resource: config.resource, resource_name: signalCDirectMailConnections[provider].name },
			registration.clientId, registration.clientSecret, id
		);
		if (!registrationResult) { throw new Error(localize('directMail.authUnavailable', "The desktop sign-in host is unavailable. Try again after Signal-C finishes opening.")); }
	}

	private scopes(provider: SignalCMailProvider, sending = false, calendar = false): string[] {
		const { scope, sendScope } = signalCMailOAuth[provider];
		return [scope, ...(sending ? [sendScope] : []), 'openid', ...(provider === 'gmail' ? ['email'] : ['profile', 'offline_access', ...(calendar ? [signalCMailOAuth.outlook.calendarScope] : [])])];
	}

	private baseProviderId(provider: SignalCMailProvider): string {
		return `${signalCMailOAuthPrefix}${provider}:${encodeURIComponent(this.profiles.currentProfile.id)}`;
	}

	private providerId(provider: SignalCMailProvider, registration?: string): string {
		const base = this.baseProviderId(provider);
		return registration ? `${base}:${registration}` : base;
	}

	private async removeProvider(id: string): Promise<void> {
		if (this.authentication.isAuthenticationProviderRegistered(id)) { this.authentication.unregisterAuthenticationProvider(id); }
		await this.grants.removeDynamicProvider(id);
	}

	private notifyCleanupFailure(name: string): void {
		this.notifications.warn(localize('directMail.cleanupFailed', "Could not remove unused {0} OAuth credentials.", name));
	}

	private connectionId(provider: SignalCMailProvider, account: string): string {
		return JSON.stringify(['signal-c-direct-mail', provider, account]);
	}

	private mailbox(provider: SignalCMailProvider): IMailbox | undefined {
		const value = this.storage.get(MAIL_STORAGE_PREFIX + provider, StorageScope.PROFILE);
		if (!value) { return undefined; }
		if (value.startsWith('{')) {
			try {
				const stored = JSON.parse(value) as Partial<IMailbox> | undefined;
				if (typeof stored?.account === 'string' && stored.account && (stored.registration === undefined || isUUID(stored.registration))) {
					return { account: stored.account, registration: stored.registration, ...(stored.canSend === true ? { canSend: true } : {}), ...(provider === 'outlook' && stored.canReadCalendar === true ? { canReadCalendar: true } : {}) };
				}
			} catch {
				// Older versions stored the provider subject as plain text.
			}
		}
		return { account: value };
	}

	private needsSignIn(provider: SignalCMailProvider): boolean {
		return this.storage.getBoolean(MAIL_STORAGE_PREFIX + provider + '.needsSignIn', StorageScope.PROFILE, false);
	}

	private setSignInRequired(provider: SignalCMailProvider, value: boolean): void {
		this.storage.store(MAIL_STORAGE_PREFIX + provider + '.needsSignIn', value, StorageScope.PROFILE, StorageTarget.MACHINE);
	}

	private signInRequired(provider: SignalCMailProvider): Error {
		return new Error(localize('directMail.signInRequired', "Sign in to {0} from Assistant’s Connections and try again.", signalCDirectMailConnections[provider].name));
	}
}

registerSingleton(ISignalCDirectMail, SignalCDirectMail, InstantiationType.Delayed);
