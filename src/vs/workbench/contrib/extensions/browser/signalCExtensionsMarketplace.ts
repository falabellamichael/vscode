/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import './media/signalCExtensionsMarketplace.css';
import * as DOM from '../../../../base/browser/dom.js';
import { Button } from '../../../../base/browser/ui/button/button.js';
import { InputBox } from '../../../../base/browser/ui/inputbox/inputBox.js';
import { CancellationToken, cancelOnDispose } from '../../../../base/common/cancellation.js';
import { getErrorMessage } from '../../../../base/common/errors.js';
import { Disposable } from '../../../../base/common/lifecycle.js';
import { isWeb } from '../../../../base/common/platform.js';
import { gte } from '../../../../base/common/semver/semver.js';
import { URI } from '../../../../base/common/uri.js';
import { localize } from '../../../../nls.js';
import { IContextViewService } from '../../../../platform/contextview/browser/contextView.js';
import { ExtensionGalleryResourceType, getExtensionGalleryManifestResourceUri, IExtensionGalleryManifestService } from '../../../../platform/extensionManagement/common/extensionGalleryManifest.js';
import { IExtensionGalleryService } from '../../../../platform/extensionManagement/common/extensionManagement.js';
import { IInstantiationService } from '../../../../platform/instantiation/common/instantiation.js';
import { IFileService } from '../../../../platform/files/common/files.js';
import { ILogService } from '../../../../platform/log/common/log.js';
import { IOpenerService } from '../../../../platform/opener/common/opener.js';
import { defaultButtonStyles, defaultInputBoxStyles } from '../../../../platform/theme/browser/defaultStyles.js';
import { isLocalhostAuthority } from '../../../../platform/url/common/trustedDomains.js';
import { ViewPane } from '../../../browser/parts/views/viewPane.js';
import { ExtensionState, IExtension, IExtensionsWorkbenchService } from '../common/extensions.js';
import { getSignalCExtensionRepository, ISignalCExtensionDownloadService, ISignalCExtensionListing, matchesSignalCExtension, signalCExtensionCatalog } from '../../../../platform/extensionManagement/common/signalCExtensionCatalog.js';
import { UpdateAction } from './extensionsActions.js';

interface ISignalCExtensionInstallServices {
	readonly downloadService: ISignalCExtensionDownloadService;
	readonly fileService: IFileService;
	readonly extensionsWorkbenchService: IExtensionsWorkbenchService;
	readonly logService: ILogService;
}

/** Downloads an exact release, checks its pinned digest, and uses the normal VSIX installation path. */
export async function installSignalCExtension(listing: ISignalCExtensionListing, services: ISignalCExtensionInstallServices, enableInAgents: boolean): Promise<IExtension> {
	const vsix = await services.downloadService.download(listing.id);
	try {
		const contents = await services.fileService.readFile(vsix);
		const digest = await crypto.subtle.digest('SHA-256', new Uint8Array(contents.value.buffer));
		const actual = Array.from(new Uint8Array(digest), byte => byte.toString(16).padStart(2, '0')).join('');
		if (actual !== listing.sha256) {
			throw new Error(localize('signalC.catalog.integrityError', "The downloaded package for {0} did not match its published catalog checksum. Installation was stopped.", listing.name));
		}
		return await services.extensionsWorkbenchService.install(vsix, enableInAgents ? { enable: true } : undefined);
	} finally {
		try {
			if (await services.fileService.exists(vsix)) {
				await services.fileService.del(vsix);
			}
		} catch (error) {
			services.logService.warn('Unable to remove Signal-C Marketplace download', error);
		}
	}
}

interface ISignalCExtensionsMarketplaceOptions {
	readonly showSearch: boolean;
	readonly enableInAgents: boolean;
	readonly onDidInstall?: (extension: IExtension) => void;
}

interface IListingRow {
	readonly listing: Omit<ISignalCExtensionListing, 'asset' | 'sha256'>;
	readonly release?: ISignalCExtensionListing;
	readonly element: HTMLElement;
	readonly button: Button;
	readonly error: HTMLElement;
}

/** Shared first-party catalog for the Extensions sidebar and the Agents window. */
export class SignalCExtensionsMarketplace extends Disposable {
	readonly element: HTMLElement;
	private readonly rows: IListingRow[] = [];
	private readonly installing = new Set<string>();
	private readonly empty: HTMLElement;
	private query = '';

	constructor(
		container: HTMLElement,
		private readonly options: ISignalCExtensionsMarketplaceOptions,
		@IExtensionsWorkbenchService private readonly extensionsWorkbenchService: IExtensionsWorkbenchService,
		@IInstantiationService private readonly instantiationService: IInstantiationService,
		@IFileService private readonly fileService: IFileService,
		@ILogService private readonly logService: ILogService,
		@IOpenerService private readonly openerService: IOpenerService,
		@IContextViewService contextViewService: IContextViewService,
		@IExtensionGalleryManifestService galleryManifestService: IExtensionGalleryManifestService,
		@IExtensionGalleryService galleryService: IExtensionGalleryService,
	) {
		super();
		this.element = DOM.append(container, DOM.$('section.signal-c-extensions-marketplace'));
		this.element.setAttribute('aria-label', localize('signalC.catalog.title', "Signal-C Marketplace"));
		DOM.append(this.element, DOM.$('h3')).textContent = localize('signalC.catalog.title', "Signal-C Marketplace");
		DOM.append(this.element, DOM.$('p.signal-c-marketplace-description')).textContent = localize('signalC.catalog.sources', "Extensions from Signal-C GitHub releases and your local Signal-C gallery.");
		if (options.showSearch) {
			const search = this._register(new InputBox(this.element, contextViewService, {
				inputBoxStyles: defaultInputBoxStyles,
				placeholder: localize('signalC.catalog.search', "Search Signal-C extensions..."),
				ariaLabel: localize('signalC.catalog.search', "Search Signal-C extensions..."),
			}));
			this._register(search.onDidChange(query => this.setQuery(query)));
		}
		const list = DOM.append(this.element, DOM.$('div.signal-c-marketplace-list'));
		for (const listing of signalCExtensionCatalog) {
			this.addListing(list, listing, listing);
		}
		this.empty = DOM.append(list, DOM.$('p.signal-c-marketplace-empty', { role: 'status' }));
		this.empty.textContent = localize('signalC.catalog.empty', "No Signal-C extensions match your search.");
		this.empty.hidden = true;
		this._register(extensionsWorkbenchService.onChange(() => this.refresh()));
		void extensionsWorkbenchService.whenInitialized.then(() => {
			if (!this._store.isDisposed) {
				this.refresh();
			}
		}).catch(error => logService.error(error));
		this.refresh();
		const token = cancelOnDispose(this._store);
		void this.loadPrivateExtensions(list, galleryManifestService, galleryService, token).catch(error => {
			if (!token.isCancellationRequested) {
				logService.debug('Unable to query the local Signal-C Marketplace', error);
			}
		});
	}

	private async loadPrivateExtensions(list: HTMLElement, manifestService: IExtensionGalleryManifestService, galleryService: IExtensionGalleryService, token: CancellationToken): Promise<void> {
		const manifest = await manifestService.getExtensionGalleryManifest();
		const endpoint = manifest && getExtensionGalleryManifestResourceUri(manifest, ExtensionGalleryResourceType.ExtensionQueryService);
		if (!endpoint || token.isCancellationRequested) {
			return;
		}
		const uri = URI.parse(endpoint);
		if ((uri.scheme !== 'http' && uri.scheme !== 'https') || !isLocalhostAuthority(uri.authority)) {
			return;
		}
		const pager = await galleryService.query({ text: 'tag:signal-c' }, token);
		for (let page = 0; page * pager.pageSize < pager.total && !token.isCancellationRequested; page++) {
			const extensions = page === 0 ? pager.firstPage : await pager.getPage(page, token);
			if (token.isCancellationRequested || !extensions.length) {
				break;
			}
			for (const extension of extensions) {
				const id = extension.identifier.id.toLowerCase();
				if (extension.private && !this.rows.some(row => row.listing.id === id)) {
					this.addListing(list, {
						id,
						name: extension.displayName || extension.name,
						description: extension.description,
						version: extension.version,
						keywords: extension.tags,
						repository: extension.assets.repository?.uri ?? '',
					});
				}
			}
			this.setQuery(this.query);
			this.refresh();
		}
	}

	private addListing(list: HTMLElement, listing: IListingRow['listing'], release?: ISignalCExtensionListing): void {
		const element = DOM.append(list, DOM.$('article.signal-c-extension-listing'));
		const title = DOM.append(element, DOM.$('div.signal-c-extension-title'));
		DOM.append(title, DOM.$('strong')).textContent = listing.name;
		const button = this._register(new Button(title, { ...defaultButtonStyles, small: true }));
		DOM.append(element, DOM.$('p.signal-c-extension-description')).textContent = listing.description;
		const details = DOM.append(element, DOM.$('div.signal-c-extension-details'));
		DOM.append(details, DOM.$('span')).textContent = `${listing.id} · v${listing.version}`;
		if (release) {
			const url = getSignalCExtensionRepository(release);
			const source = DOM.append(details, DOM.$('a', { href: url }));
			source.textContent = localize('signalC.catalog.source', "Source on GitHub");
			this._register(DOM.addDisposableListener(source, DOM.EventType.CLICK, event => {
				event.preventDefault();
				void this.openerService.open(URI.parse(url)).catch(error => this.logService.error(error));
			}));
		}
		const error = DOM.append(element, DOM.$('p.signal-c-extension-error', { role: 'status' }));
		error.hidden = true;
		const row = { listing, release, element, button, error };
		this.rows.push(row);
		this._register(button.onDidClick(() => { void this.install(row); }));
	}

	setQuery(query: string): void {
		this.query = query;
		for (const row of this.rows) {
			row.element.hidden = !matchesSignalCExtension(row.listing, query);
		}
		this.empty.hidden = this.rows.some(row => !row.element.hidden);
	}

	private refresh(): void {
		for (const row of this.rows) {
			const installed = this.extensionsWorkbenchService.local.find(extension => extension.local && extension.identifier.id.toLowerCase() === row.listing.id);
			const busy = this.installing.has(row.listing.id);
			const current = installed && gte(installed.version, row.listing.version);
			row.button.label = isWeb ? localize('signalC.catalog.desktopOnly', "Desktop Only") : busy ? localize('signalC.catalog.installing', "Installing...")
				: current ? localize('signalC.catalog.installed', "Installed")
					: installed ? localize('signalC.catalog.update', "Update") : localize('signalC.catalog.install', "Install");
			row.button.element.setAttribute('aria-label', localize('signalC.catalog.actionLabel', "{0}: {1}", row.button.label, row.listing.name));
			row.button.enabled = !isWeb && !busy && !current;
		}
	}

	private async installFromGallery(id: string): Promise<IExtension | undefined> {
		const local = this.extensionsWorkbenchService.local.find(extension => extension.local && extension.identifier.id.toLowerCase() === id);
		if (!local) {
			return this.extensionsWorkbenchService.install(id, this.options.enableInAgents ? { enable: true } : undefined);
		}
		// A VSIX installation may still be synchronizing its gallery metadata.
		const [extension] = await this.extensionsWorkbenchService.getExtensions([local.identifier], CancellationToken.None);
		if (!extension?.gallery || !extension.outdated || extension.state !== ExtensionState.Installed || extension.deprecationInfo || await this.extensionsWorkbenchService.canInstall(extension) !== true) {
			await this.extensionsWorkbenchService.open(local);
			return;
		}
		const update = this.instantiationService.createInstance(UpdateAction, false);
		try {
			update.extension = extension;
			await update.run();
		} finally {
			update.dispose();
		}
	}

	private async install(row: IListingRow): Promise<void> {
		if (!row.button.enabled || this.installing.has(row.listing.id)) {
			return;
		}
		this.installing.add(row.listing.id);
		row.error.hidden = true;
		this.refresh();
		try {
			const installed = row.release ? await installSignalCExtension(row.release, {
				downloadService: this.instantiationService.invokeFunction(accessor => accessor.get(ISignalCExtensionDownloadService)),
				fileService: this.fileService,
				extensionsWorkbenchService: this.extensionsWorkbenchService,
				logService: this.logService,
			}, this.options.enableInAgents) : await this.installFromGallery(row.listing.id);
			if (installed && !this._store.isDisposed) {
				this.options.onDidInstall?.(installed);
			}
		} catch (error) {
			this.logService.error(error);
			if (!this._store.isDisposed) {
				row.error.textContent = getErrorMessage(error);
				row.error.hidden = false;
			}
		} finally {
			this.installing.delete(row.listing.id);
			if (!this._store.isDisposed) {
				this.refresh();
			}
		}
	}
}

/** The first-party marketplace remains available even when no external gallery is configured. */
export class SignalCExtensionsMarketplaceView extends ViewPane {
	static readonly ID = 'workbench.views.extensions.signalCMarketplace';
	private marketplace: SignalCExtensionsMarketplace | undefined;
	private query = '';

	protected override renderBody(container: HTMLElement): void {
		super.renderBody(container);
		this.marketplace = this._register(this.instantiationService.createInstance(SignalCExtensionsMarketplace, container, { showSearch: false, enableInAgents: false }));
		this.marketplace.setQuery(this.query);
	}

	protected override layoutBody(height: number, width: number): void {
		super.layoutBody(height, width);
		if (this.marketplace) {
			this.marketplace.element.style.height = `${height}px`;
		}
	}

	setQuery(query: string): void {
		this.query = query;
		this.marketplace?.setQuery(query);
	}
}
