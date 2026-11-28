/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { URI } from '../../../base/common/uri.js';
import { localize } from '../../../nls.js';
import { createDecorator } from '../../instantiation/common/instantiation.js';

export const ISignalCExtensionDownloadService = createDecorator<ISignalCExtensionDownloadService>('signalCExtensionDownloadService');
export const signalCExtensionDownloadChannel = 'signalCExtensionDownloads';

export interface ISignalCExtensionDownloadService {
	readonly _serviceBrand: undefined;
	/** Download an allowlisted catalog entry to a new temporary package on the local desktop. */
	download(id: string): Promise<URI>;
}

export interface ISignalCExtensionListing {
	readonly id: string;
	readonly name: string;
	readonly description: string;
	readonly version: string;
	readonly repository: string;
	readonly asset: string;
	readonly sha256: string;
	readonly keywords: readonly string[];
}

/** Published GitHub release packages. Add new first-party extensions here once their VSIX is released and verified. */
export const signalCExtensionCatalog: readonly ISignalCExtensionListing[] = [
	{
		id: 'simplesignal.simplesignal-provider',
		name: 'SimpleSignal AI Provider',
		description: localize('signalC.catalog.provider', "Connect local and cloud AI models to chat, with endpoint management and model discovery."),
		version: '1.0.1',
		repository: 'SimpleSignal-VS-Code-Extension',
		asset: 'simplesignal-provider-1.0.1.vsix',
		sha256: 'c25043385e591d4946035f699f08d6386eee2af238e9dd9339949493b93cde44',
		keywords: ['ai', 'chat', 'models', 'ollama', 'lemonade', 'provider'],
	},
	{
		id: 'falabella.simpletheme',
		name: 'SimpleTheme',
		description: localize('signalC.catalog.theme', "Design and customize editor themes with live color controls, syntax highlighting, and curated presets."),
		version: '1.0.22',
		repository: 'SimpleSignal-Theme-Studio',
		asset: 'simpletheme-1.0.22.vsix',
		sha256: '52d4a980d245ce2bc135e83379703fcc4a2c723e720bb729303eda31f7df2ff0',
		keywords: ['theme', 'studio', 'color', 'appearance', 'syntax'],
	},
	{
		id: 'falabella.simple-gradient-studio',
		name: 'SimpleGradient Studio',
		description: localize('signalC.catalog.gradient', "Create reusable gradients for apps, pages, and panels, with live previews and profile import and export."),
		version: '0.2.0',
		repository: 'simple-gradient-studio',
		asset: 'simple-gradient-studio-0.2.0.vsix',
		sha256: '2fe7150d4e2bfbbbd6454fc9acfea57520e105eddf9e13763f667b11d9b40947',
		keywords: ['gradient', 'studio', 'theme', 'color', 'design'],
	},
];

export function getSignalCExtensionRepository(listing: ISignalCExtensionListing): string {
	return `https://github.com/falabellamichael/${listing.repository}`;
}

export function getSignalCExtensionDownload(listing: ISignalCExtensionListing): string {
	return `${getSignalCExtensionRepository(listing)}/releases/download/v${listing.version}/${listing.asset}`;
}

export function matchesSignalCExtension(listing: Pick<ISignalCExtensionListing, 'id' | 'name' | 'description' | 'repository' | 'keywords'>, query: string): boolean {
	const text = [listing.id, listing.name, listing.description, listing.repository, ...listing.keywords].join(' ').toLowerCase();
	return query.toLowerCase().trim().split(/\s+/).every(word => text.includes(word));
}
