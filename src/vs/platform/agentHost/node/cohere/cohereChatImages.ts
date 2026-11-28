/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { localize } from '../../../../nls.js';
import type { IFileService } from '../../../files/common/files.js';
import { resolveChatImages, type ChatImageAttachment, type SavedChatImage } from '../chatImageAttachments.js';

export { isChatImage as isCohereImage, type SavedChatImage as SavedCohereImage } from '../chatImageAttachments.js';

export function cohereSupportsImages(model: string): boolean {
	return model.startsWith('command-a-plus-') || model.includes('vision');
}

export function resolveCohereImages(files: Pick<IFileService, 'readFile'>, current: readonly ChatImageAttachment[], history: readonly { images?: readonly SavedChatImage[] }[]) {
	return resolveChatImages(files, current, history, {
		maxBytes: 20_000_000,
		maxImages: 20,
		limitMessage: localize('cohere.image.limit', "Cohere accepts up to 20 images and 20 MB of image data per request."),
		localFileMessage: localize('cohere.image.local', "Attach a local image file to send it to Cohere."),
		formatMessage: label => localize('cohere.image.format', "Cohere accepts PNG, JPEG, WebP and non-animated GIF images. Check the format of {0}.", label),
	});
}
