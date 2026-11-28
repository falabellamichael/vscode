/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { createHash } from 'crypto';
import { detectImageMimeType } from '../../../base/common/image.js';
import { getMediaMime } from '../../../base/common/mime.js';
import { URI } from '../../../base/common/uri.js';
import type { IFileService } from '../../files/common/files.js';
import { MessageAttachmentKind, type MessageAttachment } from '../common/state/sessionState.js';

export type ChatImageAttachment = Extract<MessageAttachment, { type: MessageAttachmentKind.Resource | MessageAttachmentKind.EmbeddedResource }>;

export interface SavedChatImage {
	readonly attachment: ChatImageAttachment;
	readonly bytes: number;
	readonly digest: string;
}

interface ChatImagePolicy {
	readonly maxImages: number;
	readonly maxBytes: number;
	readonly limitMessage: string;
	readonly localFileMessage: string;
	readonly formatMessage: (label: string) => string;
}

export function isChatImage(attachment: MessageAttachment): attachment is ChatImageAttachment {
	return (attachment.type === MessageAttachmentKind.Resource || attachment.type === MessageAttachmentKind.EmbeddedResource)
		&& attachment.displayKind !== 'directory' && (attachment.displayKind === 'image' || attachment.contentType?.startsWith('image/') === true || getMediaMime(attachment.label)?.startsWith('image/') === true);
}

/** Give current attachments priority. Replay a historical local file only if its original bytes still match. */
export async function resolveChatImages(files: Pick<IFileService, 'readFile'>, current: readonly ChatImageAttachment[], history: readonly { images?: readonly SavedChatImage[] }[], policy: ChatImagePolicy) {
	let remainingBytes = policy.maxBytes, remainingImages = policy.maxImages;
	async function read(attachment: ChatImageAttachment, previous?: SavedChatImage) {
		if (!remainingImages || (previous && previous.bytes > remainingBytes)) { throw new Error(policy.limitMessage); }
		let bytes: Uint8Array;
		if (attachment.type === MessageAttachmentKind.EmbeddedResource) {
			if (attachment.data.length > 4 * Math.ceil(remainingBytes / 3)) { throw new Error(policy.limitMessage); }
			bytes = Buffer.from(attachment.data, 'base64');
		} else {
			const uri = URI.parse(attachment.uri);
			if (uri.scheme !== 'file') { throw new Error(policy.localFileMessage); }
			bytes = (await files.readFile(uri, { limits: { size: remainingBytes } })).value.buffer;
		}
		if (bytes.byteLength > remainingBytes) { throw new Error(policy.limitMessage); }
		const mime = detectImageMimeType(bytes);
		if (!mime) { throw new Error(policy.formatMessage(attachment.label)); }
		const digest = createHash('sha256').update(bytes).digest('hex');
		if (previous && previous.digest !== digest) { return undefined; }
		remainingBytes -= bytes.byteLength;
		remainingImages--;
		return { url: 'data:' + mime + ';base64,' + Buffer.from(bytes).toString('base64'), saved: { attachment, bytes: bytes.byteLength, digest } };
	}
	if (current.length > remainingImages) { throw new Error(policy.limitMessage); }
	const urls: string[] = [], saved: SavedChatImage[] = [];
	for (const attachment of current) {
		const image = await read(attachment);
		if (image) { urls.push(image.url); saved.push(image.saved); }
	}
	const earlier = new Map<number, string[]>();
	for (let i = history.length - 1; i >= 0 && remainingImages && remainingBytes; i--) {
		const images: string[] = [];
		for (const previous of history[i].images ?? []) {
			try {
				const image = await read(previous.attachment, previous);
				if (image) { images.push(image.url); }
			} catch { /* A moved or unavailable historical image must not block the current turn. */ }
		}
		if (images.length) { earlier.set(i, images); }
	}
	return { urls, saved, earlier };
}
