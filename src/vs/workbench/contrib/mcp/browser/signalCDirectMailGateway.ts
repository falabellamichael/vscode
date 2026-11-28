/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { IReader } from '../../../../base/common/observable.js';
import { StringSHA1 } from '../../../../base/common/hash.js';
import { MCP } from '../../../../platform/mcp/common/modelContextProtocol.js';
import { IMcpGatewayLocalServer } from '../common/mcpGatewayToolBrokerChannel.js';
import { ISignalCDirectMail, ISignalCDirectMailConnection } from './signalCDirectMail.js';

export function directMailGatewayId(connection: ISignalCDirectMailConnection): string {
	const hash = new StringSHA1();
	hash.update(connection.id);
	return `signal-c-mail-gateway:${connection.provider}:${hash.digest()}`;
}

const readMail: MCP.Tool = {
	name: 'read_mail',
	description: 'Search the connected mailbox or read a message by its returned ID. Use folder all to search across the mailbox.',
	annotations: { readOnlyHint: true },
	inputSchema: { type: 'object', properties: {
		folder: { type: 'string', enum: ['inbox', 'sent', 'all'] },
		query: { type: 'string', description: 'Search terms for this mailbox.' },
		messageId: { type: 'string' },
		limit: { type: 'number', minimum: 1, maximum: 10 },
	} },
};

const sendMail: MCP.Tool = {
	name: 'send_mail',
	description: 'Send a plain-text email from the connected mailbox. Call only for an explicit send request with a recipient. Check Sent before retrying an uncertain submission.',
	annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false },
	inputSchema: { type: 'object', required: ['to', 'subject', 'body'], properties: {
		to: { type: 'string', description: 'One complete recipient email address.' },
		subject: { type: 'string' },
		body: { type: 'string' },
	} },
};

export const outlookCalendarTool: MCP.Tool = {
	name: 'list_calendar_events',
	description: 'Read events in the connected Microsoft 365 default calendar for a specific time range. Returns the first 50 events.',
	annotations: { readOnlyHint: true },
	inputSchema: { type: 'object', required: ['start', 'end'], properties: {
		start: { type: 'string', description: 'Start of range, as ISO 8601 with a time zone.' },
		end: { type: 'string', description: 'End of range, as ISO 8601 with a time zone.' },
	} },
};

export function getDirectMailGatewayServers(service: ISignalCDirectMail, reader: IReader | undefined): readonly IMcpGatewayLocalServer[] {
	return service.connections.read(reader).filter(connection => !connection.needsSignIn).map(connection => ({
		id: directMailGatewayId(connection),
		label: `signal-c-${connection.provider}`,
		requiresExplicitSelection: true,
		tools: [readMail, ...(connection.canSend ? [sendMail] : []), ...(connection.canReadCalendar ? [outlookCalendarTool] : [])],
		async callTool(name, args, token) {
			let result: string;
			if (name === readMail.name) {
				result = await service.read(connection, args, token);
			} else if (name === sendMail.name && connection.canSend) {
				result = await service.send(connection, args, token);
			} else if (name === outlookCalendarTool.name && connection.canReadCalendar) {
				result = await service.readCalendar(connection, args, token);
			} else {
				throw new Error(`Unknown mail or calendar tool: ${name}`);
			}
			return { content: [{ type: 'text', text: result }] };
		},
	}));
}
