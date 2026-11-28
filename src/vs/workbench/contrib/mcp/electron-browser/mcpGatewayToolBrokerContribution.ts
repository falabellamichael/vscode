/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { IWorkbenchContribution } from '../../../common/contributions.js';
import { IMainProcessService } from '../../../../platform/ipc/common/mainProcessService.js';
import { ILogService } from '../../../../platform/log/common/log.js';
import { derived } from '../../../../base/common/observable.js';
import { McpGatewayToolBrokerChannelName } from '../../../../platform/mcp/common/mcpGateway.js';
import { IMcpService } from '../common/mcpTypes.js';
import { McpGatewayToolBrokerChannel } from '../common/mcpGatewayToolBrokerChannel.js';
import { ISignalCDirectMail } from '../browser/signalCDirectMail.js';
import { getDirectMailGatewayServers } from '../browser/signalCDirectMailGateway.js';

export class McpGatewayToolBrokerContribution implements IWorkbenchContribution {
	constructor(
		@IMainProcessService mainProcessService: IMainProcessService,
		@IMcpService mcpService: IMcpService,
		@ILogService logService: ILogService,
		@ISignalCDirectMail directMail: ISignalCDirectMail,
	) {
		const localServers = derived(reader => getDirectMailGatewayServers(directMail, reader));
		mainProcessService.registerChannel(McpGatewayToolBrokerChannelName, new McpGatewayToolBrokerChannel(mcpService, logService, undefined, localServers));
	}
}
