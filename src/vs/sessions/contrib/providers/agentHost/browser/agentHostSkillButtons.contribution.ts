/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { registerWorkbenchContribution2, WorkbenchPhase } from '../../../../../workbench/common/contributions.js';
import { AGENT_HOST_SKILL_BUTTONS, IsAgentHostSessionContextContribution, registerAgentHostSkillButton } from '../../../../../workbench/contrib/sessions/providers/agentHost/browser/agentHostSkillButtons.js';

// Browser tool registration consults this key while the first chat is restoring.
registerWorkbenchContribution2(IsAgentHostSessionContextContribution.ID, IsAgentHostSessionContextContribution, WorkbenchPhase.BlockRestore);

for (const spec of AGENT_HOST_SKILL_BUTTONS) {
	registerAgentHostSkillButton(spec);
}
