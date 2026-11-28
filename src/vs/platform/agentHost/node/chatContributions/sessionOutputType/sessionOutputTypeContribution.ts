/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { Disposable } from '../../../../../base/common/lifecycle.js';
import type { IAgentHostChatContribution, IAgentHostChatContributionContext, IOutgoingTurn, ISendContribution } from '../../../common/agentHostChatContributionsService.js';
import { getSessionOutputKindFromMetadata, type SessionOutputKind } from '../../../common/meta/agentSessionOutputMeta.js';

const SESSION_OUTPUT_KIND_LABELS: Record<Exclude<SessionOutputKind, 'auto'>, string> = {
	website: 'Website',
	mobile: 'Mobile',
	design: 'Design',
	slides: 'Slide deck',
	animation: 'Animation',
	dataVisualization: 'Data Visualization',
	image: 'Image',
	document: 'Document',
	spreadsheet: 'Spreadsheet',
	mission: 'Mission',
};

const SESSION_OUTPUT_KIND_GUIDANCE: Record<Exclude<SessionOutputKind, 'auto'>, { files: string; inline: string }> = {
	website: {
		files: 'Build a website or web application as the primary deliverable: create the project files in the working directory, include a way to run or preview it, and verify it works before reporting.',
		inline: 'Provide a complete website or web application implementation inline, with all project files and a way to run or preview it.',
	},
	mobile: {
		files: 'Build a mobile app experience as the primary deliverable: create the project files in the working directory and include a way to run or preview it.',
		inline: 'Provide a complete mobile app experience implementation inline, with the project files needed to run or preview it.',
	},
	design: {
		files: 'Create an implementation-ready design as the primary deliverable, saved as files covering layout, components, styling, and any needed assets.',
		inline: 'Provide an implementation-ready design inline, covering layout, components, styling, and any needed assets.',
	},
	slides: {
		files: 'Create a slide deck as the primary deliverable, saved as a file the user can open: clear slide structure, concise content, and speaker notes where useful.',
		inline: 'Provide a complete slide deck inline, with a clear slide structure, concise content, and speaker notes where useful.',
	},
	animation: {
		files: 'Create an animation or animated asset as the primary deliverable, saved as files in the working directory.',
		inline: 'Provide an animation or animated asset as the primary deliverable.',
	},
	dataVisualization: {
		files: 'Create a data visualization or dashboard as the primary deliverable: save the data model and the files needed to render it in the working directory.',
		inline: 'Provide a data visualization or dashboard as the primary deliverable, including the data model and the code needed to render it.',
	},
	image: {
		files: 'Create or edit an image or visual asset as the primary deliverable, saved as files in the working directory.',
		inline: 'Create or describe an image or visual asset as the primary deliverable.',
	},
	document: {
		files: 'Create a polished document or report as the primary deliverable, saved as a file the user can open and edit.',
		inline: 'Provide a polished document or report as the primary deliverable.',
	},
	spreadsheet: {
		files: 'Create a spreadsheet, table, or data model as the primary deliverable, saved as a file the user can open (for example CSV).',
		inline: 'Provide a spreadsheet, table, or data model as the primary deliverable.',
	},
	mission: {
		files: 'Treat the request as a multi-step mission: make a plan, execute it against the working directory, and report the resulting deliverable and remaining follow-ups.',
		inline: 'Treat the request as a multi-step mission: make a plan, execute it, and report the resulting deliverable and remaining follow-ups.',
	},
};

/** Elevates the user's output-picker choice on `message._meta` into provider steering, which survives paths that drop hidden prompt attachments. */
export class SessionOutputTypeContribution extends Disposable implements IAgentHostChatContribution {
	static readonly id = 'sessionOutputType';
	readonly order = 350;

	constructor(
		protected readonly _context: IAgentHostChatContributionContext,
	) {
		super();
	}

	onOutgoingTurn(turn: IOutgoingTurn): ISendContribution | undefined {
		const kind = getSessionOutputKindFromMetadata(turn.message._meta);
		if (!kind || kind === 'auto') {
			return undefined;
		}
		const hasWorkspace = !!turn.workingDirectories?.length;
		const guidance = SESSION_OUTPUT_KIND_GUIDANCE[kind];
		return {
			instructions: [`<requested_output>\nThe user selected the "${SESSION_OUTPUT_KIND_LABELS[kind]}" output type for this request. ${hasWorkspace ? guidance.files : guidance.inline}\nTreat it as the requested primary deliverable unless the request explicitly says otherwise.${hasWorkspace ? '' : ' No workspace folder is attached, so deliver the content inline rather than claiming to have saved files.'}\n</requested_output>`],
		};
	}
}
