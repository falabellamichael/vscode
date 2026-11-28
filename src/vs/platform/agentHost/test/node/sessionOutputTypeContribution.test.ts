/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import assert from 'assert';
import { URI } from '../../../../base/common/uri.js';
import { ensureNoDisposablesAreLeakedInTestSuite } from '../../../../base/test/common/utils.js';
import type { IAgentHostChatContributionContext, IOutgoingTurn } from '../../common/agentHostChatContributionsService.js';
import { SESSION_OUTPUT_TYPE_META_KEY } from '../../common/meta/agentSessionOutputMeta.js';
import { MessageKind } from '../../common/state/sessionState.js';
import { SessionOutputTypeContribution } from '../../node/chatContributions/sessionOutputType/sessionOutputTypeContribution.js';

const stubContext: IAgentHostChatContributionContext = {
	contributionId: SessionOutputTypeContribution.id,
	memento: () => { throw new Error('memento is not used'); },
	deleteMemento: () => { throw new Error('memento is not used'); },
};

function turnWithMeta(meta: Record<string, unknown> | undefined, workingDirectories?: readonly URI[]): IOutgoingTurn {
	return {
		session: 'agenthost:/session/1',
		chat: 'agenthost:/session/1/chat/1',
		turnId: 'turn-1',
		message: { text: 'make something', origin: { kind: MessageKind.User }, ...(meta ? { _meta: meta } : {}) },
		workingDirectories,
	};
}

suite('SessionOutputTypeContribution', () => {
	ensureNoDisposablesAreLeakedInTestSuite();

	test('messages without output metadata contribute nothing', () => {
		const contribution = new SessionOutputTypeContribution(stubContext);
		try {
			assert.strictEqual(contribution.onOutgoingTurn(turnWithMeta(undefined)), undefined);
			assert.strictEqual(contribution.onOutgoingTurn(turnWithMeta({})), undefined);
			assert.strictEqual(contribution.onOutgoingTurn(turnWithMeta({ [SESSION_OUTPUT_TYPE_META_KEY]: 'auto' })), undefined);
			assert.strictEqual(contribution.onOutgoingTurn(turnWithMeta({ [SESSION_OUTPUT_TYPE_META_KEY]: 'not-a-kind' })), undefined);
		} finally {
			contribution.dispose();
		}
	});

	test('a workspace turn steers toward producing files', () => {
		const contribution = new SessionOutputTypeContribution(stubContext);
		try {
			const result = contribution.onOutgoingTurn(turnWithMeta(
				{ [SESSION_OUTPUT_TYPE_META_KEY]: 'website' },
				[URI.file('/workspace/app')],
			));
			const instruction = result?.instructions?.join('\n') ?? '';
			assert.ok(instruction.includes('Website'));
			assert.ok(instruction.includes('working directory'));
		} finally {
			contribution.dispose();
		}
	});

	test('a workspace-less turn steers toward an inline deliverable', () => {
		const contribution = new SessionOutputTypeContribution(stubContext);
		try {
			const result = contribution.onOutgoingTurn(turnWithMeta({ [SESSION_OUTPUT_TYPE_META_KEY]: 'document' }));
			const instruction = result?.instructions?.join('\n') ?? '';
			assert.ok(instruction.includes('Document'));
			assert.ok(instruction.includes('inline'));
		} finally {
			contribution.dispose();
		}
	});

	test('every explicit kind produces kind-named guidance', () => {
		const contribution = new SessionOutputTypeContribution(stubContext);
		try {
			for (const [kind, label] of [['mobile', 'Mobile'], ['design', 'Design'], ['slides', 'Slide deck'], ['animation', 'Animation'], ['dataVisualization', 'Data Visualization'], ['image', 'Image'], ['spreadsheet', 'Spreadsheet'], ['mission', 'Mission']] as const) {
				const result = contribution.onOutgoingTurn(turnWithMeta({ [SESSION_OUTPUT_TYPE_META_KEY]: kind }));
				assert.ok(result?.instructions?.[0]?.includes(label), `missing ${label} guidance`);
			}
		} finally {
			contribution.dispose();
		}
	});
});
