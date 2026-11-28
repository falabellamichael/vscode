/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import assert from 'assert';
import { VSBuffer } from '../../../../../base/common/buffer.js';
import { CancellationToken } from '../../../../../base/common/cancellation.js';
import { ResourceSet } from '../../../../../base/common/map.js';
import { constObservable } from '../../../../../base/common/observable.js';
import { URI } from '../../../../../base/common/uri.js';
import { upcastPartial } from '../../../../../base/test/common/mock.js';
import { ensureNoDisposablesAreLeakedInTestSuite } from '../../../../../base/test/common/utils.js';
import { TestConfigurationService } from '../../../../../platform/configuration/test/common/testConfigurationService.js';
import { FileService } from '../../../../../platform/files/common/fileService.js';
import { InMemoryFileSystemProvider } from '../../../../../platform/files/common/inMemoryFilesystemProvider.js';
import { NullLogService } from '../../../../../platform/log/common/log.js';
import { IWorkspaceContextService } from '../../../../../platform/workspace/common/workspace.js';
import { IWorkspaceTrustManagementService } from '../../../../../platform/workspace/common/workspaceTrust.js';
import { testWorkspace } from '../../../../../platform/workspace/test/common/testWorkspace.js';
import { IAgentPluginService } from '../../../../../workbench/contrib/chat/common/plugins/agentPluginService.js';
import { PromptsConfig } from '../../../../../workbench/contrib/chat/common/promptSyntax/config/config.js';
import { PromptFileParser } from '../../../../../workbench/contrib/chat/common/promptSyntax/promptFileParser.js';
import { PromptsType } from '../../../../../workbench/contrib/chat/common/promptSyntax/promptTypes.js';
import { IAgentSkill, IPromptsService, PromptsStorage } from '../../../../../workbench/contrib/chat/common/promptSyntax/service/promptsService.js';
import { PromptsService } from '../../../../../workbench/contrib/chat/common/promptSyntax/service/promptsServiceImpl.js';
import { ASSISTANT_CUSTOM_VIEW_ID } from '../../../../../workbench/contrib/sessions/common/chatConstants.js';
import { ISignalCDirectMail } from '../../../../../workbench/contrib/mcp/browser/signalCDirectMail.js';
import { IMcpService } from '../../../../../workbench/contrib/mcp/common/mcpTypes.js';
import { TestPathService, workbenchInstantiationService } from '../../../../../workbench/test/browser/workbenchTestServices.js';
import { TestContextService, TestWorkspaceTrustManagementService } from '../../../../../workbench/test/common/workbenchTestServices.js';
import { IAssistantAgentChats } from '../../browser/assistantAgentChats.js';
import { ASSISTANT_FIND_SKILL_TOOL, listAssistantSkills, readAssistantSkillReference } from '../../browser/assistantSkills.js';
import { collectAssistantTools, getAssistantSkillScope, getAssistantWorkspaceFolders, IAssistantToolContext } from '../../browser/assistantTools.js';

suite('AssistantSkills', () => {
	const disposables = ensureNoDisposablesAreLeakedInTestSuite();
	const root = URI.parse('file:///selected');
	const token = CancellationToken.None;
	test('keeps workflow execution and mixed actions outside named skill reference questions', () => {
		assert.strictEqual(readAssistantSkillReference('What does the installed `meeting-follow-up` skill say about missing owners?'), 'meeting-follow-up');
		assert.strictEqual(readAssistantSkillReference('What does the meeting-follow-up skill say about owners and deadlines'), 'meeting-follow-up');
		assert.strictEqual(readAssistantSkillReference('According to the meeting-follow-up skill, how should I record owners and deadlines?'), 'meeting-follow-up');
		for (const prompt of [
			'Use the meeting-follow-up skill and email the result.',
			'Read the installed meeting-follow-up skill. What does it say about owners? Email the attendees.',
			'Read the installed meeting-follow-up skill. What does it say about owners and then email the team?',
			'What does the installed meeting-follow-up skill say about deadlines and email it to Mira?',
			'According to the meeting-follow-up skill, how should I record a deadline and schedule it',
			'What does the installed ../secret skill say about owners?',
		]) {
			assert.strictEqual(readAssistantSkillReference(prompt), undefined);
		}
	});
	const skill = (name: string, overrides: Partial<IAgentSkill> = {}): IAgentSkill => ({
		uri: URI.joinPath(root, 'skills', name, 'SKILL.md'), name, description: 'Help with meeting notes',
		storage: PromptsStorage.local, disableModelInvocation: false, userInvocable: true, ...overrides,
	});

	function catalog(promptsService: IPromptsService, overrides: Partial<IAssistantToolContext> = {}) {
		return collectAssistantTools(upcastPartial<IAssistantToolContext>({
			accountId: 'fixture', agentChatTarget: { fileRoots: [root] }, mcpService: upcastPartial<IMcpService>({ servers: constObservable([]) }), promptsService, ...overrides,
			workspaceService: overrides.workspaceService ?? new TestContextService(testWorkspace(...(overrides.agentChatTarget?.fileRoots ?? [root]))),
		}));
	}

	test('discovers skills in an unopened worktree and restores personal skills shadowed by the IDE workspace', async () => {
		const open = URI.parse('file:///ide');
		const home = URI.parse('file:///home');
		const workspace = new TestContextService(testWorkspace(open));
		const config = new TestConfigurationService({ [PromptsConfig.USE_AGENT_SKILLS]: true, [PromptsConfig.SKILLS_LOCATION_KEY]: {}, [PromptsConfig.USE_CUSTOMIZATIONS_IN_PARENT_REPOS]: true });
		disposables.add(config.onDidChangeConfigurationEmitter);
		const files = disposables.add(new FileService(new NullLogService()));
		disposables.add(files.registerProvider(root.scheme, disposables.add(new InMemoryFileSystemProvider())));
		const instantiation = workbenchInstantiationService({ fileService: () => files, configurationService: () => config, pathService: () => new TestPathService(home) }, disposables);
		instantiation.stub(IWorkspaceContextService, workspace);
		instantiation.stub(IWorkspaceTrustManagementService, disposables.add(new TestWorkspaceTrustManagementService(true, new ResourceSet([root]))));
		instantiation.stub(IAgentPluginService, { plugins: constObservable([]) });
		const write = async (folder: URI, name: string, flags = '') => {
			const uri = URI.joinPath(folder, '.agents/skills', name, 'SKILL.md');
			await files.writeFile(uri, VSBuffer.fromString(`---\nname: ${name}\ndescription: From ${folder.path}\n${flags}---\nInstructions from ${folder.path}.`));
			return uri;
		};
		await write(open, 'ide-only');
		await write(open, 'fallback');
		const personal = await write(home, 'fallback');
		await write(home, 'review');
		const selected = await write(root, 'review');
		await write(root, 'interactive', 'disable-model-invocation: true\n');
		await write(root, 'automatic', 'user-invocable: false\n');
		const disabled = await write(root, 'disabled');
		await files.createFolder(URI.joinPath(root, '.git'));
		const prompts = disposables.add(instantiation.createInstance(PromptsService));
		prompts.setDisabledPromptFiles(PromptsType.skill, new ResourceSet([disabled]));
		const target = { fileRoots: [root] };
		let trusted = true;
		const chats = upcastPartial<IAssistantAgentChats>({ isResourceTrusted: async () => trusted });
		const scope = await getAssistantSkillScope(getAssistantWorkspaceFolders(workspace, target), workspace, chats);
		const manual = await listAssistantSkills(prompts, scope, 'user', token);
		assert.deepStrictEqual(manual.map(item => item.name).sort(), ['fallback', 'interactive', 'review']);
		assert.strictEqual(manual.find(item => item.name === 'fallback')?.uri.toString(), personal.toString());
		assert.strictEqual(manual.find(item => item.name === 'review')?.uri.toString(), selected.toString());
		const tool = catalog(prompts, { workspaceService: workspace, agentChats: chats, agentChatTarget: target }).entries.get(ASSISTANT_FIND_SKILL_TOOL);
		assert.ok(tool?.kind === 'builtin');
		assert.deepStrictEqual((await tool.run({}, token)).split('\n').slice(1), ['- automatic: From /selected', '- fallback: From /home', '- review: From /selected']);
		assert.match(await tool.run({ name: 'review' }, token), /Instructions from \/selected/);
		assert.match(await tool.run({ name: 'fallback' }, token), /Instructions from \/home/);

		// Parent repository discovery is enabled, but the chat selected only its subfolder.
		const nested = URI.joinPath(root, 'nested');
		await files.createFolder(nested);
		assert.ok((await prompts.listPromptFilesForStorage(PromptsType.skill, PromptsStorage.local, token, nested)).some(item => item.uri.toString() === selected.toString()));
		const nestedScope = await getAssistantSkillScope(getAssistantWorkspaceFolders(workspace, { fileRoots: [nested] }), workspace, chats);
		assert.strictEqual((await listAssistantSkills(prompts, nestedScope, 'user', token)).some(item => item.storage === PromptsStorage.local), false);
		trusted = false;
		assert.match(await tool.run({ name: 'automatic' }, token), /unavailable/);
		assert.match(await tool.run({ name: 'review' }, token), /Instructions from \/home/);
	});

	test('discovers eligible metadata on demand and only opens exact names within the captured target', async () => {
		const available = skill('review-notes', { userInvocable: false, sessionTypes: [ASSISTANT_CUSTOM_VIEW_ID] });
		const off = skill('disabled');
		const disabled = new ResourceSet([off.uri]);
		const skills = [available, off, skill('manually-invoked', { disableModelInvocation: true }), skill('other-session', { sessionTypes: ['other'] }),
			skill('elsewhere', { uri: URI.parse('file:///other/skills/elsewhere/SKILL.md') })];
		let discoveries = 0;
		const reads: string[] = [];
		const prompts = upcastPartial<IPromptsService>({
			findAgentSkills: async () => { discoveries++; return skills; },
			getDisabledPromptFiles: () => disabled,
			parseNew: async uri => { reads.push(uri.toString()); return new PromptFileParser().parse(uri, 'Use the supplied notes.'); },
		});
		const tool = catalog(prompts).entries.get(ASSISTANT_FIND_SKILL_TOOL);
		assert.ok(tool?.kind === 'builtin' && tool.safe);
		assert.strictEqual(discoveries, 0);
		assert.deepStrictEqual((await tool.run({ query: 'meeting' }, token)).split('\n').slice(1), ['- review-notes: Help with meeting notes']);
		for (const name of ['disabled', 'manually-invoked', 'other-session', 'elsewhere', '../SKILL.md']) {
			assert.ok(!(await tool.run({ name }, token)).includes('Use the supplied notes.'));
		}
		assert.deepStrictEqual(reads, []);
		assert.match(await tool.run({ name: available.name }, token), /Use the supplied notes\./);
		assert.deepStrictEqual(reads, [available.uri.toString()]);
	});

	test('pages long bodies and walks a filtered catalog without truncating later matches', async () => {
		const skills = Array.from({ length: 11 }, (_, i) => skill(`notes-${i + 10}`, { storage: PromptsStorage.builtIn }));
		const body = 'a'.repeat(4_000) + 'second page';
		let manualOnly = false;
		const disabled = new ResourceSet();
		const prompts = upcastPartial<IPromptsService>({
			findAgentSkills: async () => skills, getDisabledPromptFiles: () => disabled,
			parseNew: async uri => new PromptFileParser().parse(uri, `---\ndisable-model-invocation: ${manualOnly}\n---\n${body}`),
		});
		const tool = catalog(prompts, { agentChatTarget: { fileRoots: [] } }).entries.get(ASSISTANT_FIND_SKILL_TOOL);
		assert.ok(tool?.kind === 'builtin');
		assert.match(await tool.run({ query: 'meeting', page: 2 }, token), /notes-20/);
		const first = await tool.run({ name: 'notes-20' }, token);
		const second = await tool.run({ name: 'notes-20', page: 2 }, token);
		assert.strictEqual(first.split('\n').slice(1).join('\n') + second.split('\n').slice(1).join('\n'), body);
		assert.match(await tool.run({ name: 'notes-20', page: 3 }, token), /1 to 2/);
		manualOnly = true;
		assert.match(await tool.run({ name: 'notes-20' }, token), /unavailable/);
	});

	test('stays out of restricted runs and preserves connected-tool discovery in a full task catalog', async () => {
		const prompts = upcastPartial<IPromptsService>({});
		for (const mode of [{ connectedHistory: true }, { repositoryScope: 'workspace' as const }, { repositoryScope: 'github' as const }]) {
			assert.strictEqual(catalog(prompts, mode).entries.has(ASSISTANT_FIND_SKILL_TOOL), false);
		}
		const tools = catalog(prompts, {
			taskId: 'task', agentChats: upcastPartial({}), mailWatcher: upcastPartial({}), configuration: upcastPartial({}),
			readNewsFeeds: async () => [], readWeatherForecast: async () => { throw new Error('unused'); },
			connectApp: () => true, proposeRoutine: () => '', proposeMemory: () => '',
			directMail: upcastPartial<ISignalCDirectMail>({ connections: constObservable([
				{ id: 'gmail', provider: 'gmail', label: 'Gmail', needsSignIn: false },
				{ id: 'outlook', provider: 'outlook', label: 'Outlook', needsSignIn: false },
			]) }),
		});
		const discover = tools.entries.get('assistant.find_app_tools');
		assert.ok(discover?.kind === 'builtin');
		for (const name of [ASSISTANT_FIND_SKILL_TOOL, 'assistant.outlook_read']) {
			if (!tools.specs.some(tool => tool.name === name)) {
				assert.match(await discover.run({ name }, token), /Available in the next request/);
			}
			assert.ok(tools.specs.some(tool => tool.name === name));
			assert.ok(tools.specs.length <= 32);
		}
	});
});
