/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { CancellationToken } from '../../../../base/common/cancellation.js';
import { isCancellationError } from '../../../../base/common/errors.js';
import { URI } from '../../../../base/common/uri.js';
import { getSkillFolderName } from '../../../../workbench/contrib/chat/common/promptSyntax/config/promptFileLocations.js';
import { PromptsType } from '../../../../workbench/contrib/chat/common/promptSyntax/promptTypes.js';
import { IAgentSkill, IPromptFileDiscoveryResult, IPromptsService, matchesSessionType, PromptsStorage } from '../../../../workbench/contrib/chat/common/promptSyntax/service/promptsService.js';
import { sanitizeSlashCommandName } from '../../../../workbench/contrib/chat/common/promptSyntax/utils/promptsServiceUtils.js';
import { ASSISTANT_CUSTOM_VIEW_ID } from '../../../../workbench/contrib/sessions/common/chatConstants.js';

export const ASSISTANT_FIND_SKILL_TOOL = 'assistant.find_skill';
const LIST_PAGE_SIZE = 8;
const BODY_PAGE_SIZE = 4_000;

/** Recognize a standalone question about the contents of an explicitly named skill. */
export function readAssistantSkillReference(request: string): string | undefined {
	const text = request.trim();
	const match = /^(?:please\s+)?read\s+the\s+(?:installed\s+)?`?([\w-]{1,64})`?\s+skill[.!]\s*what\s+does\s+(?:it|the skill)\s+(?:say|recommend)\b([^?!.;\n]*)\??$/i.exec(text)
		?? /^what\s+does\s+the\s+(?:installed\s+)?`?([\w-]{1,64})`?\s+skill\s+(?:say|recommend)\b([^?!.;\n]*)\??$/i.exec(text)
		?? /^according\s+to\s+the\s+(?:installed\s+)?`?([\w-]{1,64})`?\s+skill,\s*(?:how|what|when|which|why|where|who)\b([^?!.;\n]*)\??$/i.exec(text);
	if (!match) { return undefined; }
	// Leave compound requests to the normal dispatcher so their actions remain available.
	const followUp = /\b(?:and\s+(?:then|also|please)|(?:can|could|would|will)\s+you)\b|(?:,\s*|\b(?:and|then|also)\s+)(?:send|e-?mail|schedule|book|create|draft|write|set|add|update|change|delete|remove|archive|publish|post|run|use|apply|check|find|search|open|tell|give)\b/i;
	return followUp.test(match[2]) ? undefined : match[1];
}

/** Only a successful exact read produces this header and a nonempty page body. */
export function assistantSkillPageCount(name: string, result: string): number | undefined {
	const header = /^Skill ([\w-]+) \(page 1 of (\d+)\)\.[^\n]*\n[\s\S]/.exec(result);
	return header?.[1] === name ? Number(header[2]) : undefined;
}

export interface IAssistantSkillScope {
	/** Omitted when the workbench already discovers exactly the captured roots. */
	readonly roots?: readonly URI[];
	readonly inWorkspace: (uri: URI) => boolean;
}

type AssistantSkill = Pick<IAgentSkill, 'uri' | 'name' | 'storage' | 'extension' | 'pluginLabel'> & { readonly description?: string };

export async function listAssistantSkills(prompts: IPromptsService, scope: IAssistantSkillScope, invocation: 'user' | 'model', token: CancellationToken): Promise<readonly AssistantSkill[]> {
	const disabled = prompts.getDisabledPromptFiles(PromptsType.skill);
	const eligible = (skill: Pick<IAgentSkill, 'sessionTypes' | 'storage' | 'uri'>) => !disabled.has(skill.uri)
		&& matchesSessionType(skill.sessionTypes, ASSISTANT_CUSTOM_VIEW_ID) && (skill.storage !== PromptsStorage.local || scope.inWorkspace(skill.uri));
	if (!scope.roots) {
		if (invocation === 'user') {
			const commands = await prompts.getPromptSlashCommands(token);
			return commands.filter(command => command.type === PromptsType.skill && command.userInvocable && eligible(command));
		}
		return (await prompts.findAgentSkills(token) ?? []).filter(skill => !skill.disableModelInvocation && eligible(skill));
	}

	// The global catalog already discarded duplicate names. Use its ordered discovery
	// records so a skill shadowed by the IDE workspace can win in the selected worktree.
	const [localLists, global] = await Promise.all([
		Promise.all(scope.roots.map(root => prompts.listPromptFilesForStorage(PromptsType.skill, PromptsStorage.local, token, root))),
		prompts.getDiscoveryInfo(invocation === 'user' ? PromptsType.prompt : PromptsType.skill, token),
	]);
	const candidates: IPromptFileDiscoveryResult[] = localLists.flat().map(promptPath => ({ status: 'loaded', promptPath }));
	candidates.push(...global.files.filter(file => file.promptPath.storage !== PromptsStorage.local && file.promptPath.type === PromptsType.skill
		&& (file.status === 'loaded' || file.skipReason === 'duplicate-name')));
	const seen = new Set<string>();
	const skills: AssistantSkill[] = [];
	for (const candidate of candidates) {
		const path = candidate.promptPath;
		if (!eligible(path)) { continue; }
		const rawName = path.name ?? getSkillFolderName(path.uri);
		const name = invocation === 'user' ? sanitizeSlashCommandName(rawName) : rawName;
		if (seen.has(name)) { continue; }
		let { userInvocable, disableModelInvocation } = candidate;
		let description = path.description;
		if (path.storage === PromptsStorage.local || candidate.skipReason === 'duplicate-name') {
			try {
				const { header } = await prompts.parseNew(path.uri, token);
				description = header?.description ?? description;
				userInvocable = header?.userInvocable;
				disableModelInvocation = header?.disableModelInvocation;
			} catch (error) {
				if (isCancellationError(error)) { throw error; }
				continue;
			}
		}
		seen.add(name);
		if (invocation === 'user' ? userInvocable !== false : !disableModelInvocation) {
			skills.push({ uri: path.uri, name, storage: path.storage, description: invocation === 'model' ? description?.replace(/<[^>]+>/g, '').slice(0, 1024) : description, pluginLabel: path.pluginLabel, extension: path.extension });
		}
	}
	return skills;
}

/** Resolves only names from the current eligible catalog. The caller never supplies a URI. */
export async function findAssistantSkill(prompts: IPromptsService, scope: IAssistantSkillScope, args: Record<string, unknown>, token: CancellationToken): Promise<string> {
	const name = typeof args.name === 'string' ? args.name.trim() : '';
	const query = typeof args.query === 'string' ? args.query.trim() : '';
	const page = args.page ?? 1;
	if (typeof page !== 'number' || !Number.isSafeInteger(page) || page < 1) { return 'Use a positive integer for page.'; }
	if (name.length > 64 || /[\\/:]/.test(name) || name === '.' || name === '..' || query.length > 120 || (name && query)) {
		return 'Use a search query up to 120 characters or one exact skill name from the list.';
	}
	const skills = await listAssistantSkills(prompts, scope, 'model', token);
	if (name) {
		const skill = skills.find(skill => skill.name === name);
		if (!skill) { return 'That skill is unavailable in this Assistant chat.'; }
		const parsed = await prompts.parseNew(skill.uri, token);
		if (parsed.header?.disableModelInvocation || prompts.getDisabledPromptFiles(PromptsType.skill).has(skill.uri)) {
			return 'That skill is unavailable in this Assistant chat.';
		}
		const body = parsed.body?.getContent().trim();
		if (!body) { return 'This skill has no instructions.'; }
		const pages = Math.ceil(body.length / BODY_PAGE_SIZE);
		if (page > pages) { return `Use a page from 1 to ${pages}.`; }
		const next = page < pages ? ` Request page ${page + 1} to continue.` : '';
		return `Skill ${skill.name} (page ${page} of ${pages}).${next}\n${body.slice((page - 1) * BODY_PAGE_SIZE, page * BODY_PAGE_SIZE)}`;
	}
	const terms = query.toLowerCase().split(/[\s-]+/).filter(Boolean);
	const matches = skills.filter(skill => terms.every(term => `${skill.name} ${skill.description ?? ''}`.toLowerCase().includes(term)))
		.sort((a, b) => a.name.localeCompare(b.name));
	if (!matches.length) { return 'No matching skills are available in this Assistant chat.'; }
	const pages = Math.ceil(matches.length / LIST_PAGE_SIZE);
	if (page > pages) { return `Use a page from 1 to ${pages}.`; }
	const results = matches.slice((page - 1) * LIST_PAGE_SIZE, page * LIST_PAGE_SIZE).map(skill => {
		const description = skill.description?.replace(/\s+/g, ' ').trim().slice(0, 240);
		return description ? `- ${skill.name}: ${description}` : `- ${skill.name}`;
	});
	return `Available skills (page ${page} of ${pages}). Pass an exact name to read its instructions.\n${results.join('\n')}`;
}
