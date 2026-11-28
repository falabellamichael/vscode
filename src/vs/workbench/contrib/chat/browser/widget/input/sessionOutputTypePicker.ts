/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { Codicon } from '../../../../../../base/common/codicons.js';
import { Disposable, toDisposable } from '../../../../../../base/common/lifecycle.js';
import { autorun, IObservable } from '../../../../../../base/common/observable.js';
import { ThemeIcon } from '../../../../../../base/common/themables.js';
import { localize } from '../../../../../../nls.js';
import { IActionWidgetService } from '../../../../../../platform/actionWidget/browser/actionWidget.js';
import { ActionListItemKind, IActionListDelegate, IActionListItem } from '../../../../../../platform/actionWidget/browser/actionList.js';
import { IStorageService, StorageScope, StorageTarget } from '../../../../../../platform/storage/common/storage.js';
import { IPromptTextVariableEntry } from '../../../common/attachments/chatVariableEntries.js';
import { isSessionOutputKind, SESSION_OUTPUT_CONTEXT_ID_PREFIX, type SessionOutputKind } from '../../../common/sessionOutputType.js';

export const SESSION_OUTPUT_TYPE_STORAGE_KEY = 'sessions.newChat.outputType';

interface ISessionOutputOption {
	readonly kind: SessionOutputKind;
	readonly label: string;
	readonly description: string;
	readonly icon: ThemeIcon;
	readonly instruction?: string;
}

export const SESSION_OUTPUT_OPTIONS: readonly ISessionOutputOption[] = [
	{
		kind: 'auto',
		label: localize('sessionOutputType.auto', "Output (Auto-Select)"),
		description: localize('sessionOutputType.autoDescription', "Let the agent choose the deliverable"),
		icon: Codicon.output,
	},
	{
		kind: 'website',
		label: localize('sessionOutputType.website', "Website"),
		description: localize('sessionOutputType.websiteDescription', "Build a website or web application"),
		icon: Codicon.globe,
		instruction: localize('sessionOutputType.websiteInstruction', "Build a website or web application as the primary deliverable, including the project files needed to run it."),
	},
	{
		kind: 'mobile',
		label: localize('sessionOutputType.mobile', "Mobile"),
		description: localize('sessionOutputType.mobileDescription', "Build a mobile app experience"),
		icon: Codicon.deviceMobile,
		instruction: localize('sessionOutputType.mobileInstruction', "Build a mobile app experience as the primary deliverable, including the project files needed to run or preview it."),
	},
	{
		kind: 'design',
		label: localize('sessionOutputType.design', "Design"),
		description: localize('sessionOutputType.designDescription', "Create an implementation-ready design"),
		icon: Codicon.paintcan,
		instruction: localize('sessionOutputType.designInstruction', "Create an implementation-ready design as the primary deliverable, covering layout, components, styling, and any needed assets."),
	},
	{
		kind: 'slides',
		label: localize('sessionOutputType.slides', "Slide deck"),
		description: localize('sessionOutputType.slidesDescription', "Create a presentation"),
		icon: Codicon.layoutPanel,
		instruction: localize('sessionOutputType.slidesInstruction', "Create a slide deck as the primary deliverable, with a clear slide structure, concise content, and speaker notes where useful."),
	},
	{
		kind: 'animation',
		label: localize('sessionOutputType.animation', "Animation"),
		description: localize('sessionOutputType.animationDescription', "Create an animated asset"),
		icon: Codicon.play,
		instruction: localize('sessionOutputType.animationInstruction', "Create an animation or animated asset as the primary deliverable."),
	},
	{
		kind: 'dataVisualization',
		label: localize('sessionOutputType.dataVisualization', "Data Visualization"),
		description: localize('sessionOutputType.dataVisualizationDescription', "Create a chart or dashboard"),
		icon: Codicon.graph,
		instruction: localize('sessionOutputType.dataVisualizationInstruction', "Create a data visualization or dashboard as the primary deliverable, with the data model and files needed to render it."),
	},
	{
		kind: 'image',
		label: localize('sessionOutputType.image', "Image"),
		description: localize('sessionOutputType.imageDescription', "Create a visual asset"),
		icon: Codicon.fileMedia,
		instruction: localize('sessionOutputType.imageInstruction', "Create or edit an image or visual asset as the primary deliverable."),
	},
	{
		kind: 'document',
		label: localize('sessionOutputType.document', "Document"),
		description: localize('sessionOutputType.documentDescription', "Write a document or report"),
		icon: Codicon.fileText,
		instruction: localize('sessionOutputType.documentInstruction', "Create a polished document or report as the primary deliverable."),
	},
	{
		kind: 'spreadsheet',
		label: localize('sessionOutputType.spreadsheet', "Spreadsheet"),
		description: localize('sessionOutputType.spreadsheetDescription', "Create a spreadsheet or table"),
		icon: Codicon.table,
		instruction: localize('sessionOutputType.spreadsheetInstruction', "Create a spreadsheet, table, or data model as the primary deliverable."),
	},
	{
		kind: 'mission',
		label: localize('sessionOutputType.mission', "Mission"),
		description: localize('sessionOutputType.missionDescription', "Plan and complete a multi-step objective"),
		icon: Codicon.target,
		instruction: localize('sessionOutputType.missionInstruction', "Treat the request as a multi-step mission: make a plan, execute it, and report the resulting deliverable and remaining follow-ups."),
	},
];

export function getSessionOutputOption(kind: SessionOutputKind): ISessionOutputOption {
	return SESSION_OUTPUT_OPTIONS.find(option => option.kind === kind) ?? SESSION_OUTPUT_OPTIONS[0];
}

export function readSessionOutputType(storageService: IStorageService): SessionOutputKind {
	const stored = storageService.get(SESSION_OUTPUT_TYPE_STORAGE_KEY, StorageScope.PROFILE);
	return isSessionOutputKind(stored) ? stored : 'auto';
}

export function storeSessionOutputType(storageService: IStorageService, kind: SessionOutputKind): void {
	storageService.store(SESSION_OUTPUT_TYPE_STORAGE_KEY, kind, StorageScope.PROFILE, StorageTarget.USER);
}

/** Converts an explicit output choice into hidden context sent with, but not shown in, the draft's input. */
export function toSessionOutputPromptEntry(kind: SessionOutputKind): IPromptTextVariableEntry | undefined {
	const option = getSessionOutputOption(kind);
	if (option.kind === 'auto' || !option.instruction) {
		return undefined;
	}
	return {
		kind: 'promptText',
		id: `${SESSION_OUTPUT_CONTEXT_ID_PREFIX}${option.kind}`,
		name: localize('sessionOutputType.contextName', "Requested output · {0}", option.label),
		value: localize('sessionOutputType.contextValue', "The user selected the \"{0}\" output type. {1} Treat it as the requested primary deliverable unless the request explicitly says otherwise.", option.label, option.instruction),
		modelDescription: 'The output type the user selected for this request',
		icon: option.icon,
		automaticallyAdded: true,
	};
}


export interface ISessionOutputTypePickerOptions {
	readonly selected: IObservable<SessionOutputKind>;
	readonly onSelect: (kind: SessionOutputKind) => void;
}

/** The deliverable picker opened from Add Context in a regular session composer. */
export class SessionOutputTypePicker extends Disposable {
	private _isOpen = false;

	get selected(): IObservable<SessionOutputKind> {
		return this._options.selected;
	}

	constructor(
		private readonly _options: ISessionOutputTypePickerOptions,
		@IActionWidgetService private readonly _actionWidgetService: IActionWidgetService,
	) {
		super();
		this._register(autorun(reader => {
			const selected = this._options.selected.read(reader);
			if (this._isOpen) {
				this._actionWidgetService.updateItems(this._getItems(selected));
			}
		}));
		this._register(toDisposable(() => {
			if (this._isOpen) {
				this._actionWidgetService.hide(true);
			}
		}));
	}

	showPicker(anchor: HTMLElement, restoreFocus: () => void): void {
		if (this._actionWidgetService.isVisible) {
			return;
		}
		const delegate: IActionListDelegate<SessionOutputKind> = {
			onSelect: kind => {
				this._actionWidgetService.hide();
				this._options.onSelect(kind);
			},
			onHide: () => {
				this._isOpen = false;
				if (anchor.isConnected) {
					restoreFocus();
				}
			},
		};
		this._isOpen = true;
		this._actionWidgetService.show('sessions.outputTypePicker', false, this._getItems(this._options.selected.get()), delegate, anchor, undefined, [], {
			getAriaLabel: item => item.label ?? '',
			getWidgetAriaLabel: () => localize('sessionOutputType.ariaLabel', "Output Type"),
		}, { className: 'sessions-new-chat-picker-list sessions-output-type-picker-list', minWidth: 220 });
	}

	private _getItems(selected: SessionOutputKind): IActionListItem<SessionOutputKind>[] {
		return SESSION_OUTPUT_OPTIONS.map(option => ({
			kind: ActionListItemKind.Action,
			label: option.label,
			detail: option.description,
			ariaDescription: option.description,
			labelIcon: option.icon,
			group: { title: '', icon: ThemeIcon.fromId(option.kind === selected ? Codicon.check.id : Codicon.blank.id) },
			item: option.kind,
		}));
	}
}
