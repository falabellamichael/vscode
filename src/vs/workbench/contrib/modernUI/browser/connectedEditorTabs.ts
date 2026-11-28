/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { Color } from '../../../../base/common/color.js';
import { editorBackground, foreground, iconForeground } from '../../../../platform/theme/common/colorRegistry.js';
import { registerThemingParticipant } from '../../../../platform/theme/common/themeService.js';
import { CONNECTED_EDITOR_TABS_SELECTOR } from '../../../browser/parts/editor/editor.js';
import { EDITOR_GROUP_HEADER_CONNECTED_TABS_BACKGROUND, EDITOR_GROUP_HEADER_TABS_BACKGROUND, MODERN_EDITOR_TAB_HOVER_BACKGROUND, PANEL_BACKGROUND, SIDE_BAR_BACKGROUND, TAB_ACTIVE_FOREGROUND, TAB_INACTIVE_FOREGROUND } from '../../../common/theme.js';
import './media/connectedEditorTabs.css';

const connectedTabsSelector = `.monaco-workbench.modern-ui-tabs${CONNECTED_EDITOR_TABS_SELECTOR}`;
const connectedEditorTabsSelector = `${connectedTabsSelector}.modern-ui`;

registerThemingParticipant((theme, collector) => {
	const background = theme.getColor(editorBackground);
	if (background) {
		collector.addRule(`${connectedTabsSelector} { --modern-ui-connected-tab-surface: ${background}; }`);
	}

	const defaultStripColor = theme.getColor(EDITOR_GROUP_HEADER_TABS_BACKGROUND) ?? background;
	const defaultStripBackground = background && defaultStripColor ? defaultStripColor.makeOpaque(background) : defaultStripColor;
	if (defaultStripBackground) {
		collector.addRule(`${connectedTabsSelector} { --modern-ui-connected-tab-strip-background: ${defaultStripBackground}; }`);
	}

	const hoverBackground = theme.getColor(MODERN_EDITOR_TAB_HOVER_BACKGROUND);
	if (defaultStripBackground && hoverBackground) {
		// Flatten against the strip, not the document, so the action mask and pill paint one surface.
		collector.addRule(`${connectedTabsSelector} { --modern-ui-connected-tab-upper-hover-background: ${hoverBackground.makeOpaque(defaultStripBackground)}; }`);
	}

	const editorStripColor = theme.getColor(EDITOR_GROUP_HEADER_CONNECTED_TABS_BACKGROUND) ?? defaultStripColor;
	const editorStripBackground = background && editorStripColor ? editorStripColor.makeOpaque(background) : editorStripColor;
	if (editorStripBackground) {
		collector.addRule(`${connectedEditorTabsSelector} { --modern-ui-connected-tab-strip-background: ${editorStripBackground}; }`);

		// Mixed-polarity themes pair dark editor tab strips with light sidebar/panel chrome; keep pane strips matched to their own surface.
		const sideBarBackground = theme.getColor(SIDE_BAR_BACKGROUND);
		if (sideBarBackground && sideBarBackground.isDarker() !== editorStripBackground.isDarker()) {
			const paneStrip = sideBarBackground.isDarker() ? sideBarBackground.lighten(0.06) : sideBarBackground.darken(0.04);
			collector.addRule(`${connectedEditorTabsSelector} :is(.part.sidebar, .part.auxiliarybar) { --modern-ui-connected-tab-strip-background: ${paneStrip}; }`);
		}

		const panelBackground = theme.getColor(PANEL_BACKGROUND);
		if (panelBackground && panelBackground.isDarker() !== editorStripBackground.isDarker()) {
			const panelStrip = panelBackground.isDarker() ? panelBackground.lighten(0.06) : panelBackground.darken(0.04);
			collector.addRule(`${connectedEditorTabsSelector} .part.panel { --modern-ui-connected-tab-strip-background: ${panelStrip}; }`);
		}

		const iconColor = theme.getColor(iconForeground) ?? theme.getColor(foreground);
		if (iconColor && editorStripBackground.getContrastRatio(iconColor) < 3) {
			const candidateForeground = theme.getColor(TAB_INACTIVE_FOREGROUND) ?? theme.getColor(TAB_ACTIVE_FOREGROUND);
			const flattenedCandidate = candidateForeground ? (candidateForeground.isOpaque() ? candidateForeground : candidateForeground.makeOpaque(editorStripBackground)) : undefined;
			const tabActionForeground = flattenedCandidate && editorStripBackground.getContrastRatio(flattenedCandidate) >= 3
				? flattenedCandidate
				: (editorStripBackground.isDarker() ? Color.white : Color.black);
			collector.addRule(`${connectedTabsSelector} .part.editor > .content .editor-group-container > .title.tabs :is(.editor-actions, .editor-layout-actions, .tabs-bar-add-tab) :is(.action-label, .codicon) { color: ${tabActionForeground}; }`);
		}
	}
	if (editorStripBackground && hoverBackground) {
		collector.addRule(`${connectedEditorTabsSelector} { --modern-ui-connected-tab-upper-hover-background: ${hoverBackground.makeOpaque(editorStripBackground)}; }`);
	}
});
