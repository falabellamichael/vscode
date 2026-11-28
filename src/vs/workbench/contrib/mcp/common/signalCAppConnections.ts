/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { DEFAULT_AUTH_FLOW_PORT } from '../../../../base/common/oauth.js';
import { localize } from '../../../../nls.js';
import { quickBooksSandboxRedirectUri } from '../../../../platform/native/common/native.js';
import { ICustomizationMarketplaceSourceInfo } from '../../../../platform/customizationMarketplace/common/customizationMarketplaceService.js';
import { CustomizationMarketplaceConfiguration } from '../../../../platform/customizationMarketplace/common/customizationMarketplaceSources.js';
import { SignalCMailProvider } from '../../../services/authentication/common/authentication.js';

export const signalCAppConnectionSource = {
	id: 'signalCAppConnections',
	displayName: localize('apps.source', "Signal-C App Connections"),
	enablementSetting: CustomizationMarketplaceConfiguration.MarketplaceEnabled,
} satisfies ICustomizationMarketplaceSourceInfo;

export interface ISignalCAppConnection {
	readonly id: string;
	readonly name: string;
	readonly description: string;
	readonly aliases?: readonly string[];
	readonly url?: string;
	readonly transport?: 'http' | 'sse';
	readonly documentation: string;
	readonly setup?: string;
	readonly clientIdRequired?: boolean;
	readonly requiresDesktopCallback?: boolean;
	readonly publicClientId?: string;
	readonly supportsClientSecret?: boolean;
	readonly clientSecretRequired?: boolean;
	readonly headers?: Readonly<Record<string, string>>;
	readonly blocked?: string;
	readonly local?: { readonly command: string; readonly tokenStoreVariable: string };
}

export const signalCDirectMailConnections = {
	gmail: {
		id: 'gmail-direct',
		name: localize('apps.gmail.direct', "Gmail (Direct)"),
		description: localize('apps.gmail.direct.description', "Read mail and optionally authorize sending through the Gmail API"),
		documentation: 'https://developers.google.com/identity/protocols/oauth2/native-app',
		setup: localize('apps.gmail.direct.setup', "Enable Gmail API in Google Cloud, configure consent and create a Desktop app OAuth client. Have the client ID and secret ready. If the app is in testing, add your account as a test user."),
	},
	outlook: {
		id: 'outlook-direct',
		name: localize('apps.outlook.direct', "Outlook / Microsoft 365 Mail"),
		description: localize('apps.outlook.direct.description', "Read Outlook.com and Microsoft 365 mail and optionally authorize sending"),
		documentation: 'https://learn.microsoft.com/en-us/entra/identity-platform/reply-url',
		setup: localize('apps.outlook.direct.setup', "Register a multitenant Microsoft Entra public client with delegated Microsoft Graph Mail.Read permission. Add delegated Mail.Send if you want to authorize sending. Under Mobile and desktop applications, add {0} as a redirect URI. Include personal Microsoft accounts in the app registration if you use Outlook.com.", 'http://localhost/'),
	},
} satisfies Record<SignalCMailProvider, ISignalCAppConnection>;

export function getSignalCDirectMailProvider(id: string): SignalCMailProvider | undefined {
	return id === signalCDirectMailConnections.gmail.id ? 'gmail' : id === signalCDirectMailConnections.outlook.id ? 'outlook' : undefined;
}

export function matchesSignalCAppConnection(app: ISignalCAppConnection, query: string): boolean {
	const searchable = [app.name, app.description, ...(app.aliases ?? [])].join(' ').toLowerCase();
	return query.trim().toLowerCase().split(/\s+/).every(term => searchable.includes(term));
}

export function findSignalCAppConnection(nameOrId: string): ISignalCAppConnection | undefined {
	const key = nameOrId.trim().toLowerCase();
	return signalCAppConnections.find(app => app.id === key || app.name.toLowerCase() === key
		|| app.aliases?.some(alias => alias.toLowerCase() === key));
}

/** Microsoft's published remote Work IQ connection for IDEs. */
export const microsoft365WorkIq = {
	id: 'microsoft-365',
	name: 'Microsoft 365 (Work IQ)',
	aliases: ['Microsoft 365', 'M365', 'Office 365', 'Work IQ', 'Outlook', 'Outlook Mail', 'Outlook Calendar', 'Microsoft Teams', 'Teams', 'OneDrive', 'SharePoint'],
	description: localize('apps.microsoft365', "Outlook mail and calendar, Microsoft Teams, OneDrive, and SharePoint for work or school accounts"),
	url: 'https://workiq.svc.cloud.microsoft/mcp',
	documentation: 'https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/work-iq/mcp/overview',
	setup: localize('apps.microsoft365.setup', "Sign in with your Microsoft work or school account. Your organization's administrator must grant Work IQ access and set up its Copilot Credits billing policy. Microsoft 365 permissions and your organization's policy control which Outlook, Teams, OneDrive, and SharePoint data is available."),
	// https://github.com/microsoft/work-iq/blob/main/plugins/workiq/README.md
	publicClientId: 'ba081686-5d24-4bc6-a0d6-d034ecffed87',
} satisfies ISignalCAppConnection;

const googleWorkspaceOAuth = {
	documentation: 'https://developers.google.com/workspace/guides/configure-mcp-servers',
	clientIdRequired: true,
	requiresDesktopCallback: true,
	supportsClientSecret: true,
	clientSecretRequired: true,
};
const googleWorkspaceSetup = localize('apps.googleWorkspace.desktopSetup', "Enroll your Google Cloud project in the Google Workspace Developer Preview Program, enable the MCP service, and configure the OAuth consent screen. In Google Auth Platform, create a Web application OAuth client with the authorized redirect URI {0} and select Use this client for an AI-powered agent. Have the client ID and client secret ready. Signal-C stores the secret separately from the MCP configuration.", `http://127.0.0.1:${DEFAULT_AUTH_FLOW_PORT}/`);
const amplitudeDocumentation = 'https://www.amplitude.com/docs/amplitude-ai/amplitude-mcp/other-clients';
const datadogSites = [
	['US1', 'datadoghq.com'],
	['US3', 'us3.datadoghq.com'],
	['US5', 'us5.datadoghq.com'],
	['EU1', 'datadoghq.eu'],
	['AP1', 'ap1.datadoghq.com'],
	['AP2', 'ap2.datadoghq.com'],
	['UK1', 'uk1.datadoghq.com'],
] as const;
const datadogSetup = localize('apps.datadog.setup', "Choose the site that hosts your Datadog organization. Sign in and select that organization. Tool access follows your Datadog permissions.");
const supabaseDocumentation = 'https://supabase.com/docs/guides/ai-tools/mcp';
const supabaseFullFeatureGroups = 'docs,account,database,debugging,development,functions,branching,feedback,notebooks,storage';
const docusignDocumentation = 'https://developers.docusign.com/platform/mcp-server/github-copilot';
const docusignSetup = localize('apps.docusign.setup', "Register the Signal-C desktop redirect URI {0} in your Docusign app. Enter its Integration Key as the client ID and enter its client secret when prompted. Docusign requires approved Integration Keys for production. Use the Demo connection with a developer account and its own Integration Key.", `http://127.0.0.1:${DEFAULT_AUTH_FLOW_PORT}/`);

/** Vendor-published MCP connection presets. OAuth and connection state remain owned by the MCP services. */
export const signalCAppConnections: readonly ISignalCAppConnection[] = [
	microsoft365WorkIq,
	{
		id: 'quickbooks-local', name: 'QuickBooks Online (Local Server)',
		description: localize('apps.quickbooks.local', "Official Intuit MCP server with your company's OAuth connection"),
		documentation: 'https://github.com/intuit/quickbooks-online-mcp-server',
		setup: localize('apps.quickbooks.local.setup', "Build Intuit's official QuickBooks MCP server and have its JavaScript entrypoint ready. For browser sign-in to a sandbox company, register {0} with your Intuit development app and have its Development client ID and secret ready. To use an already-authorized sandbox or production company, choose Connect Local Server and select its private token-store file. Node.js must be available on this machine.", quickBooksSandboxRedirectUri),
		local: { command: 'node', tokenStoreVariable: 'QUICKBOOKS_TOKEN_STORE_PATH' },
	},
	{
		id: 'quickbooks', name: 'QuickBooks Online (Hosted Pilot)',
		description: localize('apps.quickbooks', "Accounting, customers, invoices and reports"),
		url: 'https://mcp.quickbooks.intuit.com/mcp',
		documentation: 'https://intuitdeveloper.github.io/intuit-3p-ai-pilot/',
		setup: localize('apps.quickbooks.setup', "Intuit's hosted MCP is an invitation-only pilot. Your Intuit app, IP range and restricted MCP scopes must be approved. Use that app's public OAuth client ID; a QuickBooks subscription alone does not grant pilot access."),
		clientIdRequired: true, headers: { 'User-Agent': 'partner_app_Signal-C' },
		blocked: localize('apps.quickbooks.hosted.blocked', "Sign-in to the hosted pilot is currently unavailable in Signal-C. Choose QuickBooks Online (Local Server), or contact your Intuit Solution Engineer if you are a pilot partner."),
	},
	{ id: 'notion', name: 'Notion', description: localize('apps.notion', "Pages, databases, documents and project knowledge"), url: 'https://mcp.notion.com/mcp', documentation: 'https://developers.notion.com/guides/mcp/get-started-with-mcp' },
	{ id: 'linear', name: 'Linear', description: localize('apps.linear', "Issues, projects, cycles and planning"), url: 'https://mcp.linear.app/mcp', documentation: 'https://linear.app/docs/mcp' },
	{ id: 'github', name: 'GitHub', description: localize('apps.github', "Repositories, issues, pull requests, workflows and more"), url: 'https://api.githubcopilot.com/mcp/x/all', documentation: 'https://github.com/github/github-mcp-server/blob/main/docs/remote-server.md', supportsClientSecret: true, clientSecretRequired: true, setup: localize('apps.github.setup', "Sign in to GitHub in your browser. This connection requests the full GitHub MCP toolset. Tool access follows your account permissions and organization policies. OAuth app approval may be required.") },
	{
		id: 'figma', name: 'Figma', description: localize('apps.figma', "Designs, components, variables and canvas context"),
		url: 'https://mcp.figma.com/mcp', documentation: 'https://www.figma.com/mcp-catalog/',
		setup: localize('apps.figma.setup', "Figma requires remote MCP clients to be listed in its MCP catalog. Open Setup Guide for the client registration process."),
		blocked: localize('apps.figma.clientApproval', "Browser sign-in from Signal-C requires Figma to approve this MCP client."),
	},
	{ id: 'atlassian', name: 'Atlassian Rovo', description: localize('apps.atlassian', "Jira, Confluence, Bitbucket and team knowledge"), url: 'https://mcp.atlassian.com/v2/mcp', documentation: 'https://support.atlassian.com/atlassian-rovo-mcp-server/', setup: localize('apps.atlassian.setup', "Your Atlassian administrator may need to allow this client and the products you want to access.") },
	{ id: 'asana', name: 'Asana', description: localize('apps.asana', "Tasks, projects, portfolios and team work"), url: 'https://mcp.asana.com/v2/mcp', documentation: 'https://developers.asana.com/docs/integrating-with-asanas-mcp-server', clientIdRequired: true, supportsClientSecret: true, clientSecretRequired: true, setup: localize('apps.asana.setup', "Create an Asana MCP OAuth app and register the redirect URLs shown during authentication. Keep its client secret in the MCP credential store when requested.") },
	{ id: 'slack', name: 'Slack', description: localize('apps.slack', "Workspace messages, channels and collaboration"), url: 'https://mcp.slack.com/mcp', documentation: 'https://docs.slack.dev/ai/slack-mcp-server/', clientIdRequired: true, supportsClientSecret: true, setup: localize('apps.slack.setup', "Use an internal or Slack Marketplace app approved for the Slack MCP server. Enter its public client ID. For confidential apps, enter the client secret in the separate credential prompt. Workspace administrator approval may be required.") },
	{ id: 'canva', name: 'Canva', description: localize('apps.canva', "Designs, brand assets and presentations"), url: 'https://api.canva.com/connect/v1/mcp', documentation: 'https://www.canva.dev/docs/apps/mcp/access/', clientIdRequired: true, supportsClientSecret: true, clientSecretRequired: true, setup: localize('apps.canva.setup', "Register your AI application in the Canva Developer Portal and use its OAuth client ID. Each user authorizes their own Canva access.") },
	{ id: 'hubspot', name: 'HubSpot', description: localize('apps.hubspot', "CRM contacts, companies, deals and activities"), url: 'https://mcp.hubspot.com', documentation: 'https://developers.hubspot.com/docs/apps/developer-platform/build-apps/integrate-with-the-remote-hubspot-mcp-server', clientIdRequired: true, supportsClientSecret: true, clientSecretRequired: true, setup: localize('apps.hubspot.setup', "Create an MCP Auth App in HubSpot. Use its client ID and provide the client secret through the MCP credential prompt or Set Client Secret action.") },
	{ id: 'intercom-us', name: 'Intercom (US)', description: localize('apps.intercom.us', "US-hosted conversations, contacts and support"), url: 'https://mcp.intercom.com/mcp', documentation: 'https://developers.intercom.com/docs/guides/mcp' },
	{ id: 'intercom-eu', name: 'Intercom (EU)', description: localize('apps.intercom.eu', "EU-hosted conversations, contacts and support"), url: 'https://mcp.eu.intercom.com/mcp', documentation: 'https://developers.intercom.com/docs/guides/mcp' },
	{ id: 'stripe', name: 'Stripe', description: localize('apps.stripe', "Payments, customers, invoices and subscriptions"), url: 'https://mcp.stripe.com', documentation: 'https://docs.stripe.com/mcp' },
	{ id: 'sentry', name: 'Sentry', description: localize('apps.sentry', "Errors, performance, issues and project diagnostics"), url: 'https://mcp.sentry.dev/mcp', documentation: 'https://mcp.sentry.dev/' },
	{ id: 'cloudflare', name: 'Cloudflare', description: localize('apps.cloudflare', "DNS, Workers, R2, security and cloud services"), url: 'https://mcp.cloudflare.com/mcp', documentation: 'https://developers.cloudflare.com/agents/model-context-protocol/cloudflare/servers-for-cloudflare/' },
	{ id: 'vercel', name: 'Vercel', description: localize('apps.vercel', "Projects, deployments and hosting"), url: 'https://mcp.vercel.com', documentation: 'https://vercel.com/docs/agent-resources/vercel-mcp' },
	{ id: 'neon', name: 'Neon', description: localize('apps.neon', "Postgres databases, branches and project management"), url: 'https://mcp.neon.tech/mcp', documentation: 'https://neon.com/docs/ai/neon-mcp-server' },
	{ id: 'supabase', name: 'Supabase', description: localize('apps.supabase', "Postgres, projects, Edge Functions and development tools"), url: 'https://mcp.supabase.com/mcp', documentation: supabaseDocumentation, setup: localize('apps.supabase.setup', "Sign in to Supabase and select the organization to authorize. Tool access follows your Supabase role. Choose Supabase (All Tools) to request Storage tools as well, or Supabase (Read Only) for read-only access.") },
	{
		id: 'supabase-all-tools', name: 'Supabase (All Tools)',
		description: localize('apps.supabase.allTools', "All supported Supabase tool groups, including Storage"),
		url: `https://mcp.supabase.com/mcp?features=${supabaseFullFeatureGroups}`, documentation: supabaseDocumentation,
		setup: localize('apps.supabase.allTools.setup', "Sign in to Supabase and select the organization to authorize. This connection requests all documented hosted tool groups, including Storage. Supabase still applies your account permissions, plan, and tool confirmations."),
	},
	{ id: 'supabase-read-only', name: 'Supabase (Read Only)', description: localize('apps.supabase.readOnly', "Explore Supabase projects and database schemas in read-only mode"), url: 'https://mcp.supabase.com/mcp?read_only=true', documentation: supabaseDocumentation },
	{ id: 'huggingface', name: 'Hugging Face', aliases: ['HF', 'HuggingFace'], description: localize('apps.huggingFace', "Hub models, datasets, Spaces, papers and tools selected in your Hugging Face account"), url: 'https://huggingface.co/mcp?login', documentation: 'https://huggingface.co/docs/hub/en/agents-mcp', setup: localize('apps.huggingFace.setup', "Sign in to Hugging Face. Manage which Space tools and capabilities your connection exposes in your Hugging Face MCP settings.") },
	{ id: 'docusign', name: 'Docusign (Production)', description: localize('apps.docusign', "Agreements, envelopes and signing workflows in a production account"), url: 'https://mcp.docusign.com/mcp', documentation: docusignDocumentation, clientIdRequired: true, supportsClientSecret: true, clientSecretRequired: true, setup: docusignSetup },
	{ id: 'docusign-demo', name: 'Docusign (Demo)', description: localize('apps.docusign.demo', "Agreements, envelopes and signing tools in a Docusign developer account"), url: 'https://mcp-d.docusign.com/mcp', documentation: docusignDocumentation, clientIdRequired: true, supportsClientSecret: true, clientSecretRequired: true, setup: docusignSetup },
	{ id: 'monday', name: 'monday.com', description: localize('apps.monday', "Boards, work items, projects and team workflows"), url: 'https://mcp.monday.com/mcp', documentation: 'https://developer.monday.com/api-reference/docs/mcp-dynamic-client-registration' },
	{ id: 'todoist', name: 'Todoist', description: localize('apps.todoist', "Tasks, projects and personal planning"), url: 'https://ai.todoist.net/mcp', documentation: 'https://developer.todoist.com/api/v1/' },
	{ id: 'clickup', name: 'ClickUp', description: localize('apps.clickup', "Tasks, spaces, documents and project work"), url: 'https://mcp.clickup.com/mcp', documentation: 'https://developer.clickup.com/docs/connect-an-ai-assistant-to-clickups-mcp-server' },
	{ id: 'airtable', name: 'Airtable', description: localize('apps.airtable', "Bases, tables, records and structured work"), url: 'https://mcp.airtable.com/mcp', documentation: 'https://support.airtable.com/articles/9897799762-using-the-airtable-mcp-server' },
	{ id: 'miro', name: 'Miro', description: localize('apps.miro', "Boards, diagrams and visual collaboration"), url: 'https://mcp.miro.com/', documentation: 'https://developers.miro.com/docs/miro-mcp-server-faq-and-troubleshooting' },
	{ id: 'webflow', name: 'Webflow', description: localize('apps.webflow', "Sites, CMS content and design workflows"), url: 'https://mcp.webflow.com/mcp', documentation: 'https://developers.webflow.com/mcp/reference/getting-started', setup: localize('apps.webflow.setup', "Site owner or administrator access is required for authorization. Designer tools also require the Webflow MCP Bridge app to remain open.") },
	{ id: 'gmail', name: 'Gmail', description: localize('apps.gmail', "Google email, conversations, drafts and labels"), url: 'https://gmailmcp.googleapis.com/mcp/v1', ...googleWorkspaceOAuth, documentation: 'https://developers.google.com/workspace/gmail/api/guides/configure-mcp-server', setup: googleWorkspaceSetup },
	{ id: 'google-drive', name: 'Google Drive', description: localize('apps.googleDrive', "Files, folders and shared workspace documents"), url: 'https://drivemcp.googleapis.com/mcp/v1', ...googleWorkspaceOAuth, setup: googleWorkspaceSetup },
	{ id: 'google-docs', name: 'Google Docs', description: localize('apps.googleDocs', "Documents, content and collaborative writing"), url: 'https://docsmcp.googleapis.com/mcp/v1', ...googleWorkspaceOAuth, setup: googleWorkspaceSetup },
	{ id: 'google-sheets', name: 'Google Sheets', description: localize('apps.googleSheets', "Spreadsheets, tables and cell data"), url: 'https://sheetsmcp.googleapis.com/mcp/v1', ...googleWorkspaceOAuth, setup: googleWorkspaceSetup },
	{ id: 'google-slides', name: 'Google Slides', description: localize('apps.googleSlides', "Presentations, slides and collaborative decks"), url: 'https://slidesmcp.googleapis.com/mcp/v1', ...googleWorkspaceOAuth, setup: googleWorkspaceSetup },
	{ id: 'google-calendar', name: 'Google Calendar', description: localize('apps.googleCalendar', "Calendars, events and scheduling"), url: 'https://calendarmcp.googleapis.com/mcp/v1', ...googleWorkspaceOAuth, setup: googleWorkspaceSetup },
	{ id: 'google-chat', name: 'Google Chat', description: localize('apps.googleChat', "Workspace chat spaces, messages and collaboration"), url: 'https://chatmcp.googleapis.com/mcp/v1', ...googleWorkspaceOAuth, setup: `${googleWorkspaceSetup} ${localize('apps.googleChat.setup', "Configure a Google Chat app in the same Cloud project before using Chat MCP.")}` },
	{ id: 'google-contacts', name: 'Google Contacts', description: localize('apps.googleContacts', "People API contacts, profiles and connections"), url: 'https://people.googleapis.com/mcp/v1', ...googleWorkspaceOAuth, setup: googleWorkspaceSetup },
	{ id: 'google-classroom', name: 'Google Classroom', description: localize('apps.googleClassroom', "Classes, coursework and classroom collaboration"), url: 'https://classroommcp.googleapis.com/mcp/v1', ...googleWorkspaceOAuth, setup: `${googleWorkspaceSetup} ${localize('apps.googleClassroom.setup', "Follow the Classroom prerequisites in the setup guide, including a test class and student account for testing.")}` },
	{ id: 'box', name: 'Box', description: localize('apps.box', "Files, folders, content search and Box AI"), url: 'https://mcp.box.com', documentation: 'https://support.box.com/hc/en-us/articles/43847256139923-Managing-Box-MCP-Servers', clientIdRequired: true, supportsClientSecret: true, clientSecretRequired: true, setup: localize('apps.box.setup', "Have your Box administrator configure an MCP server integration and its redirect URI in the Admin Console. Use the generated client ID, and provide its client secret through the MCP credential flow. Box does not support automatic client registration.") },
	{ id: 'dropbox', name: 'Dropbox', description: localize('apps.dropbox', "Files, folders, document content and Dropbox Paper"), url: 'https://mcp.dropbox.com/mcp', documentation: 'https://help.dropbox.com/integrations/connect-dropbox-mcp-server', clientIdRequired: true, setup: localize('apps.dropbox.setup', "Dropbox limits automatic client registration to approved clients. For Signal-C, follow the custom OAuth app setup in Dropbox's guide and register the redirect URLs shown during authentication, then enter your app's public client ID.") },
	{ id: 'xero', name: 'Xero', description: localize('apps.xero', "Accounting, invoices, contacts and organization data"), url: 'https://api.xero.com/mcp/v1', documentation: 'https://developer.xero.com/documentation/guides/remote-mcp/overview', clientIdRequired: true, setup: localize('apps.xero.setup', "Register a Xero OAuth app with PKCE and the redirect URI shown by Signal-C. Enter its public client ID and authorize the Xero organizations you want to connect. Your existing Xero app permissions and plan requirements apply.") },
	{
		id: 'paypal', name: 'PayPal (Production)',
		description: localize('apps.paypal', "Live invoices, payments and transactions"),
		url: 'https://mcp.paypal.com/sse', transport: 'sse', documentation: 'https://developer.paypal.com/ai-tools/mcp-server/',
		setup: localize('apps.paypal.setup', "Sign in to the PayPal account you want to connect. This server uses live account data and transactions. Choose PayPal (Sandbox) to work with a test account."),
	},
	{
		id: 'paypal-sandbox', name: 'PayPal (Sandbox)',
		description: localize('apps.paypal.sandbox', "Test invoices, payments and transactions"),
		url: 'https://mcp.sandbox.paypal.com/sse', transport: 'sse', documentation: 'https://developer.paypal.com/ai-tools/mcp-server/',
		setup: localize('apps.paypal.sandbox.setup', "Sign in with a PayPal sandbox test account to work with test payments and invoices."),
	},
	{ id: 'posthog', name: 'PostHog', description: localize('apps.posthog', "Product analytics, feature flags and experiments"), url: 'https://mcp.posthog.com/mcp', documentation: 'https://posthog.com/docs/model-context-protocol/vscode' },
	{
		id: 'gitlab-com', name: 'GitLab.com',
		description: localize('apps.gitlab', "Projects, issues, merge requests and pipelines"),
		url: 'https://gitlab.com/api/v4/mcp', documentation: 'https://docs.gitlab.com/user/model_context_protocol/mcp_server/',
		setup: localize('apps.gitlab.setup', "Enable MCP server access in your GitLab.com top-level group. Sign in with the GitLab account you want to authorize."),
	},
	{ id: 'netlify', name: 'Netlify', description: localize('apps.netlify', "Sites, deployments, build logs and hosting"), url: 'https://mcp.netlify.com/mcp', documentation: 'https://docs.netlify.com/build/build-with-ai/agent-setup-guides/agent-setup-overview/' },
	{ id: 'cal-com', name: 'Cal.com', description: localize('apps.calCom', "Availability, bookings and event types"), url: 'https://mcp.cal.com/mcp', documentation: 'https://cal.com/docs/mcp-server' },
	{ id: 'attio', name: 'Attio', description: localize('apps.attio', "CRM people, companies, deals, tasks and notes"), url: 'https://mcp.attio.com/mcp', documentation: 'https://docs.attio.com/mcp/overview' },
	{ id: 'fireflies', name: 'Fireflies', description: localize('apps.fireflies', "Meeting transcripts, summaries, action items and search"), url: 'https://api.fireflies.ai/mcp', documentation: 'https://guide.fireflies.ai/articles/8272956938-learn-about-the-fireflies-mcp-server-model-context-protocol' },
	{
		id: 'apollo', name: 'Apollo',
		description: localize('apps.apollo', "Sales accounts, contacts, lists and outreach"),
		url: 'https://mcp.apollo.io/mcp', documentation: 'https://docs.apollo.io/docs/apollo-mcp',
		setup: localize('apps.apollo.setup', "Confirm model training is disabled for the model account or client that will use Apollo data. Sign in to the Apollo workspace you want to authorize. Your plan, permissions and credits apply."),
	},
	{ id: 'amplitude-us', name: 'Amplitude (US)', description: localize('apps.amplitude.us', "US-hosted product analytics, experiments and project data"), url: 'https://mcp.amplitude.com/mcp', documentation: amplitudeDocumentation },
	{ id: 'amplitude-eu', name: 'Amplitude (EU)', description: localize('apps.amplitude.eu', "EU-hosted product analytics, experiments and project data"), url: 'https://mcp.eu.amplitude.com/mcp', documentation: amplitudeDocumentation },
	...datadogSites.map(([site, domain]) => ({
		id: `datadog-${site.toLowerCase()}`, name: `Datadog (${site})`,
		description: localize('apps.datadog', "Metrics, logs, traces and incidents"),
		url: `https://mcp.${domain}/v1/mcp`, documentation: 'https://docs.datadoghq.com/mcp_server/setup/',
		setup: datadogSetup,
	})),
];
