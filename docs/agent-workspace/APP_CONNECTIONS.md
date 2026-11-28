# Signal-C app connections

Open **Assistant → Connect App**, enter `/connect`, or select **Connect** on a **Signal-C App Connections** card in **Customizations → Browse**. These paths connect the app to the current Signal-C profile for Assistant. Choose **Connect with OAuth** and finish signing in and granting access on the provider's website. Click a saved app marked **Sign-in required** to start the browser flow again. For another installation target, use **Customizations → MCP Servers → Connect App**, **Signal-C: Connect App…** from the Command Palette, or **Connect an App (OAuth)…** inside **MCP: Add Server…**. The catalog is defined in `src/vs/workbench/contrib/mcp/common/signalCAppConnections.ts`.

Search by app name or description. Selecting an app opens a standard keyboard-accessible review dialog showing its setup requirements and endpoint. **Open Setup Guide** opens the provider's documentation without changing configuration. **Cancel** or Escape exits the step without installing anything. For remote servers, continue with OAuth, name the connection, then select its destination. The suggested name includes a unique suffix so repeated app setup does not default to replacing an earlier account or manual connection; you can edit the name. Authorization happens when the installed server starts, using the owning MCP client's OAuth flow. An installed entry is not proof that authorization completed or that a provider exposed tools to a model.

Registered-client providers such as Asana, Slack, Canva and HubSpot ask for an OAuth app client ID. Register the desktop callback with the provider. When a client secret is required, enter it in the masked credential prompt. Signal-C stores it in Secret Storage and opens the browser for your account consent. Provider registration, administrator approval, licensing and regional availability still apply.

For providers that register clients automatically, Signal-C asks for a public PKCE client first. If the provider rejects that method and advertises a supported client-secret method, Signal-C retries registration over HTTPS and stores the resulting client credentials for browser sign-in and refresh.

If a connected app rejects its current access token, Signal-C can try the saved refresh grant for that approved connection. If the provider cannot refresh it, use browser sign-in again to reconnect the account.

Registered OAuth apps that require a public HTTPS callback can use **Connect with HTTPS Callback** in Assistant when the desktop build has a configured Signal-C callback. Register the exact HTTPS URL shown in the setup dialog with the provider, enter that app's client ID and secret as requested, and finish sign-in in your browser. On the return page choose **Open Signal-C**. The MCP configuration stores the public client ID and `oauth.useHttpsCallback`.

Remote MCP servers contributed by plugins use the same authorization flow. A plugin that supplies a registered `oauth.clientId` can opt into the build's HTTPS callback with `oauth.useHttpsCallback`. Start the installed server from **Customizations → MCP Servers** to sign in and load the tools the provider makes available.

GitHub's standard browser flow can ask you to copy a one-time code. If you have your own registered GitHub OAuth app and the build has a public HTTPS callback, choose **Connect with HTTPS Callback** for the authorization-code flow. New GitHub presets use the full toolset endpoint. Existing connections to the default endpoint keep their configured toolset until you add or update the connection.

### Configuring the HTTPS callback for a build

1. Serve [`signal-c-callback.html`](../../resources/oauth/signal-c-callback.html) at a stable HTTPS URL on the domain you control. Serve the file byte-for-byte as `text/html; charset=utf-8`, with `Cache-Control: no-store` and `Content-Security-Policy: frame-ancestors 'none'`. Configure callback access logs to omit query strings, which can contain short-lived authorization codes.
2. Set `signalCOAuthRedirectUri` to that exact URL in the desktop's product configuration. A development checkout reads it from `product.overrides.json`. A packaged build needs it in the packaged product configuration. The URL itself is public and the callback page needs no provider secret.
3. Register the URL with each provider app that needs it. The callback page forwards the pending sign-in's nonce, window ID, and code or error to a fixed `signal-c://` route. Signal-C exchanges the code directly with the provider using PKCE.

The workspace provides the callback page. A build without a configured public URL offers the standard desktop sign-in flow.

## Available apps

Open **Customizations → Browse** and choose **Signal-C App Connections** from the source dropdown, or select **View All** in its gallery section. Search for an app and choose **Connect** to open its setup review. The hosted QuickBooks pilot card opens its setup requirements.

Intercom and Amplitude have separate US and EU endpoints. Datadog has an entry for each supported site, PayPal has production and sandbox endpoints, and QuickBooks has an official local server and a hosted-pilot setup entry.

| App | Route | Provider setup guide |
| --- | --- | --- |
| QuickBooks Online | Official local stdio server; hosted pilot setup guide | [Intuit local server](https://github.com/intuit/quickbooks-online-mcp-server), [hosted pilot](https://intuitdeveloper.github.io/intuit-3p-ai-pilot/) |
| Notion | Remote HTTP with OAuth | [Notion](https://developers.notion.com/guides/mcp/get-started-with-mcp) |
| Linear | Remote HTTP with OAuth | [Linear](https://linear.app/docs/mcp) |
| GitHub | Remote HTTP with browser sign-in and GitHub's full MCP toolset. Tool access follows account permissions and organization policies | [GitHub remote toolsets](https://github.com/github/github-mcp-server/blob/main/docs/remote-server.md#remote-mcp-toolsets) |
| Figma | Setup guide. Figma client approval is required for the Signal-C browser sign-in | [Figma MCP Catalog](https://www.figma.com/mcp-catalog/) |
| Atlassian Rovo | Remote HTTP with OAuth | [Atlassian](https://support.atlassian.com/atlassian-rovo-mcp-server/) |
| Asana | Remote HTTP with a registered OAuth app | [Asana](https://developers.asana.com/docs/integrating-with-asanas-mcp-server) |
| Slack | Remote HTTP with a registered OAuth app | [Slack](https://docs.slack.dev/ai/slack-mcp-server/) |
| Canva | Remote HTTP with a registered OAuth app | [Canva](https://www.canva.dev/docs/apps/mcp/access/) |
| HubSpot | Remote HTTP with an MCP Auth App | [HubSpot](https://developers.hubspot.com/docs/apps/developer-platform/build-apps/integrate-with-the-remote-hubspot-mcp-server) |
| Intercom | Separate US and EU remote HTTP OAuth connections | [Intercom](https://developers.intercom.com/docs/guides/mcp) |
| Stripe | Remote HTTP with OAuth | [Stripe](https://docs.stripe.com/mcp) |
| Sentry | Remote HTTP with OAuth | [Sentry](https://mcp.sentry.dev/) |
| Cloudflare | Remote HTTP with OAuth | [Cloudflare](https://developers.cloudflare.com/agents/model-context-protocol/cloudflare/servers-for-cloudflare/) |
| Vercel | Remote HTTP with OAuth | [Vercel](https://vercel.com/docs/agent-resources/vercel-mcp) |
| Neon | Remote HTTP with OAuth | [Neon](https://neon.com/docs/ai/neon-mcp-server) |
| Supabase | Remote HTTP with browser OAuth. Choose the default, all-tools with Storage, or read-only preset. The all-tools preset keeps provider confirmations and permissions in effect | [Supabase feature groups](https://supabase.com/docs/guides/ai-tools/mcp#available-tools) |
| Hugging Face | Remote HTTP with browser OAuth. Tools follow your account's MCP settings | [Hugging Face](https://huggingface.co/docs/hub/en/agents-mcp) |
| Docusign | Separate production and demo connections with a registered Integration Key and client secret. Production access requires Docusign approval | [IDE setup](https://developers.docusign.com/platform/mcp-server/github-copilot) · [Production approval](https://developers.docusign.com/platform/mcp-server/limitations-workarounds) |
| monday.com | Remote HTTP with OAuth | [monday.com](https://developer.monday.com/api-reference/docs/mcp-dynamic-client-registration) |
| Todoist | Remote HTTP with OAuth | [Todoist](https://developer.todoist.com/api/v1/) |
| ClickUp | Remote HTTP with OAuth | [ClickUp](https://developer.clickup.com/docs/connect-an-ai-assistant-to-clickups-mcp-server) |
| Airtable | Remote HTTP with OAuth | [Airtable](https://support.airtable.com/articles/9897799762-using-the-airtable-mcp-server) |
| Miro | Remote HTTP with OAuth | [Miro](https://developers.miro.com/docs/miro-mcp-server-faq-and-troubleshooting) |
| Webflow | Remote HTTP with OAuth and site/Designer prerequisites | [Webflow](https://developers.webflow.com/mcp/reference/getting-started) |
| Gmail | Registered OAuth | [Google Workspace](https://developers.google.com/workspace/guides/configure-mcp-servers) |
| Google Drive | Registered OAuth | [Google Workspace](https://developers.google.com/workspace/guides/configure-mcp-servers) |
| Google Docs | Registered OAuth | [Google Workspace](https://developers.google.com/workspace/guides/configure-mcp-servers) |
| Google Sheets | Registered OAuth | [Google Workspace](https://developers.google.com/workspace/guides/configure-mcp-servers) |
| Google Slides | Registered OAuth | [Google Workspace](https://developers.google.com/workspace/guides/configure-mcp-servers) |
| Google Calendar | Registered OAuth | [Google Workspace](https://developers.google.com/workspace/guides/configure-mcp-servers) |
| Google Chat | Registered OAuth | [Google Workspace](https://developers.google.com/workspace/guides/configure-mcp-servers) |
| Google Contacts | Registered OAuth; People API | [Google Workspace](https://developers.google.com/workspace/guides/configure-mcp-servers) |
| Google Classroom | Registered OAuth | [Google Workspace](https://developers.google.com/workspace/guides/configure-mcp-servers) |
| Box | Registered OAuth; administrator integration setup | [Box](https://support.box.com/hc/en-us/articles/43847256139923-Managing-Box-MCP-Servers) |
| Dropbox | Registered OAuth; automatic registration limited to approved clients | [Dropbox](https://help.dropbox.com/integrations/connect-dropbox-mcp-server) |
| Xero | Registered OAuth with PKCE | [Xero](https://developer.xero.com/documentation/guides/remote-mcp/overview) |
| PayPal | Separate production and sandbox remote SSE connections with browser OAuth | [PayPal](https://developer.paypal.com/ai-tools/mcp-server/) |
| PostHog | Remote HTTP with OAuth and provider-managed region selection | [PostHog](https://posthog.com/docs/model-context-protocol/vscode) |
| GitLab.com | Remote HTTP with OAuth; MCP access enabled in the top-level group | [GitLab](https://docs.gitlab.com/user/model_context_protocol/mcp_server/) |
| Netlify | Remote HTTP with OAuth | [Netlify](https://docs.netlify.com/build/build-with-ai/agent-setup-guides/agent-setup-overview/) |
| Cal.com | Remote HTTP with OAuth | [Cal.com](https://cal.com/docs/mcp-server) |
| Attio | Remote HTTP with OAuth | [Attio](https://docs.attio.com/mcp/overview) |
| Fireflies | Remote HTTP with OAuth | [Fireflies](https://guide.fireflies.ai/articles/8272956938-learn-about-the-fireflies-mcp-server-model-context-protocol) |
| Apollo | Remote HTTP with OAuth; disable model training for the model account or client | [Apollo](https://docs.apollo.io/docs/apollo-mcp) |
| Amplitude | Separate US and EU remote HTTP OAuth connections | [Amplitude](https://www.amplitude.com/docs/amplitude-ai/amplitude-mcp/other-clients) |
| Datadog | OAuth endpoints for US1, US3, US5, EU1, AP1, AP2 and UK1 | [Datadog](https://docs.datadoghq.com/mcp_server/setup/) |

For Google Workspace, enable the relevant MCP APIs in a Cloud project, configure the documented consent scopes, and mark the OAuth client for an AI-powered agent. Signal-C starts browser authorization before retrieving the tool catalog. Google Chat writes also need a configured Chat app. Follow the Classroom test prerequisites when testing that service.

For a previously saved PayPal `/http` connection, add the updated PayPal preset and sign in again. Existing configurations retain the endpoint and account authorization selected when they were created.

## Gmail and Outlook mail sending

Choose **Gmail (Direct)** or **Outlook / Microsoft 365 Mail** in Assistant. Use **Sign In to Read and Send** to grant the mailbox's sending scope through Google or Microsoft. Existing read access remains available while you complete the new browser grant. After connection, Assistant can send a plain-text message to one recipient. In Auto mode the approval shows the recipient, subject and full message. You can return the connection to read-only access from the same mailbox menu. [Gmail sending API](https://developers.google.com/workspace/gmail/api/reference/rest/v1/users.messages/send) · [Microsoft Graph sendMail](https://learn.microsoft.com/en-us/graph/api/user-sendmail?view=graph-rest-1.0)

In a provisioned build, Signal-C supplies the app registration and launches the system browser, like SignalREACH Studio. For a local build, configure `signalCGmailOAuthClientFile` in the ignored `product.overrides.json` with the absolute path to the Studio publisher's Google Desktop OAuth registration file. Signal-C reads the registration and obtains a separate grant through browser consent. Microsoft uses the configured `signalCOutlookOAuthClientId`.

To send without running a model, type `/send-email recipient@example.com | Subject` in Assistant, press **Shift+Enter**, write the message, and submit. Review the recipient and text, then choose **Send email**. If the mailbox needs permission, the command opens mailbox setup and preserves your message.

For recurring delivery, open **Assistant Routines** and create or edit a routine. Choose **Email the result after each completed run**, select a sending mailbox, enter one recipient, and save the schedule. The routine name becomes the subject. The model produces the message and Signal-C submits it using the saved mailbox and recipient. The run reports a missing mailbox, incomplete result, or unconfirmed submission in its activity. In chat, **Always allow** authorizes future sends from the same mailbox in that conversation.

## QuickBooks local connection

1. Build the official local MCP server from [Intuit's repository](https://github.com/intuit/quickbooks-online-mcp-server). Node.js and the server's dependencies must already be installed.
2. For a sandbox company, register `http://localhost:33419/callback` in the **Development** settings for your Intuit app. In Assistant choose **QuickBooks Online (Local Server) → Sign In to Sandbox**, select the built JavaScript entrypoint, name the connection, and enter the Development client ID and secret. Finish sign-in and select the sandbox company on Intuit's website. Signal-C creates a separate credential file in its local user data and installs the connection after consent.
3. For an already-authorized sandbox or production company, choose **Connect Local Server** and select the built entrypoint and its private token-store file. Intuit requires a registered public HTTPS callback for the initial production OAuth setup. The file picker accepts local paths, and the official server reads the file.
4. Signal-C saves the `node` command, the entrypoint path, and `QUICKBOOKS_TOKEN_STORE_PATH` in local user MCP configuration. The official server reads and rotates its tokens in that file. Enter `/tools` in Assistant to see the tools the running server exposes.

To renew a sandbox connection created by Signal-C, click the saved app in Assistant and select **Sign In to QuickBooks Again**. Select the same sandbox company in Intuit's browser flow. Signal-C stops the server during renewal, updates the existing credential file once consent succeeds, then restarts the same connection.

This local preset does not send machine-specific paths to an agent host or write them to Copilot Global. An explicit workspace or agent-host target is refused before collecting paths; invoke the standalone Connect App command to use local user configuration. Starting and exposing this server still uses the existing MCP trust and tool-approval controls.

Intuit's hosted MCP is an invitation-only pilot requiring app/IP/scope onboarding. Its public protected-resource metadata currently identifies `https://mcp.quickbooks.intuit.com?resource_query=mcp`, which does not match the resource expected by RFC 9728 discovery for the documented `/mcp` endpoint. Signal-C retains resource and issuer validation and keeps this entry setup-only. Recheck the provider's metadata before enabling hosted installation; a QuickBooks subscription does not establish pilot access.

## Verification

The focused MCP configuration entry-point tests cover remote installation and start routing, registered client IDs, setup-guide actions, cancellation at each step, selected agent-host ownership and policy rejection, and QuickBooks local path-only configuration, local ownership and file-picker cancellation. No test reads a real token store or authorizes a provider account.

Run `node scripts/test-agent-workspace-native.mts --empty-workspace --keyboard --app-connections` on the prepared Windows development runtime after compiling the client. It uses a fresh hidden portable fixture, opens the app picker from the real Command Palette, checks QuickBooks local availability and the hosted-pilot review, then opens the MCP Servers management page and activates Connect App from the keyboard. It reviews Google's prerequisites, reaches the public client-ID step, cancels setup, verifies focus returns to the page button, and checks that fixture MCP configuration remains unchanged. It then verifies mode-switch keyboard dispatch, window reuse, unsaved buffers and the Agents draft. Artifacts remain under `.build/agent-workspace-native-smoke/<run>/`.

Public discovery checks use the actual `AgentHostMcpOAuthMetadataService` with credential-free HTTPS GET requests. Discovery of a resource/issuer chain does not establish OAuth grant completion, account entitlement, refresh behavior, or successful tool calls. These require separate provider-account validation. The local QuickBooks wizard uses an existing build of Intuit's server and can perform the initial sandbox OAuth handshake from the desktop.
