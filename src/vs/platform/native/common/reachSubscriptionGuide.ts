/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

/** Product reference for regular chat when the user asks about REACH Subscription. */
export const reachSubscriptionGuide = `# REACH Subscription in Signal-C

REACH Subscription is a SignalREACH account plan that gives an eligible account access to hosted AI models. Signal-C uses that account to run conversations through the REACH provider.

## Connect a chat

1. In the local desktop Agents window or IDE, open a new chat and choose **REACH** in the provider picker.
2. Reuse the saved SignalREACH sign-in, or run **Connect REACH Chat** and finish sign-in in your browser. Signal-C keeps the session in its secret store.
3. Select **Auto** to use the first eligible chat model or choose a model from the account’s catalog. Reopen an existing conversation with the same SignalREACH account to keep its history.

## How the client checks access

- \`GET /v1/account\` identifies the account. Before inference, Signal-C checks \`access.activeSubscription\` and requires either \`requestAllowance.unlimited\` or a positive \`requestAllowance.remaining\`.
- \`GET /v1/models\` returns the authenticated model catalog. The regular-chat picker includes subscription request models (\`access === 'requests'\`) that support chat. Embedding and reranking entries serve other operations.
- \`POST /v1/chat/completions\` sends the selected model ID and conversation messages using the saved bearer credential. Keep credentials out of Markdown, prompts, logs, and URLs.

## Plan and troubleshooting

REACH Subscription Plus includes unlimited completed requests to eligible subscription models. Concurrency, rate, request-size and timeout limits still apply. Manage the plan in REACH Studio. Stripe confirms purchases to the service. After checkout, refresh your account in Studio to check activation. Cancelling renewal keeps access through the end of the paid period. The public subscriptions page lists current pricing and plan terms. [REACH subscriptions](https://www.signalreach.ca/subscription.html)

If a chat cannot send, check the signed-in account, active plan, request allowance, and current model catalog in that order. If an explicitly selected model has disappeared, select an available model. Refresh the account after a billing change.
`;

export function concernsReachSubscription(prompt: string): boolean {
	return /\b(?:signal[ -]?)?reach(?:['’]s)?[ -]+(?:subscriptions?|plans?|plus)\b|\bsubscrib(?:e|ing)\s+to\s+(?:signal[ -]?)?reach\b/i.test(prompt);
}
