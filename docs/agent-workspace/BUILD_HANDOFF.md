# Signal-C build handoff

## Active Windows build

Verified on November 19, 2026. Work in `D:\vscode` on `feat/agent-workspace`. The checked-out Git base was `3107e549cee`. The running client also includes local changes compiled from the working tree. Check `git status --short`, the actual process path, and the compiled output before comparing builds. A commit hash alone cannot identify this build.

| Item | Current configuration |
| --- | --- |
| Start menu | **Signal-C**, which opens **Agents** |
| Runtime | `D:\vscode\.build\electron\Signal-C.exe`, loading the checkout and its `out\main.js` |
| Profile | `%APPDATA%\Signal-C`, the user's everyday profile |
| Normal launch | `node scripts/code-agent-workspace.ts --primary` |
| Windows OAuth return | `signal-c://` is registered per user to the same launcher with `--primary --open-url "%1"` |
| Isolated preview | `node scripts/code-agent-workspace.ts` uses `.build\agent-workspace-*` storage |

Use the primary launcher from the Start menu for everyday work. The dated standalone packages on `D:\` are snapshots and rollback copies. Older build notes describe isolated development and optimized packages. This handoff records which one is active for the user.

## What is in the build

| Area | What to look for |
| --- | --- |
| Agents and IDE | Agents opens first. The title-bar switch opens the editor or returns to Agents. Session Tools, Attention Queue, workspace selection, and Customizations live in Agents. See [mode switch details](MODE_SWITCH.md). |
| Assistant | One REACH-backed Assistant, shared chat composer, linked background tasks, image inputs, persistent Markdown documents, and delegation to regular agent chats. **Worktree** and **Agent Chats** appear inside the composer. From a folderless chat, Worktree opens a local folder picker and then lists existing worktrees by branch and path. Agent Chats explains the empty state. The composer switches Send to Stop during a reply. The view's right edge carries a floating activity bar (`.custom-view-activity-bar`, `MenuActivityBar` inside `CustomViewNode`) with the same tab entries as the chat's side strip — **Browser** and **Terminal** — driven by `Menus.AssistantActivityBar`. The descriptor declares `coveredParts: [SESSIONS_PART, AUXILIARYBAR_PART]` so the view covers only the sessions grid and side panel; the Browser tab toggles the browser editor in the uncovered editor part, and the Terminal tab toggles the terminal in the uncovered bottom panel beside the view (`ToggleAssistantBrowserAction`/`ToggleAssistantTerminalAction`; a second click hides the surface and refocuses the view). The native smoke (`--conversation`) asserts the strip docks at the view's right edge and that both surfaces open beside it (`assistant-activity-bar.png`). |
| Regular chat | The new-session composer and ongoing chat input include an **Output** dropdown (also on the sessions window's active-session chat via `ChatView`'s `renderOutputTypePicker`): Auto-Select, Website, Mobile, Design, Slide deck, Animation, Data Visualization, Image, Document, Spreadsheet, and Mission. The choice persists per profile, restores with draft/snapshot handoffs and chat input state, and reaches normal sends, follow-ups, peer chats, Send Later requests, and scheduled edits both as hidden model context and structured `vscode.sessionOutputType` request metadata. Editing a sent or queued request restores its original output into the input (from the request's hidden context, or from request metadata after a reload) without an attachment pill; resubmitting preserves unrelated `agentHostMetadata` keys. The agent host also elevates `vscode.sessionOutputType` from `message._meta` into provider-level `hostInstructions` through `SessionOutputTypeContribution`, so the deliverable hint reaches the system prompt even when a path drops hidden prompt attachments; workspace-less turns are steered to deliver inline instead of claiming saved files. The dropdown marks the selected option with a check while keeping each kind's icon via `labelIcon`, and a session switch to an input state without `sendContext` clears stale hidden entries without resetting the picker to anything but the stored default. While the selection is anything but Auto-Select, the attach-context ("+") action item carries a `.chat-output-type-indicator` icon element immediately right of the "+" (rendered by the `MenuEntryActionViewItem` subclass in `chatInputPart.ts` `_createAttachContextActionViewItem`, bound to `SessionOutputTypePicker.selected`); it re-opens the picker on activation and hides when Auto-Select is restored. The native smoke (`scripts/test-agent-workspace-native.mts --conversation` or `--protocol`) opens the Agents composer dropdown, picks Website, asserts the sent mock request carries both the output selection and the contribution's `<requested_output>` host instruction, and resets to Auto-Select so later stages are unaffected. |
| Briefings | Morning and evening briefings, Town-style routine templates, `@`-mention invocation, and recurring routines can use connected sources. Assistant can review connected history and propose memories. The weather location is saved in `sessions.assistant.weatherLocation`. The client has a built-in Weather Network path, and the current source build includes an Open-Meteo fallback for eligible forecast requests. The working tree also includes changes for multiple routine proposals and local-model text tool calls. See [Assistant details](ASSISTANT.md). |
| Connections | Assistant and the Customizations gallery expose app setup. The catalog includes Slack, Gmail, Outlook / Microsoft 365 Mail, PayPal, PostHog, and QuickBooks. QuickBooks local mode uses the official local server. The hosted pilot entry shows setup requirements. OAuth setup and provider approvals vary by connector. See [app connections](APP_CONNECTIONS.md). |
| Subscriptions | Native Codex and Claude are available in the Agents picker when configured. Devin Subscription uses its browser account flow. Devin CLI is a separate terminal entry. See [subscription setup](SUBSCRIPTIONS.md). |
| Appearance | Signal-C branding, the SimpleRAG mark in the Agents animation, and theme-aware Agents surfaces. See [branding](BRANDING.md). |


On November 27, the focused Assistant Chromium suite passed 96 tests after adding dynamic on-demand MCP activation, routine `@`-mentions, the routine template gallery, image inputs, reply document cards, the persistent document library, side-by-side editable artifacts, open-document chat context, and outcome-first response guidance. Transpilation and `git diff --check` were clean. Typecheck still reports unrelated pre-existing working-tree errors outside the touched Assistant files.

The November 19 desktop fixture verified both Assistant controls, the empty Agent Chats state, and selection of a real linked worktree from a folderless chat. The Assistant Chromium suite passed 50 tests. The November 18 cutover checks verified the Windows Gmail callback route on cold and warm launch, the Microsoft browser handoff, and Claude and Codex registration in an isolated picker. These checks used fixtures. Completing OAuth, checking account entitlement, and sending a provider-backed reply are separate checks.

On November 23, an isolated desktop profile verified the Assistant skill picker and named `/skill` commands. With fixture account authentication and the real local LFM2.5-1.2B model, an attached meeting workflow preserved Mira's Friday deadline and a plain-language recap preserved Rowan's Monday deadline. Both answers used the supplied notes without making a tool call. [Assistant skills and commands](ASSISTANT.md#attachments-and-commands)

In a separate desktop check, a question about the installed meeting workflow read its instructions and correctly explained how to report a missing deadline. Another fixture verified that `/compact` leaves the original exchanges available to search, export, undo and clear. The profile used for these checks was isolated from the user's everyday sign-in.

The idea-exploration check gave the local model one hour, paper and a pen to start a family recipe collection. It proposed collecting memories or sketching a recipe card. Its experiment paired a grandmother's story with drawing a card together and watching where she hesitated over the ingredients.

The November 23 Saved ideas desktop checks saved a Markdown reply, cleared its source chat, and attached the saved copy to an existing linked-task draft. A mouse drag inside a 9,302-character reply saved the selected 12 characters. A second launch restored the collection and unsent draft. Removing one passage preserved the other. These local actions made zero model requests, and a second synthetic account’s records stayed unchanged. [Using Saved ideas](ASSISTANT.md#saved-ideas)

In the same checkout, a local LFM2.5-1.2B desktop run received one matching passage about copper-tape leaves that unfold to reveal handwritten memories. It suggested a modular paper-tree display. Signal-C added the source chat, **Memory Tree Workshop**, below the answer. The one model request excluded an unrelated idea in the fixture account and a matching idea in a second account.

A three-launch desktop check exercised **Continue in task**. The task kept its original ID, one fixture exchange and an edited unsent draft after restart. The parent retained its own draft and attachment. After the user-confirmed task deletion, the next launch started a fresh task from the saved passage. Only the fixture’s explicit Send produced a completion request. The final regression run passed 180 Chromium tests and 18 native Assistant tests.

The same 1.2B model received a release decision where Friday offered feedback three days earlier and Monday had support. It chose Monday for the stated priority of protecting users. Its final answer incorrectly described the lack of weekend support as Monday's drawback. The actual tradeoff was receiving feedback three days later. This remains a regression case for decision quality.

| Assistant implementation | Owning files |
| --- | --- |
| Workflow instructions, discovery and worktree scope | [Bundled skills](../../src/vs/sessions/skills), [skill discovery](../../src/vs/sessions/contrib/sessions/browser/assistantSkills.ts), [target trust](../../src/vs/sessions/contrib/sessions/browser/assistantAgentChats.ts) |
| Composer commands, linked tasks and thread controls | [Assistant view](../../src/vs/sessions/contrib/sessions/browser/assistantSession.ts), [run orchestration](../../src/vs/sessions/contrib/sessions/browser/assistantRuns.ts) |
| Saved ideas, working chats and source attribution | [Saved idea and task storage](../../src/vs/sessions/contrib/sessions/browser/assistantConversation.ts), [scoped recall](../../src/vs/sessions/contrib/sessions/browser/assistantSavedIdeaRecall.ts), [lifecycle tests](../../src/vs/sessions/contrib/sessions/test/browser/assistantSavedIdeas.test.ts) |
| Memory editing, searchable history and saved summaries | [Memory operations](../../src/vs/sessions/contrib/sessions/browser/assistantMemory.ts), [conversation storage](../../src/vs/sessions/contrib/sessions/browser/assistantConversation.ts), [Assistant tools](../../src/vs/sessions/contrib/sessions/browser/assistantTools.ts) |
| Draft dictation, output selection and speech lifecycle | [Shared composer](../../src/vs/workbench/contrib/sessions/browser/chat/sharedChatInputComposer.ts), [ongoing chat input](../../src/vs/workbench/contrib/chat/browser/widget/input/chatInputPart.ts), [output picker](../../src/vs/workbench/contrib/chat/browser/widget/input/sessionOutputTypePicker.ts), [shared output kind](../../src/vs/workbench/contrib/chat/common/sessionOutputType.ts), [host output steering](../../src/vs/platform/agentHost/node/chatContributions/sessionOutputType/sessionOutputTypeContribution.ts), [dictation session](../../src/vs/workbench/contrib/chat/browser/speechToText/dictationSession.ts), [speech actions](../../src/vs/workbench/contrib/chat/browser/actions/chatSpeechToTextActions.ts) |
| REACH streaming and tool transport | [Native request contract](../../src/vs/platform/native/common/native.ts), [shared REACH transport](../../src/vs/platform/native/common/reachChatTransport.ts), [regular chat provider](../../src/vs/platform/agentHost/node/reach/reachChatAgent.ts) |

## Rebuild the active checkout

Use Windows x64, the Node version in [`.nvmrc`](../../.nvmrc), and the pnpm version in [`package.json`](../../package.json). At this snapshot these were Node 24.22.1 and pnpm 12.16.0, with Electron 43.8.9. The prepared checkout already has the native runtime and built-in extensions. For a fresh checkout, install with the pinned pnpm using `pnpm install --frozen-lockfile`, then use `node build/next/build-fast.ts` after the [native prerequisites](NATIVE_BUILD.md) are ready. If the Windows pnpm shim fails, check which executable it resolves and verify a working pinned pnpm before running install or package tasks. The client-only command below runs directly under Node with the existing dependencies.

1. Check the branch and local changes. Stage a fresh client build. `transpile` cleans its selected output folder, so use a fresh directory under `.build`.

   ```powershell
   Set-Location D:\vscode
   git status --short
   $stamp = Get-Date -Format yyyyMMdd-HHmmss
   $stage = ".build/client-stage-$stamp"
   if (Test-Path -LiteralPath $stage) { throw "Staging directory already exists: $stage" }
   node build/next/index.ts transpile --out $stage
   if ($LASTEXITCODE -ne 0 -or !(Test-Path -LiteralPath "$stage/main.js")) {
       throw 'Client build failed or did not produce main.js'
   }
   ```

2. Rebuild any changed built-in extension as well. Client transpilation only builds `src/`. The Devin auth extension can be rebuilt and tested independently.

   ```powershell
   node node_modules/@typescript/native/bin/tsc -p extensions/devin-subscription/tsconfig.json --module Node16 --moduleResolution Node16
   if ($LASTEXITCODE -ne 0) { throw 'Devin extension compilation failed' }
   node --test extensions/devin-subscription/test/auth.test.cts extensions/devin-subscription/test/extension.test.cts
   ```

3. Before switching the active build, close the running Signal-C window gracefully and check that it has exited. Confirm the staged directory and `out` resolve inside `D:\vscode`. Move `out` to an unused backup directory under `.build`, then move the staged directory to `out`. If the second move fails, restore the previous `out` before launching. Keep the profile, workspace changes, and local configuration in place.
4. Open **Signal-C** from Start. A restart is required for main process changes. If the shortcut or protocol association has drifted, verify their current targets, preserve any displaced Signal-C shortcut and registry entry for rollback, and reinstall the primary registration.

   ```powershell
   & .\scripts\install-signal-c-shortcut.ps1 -Primary
   node scripts/code-agent-workspace.ts --primary --dry-run
   (Get-Item -LiteralPath 'HKCU:\Software\Classes\signal-c\shell\open\command').GetValue('')
   ```

   Start should target `wscript.exe` with the checkout's `.build\signal-c-primary-shortcut.vbs`. The protocol command should call the checkout's `scripts\code-agent-workspace.ts --primary --open-url "%1"`. Both paths should open the same executable and profile. An old executable paired with the new shortcut's VBS argument is a broken shortcut.

## Build an optimized Windows package when requested

Use the x64 Windows SDK tools described in [startup and packaging notes](STARTUP_PERFORMANCE.md). The repository's full task is `pnpm run gulp vscode-win32-x64-min`. For a prepared build with unchanged compiled extensions, the core-only route is:

```powershell
node build/next/index.ts bundle --minify --nls --mangle-privates --out out-vscode-min --target desktop
if ($LASTEXITCODE -ne 0) { throw 'Desktop bundle failed' }
pnpm run gulp vscode-win32-x64-min-ci
```

The normal task output is `D:\VSCode-win32-x64`. Smoke-test the package with isolated storage before considering a change to the normal Start menu entry. Keep the checkout as the primary runtime unless the user asks to switch.

- **Outlook registration.** The primary checkout reads a local, ignored `product.overrides.json` containing the public `signalCOutlookOAuthClientId`. Keep the approved value in local configuration. An optimized package needs that field in `resources/app/product.json` through its packaging or staging step. If the field is dropped, Outlook asks for a client ID. Check only presence and agreement with the approved registration. Keep credentials and callback parameters out of source, logs, and handoffs.
- **Claude and Codex.** A standalone build needs its pinned Windows runtimes under `resources/app/agent-sdks` with the required completion markers, or supported download metadata. Pins live under [`build/agent-sdk/agents`](../../build/agent-sdk/agents). The checkout's optimized preview launcher can use locally installed SDKs. Verify both provider entries from the final executable.
- **Dependencies and integrity.** Verify the built-in extensions, Windows native dependencies, and recorded product checksums in the assembled package. Run the terminal smoke check when changing node-pty or the dependency archive. Confirm the Assistant controls in the packaged window before installing over a working build.
- **November 27 installer checkpoint.** `D:\vscode\.build\win32-x64\user-setup\SignalCSetup.exe` was built and silently installed to `%LOCALAPPDATA%\Programs\Signal-C`. The staged executable is `D:\VSCode-win32-x64\Signal-C.exe`. The package passed `--optimized --empty-workspace --keyboard --app-connections` with report `.build/agent-workspace-native-smoke/2026-11-27T10-21-51-096Z-60664/report.json`; the installed product retained the approved Outlook registration. The local installer is unsigned. Document Copy/Export actions were added after that installer snapshot.
- **November 27 second installer.** `SignalCSetup.exe` was rebuilt (~328 MB) with the full Output-picker lifecycle (new-session composer + ongoing chat picker, `sendContext` draft persistence, Send Later/scheduled-edit capture, request-edit restore with `agentHostMetadata` fallback, `vscode.sessionOutputType` request metadata). Packaging steps: `node build/next/index.ts bundle --minify --nls --mangle-privates --out out-vscode-min --target desktop`, then `gulp vscode-win32-x64-min-ci`, merge `product.overrides.json` into `D:\VSCode-win32-x64\resources\app\product.json`, `gulp vscode-win32-x64-inno-updater` (required before user-setup or Inno fails on missing `tools\*`), then `gulp vscode-win32-x64-user-setup`. Invoke gulp directly as `node --experimental-strip-types --max-old-space-size=8192 ./node_modules/gulp/bin/gulp.js <task>` because the pnpm shim is broken, and prepend the Windows SDK signtool dir `C:\Program Files (x86)\Windows Kits\10\bin\10.0.26100.0\x64` to PATH or packaging fails with `spawn signtool.exe ENOENT`. Silent install completed with `/VERYSILENT /NORESTART /MERGETASKS=!runcode`; installed bundle MD5 matches the staged package and contains `vscode.sessionOutputType`, `sessions.newChat.outputType`, and the picker NLS strings.

## Verify before handing off

```powershell
node --test scripts/code-agent-workspace.test.ts
node test/unit/browser/index.js --browser chromium --run src/vs/sessions/contrib/sessions/test/browser/assistantSession.test.ts --reporter dot
git diff --check
```

For a desktop check, open the build through the entry point the user will use. Confirm **Agents → Assistant** shows **Worktree**, **New task**, and **Agent Chats**. In an isolated folderless chat, pick a small Git project with two worktrees and confirm that selecting one updates the composer label. After a launcher change, verify the actual running executable and profile and send a synthetic `signal-c://` return link through Windows. Check real provider login separately when the user can complete consent.

Local evidence on this machine:

- `.build/signal-c-assistant-controls-proof-20261119/assistant-controls-report.json` for the folderless composer and linked-worktree check.
- `.build/signal-c-cutover-20261118-234951/cutover-report.json` for the Windows callback and shortcut checks.
- `.build/signal-c-cutover-20261118-234951/microsoft-sign-in-report.json` and `provider-picker-report.json` in the same directory for provider entry-point checks.

## Agent-host unit tests

Run a targeted suite with the source path (not the out/ path):

```
node test/unit/node/index.ts --run src/vs/platform/agentHost/test/node/reachChatAgent.test.ts
```

Tests execute from compiled output in `out/vs/platform/agentHost/test/node/`. After editing a test or the code under test, retranspile (`node build/next/index.ts transpile --out .build/client-test-XXX`) and copy the compiled test file into `out/` — or do a full swap if production code changed. A mixed/stale `out/` causes confusing deterministic test failures that do not reflect the real code state.

## REACH/Cohere chat pipeline — failure-mode map (2026-11-25)

The REACH and Cohere providers share `runReachChatToolTurn` (src/vs/platform/agentHost/node/reach/reachChatToolTurn.ts). Key mechanics when debugging agent-chat failures:

- **Inkling output cap**: inkling-2-mini defaults to 4096 output tokens server-side. The client raises it to 16384 for every chat round (reachChatTransport.ts). Reasoning consumes the same budget — output-limit deaths mid-turn usually mean this cap regressed.
- **Edit strictness**: reach_edit_file oldText must match exactly once — unless `nearLine` is given, which picks the occurrence nearest that line (reachWorkspaceFiles.contribution.ts). nearLine is the primary ambiguity escape for small models: the ambiguous error teaches it with a concrete example (`nearLine: N picks the occurrence at line N`), and the circuit-breaker steering reinforces it after 2 failures. Repetitive/duplicated code makes the widen-oldText strategy nearly impossible for small models — nearLine is the only practical disambiguation. Errors are differentiated: not-found vs ambiguous (exact) vs ambiguous-tolerant (whitespace-normalized). Whitespace-drifted passages locate tolerantly when unique. Revision-mode edits (fileRevision from a read) are positional and never fail on matching; successful edits return a fresh post-edit revision so chained edits skip re-reads. Ranged revisions validate up to 256KB sections; full-file revisions up to 256KB files.
- **Read caps**: file reads pass up to the provider's read budget (REACH 64KB, Cohere 64-128KB). Larger files become excerpts with startLine/endLine/totalLines/fileRevision plus a readGuide teaching ranged reads. Terminal-spam dedup collapses 3+ consecutive identical lines.
- **Round caps**: only write-free rounds count — rounds with successful writes never approach the ceiling. 80 write-free rounds = graceful exit: tools are stripped and the model writes a completion summary (no error death). Anti-paralysis steering fires after 8 write-free rounds on modification tasks. The stall guard (32 identical-result rounds) exits with a reply.
- **Follow-up history**: completed actions persist per turn (`actions` in SavedTurn) and replay in follow-ups — thin replay causes re-exploration loops.
- **Tests**: reach 47 + cohere 70 in the node runner. Workbench tool tests (reachWorkspaceFiles.test.ts) are electron-browser — excluded from the node runner; verify by transpile.


The November 27 Assistant parity update added composer image inputs, a personalized home greeting, and routine `@`-mentions with completion and routine-scoped on-demand app activation. Its focused Assistant Chromium suite passed 91 tests. The same check fixed a customization-gallery `type is not defined` regression; the isolated native `--empty-workspace --keyboard --app-connections` smoke then passed with every app-connection fixture and both folderless round trips verified. Broader TypeScript checking still reports unrelated pre-existing working-tree errors outside the Assistant files.
