# Signal-C Assistant — Agent Capabilities

The Assistant view (`Assistant` in the Sessions sidebar, `OPEN_ASSISTANT_ACTION_ID`) is a conversational
agent backed by the user's REACH Subscription. Inference, account identity, subscription entitlement,
model catalog, and request allowance all stay on the existing SignalREACH path — the agent loop only
extends *what* a single prompt can accomplish.

Inspired by Town's assistant (town.com), the following agent capabilities are implemented on Signal-C
substrate:

## Multi-turn threads

- One root conversation per verified REACH account, plus linked **task** conversations
  (New task button, `/task <description>`, or a task row click).
- The empty home thread uses a time-aware greeting. When the user profile contains an explicit
  "my name is…" or "call me…" preference it uses that name, and the greeting surfaces local
  need-to-know counts for running linked tasks and important-mail alerts without starting any
  connected apps.
- Turns are persisted per account (`AssistantConversationStore`), including per-thread model choice,
  draft text, permission mode, and always-allowed tools.
- Task rows can be pinned, renamed, and deleted (with confirmation); running tasks show live status
  and a one-line completion update on the parent thread. Tasks interrupted by a restart show a
  resume button that replays their last prompt. When a background task finishes (or fails) while
  you're on another thread, a status-line nudge announces it so completed work isn't missed.
- Turns carry timestamps (relative display, exact time on hover) and record their run
  duration, shown in the turn meta as "worked 12s" (minutes/hours when longer). Sending
  while a run is in flight
  queues the message — several can stack (position is announced); they send FIFO as each reply
  finishes. The "Queued: …" preview shows the head message plus a "(+N more)" count with edit
  (pull head into composer) and remove actions. Stop or Escape clears the queue and cancels the run.
- The pending card shows the live status with elapsed seconds ("Working (step 2)… · 12s"), streamed
  reply text, and per-tool activity chips (kind icon, label, colored status icon). Completed
  calls ≥1s show an inline duration ("2m 5s" via the same formatter as turn duration), and every
  chip's tooltip carries `label — status · duration`; durations persist on saved turns.
- When dictation is configured and the microphone button is enabled, use it to dictate into the Assistant draft. Send waits for the final transcript. Switching chats or accounts releases dictation for the old draft.
- A completed reply plays `AccessibilitySignal.chatResponseReceived` (stopped and failed
  runs don't), honoring the user's `accessibility.signals.chatResponseReceived`
  sound/announcement setting. A new approval request plays
  `chatUserActionRequired` (deduped per request id — re-renders don't re-signal), and a
  background task ending on a non-open thread plays `taskCompleted` or `taskFailed`
  alongside its status-line nudge.
- **Regenerate from here** sends the selected request with its earlier conversation. Regenerating a reply or editing an earlier prompt keeps the original replies until the replacement succeeds. If the saved transcript has changed by then, Assistant keeps the newer transcript.
- **Open as document** moves any reply into a pinned Markdown editor as an untitled document. It keeps the conversational copy in the thread while giving long reports, drafts, and plans a full detail surface that can be edited or saved like a normal document. Document-like replies also get a titled artifact card with a word count.
- **Assistant documents** are persistent Markdown artifacts stored per account. When a deliverable is substantial, the model is directed to discover `assistant.create_document`, save the full artifact, and keep the chat reply concise. Follow-up revisions can use `assistant.read_document` and `assistant.update_document` without losing the document identity. Created artifacts appear as cards in the transcript; `/documents` lists the library, opens the current saved copy beside the conversation in a pinned Markdown editor whose edits are written back after a short pause, copies it, exports it as a Markdown file, or deletes it after confirmation. While that document remains open, each new prompt in the same thread carries as much of its current content as fits in the fixed-size request, so follow-up requests such as “make the introduction shorter” refer to the visible artifact. The local library holds up to 50 documents of 64,000 characters each.

## Routine mentions

Typing `@` in the composer offers the account's saved routines by name (disabled routines are labeled as such). Accepting a completion inserts the exact routine name; a submit whose `@Routine Name` matches a saved routine case-insensitively (flexible internal whitespace) invokes that routine's saved instructions as a hidden attachment and scopes the run to the routine's selected app connections. Any extra words after the mention become the immediate request, while the mention itself is removed from the visible typed text. Email-style text such as `person@example.com` is not treated as a mention. Routine-scoped requests follow the same queueing, task-linking, image, and retry behavior as ordinary sends, and they do not bulk-start unrelated MCP servers.

`/routine` → **New routine** now opens a small Town-style template gallery before the editor. Templates include Morning briefing, Meeting briefing, Inbox triage, GitHub report, Newsletter digest, and Relationship reconnect. Choosing one pre-fills the name, instructions, cadence, and time; matching connected apps are preselected by label when available, and every field remains editable before the routine is saved. **Start from scratch** keeps the original blank flow.

## Saved ideas

Use **Save reply** to keep the original Markdown. Highlight part of a reply before clicking **Save reply** to keep just the visible words, or enter `/save <text>` to capture your own idea. Each entry records its source chat and the time you saved it. Clearing a conversation or deleting its linked task preserves the saved copy.

Open **Saved ideas** in any Assistant chat under the same REACH account, or enter `/saved <words>` to search passage text and source titles. **Attach to draft** keeps your existing message and attachments. You can also copy the passage or remove it from the collection. Saving, searching in the picker, and attaching work locally. The collection uses the local profile and supports 100 entries of up to 8,000 characters each.

Choose **Continue in task** to give an idea its own working chat. The first time, Assistant opens a linked task with the passage attached. Later, the same action restores that task’s conversation and unsent draft. You decide when to send. If you delete the working task, you can start a fresh one from the saved passage.

For conversational recall, try “What ideas have I saved about copper?” or “Using my saved ideas about copper, suggest a prototype.” Assistant searches the account’s collection for the requested words and gives the selected model excerpts from the matching ideas. The model can read more pages of those matches. The answer’s **Source chats consulted** line identifies the chats whose passages were shown. Routine runs and unrelated requests do not receive the collection.

## Tool calling

Each model request can both reply and request tool calls; `AssistantRuns` loops up to 6 steps,
executing calls and feeding results back until the model produces a final answer.

- **Native function calling**: when the selected REACH model supports `tools`, the catalog is sent as
  OpenAI `tools` specs and `tool_calls` replies are executed directly (echoed back as
  `assistant` + `tool` messages).
- **Text fallback** uses a fenced `tool_call` block for models that return a call as text. Models declared text-only receive the tool index with their prompt. Results return in a bounded `<tool_result>` message.
- When the user supplies notes, a transcript or other text to rewrite, Assistant can answer from the source and their saved response preferences. It defers app startup and initially offers tool discovery plus skill search when a workflow hasn't already been attached. A request to retrieve source material from an app can still use that connection.
- **Loop guards**: max 6 model steps, max 4 calls per step, same-call-signature breaker, and a final
  "summarize without tools" step if the budget is exhausted.

### Tool sources

- **MCP connections**: enabled `IMcpService` servers contribute their model-visible cached tools
  (`McpToolVisibility.Model`), namespaced by `tool.id` and labeled `Server · Tool`, without requiring
  every server process to run. Interactive chats do **not** bulk-autostart the configured MCP set.
  Naming a configured app in the prompt activates just that connection before the first model request;
  if its tools are cold or incomplete, `assistant.find_app_tools` can activate the matching connection
  on demand. Actual MCP tool execution also starts its owning server through the normal MCP lifecycle.
  Scheduled routines still pre-start only the sources selected for that routine. MCP annotations
  classify safety: `readOnlyHint && !destructiveHint` counts as a read-only action. Each enabled
  server's `serverInstructions` (capped at 2k chars) is injected into the system prompt under
  "Guidance published by connected tools", so app-provided usage hints reach the model. App tiles show
  inactive enabled connections as **On-demand**; their menu can start, stop, enable, or disable the
  connection through the existing MCP enablement model.
- **Built-ins** (all workspace access is confined to workspace folders):
  - `assistant.get_current_time` — get the local date, weekday, time, and UTC offset for the device or a requested time zone, such as `Asia/Kathmandu`.
  - `assistant.ask_user` — pause the run and ask a clarifying question (optional answer
    choices); feeds the answer back as a tool result. Refused inside background tasks.
  - `assistant.update_plan` — post a live checklist (≤8 steps, pending / inProgress /
    done) rendered in the pending card while the run executes, for multi-step work.
    The final plan persists on the completed turn, shown as a collapsed
    "Plan — N/M done" details block, and survives conversation reloads.
  - `assistant.memory` — add, edit, or remove durable facts in the account's user profile or Assistant notes.
  - `assistant.search_conversation` — find earlier exchanges in the current chat by a word, phrase, or local calendar date range. For example, ask “What decisions did we discuss yesterday?” Each page contains up to five short excerpts and a continuation when earlier matches exist. Attachment bodies are omitted.
  - `assistant.find_skill` — search enabled skills by name or description and load a matching workflow when helpful. Long instructions are available in pages. Manual-only skills remain available through the composer.
  - `assistant.list_tasks` / `assistant.read_task` — list linked tasks and read a task's
    status + latest reply, so the parent thread can orchestrate background work.
  - `assistant.create_task` — start a linked task conversation in the background
    (side-effect: asks in Auto mode).
  - `assistant.rename_task` / `assistant.report_progress` — a running task can retitle itself
    and post one-line progress updates the user sees in the task list.
  - `assistant.list_files`, `assistant.read_file`, `assistant.find_files`,
    `assistant.search_files` — read-only workspace access via `IFileService`/`ISearchService`
    across every workspace folder (multi-root paths use `folderName/path` form).
  - `assistant.write_file`, `assistant.edit_file` — create/overwrite or exact-span replace in
    workspace files (side-effect: asks in Auto mode, refused in Safe-only). Results report
    change stats — "Wrote path: N lines, M characters" / "Edited path: +N −M lines" — which
    also surface on the activity chip's detail line.
  - `assistant.web_fetch` — read-only page fetch via `IRequestService`.
  - `assistant.open_url` — open an http(s) URL in the user's default browser via
    `IOpenerService` (side-effect: asks in Auto mode).
  - `assistant.open_file` — reveal a workspace file in the editor via `IOpenerService`;
    same workspace confinement as the other file tools (side-effect: asks in Auto).
    Optional `line`/`column` args jump to a position (`withSelection` URI fragment).
  - `assistant.list_documents`, `assistant.read_document`, `assistant.create_document`,
    `assistant.update_document` — manage persistent Markdown documents. These are
    discovery-only built-ins so they do not displace core tools in the first 32-tool batch;
    the system prompt directs document work through `assistant.find_app_tools` by exact name.
- Each model request carries up to 32 tool definitions. Assistant can find a connected app's other tools with `assistant.find_app_tools`, select a page of matching tools or load a tool by exact ID. `/tools` shows the full available catalog. Tool results are truncated at 6 000 characters and each call times out after 60 s.

## Approvals and permission modes

Per-thread permission mode (shield button):

| Mode | Read-only tools | Side-effect tools |
|---|---|---|
| Auto (default) | run immediately | inline approval card — Allow once / Always allow / Reject / Reject with note… |
| Safe only | run immediately | auto-declined |
| Allow all | run immediately | run immediately |

- "Always allow" records the tool in the thread's `allowedTools` list and skips future prompts;
  the permission picker's "Allowed tools…" lists every side-effect tool in the
  catalog (MCP-tagged) with checkboxes — check to pre-authorize, uncheck to revoke.
- MCP calls also offer **"Always allow {App} tools"** — merges all of that server's
  model-visible tools into `allowedTools` at once (Codex's "trust app"), then allows
  the pending call. Revocable through the same "Allowed tools…" picker or the tile's
  **Trust all tools / Require approval again** toggle in its context menu.
- Every call is rendered as an activity chip under the pending run and on the completed turn
  (pending → awaitingApproval → running → done / rejected / failed).
- Stop cancels the in-flight request *and* any waiting approval; pending approval requests
  survive thread switches (`getPendingApproval` restores the card).

## Connected apps

A tile strip above the thread lists every MCP server (`IMcpService.servers`) with its icon,
tool count and live state (Connected / Connecting / Needs sign-in / Failed / Disabled).
Click a tile marked **Sign-in required** to start authorization, or click a connected app to manage it. The trailing **Connect App** tile opens the app catalog. Right-clicking a
tile opens a context menu with Options…, Start/Stop (state-dependent), Restart, Show
Output, and Connect app, plus content actions scoped to that server: **Use prompt…**
(enabled when the server publishes prompts), **Attach resource…**, and **List tools…**
(safety-badged), and **Pin app**/**Unpin app**. Pinned servers float to the front of the
strip (then state rank within the group) and show a pin glyph next to the name; pins
persist in profile storage (`sessions.assistant.pinnedServers`) across sessions and
accounts. The strip updates live as servers start, stop, or are enabled.

## Memory

- `/memory` opens the account's saved profile and Assistant notes. Select an entry to edit or remove it, or add a new fact. A batch of model edits must fit within the selected memory bank's character limit before any change is saved.
- `/forget` lets you uncheck individual entries in either bank. Notes added or edited while the picker is open are preserved.

## Attachments and commands

- With skills enabled, choose **Use a skill…** in an empty home conversation or enter `/skill` to browse. A named command such as `/skill meeting-follow-up Summarize these notes: Mira will send the draft Friday.` attaches the instructions and leaves the request editable. Send when ready. Enabled user and plugin skills are also available. Workspace skills come from the selected trusted project or worktree.
- Three bundled skills support different stages of work. `idea-exploration` develops a rough idea into distinct options and a small experiment. `decision-brief` compares alternatives against the criteria you supply. `meeting-follow-up` turns notes into decisions, agreed actions, and open questions.
- To ask about a workflow, try “Read the installed meeting-follow-up skill. What does it say to do when a deadline is missing?” Assistant reads the eligible skill before answering. For a multi-page skill it can request more instructions from that same skill.
- To start a linked task with the material already attached, enter `/task` followed by the task instructions. The task receives the attachments, and its title comes from your instructions.
- `/attach <path> [path2 …]` inserts workspace files into the composer as
  `<file path="…">…</file>` blocks (each truncated at 12 000 chars; confined to workspace
  folders). `/attach` alone opens a multi-select workspace file picker (`ISearchService`
  file query, multi-root aware — pick entries are prefixed `rootName/` when more than one
  folder is open).
- Composer attachments (drag-and-drop or the attach control) are accepted on send:
  `URI`-valued entries are read through `IFileService` under the same workspace-folder
  confinement and 12k-char truncation as `/attach` and wrapped in `<file>` blocks;
  string-valued entries (e.g. pasted text) are wrapped in `<attachment>` blocks; image
  entries are resized by the shared composer, converted to vision inputs, and sent with
  the request instead of being flattened to text. Up to four PNG, JPEG, GIF or WebP
  images (12 MB total) are accepted per turn; unsupported or excess items are skipped
  and reported. Consumed attachments are cleared from the composer; if nothing readable
  was attached the send aborts with a hint to paste the content instead. Sent turns render
  `<file>`/`<attachment>`/`<resource>` blocks as labeled chips (icon + basename, full name
  on hover) and images as thumbnails above the message text instead of raw markup;
  "Edit prompt in composer" still replays the full block text so attachments stay editable.
  Image inputs are stored with their turn, replay in follow-up context, and are inherited
  by retry/regeneration when no replacement image is supplied.
- `/task <description>` creates and opens a linked task conversation; `/tasks` opens a
  picker of saved tasks (status, title, latest update) to jump between threads.
  Untitled tasks take their title from the first prompt sent (the model can refine it
  via `assistant.rename_task`).
- `/connect` opens the app catalog (`McpCommandIds.ConnectApp`) — same as the "Connect
  app" tile.
- `/prompts` lists prompt templates published by connected apps (`IMcpServer.prompts`),
  asks for any declared arguments, resolves the template, and drops the result into the
  composer for review before sending.
- `/resources` lists readable content published by connected apps
  (`IMcpServer.resources()`), multi-selects entries, reads them through `IFileService`
  (32KB cap, 12k-char truncation each), and wraps them as `<resource>` blocks in the
  composer — the app-content equivalent of `/attach`.
- `/model` opens the REACH model picker with the active choice selected. Models with `allow_tools: false` use text-based tool calls. Models with native support or an unspecified capability receive the native tool definitions.
- The approval card shows a collapsible Arguments block (auto-expanded under 200
  chars), so long tool payloads don't flood the card.
- "Reject with note…" asks for an optional reason which is relayed to the model as
  `Reason: <note>` in the tool result and shown on the declined activity chip — the
  run continues with the guidance instead of stopping.
- After a failed run a "What went wrong?" follow-up chip offers to diagnose the
  failure instead of retrying blindly.
- `/tools` lists the current tool catalog with safety hints; `/clear` resets the thread.
- Failure status lines carry an inline "Retry" action link that resends the failed
  prompt — `/retry` remains the keyboard path. The feedback line supports action
  links generally (message + link).
- `/search <text>` filters the transcript to turns whose prompt or reply contains the
  text (case-insensitive, attachment block contents included); the status line reports
  the match count and `/search` alone clears the filter. `/export` copies the
  transcript as Markdown. `/copy` copies the latest reply and `/retry` resends the failed prompt.
- `/regenerate` requests another answer to the last prompt. `/undo` removes the last turn and restores its prompt for review. `/context` shows how much of the retained conversation the model will see.
- `/compact` saves a summary and collapses earlier turns in the normal view. Future replies carry the latest summary and recent exchanges. Transcript search, dated conversation recall, and export can still reach the original turns. Undoing a summary reveals the earlier conversation.
- `/stop` or Escape cancels the run. `/help` opens a command picker that fills the composer.
- Replies render as Markdown while streaming. Reply actions include copy, quote, edit and regenerate. Editing an earlier prompt shows which part of the conversation a successful replacement will change. **Cancel edit** returns to ordinary chat.
- Suggested follow-up chips appear under the last reply based on which tools ran
  (file edits → "Summarize the changes", task creation → "Check on my tasks", searches →
  "Find related files").

## Suggestions

The empty home thread shows server-aware suggestion chips (e.g. email summary when a mail MCP
server is connected) plus a "What can you do?" discovery prompt. Servers that publish
prompts also get a "{Server} prompts…" chip that opens the server-scoped template picker.
Clicking a suggested question sends that question and keeps your existing draft and its attachments in the composer.

## Protocol

`ReachAssistantRequest` = `{ messages, modelId?, tools?, textToolIndex?, isTask? }` where messages may be
`system | user | assistant | tool` and assistant messages may carry `toolCalls`
(`{id, name, arguments}`). `ReachAssistantResponse.reply` may carry parsed `toolCalls`.
The account/subscription/allowance gate in `requestReachAssistant` is unchanged; busy protection
is keyed per `[accountId, conversationId]` and stop aborts by request id per window.

Native tool-capable models receive tool descriptions in the function definitions. Text-only models receive the tool index in the system message. Multiple text tool results share the 8,000-character message budget while preserving short results.

Replies stream (`stream: true` upstream): SSE deltas travel main → renderer through
`INativeHostService.onDidReceiveReachAssistantChunk` (keyed by request id), accumulate in
`IAssistantPendingRun.partial`, and render as live text in the pending card. Streamed
`tool_calls` fragments are reassembled by index. Backends that ignore `stream` fall back
to the normal JSON path automatically (the stream parser only runs on `text/event-stream`
responses).

Remaining Town comparison items include document collaboration, richer conflict handling and selection-aware editing, a code sandbox, and shared team threads.
