# ADR-0001: Split the shared taskpane into a shared shell plus per-host taskpanes

- **Status:** Proposed
- **Date:** 2026-08-29
- **Author:** C. Schiele (github:shazbot996, ldap:schiele) — Excel adapter track
- **Affects:** `microsoft-addin/src/taskpane/*`, `microsoft-addin/webpack.config.js`, `microsoft-addin/manifest*.xml` (optional, Phase 3)
- **Does not affect:** `geminiproxy/`, `microsoft-addin/src/core/geminiClient.js` wire protocol

---

## 1. Context

### 1.1 Where we are

`microsoft-addin/` ships **one** taskpane that serves all three Office hosts:

```
src/taskpane/taskpane.html   141 lines   one DOM for Word + PowerPoint + Excel
src/taskpane/taskpane.js     499 lines   Office.onReady, chat state, chips, insertion
src/taskpane/taskpane.css      79 lines   orphaned — nothing imports or links it
src/adapters/{Word,PPT,Excel}Adapter.js   host-specific document I/O
src/adapters/HostAdapterFactory.js        runtime host detection
src/core/{geminiClient,markdownParser}.js shared transport + rendering
```

`webpack.config.js` defines a single `taskpane` entry and a single `HtmlWebpackPlugin`
emitting `/taskpane.html`. `manifest.xml` points every host's
`<Action xsi:type="ShowTaskpane"><SourceLocation resid="Taskpane.Url"/>` at that
same URL.

The adapter layer is already a clean seam: `HostAdapterFactory.getAdapter()` returns
an object exposing `getSelectedText()`, `getFullDocumentText()`,
`insertContent(html, rawText, mode)`, and `checkInDocumentCommands(force, callbacks)`.
The Excel work described in [`excel-core-changes.md`](./excel-core-changes.md) landed
entirely behind that seam with zero core edits — which is exactly why the seam is
worth preserving and extending.

### 1.2 The problem

The seam stops at document I/O. **Everything above it is shared and host-branched.**
`taskpane.js` contains at least four separate `if (hostName === "Excel") … else if
(hostName === "PowerPoint") … else …` ladders:

| Location | What branches |
| :--- | :--- |
| `adaptUIForHost()` (L69–L122) | Banner text, section titles, 4 chip labels, welcome copy, status text |
| `setupDocToolsChips()` (L134–L184) | 4 chips × 3 hosts = 12 hardcoded prompt strings |
| `appendAssistantBubble()` (L410–L419) | Insert button label and tooltip |
| `performDocumentInsertion()` (L472) | Loading message |

The HTML is worse than branched — it is *fixed*. There are exactly four
`#chipSummarize`/`#chipRisks`/`#chipActionItems`/`#chipExecBox` slots and exactly five
`#selRewrite`…`#selTable` selection chips. A host cannot add a fifth doc chip, remove
one it does not need, or introduce a control of a different shape (a range picker, a
formula preview strip, a target-cell indicator) without editing markup that Word and
PowerPoint also render.

Concretely, the three Excel enhancements already recommended in
`excel-core-changes.md` §3 are all blocked on this:

1. Relabel the insert button to **"➕ Insert Formula"** when the response contains a
   detected formula → requires editing `appendAssistantBubble()`, shared code.
2. Add **✨ Formula Generator**, **🔍 XLOOKUP Builder**, **📊 Data Anomaly Scan** chips
   → requires adding elements to the shared `taskpane.html`.
3. Show the resolved target cell (`E2`) and the inferred column schema before the user
   commits an insertion → no slot exists in the shared DOM at all.

### 1.3 The organizational constraint

Three teams currently hold **duplicate working copies of `microsoft-addin/`**. Every
edit to `src/taskpane/*` is a guaranteed three-way merge conflict and a
cross-team coordination event. The current design therefore has the worst possible
property: the files that each team most needs to change are the files every team
shares. Ownership is undefined precisely where churn is highest.

The `geminiproxy/` service is owned by another team and is out of scope here. Nothing
in this ADR changes the request/response contract in `geminiClient.js`.

---

## 2. Decision

**Split the taskpane into a shared *shell* and one *host taskpane* per Office host,
each with its own webpack entry, its own HTML document, and a single owning team.**

The shell owns everything that is genuinely common — chat transport, session state,
the message log, markdown rendering, error handling, boot sequence. Each host taskpane
owns its own DOM, its own copy, its own chips, and its own response actions, and
registers them with the shell through a declarative contract.

### 2.1 Target layout

```
microsoft-addin/src/
  core/                          # UNCHANGED · owner: core team
    geminiClient.js
    markdownParser.js

  shell/                         # NEW · shared kernel · owner: core team
    TaskpaneShell.js             # boot(hostModule); owns lifecycle + wiring
    ChatSession.js               # sessionId, history, executeWorkflow(), reset()
    ChatLog.js                   # appendUserBubble / appendAssistantBubble / appendSystem
    SelectionWatcher.js          # DocumentSelectionChanged → hostModule.onSelectionChange
    shell.css                    # bubbles, layout, loading, buttons (from taskpane.css)
    shell-frame.html             # the ~20 shared elements, included by each host page

  taskpanes/
    excel/                       # owner: Excel team (cbs)      ← our lane
      index.html                 # entry document; Excel-only markup
      index.js                   # exports the Excel host module
      excel.css
      chips.js                   # Excel doc-tool + selection chip definitions
      responseActions.js         # "Insert Formula" vs "Insert into Sheet", schema preview
    word/                        # owner: Word team
      index.html  index.js  word.css  chips.js
    powerpoint/                  # owner: PowerPoint team
      index.html  index.js  powerpoint.css  chips.js
    dispatcher/                  # NEW · owner: core team
      index.html                 # emitted as /taskpane.html — see §2.4

  adapters/                      # UNCHANGED ownership, one file per team
    ExcelAdapter.js WordAdapter.js PPTAdapter.js HostAdapterFactory.js
```

Build output becomes `/excel.html`, `/word.html`, `/powerpoint.html`,
`/commands.html`, and a compatibility `/taskpane.html`.

### 2.2 The host-module contract

Each host taskpane's `index.js` default-exports a plain descriptor. This is the entire
API surface between a host team and the core team:

```js
// src/taskpanes/excel/index.js
import { ExcelAdapter } from "../../adapters/ExcelAdapter.js";
import { docChips, selectionChips } from "./chips.js";
import { renderResponseActions } from "./responseActions.js";

export default {
  hostName: "Excel",                    // asserted against Office.context.host at boot

  strings: {
    tipBanner:   "💡 Type <code>@gemini &lt;prompt&gt;</code> in Cell",
    toolsTitle:  "📈 <strong>Chat with Spreadsheet:</strong>",
    welcome:     "Type a prompt, click a <strong>Sheet Chip</strong>, or select cells.",
    readyStatus: "Excel Ready",
  },

  createAdapter: () => new ExcelAdapter(),

  docChips,                             // [{ id, label, title, prompt }]
  selectionChips,                       // [{ id, label, instruction }]

  // Shell calls this after rendering an assistant response. The host decides
  // which buttons exist and what they say. Excel inspects the response for a
  // formula and swaps the primary action accordingly.
  renderResponseActions,                // (ctx) => HTMLElement

  // Optional per-host hooks. Shell no-ops if absent.
  onSelectionChange: async (selection, ui) => { /* Excel: show target cell + schema */ },
  onReady:           async (ui) => {},
};
```

And the shell entry, identical in all three host pages:

```js
// src/taskpanes/excel/boot.js  (three lines, one per host)
import { TaskpaneShell } from "../../shell/TaskpaneShell.js";
import excel from "./index.js";
TaskpaneShell.boot(excel);
```

`TaskpaneShell.boot()` performs `Office.onReady`, instantiates the adapter, renders
the declared chips into the host page's containers, attaches the selection handler,
and wires the send button to `ChatSession.executeWorkflow()`. Every `if (hostName ===
…)` ladder in today's `taskpane.js` disappears — it becomes data in the host module.

**Contract rule:** the shell may never branch on `hostName`. If the shell needs to
know which host it is running in, that is a signal the contract is missing a hook.
This rule is what keeps the shell from re-accreting host logic.

### 2.3 Webpack changes

The entry list and HTML plugins are derived from one array, so adding a host is a
one-line change and no host team edits another host's build config:

```js
const HOSTS = ["excel", "word", "powerpoint"];

entry: {
  polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
  commands: "./src/commands/commands.js",
  dispatcher: "./src/taskpanes/dispatcher/dispatcher.js",
  ...Object.fromEntries(HOSTS.map(h => [h, `./src/taskpanes/${h}/boot.js`])),
},

plugins: [
  ...HOSTS.map(h => new HtmlWebpackPlugin({
    filename: `${h}.html`,
    template: `./src/taskpanes/${h}/index.html`,
    chunks: ["polyfill", h],
    minify: false,
  })),
  new HtmlWebpackPlugin({
    filename: "taskpane.html",
    template: "./src/taskpanes/dispatcher/index.html",
    chunks: ["polyfill", "dispatcher"],
  }),
  // …commands.html, CopyWebpackPlugin unchanged
]
```

`nginx.conf` and the `Dockerfile` need no changes — the container already serves the
webpack output directory as static files.

### 2.4 The dispatcher: why the manifest change is optional

`/taskpane.html` is **kept**, and re-emitted as a ~15-line dispatcher:

```js
Office.onReady((info) => {
  const target = { [Office.HostType.Excel]:      "excel.html",
                   [Office.HostType.Word]:       "word.html",
                   [Office.HostType.PowerPoint]: "powerpoint.html" }[info.host] || "word.html";
  window.location.replace(target + window.location.search);
});
```

This preserves the query-string passthrough that `geminiClient.getActiveProxyUrl()`
relies on (`?backend=streamassist`).

The consequence is important for our constraint: **every already-sideloaded manifest
keeps working unchanged.** No user re-installs, no coordinated manifest release, and
critically, no edit to the shared `manifest.xml` is required to ship this. Pointing
each host's `<SourceLocation>` at its own URL becomes a later, purely-optional
optimization (Phase 3) that removes one redirect hop.

---

## 3. Options considered

### Option A — Status quo: one taskpane, runtime host branching

*Rejected.* This is the state described in §1.2. Zero migration cost, but every host
feature is a shared-file edit across three duplicated working copies, the HTML cannot
express host-specific controls at all, and the branch ladders grow superlinearly as
each team adds features. It already blocks all three Excel enhancements we have queued.

### Option B — One entry, host module loaded at runtime via dynamic `import()`

Keep a single `/taskpane.html` and a single entry; `Office.onReady` dynamically imports
`./taskpanes/${host}/index.js` and hands it to the shell.

*Rejected as the primary, retained as a fallback.* It gets us the ownership split in
JS and needs no manifest thinking at all, which is genuinely attractive. But the HTML
document stays shared and therefore stays a merge-conflict surface and a constraint on
host-specific layout — which is half of the actual problem. It also keeps three hosts'
CSS in one cascade, so a Word style change can visually regress Excel. If Phase 1
review objects to emitting multiple HTML documents, B is the graceful degradation:
the host-module contract in §2.2 is identical either way, so the decision is reversible.

### Option C — Per-host entries + per-host HTML + compatibility dispatcher (**chosen**)

Described in §2. Each host owns a directory containing its own document, script, and
stylesheet. Costs one extra HTML file and one extra entry per host, and roughly
doubles the shared-shell design work relative to B. Buys real file-level ownership:
after migration, an Excel feature touches only `src/taskpanes/excel/**` and
`src/adapters/ExcelAdapter.js`, which are files no other team's copy diverges on.

### Option D — Separate packages or repositories per host

*Rejected.* Correct destination if this ever becomes three separately-shipped add-ins,
but today all three hosts ship in one manifest, one container, and one Cloud Run
service. Splitting the repo now would force publishing `core/` and `shell/` as a
versioned internal package and would fragment a deployment that is currently
single-artifact. Revisit only if the hosts' release cadences actually diverge.

---

## 4. Consequences

### Positive

- **Ownership becomes file-level.** Excel work lands in `src/taskpanes/excel/**` and
  `src/adapters/ExcelAdapter.js`. Word and PowerPoint copies never diverge there, so
  the three-way merge problem on `taskpane.js` disappears for host feature work.
- **Host-specific UI becomes possible**, not just host-specific strings. The Excel pane
  can render a target-cell indicator, a formula preview with the resolved schema, and
  an anomaly-scan chip without any other host rendering dead DOM.
- **The shell gets simpler as it generalizes.** ~120 lines of branch ladder in
  `taskpane.js` become ~40 lines of data spread across three host modules.
- **Excel's queued enhancements unblock** — all three items in `excel-core-changes.md`
  §3 become Excel-team-local changes.
- **CSS isolation.** Each host loads `shell.css` plus its own sheet; a Word restyle
  cannot regress the Excel pane.
- **Payload shrinks per host.** Today every host downloads all three hosts' logic and
  adapters. After the split, `excel.html` pulls only `ExcelAdapter` (once
  `HostAdapterFactory`'s static three-way import is bypassed by the host module's
  `createAdapter()` — see §5, Phase 2).

### Negative / costs

- **This ADR proposes changing files outside the Excel lane** —
  `src/taskpane/*` and `webpack.config.js` are shared. That is precisely why this
  record exists, and why §5 sequences the change so that Word and PowerPoint keep
  working byte-identically at every step.
- **One-time coordination.** All three teams must re-sync their duplicated
  `microsoft-addin/` copies once, at the Phase 1 boundary. This is a scheduled cost
  paid once, replacing an unscheduled cost paid on every feature.
- **Three HTML documents to keep visually coherent.** Mitigated by keeping all shared
  chrome in `shell-frame.html` and `shell.css`; host pages should contain only their
  own controls.
- **A redirect hop** on taskpane open until Phase 3 lands. Local static redirect,
  effectively imperceptible, and removable.

### Neutral / notes

- `src/taskpane/taskpane.css` (79 lines) is currently **dead** — no `<link>` and no
  `import` references it; all live styles are inlined in `taskpane.html`. It should be
  reconciled into `shell/shell.css` during Phase 1 rather than carried forward
  unexamined.
- `commands.js` / `commands.html` are untouched by this ADR.
- No change to the `geminiproxy/` contract: `askGeminiEnterprise(prompt, history,
  sessionId, enableGrounding)` and its payload shape are unchanged. The other team
  needs no notification.
- Authentication is assumed to be handled by a separate layer and is out of scope.

---

## 5. Migration plan

Each phase is independently shippable and independently revertible.

**Phase 0 — Agreement (no code).** Circulate this ADR. The only thing all three teams
must agree on is the host-module contract in §2.2; everything else is mechanical.

**Phase 1 — Extract the shell, keep behavior identical.** Create `src/shell/**` by
lifting the host-neutral half of `taskpane.js` (chat session state,
`executeGeminiWorkflow`, `appendBubble`, `appendAssistantBubble`,
`performDocumentInsertion`, the selection watcher). Reconcile the dead `taskpane.css`
into `shell.css`. Create the three host modules by moving each branch of the existing
ladders into data. Emit `/excel.html`, `/word.html`, `/powerpoint.html` plus the
`/taskpane.html` dispatcher. **Acceptance: Word and PowerPoint behave exactly as they
do today; no manifest change; no adapter change.** This is the only phase requiring
cross-team review, and it is the re-sync point for the duplicated folders.

**Phase 2 — Excel feature work (Excel team only).** Land `excel-core-changes.md` §3
items 1–3 in `src/taskpanes/excel/**`. Replace the Excel page's use of
`HostAdapterFactory` with the module's own `createAdapter()` so the bundle stops
pulling `WordAdapter` and `PPTAdapter`. `HostAdapterFactory` stays for the dispatcher
and for any consumer not yet migrated.

**Phase 3 — Optional manifest optimization.** Add per-host `bt:Url` resids
(`Taskpane.Url.Excel`, `…Word`, `…PowerPoint`) and point each `<Host>`'s
`<Action><SourceLocation>` at its own page, dropping the redirect. Requires a manifest
version bump and re-sideload, so schedule it with a release rather than standalone.
`<DefaultSettings><SourceLocation>` keeps pointing at the dispatcher as the fallback
for hosts that ignore `VersionOverrides`.

---

## 6. Open questions

1. **Shared-chrome mechanism.** `shell-frame.html` as an html-loader partial included
   by each host page, or shell-rendered DOM injected into a single `<div id="shell">`
   mount? Injection keeps host pages smaller and prevents chrome drift; partials keep
   the markup greppable. Leaning injection; core team's call.
2. **Chip prompt ownership.** Prompt text is product copy, not code. Do the ~12 doc-chip
   prompts stay in each host's `chips.js`, or move to a per-host JSON that a PM can edit
   without a JS review?
3. **`hostName` assertion failure.** If `excel.html` is somehow opened in Word, does the
   shell hard-fail with a clear message or fall back to the dispatcher? Proposing an
   explicit error — silent fallback would mask manifest misconfiguration.
4. **Testing.** There is currently no test harness for the taskpane layer. The
   host-module contract is exactly the seam a smoke test should target (mock adapter +
   mock `askGeminiEnterprise`, assert chips render and workflow dispatches). Worth
   scoping as follow-up work.
