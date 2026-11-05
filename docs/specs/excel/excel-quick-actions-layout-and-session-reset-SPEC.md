---
type: spec
title: "Feature Specification: Excel Quick Actions Component Repositioning, Layout Deduplication & Session Reset Behavior"
scope: "microsoft-addin/src/taskpanes/excel/*, microsoft-addin/src/shell/*, microsoft-addin/src/taskpanes/*"
date: "2026-09-14"
---

# Feature Specification: Excel Quick Actions Repositioning, Layout Deduplication & Session Reset Behavior

## 1. Problem Statement

### 1.1 Background & Current Layout
In the Gemini for Microsoft 365 Excel Taskpane, user interactions are organized across several visual components:
1. **Top Target Range & Chart Tracker (`#excelSchemaPreview`):** Continuously tracks and displays the active target cell (e.g., `🎯 Target Cell: B2`), highlighted range (e.g., `📊 Target Range: A1:D50`), active structured table (e.g., `📑 Target Table: Transactions`), and sticky chart context (e.g., `📊 Chart: Revenue Summary (Pie)`).
2. **Yellow Quick Actions Box (`#selectionToolbar`):** A context-sensitive container rendering high-frequency action chips (e.g. *Insights*, *Exec Summary*, *Categorize Taxonomy*, *Visualize Selection*, *Style & Format*).
3. **Scrollable Chat Buffer (`#chatHistory`):** The conversational viewport displaying message bubbles (user queries, system notifications, markdown AI responses, and interactive action buttons).

### 1.2 Identified UI Flaws & Redundancies
* **Duplicate Cell Range Displays:** The yellow Quick Actions box previously included `#selectionPreviewText` and `#selectionWordCount` (e.g., `"12 words"`, `0 cells`, `"...range text..."`). Because the top tracker `#excelSchemaPreview` already tracks the active target cell and range in real time, displaying the range address and word count inside the yellow box introduced redundant visual clutter.
* **Fixed Bottom Positioning & Viewport Clutter:** Previously, `#selectionToolbar` was positioned statically outside and below the scrollable `#chatHistory` container (above the prompt input). In this location, it occupied permanent vertical screen real estate, compressing the active chat viewport and obscuring conversation history.
* **Inconsistent Reset Experience:** When clearing a chat session or starting a new topic, `#selectionToolbar` remained in a disconnected layout state rather than presenting the initial, clean starting view at the top of the scroll container.

---

## 2. Goals & Non-Goals

### 2.1 Goals
* **Eliminate Redundant UI Elements:**
  - **Retain Top Range Tracker (`#excelSchemaPreview`):** Preserve the dedicated top target bar with `#targetRangeBadge`, `#targetChartBadge`, and `#sheetSchemaBadge` for real-time tracking of cursor position, range bounds, and sticky chart targets.
  - **Clean Yellow Quick Actions Box (`#selectionToolbar`):** Remove duplicate cell range preview (`#selectionPreviewText`) and word/cell counts (`#selectionWordCount`).
* **Reposition Quick Actions to Top of Scroll Buffer:**
  - Move `#selectionToolbar` inside `#chatHistory` as its topmost child element directly beneath the top target tracker bar.
  - Allow the Quick Actions box to scroll naturally upward and out of view as conversation history progresses.
* **Session Reset & Scroll Top Behavior:**
  - When the user clicks "New Chat" (`#clearSession`) or initiates a session reset:
    - Clear conversation message bubbles.
    - Preserve/re-mount the yellow Quick Actions box as the topmost element in `#chatHistory`.
    - Reset the scroll buffer to the top (`scrollTop = 0`) so Quick Actions is immediately visible.
* **Cross-Host Alignment:** Apply consistent scroll-integrated Quick Actions layout across Excel, Word, PowerPoint, and the shared shell architecture.

### 2.2 Non-Goals
* Modifying backend prompt construction or LLM token schemas.
* Altering the execution logic of individual action chips (e.g. Insights, Charts, Taxonomy).

---

## 3. High-Level Architecture & Layout Progression

### 3.1 DOM Hierarchy Comparison

#### Before
```text
Taskpane Body
├── Header (.header: Title + "New Chat" button)
├── Top Range Tracker (#excelSchemaPreview)
├── Document Tools Bar (#docToolsBar - Collapsible)
├── Scroll Container (#chatHistory)
│   ├── Welcome System Bubble
│   └── User/Assistant Message Bubbles
├── [REDUNDANT] Yellow Box (#selectionToolbar: with Range Text & Word Count)
└── Prompt Input Container (.input-container)
```

#### After (Optimized Layout)
```text
Taskpane Body
├── Header (.header: Title + "New Chat" button)
├── Top Range Tracker (#excelSchemaPreview - Continuous Target Cell/Range Tracking)
├── Document Tools Bar (#docToolsBar - Collapsible)
├── Scroll Container (#chatHistory - overflow-y: auto)
│   ├── [REPOSITIONED & CLEAN] Yellow Quick Actions Box (#selectionToolbar)
│   │   ├── Title: ⚡ Quick Actions (#selectionTitle)
│   │   └── Chips: #selectionChipsContainer (Insights, Taxonomy, Charts, etc.)
│   ├── Welcome System Bubble (#welcomeSystemBubble)
│   └── User/Assistant Message Bubbles (scrolls naturally downwards)
└── Prompt Input Container (.input-container - Fixed at bottom)
```

### 3.2 Scrolling Lifecycle Progression

```mermaid
sequenceDiagram
    autonumber
    actor User
    participant Taskpane as Taskpane Viewport
    participant ScrollBuffer as #chatHistory Container
    participant QuickActions as Yellow Quick Actions (#selectionToolbar)
    participant ChatLog as ChatLog Subsystem

    Note over Taskpane,QuickActions: State 1: New / Reset Session
    User->>Taskpane: Opens Taskpane or Clicks "New Chat"
    ChatLog->>ScrollBuffer: Reset scroll buffer (scrollTop = 0)
    ScrollBuffer-->>Taskpane: Quick Actions visible at top of scroll buffer

    Note over Taskpane,QuickActions: State 2: Active Conversation
    User->>Taskpane: Clicks Quick Action Chip or Sends Prompt
    ChatLog->>ScrollBuffer: appendUserBubble(prompt)
    ChatLog->>ScrollBuffer: appendAssistantBubble(response)
    ChatLog->>ScrollBuffer: scrollToBottom()
    ScrollBuffer-->>Taskpane: Buffer scrolls down; Quick Actions scrolls naturally out of view

    Note over Taskpane,QuickActions: State 3: Session Reset
    User->>Taskpane: Clicks "New Chat" (#clearSession)
    ChatLog->>ScrollBuffer: clear() [Preserves #selectionToolbar, clears bubbles]
    ChatLog->>ScrollBuffer: scrollToTop() [scrollTop = 0]
    ScrollBuffer-->>Taskpane: Quick Actions immediately visible again
```

---

## 4. Technical Design & Implementation

### 4.1 HTML Template Refactoring (`index.html`)

```html
<!-- Excel Target Range, Chart & Schema Preview Bar (Retained at Top) -->
<div id="excelSchemaPreview">
    <span class="excel-pill target-range" id="targetRangeBadge">🎯 Target: Cell Active</span>
    <span class="excel-pill target-chart" id="targetChartBadge" style="display: none;">
        <span id="targetChartText">📊 Chart: Active Chart</span>
        <button id="removeTargetChartBtn" class="chip-close-btn" title="Remove chart from context" aria-label="Remove chart target">&times;</button>
    </span>
    <span class="excel-pill schema-info" id="sheetSchemaBadge">📄 Active Worksheet</span>
</div>

<!-- Scrollable Chat Buffer -->
<div id="chatHistory">
    <!-- Yellow Quick Actions Box (Repositioned at Top of Scroll Buffer) -->
    <div id="selectionToolbar" class="quick-actions-toolbar">
        <div class="selection-header">
            <span id="selectionTitle">⚡ Quick Actions</span>
        </div>
        <div class="chips-container" id="selectionChipsContainer"></div>
    </div>

    <div class="chat-bubble system" id="welcomeSystemBubble">
        Type a prompt below, click an action above, or select cells/charts in Excel!
    </div>
</div>
```

### 4.2 Shared Shell Stylesheet (`shell.css`)

```css
/* Yellow Quick Actions Box (Top of scrollable #chatHistory) */
#selectionToolbar,
.quick-actions-toolbar {
  display: block;
  background-color: #fff4ce;
  border: 1px solid #fce100;
  border-radius: 6px;
  padding: 8px 10px;
  margin-bottom: 6px;
  font-size: 11px;
  flex-shrink: 0;
}

.selection-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 6px;
  font-weight: 600;
  color: #795b00;
  font-size: 11px;
}

.selection-preview {
  display: none;
}
```

### 4.3 ChatLog DOM Engine (`ChatLog.js`)

```javascript
export class ChatLog {
  constructor(container = "#chatHistory") {
    this.container = typeof container === 'string' ? document.querySelector(container) : container;
  }

  /**
   * Resets chat history, preserving top Quick Actions toolbar,
   * rendering the welcome message, and resetting scroll to top.
   * @param {string} welcomeHtml 
   */
  clear(welcomeHtml = null) {
    if (!this.container) return;
    const selectionToolbar = this.container.querySelector("#selectionToolbar, .quick-actions-toolbar");
    this.container.innerHTML = "";
    if (selectionToolbar) {
      this.container.appendChild(selectionToolbar);
    }
    if (welcomeHtml) {
      this.appendSystem(welcomeHtml, true);
    }
    this.scrollToTop();
  }

  /**
   * Resets scroll position to the top of the container.
   */
  scrollToTop() {
    if (this.container) {
      this.container.scrollTop = 0;
    }
  }

  /**
   * Scrolls chat history viewport to the bottom.
   */
  scrollToBottom() {
    if (this.container) {
      this.container.scrollTop = this.container.scrollHeight;
    }
  }
}
```

### 4.4 Document Selection Watcher (`SelectionWatcher.js`)

```javascript
// Remove duplicate text and word count updates; keep toolbar persistently mounted in scroll buffer
if (this.currentSelectedText.length > 0) {
  // Target tracking is delegated to top #excelSchemaPreview bar
}
```

---

## 5. UI / UX Validation & Expected Behaviors

| Feature State | User Action | Expected UI Behavior |
|---|---|---|
| **Initial Load** | Taskpane boots | `#excelSchemaPreview` displays active cell; `#selectionToolbar` is displayed as topmost item in `#chatHistory`; viewport shows Quick Actions. |
| **Active Chat** | User submits prompt or executes chip | Assistant message appended; `#chatHistory.scrollTop` animates to bottom; Quick Actions scrolls up out of view. |
| **Scroll History** | User scrolls up in chat log | Conversation history scrolls down; yellow Quick Actions box is accessible at top of scroll buffer. |
| **New Chat** | User clicks "New Chat" | Messages cleared; Quick Actions box preserved at top of `#chatHistory`; scroll buffer resets to `scrollTop = 0`. |
| **Cell Selection** | User changes selection in Excel | Top `#excelSchemaPreview` updates badge immediately (e.g., `📊 Target Range: A1:D50`); no duplicate text rendered inside Quick Actions box. |

---

## 6. Testing & Validation Strategy

1. **Automated Unit Tests ([`microsoft-addin/tests/excel.test.mjs`](file:///Users/ravindranv/Code-Local/For-Excel-GE-Add-In/excel_pod_release/gemini-for-office-365/microsoft-addin/tests/excel.test.mjs))**:
   - `UT-31`: Verifies Excel taskpane 2-row header layout conformance (`.primary-row` and `.chart-row`), existence of `#excelSchemaPreview`, and clean removal of redundant toolbar elements.
   - `UT-32`: Validates all Spreadsheet Tools & Analysis chips exist (`chipInsights`, `chipExecSummary`, `chipChartGen`, `chipTaxonomy`, `chipFormulaGen`, `chipXlookup`, `chipFormatRange`, etc.).
   - `UT-33`: Validates `TaskpaneShell.populateDocChips` context ingestion with active selection and document text.
   - `UT-34`: Validates `excelModule.onSessionReset` hides target chart badges and resets `chatHistory.scrollTop = 0`.
   - `UT-35`: Validates `ChatLog.clear()` clears messages, inserts welcome bubble, and scrolls to top (`scrollTop = 0`).

