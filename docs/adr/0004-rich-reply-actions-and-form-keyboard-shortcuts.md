# ADR-0002: Removal of "Insert Into Sheet" for Rich HTML Replies and Form Keyboard Interaction

- **Status:** Accepted
- **Date:** 2026-09-14
- **Author:** C. Schiele (shazbot996, ldap:schiele) — Excel & Shell Architecture Track
- **Affects:** `microsoft-addin/src/taskpanes/excel/responseActions.js`, `microsoft-addin/src/shell/TaskpaneShell.js`, `microsoft-addin/src/taskpane/taskpane.js`
- **Related Specs:** [`docs/specs/rich-clipboard-image-export-SPEC.md`](../specs/rich-clipboard-image-export-SPEC.md), [`docs/specs/excel-rich-content-insertion-SPEC.md`](../specs/excel-rich-content-insertion-SPEC.md)

---

## 1. Context & Problem Statement

### 1.1 Rich HTML vs. Spreadsheet Cell Mismatch
When Gemini generates rich analytical summaries, insights, or explanations, the output is rich HTML comprising headers, callout boxes, markdown tables, and embedded visual images. 

In Microsoft Excel, cells operate as a discrete calculation matrix. They cannot render arbitrary HTML trees or inline images. Previously, the add-in presented a generic `➕ Insert into Sheet` button on all replies. Clicking this button dumped the raw HTML or unformatted multiline text into the active worksheet cell, resulting in an unreadable and broken user experience.

External document processors (Google Docs, Microsoft Word) render rich HTML clipboard payloads natively with full styling and images.

### 1.2 Form Input & Keyboard Interaction Ergonomics
The add-in input interface utilizes a `<textarea>` for multiline prompt entry. Without explicit keyboard event handling, hitting the `Enter` key inserted a newline rather than submitting the prompt, requiring users to switch from the keyboard to the mouse to click the "Send" button.

Users expect standard conversational AI keyboard interactions:
- `Enter` submits the form and executes the agent.
- `Shift + Enter` or `Ctrl + Enter` (and `Cmd + Enter` on macOS) inserts a newline into the editor.

---

## 2. Decision

We have agreed upon and implemented the following two core architectural decisions:

### Decision 1: Remove "Insert Into Sheet" for Rich Text / Formatted HTML Replies
1. **Deprecate Generic Sheet Insertion:** Remove the `➕ Insert into Sheet` button for generic rich text / formatted HTML replies in the Excel taskpane.
2. **Retain Rich Copy Action:** Retain the `📋 Copy to Clipboard` button as the primary action for rich text / HTML replies. When copied, the multimodal clipboard utility (`clipboardUtils.js`) exports a rich `text/html` document envelope and sanitized `text/plain`, enabling one-click pasting into Microsoft Word, Google Docs, Apple Pages, and email clients.
3. **Preserve Native Object Insertion:** Continue providing dedicated, intent-driven action buttons when Gemini produces native Excel structures:
   - `➕ Insert Formula` for evaluated Excel formulas (`=SUMIFS(...)`, `=XLOOKUP(...)`).
   - `📊 Insert Native Chart` for interactive Excel chart specifications.
   - `🔄 Update Chart` for in-place chart modifications.
   - `🎨 Apply Range Format` for grid styling and highlight specifications.

### Decision 2: Form Keyboard Shortcuts for Prompt Submission and Multiline Editing
1. **`Enter` Keydown:** When `Enter` is pressed without modifiers (`shiftKey`, `ctrlKey`, `metaKey` are false), intercept the event, call `e.preventDefault()`, and trigger prompt submission (`runBtn.click()`).
2. **`Shift + Enter` Keydown:** Allow the default textarea behavior to insert a newline `\n` without submitting.
3. **`Ctrl + Enter` / `Cmd + Enter` Keydown:** Intercept the event, insert a newline `\n` at the cursor position (`selectionStart` / `selectionEnd`), and prevent form submission.

---

## 3. Consequences

### Positive
- **Eliminates Cell Corruption:** Users will no longer encounter raw HTML markup or unformatted text dumps in spreadsheet cells.
- **Clear User Expectations:** Cleanly separates document/narrative output (destined for clipboard / Docs / Word) from native grid operations (formulas, charts, formats).
- **Fast Keyboard Navigation:** Users can type prompts and press `Enter` to execute immediately, while retaining the ability to write structured, multiline prompts with `Shift + Enter` or `Ctrl + Enter`.

### Negative / Mitigations
- Users seeking to insert tabular data into Excel can use formula/table generation prompts or copy-paste directly, which produces a properly formatted grid.
