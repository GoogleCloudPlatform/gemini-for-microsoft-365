---
type: spec
title: "Feature Specification: Excel Zero-Shot Taxonomy Categorization via Office.js"
scope: "microsoft-addin/src/adapters/ExcelAdapter.js, microsoft-addin/src/taskpanes/excel/*, geminiproxy/index.js"
date: "2026-09-10"
---

# Feature Specification: Excel Zero-Shot Taxonomy Categorization via Office.js

## 1. Problem Statement

### 1.1 Background and Current Behavior
Spreadsheet users frequently encounter unstandardized, high-cardinality unstructured text columns (e.g., bank transaction descriptions, vendor names, customer support logs, product descriptions, or raw CRM notes) that require structured classification for reporting, aggregation, or accounting.

Historically, users had to:
1. Manually write convoluted nested `=IF(...)` or `=SWITCH(...)` formulas with regex or fuzzy string matching.
2. Manually tag hundreds of rows one by one.
3. Copy-paste external classifications back and forth between web AI tools and Excel, leading to row alignment errors and clipboard truncation.

### 1.2 User Impact & Limitations
* **Laborious & Error-Prone:** Classifying hundreds of transactions or line items manually consumes hours and is prone to human inconsistency.
* **Complex Multi-Class Logics:** Traditional Excel formulas struggle with semantic ambiguity (e.g., classifying "Uber Eats" under *Dining / Meals* vs "Uber Trip" under *Travel / Transportation*).
* **Grid Misalignment Risk:** Copy-pasting AI output from an external browser window often misaligns rows when empty or multiline cells exist.

### 1.3 Target Persona & Impact
* **Accounting & Finance Teams:** Categorizing GL transaction lines, expense reports, and accounts payable.
* **Data Analysts & Operations:** Tagging customer feedback, operational tickets, SKUs, and inventory logs.

---

## 2. Goals & Non-Goals

### 2.1 Goals
* **Zero-Shot Semantic Classification:** Use Gemini Enterprise to analyze textual descriptions in an Excel column and assign standard category labels.
* **Declarative Taxonomy Schema (` ```excel-taxonomy `):** Standardize the JSON response format including source column, target column, inserted header title, row ranges, value arrays, and distribution summaries.
* **Native Column Insertion:** Automatically shift existing columns right (`sheet.getRange("C:C").insert(Excel.InsertShiftDirection.right)`) to avoid overwriting existing data.
* **Corporate Header Styling:** Format the new category column header with professional Office Blue styling (`#004E8C` text, `#EBF3FC` fill, bold Segoe UI, centered).
* **One-Click Batch Execution:** Insert all categorized rows into the worksheet in a single atomic matrix write via `targetDataRange.values = matrix` followed by `format.autofitColumns()`.
* **Action Button & UI Card:** Mount a dedicated **"🏷️ Insert Categories Column"** response button and render a summary distribution chip in the taskpane.

### 2.2 Non-Goals
* **Machine Learning Model Training/Fine-Tuning:** The feature performs zero-shot and few-shot inference through the LLM prompt rather than training custom classifiers on user infrastructure.
* **Persistent External Taxonomy Store:** Workbooks are standalone; storing taxonomy definitions in external corporate MDM systems is out of scope.

---

## 3. Proposed Solution

### 3.1 High-Level Architecture

When the user requests taxonomy categorization (e.g. by clicking the **"🏷️ Categorize Taxonomy"** prompt chip or typing a custom request):
1. The add-in extracts the selected column values and active worksheet schema.
2. The prompt requests a standardized ````excel-taxonomy```` JSON block with mapped categories.
3. `extractExcelTaxonomySpec()` parses the response and `responseActions.js` displays the **"🏷️ Insert Categories Column"** button.
4. When clicked, `applyTaxonomyCategorization()` executes within `Excel.run()`, shifts columns right, styles the header, writes the values matrix, and autofits the column.

```mermaid
sequenceDiagram
    autonumber
    actor User
    participant Taskpane as Excel Taskpane UI
    participant Backend as Gemini Backend Proxy
    participant Adapter as ExcelAdapter.js
    participant Office as Excel Office.js Runtime

    User->>Taskpane: Selects column B & clicks "🏷️ Categorize Taxonomy"
    Taskpane->>Adapter: getSelectedText() / getWorksheetSchema()
    Adapter-->>Taskpane: Column B data (Rows 2-50)
    Taskpane->>Backend: POST /askGeminiEnterprise (prompt + column data)
    Backend-->>Taskpane: Response with ```excel-taxonomy { sourceColumn: "B", targetColumn: "C", categories: [...] }```
    Taskpane->>Taskpane: extractExcelTaxonomySpec() -> Mounts "🏷️ Insert Categories Column"
    User->>Taskpane: Clicks "Insert Categories Column"
    Taskpane->>Adapter: applyTaxonomyCategorization(spec)
    Adapter->>Office: Excel.run()
    Office->>Office: sheet.getRange("C:C").insert(right)
    Office->>Office: header.values = [["Category"]]; header.format.fill.color = "#EBF3FC"
    Office->>Office: dataRange.values = [["Meals"], ["Software"], ...]
    Office->>Office: dataRange.format.autofitColumns()
    Office-->>Adapter: Complete & synchronized
    Adapter-->>Taskpane: Success notification ("✅ Inserted 49 categories into Column C")
```

---

## 4. Technical Design

### 4.1 Data Structures & Specification Schema

```typescript
interface ExcelTaxonomySpec {
  action: "categorize";
  sourceColumn: string;       // e.g. "B"
  targetColumn: string;       // e.g. "C"
  insertColumn?: boolean;     // default: true (shifts adjacent columns right)
  header: string;             // e.g. "Category" or "Expense Type"
  startRow: number;           // e.g. 2 (1-based index)
  endRow?: number;            // e.g. 50
  categories: Array<{
    sourceValue?: string;
    category: string;
  }> | string[];
  categoryValues: string[];   // Normalized flat array of categories for each row
  count: number;              // Total number of classified rows
  taxonomySummary?: Record<string, number>; // Distribution map (e.g. { "Software": 12, "Meals": 8 })
}
```

### 4.2 Key Functions & Algorithms

#### 4.2.1 Extraction & Normalization Engine
Located in [`microsoft-addin/src/adapters/ExcelAdapter.js`](file:///Users/ravindranv/Code-Local/For-Excel-GE-Add-In/excel_pod_release/gemini-for-office-365/microsoft-addin/src/adapters/ExcelAdapter.js):
```javascript
export function normalizeTaxonomySpec(spec) {
  if (!spec || typeof spec !== "object") return null;

  let sourceCol = String(spec.sourceColumn || spec.sourceCol || "B").toUpperCase().trim();
  let targetCol = String(spec.targetColumn || spec.targetCol || "").toUpperCase().trim();
  if (!targetCol) {
    const sIdx = letterToColumnIndex(sourceCol);
    targetCol = columnIndexToLetter(sIdx + 1);
  }

  const header = spec.header || spec.columnName || "Category";
  const startRow = typeof spec.startRow === "number" && spec.startRow >= 1 ? spec.startRow : 2;

  let categories = spec.categories || spec.mappings || spec.values || [];
  const categoryValues = categories.map(item => {
    if (typeof item === "string") return item;
    if (item && typeof item === "object") return item.category || item.categoryName || "";
    return String(item || "");
  });

  return {
    action: "categorize",
    sourceColumn: sourceCol,
    targetColumn: targetCol,
    insertColumn: spec.insertColumn !== false,
    header,
    startRow,
    endRow: startRow + categoryValues.length - 1,
    categories,
    categoryValues,
    count: categoryValues.length,
    taxonomySummary: spec.taxonomySummary || null
  };
}
```

#### 4.2.2 Batch Column Insertion (`applyTaxonomyCategorization`)
```javascript
export async function applyTaxonomyCategorization(context, sheet, spec) {
  const normalized = normalizeTaxonomySpec(spec);
  const targetCol = normalized.targetColumn;
  const header = normalized.header;
  const startRow = normalized.startRow;
  const headerRow = startRow > 1 ? startRow - 1 : 1;
  const categoryValues = normalized.categoryValues;
  const endRow = startRow + categoryValues.length - 1;

  // 1. Shift existing columns right
  if (normalized.insertColumn) {
    const colRange = sheet.getRange(`${targetCol}:${targetCol}`);
    colRange.insert(Excel.InsertShiftDirection.right);
    await context.sync();
  }

  // 2. Set Header and Executive Office Styling
  const headerCell = sheet.getRange(`${targetCol}${headerRow}`);
  headerCell.values = [[header]];
  headerCell.format.font.bold = true;
  headerCell.format.font.name = "Segoe UI";
  headerCell.format.font.color = "#004e8c";
  headerCell.format.fill.color = "#EBF3FC";
  headerCell.format.horizontalAlignment = "Center";

  // 3. Batch write all categorized rows
  const targetRange = sheet.getRange(`${targetCol}${startRow}:${targetCol}${endRow}`);
  targetRange.values = categoryValues.map(cat => [cat]);
  targetRange.format.font.name = "Segoe UI";
  targetRange.format.verticalAlignment = "Center";

  // 4. Autofit column width
  const fullRange = sheet.getRange(`${targetCol}${headerRow}:${targetCol}${endRow}`);
  fullRange.format.autofitColumns();

  await context.sync();
  return { success: true, address: `${targetCol}${headerRow}:${targetCol}${endRow}`, count: categoryValues.length };
}
```

---

## 5. UI / UX Integration

1. **Quick Action Chips ([`chips.js`](file:///Users/ravindranv/Code-Local/For-Excel-GE-Add-In/excel_pod_release/gemini-for-office-365/microsoft-addin/src/taskpanes/excel/chips.js))**:
   - `selectionChips.chipCategorizeTaxonomy` mounts **"🏷️ Categorize Taxonomy"** when text columns are selected.
2. **Response Action Button ([`responseActions.js`](file:///Users/ravindranv/Code-Local/For-Excel-GE-Add-In/excel_pod_release/gemini-for-office-365/microsoft-addin/src/taskpanes/excel/responseActions.js))**:
   - Mounts `<button class="action-btn insert-taxonomy-btn">🏷️ Insert Categories Column</button>`.
   - Includes strict intent and anti-collision guards to avoid triggering on chart edits or summary insights.

---

## 6. Testing & Validation Strategy

1. **Unit Test Coverage ([`microsoft-addin/tests/excel.test.mjs`](file:///Users/ravindranv/Code-Local/For-Excel-GE-Add-In/excel_pod_release/gemini-for-office-365/microsoft-addin/tests/excel.test.mjs))**:
   - `UT-7`: `normalizeTaxonomySpec` defaults target column to adjacent column (`Col B -> Col C`), validates row spans, and computes category frequency distributions.
   - `UT-8`: `extractExcelTaxonomySpec` extracts and validates fenced ````excel-taxonomy```` JSON blocks.
   - `UT-9`: `extractExcelTaxonomySpec` accurately parses fallback markdown classification tables.
   - `UT-10`: `applyTaxonomyCategorization` inserts a new column (`C:C`), sets professional header styling (`#EBF3FC`), writes data in batch, and triggers autofit.
   - `UT-11`: `parseMarkdown` renders interactive visual preview cards for taxonomy specifications.
   - `UT-13`: `selectionChips` defines `selTaxonomy` (*"🏷️ Categorize Taxonomy"*) as the primary selection action.
   - `UT-15`: `renderResponseActions` renders **"🏷️ Insert Categories Column"** for taxonomy specifications.
   - `UT-18`: Scoped action button dispatcher isolates taxonomy categorizer from insights summary tables.
   - `UT-20` & `UT-22`: Anti-collision validation ensuring chart updates and chart JSON containing `"categories"` keys do NOT false-trigger taxonomy categorization.

