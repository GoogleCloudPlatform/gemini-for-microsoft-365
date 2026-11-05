---
type: spec
title: "Feature Specification: Excel Conversational Chart Modification via Office.js"
scope: "microsoft-addin/src/adapters/ExcelAdapter.js, microsoft-addin/src/taskpanes/excel/*, microsoft-addin/src/core/markdownParser.js, geminiproxy/index.js"
date: "2026-09-10"
---

# Feature Specification: Excel Conversational Chart Modification via Office.js

## 1. Problem Statement

### 1.1 Background and Current Behavior
When users generate native Excel charts using the Gemini add-in, they frequently require iterative refinements based on executive review, presentation styling, or exploratory analysis. Typical modification requests include:
* *"Change the chart to a 3D Pie chart"*
* *"Update title to '2026 Projected Regional Sales'"*
* *"Switch the chart to a stacked column and move legend to the bottom"*
* *"Filter series to show only Cloud and Services"*

Previously, when a user prompted to alter an existing chart on the worksheet:
1. The AI generated a complete new chart or unstructured text response.
2. The UI response actions misclassified chart edit responses as taxonomy categorization, displaying an erroneous `"🏷️ Insert Categories Column"` button that caused JavaScript runtime exceptions.
3. Users had to manually delete previous charts or use native Excel ribbon menus to recreate or alter chart properties.

### 1.2 User Impact & Limitations
* **Broken Conversational Flow:** Users could not iteratively tweak visualizations conversationally without leaving the taskpane.
* **Misrouted Action Buttons:** False positive matches in response action button dispatchers triggered invalid operations.
* **Loss of Chart State:** Creating a whole new chart often led to overlapping canvas artifacts and orphaned chart shapes.

### 1.3 Target Persona & Impact
* **Financial Analysts & Executives:** Need to rapidly iterate on chart presentations, adjusting chart types, styling palettes, and titles without rebuilding ranges from scratch.

---

## 2. Goals & Non-Goals

### 2.1 Goals
* **Conversational Chart Mutation:** Enable incremental modification of existing charts on the active worksheet via natural language.
* **Declarative Edit Specification (` ```excel-chart-edit `):** Support structured JSON edit manifests specifying mutation targets (title, chartType, legend, series, palette, data range, dimensions).
* **Smart Intent Conversion:** Automatically convert standard ` ```excel-chart ` responses into edit specifications when user prompt intent indicates an update or modification.
* **Scoped Action Button:** Display a dedicated **"🔄 Update Chart"** button on the response bubble when a chart modification specification is detected.
* **Active Chart Discovery:** Automatically inspect the active worksheet to locate existing charts (`sheet.charts.getItemAt(0)` or target by name/index).
* **Intelligent Sub-Range & Backslash Sanitization:** Handle escaped backslashes, range trimming, and categorical sub-range filtering.

### 2.2 Non-Goals
* **Multi-Sheet Cross-Workbook Chart Migration:** Moving charts across different `.xlsx` files is out of scope.
* **3D Surface & Contour Charts:** Unsupported by standard Office.js Web API.

---

## 3. Proposed Solution

### 3.1 High-Level Architecture

When the user enters a chart update request:
1. The taskpane sends the prompt with conversational history and active chart context to Gemini Enterprise / Vertex AI.
2. The model returns an ````excel-chart-edit```` JSON block (or an ````excel-chart```` block with update intent).
3. `extractExcelChartEditSpec()` extracts and normalizes the edit specification.
4. `responseActions.js` renders the **"🔄 Update Chart"** action button.
5. Clicking the button invokes `ExcelAdapter.applyChartModification()`, which loads the target chart via `Excel.run()`, mutates properties, and synchronizes the worksheet.

```mermaid
sequenceDiagram
    autonumber
    actor User
    participant Taskpane as Excel Taskpane UI
    participant Backend as Gemini Backend Proxy
    participant Adapter as ExcelAdapter.js
    participant Office as Excel Office.js Runtime

    User->>Taskpane: "Change the chart to a Donut chart and set title to 2026 Sales"
    Taskpane->>Backend: POST /askGeminiEnterprise (prompt + history)
    Backend-->>Taskpane: Response with ```excel-chart-edit { "chartType": "Doughnut", "title": "2026 Sales" }```
    Taskpane->>Taskpane: extractExcelChartEditSpec() -> Mounts "🔄 Update Chart" Button
    User->>Taskpane: Clicks "🔄 Update Chart"
    Taskpane->>Adapter: applyChartModification(editSpec)
    Adapter->>Office: Excel.run() -> sheet.charts.getItemAt(0)
    Office->>Office: chart.chartType = "Doughnut"; chart.title.text = "2026 Sales"
    Office-->>Adapter: Chart synchronized
    Adapter-->>Taskpane: Success notification ("✅ Chart updated")
```

---

## 4. Technical Design

### 4.1 Data Structures & Schema

```typescript
interface ExcelChartEditSpec {
  /** Target chart identifier: 'active' | index (number) | name (string) */
  targetChart?: string | number;

  /** New chart type if changing geometry */
  chartType?: string;
  type?: string;

  /** New chart title */
  title?: string;

  /** Updated data range (e.g., "A1:C10") */
  dataRange?: string;
  categoriesRange?: string;
  seriesRange?: string;

  /** Updated series orientation */
  seriesBy?: "auto" | "columns" | "rows";

  /** Updated legend settings */
  legend?: {
    visible?: boolean;
    position?: "Top" | "Bottom" | "Left" | "Right" | "Corner";
  };

  /** Updated dimensions */
  dimensions?: {
    width?: number;
    height?: number;
  };

  /** Color palette / style index */
  style?: number | string;
}
```

### 4.2 Key Functions & Algorithms

#### 4.2.1 Extraction Engine (`extractExcelChartEditSpec`)
Located in [`microsoft-addin/src/adapters/ExcelAdapter.js`](file:///Users/ravindranv/Code-Local/For-Excel-GE-Add-In/excel_pod_release/gemini-for-office-365/microsoft-addin/src/adapters/ExcelAdapter.js):
```javascript
export function extractExcelChartEditSpec(text, htmlContent = "", prompt = "") {
  const combined = (text || "") + "\n" + (htmlContent ? htmlContent.replace(/<[^>]+>/g, " ") : "");
  if (!combined.trim()) return null;

  // 1. Direct match on ```excel-chart-edit or ```chart-edit code blocks
  const editBlockRegex = /```(?:excel-chart-edit|chart-edit|excel_chart_edit)\s*([\s\S]*?)```/i;
  const match = combined.match(editBlockRegex);
  if (match && match[1]) {
    try {
      const parsed = JSON.parse(match[1].trim());
      if (parsed) return normalizeChartEditSpec(parsed);
    } catch (e) {
      console.warn("Failed to parse chart-edit JSON:", e);
    }
  }

  // 2. Intent conversion: If user asked to modify/update/change and response has ```excel-chart
  const isModificationPrompt = /\b(change|update|switch|convert|modify|turn into|make it|adjust|replace)\b/i.test(prompt || "");
  if (isModificationPrompt) {
    const regularSpec = extractExcelChartSpec(text, htmlContent);
    if (regularSpec) {
      return {
        targetChart: "active",
        chartType: regularSpec.type,
        title: regularSpec.title,
        dataRange: regularSpec.dataRange,
        categoriesRange: regularSpec.categoriesRange,
        seriesRange: regularSpec.seriesRange,
        legend: regularSpec.legend,
        dimensions: regularSpec.dimensions,
        seriesBy: regularSpec.seriesBy
      };
    }
  }

  return null;
}
```

#### 4.2.2 Chart Mutation Execution (`applyChartModification`)
```javascript
export async function applyChartModification(editSpec) {
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const charts = sheet.charts;
    charts.load("count, items/name, items/id");
    await context.sync();

    if (charts.count === 0) {
      throw new Error("No chart found on the active worksheet to update.");
    }

    // Target active/first chart
    const targetChart = charts.items[0];

    if (editSpec.chartType) {
      targetChart.chartType = normalizeChartType(editSpec.chartType);
    }
    if (editSpec.title) {
      targetChart.title.text = editSpec.title;
      targetChart.title.visible = true;
    }
    if (editSpec.legend) {
      if (editSpec.legend.visible !== undefined) targetChart.legend.visible = editSpec.legend.visible;
      if (editSpec.legend.position) targetChart.legend.position = editSpec.legend.position;
    }
    if (editSpec.dataRange) {
      const range = sheet.getRange(editSpec.dataRange);
      targetChart.setData(range, editSpec.seriesBy === "rows" ? Excel.ChartSeriesBy.rows : Excel.ChartSeriesBy.columns);
    }
    if (editSpec.dimensions) {
      if (editSpec.dimensions.width) targetChart.width = editSpec.dimensions.width;
      if (editSpec.dimensions.height) targetChart.height = editSpec.dimensions.height;
    }

    await context.sync();
    return { success: true, chartName: targetChart.name };
  });
}
```

---

## 5. UI / UX Integration

1. **2-Row Header Preview Bar & Sticky Chart Target Badge**:
   - When a chart is selected or generated, Row 2 of the top preview header (`#targetChartRow`) displays a dedicated chart badge: `📊 Target Chart: [Chart Name] ✕`.
   - Clicking the dismiss icon (`#removeTargetChartBtn`) clears the chart target and collapses Row 2.
   - On session reset or prompt clearing, `onSessionReset()` dismisses `#targetChartRow` and resets the scroll position (`UT-34`).
2. **Action Button Rendering**:
   - `renderResponseActions()` inspects response content. If `extractExcelChartEditSpec()` returns valid data, it renders the primary button:
     `<button class="action-btn chart-edit" id="btnUpdateChart">🔄 Update Chart</button>`.
   - Prevents conflicting buttons (e.g. taxonomy categorizer) from rendering simultaneously (`UT-20`, `UT-22`).

---

## 6. Testing & Validation Strategy

1. **Unit Test Coverage ([`microsoft-addin/tests/excel.test.mjs`](file:///Users/ravindranv/Code-Local/For-Excel-GE-Add-In/excel_pod_release/gemini-for-office-365/microsoft-addin/tests/excel.test.mjs))**:
   - `UT-5`: Parses fenced ````excel-chart-edit```` blocks and normalizes properties.
   - `UT-20`: Verifies natural language chart modification prompts render **"🔄 Update Chart"** without taxonomy collisions.
   - `UT-21`: Validates intent-based conversion from regular chart blocks to edit specs.
   - `UT-22`: Guards chart JSON containing `"categories"` keys against false positive taxonomy categorization.
   - `UT-25`: Validates chart prompt intent detection and button labeling for natural language chart prompts.
   - `UT-31`: Verifies 2-row header structure with Row 2 chart target row and badge dismissal.
   - `UT-34`: Verifies `onSessionReset()` clears the locked chart target and hides the Row 2 target chart badge.

