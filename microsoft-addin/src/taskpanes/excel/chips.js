/**
 * Copyright 2026 Google LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

/**
 * Gemini for Microsoft Excel - Quick Action Chips & Spreadsheet Tools
 */

export const docChips = [
  {
    id: "chipInsights",
    label: "📈 Table Insights",
    title: "Perform deep statistical insights analysis on the active worksheet or selected range",
    isInsights: true,
    scope: "auto",
    prompt: "Perform deep Table Insights on the spreadsheet data with summary statistics, anomalies, key trends, and a visual graph directive."
  },
  {
    id: "chipChartGen",
    label: "📊 Native Chart Builder",
    title: "Configure and generate an interactive native Excel chart specification",
    isChartBuilder: true,
    prompt: "Analyze the active worksheet and selected data to generate an optimal native Excel chart specification (ColumnClustered, Line, BarClustered, or Pie) visualizing key metrics across categories. Return the specification as an ```excel-chart code block."
  },
  {
    id: "chipFormulaGen",
    label: "✨ Formula Generator",
    title: "Analyze selected data (or sheet) and get top recommended formulas with plain-English explanations",
    isFormulaGenerator: true,
    prompt: "Analyze the selected spreadsheet range (or active worksheet) and recommend the top 4-5 most useful Excel formulas with clear explanations."
  },
  {
    id: "chipTaxonomy",
    label: "🏷️ Taxonomy Categorizer",
    title: "Map text column entries to standardized taxonomy categories",
    isTaxonomy: true,
    scope: "auto",
    prompt: "Categorize the data in this active worksheet or selected range. Identify the primary text column (e.g., Opportunity Name, Merchant, Product, Vendor, Description) or selected column and categorize each row into standardized categories. Return a declarative ```excel-taxonomy JSON specification with sourceColumn, targetColumn, insertColumn: true, header, startRow, and categories array with { row, sourceValue, category }."
  }
];

// Preserved for backward compatibility
export const selectionChips = [
  {
    id: "selInsightsSel",
    label: "📊 Insights (Selection)",
    title: "Compute statistics and extract deep insights from highlighted cells",
    isInsights: true,
    scope: "selection",
    instruction: "Perform deep Table Insights on the highlighted cells with summary statistics, anomalies, key trends, and a visual graph directive."
  },
  {
    id: "selExecSel",
    label: "💡 Exec Summary (Selection)",
    title: "Generate executive KPI summary for highlighted cells",
    isExecSummary: true,
    scope: "selection",
    instruction: "Generate an Executive Summary for corporate leadership based on the highlighted cells with pre-calculated metrics, key findings, and strategic recommendations."
  },
  {
    id: "selTaxonomy",
    label: "🏷️ Categorize Taxonomy",
    isTaxonomy: true,
    scope: "selection",
    instruction: "Categorize the entries in this highlighted column row-by-row into standardized taxonomy categories (e.g., Expense types, Opportunity Types, Departments). Return a declarative ```excel-taxonomy JSON specification with sourceColumn, targetColumn, insertColumn: true, header, startRow, and categories array with { row, sourceValue, category }."
  },
  {
    id: "selChart",
    label: "📈 Visualize Selection",
    instruction: "Create a native Excel chart visualizing the selected data range. Determine the best chart type based on the series dimensions and return an ```excel-chart JSON specification."
  },
  {
    id: "selExplainFormula",
    label: "🧮 Explain Formula/Data",
    instruction: "Explain what this selected Excel formula or data range does, step-by-step, in plain English."
  },
  {
    id: "selFormatRange",
    label: "🎨 Style & Format",
    instruction: "Format and style this selected range with professional Excel formatting (header fill color, clear font styling, borders, and number formatting). Return an ```excel-range-format JSON specification."
  },
  {
    id: "selShorten",
    label: "📉 Summarize Range",
    instruction: "Summarize the key takeaways and stats from the selected range concisely."
  }
];
