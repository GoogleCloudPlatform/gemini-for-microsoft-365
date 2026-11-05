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
 * Host Adapter Implementation for Microsoft Excel (Excel.run)
 * 
 * Provides native Excel spreadsheet intelligence, automated formula generation
 * (e.g., XLOOKUP, SUMIFS, COUNTIFS, AVERAGEIFS, INDEX/MATCH), worksheet schema &
 * column header context extraction, cell matrix injection, and in-cell @gemini triggers.
 * Integrates with ExcelInsightsService for table insights, executive summaries,
 * statistical pre-parsing, and worksheet manipulation.
 * 
 * @author Sathya AG, Principal Architect, Google
 */

import { ExcelInsightsService } from './ExcelInsightsService.js';

/**
 * Converts a 0-indexed column number to Excel column letters (e.g., 0 -> 'A', 2 -> 'C', 27 -> 'AB')
 * @param {number} colIndex 
 * @returns {string}
 */
export function columnIndexToLetter(colIndex) {
  let temp;
  let letter = "";
  let col = colIndex + 1;
  while (col > 0) {
    temp = (col - 1) % 26;
    letter = String.fromCharCode(65 + temp) + letter;
    col = Math.floor((col - temp) / 26);
  }
  return letter;
}

/**
 * Converts Excel column letter(s) to 0-indexed column number (e.g., 'A' -> 0, 'C' -> 2, 'AB' -> 27)
 * @param {string} letter 
 * @returns {number}
 */
export function letterToColumnIndex(letter) {
  let col = 0;
  const clean = String(letter || "").toUpperCase().trim();
  for (let i = 0; i < clean.length; i++) {
    col = col * 26 + (clean.charCodeAt(i) - 64);
  }
  return col - 1;
}

/**
 * Parses an Excel range string (e.g. "B1:B50", "A:D", "Sheet1!$C$2:$E$10", "B") into column/row coordinates
 * @param {string} rangeStr 
 * @returns {Object|null}
 */
export function parseRangeAddress(rangeStr) {
  if (!rangeStr) return null;
  let clean = String(rangeStr).replace(/^=|^["']|["']$/g, "").replace(/[$]/g, "").trim();
  if (clean.includes("!")) {
    clean = clean.split("!").pop().trim();
  }
  if (!clean) return null;

  const parseCellOrCol = (s) => {
    const raw = String(s).trim();
    // 1. Cell with column and row: "B2", "AA10"
    const mCell = raw.match(/^([A-Za-z]+)([0-9]+)$/);
    if (mCell) {
      const colLetter = mCell[1].toUpperCase();
      return {
        colLetter,
        colIndex: letterToColumnIndex(colLetter),
        row: parseInt(mCell[2], 10)
      };
    }
    // 2. Column letters only: "B", "AA"
    const mCol = raw.match(/^([A-Za-z]+)$/);
    if (mCol) {
      const colLetter = mCol[1].toUpperCase();
      return {
        colLetter,
        colIndex: letterToColumnIndex(colLetter),
        row: null
      };
    }
    return null;
  };

  if (!clean.includes(":")) {
    const c = parseCellOrCol(clean);
    if (!c) return null;
    return {
      startCol: c.colLetter,
      startColIndex: c.colIndex,
      startRow: c.row,
      endCol: c.colLetter,
      endColIndex: c.colIndex,
      endRow: c.row,
      address: c.row !== null ? `${c.colLetter}${c.row}:${c.colLetter}${c.row}` : `${c.colLetter}:${c.colLetter}`
    };
  }

  const parts = clean.split(":");
  const c1 = parseCellOrCol(parts[0]);
  const c2 = parseCellOrCol(parts[1]);
  if (!c1 || !c2) return null;

  const minColIdx = Math.min(c1.colIndex, c2.colIndex);
  const maxColIdx = Math.max(c1.colIndex, c2.colIndex);
  const startCol = columnIndexToLetter(minColIdx);
  const endCol = columnIndexToLetter(maxColIdx);

  let startRow = null;
  let endRow = null;
  if (c1.row !== null && c2.row !== null) {
    startRow = Math.min(c1.row, c2.row);
    endRow = Math.max(c1.row, c2.row);
  } else if (c1.row !== null) {
    startRow = c1.row;
  } else if (c2.row !== null) {
    endRow = c2.row;
  }

  const addrStr = (startRow !== null && endRow !== null)
    ? `${startCol}${startRow}:${endCol}${endRow}`
    : (startRow !== null ? `${startCol}${startRow}:${endCol}` : `${startCol}:${endCol}`);

  return {
    startCol,
    startColIndex: minColIdx,
    startRow,
    endCol,
    endColIndex: maxColIdx,
    endRow,
    address: addrStr
  };
}

/**
 * Resolves a range address with fallback row bounds from worksheet used range
 * @param {string|Object} rangeInput 
 * @param {Object} defaultBounds - { startRow: number, endRow: number }
 * @returns {Object|null}
 */
export function resolveRangeAddress(rangeInput, defaultBounds = { startRow: 1, endRow: 50 }) {
  const parsed = typeof rangeInput === "string" ? parseRangeAddress(rangeInput) : rangeInput;
  if (!parsed) return null;
  const startRow = (parsed.startRow !== null && parsed.startRow !== undefined) ? parsed.startRow : (defaultBounds.startRow || 1);
  const endRow = (parsed.endRow !== null && parsed.endRow !== undefined) ? parsed.endRow : (defaultBounds.endRow || defaultBounds.startRow || 50);
  return {
    ...parsed,
    startRow,
    endRow,
    address: `${parsed.startCol}${startRow}:${parsed.endCol}${endRow}`
  };
}

/**
 * Normalizes an Excel formula string:
 * - Ensures leading '='
 * - Strips non-breaking spaces and unicode spaces
 * - Replaces smart/curly quotes with standard ASCII quotes
 * - Normalizes unicode minus and operators
 * - Replaces single-quoted string criteria with double quotes (e.g., 'Approved' -> "Approved")
 *   while preserving single quotes for sheet names (e.g., 'Sheet 1'!A1)
 * - Collapses multi-line breaks and excess whitespace
 * - Strips trailing punctuation (periods, semicolons)
 * @param {string} formula 
 * @returns {string}
 */
export function normalizeFormula(formula) {
  if (!formula) return "";
  let clean = formula.trim();

  // Strip non-breaking spaces and unicode spaces
  clean = clean.replace(/[\u00A0\u2000-\u200B\uFEFF]/g, " ");

  // Normalize smart/curly quotes
  clean = clean.replace(/[\u201C\u201D\u201E\u00AB\u00BB]/g, "\"");
  clean = clean.replace(/[\u2018\u2019\u201A]/g, "'");

  // Normalize unicode math symbols
  clean = clean.replace(/[\u2212\u2013\u2014]/g, "-");
  clean = clean.replace(/[\u00D7]/g, "*");
  clean = clean.replace(/[\u00F7]/g, "/");

  // Strip wrapping backticks or quotes if any
  clean = clean.replace(/^`+|`+$/g, "").trim();

  // Strip trailing period or semicolon
  clean = clean.replace(/[.;]+$/, "").trim();

  // Normalize multi-line breaks and excess whitespace to single spaces
  clean = clean.replace(/\r?\n+/g, " ").replace(/\s+/g, " ").trim();

  // Ensure no space immediately after leading "="
  clean = clean.replace(/^=\s+/, "=");

  // Ensure leading '='
  if (!clean.startsWith("=")) {
    clean = "=" + clean;
  }

  // Normalize single-quoted string literals to double quotes for criteria
  // e.g. =SUMIFS(C:C, D:D, 'Approved') -> =SUMIFS(C:C, D:D, "Approved")
  // Excludes sheet references like 'Sheet 1'!A1 by checking for non-'!' content not followed by '!'
  clean = clean.replace(/'([^'!\r\n]+)'(?!\s*!)/g, '"$1"');

  return clean;
}

/**
 * Validates whether a candidate string is a plausible Excel formula
 * @param {string} f
 * @returns {boolean}
 */
export function isFormulaValid(f) {
  if (!f || typeof f !== "string") return false;
  let str = f.trim();
  str = str.replace(/[\u00A0\u2000-\u200B\uFEFF]/g, " ");
  if (!str.startsWith("=")) return false;

  // Reject JS code, Office.js objects, method chains, statements
  if (/targetRange|sheet\.|context\.|workbook\.|Excel\.|getFormat|setColor|getFill|getFont|getItem|select|format\.|autofit|innerHTML|document\.|window\.|console\.|=>|function|\{|\}|;|\[\[|\]\]/i.test(str)) {
    return false;
  }
  // Reject if it looks like JS assignment or method call e.g. =x.y( or =x=y or =const or =let
  if (/^[A-Za-z0-9_]+\.[A-Za-z0-9_.]+/i.test(str.substring(1)) || /^=(?:const|let|var|return)\b/i.test(str)) {
    return false;
  }

  // Check parenthesis balance
  const openCount = (str.match(/\(/g) || []).length;
  const closeCount = (str.match(/\)/g) || []).length;
  if (openCount !== closeCount) return false;

  // Formula should start with standard Excel function or arithmetic expression or cell reference
  return /^=[_A-Za-z0-9_.]+\(/i.test(str) ||
         /^=[_A-Za-z0-9_!$'"\s]+[\+\-\*\/]/i.test(str) ||
         /^=[_A-Za-z0-9_!$'"]+$/i.test(str) ||
         /^=\s*(?:_xlfn\.)?(?:SUMIFS|COUNTIFS|AVERAGEIFS|XLOOKUP|VLOOKUP|HLOOKUP|INDEX|MATCH|SUM|AVERAGE|COUNT|COUNTA|MAX|MIN|IF|IFS|SWITCH|CONCATENATE|TEXTJOIN|FILTER|UNIQUE|SORT|LET|LAMBDA|OFFSET|INDIRECT|UPPER|LOWER|PROPER|TRIM|LEN|LEFT|RIGHT|MID|DATE|TODAY|NOW|YEAR|MONTH|DAY|AND|OR|NOT|ISBLANK|ISNUMBER|ISTEXT|ISERROR|IFERROR|HYPERLINK|ROUND|ROUNDUP|ROUNDDOWN|INT|ABS)\s*\(/i.test(str);
}

/**
 * Robustly extracts an Excel formula from AI response text or HTML
 * Handles code blocks, inline code, standalone lines, and embedded formula expressions.
 * @param {string} text - Raw AI response text
 * @param {string} htmlContent - Optional parsed HTML content
 * @returns {string|null} - Clean normalized formula or null
 */
export function extractExcelFormula(text, htmlContent = "") {
  let raw = (text || "");
  // Ignore non-Excel code blocks (js, python, json, etc.) before scanning general text
  const cleanedText = raw.replace(/```(?:javascript|js|typescript|ts|json|python|py|html|css|bash|sh)[\s\S]*?```/gi, "");
  const combined = cleanedText + "\n" + (htmlContent ? htmlContent.replace(/<[^>]+>/g, " ") : "");
  if (!combined.trim()) return null;

  // 1. Check for fenced code blocks containing a formula: ```excel\n=SUMIFS(...)\n```
  const codeBlockRegex = /```(?:excel|formula|vba|xlsx)?\s*[\r\n]+([\s\S]*?)[\r\n]+```/gi;
  let cbMatch;
  while ((cbMatch = codeBlockRegex.exec(combined)) !== null) {
    const rawBlock = cbMatch[1];
    // Filter out comment lines starting with // or # or rem
    const cleanLines = rawBlock.split(/\r?\n/)
      .map(l => l.trim())
      .filter(l => l && !l.startsWith("//") && !l.startsWith("#") && !/^rem\b/i.test(l));
    
    // First try full joined block (handles multiline formulas)
    const joined = cleanLines.join(" ");
    const candidateJoined = joined.startsWith("=") ? joined : "=" + joined;
    if (isFormulaValid(candidateJoined)) {
      return normalizeFormula(candidateJoined);
    }

    // Next try individual lines
    for (const line of cleanLines) {
      const candidate = line.startsWith("=") ? line : "=" + line;
      if (isFormulaValid(candidate)) {
        return normalizeFormula(candidate);
      }
    }
  }

  // 2. Check for inline code containing a formula: `=SUMIFS(...)` or `=XLOOKUP(...)`
  const inlineCodeMatch = combined.match(/`(=[A-Z_]+[A-Z0-9_.]*\([^`]*\))`|`(=[^`]+)`/i);
  if (inlineCodeMatch) {
    const candidate = (inlineCodeMatch[1] || inlineCodeMatch[2] || "").trim();
    if (candidate && isFormulaValid(candidate)) return normalizeFormula(candidate);
  }

  // 3. Check for standalone line starting with an Excel formula
  const lineMatch = combined.match(/^[ \t]*(=[A-Z_]+[A-Z0-9_.]*\(.*?\))[ \t]*$/m);
  if (lineMatch && isFormulaValid(lineMatch[1].trim())) {
    return normalizeFormula(lineMatch[1].trim());
  }

  // 4. Check for any standard Excel formula pattern in the text (not inside JS strings or quotes)
  const funcPattern = /(?:^|[\s`(\[])(=\s*(?:_xlfn\.)?(?:SUMIFS|COUNTIFS|AVERAGEIFS|XLOOKUP|VLOOKUP|HLOOKUP|INDEX|MATCH|SUM|AVERAGE|COUNT|COUNTA|MAX|MIN|IF|IFS|SWITCH|CONCATENATE|TEXTJOIN|FILTER|UNIQUE|SORT|LET|LAMBDA|OFFSET|INDIRECT|UPPER|LOWER|PROPER|TRIM|LEN|LEFT|RIGHT|MID|DATE|TODAY|NOW|YEAR|MONTH|DAY|AND|OR|NOT|ISBLANK|ISNUMBER|ISTEXT|ISERROR|IFERROR|HYPERLINK|ROUND|ROUNDUP|ROUNDDOWN|INT|ABS)\s*\([^)\r\n]*\))/i;
  const funcMatch = combined.match(funcPattern);
  if (funcMatch && isFormulaValid(funcMatch[1].trim())) {
    return normalizeFormula(funcMatch[1].trim());
  }

  // 5. If the entire text starts with '=' (trimmed)
  const trimmed = (text || "").trim();
  if (trimmed.startsWith("=") && trimmed.length > 2 && isFormulaValid(trimmed)) {
    return normalizeFormula(trimmed);
  }

  return null;
}

/**
 * Map of friendly/LLM chart types to official Excel.ChartType PascalCase strings
 */
export const CHART_TYPE_MAP = {
  columnclustered: "ColumnClustered",
  columnstacked: "ColumnStacked",
  columnstacked100: "ColumnStacked100",
  column: "ColumnClustered",
  barclustered: "BarClustered",
  barstacked: "BarStacked",
  barstacked100: "BarStacked100",
  bar: "BarClustered",
  line: "Line",
  linemarkers: "LineMarkers",
  linestacked: "LineStacked",
  pie: "Pie",
  pieexploded: "PieExploded",
  doughnut: "Doughnut",
  area: "Area",
  areastacked: "AreaStacked",
  scatter: "XYScatter",
  xyscatter: "XYScatter",
  waterfall: "Waterfall",
  treemap: "Treemap",
  sunburst: "Sunburst",
  histogram: "Histogram",
  funnel: "Funnel"
};

/**
 * Normalizes user or model generated chart type to exact Excel.ChartType PascalCase string
 * @param {string} rawType 
 * @returns {string}
 */
export function normalizeChartType(rawType) {
  if (!rawType) return "ColumnClustered";
  const clean = String(rawType).replace(/[^a-zA-Z0-9]/g, "").toLowerCase();
  return CHART_TYPE_MAP[clean] || "ColumnClustered";
}

/**
 * Normalizes a chart specification object into a clean structured schema
 * @param {Object} spec 
 * @returns {Object|null}
 */
export function normalizeChartSpec(spec) {
  if (!spec) return null;
  const rawType = spec.type || spec.chartType || "ColumnClustered";

  // Helper to extract clean range strings from strings, arrays, or object lists
  const extractRangeStr = (val) => {
    if (!val) return "";
    if (typeof val === "string") return val;
    if (Array.isArray(val)) {
      return val.map(item => {
        if (typeof item === "string") return item;
        if (item && typeof item === "object") return item.range || item.address || item.column || item.name || "";
        return "";
      }).filter(Boolean).join(", ");
    }
    if (typeof val === "object") {
      return val.range || val.address || val.column || "";
    }
    return String(val);
  };

  let rawRange = extractRangeStr(spec.dataRange || spec.range || spec.sourceRange || spec.data || spec.targetRange || spec.tableRange || spec.dataSource || "");
  let categoriesRange = extractRangeStr(spec.categoriesRange || spec.categoryRange || spec.labelRange || spec.labelsRange || spec.xRange || spec.categoryAxis || spec.categories || spec.labels || spec.xAxis || spec.x_axis || "");
  let seriesRange = extractRangeStr(spec.seriesRange || spec.valuesRange || spec.valueRange || spec.yRange || spec.metricsRange || spec.metricRange || spec.series || spec.values || spec.yAxis || spec.y_axis || spec.dataSeries || "");

  // If dataRange contains comma-separated ranges e.g. "B1:B50, E1:E50"
  if (rawRange && String(rawRange).includes(",") && !categoriesRange && !seriesRange) {
    const parts = String(rawRange).split(",").map(p => p.trim()).filter(Boolean);
    if (parts.length >= 2) {
      categoriesRange = parts[0];
      seriesRange = parts.slice(1).join(", ");
      rawRange = "";
    }
  }

  // Clean strings
  const cleanRange = (r) => r ? String(r).replace(/[$]/g, "").replace(/^=|^["']|["']$/g, "").trim() : "";
  categoriesRange = cleanRange(categoriesRange);
  seriesRange = cleanRange(seriesRange);
  rawRange = cleanRange(rawRange);

  let aggregate = null;
  const rawAgg = spec.aggregate || spec.aggregation || spec.groupBy || spec.group_by || spec.agg || null;
  if (rawAgg) {
    if (typeof rawAgg === "string") {
      aggregate = rawAgg.toLowerCase().trim();
    } else if (rawAgg === true) {
      aggregate = "sum";
    }
  }

  return {
    type: normalizeChartType(rawType),
    rawType: rawType,
    dataRange: rawRange,
    categoriesRange: categoriesRange,
    seriesRange: seriesRange,
    aggregate: aggregate,
    title: spec.title ? String(spec.title).trim() : "",
    hasHeaders: spec.hasHeaders !== false,
    seriesBy: spec.seriesBy === "rows" ? "rows" : "columns",
    legend: {
      visible: spec.legend ? spec.legend.visible !== false : true,
      position: spec.legend && spec.legend.position ? spec.legend.position : "Right"
    },
    position: {
      placement: spec.position && spec.position.placement ? spec.position.placement : "adjacent",
      targetCell: spec.position && spec.position.targetCell ? spec.position.targetCell : "",
      offsetColumns: spec.position && typeof spec.position.offsetColumns === "number" ? spec.position.offsetColumns : 2,
      offsetRows: spec.position && typeof spec.position.offsetRows === "number" ? spec.position.offsetRows : 2
    },
    dimensions: {
      width: spec.dimensions && typeof spec.dimensions.width === "number" ? spec.dimensions.width : 480,
      height: spec.dimensions && typeof spec.dimensions.height === "number" ? spec.dimensions.height : 300
    },
    dataLabels: spec.dataLabels ? {
      visible: spec.dataLabels.visible !== false,
      showValue: spec.dataLabels.showValue !== false,
      showPercentage: spec.dataLabels.showPercentage === true,
      showCategoryName: spec.dataLabels.showCategoryName === true
    } : null
  };
}

/**
 * Robustly extracts an Excel chart specification from AI response text or HTML
 * Handles ```excel-chart, ```chart, ```excel_chart, ```json code blocks and raw JSON objects.
 * @param {string} text - Raw AI response text
 * @param {string} htmlContent - Optional parsed HTML content
 * @returns {Object|null} - Normalized chart specification or null
 */
export function extractExcelChartSpec(text, htmlContent = "") {
  const combined = (text || "") + "\n" + (htmlContent ? htmlContent.replace(/<[^>]+>/g, " ") : "");
  if (!combined.trim()) return null;

  // 1. Check for fenced code blocks: ```excel-chart\n{...}\n``` or ```json\n{...}\n```
  const chartBlockRegex = /```(?:excel-chart|chart|excel_chart|json)?\s*([\s\S]*?)```/gi;
  let blockMatch;
  while ((blockMatch = chartBlockRegex.exec(combined)) !== null) {
    if (blockMatch[1]) {
      try {
        const cleaned = blockMatch[1].trim();
        if (cleaned.startsWith("{") && cleaned.endsWith("}")) {
          const parsed = JSON.parse(cleaned);
          if (parsed && (parsed.type || parsed.chartType || parsed.dataRange || parsed.range || parsed.categoriesRange || parsed.categoryRange || parsed.seriesRange || parsed.valuesRange || parsed.categories || parsed.series || parsed.sourceRange)) {
            return normalizeChartSpec(parsed);
          }
        }
      } catch (e) {
        // Continue searching
      }
    }
  }

  // 2. Fallback: Search for any JSON object containing chart type or range keys in any order
  const jsonObjectRegex = /\{[\s\S]*?\}/g;
  let objMatch;
  while ((objMatch = jsonObjectRegex.exec(combined)) !== null) {
    try {
      const parsed = JSON.parse(objMatch[0].trim());
      if (parsed && (parsed.type || parsed.chartType) && (parsed.dataRange || parsed.range || parsed.categoriesRange || parsed.categoryRange || parsed.seriesRange || parsed.valuesRange || parsed.categories || parsed.series || parsed.sourceRange || parsed.title)) {
        return normalizeChartSpec(parsed);
      }
    } catch (e) {
      // Continue
    }
  }

  return null;
}

/**
 * Normalizes a chart edit specification object into a clean structured schema
 * @param {Object} spec 
 * @returns {Object|null}
 */
export function normalizeChartEditSpec(spec) {
  if (!spec) return null;

  const rawType = spec.chartType || spec.type || "";
  let rawRange = spec.dataRange || spec.range || "";
  let categoriesRange = spec.categoriesRange || spec.categoryRange || spec.labelRange || spec.labelsRange || spec.xRange || spec.categoryAxis || "";
  let seriesRange = spec.seriesRange || spec.valuesRange || spec.valueRange || spec.yRange || spec.metricsRange || spec.metricRange || "";

  if (rawRange && String(rawRange).includes(",") && !categoriesRange && !seriesRange) {
    const parts = String(rawRange).split(",").map(p => p.trim()).filter(Boolean);
    if (parts.length >= 2) {
      categoriesRange = parts[0];
      seriesRange = parts.slice(1).join(", ");
      rawRange = "";
    }
  }

  const cleanRange = (r) => r ? String(r).replace(/[$]/g, "").replace(/^=|^["']|["']$/g, "").trim() : "";
  categoriesRange = cleanRange(categoriesRange);
  seriesRange = cleanRange(seriesRange);
  rawRange = cleanRange(rawRange);

  return {
    action: "edit",
    targetChart: spec.targetChart ? String(spec.targetChart).trim() : "",
    chartId: spec.chartId ? String(spec.chartId).trim() : "",
    chartType: rawType ? normalizeChartType(rawType) : undefined,
    rawType: rawType,
    title: spec.title !== undefined ? String(spec.title).trim() : undefined,
    dataRange: rawRange || undefined,
    categoriesRange: categoriesRange || undefined,
    seriesRange: seriesRange || undefined,
    seriesBy: spec.seriesBy === "rows" ? "rows" : (spec.seriesBy === "columns" ? "columns" : undefined),
    legend: spec.legend ? {
      visible: spec.legend.visible !== undefined ? Boolean(spec.legend.visible) : undefined,
      position: spec.legend.position ? String(spec.legend.position).trim() : undefined
    } : undefined,
    dataLabels: spec.dataLabels ? {
      visible: spec.dataLabels.visible !== undefined ? Boolean(spec.dataLabels.visible) : undefined,
      showCategoryName: spec.dataLabels.showCategoryName !== undefined ? Boolean(spec.dataLabels.showCategoryName) : undefined,
      showPercentage: spec.dataLabels.showPercentage !== undefined ? Boolean(spec.dataLabels.showPercentage) : undefined,
      showValue: spec.dataLabels.showValue !== undefined ? Boolean(spec.dataLabels.showValue) : undefined
    } : undefined,
    dimensions: spec.dimensions ? {
      width: typeof spec.dimensions.width === 'number' ? spec.dimensions.width : undefined,
      height: typeof spec.dimensions.height === 'number' ? spec.dimensions.height : undefined
    } : undefined,
    position: spec.position ? {
      targetCell: spec.position.targetCell ? cleanRange(spec.position.targetCell) : undefined
    } : undefined
  };
}

/**
 * Extracts an Excel chart edit specification from AI response text or HTML
 * Handles ```excel-chart-edit, ```chart-edit code blocks, chart update intent, or JSON objects with edit actions.
 * @param {string} text - Raw AI response text
 * @param {string} htmlContent - Optional parsed HTML content
 * @param {Object} options - Optional context parameters including userPrompt and prompt
 * @returns {Object|null} - Normalized chart edit specification or null
 */
export function extractExcelChartEditSpec(text, htmlContent = "", options = {}) {
  const combined = (text || "") + "\n" + (htmlContent ? htmlContent.replace(/<[^>]+>/g, " ") : "");
  if (!combined.trim()) return null;

  const userPrompt = (options && options.userPrompt) || "";
  const prompt = (options && options.prompt) || "";

  const isEditIntent = Boolean(
    /update.*chart|change.*chart|modify.*chart|switch.*chart|edit.*chart|convert.*chart|make\s+it\s+a.*chart|re-style.*chart|update\s+title|change.*title/i.test(userPrompt) ||
    /update.*chart|change.*chart|modify.*chart|switch.*chart|edit.*chart/i.test(prompt) ||
    /update.*chart|change.*chart|modify.*chart|switch.*chart|edit.*chart|updated.*chart|changed.*chart|modified.*chart/i.test(combined) ||
    /```(?:excel-chart-edit|chart-edit)/i.test(combined)
  );

  // 1. Check for fenced code blocks: ```excel-chart-edit\n{...}\n``` or ```chart-edit\n{...}\n```
  const editBlockRegex = /```(?:excel-chart-edit|chart-edit)\s*([\s\S]*?)```/gi;
  let blockMatch;
  while ((blockMatch = editBlockRegex.exec(combined)) !== null) {
    if (blockMatch[1]) {
      try {
        const cleaned = blockMatch[1].trim();
        if (cleaned.startsWith("{") && cleaned.endsWith("}")) {
          const parsed = JSON.parse(cleaned);
          if (parsed) {
            return normalizeChartEditSpec(parsed);
          }
        }
      } catch (e) {
        // Continue searching
      }
    }
  }

  // 2. Check for general code blocks with action: "edit" or targetChart or chart update intent
  const codeBlockRegex = /```(?:excel-chart|chart|json)?\s*([\s\S]*?)```/gi;
  while ((blockMatch = codeBlockRegex.exec(combined)) !== null) {
    if (blockMatch[1]) {
      try {
        const cleaned = blockMatch[1].trim();
        if (cleaned.startsWith("{") && cleaned.endsWith("}")) {
          const parsed = JSON.parse(cleaned);
          if (parsed && (parsed.action === "edit" || (parsed.targetChart && (parsed.chartType || parsed.type || parsed.title || parsed.categoriesRange || parsed.seriesRange || parsed.dataRange || parsed.legend || parsed.dataLabels)))) {
            return normalizeChartEditSpec(parsed);
          }
          if (isEditIntent && parsed && (parsed.type || parsed.chartType || parsed.title || parsed.dataRange || parsed.categoriesRange || parsed.seriesRange)) {
            return normalizeChartEditSpec({ ...parsed, action: "edit", targetChart: parsed.targetChart || "active" });
          }
        }
      } catch (e) {
        // Continue searching
      }
    }
  }

  // 3. Fallback: Search for any JSON object with action: "edit" and targetChart or chartType
  const jsonObjectRegex = /\{[\s\S]*?\}/g;
  let objMatch;
  while ((objMatch = jsonObjectRegex.exec(combined)) !== null) {
    try {
      const parsed = JSON.parse(objMatch[0].trim());
      if (parsed && (parsed.action === "edit" || (parsed.targetChart && (parsed.chartType || parsed.type || parsed.title)))) {
        return normalizeChartEditSpec(parsed);
      }
      if (isEditIntent && parsed && (parsed.chartType || parsed.type) && (parsed.title || parsed.dataRange || parsed.categoriesRange || parsed.seriesRange || parsed.targetChart || isEditIntent)) {
        return normalizeChartEditSpec({ ...parsed, action: "edit", targetChart: parsed.targetChart || "active" });
      }
    } catch (e) {
      // Continue
    }
  }

  // 4. If user asked for chart edit/update and a regular chart spec or natural language update was returned
  if (isEditIntent) {
    const regularChartSpec = extractExcelChartSpec(text, htmlContent);
    if (regularChartSpec) {
      return normalizeChartEditSpec({
        action: "edit",
        targetChart: "active",
        chartType: regularChartSpec.type,
        rawType: regularChartSpec.rawType,
        title: regularChartSpec.title,
        dataRange: regularChartSpec.dataRange,
        categoriesRange: regularChartSpec.categoriesRange,
        seriesRange: regularChartSpec.seriesRange,
        legend: regularChartSpec.legend,
        dataLabels: regularChartSpec.dataLabels,
        seriesBy: regularChartSpec.seriesBy
      });
    }

    const allText = userPrompt + " " + prompt + " " + combined;
    const typeMatch = allText.match(/(?:to|into|as|make it|change to|switch to|update to)\s+(?:a\s+|an\s+)?(column|bar|pie|line|doughnut|donut|area|scatter|waterfall)/i)
      || allText.match(/(column|bar|pie|line|doughnut|donut|area|scatter|waterfall)\s+chart/i);
    const titleMatch = combined.match(/(?:title|named)\s+(?:to\s+)?["']([^"'\n\r]+)["']/i);
    
    let resolvedType = typeMatch ? typeMatch[1].toLowerCase() : undefined;
    if (resolvedType === "donut") resolvedType = "doughnut";

    if (resolvedType || titleMatch) {
      return normalizeChartEditSpec({
        action: "edit",
        targetChart: "active",
        chartType: resolvedType,
        title: titleMatch ? titleMatch[1] : undefined
      });
    }

    // Default chart edit spec when user explicitly requested chart modification
    return normalizeChartEditSpec({
      action: "edit",
      targetChart: "active"
    });
  }

  return null;
}

/**
 * Modifies an existing native Excel chart object in-place on the worksheet
 * @param {Excel.RequestContext} context - Active Office.js request context
 * @param {Excel.Worksheet} sheet - Target worksheet
 * @param {Object} spec - Normalized chart edit specification
 * @returns {Promise<Excel.Chart>}
 */
export async function editNativeChart(context, sheet, spec) {
  if (!spec) {
    throw new Error("Invalid chart edit specification: specification is empty.");
  }

  let chart = null;

  // 1. Locate Target Chart by ID
  if (spec.chartId) {
    try {
      chart = sheet.charts.getItem(spec.chartId);
      chart.load("id, name, chartType, title/text, width, height");
      await context.sync();
    } catch (e) {
      chart = null;
    }
  }

  // 2. Locate Target Chart by Name
  if (!chart && spec.targetChart && spec.targetChart.toLowerCase() !== "active" && spec.targetChart.toLowerCase() !== "selected") {
    try {
      chart = sheet.charts.getItem(spec.targetChart);
      chart.load("id, name, chartType, title/text, width, height");
      await context.sync();
    } catch (e) {
      chart = null;
    }
  }

  // 3. Locate Target Chart from active selection
  if (!chart) {
    try {
      const active = context.workbook.getActiveChartOrNullObject();
      active.load("id, name, chartType, title/text, width, height");
      await context.sync();
      if (!active.isNullObject) {
        chart = active;
      }
    } catch (e) {
      chart = null;
    }
  }

  // 4. Fallback: Target the most recent chart on the worksheet (or the single chart)
  if (!chart) {
    try {
      const allCharts = sheet.charts;
      allCharts.load("items/id, items/name, items/chartType, items/title/text");
      await context.sync();
      if (allCharts.items && allCharts.items.length > 0) {
        chart = allCharts.items[allCharts.items.length - 1];
      }
    } catch (e) {
      chart = null;
    }
  }

  if (!chart) {
    throw new Error(`Target chart "${spec.targetChart || spec.chartId || 'active'}" could not be found on the worksheet.`);
  }

  const cleanAddr = (addr) => {
    if (!addr) return "";
    let s = String(addr).replace(/^=|^["']|["']$/g, "").trim();
    if (s.includes("!")) {
      s = s.split("!").pop().trim();
    }
    return s.replace(/[$]/g, "");
  };

  // 2. Morph Chart Type if specified
  const rawType = spec.chartType || spec.rawType;
  if (rawType) {
    const newType = normalizeChartType(rawType);
    chart.chartType = newType;
    console.log(`[Chart Edit] Mutated chart type to: ${newType}`);
  }

  // 3. Update Title if specified
  if (spec.title !== undefined) {
    const titleStr = String(spec.title).trim();
    if (titleStr) {
      chart.title.text = titleStr;
      chart.title.visible = true;
    } else {
      chart.title.visible = false;
    }
    console.log(`[Chart Edit] Updated chart title: "${titleStr}"`);
  }

  // 4. Update Source Data Ranges if specified
  const catAddr = cleanAddr(spec.categoriesRange);
  const seriesAddr = cleanAddr(spec.seriesRange);
  const dataAddr = cleanAddr(spec.dataRange);

  const activeType = chart.chartType ? chart.chartType : (rawType ? normalizeChartType(rawType) : "");
  const isPieOrDoughnut = activeType.toLowerCase().includes("pie") || activeType.toLowerCase().includes("doughnut");
  const editSeriesBy = spec.seriesBy === "rows" ? "Rows" : "Columns";

  if (seriesAddr) {
    const seriesParts = seriesAddr.split(",").map(p => cleanAddr(p)).filter(Boolean);
    const primarySeries = seriesParts[0];
    const catParsed = parseRangeAddress(catAddr);
    const seriesParsed = parseRangeAddress(primarySeries);
    const rowsAlign = catParsed && seriesParsed && (catParsed.startRow === seriesParsed.startRow) && (catParsed.endRow === seriesParsed.endRow);
    const isAdjacent = rowsAlign && (
      (catParsed.endColIndex + 1 === seriesParsed.startColIndex) ||
      (seriesParsed.endColIndex + 1 === catParsed.startColIndex)
    );

    if (isAdjacent) {
      const minColIdx = Math.min(catParsed.startColIndex, seriesParsed.startColIndex);
      const maxColIdx = Math.max(catParsed.endColIndex, seriesParsed.endColIndex);
      const minRow = Math.min(catParsed.startRow, seriesParsed.startRow);
      const maxRow = Math.max(catParsed.endRow, seriesParsed.endRow);
      const contiguousAddr = `${columnIndexToLetter(minColIdx)}${minRow}:${columnIndexToLetter(maxColIdx)}${maxRow}`;
      chart.setData(sheet.getRange(contiguousAddr), "Columns");
      console.log(`[Chart Edit] Rebound to contiguous range: ${contiguousAddr}`);
    } else {
      chart.setData(sheet.getRange(seriesAddr), "Columns");
      if (!isPieOrDoughnut && catAddr) {
        try {
          chart.axes.categoryAxis.setCategoryNames(sheet.getRange(catAddr));
        } catch (axErr) {
          console.debug("Category axis notice:", axErr);
        }
      }
    }
  } else if (dataAddr) {
    chart.setData(sheet.getRange(dataAddr), editSeriesBy);
    if (!isPieOrDoughnut && catAddr) {
      try {
        chart.axes.categoryAxis.setCategoryNames(sheet.getRange(catAddr));
      } catch (axErr) {
        console.debug("Category axis notice:", axErr);
      }
    }
  }

  // 5. Configure Legend
  if (spec.legend) {
    try {
      if (spec.legend.visible !== undefined) {
        chart.legend.visible = Boolean(spec.legend.visible);
      }
      if (spec.legend.position) {
        chart.legend.position = spec.legend.position;
      }
    } catch (legErr) {
      console.debug("Legend edit notice:", legErr);
    }
  }

  // 6. Configure Data Labels
  if (spec.dataLabels) {
    try {
      const series0 = chart.series.getItemAt(0);
      if (spec.dataLabels.visible !== undefined) {
        series0.hasDataLabels = Boolean(spec.dataLabels.visible);
      }
      if (spec.dataLabels.showCategoryName !== undefined) {
        series0.dataLabels.showCategoryName = Boolean(spec.dataLabels.showCategoryName);
      }
      if (spec.dataLabels.showPercentage !== undefined) {
        series0.dataLabels.showPercentage = Boolean(spec.dataLabels.showPercentage);
      }
      if (spec.dataLabels.showValue !== undefined) {
        series0.dataLabels.showValue = Boolean(spec.dataLabels.showValue);
      }
    } catch (dlErr) {
      console.debug("Data labels edit notice:", dlErr);
    }
  }

  // 7. Configure Dimensions & Placement
  if (spec.dimensions) {
    try {
      if (typeof spec.dimensions.width === "number" && spec.dimensions.width >= 200 && spec.dimensions.width <= 1200) {
        chart.width = spec.dimensions.width;
      }
      if (typeof spec.dimensions.height === "number" && spec.dimensions.height >= 150 && spec.dimensions.height <= 800) {
        chart.height = spec.dimensions.height;
      }
    } catch (dimErr) {
      console.debug("Dimensions edit notice:", dimErr);
    }
  }

  if (spec.position && spec.position.targetCell) {
    try {
      const cellRef = cleanAddr(spec.position.targetCell);
      if (/^[A-Za-z]+[0-9]+$/.test(cellRef)) {
        chart.setPosition(cellRef);
      }
    } catch (posErr) {
      console.debug("Position edit notice:", posErr);
    }
  }

  await context.sync();
  return chart;
}

/**
 * Automatically creates a formula-linked Aggregated Summary Table (=SUMIF / =AVERAGEIF / =COUNTIF)
 * and generates a native Excel chart bound to the aggregated summary data.
 * 
 * @param {Excel.RequestContext} context
 * @param {Excel.Worksheet} sheet
 * @param {Object} spec
 * @param {Object} catResolved
 * @param {Array<Object>} seriesResolvedList
 * @param {Object} usedBounds
 * @param {string} chartTypeName
 * @param {boolean} isPieOrDoughnut
 * @param {string} seriesBy
 * @returns {Promise<Excel.Chart|null>}
 */
export async function createAggregatedSummaryTableAndChart(
  context,
  sheet,
  spec,
  catResolved,
  seriesResolvedList,
  usedBounds,
  chartTypeName,
  isPieOrDoughnut,
  seriesBy = "Columns"
) {
  if (!catResolved || !seriesResolvedList || seriesResolvedList.length === 0) {
    return null;
  }

  try {
    const catRange = sheet.getRange(catResolved.address);
    catRange.load("values, text, numberFormat, address, rowIndex, columnIndex, rowCount, columnCount");

    const seriesRanges = seriesResolvedList.map(s => {
      const r = sheet.getRange(s.address);
      r.load("values, text, numberFormat, address, rowIndex, columnIndex, rowCount, columnCount");
      return r;
    });

    await context.sync();

    const catValues = catRange.values || [];
    if (catValues.length === 0) return null;

    // Determine if first row is header
    let catHeaderName = "";
    let dataStartIdx = 0; // index in catValues array

    if (catResolved.startRow === 1 && spec.hasHeaders !== false) {
      catHeaderName = String(catValues[0][0] || "Category").trim();
      dataStartIdx = 1;
    } else if (catResolved.startRow > 1) {
      // Range starts at row > 1 (e.g. D2:D7). Check if row 1 has a header for this column.
      try {
        const hCell = sheet.getRange(`${catResolved.startCol}1`);
        hCell.load("values");
        await context.sync();
        if (hCell.values && hCell.values[0] && hCell.values[0][0]) {
          catHeaderName = String(hCell.values[0][0]).trim();
        }
      } catch (e) {}
      dataStartIdx = 0;
    }
    if (!catHeaderName) catHeaderName = "Category";

    // Extract category data items
    const categoryItems = [];
    for (let r = dataStartIdx; r < catValues.length; r++) {
      const cellVal = catValues[r][0];
      if (cellVal !== null && cellVal !== undefined && String(cellVal).trim() !== "") {
        categoryItems.push(String(cellVal).trim());
      }
    }

    if (categoryItems.length === 0) return null;

    // Find unique categories (preserving first-seen order)
    const uniqueCategories = [];
    categoryItems.forEach(item => {
      if (!uniqueCategories.includes(item)) {
        uniqueCategories.push(item);
      }
    });

    const isAggregateExplicit = Boolean(
      spec.aggregate || spec.aggregation || spec.groupBy || spec.group_by ||
      (spec.title && /total|subtotal|distribution|summary|by\s+category/i.test(spec.title))
    );

    const hasDuplicates = uniqueCategories.length < categoryItems.length && uniqueCategories.length > 0;

    // If neither explicit aggregation requested nor duplicates exist, fallback to regular chart
    if (!isAggregateExplicit && !hasDuplicates) {
      return null;
    }

    console.log(`[Aggregation Engine] Detected ${uniqueCategories.length} unique categories across ${categoryItems.length} rows. Building Aggregated Summary Table...`);

    // Determine aggregation function
    const rawAgg = String(spec.aggregate || spec.aggregation || "SUM").toUpperCase();
    let funcName = "SUMIF";
    let funcPrefix = "Total ";
    if (rawAgg.includes("AVG") || rawAgg.includes("AVERAGE")) {
      funcName = "AVERAGEIF";
      funcPrefix = "Average ";
    } else if (rawAgg.includes("COUNT")) {
      funcName = "COUNTIF";
      funcPrefix = "Count of ";
    }

    // Prepare metadata for each series column
    const seriesMetaList = [];
    for (let sIdx = 0; sIdx < seriesRanges.length; sIdx++) {
      const sRange = seriesRanges[sIdx];
      const sResolved = seriesResolvedList[sIdx];
      let seriesHeaderName = "";

      if (sResolved.startRow === 1 && spec.hasHeaders !== false && sRange.values && sRange.values[0]) {
        seriesHeaderName = String(sRange.values[0][0] || `Metric ${sIdx + 1}`).trim();
      } else if (sResolved.startRow > 1) {
        try {
          const hCell = sheet.getRange(`${sResolved.startCol}1`);
          hCell.load("values");
          await context.sync();
          if (hCell.values && hCell.values[0] && hCell.values[0][0]) {
            seriesHeaderName = String(hCell.values[0][0]).trim();
          }
        } catch (e) {}
      }
      if (!seriesHeaderName) seriesHeaderName = `Value ${sIdx + 1}`;

      // Clean header prefix if seriesHeaderName already has "Total" / "Amount"
      let finalSeriesHeader = seriesHeaderName;
      if (!seriesHeaderName.toLowerCase().startsWith(funcPrefix.toLowerCase().trim())) {
        finalSeriesHeader = `${funcPrefix}${seriesHeaderName}`;
      }

      // Detect number format
      let numFmt = "$#,##0.00";
      if (sRange.numberFormat && Array.isArray(sRange.numberFormat) && sRange.numberFormat.length > 0) {
        const rowFmt = sRange.numberFormat[dataStartIdx] || sRange.numberFormat[0];
        if (rowFmt && rowFmt[0] && rowFmt[0] !== "General") {
          numFmt = rowFmt[0];
        }
      }

      const srcCatStartRow = (catResolved.startRow === 1 && spec.hasHeaders !== false) ? 2 : catResolved.startRow;
      const srcCatEndRow = catResolved.endRow;
      const srcSeriesStartRow = (sResolved.startRow === 1 && spec.hasHeaders !== false) ? 2 : sResolved.startRow;
      const srcSeriesEndRow = sResolved.endRow;

      seriesMetaList.push({
        header: finalSeriesHeader,
        numFmt: numFmt,
        srcCatRange: `$${catResolved.startCol}$${srcCatStartRow}:$${catResolved.endCol}$${srcCatEndRow}`,
        srcSeriesRange: `$${sResolved.startCol}$${srcSeriesStartRow}:$${sResolved.endCol}$${srcSeriesEndRow}`
      });
    }

    // Position summary table at the user's selected cell location if available
    let summaryStartColIdx = (usedBounds.endColIndex || 3) + 2;
    let summaryStartRow = Math.min(catResolved.startRow, usedBounds.startRow || 1);

    try {
      const selectedRange = context.workbook.getSelectedRange();
      selectedRange.load("rowIndex, columnIndex, rowCount, columnCount");
      await context.sync();

      if (selectedRange && selectedRange.rowIndex !== undefined && selectedRange.columnIndex !== undefined) {
        const isSingleCell = (selectedRange.rowCount === 1 && selectedRange.columnCount === 1);
        const isDisjointFromSource = (
          selectedRange.columnIndex > (usedBounds.endColIndex || 0) ||
          selectedRange.columnIndex < (usedBounds.startColIndex || 0) ||
          selectedRange.rowIndex > (usedBounds.endRow || 0) ||
          selectedRange.rowIndex < ((usedBounds.startRow || 1) - 1)
        );

        if (isSingleCell || isDisjointFromSource) {
          summaryStartColIdx = selectedRange.columnIndex;
          summaryStartRow = selectedRange.rowIndex + 1;
        }
      }
    } catch (selErr) {
      console.debug("Selected range query notice:", selErr);
    }

    const totalSummaryCols = 1 + seriesMetaList.length;
    const summaryEndColIdx = summaryStartColIdx + totalSummaryCols - 1;
    const summaryEndRow = summaryStartRow + uniqueCategories.length;

    const summaryStartColLetter = columnIndexToLetter(summaryStartColIdx);
    const summaryEndColLetter = columnIndexToLetter(summaryEndColIdx);
    const catColLetter = summaryStartColLetter;
    const summaryTableAddr = `${summaryStartColLetter}${summaryStartRow}:${summaryEndColLetter}${summaryEndRow}`;

    // 1. Write Header Row
    const headerRowAddr = `${summaryStartColLetter}${summaryStartRow}:${summaryEndColLetter}${summaryStartRow}`;
    const headerRange = sheet.getRange(headerRowAddr);
    headerRange.values = [[catHeaderName, ...seriesMetaList.map(s => s.header)]];
    headerRange.format.font.bold = true;
    headerRange.format.font.color = "#1F2937";
    headerRange.format.fill.color = "#EBF3FC";
    headerRange.format.horizontalAlignment = "Center";

    // 2. Populate Category Rows & Formulas
    for (let i = 0; i < uniqueCategories.length; i++) {
      const rowNum = summaryStartRow + 1 + i;
      const catVal = uniqueCategories[i];

      // Set Category Cell
      const catCell = sheet.getRange(`${catColLetter}${rowNum}`);
      catCell.values = [[catVal]];
      catCell.format.horizontalAlignment = "Left";

      // Set Series Formula Cells
      for (let sIdx = 0; sIdx < seriesMetaList.length; sIdx++) {
        const meta = seriesMetaList[sIdx];
        const sColLetter = columnIndexToLetter(summaryStartColIdx + 1 + sIdx);
        const sCell = sheet.getRange(`${sColLetter}${rowNum}`);

        let formula = "";
        if (funcName === "COUNTIF") {
          formula = `=${funcName}(${meta.srcCatRange}, ${catColLetter}${rowNum})`;
        } else {
          formula = `=${funcName}(${meta.srcCatRange}, ${catColLetter}${rowNum}, ${meta.srcSeriesRange})`;
        }

        sCell.formulas = [[formula]];
        sCell.numberFormat = [[meta.numFmt]];
        sCell.format.horizontalAlignment = "Right";
      }
    }

    // 3. Format Borders and Autofit Summary Table
    const summaryRange = sheet.getRange(summaryTableAddr);
    try {
      summaryRange.format.borders.getItem("InsideHorizontal").style = "Continuous";
      summaryRange.format.borders.getItem("InsideHorizontal").color = "#E5E7EB";
      summaryRange.format.borders.getItem("EdgeBottom").style = "Continuous";
      summaryRange.format.borders.getItem("EdgeBottom").color = "#9CA3AF";
      summaryRange.format.autofitColumns();
    } catch (fmtErr) {
      console.debug("Summary table format notice:", fmtErr);
    }

    await context.sync();

    // 4. Create Native Chart from the Aggregated Summary Range
    const chart = sheet.charts.add(chartTypeName, summaryRange, "Columns");

    // Position chart in-line, one row below the summary table (2 rows below endRow)
    if (!spec.position || !spec.position.targetCell) {
      const chartPlacementRow = summaryEndRow + 2;
      spec.position = {
        ...spec.position,
        placement: "targetCell",
        targetCell: `${summaryStartColLetter}${chartPlacementRow}`
      };
    }

    console.log(`[Aggregated Summary Chart] Created ${chartTypeName} chart bound to summary table: ${summaryTableAddr}`);
    return chart;
  } catch (err) {
    console.warn("Aggregated summary table creation error, falling back:", err);
    return null;
  }
}

/**
 * Inserts a native Excel chart object onto the worksheet
 * Supports targeted multi-range field selection (e.g., categoriesRange + seriesRange)
 * as well as single contiguous dataRange tables and intelligent used-range inference.
 * 
 * @param {Excel.RequestContext} context - Active Office.js request context
 * @param {Excel.Worksheet} sheet - Target worksheet
 * @param {Object} spec - Normalized chart specification
 * @returns {Promise<Excel.Chart>}
 */
export async function insertNativeChart(context, sheet, spec, adapterInstance = null) {
  if (!spec) {
    throw new Error("Invalid chart specification: specification is empty.");
  }

  const cleanAddr = (addr) => {
    if (!addr) return "";
    let s = String(addr).replace(/^=|^["']|["']$/g, "").trim();
    if (s.includes("!")) {
      s = s.split("!").pop().trim();
    }
    return s.replace(/[$]/g, "");
  };

  const chartTypeName = normalizeChartType(spec.type || spec.rawType);
  const isPieOrDoughnut = chartTypeName.toLowerCase().includes("pie") || chartTypeName.toLowerCase().includes("doughnut");
  const seriesBy = spec.seriesBy === "rows" ? "Rows" : "Columns";

  // 1. Inspect active worksheet used range to obtain ground-truth bounding coordinates
  let usedBounds = {
    startColIndex: 0,
    endColIndex: 3,
    startRow: 1,
    endRow: 10,
    rowCount: 10,
    columnCount: 4,
    hasData: false
  };

  try {
    const usedRange = sheet.getUsedRange(true);
    usedRange.load("address, rowCount, columnCount, rowIndex, columnIndex, values");
    await context.sync();
    if (usedRange && usedRange.rowCount > 0 && usedRange.columnCount > 0) {
      usedBounds = {
        startColIndex: usedRange.columnIndex || 0,
        endColIndex: (usedRange.columnIndex || 0) + (usedRange.columnCount || 1) - 1,
        startRow: (usedRange.rowIndex || 0) + 1,
        endRow: (usedRange.rowIndex || 0) + (usedRange.rowCount || 1),
        rowCount: usedRange.rowCount,
        columnCount: usedRange.columnCount,
        hasData: Array.isArray(usedRange.values) && usedRange.values.length > 0 && usedRange.values.some(row => Array.isArray(row) && row.some(cell => cell !== null && cell !== undefined && String(cell).trim() !== ""))
      };
    }
  } catch (e) {
    console.debug("Used range query notice:", e);
  }

  let chart = null;

  // 2. Priority 1A: Targeted Categories + Series Ranges
  const catAddr = cleanAddr(spec.categoriesRange);
  const seriesAddr = cleanAddr(spec.seriesRange);

  if (seriesAddr) {
    try {
      const seriesParts = seriesAddr.split(",").map(p => cleanAddr(p)).filter(Boolean);
      const catResolved = catAddr ? resolveRangeAddress(catAddr, usedBounds) : null;
      const seriesResolvedList = seriesParts.map(p => resolveRangeAddress(p, usedBounds)).filter(Boolean);

      if (seriesResolvedList.length > 0) {
        const primarySeries = seriesResolvedList[0];

        // Priority 1A-Agg: Check for duplicate category rows or explicit aggregation
        if (catResolved) {
          const aggChart = await createAggregatedSummaryTableAndChart(
            context,
            sheet,
            spec,
            catResolved,
            seriesResolvedList,
            usedBounds,
            chartTypeName,
            isPieOrDoughnut,
            seriesBy
          );
          if (aggChart) {
            chart = aggChart;
          }
        }

        if (!chart) {
          // Determine whether categories and all series ranges form a contiguous column block
          let isContiguousBlock = false;
          let contiguousAddr = "";

          if (catResolved) {
            const allCols = [];
            for (let c = catResolved.startColIndex; c <= catResolved.endColIndex; c++) allCols.push(c);
            seriesResolvedList.forEach(s => {
              for (let c = s.startColIndex; c <= s.endColIndex; c++) allCols.push(c);
            });
            const uniqueCols = [...new Set(allCols)].sort((a, b) => a - b);
            const minCol = uniqueCols[0];
            const maxCol = uniqueCols[uniqueCols.length - 1];
            const minRow = Math.min(catResolved.startRow, ...seriesResolvedList.map(s => s.startRow));
            const maxRow = Math.max(catResolved.endRow, ...seriesResolvedList.map(s => s.endRow));

            // Contiguous if uniqueCols covers all columns from minCol to maxCol without gaps
            if (uniqueCols.length === (maxCol - minCol + 1)) {
              isContiguousBlock = true;
              contiguousAddr = `${columnIndexToLetter(minCol)}${minRow}:${columnIndexToLetter(maxCol)}${maxRow}`;
            }
          }

          if (isContiguousBlock && contiguousAddr) {
            // Contiguous multi-column block (e.g. Column A category + Column B..D series -> A1:D10)
            // Office.js sheet.charts.add natively binds leftmost column as category labels and subsequent columns as series
            chart = sheet.charts.add(chartTypeName, sheet.getRange(contiguousAddr), seriesBy);
            console.log(`[Contiguous Multi-Range Chart] Created chart for contiguous range: ${contiguousAddr}`);

            // If not pie chart, reinforce category axis
            if (!isPieOrDoughnut && catResolved) {
              try {
                let catLabelAddr = catResolved.address;
                if (primarySeries.startRow === 1 && catResolved.startRow === 1 && catResolved.endRow > 1) {
                  catLabelAddr = `${catResolved.startCol}2:${catResolved.endCol}${catResolved.endRow}`;
                }
                chart.axes.categoryAxis.setCategoryNames(sheet.getRange(catLabelAddr));
              } catch (e) {
                console.debug("Category axis notice:", e);
              }
            }
          } else if (isPieOrDoughnut && catResolved && primarySeries) {
            // Pie / Doughnut chart with non-adjacent columns (e.g. Column A for Labels, Column D for Values)
            // Since Pie charts do not have a category axis, create a dynamically formula-linked 2-column helper range
            const helperColStart = usedBounds.endColIndex + 3;
            const hCol1 = columnIndexToLetter(helperColStart);
            const hCol2 = columnIndexToLetter(helperColStart + 1);
            const hStartRow = Math.min(catResolved.startRow, primarySeries.startRow);
            const hEndRow = Math.max(catResolved.endRow, primarySeries.endRow);
            const helperRange = sheet.getRange(`${hCol1}${hStartRow}:${hCol2}${hEndRow}`);

            const formulas = [];
            for (let r = hStartRow; r <= hEndRow; r++) {
              formulas.push([`=${catResolved.startCol}${r}`, `=${primarySeries.startCol}${r}`]);
            }
            helperRange.formulas = formulas;
            await context.sync();

            chart = sheet.charts.add(chartTypeName, helperRange, "Columns");
            console.log(`[Pie Helper Range] Bound non-adjacent categories (${catResolved.address}) and series (${primarySeries.address}) via helper range ${hCol1}${hStartRow}:${hCol2}${hEndRow}`);
          } else {
            // Disjoint Cartesian chart (Column, Bar, Line, Area, etc.)
            const primaryRange = sheet.getRange(primarySeries.address);
            chart = sheet.charts.add(chartTypeName, primaryRange, "Columns");
            console.log(`[Disjoint Chart] Created primary series range: ${primarySeries.address}`);

            // Attach extra series columns if specified
            if (seriesResolvedList.length > 1) {
              for (let i = 1; i < seriesResolvedList.length; i++) {
                try {
                  const extraRange = sheet.getRange(seriesResolvedList[i].address);
                  const addedSeries = chart.series.add();
                  addedSeries.setValues(extraRange);
                } catch (extraErr) {
                  console.warn(`Could not attach extra series range "${seriesResolvedList[i].address}":`, extraErr);
                }
              }
            }

            // Set category axis labels
            if (catResolved) {
              try {
                let catLabelAddr = catResolved.address;
                if (primarySeries.startRow === 1 && catResolved.startRow === 1 && catResolved.endRow > 1) {
                  catLabelAddr = `${catResolved.startCol}2:${catResolved.endCol}${catResolved.endRow}`;
                }
                const catRange = sheet.getRange(catLabelAddr);
                chart.axes.categoryAxis.setCategoryNames(catRange);
                console.log(`[Disjoint Chart] Set category axis labels from: ${catLabelAddr}`);
              } catch (catErr) {
                console.warn("Could not set category names axis:", catErr);
              }
            }
          }
        }
      }
    } catch (multiErr) {
      console.warn("Multi-range chart creation notice, falling back to contiguous range:", multiErr);
      chart = null;
    }
  }

  // 3. Priority 1B & 2: Contiguous dataRange or Worksheet Content Inference
  if (!chart) {
    let rangeAddr = cleanAddr(spec.dataRange);
    let sourceRange = null;

    if (rangeAddr) {
      const resolved = resolveRangeAddress(rangeAddr, usedBounds);
      if (resolved) {
        try {
          sourceRange = sheet.getRange(resolved.address);
          console.log(`[Data Range] Using explicit data range: ${resolved.address}`);
        } catch (e) {
          sourceRange = null;
        }
      }
    }

    if (!sourceRange && usedBounds.hasData) {
      try {
        const startCol = columnIndexToLetter(usedBounds.startColIndex);
        const endCol = columnIndexToLetter(usedBounds.endColIndex);
        const autoAddr = `${startCol}${usedBounds.startRow}:${endCol}${usedBounds.endRow}`;
        sourceRange = sheet.getRange(autoAddr);
        console.log(`[Worksheet Content] Inferred data range from sheet bounds: ${autoAddr}`);
      } catch (usedErr) {
        sourceRange = null;
      }
    }

    if (!sourceRange) {
      try {
        const sel = context.workbook.getSelectedRange();
        sel.load("rowCount, columnCount, address");
        await context.sync();
        if ((sel.rowCount > 1 || sel.columnCount > 1)) {
          sourceRange = sel;
          console.log(`[Selection Fallback] Using active selected range`);
        }
      } catch (selErr) {
        sourceRange = null;
      }
    }

    if (!sourceRange) {
      throw new Error("Unable to create chart: No data table found on worksheet. Please populate or select spreadsheet data first.");
    }

    // Check if sourceRange has multiple columns and needs automatic aggregation
    try {
      sourceRange.load("columnCount, rowCount, columnIndex, rowIndex, address");
      await context.sync();
      if (sourceRange.columnCount >= 2 && sourceRange.rowCount >= 2) {
        const startColIdx = sourceRange.columnIndex;
        const endColIdx = sourceRange.columnIndex + sourceRange.columnCount - 1;
        const startR = sourceRange.rowIndex + 1;
        const endR = sourceRange.rowIndex + sourceRange.rowCount;
        const firstColLetter = columnIndexToLetter(startColIdx);
        const catResolvedAuto = {
          startCol: firstColLetter,
          startColIndex: startColIdx,
          startRow: startR,
          endCol: firstColLetter,
          endColIndex: startColIdx,
          endRow: endR,
          address: `${firstColLetter}${startR}:${firstColLetter}${endR}`
        };
        const seriesResolvedAuto = [];
        for (let c = startColIdx + 1; c <= endColIdx; c++) {
          const colL = columnIndexToLetter(c);
          seriesResolvedAuto.push({
            startCol: colL,
            startColIndex: c,
            startRow: startR,
            endCol: colL,
            endColIndex: c,
            endRow: endR,
            address: `${colL}${startR}:${colL}${endR}`
          });
        }
        const aggChart = await createAggregatedSummaryTableAndChart(
          context,
          sheet,
          spec,
          catResolvedAuto,
          seriesResolvedAuto,
          usedBounds,
          chartTypeName,
          isPieOrDoughnut,
          seriesBy
        );
        if (aggChart) {
          chart = aggChart;
        }
      }
    } catch (autoAggErr) {
      console.debug("Auto-aggregation check notice:", autoAggErr);
    }

    if (!chart) {
      chart = sheet.charts.add(chartTypeName, sourceRange, seriesBy);

      // If categoriesRange was explicitly provided on a contiguous range
      if (catAddr && !isPieOrDoughnut) {
        try {
          const catResolved = resolveRangeAddress(catAddr, usedBounds);
          if (catResolved) {
            chart.axes.categoryAxis.setCategoryNames(sheet.getRange(catResolved.address));
          }
        } catch (catErr) {
          console.debug("Category axis notice:", catErr);
        }
      }
    }
  }

  // 4. Title Configuration
  if (spec.title) {
    try {
      chart.title.text = String(spec.title).trim();
      chart.title.visible = true;
    } catch (titleErr) {
      console.warn("Chart title configuration notice:", titleErr);
    }
  }

  // 5. Legend Configuration
  try {
    const showLegend = spec.legend ? spec.legend.visible !== false : true;
    chart.legend.visible = showLegend;
    if (showLegend && spec.legend && spec.legend.position) {
      chart.legend.position = spec.legend.position;
    }
  } catch (legErr) {
    console.debug("Legend setup notice:", legErr);
  }

  // 6. Data Labels Configuration
  if (isPieOrDoughnut) {
    try {
      const series0 = chart.series.getItemAt(0);
      series0.hasDataLabels = true;
      series0.dataLabels.showCategoryName = true;
      series0.dataLabels.showPercentage = true;
      series0.dataLabels.showValue = false;
    } catch (pieLabelErr) {
      console.debug("Pie chart data label setup notice:", pieLabelErr);
    }
  } else if (spec.dataLabels) {
    try {
      const series0 = chart.series.getItemAt(0);
      series0.hasDataLabels = spec.dataLabels.visible !== false;
      if (spec.dataLabels.showValue !== undefined) series0.dataLabels.showValue = spec.dataLabels.showValue;
      if (spec.dataLabels.showPercentage !== undefined) series0.dataLabels.showPercentage = spec.dataLabels.showPercentage;
      if (spec.dataLabels.showCategoryName !== undefined) series0.dataLabels.showCategoryName = spec.dataLabels.showCategoryName;
    } catch (lblErr) {
      console.debug("Data label setup notice:", lblErr);
    }
  }

  // 7. Optimal Chart Placement
  try {
    const pos = spec.position || {};
    let targetCellAddr = "";

    if (pos.placement === "targetCell" && pos.targetCell) {
      let cellRef = cleanAddr(pos.targetCell);
      if (cellRef.includes(":")) cellRef = cellRef.split(":")[0].trim();
      if (/^[A-Za-z]+[0-9]+$/.test(cellRef)) targetCellAddr = cellRef;
    }

    if (!targetCellAddr) {
      // Check if user selected a cell
      let selectedRange = null;
      try {
        selectedRange = context.workbook.getSelectedRange();
        selectedRange.load("rowIndex, columnIndex, rowCount, columnCount, address");
        await context.sync();
      } catch (e) {}

      if (selectedRange && selectedRange.rowIndex !== undefined && selectedRange.columnIndex !== undefined) {
        const isSingleCell = (selectedRange.rowCount === 1 && selectedRange.columnCount === 1);
        const isDisjointFromSource = (
          selectedRange.columnIndex > (usedBounds.endColIndex || 0) ||
          selectedRange.columnIndex < (usedBounds.startColIndex || 0) ||
          selectedRange.rowIndex > (usedBounds.endRow || 0) ||
          selectedRange.rowIndex < ((usedBounds.startRow || 1) - 1)
        );

        if (isSingleCell || isDisjointFromSource) {
          targetCellAddr = `${columnIndexToLetter(selectedRange.columnIndex)}${selectedRange.rowIndex + 1}`;
        }
      }

      if (!targetCellAddr) {
        // Position safely in-line, one row below the data table (2 rows below endRow)
        const chartPlacementRow = (usedBounds.endRow || 1) + 2;
        const colLetter = columnIndexToLetter(usedBounds.startColIndex !== undefined ? usedBounds.startColIndex : 0);
        targetCellAddr = `${colLetter}${chartPlacementRow}`;
      }
    }

    chart.setPosition(targetCellAddr);
    console.log(`[Chart Position] Positioned chart at cell: ${targetCellAddr}`);
  } catch (posErr) {
    console.warn("Chart positioning notice:", posErr);
  }

  // 8. Dimensions
  if (spec.dimensions) {
    try {
      if (typeof spec.dimensions.width === "number" && spec.dimensions.width >= 200 && spec.dimensions.width <= 1200) {
        chart.width = spec.dimensions.width;
      }
      if (typeof spec.dimensions.height === "number" && spec.dimensions.height >= 150 && spec.dimensions.height <= 800) {
        chart.height = spec.dimensions.height;
      }
    } catch (dimErr) {
      console.warn("Chart dimensions notice:", dimErr);
    }
  }

  // 9. Load Chart Metadata & Lock Chart on Adapter
  try {
    chart.load("id, name, chartType, width, height");
    await context.sync();
    if (adapterInstance) {
      adapterInstance.lockedChart = {
        id: chart.id,
        name: chart.name,
        chartType: chartTypeName,
        title: spec.title || chart.name,
        width: Math.round(chart.width || spec.dimensions?.width || 480),
        height: Math.round(chart.height || spec.dimensions?.height || 300)
      };
    }
  } catch (e) {
    console.debug("Chart sync notice:", e);
  }

  return chart;
}

/**
 * Map of friendly/common color names to standard hex color codes
 */
export const NAMED_COLORS = {
  lightblue: "#ADD8E6",
  lightgreen: "#90EE90",
  lightyellow: "#FFFFE0",
  lightgray: "#D3D3D3",
  lightgrey: "#D3D3D3",
  lightcoral: "#F08080",
  lightcyan: "#E0FFFF",
  lightpink: "#FFB6C1",
  yellow: "#FFFF00",
  red: "#FF0000",
  green: "#008000",
  darkgreen: "#006400",
  blue: "#0000FF",
  darkblue: "#00008B",
  navy: "#000080",
  teal: "#008080",
  cyan: "#00FFFF",
  orange: "#FFA500",
  darkorange: "#FF8C00",
  purple: "#800080",
  pink: "#FFC0CB",
  magenta: "#FF00FF",
  gold: "#FFD700",
  gray: "#808080",
  grey: "#808080",
  darkgray: "#A9A9A9",
  darkgrey: "#A9A9A9",
  white: "#FFFFFF",
  black: "#000000",
  transparent: ""
};

/**
 * Normalizes a color string to hex or clean CSS color
 * @param {string} val 
 * @returns {string}
 */
export function normalizeColor(val) {
  if (!val) return "";
  let clean = String(val).trim().toLowerCase();
  clean = clean.replace(/^["']|["']$/g, "").trim();
  if (NAMED_COLORS[clean] !== undefined) {
    return NAMED_COLORS[clean];
  }
  if (clean.startsWith("#")) {
    return clean;
  }
  if (/^[0-9a-f]{6}$/i.test(clean) || /^[0-9a-f]{3}$/i.test(clean)) {
    return "#" + clean;
  }
  return clean;
}

/**
 * Normalizes a number format string or preset alias to Excel number format pattern
 * @param {string} fmt 
 * @returns {string}
 */
export function normalizeNumberFormat(fmt) {
  if (!fmt) return "";
  let clean = String(fmt).trim();
  clean = clean.replace(/^["']|["']$/g, "").trim();
  const lower = clean.toLowerCase();

  switch (lower) {
    case "currency":
    case "dollar":
    case "$":
      return "$#,##0.00";
    case "currency-int":
    case "currency0":
      return "$#,##0";
    case "percent":
    case "percentage":
    case "%":
      return "0.0%";
    case "percent2":
    case "percentage2":
      return "0.00%";
    case "number":
    case "integer":
    case "comma":
      return "#,##0";
    case "decimal":
      return "#,##0.00";
    case "date":
      return "yyyy-mm-dd";
    case "short-date":
      return "m/d/yyyy";
    case "text":
      return "@";
    case "accounting":
      return "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)";
    default:
      return clean;
  }
}

/**
 * Normalizes a range formatting/modification specification into a clean structured schema
 * @param {Object} spec 
 * @returns {Object|null}
 */
export function normalizeRangeFormatSpec(spec) {
  if (!spec) return null;

  // Range address (clean)
  let targetRange = spec.targetRange || spec.range || spec.address || "";
  if (targetRange && (targetRange.toLowerCase() === "active" || targetRange.toLowerCase() === "selected" || targetRange.toLowerCase() === "selection")) {
    targetRange = "";
  }
  if (targetRange) {
    targetRange = String(targetRange).replace(/[$]/g, "").replace(/^=|^["']|["']$/g, "").trim();
    if (targetRange.includes("!")) {
      targetRange = targetRange.split("!").pop().trim();
    }
  }

  // Fill Color
  let fillColor = "";
  if (spec.fill && typeof spec.fill === "object") {
    fillColor = spec.fill.color || spec.fill.fillColor || "";
  } else if (typeof spec.fill === "string") {
    fillColor = spec.fill;
  } else if (spec.fillColor) {
    fillColor = spec.fillColor;
  } else if (spec.backgroundColor || spec.background) {
    fillColor = spec.backgroundColor || spec.background;
  }
  fillColor = normalizeColor(fillColor);

  // Font
  let fontObj = null;
  const rawFont = spec.font || {};
  const hasFont = spec.font || spec.fontBold !== undefined || spec.bold !== undefined || spec.italic !== undefined ||
    spec.fontColor || spec.textColor || spec.fontSize || spec.fontName || spec.underline !== undefined || spec.strikethrough !== undefined;

  if (hasFont) {
    fontObj = {
      bold: rawFont.bold !== undefined ? Boolean(rawFont.bold) : (spec.bold !== undefined ? Boolean(spec.bold) : (spec.fontBold !== undefined ? Boolean(spec.fontBold) : undefined)),
      italic: rawFont.italic !== undefined ? Boolean(rawFont.italic) : (spec.italic !== undefined ? Boolean(spec.italic) : undefined),
      underline: rawFont.underline !== undefined ? (rawFont.underline === true ? "Single" : (rawFont.underline === false ? "None" : rawFont.underline)) : (spec.underline !== undefined ? (spec.underline === true ? "Single" : (spec.underline === false ? "None" : spec.underline)) : undefined),
      strikethrough: rawFont.strikethrough !== undefined ? Boolean(rawFont.strikethrough) : (spec.strikethrough !== undefined ? Boolean(spec.strikethrough) : undefined),
      color: normalizeColor(rawFont.color || rawFont.fontColor || spec.fontColor || spec.textColor || ""),
      size: typeof rawFont.size === "number" ? rawFont.size : (typeof spec.fontSize === "number" ? spec.fontSize : (typeof spec.size === "number" ? spec.size : undefined)),
      name: rawFont.name || rawFont.fontName || rawFont.fontFamily || spec.fontName || spec.fontFamily || undefined
    };
    if (fontObj.bold === undefined && fontObj.italic === undefined && fontObj.underline === undefined &&
        fontObj.strikethrough === undefined && !fontObj.color && fontObj.size === undefined && !fontObj.name) {
      fontObj = null;
    }
  }

  // Alignment
  let horizontalAlignment = spec.horizontalAlignment || spec.horizontal || spec.align || (spec.alignment && spec.alignment.horizontal) || undefined;
  if (horizontalAlignment) {
    const h = String(horizontalAlignment).trim().toLowerCase();
    if (h === "center") horizontalAlignment = "Center";
    else if (h === "left") horizontalAlignment = "Left";
    else if (h === "right") horizontalAlignment = "Right";
    else if (h === "justify") horizontalAlignment = "Justify";
    else if (h === "general") horizontalAlignment = "General";
  }

  let verticalAlignment = spec.verticalAlignment || spec.vertical || spec.valign || (spec.alignment && spec.alignment.vertical) || undefined;
  if (verticalAlignment) {
    const v = String(verticalAlignment).trim().toLowerCase();
    if (v === "center" || v === "middle") verticalAlignment = "Center";
    else if (v === "top") verticalAlignment = "Top";
    else if (v === "bottom") verticalAlignment = "Bottom";
    else if (v === "justify") verticalAlignment = "Justify";
  }

  const wrapText = spec.wrapText !== undefined ? Boolean(spec.wrapText) : (spec.alignment && spec.alignment.wrapText !== undefined ? Boolean(spec.alignment.wrapText) : undefined);

  // Number Format
  let numberFormat = spec.numberFormat || spec.formatString || spec.numFormat || undefined;
  if (numberFormat) {
    numberFormat = normalizeNumberFormat(numberFormat);
  }

  // Borders
  let borders = null;
  if (spec.borders) {
    if (spec.borders === true || spec.borders === "all" || spec.borders === "box") {
      borders = { edge: spec.borders === "box" ? "Outline" : "All", style: "Continuous", weight: "Thin", color: "#D3D3D3" };
    } else if (typeof spec.borders === "object") {
      borders = {
        edge: spec.borders.edge || "All",
        style: spec.borders.style || "Continuous",
        weight: spec.borders.weight || "Thin",
        color: normalizeColor(spec.borders.color || "#D3D3D3")
      };
    }
  }

  // Auto-fit & Dimensions
  const autofitColumns = spec.autofitColumns === true || spec.autofit === true;
  const autofitRows = spec.autofitRows === true;
  const rowHeight = typeof spec.rowHeight === "number" ? spec.rowHeight : undefined;
  const columnWidth = typeof spec.columnWidth === "number" ? spec.columnWidth : undefined;

  // Clear options
  const clear = spec.clear ? String(spec.clear).toLowerCase() : undefined;

  // Values / formulas
  const values = spec.values !== undefined ? spec.values : undefined;
  const formulas = spec.formulas !== undefined ? spec.formulas : (spec.formula ? spec.formula : undefined);

  // If nothing was specified, return null
  if (!fillColor && !fontObj && !horizontalAlignment && !verticalAlignment && wrapText === undefined &&
      !numberFormat && !borders && !autofitColumns && !autofitRows && rowHeight === undefined &&
      columnWidth === undefined && !clear && values === undefined && formulas === undefined) {
    return null;
  }

  return {
    action: "format-range",
    targetRange,
    fill: fillColor ? { color: fillColor } : null,
    font: fontObj,
    horizontalAlignment,
    verticalAlignment,
    wrapText,
    numberFormat,
    borders,
    autofitColumns,
    autofitRows,
    rowHeight,
    columnWidth,
    clear,
    values,
    formulas
  };
}

/**
 * Robustly extracts an Excel range format / modification specification from AI response text or HTML
 * Handles ```excel-range-format, ```range-format, ```excel-format, JSON blocks, and JavaScript Office.js statements.
 * @param {string} text - Raw AI response text
 * @param {string} htmlContent - Optional parsed HTML content
 * @returns {Object|null} - Normalized range format specification or null
 */
export function extractExcelRangeFormatSpec(text, htmlContent = "") {
  const combined = (text || "") + "\n" + (htmlContent ? htmlContent.replace(/<[^>]+>/g, " ") : "");
  if (!combined.trim()) return null;

  // 1. Check for dedicated fenced code blocks: ```excel-range-format, ```range-format, ```excel-format, ```range-modify, ```excel-range
  const rangeBlockRegex = /```(?:excel-range-format|range-format|excel-format|excel-range-modify|range-modify|excel-range|format)\s*([\s\S]*?)```/gi;
  let blockMatch;
  while ((blockMatch = rangeBlockRegex.exec(combined)) !== null) {
    if (blockMatch[1]) {
      try {
        const cleaned = blockMatch[1].trim();
        if (cleaned.startsWith("{") && cleaned.endsWith("}")) {
          const parsed = JSON.parse(cleaned);
          const normalized = normalizeRangeFormatSpec(parsed);
          if (normalized) return normalized;
        }
      } catch (e) {
        // Continue searching
      }
    }
  }

  // 2. Check for general JSON code blocks or raw JSON containing range formatting properties
  const jsonBlockRegex = /```(?:json)?\s*([\s\S]*?)```/gi;
  while ((blockMatch = jsonBlockRegex.exec(combined)) !== null) {
    if (blockMatch[1]) {
      try {
        const cleaned = blockMatch[1].trim();
        if (cleaned.startsWith("{") && cleaned.endsWith("}")) {
          const parsed = JSON.parse(cleaned);
          if (parsed && (parsed.fill || parsed.fillColor || parsed.backgroundColor || parsed.horizontalAlignment || parsed.verticalAlignment || parsed.numberFormat || parsed.autofitColumns || parsed.action === "format-range" || parsed.action === "modify-range")) {
            const normalized = normalizeRangeFormatSpec(parsed);
            if (normalized) return normalized;
          }
        }
      } catch (e) {
        // Continue searching
      }
    }
  }

  // 3. Check for standalone JSON objects in the text
  const jsonObjectRegex = /\{[\s\S]*?\}/g;
  let objMatch;
  while ((objMatch = jsonObjectRegex.exec(combined)) !== null) {
    try {
      const parsed = JSON.parse(objMatch[0].trim());
      if (parsed && (parsed.fill || parsed.fillColor || parsed.backgroundColor || parsed.horizontalAlignment || parsed.verticalAlignment || parsed.numberFormat || parsed.autofitColumns || parsed.action === "format-range" || parsed.action === "modify-range")) {
        const normalized = normalizeRangeFormatSpec(parsed);
        if (normalized) return normalized;
      }
    } catch (e) {
      // Continue
    }
  }

  // 4. Fallback: Parse JavaScript / Office.js Code Statements or Pseudo-code
  // Examples:
  // =targetRange.getFormat().getFill().setColor("lightblue")
  // targetRange.format.fill.color = "lightblue";
  // targetRange.format.font.bold = true;
  // targetRange.format.autofitColumns();
  // range.numberFormat = "$#,##0.00";
  const jsStatements = combined.split(/\r?\n/);
  const synthesizedSpec = {};
  let foundAnyJsFormat = false;

  for (const rawLine of jsStatements) {
    const line = rawLine.trim().replace(/^=/, "").trim(); // Strip optional leading '='
    if (!line) continue;

    // Fill Color: .getFill().setColor("...") or .fill.color = "..." or .fillColor = "..."
    const fillMatch = line.match(/(?:\.getFill\(\)\.setColor|\.fill\.color|\.fillColor|\.backgroundColor|\.background)\s*(?:\(|\=)\s*["']([^"']+)["']/i);
    if (fillMatch) {
      synthesizedSpec.fillColor = fillMatch[1];
      foundAnyJsFormat = true;
    }

    // Font Bold: .getFont().setBold(true) or .font.bold = true
    const boldMatch = line.match(/(?:\.getFont\(\)\.setBold|\.font\.bold|\.bold)\s*(?:\(|\=)\s*(true|false)/i);
    if (boldMatch) {
      synthesizedSpec.bold = boldMatch[1].toLowerCase() === "true";
      foundAnyJsFormat = true;
    }

    // Font Italic: .getFont().setItalic(true) or .font.italic = true
    const italicMatch = line.match(/(?:\.getFont\(\)\.setItalic|\.font\.italic|\.italic)\s*(?:\(|\=)\s*(true|false)/i);
    if (italicMatch) {
      synthesizedSpec.italic = italicMatch[1].toLowerCase() === "true";
      foundAnyJsFormat = true;
    }

    // Font Color: .getFont().setColor("...") or .font.color = "..."
    const fontColorMatch = line.match(/(?:\.getFont\(\)\.setColor|\.font\.color|\.fontColor|\.textColor)\s*(?:\(|\=)\s*["']([^"']+)["']/i);
    if (fontColorMatch) {
      synthesizedSpec.fontColor = fontColorMatch[1];
      foundAnyJsFormat = true;
    }

    // Font Size: .getFont().setSize(12) or .font.size = 12
    const fontSizeMatch = line.match(/(?:\.getFont\(\)\.setSize|\.font\.size|\.fontSize)\s*(?:\(|\=)\s*([0-9]+)/i);
    if (fontSizeMatch) {
      synthesizedSpec.fontSize = parseInt(fontSizeMatch[1], 10);
      foundAnyJsFormat = true;
    }

    // Font Name: .getFont().setName("...") or .font.name = "..."
    const fontNameMatch = line.match(/(?:\.getFont\(\)\.setName|\.font\.name|\.fontName|\.fontFamily)\s*(?:\(|\=)\s*["']([^"']+)["']/i);
    if (fontNameMatch) {
      synthesizedSpec.fontName = fontNameMatch[1];
      foundAnyJsFormat = true;
    }

    // Alignment: .setHorizontalAlignment("...") or .horizontalAlignment = "..."
    const hAlignMatch = line.match(/(?:\.setHorizontalAlignment|\.horizontalAlignment|\.horizontal|\.align)\s*(?:\(|\=)\s*["']([^"']+)["']/i);
    if (hAlignMatch) {
      synthesizedSpec.horizontalAlignment = hAlignMatch[1];
      foundAnyJsFormat = true;
    }

    // Vertical Alignment: .setVerticalAlignment("...") or .verticalAlignment = "..."
    const vAlignMatch = line.match(/(?:\.setVerticalAlignment|\.verticalAlignment|\.vertical|\.valign)\s*(?:\(|\=)\s*["']([^"']+)["']/i);
    if (vAlignMatch) {
      synthesizedSpec.verticalAlignment = vAlignMatch[1];
      foundAnyJsFormat = true;
    }

    // Wrap Text: .setWrapText(true) or .wrapText = true
    const wrapMatch = line.match(/(?:\.setWrapText|\.wrapText)\s*(?:\(|\=)\s*(true|false)/i);
    if (wrapMatch) {
      synthesizedSpec.wrapText = wrapMatch[1].toLowerCase() === "true";
      foundAnyJsFormat = true;
    }

    // Number Format: .setNumberFormat("...") or .numberFormat = "..."
    const numFmtMatch = line.match(/(?:\.setNumberFormat|\.numberFormat)\s*(?:\(|\=)\s*["']([^"']+)["']/i);
    if (numFmtMatch) {
      synthesizedSpec.numberFormat = numFmtMatch[1];
      foundAnyJsFormat = true;
    }

    // Autofit Columns / Rows
    if (/\.autofitColumns\(\)/i.test(line)) {
      synthesizedSpec.autofitColumns = true;
      foundAnyJsFormat = true;
    }
    if (/\.autofitRows\(\)/i.test(line)) {
      synthesizedSpec.autofitRows = true;
      foundAnyJsFormat = true;
    }

    // Target Range: sheet.getRange("A1:D10") or getRange("A1:D10")
    const rangeMatch = line.match(/(?:getRange)\s*\(\s*["']([A-Za-z0-9:!$]+)["']\s*\)/i);
    if (rangeMatch && !synthesizedSpec.targetRange) {
      synthesizedSpec.targetRange = rangeMatch[1];
    }
  }

  if (foundAnyJsFormat) {
    const normalized = normalizeRangeFormatSpec(synthesizedSpec);
    if (normalized) return normalized;
  }

  return null;
}

/**
 * Applies formatting and styling modifications to a target range on the active worksheet
 * 
 * @param {Excel.RequestContext} context - Active Office.js request context
 * @param {Excel.Worksheet} sheet - Target worksheet
 * @param {Object} spec - Normalized range format specification
 * @returns {Promise<Excel.Range>}
 */
export async function applyRangeFormatting(context, sheet, spec) {
  if (!spec) {
    throw new Error("Invalid range formatting specification: specification is empty.");
  }

  let targetRange = null;

  // 1. Locate Target Range by Address if provided
  if (spec.targetRange && spec.targetRange.includes(":") && /^[A-Za-z]+[0-9]+:[A-Za-z]+[0-9]+$/.test(spec.targetRange)) {
    try {
      targetRange = sheet.getRange(spec.targetRange);
    } catch (e) {
      targetRange = null;
    }
  } else if (spec.targetRange && /^[A-Za-z]+[0-9]+$/.test(spec.targetRange)) {
    try {
      targetRange = sheet.getRange(spec.targetRange);
    } catch (e) {
      targetRange = null;
    }
  }

  // 2. Fallback to active selected range
  if (!targetRange) {
    targetRange = context.workbook.getSelectedRange();
  }

  targetRange.load(["address", "rowCount", "columnCount"]);
  await context.sync();

  console.log(`[Range Formatting] Applying format to range: ${targetRange.address}`);

  // 3. Clear formatting or contents if requested
  if (spec.clear) {
    try {
      if (spec.clear === "formats") {
        targetRange.clear("Formats");
      } else if (spec.clear === "contents") {
        targetRange.clear("Contents");
      } else if (spec.clear === "all") {
        targetRange.clear("All");
      }
    } catch (cErr) {
      console.warn("Clear range notice:", cErr);
    }
  }

  // 4. Apply Fill / Background Color
  if (spec.fill && spec.fill.color !== undefined) {
    try {
      if (spec.fill.color) {
        targetRange.format.fill.color = spec.fill.color;
      } else {
        targetRange.format.fill.clear();
      }
    } catch (fillErr) {
      console.warn("Fill color application notice:", fillErr);
    }
  }

  // 5. Apply Font Styling
  if (spec.font) {
    try {
      if (spec.font.bold !== undefined) {
        targetRange.format.font.bold = Boolean(spec.font.bold);
      }
      if (spec.font.italic !== undefined) {
        targetRange.format.font.italic = Boolean(spec.font.italic);
      }
      if (spec.font.underline !== undefined) {
        targetRange.format.font.underline = spec.font.underline;
      }
      if (spec.font.strikethrough !== undefined) {
        targetRange.format.font.strikethrough = Boolean(spec.font.strikethrough);
      }
      if (spec.font.color) {
        targetRange.format.font.color = spec.font.color;
      }
      if (typeof spec.font.size === "number" && spec.font.size >= 6 && spec.font.size <= 72) {
        targetRange.format.font.size = spec.font.size;
      }
      if (spec.font.name) {
        targetRange.format.font.name = spec.font.name;
      }
    } catch (fontErr) {
      console.warn("Font formatting application notice:", fontErr);
    }
  }

  // 6. Apply Text Alignment
  if (spec.horizontalAlignment) {
    try {
      targetRange.format.horizontalAlignment = spec.horizontalAlignment;
    } catch (alignErr) {
      console.warn("Horizontal alignment notice:", alignErr);
    }
  }
  if (spec.verticalAlignment) {
    try {
      targetRange.format.verticalAlignment = spec.verticalAlignment;
    } catch (vAlignErr) {
      console.warn("Vertical alignment notice:", vAlignErr);
    }
  }
  if (spec.wrapText !== undefined) {
    try {
      targetRange.format.wrapText = Boolean(spec.wrapText);
    } catch (wrapErr) {
      console.warn("Wrap text notice:", wrapErr);
    }
  }

  // 7. Apply Borders
  if (spec.borders) {
    const borderEdges = [];
    const edge = (spec.borders.edge || "All").toLowerCase();
    if (edge === "all") {
      borderEdges.push("EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideHorizontal", "InsideVertical");
    } else if (edge === "outline" || edge === "box") {
      borderEdges.push("EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight");
    } else if (edge === "top") {
      borderEdges.push("EdgeTop");
    } else if (edge === "bottom") {
      borderEdges.push("EdgeBottom");
    } else if (edge === "left") {
      borderEdges.push("EdgeLeft");
    } else if (edge === "right") {
      borderEdges.push("EdgeRight");
    } else if (edge === "inside") {
      borderEdges.push("InsideHorizontal", "InsideVertical");
    }

    const bStyle = spec.borders.style || "Continuous";
    const bColor = spec.borders.color || "#D3D3D3";
    const bWeight = spec.borders.weight || "Thin";

    for (const bEdge of borderEdges) {
      try {
        const borderItem = targetRange.format.borders.getItem(bEdge);
        borderItem.style = bStyle;
        borderItem.color = bColor;
        borderItem.weight = bWeight;
      } catch (bErr) {
        console.debug(`Border edge ${bEdge} notice:`, bErr);
      }
    }
  }

  // 8. Apply Number Formatting
  if (spec.numberFormat) {
    try {
      const numFmtMatrix = Array.from({ length: targetRange.rowCount || 1 }, () => Array(targetRange.columnCount || 1).fill(spec.numberFormat));
      targetRange.numberFormat = numFmtMatrix;
    } catch (numErr) {
      try {
        targetRange.numberFormat = [[spec.numberFormat]];
      } catch (numErr2) {
        console.warn("Number format application notice:", numErr2);
      }
    }
  }

  // 9. Apply Dimensions / Auto-fit
  if (typeof spec.rowHeight === "number") {
    try {
      targetRange.format.rowHeight = spec.rowHeight;
    } catch (rhErr) {
      console.warn("Row height notice:", rhErr);
    }
  }
  if (typeof spec.columnWidth === "number") {
    try {
      targetRange.format.columnWidth = spec.columnWidth;
    } catch (cwErr) {
      console.warn("Column width notice:", cwErr);
    }
  }
  if (spec.autofitColumns) {
    try {
      targetRange.format.autofitColumns();
    } catch (afcErr) {
      console.warn("Autofit columns notice:", afcErr);
    }
  }
  if (spec.autofitRows) {
    try {
      targetRange.format.autofitRows();
    } catch (afrErr) {
      console.warn("Autofit rows notice:", afrErr);
    }
  }

  // 10. Apply Values or Formulas if specified
  if (spec.formulas) {
    try {
      if (Array.isArray(spec.formulas)) {
        targetRange.formulas = spec.formulas;
      } else if (typeof spec.formulas === "string") {
        const norm = normalizeFormula(spec.formulas);
        const fMatrix = Array.from({ length: targetRange.rowCount || 1 }, () => Array(targetRange.columnCount || 1).fill(norm));
        targetRange.formulas = fMatrix;
      }
    } catch (fErr) {
      console.warn("Range formulas application notice:", fErr);
    }
  } else if (spec.values !== undefined) {
    try {
      if (Array.isArray(spec.values)) {
        targetRange.values = spec.values;
      } else {
        const vMatrix = Array.from({ length: targetRange.rowCount || 1 }, () => Array(targetRange.columnCount || 1).fill(spec.values));
        targetRange.values = vMatrix;
      }
    } catch (vErr) {
      console.warn("Range values application notice:", vErr);
    }
  }

  await context.sync();
  return targetRange;
}

/**
 * Normalizes a taxonomy categorization specification object into a clean structured schema
 * @param {Object} spec 
 * @returns {Object|null}
 */
export function normalizeTaxonomySpec(spec) {
  if (!spec || typeof spec !== "object") return null;

  let sourceColRaw = String(spec.sourceColumn || spec.sourceCol || spec.sourceRange || spec.source || "").trim();
  if (sourceColRaw.includes(":")) sourceColRaw = sourceColRaw.split(":")[0];

  let sourceCol = "B";
  let sourceColumnName = null;

  // Check if source is a clean column letter like "B" or "AB"
  const cleanColMatch = sourceColRaw.match(/^([A-Za-z]{1,3})([0-9]*)$/);
  if (cleanColMatch && letterToColumnIndex(cleanColMatch[1]) < 1000) {
    sourceCol = cleanColMatch[1].toUpperCase();
  } else if (sourceColRaw) {
    sourceColumnName = sourceColRaw;
    sourceCol = "B";
  }

  let targetColRaw = String(spec.targetColumn || spec.targetCol || spec.targetRange || spec.target || "").trim();
  if (targetColRaw.includes(":")) targetColRaw = targetColRaw.split(":")[0];

  let targetCol = "";
  const cleanTargetMatch = targetColRaw.match(/^([A-Za-z]{1,3})([0-9]*)$/);
  if (cleanTargetMatch && letterToColumnIndex(cleanTargetMatch[1]) < 1000) {
    targetCol = cleanTargetMatch[1].toUpperCase();
  } else {
    const sIdx = letterToColumnIndex(sourceCol);
    targetCol = columnIndexToLetter(sIdx + 1);
  }

  let header = spec.header || spec.headerTitle || spec.categoryHeader || spec.columnName || spec.title || "Category";
  const insertColumn = spec.insertColumn !== false; // default: true
  const startRow = typeof spec.startRow === "number" && spec.startRow >= 1 ? spec.startRow : 2;

  let categories = [];
  if (Array.isArray(spec.categories)) {
    categories = spec.categories;
  } else if (Array.isArray(spec.taxonomy)) {
    // If it's an array of column definitions e.g. [ { columnName: "Sales Rep", mapping: [...] } ]
    const firstCol = spec.taxonomy[0];
    if (firstCol && Array.isArray(firstCol.mapping)) {
      if (firstCol.columnName && !sourceColumnName) {
        sourceColumnName = firstCol.columnName;
      }
      categories = firstCol.mapping.map((m, idx) => ({
        row: startRow + idx,
        sourceValue: m.value || m.sourceValue || m.from || "",
        category: m.category || m.targetCategory || m.to || ""
      }));
      if (!spec.header && firstCol.columnName) {
        header = `${firstCol.columnName} Category`;
      }
    } else if (firstCol && (firstCol.category || firstCol.sourceValue)) {
      categories = spec.taxonomy;
    }
  } else if (spec.taxonomy && typeof spec.taxonomy === "object" && !Array.isArray(spec.taxonomy)) {
    // Key-value object mapping
    categories = Object.entries(spec.taxonomy).map(([k, v], idx) => ({
      row: startRow + idx,
      sourceValue: k,
      category: typeof v === "object" ? (v.category || v.value || JSON.stringify(v)) : String(v)
    }));
  } else if (Array.isArray(spec.mapping)) {
    categories = spec.mapping.map((m, idx) => ({
      row: startRow + idx,
      sourceValue: m.value || m.sourceValue || m.from || "",
      category: m.category || m.targetCategory || m.to || ""
    }));
  } else if (Array.isArray(spec.mappings)) {
    categories = spec.mappings;
  } else if (Array.isArray(spec.values)) {
    categories = spec.values;
  } else if (Array.isArray(spec.data)) {
    categories = spec.data;
  } else if (spec.categories && typeof spec.categories === "object") {
    // Key-value object mapping { "Uber": "Travel", "AWS": "Cloud" }
    categories = Object.entries(spec.categories).map(([k, v]) => ({ sourceValue: k, category: v }));
  }

  if (!categories || categories.length === 0) return null;

  // Compute clean list of category strings
  const categoryValues = categories.map(item => {
    if (typeof item === "string") return item;
    if (item && typeof item === "object") {
      return item.category || item.categoryName || item.taxonomy || item.type || item.value || item.to || "";
    }
    return String(item || "");
  });

  // Extract or synthesize taxonomy summary
  let taxonomySummary = spec.taxonomySummary || spec.summary || null;
  if (!taxonomySummary && categoryValues.length > 0) {
    taxonomySummary = {};
    categoryValues.forEach(cat => {
      const c = String(cat || "Uncategorized").trim();
      if (c) taxonomySummary[c] = (taxonomySummary[c] || 0) + 1;
    });
  }

  return {
    action: "categorize",
    sourceColumn: sourceCol,
    sourceColumnName,
    targetColumn: targetCol,
    insertColumn,
    header,
    startRow,
    endRow: typeof spec.endRow === "number" && spec.endRow >= startRow ? spec.endRow : (startRow + categories.length - 1),
    categories,
    categoryValues,
    count: categories.length,
    taxonomySummary
  };
}

/**
 * Robustly extracts an Excel taxonomy categorization specification from AI response text or HTML
 * Handles ```excel-taxonomy, ```excel-categorize, ```taxonomy, ```categorize code blocks, JSON objects, and markdown mapping tables.
 * @param {string} text - Raw AI response text
 * @param {string} htmlContent - Optional parsed HTML content
 * @returns {Object|null} - Normalized taxonomy specification or null
 */
export function extractExcelTaxonomySpec(text, htmlContent = "", options = {}) {
  const combined = (text || "") + "\n" + (htmlContent ? htmlContent.replace(/<[^>]+>/g, " ") : "");
  if (!combined.trim()) return null;

  const userPrompt = (options && options.userPrompt) || "";
  const prompt = (options && options.prompt) || "";

  // Guard 1: If user is asking for chart updates or chart operations, NEVER treat as taxonomy
  if (/update.*chart|change.*chart|modify.*chart|switch.*chart|edit.*chart|re-style.*chart|chart/i.test(userPrompt) ||
      /update.*chart|change.*chart|modify.*chart|switch.*chart|edit.*chart/i.test(prompt) ||
      /```(?:excel-chart|chart|excel_chart|excel-chart-edit|chart-edit)/i.test(combined) ||
      /"chartType"|"targetChart"|"categoriesRange"|"seriesRange"|"dataRange"/i.test(combined)) {
    return null;
  }

  // Guard 2: If this is an Insights summary response, do not treat it as a taxonomy spec
  if (/structured insights|insights summary|key table patterns|anomalies.*outliers/i.test(combined) ||
      /insights/i.test(userPrompt) || /insights/i.test(prompt)) {
    return null;
  }

  const hasExplicitTaxonomyBlock = /```(?:excel-taxonomy|excel-categorize|taxonomy|categorize|excel-classify|classify|excel-column-insert)/i.test(combined);
  const isExplicitTaxonomyRequest = Boolean(
    (options && options.isTaxonomy) ||
    /taxonomy|categoriz|categories\s+column|expense\s+categor/i.test(userPrompt) ||
    /taxonomy|categoriz|categories\s+column|expense\s+categor/i.test(prompt) ||
    /taxonomy|categoriz|categories\s+column|expense\s+categor/i.test(combined) ||
    hasExplicitTaxonomyBlock
  );

  const isExplicitTaxonomyJson = (parsed) => {
    if (!parsed || typeof parsed !== 'object') return false;
    if (parsed.action === "categorize" || parsed.action === "taxonomy" || parsed.taxonomySummary) return true;
    if (parsed.insertColumn && parsed.header && Array.isArray(parsed.categories)) return true;
    if (Array.isArray(parsed.categories) && parsed.categories.length > 0 && typeof parsed.categories[0] === 'object' && (parsed.categories[0].category || parsed.categories[0].categoryName || parsed.categories[0].sourceValue)) return true;
    if (Array.isArray(parsed.taxonomy)) return true;
    if (parsed.taxonomy && typeof parsed.taxonomy === 'object') return true;
    if (Array.isArray(parsed.mapping)) return true;
    return false;
  };

  // 1. Check for dedicated fenced code blocks: ```excel-taxonomy, ```taxonomy, ```excel-categorize, ```categorize, ```excel-classify
  const taxonomyBlockRegex = /```(?:excel-taxonomy|excel-categorize|taxonomy|categorize|excel-classify|classify|excel-column-insert)\s*([\s\S]*?)```/gi;
  let blockMatch;
  while ((blockMatch = taxonomyBlockRegex.exec(combined)) !== null) {
    if (blockMatch[1]) {
      try {
        const cleaned = blockMatch[1].trim();
        if (cleaned.startsWith("{") && cleaned.endsWith("}")) {
          const parsed = JSON.parse(cleaned);
          const normalized = normalizeTaxonomySpec(parsed);
          if (normalized) return normalized;
        }
      } catch (e) {
        // Continue searching
      }
    }
  }

  // 2. Check for general JSON code blocks containing explicit taxonomy properties
  const jsonBlockRegex = /```(?:json)?\s*([\s\S]*?)```/gi;
  while ((blockMatch = jsonBlockRegex.exec(combined)) !== null) {
    if (blockMatch[1]) {
      try {
        const cleaned = blockMatch[1].trim();
        if (cleaned.startsWith("{") && cleaned.endsWith("}")) {
          const parsed = JSON.parse(cleaned);
          if (isExplicitTaxonomyJson(parsed)) {
            const normalized = normalizeTaxonomySpec(parsed);
            if (normalized) return normalized;
          }
        }
      } catch (e) {
        // Continue searching
      }
    }
  }

  // 3. Check for standalone JSON objects in the text (only when explicit taxonomy request)
  if (isExplicitTaxonomyRequest) {
    const jsonObjectRegex = /\{[\s\S]*?\}/g;
    let objMatch;
    while ((objMatch = jsonObjectRegex.exec(combined)) !== null) {
      try {
        const parsed = JSON.parse(objMatch[0].trim());
        if (isExplicitTaxonomyJson(parsed)) {
          const normalized = normalizeTaxonomySpec(parsed);
          if (normalized) return normalized;
        }
      } catch (e) {
        // Continue
      }
    }
  }

  // 4. Fallback: Parse markdown table with categorization columns ONLY if explicit taxonomy intent was requested
  if (isExplicitTaxonomyRequest && combined.includes("|")) {
    const lines = combined.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
    const tableLines = lines.filter(l => l.startsWith("|") && l.endsWith("|"));
    if (tableLines.length >= 3) {
      const headerLine = tableLines[0];
      const headers = headerLine.split("|").map(h => h.trim()).filter(Boolean);
      const catIdx = headers.findIndex(h => /category|taxonomy|classification|type|mapped/i.test(h) && !/total|amount|sum|count|subtotal|avg|metric/i.test(h));
      const srcIdx = headers.findIndex((h, idx) => idx !== catIdx && /source|merchant|vendor|raw|item|text|original|entry|description|name|product/i.test(h));

      if (catIdx !== -1 && (srcIdx !== -1 || headers.length >= 2)) {
        const categories = [];
        for (let i = 2; i < tableLines.length; i++) {
          const cells = tableLines[i].split("|").map(c => c.trim()).filter(Boolean);
          if (cells.length > catIdx) {
            const catVal = cells[catIdx];
            const srcVal = srcIdx !== -1 && cells.length > srcIdx ? cells[srcIdx] : `Row ${i}`;
            if (catVal && !/^[-:]+$/.test(catVal)) {
              categories.push({
                row: i,
                sourceValue: srcVal,
                category: catVal
              });
            }
          }
        }
        if (categories.length >= 1) {
          return normalizeTaxonomySpec({
            action: "categorize",
            sourceColumn: "B",
            targetColumn: "C",
            insertColumn: true,
            header: headers[catIdx] || "Category",
            startRow: 2,
            categories
          });
        }
      }
    }
  }

  return null;
}

/**
 * Extracts Structured Insights Summary Table from text/HTML response
 * @param {string} text - AI response text
 * @param {string} htmlContent - Optional parsed HTML
 * @returns {Array<Array<string>>|null}
 */
export function extractInsightsSummaryTable(text, htmlContent = "") {
  return ExcelInsightsService.extractInsightsSummaryTable(text, htmlContent);
}

/**
 * Normalizes table matrix converting currency, percentages, numbers to JS numbers and formatting
 * @param {Array<Array<string>>} rawMatrix 
 */
export function normalizeInsightsTableData(rawMatrix) {
  return ExcelInsightsService.normalizeInsightsTableData(rawMatrix);
}

/**
 * Executes taxonomy categorization on the active worksheet:
 * 1. Resolves source & target columns.
 * 2. Inserts a new adjacent column (e.g. Column C).
 * 3. Sets the category header with corporate Office styling in Row 1.
 * 4. Populates all mapped taxonomy values for each corresponding row in a single batch matrix.
 * 5. Applies autofit and formatting.
 * 
 * @param {Excel.RequestContext} context - Active Office.js request context
 * @param {Excel.Worksheet} sheet - Target worksheet
 * @param {Object} spec - Normalized taxonomy specification
 * @returns {Promise<Object>} Categorization result summary
 */
export async function applyTaxonomyCategorization(context, sheet, spec) {
  if (!spec) {
    throw new Error("Invalid taxonomy specification: specification is empty.");
  }

  const normalized = normalizeTaxonomySpec(spec);
  if (!normalized || !normalized.categoryValues || normalized.categoryValues.length === 0) {
    throw new Error("No category values found in taxonomy specification.");
  }

  // 1. Inspect worksheet used range to dynamically resolve column letters, headers, and row bounds
  let usedRange = null;
  let headers = [];
  let sheetStartRow = 1;
  let sheetEndRow = null;
  let sheetStartColIdx = 0;
  try {
    if (sheet && typeof sheet.getUsedRange === 'function') {
      usedRange = sheet.getUsedRange();
      if (usedRange && typeof usedRange.load === 'function') {
        usedRange.load(["address", "rowCount", "columnCount", "rowIndex", "columnIndex", "values"]);
        await context.sync();
        if (usedRange) {
          sheetStartRow = (usedRange.rowIndex || 0) + 1;
          sheetEndRow = (usedRange.rowIndex || 0) + (usedRange.rowCount || 1);
          sheetStartColIdx = usedRange.columnIndex || 0;
          if (usedRange.values && usedRange.values.length > 0) {
            headers = usedRange.values[0] || [];
          }
        }
      }
    }
  } catch (uErr) {
    console.debug("[Excel Taxonomy] Used range inspection notice:", uErr);
  }

  let sourceCol = normalized.sourceColumn || "B";
  let targetCol = normalized.targetColumn || "C";
  let header = normalized.header || "Category";
  let startRow = normalized.startRow || 2;
  let endRow = normalized.endRow || (startRow + normalized.categoryValues.length - 1);

  // If sourceColumn was specified as a column header name (e.g. "Sales Rep", "Opportunity Name")
  // resolve against actual sheet headers
  const sourceNameCandidate = normalized.sourceColumnName;
  if (sourceNameCandidate && headers && headers.length > 0) {
    const cleanCand = String(sourceNameCandidate).toLowerCase().trim();
    const hIdx = headers.findIndex(h => {
      const cleanH = String(h || "").toLowerCase().trim();
      return cleanH === cleanCand || cleanH.includes(cleanCand) || cleanCand.includes(cleanH);
    });
    if (hIdx !== -1) {
      sourceCol = columnIndexToLetter(sheetStartColIdx + hIdx);
      targetCol = columnIndexToLetter(sheetStartColIdx + hIdx + 1);
      console.log(`[Excel Taxonomy] Resolved header "${sourceNameCandidate}" ➔ Column ${sourceCol}, target Column ${targetCol}`);
    }
  }

  // Validate column letters
  if (!/^[A-Za-z]{1,3}$/.test(sourceCol) || letterToColumnIndex(sourceCol) > 16383) {
    sourceCol = "B";
  }
  if (!/^[A-Za-z]{1,3}$/.test(targetCol) || letterToColumnIndex(targetCol) > 16383) {
    targetCol = columnIndexToLetter(letterToColumnIndex(sourceCol) + 1);
  }

  // If taxonomy contains mapping dictionary or fewer rows than the sheet, extend endRow to match full table
  const mappingMap = new Map();
  if (Array.isArray(normalized.categories)) {
    normalized.categories.forEach(item => {
      if (item && item.sourceValue && item.category) {
        mappingMap.set(String(item.sourceValue).toLowerCase().trim(), String(item.category).trim());
      }
    });
  }

  if (mappingMap.size > 0 && typeof sheetEndRow === "number" && sheetEndRow > endRow) {
    endRow = sheetEndRow;
  }

  const headerRow = startRow > 1 ? startRow - 1 : 1;
  const targetRowCount = Math.max(1, endRow - startRow + 1);
  const rowCount = targetRowCount;

  console.log(`[Excel Taxonomy] Categorizing Column ${sourceCol} ➔ Column ${targetCol} (Rows ${startRow}-${endRow}, ${targetRowCount} data rows)`);

  // 2. Read Source Column values BEFORE column insertion so column index shift cannot corrupt reading
  let sourceValues = [];
  try {
    const srcRange = sheet.getRange(`${sourceCol}${startRow}:${sourceCol}${endRow}`);
    srcRange.load("values");
    await context.sync();
    if (srcRange.values && Array.isArray(srcRange.values)) {
      sourceValues = srcRange.values.map(r => (r && r[0] !== undefined ? String(r[0]).trim() : ""));
    }
  } catch (srcErr) {
    console.debug("[Excel Taxonomy] Source value pre-read notice:", srcErr);
  }

  // Compute final categorized values array for all rows
  let finalCategoryValues = [];
  if (normalized.categoryValues && normalized.categoryValues.length >= targetRowCount) {
    // 1. Direct 1-to-1 row categorization returned by the AI model
    finalCategoryValues = normalized.categoryValues.slice(0, targetRowCount);
  } else if (Array.isArray(normalized.categories) && normalized.categories.some(c => typeof c === "object" && typeof c?.row === "number")) {
    // 2. Explicit row-numbered mapping entries (e.g. { row: 2, category: "High Value" })
    finalCategoryValues = Array(targetRowCount).fill("");
    normalized.categories.forEach((item, i) => {
      if (item && typeof item === "object") {
        const cat = item.category || item.categoryName || item.taxonomy || item.type || item.value || item.to || "";
        const rIdx = typeof item.row === "number" ? (item.row - startRow) : i;
        if (rIdx >= 0 && rIdx < targetRowCount) {
          finalCategoryValues[rIdx] = cat;
        }
      }
    });
  } else if (sourceValues.length > 0 && mappingMap.size > 0) {
    // 3. Dictionary / lookup mapping against pre-read source column cells
    finalCategoryValues = sourceValues.map((rawCell, idx) => {
      if (rawCell === undefined || rawCell === null || rawCell === "") {
        return (normalized.categoryValues && normalized.categoryValues[idx]) || "";
      }
      const rawStr = String(rawCell).trim();
      const rawLower = rawStr.toLowerCase();
      // Exact match
      if (mappingMap.has(rawLower)) {
        return mappingMap.get(rawLower);
      }
      // Exact numeric match
      const rawNum = parseFloat(rawStr.replace(/[^0-9.-]+/g, ""));
      if (!isNaN(rawNum) && mappingMap.has(String(rawNum))) {
        return mappingMap.get(String(rawNum));
      }
      // Non-numeric string keyword contains match (minimum 3 chars to prevent false positives)
      if (isNaN(rawNum) && rawStr.length >= 3) {
        for (const [k, v] of mappingMap.entries()) {
          if (k.length >= 3 && isNaN(Number(k))) {
            if (rawLower.includes(k) || k.includes(rawLower)) {
              return v;
            }
          }
        }
      }
      return (normalized.categoryValues && normalized.categoryValues[idx]) || rawStr;
    });
  } else {
    finalCategoryValues = normalized.categoryValues ? [...normalized.categoryValues] : [];
  }

  // Adjust row length to match targetRowCount exactly
  while (finalCategoryValues.length < targetRowCount) {
    finalCategoryValues.push("");
  }
  if (finalCategoryValues.length > targetRowCount) {
    finalCategoryValues = finalCategoryValues.slice(0, targetRowCount);
  }

  // 3. Insert New Column if requested (default: true)
  if (normalized.insertColumn !== false) {
    try {
      const targetColumnRange = sheet.getRange(`${targetCol}:${targetCol}`);
      targetColumnRange.insert(Excel.InsertShiftDirection.right);
      await context.sync();
      console.log(`[Excel Taxonomy] Successfully inserted new column at Column ${targetCol}`);
    } catch (insErr) {
      console.warn("[Excel Taxonomy] Column insertion notice (proceeding with direct write):", insErr);
    }
  }

  // 4. Set Category Header and Corporate Styling
  try {
    const headerCell = sheet.getRange(`${targetCol}${headerRow}`);
    headerCell.values = [[ExcelInsightsService.cleanCellText(header)]];
    headerCell.format.font.bold = true;
    headerCell.format.font.name = "Segoe UI";
    headerCell.format.font.size = 11;
    headerCell.format.font.color = "#004e8c";
    headerCell.format.fill.color = "#EBF3FC";
    headerCell.format.horizontalAlignment = "Center";
    headerCell.format.verticalAlignment = "Center";
  } catch (hErr) {
    console.warn("[Excel Taxonomy] Header formatting notice:", hErr);
  }

  // 5. Write Categorized Values into Target Column in a single batch
  try {
    const targetDataRangeAddress = `${targetCol}${startRow}:${targetCol}${endRow}`;
    const targetDataRange = sheet.getRange(targetDataRangeAddress);
    const matrix = finalCategoryValues.map(cat => [ExcelInsightsService.cleanCellText(cat)]);
    targetDataRange.values = matrix;
    targetDataRange.format.font.name = "Segoe UI";
    targetDataRange.format.font.size = 11;
    targetDataRange.format.verticalAlignment = "Center";
  } catch (vErr) {
    console.warn("[Excel Taxonomy] Target data write notice:", vErr);
  }

  // 6. Auto-fit the newly populated column
  try {
    const fullRange = sheet.getRange(`${targetCol}${headerRow}:${targetCol}${endRow}`);
    fullRange.format.autofitColumns();
  } catch (afErr) {
    console.warn("[Excel Taxonomy] Autofit notice:", afErr);
  }

  await context.sync();

  return {
    sourceColumn: sourceCol,
    targetColumn: targetCol,
    header,
    startRow,
    endRow,
    count: rowCount,
    address: `${targetCol}${headerRow}:${targetCol}${endRow}`
  };
}

export class ExcelAdapter {
  constructor() {
    this.name = "Excel";
    this.insightsService = ExcelInsightsService;
    this.agentMode = "create-formula"; // "create-formula" | "modify-range" | "modify-chart"
    this.lockedChart = null; // Sticky chart metadata { id, name, chartType, title, width, height, top, left }
  }

  setAgentMode(mode) {
    this.agentMode = mode || "create-formula";
  }

  getAgentMode() {
    return this.agentMode || "create-formula";
  }

  setLockedChart(chartMeta) {
    this.lockedChart = chartMeta || null;
  }

  getLockedChart() {
    return this.lockedChart || null;
  }

  /**
   * Helper to inspect the active worksheet schema (used range, column letters, headers, sample data)
   * @param {Excel.RequestContext} context 
   * @param {Excel.Worksheet} sheet 
   * @returns {Promise<Object>} Schema metadata object
   */
  async getWorksheetSchema(context, sheet) {
    const schema = {
      sheetName: "Sheet1",
      usedRangeAddress: "",
      rowCount: 0,
      columnCount: 0,
      startRow: 1,
      endRow: 1,
      startCol: 0,
      columns: [],
      sampleRows: [],
      gridData: [],
      tables: [],
      charts: []
    };

    try {
      sheet.load("name");
      const usedRange = sheet.getUsedRange();
      usedRange.load("address, values, formulas, rowCount, columnCount, rowIndex, columnIndex");
      
      const tables = sheet.tables;
      tables.load("items/name, items/columns/items/name");

      const charts = sheet.charts;
      charts.load("items/id, items/name, items/chartType, items/width, items/height");

      await context.sync();

      schema.sheetName = sheet.name;
      schema.usedRangeAddress = usedRange.address;
      schema.rowCount = usedRange.rowCount;
      schema.columnCount = usedRange.columnCount;
      schema.startRow = (usedRange.rowIndex || 0) + 1;
      schema.endRow = (usedRange.rowIndex || 0) + (usedRange.rowCount || 1);
      schema.startCol = usedRange.columnIndex || 0;

      if (usedRange.values && usedRange.values.length > 0) {
        schema.gridData = usedRange.values;
        const headerRowIndex = 0;
        const headers = usedRange.values[headerRowIndex] || [];
        const startCol = usedRange.columnIndex;

        // Map column letters to header names and sample values
        for (let c = 0; c < usedRange.columnCount; c++) {
          const actualColIndex = startCol + c;
          const colLetter = columnIndexToLetter(actualColIndex);
          const headerName = headers[c] ? String(headers[c]).trim() : `Column ${colLetter}`;
          
          // Grab up to 3 sample values from subsequent rows
          const sampleVals = [];
          for (let r = 1; r < Math.min(usedRange.rowCount, 4); r++) {
            const val = usedRange.values[r] ? usedRange.values[r][c] : "";
            if (val !== undefined && val !== null && String(val).trim().length > 0) {
              sampleVals.push(String(val).trim());
            }
          }

          schema.columns.push({
            letter: colLetter,
            index: actualColIndex,
            header: headerName,
            samples: sampleVals
          });
        }

        // Extract up to 5 sample rows for tabular preview
        const maxSampleRows = Math.min(usedRange.rowCount, 6);
        for (let r = 0; r < maxSampleRows; r++) {
          schema.sampleRows.push(usedRange.values[r] || []);
        }
      }

      // Collect named tables if present
      if (tables.items && tables.items.length > 0) {
        for (const tbl of tables.items) {
          const colNames = tbl.columns && tbl.columns.items ? tbl.columns.items.map(c => c.name) : [];
          schema.tables.push({
            name: tbl.name,
            columns: colNames
          });
        }
      }

      // Collect existing charts on worksheet
      if (charts.items && charts.items.length > 0) {
        for (const ch of charts.items) {
          schema.charts.push({
            id: ch.id,
            name: ch.name,
            type: ch.chartType,
            title: ch.name,
            width: Math.round(ch.width || 0),
            height: Math.round(ch.height || 0)
          });
        }
      }
    } catch (e) {
      console.warn("Error reading worksheet schema:", e);
    }

    return schema;
  }

  /**
   * Safe helper to detect active target (Chart, Table, Range, or Cell).
   * Populates both active/sticky chart and active worksheet range.
   * @param {Excel.RequestContext} context
   * @param {string} [overrideMode]
   * @param {Object} [overrideLockedChart]
   * @returns {Promise<Object>} Target metadata
   */
  async getSelectionTargetInfo(context, overrideMode = null, overrideLockedChart = undefined) {
    const stickyChart = overrideLockedChart !== undefined ? overrideLockedChart : (this.lockedChart || null);

    const targetInfo = {
      type: "cell", // "cell" | "range" | "table"
      address: "A1",
      cleanAddress: "A1",
      sheetName: "Sheet1",
      rowCount: 1,
      columnCount: 1,
      cellCount: 1,
      rawText: "",
      chart: null,
      table: null
    };

    try {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");

      // 1. Check if user currently has a chart actively selected or locked
      let activeChart = null;
      try {
        if (typeof context.workbook.getActiveChartOrNullObject === 'function') {
          activeChart = context.workbook.getActiveChartOrNullObject();
          activeChart.load(["id", "name", "chartType", "width", "height", "top", "left", "isNullObject"]);
        }
      } catch (e) {
        console.debug("Active chart lookup notice:", e);
      }

      await context.sync();
      targetInfo.sheetName = sheet.name || "Sheet1";

      if (activeChart && !activeChart.isNullObject) {
        let chartTitle = activeChart.name || "Chart";
        try {
          if (activeChart.title) {
            activeChart.title.load(["text", "visible"]);
            await context.sync();
            if (activeChart.title.visible && activeChart.title.text) {
              chartTitle = activeChart.title.text;
            }
          }
        } catch (tErr) {}

        const chartMeta = {
          id: activeChart.id,
          name: activeChart.name,
          chartType: activeChart.chartType,
          title: chartTitle,
          width: Math.round(activeChart.width || 0),
          height: Math.round(activeChart.height || 0),
          top: Math.round(activeChart.top || 0),
          left: Math.round(activeChart.left || 0)
        };
        this.lockedChart = chartMeta;
        targetInfo.chart = chartMeta;
      } else if (stickyChart) {
        targetInfo.chart = stickyChart;
        this.lockedChart = stickyChart;
      }

      // 2. Always Detect Active Range / Table / Cell Target
      const selectedRange = context.workbook.getSelectedRange();
      selectedRange.load(["address", "cellCount", "rowCount", "columnCount", "rowIndex", "columnIndex"]);

      let overlappingTables = null;
      try {
        if (typeof selectedRange.getTables === 'function') {
          overlappingTables = selectedRange.getTables(false);
          overlappingTables.load("items/id, items/name");
        }
      } catch (e) {
        console.debug("Table lookup notice:", e);
      }

      await context.sync();

      targetInfo.sheetName = sheet.name || "Sheet1";
      const fullAddress = selectedRange.address || "A1";
      const cleanAddress = fullAddress.includes("!") ? fullAddress.split("!").pop() : fullAddress;
      targetInfo.address = fullAddress;
      targetInfo.cleanAddress = cleanAddress;
      targetInfo.rowCount = selectedRange.rowCount || 1;
      targetInfo.columnCount = selectedRange.columnCount || 1;
      targetInfo.cellCount = selectedRange.cellCount || (targetInfo.rowCount * targetInfo.columnCount);

      if (overlappingTables && overlappingTables.items && overlappingTables.items.length > 0) {
        const tbl = overlappingTables.items[0];
        targetInfo.type = "table";
        targetInfo.table = {
          id: tbl.id,
          name: tbl.name,
          range: cleanAddress
        };
      } else if (targetInfo.cellCount > 1) {
        targetInfo.type = "range";
      } else {
        targetInfo.type = "cell";
      }

      try {
        if (targetInfo.cellCount <= 5000) {
          selectedRange.load("values");
          await context.sync();
          if (selectedRange.values && Array.isArray(selectedRange.values)) {
            targetInfo.rawText = selectedRange.values
              .map(row => row.map(v => (v !== null && v !== undefined) ? String(v).trim() : "").filter(Boolean).join("\t"))
              .filter(line => line.length > 0)
              .join("\n")
              .trim();
          }
        } else {
          const usedRange = sheet.getUsedRangeOrNullObject(true);
          const boundedRange = selectedRange.getIntersectionOrNullObject(usedRange);
          boundedRange.load(["isNullObject", "values"]);
          await context.sync();
          if (!boundedRange.isNullObject && boundedRange.values && Array.isArray(boundedRange.values)) {
            targetInfo.rawText = boundedRange.values
              .map(row => row.map(v => (v !== null && v !== undefined) ? String(v).trim() : "").filter(Boolean).join("\t"))
              .filter(line => line.length > 0)
              .join("\n")
              .trim();
          }
        }
      } catch (valErr) {
        console.debug("Selected range values lookup notice:", valErr);
      }

      return targetInfo;
    } catch (err) {
      console.warn("getSelectionTargetInfo error:", err);
      return targetInfo;
    }
  }

  /**
   * Formats the worksheet schema into structured prompt context for Gemini / StreamAssist
   * @param {Object} schema 
   * @param {string} targetCellAddress 
   * @param {Object} activeTarget - Optional active target metadata
   * @param {string} [mode] - Optional legacy mode string
   * @returns {string} Formatted context string
   */
  formatSchemaForPrompt(schema, targetCellAddress = "", activeTarget = null, mode = null) {
    if (!schema || !schema.columns || schema.columns.length === 0) {
      return "";
    }

    const lines = [];
    const headerRow = schema.startRow || 1;
    const dataStartRow = (schema.rowCount > 1) ? (headerRow + 1) : headerRow;
    const endRow = schema.endRow || (headerRow + (schema.rowCount || 1) - 1);

    lines.push(`Worksheet: "${schema.sheetName}" | Table Range: Rows ${headerRow} to ${endRow} (Headers: Row ${headerRow}, Data: Rows ${dataStartRow} to ${endRow})`);

    // 1. Target Range Selection Context (Always Active)
    if (activeTarget && activeTarget.type === "table" && activeTarget.table) {
      lines.push(`\n🎯 TARGET RANGE (Grid Selection): Excel Structured Table "${activeTarget.table.name}" (Range: ${activeTarget.table.range})`);
      lines.push(`- Instructions: When asked to format, modify, or process table data, operate on this targeted selection.`);
    } else if (activeTarget && activeTarget.type === "range") {
      lines.push(`\n🎯 TARGET RANGE (Grid Selection): Multi-Cell Range (${activeTarget.cleanAddress})`);
      lines.push(`- Instructions: When asked to format, style, fill, or modify data, operate on this targeted selection.`);
    } else {
      const cellAddr = targetCellAddress || (activeTarget ? activeTarget.cleanAddress : 'A1');
      lines.push(`\n🎯 TARGET RANGE (Grid Selection): Cell ${cellAddr}`);
      lines.push(`- Instructions: Generate a standard Excel formula starting with '=' or calculate values for this target cell.`);
    }

    // 2. Target Chart Object (If Included in Context)
    const activeChartObj = (activeTarget && activeTarget.chart) ? activeTarget.chart : this.lockedChart;
    if (activeChartObj) {
      lines.push(`\n📊 TARGET CHART OBJECT (In Context): Excel Native Chart "${activeChartObj.name}" (ID: "${activeChartObj.id}")`);
      lines.push(`- Current Type: ${activeChartObj.chartType}`);
      lines.push(`- Current Title: "${activeChartObj.title}"`);
      lines.push(`- Dimensions: ${activeChartObj.width}pt x ${activeChartObj.height}pt`);
      lines.push(`- Worksheet: "${schema.sheetName}"`);
      lines.push(`- Chart Action Instructions (CRITICAL): The user has included this chart in their context. When the user's prompt references or requests modifying/updating a chart (e.g. changing type, updating title, adjusting series/legend/labels), emit a declarative \`\`\`excel-chart-edit JSON block targeting this chart.`);
    }

    lines.push("\nColumn Schema (Letter -> Full Column Range -> Header -> Sample Data):");
    schema.columns.forEach(col => {
      const colRange = `${col.letter}${headerRow}:${col.letter}${endRow}`;
      const sampleText = col.samples.length > 0 ? ` (Samples: ${col.samples.join(", ")})` : "";
      lines.push(`- Column ${col.letter} (${colRange}): "${col.header}"${sampleText}`);
    });

    if (schema.tables && schema.tables.length > 0) {
      lines.push("\nNamed Excel Tables:");
      schema.tables.forEach(tbl => {
        lines.push(`- Table "${tbl.name}" with columns: [${tbl.columns.join(", ")}]`);
      });
    }

    if (schema.charts && schema.charts.length > 0) {
      lines.push("\nExisting Charts on Worksheet:");
      schema.charts.forEach(ch => {
        lines.push(`- Chart "${ch.name}" (ID: ${ch.id}) | Type: ${ch.type} | Title: "${ch.title}" (${ch.width}x${ch.height}pt)`);
      });
    }

    if (schema.gridData && Array.isArray(schema.gridData) && schema.gridData.length > 0) {
      lines.push("\nWorksheet Grid Data:");
      let currentLength = lines.join("\n").length;
      for (let idx = 0; idx < schema.gridData.length; idx++) {
        const row = schema.gridData[idx];
        const rowNum = headerRow + idx;
        const rowStr = row.map(v => (v !== null && v !== undefined) ? String(v).replace(/\t/g, " ") : "").join(" | ");
        const line = `Row ${rowNum}: | ${rowStr} |`;
        if (currentLength + line.length > 3900000) {
          lines.push(`... [Truncated: dataset exceeds 4,000,000 characters]`);
          break;
        }
        lines.push(line);
        currentLength += line.length + 1;
      }
    } else if (schema.sampleRows && schema.sampleRows.length > 0) {
      lines.push("\nSample Worksheet Grid (First Rows):");
      schema.sampleRows.forEach((row, idx) => {
        const rowNum = headerRow + idx;
        const rowStr = row.map(v => (v !== null && v !== undefined) ? String(v).replace(/\t/g, " ") : "").join(" | ");
        lines.push(`Row ${rowNum}: | ${rowStr} |`);
      });
    }

    lines.push("\nFormula Generation Instructions (CRITICAL):");
    lines.push("1. When asked for a calculation or formula, return the exact Excel formula starting with '=' inside a dedicated ```excel code block (e.g. ```excel\n=SUM(I2:I15)\n```).");
    lines.push("2. Use standard uppercase Excel functions (SUM, SUMIFS, COUNTIFS, AVERAGEIFS, XLOOKUP, INDEX/MATCH, IF, IFS).");
    lines.push("3. Wrap string criteria inside double quotes (e.g. \"Approved\").");
    lines.push("4. Reference the correct column letters (e.g. I2:I15) based on the Column Schema and Grid Data provided above.");
    lines.push("5. Follow the formula code block with a clear, concise description of how the formula works.");
    lines.push("6. DO NOT output extraneous ```excel-range-format blocks, formatting recommendations, summary tables, or chart blocks when the user asks for a formula.");

    lines.push("\nNative Chart Creation Instructions:");
    lines.push("1. When asked to create, plot, or visualize data as a chart or graph in Excel, return a declarative ```excel-chart JSON specification code block.");
    lines.push("2. Supported 'type' values: ColumnClustered, ColumnStacked, BarClustered, BarStacked, Line, LineMarkers, Pie, Doughnut, Area, Scatter, Waterfall, Treemap, Sunburst, Histogram, Funnel.");
    lines.push("3. TARGETED MULTI-RANGE SELECTION RULES (CRITICAL):");
    lines.push("   - When the user asks to chart specific fields (e.g., 'Actual Sales by Employee Name', 'Plot Revenue per Product', 'Compare Quota vs Actual'):");
    lines.push("     * Identify the category/label/display name column (e.g. Column B) and specify 'categoriesRange': \"B1:B50\" (or \"B2:B50\").");
    lines.push("     * Identify the metric/value column(s) (e.g. Column E) and specify 'seriesRange': \"E1:E50\" (or \"E2:E50\").");
    lines.push("     * If multiple metrics are requested, specify their range (e.g. 'seriesRange': \"D1:E50\" or \"D1:D50, F1:F50\").");
    lines.push("     * DO NOT select the entire worksheet or unrelated columns!");
    lines.push("   - CATEGORY TOTALS & AGGREGATIONS (CRITICAL FOR LINE-ITEM/TRANSACTION DATA):");
    lines.push("     * When the user asks to chart totals, subtotals, sums, or distributions by category from transaction rows where categories repeat, specify 'aggregate': \"sum\" (or \"average\", \"count\").");
    lines.push("     * The Excel add-in automatically creates a formula-linked Aggregated Summary Table (=SUMIF) and generates the chart from the unique category subtotals!");
    lines.push("   - For contiguous tables covering all columns in range, specify 'dataRange': \"A1:C50\".");
    lines.push("4. Example Targeted Chart Block with Aggregation:\n```excel-chart\n{\n  \"type\": \"BarClustered\",\n  \"categoriesRange\": \"D1:D50\",\n  \"seriesRange\": \"E1:E50\",\n  \"aggregate\": \"sum\",\n  \"title\": \"Distribution of Total Transaction Amount by Category\",\n  \"hasHeaders\": true,\n  \"position\": { \"placement\": \"adjacent\", \"offsetColumns\": 2 }\n}\n```");

    lines.push("\nNative Chart Modification & Editing Instructions:");
    lines.push("1. When asked to modify, update, re-title, switch type, rebind series, toggle legend, or format an existing chart or active chart, return a declarative ```excel-chart-edit JSON specification code block.");
    lines.push("2. Supported chart edit properties:");
    lines.push(`   - "targetChart": Chart name (e.g. "${activeChartObj ? activeChartObj.name : 'Chart 1'}") or "active"`);
    lines.push(`   - "chartId": Optional chart ID (e.g. "${activeChartObj ? activeChartObj.id : ''}")`);
    lines.push(`   - "chartType": New chart type (e.g. "Pie", "Doughnut", "ColumnClustered", "Line", "BarClustered", "Area", "Waterfall", etc.)`);
    lines.push(`   - "title": New title string`);
    lines.push(`   - "categoriesRange": New category labels range (e.g. "B1:B50")`);
    lines.push(`   - "seriesRange": New metric data range(s) (e.g. "E1:E50")`);
    lines.push(`   - "dataRange": New contiguous data range (e.g. "A1:E50")`);
    lines.push(`   - "legend": { "visible": true|false, "position": "Top"|"Bottom"|"Left"|"Right"|"Corner" }`);
    lines.push(`   - "dataLabels": { "visible": true|false, "showCategoryName": true|false, "showPercentage": true|false, "showValue": true|false }`);
    lines.push(`   - "position": { "targetCell": "G1" }`);
    lines.push(`   - "dimensions": { "width": 480, "height": 300 }`);
    lines.push("3. Example Chart Edit Block:\n```excel-chart-edit\n{\n  \"targetChart\": \"" + (activeChartObj ? activeChartObj.name : 'Chart 1') + "\",\n  \"chartType\": \"Doughnut\",\n  \"title\": \"Updated Revenue Distribution\",\n  \"legend\": { \"visible\": true, \"position\": \"Bottom\" },\n  \"dataLabels\": { \"visible\": true, \"showPercentage\": true, \"showCategoryName\": true }\n}\n```");

    lines.push("\nRange Formatting & Cell Modification Instructions (CRITICAL):");
    lines.push("1. ONLY when the user explicitly asks to format, style, fill color, highlight, bold, align, set borders, format numbers, or modify cell styles in Excel, return a declarative ```excel-range-format JSON specification code block. DO NOT output ```excel-range-format blocks when answering formula requests.");
    lines.push("2. DO NOT output JavaScript code or Office.js API calls (such as targetRange.format.fill.color = ... or =targetRange.getFormat().getFill().setColor(...)); ALWAYS output the ```excel-range-format JSON specification block.");
    lines.push("3. Supported range format properties:");
    const targetRangeClean = activeTarget ? activeTarget.cleanAddress : (targetCellAddress || "A1:D10");
    lines.push(`   - "targetRange": "${targetRangeClean}" (or omit to apply directly to active selection)`);
    lines.push(`   - "fill": { "color": "lightblue" | "lightgreen" | "yellow" | "#E0F2FE" | "#HEX" }`);
    lines.push(`   - "font": { "bold": true|false, "italic": true|false, "underline": true|false, "color": "#004e8c" | "#HEX", "size": number, "name": "Segoe UI" }`);
    lines.push(`   - "horizontalAlignment": "Left"|"Center"|"Right"|"Justify"`);
    lines.push(`   - "verticalAlignment": "Top"|"Center"|"Bottom"`);
    lines.push(`   - "wrapText": true|false`);
    lines.push(`   - "numberFormat": "$#,##0.00" | "0.0%" | "#,##0" | "yyyy-mm-dd" | "@" | "currency" | "percentage" | "accounting"`);
    lines.push(`   - "borders": { "edge": "All"|"Outline"|"Bottom", "style": "Continuous", "color": "#D3D3D3", "weight": "Thin" } or true`);
    lines.push(`   - "autofitColumns": true|false`);
    lines.push(`   - "autofitRows": true|false`);
    lines.push(`   - "clear": "formats"|"contents"|"all" (optional)`);
    lines.push(`4. Example Range Format Block:\n\`\`\`excel-range-format\n{\n  "targetRange": "${targetRangeClean}",\n  "fill": { "color": "lightblue" },\n  "font": { "bold": true, "color": "#003366" },\n  "horizontalAlignment": "Center",\n  "numberFormat": "currency",\n  "autofitColumns": true\n}\n\`\`\``);

    lines.push("\nTaxonomy Categorization & Column Insertion Instructions (CRITICAL):");
    lines.push("1. When asked to categorize, classify, tag, clean, standardize, or assign taxonomy categories to a text column in Excel (e.g., Raw Merchant Names into Expense Types, Products into Categories, Customer Inquiries into Intents, Job Roles into Departments):");
    lines.push("   - Return a declarative ```excel-taxonomy JSON specification code block.");
    lines.push("   - Specify 'sourceColumn' (e.g. \"B\"), 'targetColumn' (e.g. \"C\"), 'insertColumn': true, 'header' (e.g. \"Expense Category\"), 'startRow': 2, and 'categories': [ { \"row\": 2, \"sourceValue\": \"Uber *Trip\", \"category\": \"Travel / Transportation\" }, ... ].");
    lines.push("   - Include a 'taxonomySummary' dictionary showing the total distribution count for each standardized category.");
    lines.push("   - Provide a clean summary table in your response explaining the mapped categories.");
    lines.push(`2. Example Taxonomy Categorization Block:\n\`\`\`excel-taxonomy\n{\n  "action": "categorize",\n  "sourceColumn": "B",\n  "targetColumn": "C",\n  "insertColumn": true,\n  "header": "Expense Category",\n  "startRow": 2,\n  "categories": [\n    { "row": 2, "sourceValue": "Uber *Trip 9823", "category": "Travel & Transportation" },\n    { "row": 3, "sourceValue": "AWS EMEA Cloud", "category": "Cloud Infrastructure" },\n    { "row": 4, "sourceValue": "Sweetgreen SF", "category": "Meals & Entertainment" },\n    { "row": 5, "sourceValue": "Delta Airlines Flight", "category": "Travel & Transportation" }\n  ],\n  "taxonomySummary": {\n    "Travel & Transportation": 2,\n    "Cloud Infrastructure": 1,\n    "Meals & Entertainment": 1\n  }\n}\n\`\`\``);

    return lines.join("\n");
  }

  /**
   * Read only the raw cell text values from the currently selected cell or range in Excel.
   * Returns an empty string when the selected cell(s) have no text content.
   * @returns {Promise<string>}
   */
  async getSelectedRawText() {
    let rawText = "";
    try {
      if (typeof Excel !== 'undefined' && Excel.run) {
        await Excel.run(async (context) => {
          const targetInfo = await this.getSelectionTargetInfo(context);
          rawText = (targetInfo && targetInfo.rawText) ? targetInfo.rawText.trim() : "";
        });
      }
    } catch (err) {
      console.warn("Excel raw selection read error:", err);
    }
    return rawText;
  }

  /**
   * Read currently selected cell range text & active cell context in Excel.
   * For single cells and multi-cell ranges, provides range address, dimensions, content, and worksheet schema.
   * Also detects if an active or sticky locked chart or structured table is selected on the worksheet.
   * @returns {Promise<string>}
   */
  async getSelectedText() {
    let selectedText = "";
    try {
      if (typeof Excel !== 'undefined') {
        await Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          const targetInfo = await this.getSelectionTargetInfo(context);
          const schema = await this.getWorksheetSchema(context, sheet);

          const range = context.workbook.getSelectedRange();
          range.load("values, formulas, address, rowCount, columnCount, rowIndex, columnIndex");
          await context.sync();

          const schemaContext = this.formatSchemaForPrompt(schema, targetInfo.address, targetInfo);

          const startColLetter = columnIndexToLetter(range.columnIndex || 0);
          const endColLetter = columnIndexToLetter((range.columnIndex || 0) + (range.columnCount || 1) - 1);
          const colSpan = startColLetter === endColLetter ? `${startColLetter}:${startColLetter}` : `${startColLetter}:${endColLetter}`;

          const hasValues = range.values && Array.isArray(range.values);
          const rawCellText = hasValues
            ? range.values.map(row => row.map(v => (v !== null && v !== undefined) ? String(v) : "").join("\t")).filter(line => line.trim().length > 0).join("\n")
            : "";

          const ch = targetInfo.chart;
          let chartHeader = "";
          if (ch) {
            chartHeader = `📊 TARGET CHART (Included in Context): Chart "${ch.name}" (ID: ${ch.id}) | Type: ${ch.chartType} | Title: "${ch.title}" (${ch.width}x${ch.height}pt)\n` +
              `- Note: A target chart is included in the context. If the user prompt references or asks to modify/update a chart, emit a \`\`\`excel-chart-edit JSON block targeting this chart.\n\n`;
          }

          if (targetInfo.type === "table" && targetInfo.table) {
            selectedText = `${chartHeader}🎯 TARGET RANGE (Table): "${targetInfo.table.name}" (Range: ${targetInfo.cleanAddress})\n` +
              `Table Span: ${targetInfo.rowCount} rows x ${targetInfo.columnCount} columns\n` +
              `Worksheet: "${targetInfo.sheetName}"\n\n` +
              `${schemaContext}`;
            return;
          }

          if (targetInfo.type === "cell") {
            // Single cell selected
            if (rawCellText.trim()) {
              selectedText = `${chartHeader}🎯 TARGET RANGE (Cell): ${targetInfo.cleanAddress} (${targetInfo.address}) | Column: ${startColLetter}\n` +
                `Cell Content: "${rawCellText.trim()}"\n\n` +
                `${schemaContext}`;
            } else {
              selectedText = `${chartHeader}🎯 TARGET RANGE (Empty Cell): ${targetInfo.cleanAddress} (${targetInfo.address}) | Column: ${startColLetter}\n\n` +
                `${schemaContext}`;
            }
          } else {
            // Multi-cell range selected (e.g. B1:B30, A2:D50)
            const rowCount = targetInfo.rowCount;
            const colCount = targetInfo.columnCount;

            let contentDisplay = rawCellText;
            if (rowCount > 30) {
              const firstRows = (range.values || []).slice(0, 15).map(r => r.join("\t")).join("\n");
              const lastRows = (range.values || []).slice(-10).map(r => r.join("\t")).join("\n");
              contentDisplay = `${firstRows}\n... [${rowCount - 25} rows omitted] ...\n${lastRows}`;
            }

            selectedText = `${chartHeader}🎯 TARGET RANGE: ${targetInfo.cleanAddress} (${targetInfo.address})\n` +
              `Range Span: Column(s) ${colSpan} (${rowCount} rows x ${colCount} cols)\n` +
              `Selected Range Content:\n"""\n${contentDisplay}\n"""\n\n` +
              `${schemaContext}`;
          }
        });
      }
    } catch (err) {
      console.warn("Excel selection read error:", err);
    }
    return selectedText;
  }

  /**
   * Read full active worksheet text, columns, and data context for Q&A and spreadsheet intelligence
   * @returns {Promise<string>}
   */
  async getFullDocumentText() {
    try {
      const fullSheetData = await ExcelInsightsService.getFullSheetData();
      if (fullSheetData && fullSheetData.formattedText) {
        return fullSheetData.formattedText;
      }
    } catch (err) {
      console.warn("Excel full document text read error:", err);
    }

    let fullText = "";
    try {
      if (typeof Excel !== 'undefined') {
        await Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          const activeRange = context.workbook.getSelectedRange();
          activeRange.load("address");
          
          const schema = await this.getWorksheetSchema(context, sheet);
          await context.sync();

          fullText = this.formatSchemaForPrompt(schema, activeRange.address);
        });
      }
    } catch (err) {
      console.warn("Excel full document read error:", err);
    }
    return fullText || this.getSelectedText();
  }

  /**
   * Build Table Insights Prompt (delegates to ExcelInsightsService)
   */
  async getTableInsightsPrompt(scope = "auto", customInstruction = "") {
    return await ExcelInsightsService.buildTableInsightsPrompt(scope, customInstruction);
  }

  /**
   * Build Executive Summary Prompt (delegates to ExcelInsightsService)
   */
  async getExecutiveSummaryPrompt(scope = "auto", customInstruction = "") {
    return await ExcelInsightsService.buildExecutiveSummaryPrompt(scope, customInstruction);
  }

  /**
   * Insert Executive Summary Sheet into Workbook (delegates to ExcelInsightsService)
   */
  async insertExecutiveSummaryWorksheet(summaryTitle, textContent) {
    return await ExcelInsightsService.insertExecutiveSummaryWorksheet(summaryTitle, textContent);
  }

  /**
   * Insert native visual chart/graph image into Excel active worksheet
   */
  async insertVisualChart(base64DataUri) {
    return await ExcelInsightsService.insertVisualImage(base64DataUri);
  }

  /**
   * Insert native interactive Excel chart linked directly to concise AI tabular dataset
   */
  async insertNativeInteractiveChart(chartTitle = "Executive Summary Analysis", customType = null, responseText = "") {
    return await ExcelInsightsService.insertNativeInteractiveChart(chartTitle, customType, responseText);
  }

  /**
   * Insert Structured Insights Summary Table directly at the user's selected cell location
   */
  async insertInsightsSummaryTable(rawText = "", htmlContent = "") {
    return await ExcelInsightsService.insertInsightsSummaryTable(rawText, htmlContent);
  }

  /**
   * Insert and populate Taxonomy Categories column into worksheet
   */
  async insertTaxonomyColumn(rawTextOrSpec = "", htmlContent = "", options = {}) {
    if (typeof Excel === 'undefined') return;
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      let spec = null;
      if (rawTextOrSpec && typeof rawTextOrSpec === 'object' && (rawTextOrSpec.categories || rawTextOrSpec.taxonomy)) {
        spec = rawTextOrSpec;
      } else if (options && options.spec && (options.spec.categories || options.spec.taxonomy)) {
        spec = options.spec;
      } else {
        spec = extractExcelTaxonomySpec(rawTextOrSpec || "", htmlContent, { isTaxonomy: true, ...options });
      }
      if (!spec) {
        throw new Error("No taxonomy categories found in response.");
      }
      return await applyTaxonomyCategorization(context, sheet, spec);
    });
  }

  /**
   * Edit existing native Excel chart in-place
   */
  async editChart(specOrRawText, htmlContent = "") {
    if (typeof Excel === 'undefined') return;
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      let spec = null;
      if (specOrRawText && typeof specOrRawText === "object" && specOrRawText.action) {
        spec = specOrRawText;
      } else {
        spec = extractExcelChartEditSpec(specOrRawText || "", htmlContent);
      }
      if (!spec) {
        throw new Error("No chart update instructions could be extracted.");
      }
      return await editNativeChart(context, sheet, spec);
    });
  }

  /**
   * Delegate table insights prompt construction to ExcelInsightsService
   * @param {string} scope - "auto", "selection", or "sheet"
   * @param {string} customInstruction
   * @returns {Promise<string>}
   */
  async getTableInsightsPrompt(scope = "auto", customInstruction = "") {
    return ExcelInsightsService.buildTableInsightsPrompt(scope, customInstruction);
  }

  /**
   * Delegate executive summary prompt construction to ExcelInsightsService
   * @param {string} scope - "auto", "selection", or "sheet"
   * @param {string} customInstruction
   * @returns {Promise<string>}
   */
  async getExecutiveSummaryPrompt(scope = "auto", customInstruction = "") {
    return ExcelInsightsService.buildExecutiveSummaryPrompt(scope, customInstruction);
  }

  /**
   * Delegate taxonomy categorizer prompt construction to ExcelInsightsService
   * @param {string} scope - "auto", "selection", or "sheet"
   * @param {string} customInstruction
   * @returns {Promise<string>}
   */
  async getTaxonomyPrompt(scope = "auto", customInstruction = "") {
    return ExcelInsightsService.buildTaxonomyPrompt(scope, customInstruction);
  }

  /**
   * Insert generated Executive Summary as a new dedicated worksheet
   * @param {string} summaryTitle
   * @param {string} textContent
   */
  async insertExecutiveSummaryWorksheet(summaryTitle = "Executive Summary", textContent = "") {
    return ExcelInsightsService.insertExecutiveSummaryWorksheet(summaryTitle, textContent);
  }

  /**
   * Insert AI response into Excel:
   * - Evaluates chart edits on existing charts in-place
   * - Inserts native Excel charts
   * - Evaluates formulas directly into the selected top-left cell (e.g. =SUMIFS(C:C, D:D, "Approved"))
   * - Injects tables and long-form analysis (headings, bullet points, paragraphs) row-by-row into a normalized rectangular matrix
   * - Avoids dimension mismatch errors by sizing target range to the exact rectangular matrix dimensions
   * 
   * @param {string} htmlContent - Parsed HTML content
   * @param {string} rawText - Raw markdown / text response from Gemini
   * @param {string} mode - "smart", "replace_draft", "chart_edit", "chart", etc.
   */
  async insertContent(htmlContent, rawText = "", mode = "smart") {
    if (!htmlContent && !rawText && mode === "smart") return;
    try {
      if (typeof Excel !== 'undefined') {
        await Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          const range = context.workbook.getSelectedRange();
          range.load("rowIndex, columnIndex, address, rowCount, columnCount");
          await context.sync();

          const tempDiv = typeof document !== "undefined" && document.createElement
            ? document.createElement("div")
            : { innerHTML: "", children: [], innerText: "" };
          tempDiv.innerHTML = htmlContent || rawText || "";

          // 1. Priority 1: Check if mode is chart_edit or if chart edit specification is present
          const detectedChartEdit = (mode === "chart_edit")
            ? (extractExcelChartEditSpec(rawText, htmlContent) || normalizeChartEditSpec({ action: "edit", targetChart: "active" }))
            : extractExcelChartEditSpec(rawText, htmlContent);
          if (detectedChartEdit) {
            console.log(`Editing existing native Excel chart (${detectedChartEdit.targetChart || detectedChartEdit.chartId || 'active'})`);
            await editNativeChart(context, sheet, detectedChartEdit);
            return;
          }

          // 2. Priority 2: Check if the response contains a Native Excel Chart specification
          let detectedChartSpec = extractExcelChartSpec(rawText, htmlContent);
          if (!detectedChartSpec && mode === "chart") {
            detectedChartSpec = normalizeChartSpec({
              type: "ColumnClustered",
              title: "Data Visualization"
            });
          }
          if (detectedChartSpec) {
            console.log(`Creating native Excel chart (${detectedChartSpec.type}) for range ${detectedChartSpec.dataRange || detectedChartSpec.seriesRange}`);
            await insertNativeChart(context, sheet, detectedChartSpec, this);
            return;
          }

          // 3. Priority 3: Check if the response contains an Excel Range Format / Modification specification
          const detectedRangeFormat = extractExcelRangeFormatSpec(rawText, htmlContent);
          if (detectedRangeFormat) {
            console.log(`Applying range formatting:`, detectedRangeFormat);
            await applyRangeFormatting(context, sheet, detectedRangeFormat);
            return;
          }

          // 4. Priority 4: Check if the response contains an Excel Taxonomy Categorization specification
          const detectedTaxonomy = extractExcelTaxonomySpec(rawText, htmlContent, { isTaxonomy: (mode === "taxonomy") });
          if (detectedTaxonomy) {
            console.log(`Applying taxonomy categorization from Column ${detectedTaxonomy.sourceColumn} to Column ${detectedTaxonomy.targetColumn}`);
            await applyTaxonomyCategorization(context, sheet, detectedTaxonomy);
            return;
          }

          // 5. Priority 5: Check if the response contains an Excel Formula (e.g., =SUMIFS, =XLOOKUP)
          const detectedFormula = extractExcelFormula(rawText, htmlContent);
          if (detectedFormula) {
            let formulaTargetRange = range;
            const preferredTarget = (typeof window !== "undefined" && window.__preferredFormulaTargetCell)
              ? String(window.__preferredFormulaTargetCell).trim()
              : "";
            if (preferredTarget) {
              try {
                const cleanTargetAddr = preferredTarget.includes("!")
                  ? preferredTarget.split("!").pop().trim()
                  : preferredTarget;
                formulaTargetRange = sheet.getRange(cleanTargetAddr);
                formulaTargetRange.load("rowIndex, columnIndex, address, rowCount, columnCount");
                await context.sync();
                formulaTargetRange.select();
              } catch (targetErr) {
                console.warn(`Could not resolve preferred target cell "${preferredTarget}", falling back to active selection:`, targetErr);
                formulaTargetRange = range;
              }
            }

            console.log(`Setting Excel formula in selection (${formulaTargetRange.rowIndex}, ${formulaTargetRange.columnIndex}): ${detectedFormula}`);
            
            const rCount = Math.max(1, formulaTargetRange.rowCount || 1);
            const cCount = Math.max(1, formulaTargetRange.columnCount || 1);
            const formulaMatrix = Array.from({ length: rCount }, () => Array(cCount).fill(detectedFormula));

            // Resilient formula assignment with 3-tier fallback (formulas -> formulasLocal -> values)
            try {
              formulaTargetRange.formulas = formulaMatrix;
              await context.sync();
            } catch (formErr) {
              console.warn("Standard formula insertion notice, attempting formulasLocal fallback:", formErr);
              try {
                formulaTargetRange.formulasLocal = formulaMatrix;
                await context.sync();
              } catch (locErr) {
                console.warn("formulasLocal insertion notice, attempting values fallback:", locErr);
                formulaTargetRange.values = formulaMatrix;
                await context.sync();
              }
            }

            // Safe formatting and autofit (isolated in try-catch so styling never fails insertion)
            try {
              formulaTargetRange.format.font.name = "Segoe UI";
              formulaTargetRange.format.font.size = 11;
              formulaTargetRange.format.autofitColumns();
              await context.sync();
            } catch (fmtErr) {
              console.warn("Formula cell formatting notice:", fmtErr);
            }
            return;
          }

          // 6. Priority 6: Structured Document / Table / Analysis multi-row insertion (Clipboard-like behavior)
          const rows = [];
          const formatting = []; // { rowIndex, isHeader, isBold, fontSize, color }

          const childNodes = Array.from(tempDiv.children || []);

          if (childNodes.length > 0) {
            childNodes.forEach(node => {
              const tag = node.tagName ? node.tagName.toUpperCase() : "";
              const text = ExcelInsightsService.cleanCellText(node.innerText || node.textContent || "");

              if (!text && tag !== "TABLE") return;

              if (tag === "H1" || tag === "H2" || tag === "H3") {
                rows.push([text]);
                formatting.push({
                  rowIndex: rows.length - 1,
                  isBold: true,
                  fontSize: tag === "H1" ? 13 : 12,
                  color: "#004e8c"
                });
              } else if (tag === "P") {
                rows.push([text]);
              } else if (tag === "UL" || tag === "OL") {
                const lis = Array.from(node.querySelectorAll("li"));
                lis.forEach((li, idx) => {
                  const prefix = tag === "OL" ? `${idx + 1}. ` : "• ";
                  const liText = ExcelInsightsService.cleanCellText(li.innerText || li.textContent || "");
                  if (liText) rows.push([prefix + liText]);
                });
              } else if (tag === "TABLE") {
                const trs = Array.from(node.querySelectorAll("tr"));
                trs.forEach(tr => {
                  const cells = Array.from(tr.querySelectorAll("th, td")).map(td => ExcelInsightsService.cleanCellText(td.innerText || td.textContent || ""));
                  if (cells.length > 0) {
                    rows.push(cells);
                    if (tr.querySelector("th")) {
                      formatting.push({ rowIndex: rows.length - 1, isHeader: true });
                    }
                  }
                });
              } else if (tag === "BLOCKQUOTE" || node.classList.contains("note")) {
                rows.push([`💡 ${text}`]);
                formatting.push({ rowIndex: rows.length - 1, isBold: true, color: "#106ebe" });
              } else {
                rows.push([text]);
              }
            });
          }

          // Fallback if no DOM children were generated
          if (rows.length === 0) {
            const fallbackText = (tempDiv.innerText || rawText || "").trim();
            if (fallbackText) {
              const lines = fallbackText.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
              lines.forEach(l => {
                const cleanLine = ExcelInsightsService.cleanCellText(l);
                if (cleanLine.includes("\t")) {
                  rows.push(cleanLine.split("\t").map(c => ExcelInsightsService.cleanCellText(c)));
                } else {
                  rows.push([cleanLine]);
                }
              });
            }
          }

          if (rows.length === 0) {
            rows.push([""]);
          }

          // 7. Normalize into a strictly rectangular matrix [rowCount x maxColCount]
          const rowCount = rows.length;
          const maxColCount = Math.max(1, ...rows.map(r => r.length));

          const matrix = rows.map(r => {
            const rowCopy = r.map(c => (typeof c === "string" ? ExcelInsightsService.cleanCellText(c) : c));
            while (rowCopy.length < maxColCount) {
              rowCopy.push("");
            }
            return rowCopy;
          });

          // 8. Target Range based exactly on matrix dimensions starting at active cell
          const startRow = range.rowIndex;
          const startCol = range.columnIndex;
          const targetRange = sheet.getRangeByIndexes(startRow, startCol, rowCount, maxColCount);

          targetRange.values = matrix;

          // Apply cell styling for headers and headings
          if (formatting.length > 0) {
            for (const fmt of formatting) {
              const rowSubRange = sheet.getRangeByIndexes(startRow + fmt.rowIndex, startCol, 1, maxColCount);
              if (fmt.isHeader) {
                rowSubRange.format.font.bold = true;
                rowSubRange.format.font.color = "#004e8c";
                rowSubRange.format.fill.color = "#EBF3FC";
              }
              if (fmt.isBold) {
                rowSubRange.format.font.bold = true;
              }
              if (fmt.fontSize) {
                rowSubRange.format.font.size = fmt.fontSize;
              }
              if (fmt.color) {
                rowSubRange.format.font.color = fmt.color;
              }
            }
          }

          try {
            targetRange.format.autofitColumns();
          } catch (afErr) {
            console.warn("Autofit notice:", afErr);
          }
          await context.sync();
        });
      }
    } catch (err) {
      console.error("Excel insertion error:", err);
      throw err;
    }
  }

  /**
   * Scan and execute in-cell @gemini commands or plain calculation prompts in Excel
   * @param {boolean} forceRun - Force execution even if cell is currently active
   * @param {Object} callbacks - Callback functions { onStatus, executePrompt }
   */
  async checkInDocumentCommands(forceRun = false, callbacks = {}) {
    let targetPrompt = "";
    let targetAddress = "";

    try {
      if (typeof Excel !== 'undefined') {
        await Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          const range = context.workbook.getSelectedRange();
          range.load("values, formulas, address, rowCount, columnCount");
          await context.sync();

          if (range.values && range.values.length > 0 && range.values[0].length > 0) {
            const cellValue = String(range.values[0][0] || "").trim();
            const cellFormula = String(range.formulas[0][0] || "").trim();
            const textToInspect = cellValue || cellFormula;

            const geminiIdx = textToInspect.toLowerCase().indexOf("@gemini");
            if (geminiIdx !== -1) {
              const userPrompt = textToInspect.substring(geminiIdx + 7).replace(/^[:\s]+/, "").trim();
              if (userPrompt.length >= 2) {
                targetPrompt = userPrompt;
                targetAddress = range.address;

                if (callbacks.onStatus) {
                  callbacks.onStatus(`Found: "@gemini ${userPrompt.substring(0, 25)}..." in ${targetAddress}`);
                }

                // Temporary placeholder while generating
                range.values = [["⚡ [Gemini generating formula...]"]];
                await context.sync();
              }
            } else if (forceRun && textToInspect.length > 3) {
              // If user explicitly clicked "Run @gemini" button on an active cell with instructions
              targetPrompt = textToInspect;
              targetAddress = range.address;

              if (callbacks.onStatus) {
                callbacks.onStatus(`Processing cell ${targetAddress}: "${targetPrompt.substring(0, 25)}..."`);
              }

              range.values = [["⚡ [Gemini generating formula...]"]];
              await context.sync();
            }
          }
        });

        if (targetPrompt && callbacks.executePrompt) {
          // Read full schema context to ensure accurate formula generation
          let schemaContext = "";
          await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const schema = await this.getWorksheetSchema(context, sheet);
            schemaContext = this.formatSchemaForPrompt(schema, targetAddress);
          });

          const fullPrompt = `Task: Generate an Excel formula or calculation for cell ${targetAddress}.\nInstruction: "${targetPrompt}"\n\n${schemaContext}`;
          
          const formattedHtml = await callbacks.executePrompt(fullPrompt);

          if (formattedHtml) {
            await Excel.run(async (context) => {
              const sheet = context.workbook.worksheets.getActiveWorksheet();
              const range = sheet.getRange(targetAddress);
              range.load("address, rowCount, columnCount");
              await context.sync();

              // 1. Check if in-cell response generated an Excel chart edit
              const chartEditSpec = extractExcelChartEditSpec(formattedHtml);
              if (chartEditSpec) {
                range.values = [[""]]; // Clear the placeholder text
                await context.sync();
                await editNativeChart(context, sheet, chartEditSpec);
                if (callbacks.onStatus) {
                  callbacks.onStatus(`✅ Updated chart (${chartEditSpec.targetChart || 'active'})`);
                }
                return;
              }

              // 2. Check if in-cell response generated an Excel chart
              const chartSpec = extractExcelChartSpec(formattedHtml);
              if (chartSpec) {
                range.values = [[""]]; // Clear the placeholder text
                await context.sync();
                await insertNativeChart(context, sheet, chartSpec, this);
                if (callbacks.onStatus) {
                  callbacks.onStatus(`✅ Created native Excel chart (${chartSpec.type})`);
                }
                return;
              }

              // 3. Check if in-cell response generated range formatting
              const rangeFormatSpec = extractExcelRangeFormatSpec(formattedHtml);
              if (rangeFormatSpec) {
                range.values = [[""]]; // Clear the placeholder text
                await context.sync();
                await applyRangeFormatting(context, sheet, rangeFormatSpec);
                if (callbacks.onStatus) {
                  callbacks.onStatus(`✅ Applied formatting to range`);
                }
                return;
              }

              // 4. Check if in-cell response generated taxonomy categorization
              const taxonomySpec = extractExcelTaxonomySpec(formattedHtml);
              if (taxonomySpec) {
                range.values = [[""]]; // Clear the placeholder text
                await context.sync();
                await applyTaxonomyCategorization(context, sheet, taxonomySpec);
                if (callbacks.onStatus) {
                  callbacks.onStatus(`✅ Categorized ${taxonomySpec.count || ''} rows into Column ${taxonomySpec.targetColumn}`);
                }
                return;
              }

              const formula = extractExcelFormula(formattedHtml);
              if (formula) {
                const rCount = Math.max(1, range.rowCount || 1);
                const cCount = Math.max(1, range.columnCount || 1);
                const formulaMatrix = Array.from({ length: rCount }, () => Array(cCount).fill(formula));

                try {
                  range.formulas = formulaMatrix;
                  await context.sync();
                } catch (fErr) {
                  try {
                    range.formulasLocal = formulaMatrix;
                    await context.sync();
                  } catch (flErr) {
                    range.values = formulaMatrix;
                    await context.sync();
                  }
                }
                try {
                  range.format.font.name = "Segoe UI";
                  range.format.font.size = 11;
                  range.format.autofitColumns();
                  await context.sync();
                } catch (fmtErr) {}
                if (callbacks.onStatus) {
                  callbacks.onStatus(`✅ Evaluated formula in ${targetAddress}: ${formula}`);
                }
              } else {
                // If HTML table or text
                const tempDiv = typeof document !== "undefined" && document.createElement
                  ? document.createElement("div")
                  : { innerHTML: "", innerText: String(formattedHtml || "").replace(/<[^>]+>/g, " ") };
                tempDiv.innerHTML = formattedHtml;
                const plainText = tempDiv.innerText || tempDiv.textContent || "";
                range.values = [[plainText.trim()]];
                try {
                  range.format.font.name = "Segoe UI";
                  await context.sync();
                } catch (fErr) {}
                if (callbacks.onStatus) {
                  callbacks.onStatus(`✅ Updated cell ${targetAddress}`);
                }
              }
            });
          }
        }
      }
    } catch (err) {
      console.warn("Excel in-cell command scan error:", err);
      if (callbacks.onStatus) callbacks.onStatus(`Excel Ready`);
    }
  }
}

