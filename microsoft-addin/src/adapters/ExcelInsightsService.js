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
 * Excel Table Insights & Executive Summary Engine
 * 
 * Specialized data extractor, statistical pre-parser, and prompt generator
 * for Microsoft Excel table insights and executive summary workflows.
 * Supports both highlighted cell selection and entire worksheet scopes.
 * 
 * @author Sathya AG, Principal Architect, Google
 */

export class ExcelInsightsService {
  /**
   * Extract structured data from currently highlighted cells (selected range)
   * Includes values, formulas, header auto-detection, and statistical pre-parsing.
   */
  static async getSelectedRangeData() {
    let data = {
      scope: "selection",
      address: "",
      rowCount: 0,
      columnCount: 0,
      matrix: [],
      formulas: [],
      headers: [],
      stats: [],
      formattedText: ""
    };

    try {
      if (typeof Excel !== 'undefined') {
        await Excel.run(async (context) => {
          const range = context.workbook.getSelectedRange();
          range.load(["address", "rowCount", "columnCount", "values", "formulas"]);
          await context.sync();

          data.address = range.address;
          data.rowCount = range.rowCount;
          data.columnCount = range.columnCount;
          data.matrix = range.values || [];
          data.formulas = range.formulas || [];

          if (data.matrix.length > 0) {
            data.headers = this.extractHeaders(data.matrix);
            data.stats = this.computeColumnStats(data.matrix, data.headers);
            data.formattedText = this.matrixToMarkdownTable(data.matrix);
          }
        });
      }
    } catch (err) {
      console.warn("ExcelInsightsService getSelectedRangeData error:", err);
    }
    return data;
  }

  /**
   * Extract structured data from the entire active worksheet (used range)
   * Includes sheet metadata, native Excel tables, values, formulas, and stats.
   */
  static async getFullSheetData() {
    let data = {
      scope: "sheet",
      sheetName: "",
      address: "",
      rowCount: 0,
      columnCount: 0,
      tableCount: 0,
      tableNames: [],
      matrix: [],
      formulas: [],
      headers: [],
      stats: [],
      formattedText: ""
    };

    try {
      if (typeof Excel !== 'undefined') {
        await Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          sheet.load("name");
          
          const usedRange = sheet.getUsedRange();
          usedRange.load(["address", "rowCount", "columnCount", "values", "formulas"]);
          
          const tables = sheet.tables;
          tables.load("items/name");

          await context.sync();

          data.sheetName = sheet.name;
          data.address = usedRange.address;
          data.rowCount = usedRange.rowCount;
          data.columnCount = usedRange.columnCount;
          data.matrix = usedRange.values || [];
          data.formulas = usedRange.formulas || [];
          data.tableCount = tables.items ? tables.items.length : 0;
          data.tableNames = tables.items ? tables.items.map(t => t.name) : [];

          if (data.matrix.length > 0) {
            data.headers = this.extractHeaders(data.matrix);
            data.stats = this.computeColumnStats(data.matrix, data.headers);
            data.formattedText = this.matrixToMarkdownTable(data.matrix);
          }
        });
      }
    } catch (err) {
      console.warn("ExcelInsightsService getFullSheetData error:", err);
    }
    return data;
  }

  /**
   * Determine target scope ('selection' vs 'sheet') based on user parameter or active selection size.
   */
  static async resolveScope(requestedScope = "auto") {
    if (requestedScope === "selection" || requestedScope === "sheet") {
      return requestedScope;
    }
    // Auto mode: If user highlighted > 1 cell, use selection; otherwise analyze full sheet
    try {
      let isMultiCellSelection = false;
      if (typeof Excel !== 'undefined') {
        await Excel.run(async (context) => {
          const range = context.workbook.getSelectedRange();
          range.load(["rowCount", "columnCount"]);
          await context.sync();
          if (range.rowCount > 1 || range.columnCount > 1) {
            isMultiCellSelection = true;
          }
        });
      }
      return isMultiCellSelection ? "selection" : "sheet";
    } catch (e) {
      return "sheet";
    }
  }

  /**
   * Generate Table Insights prompt context for Gemini
   * Supports 'selection', 'sheet', or 'auto' scope.
   */
  static async buildTableInsightsPrompt(scope = "auto", customInstruction = "") {
    const resolvedScope = await this.resolveScope(scope);
    const data = resolvedScope === "selection" 
      ? await this.getSelectedRangeData() 
      : await this.getFullSheetData();

    const scopeTitle = resolvedScope === "selection"
      ? `Highlighted Cells Range (${data.address})`
      : `Full Active Worksheet "${data.sheetName}" (${data.address})`;

    let prompt = `System Context: You are performing a deep TABLE INSIGHTS ANALYSIS on Microsoft Excel data.\n`;
    prompt += `Target Scope: ${scopeTitle}\n`;
    prompt += `Grid Size: ${data.rowCount} rows × ${data.columnCount} columns\n`;
    if (data.tableNames && data.tableNames.length > 0) {
      prompt += `Excel Structured Tables Detected: ${data.tableNames.join(", ")}\n`;
    }

    if (data.stats && data.stats.length > 0) {
      prompt += `\nCalculated Column Statistics (Pre-parsed):\n`;
      data.stats.forEach(st => {
        if (st.type === "numeric") {
          prompt += `- Column "${st.header}": Sum=${st.sum}, Avg=${st.avg.toFixed(2)}, Min=${st.min}, Max=${st.max}\n`;
        } else {
          prompt += `- Column "${st.header}": ${st.distinctCount} unique values (${st.nonEmptyCount} filled)\n`;
        }
      });
    }

    prompt += `\nRaw Data Matrix (Markdown Table format):\n`;
    // Cap row text to prevent token overflow for massive sheets while preserving top & bottom context
    if (data.matrix.length > 100) {
      const topMatrix = data.matrix.slice(0, 50);
      const bottomMatrix = data.matrix.slice(-20);
      prompt += this.matrixToMarkdownTable(topMatrix);
      prompt += `\n\n... [${data.matrix.length - 70} rows omitted for brevity] ...\n\n`;
      prompt += this.matrixToMarkdownTable(bottomMatrix);
    } else {
      prompt += data.formattedText;
    }

    prompt += `\n\nTask & Required Output Format:`;
    prompt += `\n1. ## 📊 Key Table Patterns & Trends`;
    prompt += `\n   - Provide 3-4 high-impact insights with **bold metric lead-ins** and exact numbers. Format currency cleanly as $1,234 or $500 (never escape with backslashes like \\$).`;
    prompt += `\n2. ## ⚠️ Anomalies, Outliers & Data Risks`;
    prompt += `\n   - Flag negative variances, unexpected outliers, blank cell gaps, or formula risks.`;
    prompt += `\n3. ## 📋 Structured Insights Summary Table`;
    prompt += `\n   - Provide a clean markdown table summarizing Key Metrics / Categories, Numeric Values, and Trend Status (e.g. | Category / Metric | Value | Trend Status |).`;
    prompt += `\n   - CRITICAL: Keep numeric/currency values clean in their own dedicated column (e.g. $1,672,000, 52%, $996,700) without embedding long narrative sentences inside the numeric cell, so that Excel can chart the values directly. Put descriptive commentary or narrative analysis in a separate column or bulleted sections.`;
    prompt += `\n4. Include an executive chart visualization directive if relevant: \`![Table Trend Chart](image: A modern flat 2D financial chart showing...)\``;

    if (customInstruction) {
      prompt += `\n\nAdditional User Request: ${customInstruction}`;
    }

    return prompt;
  }

  /**
   * Converts 0-based column index to Excel column letter (e.g. 0 -> 'A', 1 -> 'B')
   */
  static colIndexToLetter(idx) {
    let letter = "";
    let n = idx;
    while (n >= 0) {
      letter = String.fromCharCode((n % 26) + 65) + letter;
      n = Math.floor(n / 26) - 1;
    }
    return letter || "A";
  }

  /**
   * Converts Excel column letter to 0-based index (e.g. 'A' -> 0, 'B' -> 1)
   */
  static letterToColIndex(letter) {
    const s = String(letter || "").toUpperCase().trim();
    let col = 0;
    for (let i = 0; i < s.length; i++) {
      col = col * 26 + (s.charCodeAt(i) - 64);
    }
    return Math.max(0, col - 1);
  }

  /**
   * Generate Taxonomy Categorization prompt context for Gemini
   * Supports 'selection', 'sheet', or 'auto' scope.
   */
  static async buildTaxonomyPrompt(scope = "auto", customInstruction = "") {
    const resolvedScope = await this.resolveScope(scope);
    const data = resolvedScope === "selection" 
      ? await this.getSelectedRangeData() 
      : await this.getFullSheetData();

    const scopeTitle = resolvedScope === "selection"
      ? `Highlighted Cells Range (${data.address})`
      : `Full Active Worksheet "${data.sheetName}" (${data.address})`;

    let prompt = `System Context: You are performing a TAXONOMY CATEGORIZATION & DATA CLASSIFICATION analysis on Microsoft Excel data.\n`;
    prompt += `Target Scope: ${scopeTitle}\n`;
    prompt += `Grid Size: ${data.rowCount} rows × ${data.columnCount} columns\n`;
    if (data.headers && data.headers.length > 0) {
      prompt += `Columns Detected: ${data.headers.join(", ")}\n`;
    }

    prompt += `\nRaw Data Matrix (Markdown Table format):\n`;
    if (data.matrix.length > 100) {
      const topMatrix = data.matrix.slice(0, 50);
      const bottomMatrix = data.matrix.slice(-20);
      prompt += this.matrixToMarkdownTable(topMatrix);
      prompt += `\n\n... [${data.matrix.length - 70} rows omitted for brevity] ...\n\n`;
      prompt += this.matrixToMarkdownTable(bottomMatrix);
    } else {
      prompt += data.formattedText;
    }

    // Auto-detect best column for categorization
    let targetColHint = "B";
    let targetHeaderHint = "Opportunity Name";
    if (data.headers && data.headers.length > 0) {
      const foundIdx = data.headers.findIndex(h => /opportunity|name|merchant|vendor|description|product|title|item|account|lead\s*source/i.test(h));
      if (foundIdx !== -1) {
        targetColHint = this.colIndexToLetter(foundIdx);
        targetHeaderHint = data.headers[foundIdx];
      }
    }
    const destColHint = this.colIndexToLetter(this.letterToColIndex(targetColHint) + 1);

    prompt += `\n\nTask & Required Output Format:`;
    prompt += `\n1. Analyze each text entry row-by-row in Column ${targetColHint} ("${targetHeaderHint}") across all data rows (Rows 2 to ${Math.max(2, data.rowCount)}) and map each entry into standardized, professional taxonomy categories.`;
    prompt += `\n2. Provide a clean summary markdown table showing the category breakdown, count of rows in each category, and brief analysis.`;
    prompt += `\n3. Return a declarative \`\`\`excel-taxonomy JSON specification code block containing:`;
    prompt += `\n\`\`\`excel-taxonomy`;
    prompt += `\n{`;
    prompt += `\n  "action": "categorize",`;
    prompt += `\n  "sourceColumn": "${targetColHint}",`;
    prompt += `\n  "targetColumn": "${destColHint}",`;
    prompt += `\n  "insertColumn": true,`;
    prompt += `\n  "header": "${targetHeaderHint} Category",`;
    prompt += `\n  "startRow": 2,`;
    prompt += `\n  "categories": [`;
    prompt += `\n    { "row": 2, "sourceValue": "...", "category": "..." }`;
    prompt += `\n  ],`;
    prompt += `\n  "taxonomySummary": {`;
    prompt += `\n    "Category Name": count`;
    prompt += `\n  }`;
    prompt += `\n}`;
    prompt += `\n\`\`\``;

    if (customInstruction) {
      prompt += `\n\nAdditional User Request: ${customInstruction}`;
    }

    return prompt;
  }

  /**
   * Generate Executive Summary prompt context for Gemini
   * Supports 'selection', 'sheet', or 'auto' scope.
   */
  static async buildExecutiveSummaryPrompt(scope = "auto", customInstruction = "") {
    const resolvedScope = await this.resolveScope(scope);
    const data = resolvedScope === "selection" 
      ? await this.getSelectedRangeData() 
      : await this.getFullSheetData();

    const scopeTitle = resolvedScope === "selection"
      ? `Highlighted Cell Selection (${data.address})`
      : `Entire Sheet "${data.sheetName}" (${data.address})`;

    let prompt = `System Context: You are generating an EXECUTIVE SUMMARY for corporate leadership based on Microsoft Excel financial/operational data.\n`;
    prompt += `Target Scope: ${scopeTitle}\n`;
    prompt += `Data Volume: ${data.rowCount} rows, ${data.columnCount} columns\n`;

    if (data.stats && data.stats.length > 0) {
      prompt += `\nPre-calculated Key Metrics:\n`;
      data.stats.forEach(st => {
        if (st.type === "numeric") {
          prompt += `- ${st.header}: Total = ${st.sum.toLocaleString()}, Average = ${st.avg.toFixed(2)} (Range: ${st.min} to ${st.max})\n`;
        }
      });
    }

    prompt += `\nData Matrix Snapshot:\n`;
    if (data.matrix.length > 80) {
      prompt += this.matrixToMarkdownTable(data.matrix.slice(0, 40));
      prompt += `\n\n... [${data.matrix.length - 40} rows summarized in metrics above] ...\n\n`;
    } else {
      prompt += data.formattedText;
    }

    prompt += `\n\nRequired Executive Output Structure:`;
    prompt += `\n1. Executive Callout Box: Start with \`> [!NOTE]\` containing a 2-sentence C-level summary highlighting total business impact and top takeaway.`;
    prompt += `\n2. ## 📈 Performance Highlights`;
    prompt += `\n   - Concise, bulleted executive key metrics with bold numbers (format currency as $500, never \\$500).`;
    prompt += `\n3. ## ⚠️ Strategic Risks & Operational Gaps`;
    prompt += `\n   - 2-3 critical risks or anomalies identified in the data.`;
    prompt += `\n4. ## 🎯 Recommended Leadership Actions`;
    prompt += `\n   - Actionable next steps with Priority (High/Medium) and Expected Impact.`;
    prompt += `\n5. ## 💡 Executive KPI Summary Table`;
    prompt += `\n   - A clean, concise markdown table (3 to 6 key rows max, numeric values in columns, clean unescaped currency): | Metric / Category | Current Value | Target / Benchmark |`;
    prompt += `\n6. Include an executive infographic visual chart directive with exact metric comparisons: \`![Executive Financial Chart](image: A modern 2D corporate financial chart showing [Category A: Value A, Category B: Value B, Category C: Value C]...)\``;

    if (customInstruction) {
      prompt += `\n\nSpecific Executive Directives: ${customInstruction}`;
    }

    return prompt;
  }

  /**
   * Helper: Auto-detect headers from first row of matrix
   */
  static extractHeaders(matrix) {
    if (!matrix || matrix.length === 0) return [];
    const firstRow = matrix[0];
    return firstRow.map((cell, idx) => {
      if (cell !== null && cell !== undefined && String(cell).trim() !== "") {
        return String(cell).trim();
      }
      return `Column ${idx + 1}`;
    });
  }

  /**
   * Helper: Compute basic statistical metrics for columns
   */
  static computeColumnStats(matrix, headers) {
    if (!matrix || matrix.length <= 1) return [];

    const numCols = headers.length;
    const stats = [];

    for (let c = 0; c < numCols; c++) {
      let isNumeric = true;
      let sum = 0;
      let numericCount = 0;
      let min = Infinity;
      let max = -Infinity;
      let nonBlankCount = 0;
      const uniqueVals = new Set();

      // Examine data rows (skipping row 0 header)
      for (let r = 1; r < matrix.length; r++) {
        const val = matrix[r] ? matrix[r][c] : null;
        if (val !== null && val !== undefined && String(val).trim() !== "") {
          nonBlankCount++;
          uniqueVals.add(String(val));
          const num = Number(val);
          if (!isNaN(num)) {
            sum += num;
            numericCount++;
            if (num < min) min = num;
            if (num > max) max = num;
          } else {
            isNumeric = false;
          }
        }
      }

      if (numericCount > 0 && numericCount >= (matrix.length - 1) * 0.5) {
        stats.push({
          header: headers[c] || `Column ${c + 1}`,
          type: "numeric",
          sum: sum,
          avg: numericCount > 0 ? sum / numericCount : 0,
          min: min === Infinity ? 0 : min,
          max: max === -Infinity ? 0 : max,
          count: numericCount
        });
      } else {
        stats.push({
          header: headers[c] || `Column ${c + 1}`,
          type: "categorical",
          distinctCount: uniqueVals.size,
          nonEmptyCount: nonBlankCount
        });
      }
    }

    return stats;
  }

  /**
   * Helper: Convert cell matrix to Markdown table string
   */
  static matrixToMarkdownTable(matrix) {
    if (!matrix || matrix.length === 0) return "";

    const headers = matrix[0].map((cell, idx) => (cell !== null && cell !== undefined && String(cell).trim()) ? String(cell).trim() : `Col ${idx + 1}`);
    let md = "| " + headers.join(" | ") + " |\n";
    md += "| " + headers.map(() => "---").join(" | ") + " |\n";

    for (let r = 1; r < matrix.length; r++) {
      const row = matrix[r] || [];
      const formattedRow = headers.map((_, colIdx) => {
        const val = row[colIdx];
        if (val === null || val === undefined) return "";
        return String(val).replace(/\|/g, "\\|").replace(/\n/g, " ");
      });
      md += "| " + formattedRow.join(" | ") + " |\n";
    }

    return md;
  }

  /**
   * Insert generated Executive Summary as a new native Excel worksheet
   */
  static async insertExecutiveSummaryWorksheet(summaryTitle = "Executive Summary", textContent = "") {
    try {
      if (typeof Excel !== 'undefined') {
        return await Excel.run(async (context) => {
          const sheets = context.workbook.worksheets;
          sheets.load("items/name");
          await context.sync();

          // Generate a clean, unique worksheet name (max 31 chars, valid Excel characters)
          const baseName = summaryTitle.replace(/[:\\/?*\[\]]/g, "").substring(0, 20).trim() || "Summary";
          let sheetName = baseName;
          let counter = 1;
          const existingNames = new Set(sheets.items.map(s => s.name.toLowerCase()));
          while (existingNames.has(sheetName.toLowerCase())) {
            counter++;
            sheetName = `${baseName} ${counter}`.substring(0, 31);
          }

          const summarySheet = sheets.add(sheetName);
          summarySheet.activate();

          // 1. Title Header Banner in A1
          const titleCell = summarySheet.getRange("A1");
          titleCell.values = [[`📊 ${summaryTitle.toUpperCase()} — GEMINI INTELLIGENCE`]];
          titleCell.format.font.bold = true;
          titleCell.format.font.name = "Segoe UI";
          titleCell.format.font.size = 13;
          titleCell.format.font.color = "#004E8C";

          // 2. Subtitle in A2
          const subtitleCell = summarySheet.getRange("A2");
          subtitleCell.values = [[`Generated on ${new Date().toLocaleString()} | Powered by Gemini`]];
          subtitleCell.format.font.italic = true;
          subtitleCell.format.font.name = "Segoe UI";
          subtitleCell.format.font.size = 9;
          subtitleCell.format.font.color = "#666666";

          let currentRow = 3;
          let summaryTableStartRow = null;
          let summaryTableRows = 0;
          let summaryTableCols = 0;
          let lastNormalizedInfo = null;

          // 3. Parse and format markdown content into structured Excel cells
          const rawLines = (textContent || "").split(/\r?\n/);
          let i = 0;

          while (i < rawLines.length) {
            const line = rawLines[i].trim();

            if (!line) {
              i++;
              continue;
            }

            // Skip raw chart directives, markdown image syntax, HTML tags, or base64 data
            if (/^#*\s*(?:📈\s*)?(?:Executive\s+)?Chart\s+Visualization\s+Directive/i.test(line) ||
                /^(?:📈\s*)?Executive Chart Visualization/i.test(line) ||
                line.startsWith("![") ||
                line.includes("image:") ||
                line.includes("image-prompt:") ||
                line.includes("imagen:") ||
                line.includes("data:image/") ||
                line.includes("base64,") ||
                line.includes("<img") ||
                line.includes("<div") ||
                line.includes("</div>") ||
                line.includes("<table") ||
                line.includes("</table>") ||
                line.includes("<tr") ||
                line.includes("<td") ||
                line.includes("office-visual-")) {
              i++;
              continue;
            }

            // A. Markdown Headers (#, ##, ###, ####)
            if (line.startsWith("#")) {
              const headerText = this.cleanCellText(line.replace(/^#+\s*/, ""));
              if (/^(?:📈\s*)?(?:Executive\s+)?Chart\s+Visualization\s+Directive/i.test(headerText)) {
                i++;
                continue;
              }
              const hCell = summarySheet.getRangeByIndexes(currentRow, 0, 1, 6);
              hCell.values = [[headerText, "", "", "", "", ""]];
              hCell.format.font.bold = true;
              hCell.format.font.name = "Segoe UI";
              hCell.format.font.size = 11;
              hCell.format.font.color = "#0078D4";
              hCell.format.fill.color = "#F0F4F8";
              currentRow += 1;
              i++;
              continue;
            }

            // B. Markdown Table (| ... |)
            if (line.startsWith("|") && line.endsWith("|")) {
              const tableRows = [];
              while (i < rawLines.length && rawLines[i].trim().startsWith("|") && rawLines[i].trim().endsWith("|")) {
                const tLine = rawLines[i].trim();
                // Skip markdown separator row (|---|---|)
                if (!/^\|[\s\-:|]+\|$/.test(tLine)) {
                  const cells = tLine.slice(1, -1).split("|").map(c => this.cleanCellText(c));
                  if (cells.length > 0) {
                    tableRows.push(cells);
                  }
                }
                i++;
              }

              if (tableRows.length > 0) {
                const normalizedInfo = this.normalizeInsightsTableData(tableRows);
                const normalizedMatrix = (normalizedInfo && normalizedInfo.cleanedMatrix && normalizedInfo.cleanedMatrix.length > 0)
                  ? normalizedInfo.cleanedMatrix
                  : tableRows;
                const maxCols = Math.max(...normalizedMatrix.map(r => r.length));

                summaryTableStartRow = currentRow;
                summaryTableRows = normalizedMatrix.length;
                summaryTableCols = maxCols;
                lastNormalizedInfo = normalizedInfo;

                const tRange = summarySheet.getRangeByIndexes(currentRow, 0, normalizedMatrix.length, maxCols);
                tRange.values = normalizedMatrix;

                // Style table header row
                const thRange = summarySheet.getRangeByIndexes(currentRow, 0, 1, maxCols);
                thRange.format.font.bold = true;
                thRange.format.font.name = "Segoe UI";
                thRange.format.font.size = 10;
                thRange.format.fill.color = "#0078D4";
                thRange.format.font.color = "#FFFFFF";

                // Style table body rows
                if (normalizedMatrix.length > 1) {
                  const tdRange = summarySheet.getRangeByIndexes(currentRow + 1, 0, normalizedMatrix.length - 1, maxCols);
                  tdRange.format.font.name = "Segoe UI";
                  tdRange.format.font.size = 10;
                  tdRange.format.fill.color = "#FAFAFA";

                  // Apply column number formats if available
                  if (normalizedInfo && normalizedInfo.columnFormats) {
                    normalizedInfo.columnFormats.forEach((fmt, colIdx) => {
                      if (fmt && colIdx < maxCols) {
                        try {
                          const colRange = summarySheet.getRangeByIndexes(currentRow + 1, colIdx, normalizedMatrix.length - 1, 1);
                          colRange.numberFormat = fmt;
                        } catch (_) {}
                      }
                    });
                  }
                }

                currentRow += normalizedMatrix.length + 1;
              }
              continue;
            }

            // C. Bullet points (*, -, 1., 2.)
            if (/^[\*\-]\s+/.test(line) || /^\d+\.\s+/.test(line)) {
              const cleanBullet = this.cleanCellText(line.replace(/^[\*\-]\s+/, "• "));
              const bCell = summarySheet.getRangeByIndexes(currentRow, 0, 1, 1);
              bCell.values = [[cleanBullet]];
              bCell.format.font.name = "Segoe UI";
              bCell.format.font.size = 10;
              currentRow += 1;
              i++;
              continue;
            }

            // D. Standard text paragraph
            const cleanText = this.cleanCellText(line);
            if (cleanText) {
              const pCell = summarySheet.getRangeByIndexes(currentRow, 0, 1, 1);
              pCell.values = [[cleanText]];
              pCell.format.font.name = "Segoe UI";
              pCell.format.font.size = 10;
              currentRow += 1;
            }
            i++;
          }

          // 4. Auto-create Native Excel Interactive Chart on the summary tab if table data is available
          if (summaryTableStartRow !== null && summaryTableRows >= 2) {
            try {
              const norm = lastNormalizedInfo;
              let selectedChartType = Excel.ChartType.columnClustered;
              let chartCols = summaryTableCols;
              let chartTitleText = `${summaryTitle} Visual Overview`;

              if (norm) {
                if (norm.numericColumns && norm.numericColumns.length > 0) {
                  let consecutiveNumCols = 0;
                  for (let c = 1; c < summaryTableCols; c++) {
                    if (norm.numericColumns.includes(c)) consecutiveNumCols++;
                    else break;
                  }
                  chartCols = 1 + Math.max(1, consecutiveNumCols);
                }
                if (norm.chartType === "pie") {
                  selectedChartType = Excel.ChartType.pie;
                } else if (norm.chartType === "line") {
                  selectedChartType = Excel.ChartType.lineMarkers || Excel.ChartType.line;
                } else if (norm.chartType === "barClustered") {
                  selectedChartType = Excel.ChartType.barClustered;
                }
              }

              let chartTargetRange = null;
              if (norm && norm.chartRangeInfo && norm.chartRangeInfo.isSubSection) {
                const subRow = summaryTableStartRow + norm.chartRangeInfo.startRowOffset;
                chartTargetRange = summarySheet.getRangeByIndexes(
                  subRow,
                  norm.chartRangeInfo.startColOffset,
                  norm.chartRangeInfo.rowCount,
                  norm.chartRangeInfo.colCount
                );
                if (norm.chartRangeInfo.title) chartTitleText = norm.chartRangeInfo.title;
              } else {
                chartTargetRange = summarySheet.getRangeByIndexes(
                  summaryTableStartRow,
                  0,
                  summaryTableRows,
                  Math.min(chartCols, summaryTableCols)
                );
              }

              chartTargetRange.load(["left", "top", "width", "height"]);
              await context.sync();

              const chart = summarySheet.charts.add(selectedChartType, chartTargetRange, Excel.ChartSeriesBy.columns);
              chart.name = "Gemini_Summary_Chart_" + Date.now();
              chart.title.text = chartTitleText;
              chart.title.format.font.name = "Segoe UI";
              chart.title.format.font.size = 12;
              chart.title.format.font.bold = true;
              chart.legend.position = (selectedChartType === Excel.ChartType.pie)
                ? Excel.ChartLegendPosition.right
                : Excel.ChartLegendPosition.top;
              chart.legend.visible = (selectedChartType === Excel.ChartType.pie);
              chart.dataLabels.showValue = true;
              chart.format.fill.setSolidColor("#FFFFFF");

              // Position chart in-line, one row below the summary table
              const chartRow = summaryTableStartRow + summaryTableRows + 1;
              const chartPosRange = summarySheet.getRangeByIndexes(chartRow, 0, 1, 1);
              chartPosRange.load(["left", "top"]);
              await context.sync();

              chart.left = (chartPosRange.left !== undefined) ? chartPosRange.left : 10;
              chart.top = (chartPosRange.top !== undefined) ? (chartPosRange.top + 5) : (chartRow * 20 + 20);
              chart.width = 540;
              chart.height = 300;
            } catch (chartErr) {
              console.warn("Notice: could not auto-create native chart on summary tab:", chartErr);
            }
          }

          // 5. If response contains a generated base64 image, insert the image shape onto summary sheet
          const base64Match = (textContent || "").match(/data:image\/[^;]+;base64,([A-Za-z0-9+/=]+)/);
          if (base64Match && base64Match[1]) {
            try {
              const cleanBase64 = base64Match[1].replace(/[\r\n\s]+/g, "").trim();
              const shape = summarySheet.shapes.addImage(cleanBase64);
              shape.name = "Gemini_Visual_Graph_" + Date.now();
              const chartWasCreated = (summaryTableStartRow !== null && summaryTableRows >= 2);
              const imgRow = chartWasCreated ? (summaryTableStartRow + summaryTableRows + 18) : (currentRow + 1);
              const imgPosRange = summarySheet.getRangeByIndexes(imgRow, 0, 1, 1);
              imgPosRange.load(["left", "top"]);
              await context.sync();

              shape.left = (imgPosRange.left !== undefined) ? imgPosRange.left : 10;
              shape.top = (imgPosRange.top !== undefined) ? (imgPosRange.top + 5) : (chartWasCreated ? 400 : 80);
              shape.width = 540;
              shape.height = 300;
            } catch (imgErr) {
              console.warn("Notice: could not insert image shape on summary tab:", imgErr);
            }
          }

          // Auto-fit columns across the populated range
          summarySheet.getRange(`A1:H${Math.max(currentRow + 2, 20)}`).format.autofitColumns();
          await context.sync();
          return true;
        });
      }
    } catch (err) {
      console.error("Error creating executive summary worksheet:", err);
      throw err;
    }
    return false;
  }

  /**
   * Insert Base64 Image into active Excel Worksheet at user's selected cell location
   */
  static async insertVisualImage(base64DataUri) {
    if (!base64DataUri) return false;

    // Clean data URI prefix and whitespace
    const cleanBase64 = base64DataUri.replace(/^data:image\/[^;]+;base64,/i, "").replace(/[\r\n\s]+/g, "").trim();

    try {
      if (typeof Excel !== 'undefined') {
        return await Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          const range = context.workbook.getSelectedRange();
          range.load(["left", "top", "width", "height", "rowIndex", "columnIndex"]);
          await context.sync();

          const shape = sheet.shapes.addImage(cleanBase64);
          shape.name = "Gemini_Visual_Graph_" + Date.now();

          // Position directly at active user selection
          const leftPos = (range.left !== undefined) ? range.left : 20;
          const topPos = (range.top !== undefined) ? range.top : 20;

          shape.left = leftPos;
          shape.top = topPos;
          shape.width = 520;
          shape.height = 290;

          await context.sync();
          return true;
        });
      }
    } catch (err) {
      console.error("Excel insertVisualImage error:", err);
      throw err;
    }
    return false;
  }

  /**
   * Helper: Parse string numbers with units like "$94.5B", "24.7M", "15.2%", "1,200", "$1,672,000 total..."
   */
  static parseNumberWithUnit(val) {
    if (typeof val === "number") return { num: val, rawNum: val, unit: "", isCurrency: false, hasDecimals: (val % 1 !== 0) };
    if (!val || typeof val !== "string") return null;
    const clean = this.cleanCellText(val);
    if (!clean) return null;

    // 1. Strict full string match: $1,234.50, €500, 15.2%, 1,200, 94.5B, -$50.25, ($100), 50.5M
    let match = clean.match(/^([\$€£¥]?)\s*([+-]?[0-9,]+(?:\.[0-9]+)?)\s*([BMKbmk%]|\bBillion\b|\bMillion\b|\bThousand\b)?(?:\s*\(.*?\))?$/i);
    
    // 2. Accounting negative parentheses: ($1,234.50) or (15.2%)
    if (!match) {
      match = clean.match(/^\(\s*([\$€£¥]?)\s*([0-9,]+(?:\.[0-9]+)?)\s*([BMKbmk%]|\bBillion\b|\bMillion\b|\bThousand\b)?\s*\)$/i);
      if (match) {
        const currencyPrefix = match[1] || "";
        const cleanNumStr = match[2].replace(/,/g, "");
        let rawNum = -parseFloat(cleanNumStr);
        let unit = (match[3] || "").toUpperCase();
        if (unit === "BILLION") unit = "B";
        if (unit === "MILLION") unit = "M";
        if (unit === "THOUSAND") unit = "K";
        if (!['B', 'M', 'K', '%'].includes(unit)) unit = "";
        if (!isNaN(rawNum)) {
          return {
            num: rawNum,
            rawNum: rawNum,
            unit: unit,
            isCurrency: !!currencyPrefix || /[\$€£¥]/.test(clean),
            hasDecimals: cleanNumStr.includes(".")
          };
        }
      }
    }

    // 3. String starting with explicit currency (e.g. "$1,672,000 total pipeline across 14 deals")
    if (!match) {
      match = clean.match(/^([\$€£¥])\s*([0-9,]+(?:\.[0-9]+)?)\s*([BMKbmk%]|\bBillion\b|\bMillion\b|\bThousand\b)?/i);
    }

    // 4. Entity prefix with currency in parens (e.g. "Elena Rostova ($562,000 across 4 opportunities)")
    if (!match) {
      match = clean.match(/^[A-Za-z\s]{2,25}\s*\(\s*([\$€£¥])\s*([0-9,]+(?:\.[0-9]+)?)\s*([BMKbmk%]|\bBillion\b|\bMillion\b|\bThousand\b)?/i);
    }

    if (match && match[2]) {
      const currencyPrefix = match[1] || "";
      const cleanNumStr = match[2].replace(/,/g, "");
      const rawNum = parseFloat(cleanNumStr);
      let unit = (match[3] || "").toUpperCase();
      if (unit === "BILLION") unit = "B";
      if (unit === "MILLION") unit = "M";
      if (unit === "THOUSAND") unit = "K";
      if (!['B', 'M', 'K', '%'].includes(unit)) unit = "";
      
      if (!isNaN(rawNum)) {
        return { 
          num: rawNum, 
          rawNum: rawNum, 
          unit: unit, 
          isCurrency: !!currencyPrefix || /[\$€£¥]/.test(clean),
          hasDecimals: cleanNumStr.includes(".")
        };
      }
    }
    return null;
  }

  /**
   * Helper: Extract concise, workable chart dataset (3-6 items) directly from AI response or visual image directive
   */
  static extractChartDataset(responseText) {
    if (!responseText) return null;

    // 1. Try to extract from the visual image directive: ![...](image: ...) or ```image ... ```
    const imgDirectiveMatch = responseText.match(/!\[.*?\]\(image:\s*([\s\S]*?)\)/i) || 
                             responseText.match(/```(?:image|imagen|visual)\s*([\s\S]*?)```/i);
    
    if (imgDirectiveMatch && imgDirectiveMatch[1]) {
      const imgPrompt = imgDirectiveMatch[1];
      const pairRegex = /(?:comparing|showing|of)?\s*([A-Za-z0-9\s&/]{2,25}?)(?:\s*[:\-=]|\s*\()\s*[\$€£]?([0-9]+(?:\.[0-9]+)?)\s*([BMKbmk%]?)\)?/g;
      const extracted = [];
      let m;
      while ((m = pairRegex.exec(imgPrompt)) !== null) {
        const rawLabel = this.cleanCellText(m[1]).replace(/^(and|vs|the|with|for)\s+/i, "");
        const num = parseFloat(m[2]);
        const unit = (m[3] || "").toUpperCase();
        if (rawLabel.length >= 2 && !isNaN(num) && num > 0) {
          if (!extracted.some(e => e.label.toLowerCase() === rawLabel.toLowerCase())) {
            extracted.push({
              label: rawLabel,
              value: num,
              unit: unit ? ` (${unit})` : ""
            });
          }
        }
      }

      if (extracted.length >= 2 && extracted.length <= 8) {
        const unitLabel = extracted[0].unit || "";
        const header = ["Category", `Value${unitLabel}`];
        const rows = extracted.map(e => [e.label, e.value]);
        return {
          matrix: [header, ...rows],
          chartType: (extracted.length <= 4) ? "pie" : "columnClustered",
          title: "Executive Visual Breakdown"
        };
      }
    }

    // 2. Parse Markdown Table, filtering to 1 Label Column + 1-2 Pure Numeric Columns
    const rawTable = this.parseMarkdownTableFromResponse(responseText);
    if (rawTable && rawTable.length >= 2) {
      const headers = rawTable[0].map(h => this.cleanCellText(h));
      const dataRows = rawTable.slice(1);

      // Find columns that contain valid numbers (including $94.5B, 15%, etc.)
      const numericColIndices = [];
      const colUnits = {};

      for (let c = 1; c < headers.length; c++) {
        let numericCount = 0;
        let detectedUnit = "";
        for (let r = 0; r < dataRows.length; r++) {
          const val = dataRows[r][c];
          const parsed = this.parseNumberWithUnit(val);
          if (parsed !== null) {
            numericCount++;
            if (parsed.unit && ['B', 'M', 'K', '%'].includes(parsed.unit) && !detectedUnit) {
              detectedUnit = parsed.unit;
            }
          }
        }
        if (numericCount >= dataRows.length * 0.4) {
          numericColIndices.push(c);
          colUnits[c] = detectedUnit ? ` (${detectedUnit})` : "";
        }
      }

      if (numericColIndices.length > 0) {
        const selectedNumericCols = numericColIndices.slice(0, 2);
        const cleanHeader = [
          headers[0],
          ...selectedNumericCols.map(idx => `${headers[idx]}${colUnits[idx] || ""}`)
        ];
        
        const cleanRows = dataRows.slice(0, 8).map(row => {
          let cat = this.cleanCellText(row[0]).substring(0, 28).trim();
          const nums = selectedNumericCols.map(idx => {
            const v = row[idx];
            const parsed = this.parseNumberWithUnit(v);
            if (parsed !== null) {
              const entityMatch = String(v).match(/^([A-Za-z\s]{3,20})\s*[\(\$]/);
              if (entityMatch && !cat.toLowerCase().includes(entityMatch[1].trim().toLowerCase())) {
                cat = `${cat} (${entityMatch[1].trim()})`.substring(0, 32);
              }
              return parsed.num;
            }
            return 0;
          });
          return [cat, ...nums];
        });

        let chartType = "columnClustered";
        if (selectedNumericCols.length === 1 && cleanRows.length <= 4) {
          chartType = "pie";
        } else if (cleanRows.some(r => /Q[1-4]|202\d|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec/i.test(String(r[0])))) {
          chartType = "line";
        }

        return {
          matrix: [cleanHeader, ...cleanRows],
          chartType: chartType,
          title: "Executive Key Metrics"
        };
      }
    }

    // 3. Parse bullet points with metrics (e.g. "* **Search:** $54.2B", "* **Cloud ($24.7B):**...")
    const bulletRegex = /^\s*[\*\-]\s+(?:\*\*)?([A-Za-z0-9\s&/]{2,25}?)(?:\*\*)?(?:\s*\(?[\$€£]?([0-9]+(?:\.[0-9]+)?)\s*([BMKbmk%]?)\)?|\s*[:\-=]\s*[\$€£]?([0-9]+(?:\.[0-9]+)?)\s*([BMKbmk%]?))/gm;
    const bulletItems = [];
    let bm;
    while ((bm = bulletRegex.exec(responseText)) !== null) {
      const rawLabel = bm[1].trim().replace(/^(and|vs|the|with|for)\s+/i, "");
      const numStr = bm[2] || bm[4];
      const unitStr = (bm[3] || bm[5] || "").toUpperCase();
      const num = parseFloat(numStr);
      if (rawLabel.length >= 2 && !isNaN(num) && num > 0) {
        if (!bulletItems.some(b => b.label.toLowerCase() === rawLabel.toLowerCase())) {
          bulletItems.push({
            label: rawLabel,
            value: num,
            unit: unitStr ? ` (${unitStr})` : ""
          });
        }
      }
    }

    if (bulletItems.length >= 2 && bulletItems.length <= 8) {
      const unitLabel = bulletItems[0].unit || "";
      const header = ["Metric / Segment", `Value${unitLabel}`];
      const rows = bulletItems.map(b => [b.label, b.value]);
      return {
        matrix: [header, ...rows],
        chartType: (bulletItems.length <= 4) ? "pie" : "columnClustered",
        title: "Executive Key Metrics"
      };
    }

    return null;
  }

  /**
   * Helper: Parse raw Markdown table lines from response
   */
  static parseMarkdownTableFromResponse(text) {
    if (!text) return null;
    const lines = text.split("\n");
    const rows = [];
    let inTable = false;

    for (const line of lines) {
      const trimmed = line.trim();
      if (trimmed.startsWith("|") && trimmed.endsWith("|")) {
        // Skip markdown separator row (|---|---|)
        if (/^\|[\s\-:|]+\|$/.test(trimmed)) {
          inTable = true;
          continue;
        }
        const cells = trimmed
          .slice(1, -1)
          .split("|")
          .map(c => this.cleanCellText(c));

        if (cells.length >= 2) {
          rows.push(cells);
        }
      } else if (inTable && rows.length >= 2) {
        break;
      }
    }

    return rows.length >= 2 ? rows : null;
  }

  /**
   * Helper: Normalize raw table matrix into numeric data types, number formats, and chart properties
   * Ensures numbers, currencies, and percentages are genuine JavaScript numbers so Excel can chart them.
   * @param {Array<Array<string>>} rawMatrix 
   * @returns {Object} { headers, cleanedMatrix, columnFormats, numericColumns, chartType }
   */
  /**
   * Helper: Cleans text in a table cell, removing markdown formatting and backslash escapes
   * @param {*} val 
   * @returns {string}
   */
  static cleanCellText(val) {
    if (val === null || val === undefined) return "";
    let s = String(val).trim();
    // Strip bold/italic markdown markers
    s = s.replace(/\*\*/g, "").replace(/\*/g, "");
    // Unescape backslash before special characters, math/currency symbols, or markdown formatting
    // Handles \$500 -> $500, \%10 -> %10, \_foo\_ -> _foo_, \# -> #, etc.
    s = s.replace(/\\+([$¥€£%_*`~#\-\(\)\[\]{}<>.,:;!+=\/\\|'"])/g, "$1");
    // Strip backslash right before digits or currency (e.g. \500 -> 500, \$500 -> $500)
    s = s.replace(/\\+(\$|[0-9])/g, "$1");
    // Strip leading or trailing backslashes
    s = s.replace(/^\\+|\\+$/g, "");
    // Strip isolated backslashes followed by whitespace or string boundaries
    s = s.replace(/\\+(?=\s|$)/g, "");
    return s.trim();
  }

  /**
   * Cleans and converts raw markdown table values into typed JavaScript data structures
   * (numbers, currencies, percentages, strings) and auto-detects optimal number formats
   * and intelligent chart sub-ranges (e.g. Category Performance).
   * @param {Array<Array<string>>} rawMatrix - Extracted table matrix from markdown
   * @returns {Object} { headers, cleanedMatrix, columnFormats, numericColumns, chartType, chartRangeInfo, sectionRows }
   */
  static normalizeInsightsTableData(rawMatrix) {
    if (!rawMatrix || rawMatrix.length === 0) {
      return {
        headers: [],
        cleanedMatrix: [],
        columnFormats: [],
        numericColumns: [],
        chartType: "columnClustered",
        chartRangeInfo: null,
        sectionRows: []
      };
    }

    const numCols = Math.max(...rawMatrix.map(r => r.length));
    const paddedMatrix = rawMatrix.map(r => {
      const row = r.map(c => this.cleanCellText(c));
      while (row.length < numCols) row.push("");
      return row;
    });

    const headers = paddedMatrix[0].map((h, idx) => {
      const clean = this.cleanCellText(h);
      return clean || `Column ${idx + 1}`;
    });

    const dataRows = paddedMatrix.slice(1);
    const columnFormats = []; // array of number format strings or null
    const numericColumns = []; // column indices that are numeric

    // Identify section subheader rows in matrix (e.g., "| Key Metrics |", "| Category Performance |", "| Trend Status |")
    // A row in dataRows is a section subheader if only 1 non-empty cell or explicit section header name
    const sectionRows = []; // 1-indexed row indices in cleanedMatrix
    for (let r = 0; r < dataRows.length; r++) {
      const nonEmpty = dataRows[r].filter(c => c.trim() !== "");
      if (nonEmpty.length === 1 && numCols > 1) {
        sectionRows.push(r + 1); // +1 because row 0 is headers
      } else if (dataRows[r][0] && /^(?:key metrics|category performance|spending distribution|trend status|category breakdown|performance breakdown|summary metrics)/i.test(dataRows[r][0]) && (!dataRows[r][1] || isNaN(parseFloat(dataRows[r][1].replace(/[\$,€£¥%\s]/g, ""))))) {
        sectionRows.push(r + 1);
      }
    }

    // Analyze each column for numeric / currency / percentage patterns
    for (let c = 0; c < numCols; c++) {
      let numericCount = 0;
      let currencyCount = 0;
      let percentCount = 0;
      let hasDecimals = false;
      let nonBlankCount = 0;

      for (let r = 0; r < dataRows.length; r++) {
        if (sectionRows.includes(r + 1)) continue; // skip section subheaders in column type inference
        const rawVal = this.cleanCellText(dataRows[r][c]);
        if (!rawVal) continue;
        nonBlankCount++;

        const parsed = this.parseNumberWithUnit(rawVal);
        if (parsed !== null) {
          numericCount++;
          if (parsed.isCurrency) currencyCount++;
          if (parsed.unit === "%") percentCount++;
          if (parsed.hasDecimals) hasDecimals = true;
        }
      }

      if (nonBlankCount > 0 && numericCount >= nonBlankCount * 0.4) {
        numericColumns.push(c);
        if (currencyCount >= numericCount * 0.3) {
          columnFormats[c] = hasDecimals ? "$#,##0.00" : "$#,##0";
        } else if (percentCount >= numericCount * 0.3) {
          columnFormats[c] = "0.0%";
        } else {
          columnFormats[c] = hasDecimals ? "#,##0.00" : "#,##0";
        }
      } else {
        columnFormats[c] = null;
      }
    }

    // Convert data rows: convert numeric cells to genuine JS numbers
    const cleanedRows = dataRows.map((row, rIdx) => {
      const isSubheader = sectionRows.includes(rIdx + 1);
      let enrichedCategory = null;
      const mappedRow = row.map((cell, colIdx) => {
        const cleanStr = this.cleanCellText(cell);
        if (!isSubheader && cleanStr !== "") {
          const parsed = this.parseNumberWithUnit(cleanStr);
          if (parsed !== null && numericColumns.includes(colIdx)) {
            // If there was an entity name in the numeric cell, enrich the category in Col 0
            if (colIdx > 0 && typeof cleanStr === "string") {
              const entityMatch = cleanStr.match(/^([A-Za-z\s]{3,24})\s*[\(\$]/);
              if (entityMatch && entityMatch[1]) {
                const entity = entityMatch[1].trim();
                const currentCat = this.cleanCellText(row[0]);
                if (currentCat && !currentCat.toLowerCase().includes(entity.toLowerCase())) {
                  enrichedCategory = `${currentCat} (${entity})`;
                }
              }
            }
            let finalVal = parsed.num;
            if (parsed.unit === "%" || columnFormats[colIdx] === "0.0%") {
              finalVal = finalVal / 100;
            }
            return finalVal;
          }
        }
        return cleanStr;
      });
      if (enrichedCategory && mappedRow.length > 0) {
        mappedRow[0] = enrichedCategory;
      }
      return mappedRow;
    });

    const cleanedMatrix = [headers, ...cleanedRows];

    // Sub-section analysis for intelligent interactive charting
    // Look for sections like "Category Performance", "Breakdown", "Distribution", or groups of rows with category name + numeric amount
    let chartRangeInfo = null;

    if (sectionRows.length > 0) {
      // Find each section slice
      const sections = [];
      for (let s = 0; s < sectionRows.length; s++) {
        const subheaderRowIdx = sectionRows[s]; // 1-indexed in cleanedMatrix
        const sectionTitle = cleanedMatrix[subheaderRowIdx][0] || `Section ${s + 1}`;
        const nextSubheaderRowIdx = (s + 1 < sectionRows.length) ? sectionRows[s + 1] : cleanedMatrix.length;
        const startDataRow = subheaderRowIdx + 1;
        const endDataRow = nextSubheaderRowIdx; // exclusive
        const rowCount = endDataRow - startDataRow;

        if (rowCount > 0) {
          const secRows = cleanedMatrix.slice(startDataRow, endDataRow);
          // Check if secRows are homogeneous categorical numeric data (e.g. Travel: $1054.75, Cloud: $300, etc.)
          let secNumericCount = 0;
          for (const sRow of secRows) {
            const val = sRow[1];
            if (typeof val === "number" || (typeof val === "string" && !isNaN(parseFloat(val.replace(/[\$,€£¥%\s]/g, ""))))) {
              secNumericCount++;
            }
          }
          const isCategoryDistribution = secNumericCount >= Math.min(rowCount, 2) && rowCount >= 2;
          sections.push({
            title: sectionTitle,
            startRowOffset: startDataRow, // 0-indexed relative to matrix top
            rowCount,
            isCategoryDistribution,
            secNumericCount
          });
        }
      }

      // Prioritize "Category Performance", "Breakdown", "Distribution", or section with highest categorical numeric count
      const categorySection = sections.find(sec => /category|performance|breakdown|distribution|expense|revenue|segment/i.test(sec.title) && sec.isCategoryDistribution)
        || sections.find(sec => sec.isCategoryDistribution);

      if (categorySection) {
        const cType = categorySection.rowCount <= 4 ? "barClustered" : "barClustered";
        chartRangeInfo = {
          isSubSection: true,
          startRowOffset: categorySection.startRowOffset,
          rowCount: categorySection.rowCount,
          startColOffset: 0,
          colCount: 2,
          title: `${categorySection.title} Breakdown`,
          chartType: cType
        };
      }
    }

    // Default chart type determination for whole table if not sub-section
    let defaultChartType = "columnClustered";
    const firstColValues = dataRows.map(r => String(r[0] || "").toLowerCase());
    const isTimeline = firstColValues.some(v => /q[1-4]|202\d|jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec|week|month|year/i.test(v));
    if (isTimeline) {
      defaultChartType = "line";
    } else if (numericColumns.length === 1 && dataRows.length <= 4) {
      defaultChartType = "pie";
    } else if (dataRows.length > 6) {
      defaultChartType = "barClustered";
    }

    if (!chartRangeInfo) {
      let chartCols = 2;
      if (numericColumns.length > 0) {
        let consecutiveNumCols = 0;
        for (let c = 1; c < numCols; c++) {
          if (numericColumns.includes(c)) consecutiveNumCols++;
          else break;
        }
        chartCols = 1 + Math.max(1, consecutiveNumCols);
      }
      chartRangeInfo = {
        isSubSection: false,
        startRowOffset: 0,
        rowCount: cleanedMatrix.length,
        startColOffset: 0,
        colCount: Math.min(chartCols, numCols),
        title: "Executive Insights Summary",
        chartType: defaultChartType
      };
    }

    return {
      headers,
      cleanedMatrix,
      columnFormats,
      numericColumns,
      chartType: chartRangeInfo.chartType,
      chartRangeInfo,
      sectionRows
    };
  }

  /**
   * Robustly extracts Structured Insights Summary Table from markdown or HTML response
   * @param {string} text - Markdown response text
   * @param {string} htmlContent - Optional HTML content
   * @returns {Array<Array<string>>|null}
   */
  static extractInsightsSummaryTable(text = "", htmlContent = "") {
    if (!text && !htmlContent) return null;

    // 1. Check for dedicated Structured Insights Summary Table section in text
    if (text && text.includes("|")) {
      const lines = text.split(/\r?\n/);
      let tableLines = [];
      let inSummarySection = false;

      for (let i = 0; i < lines.length; i++) {
        const line = lines[i].trim();
        if (/^#+\s*.*(?:Structured Insights|Insights Summary|Summary Table|Key Metrics|Table Patterns)/i.test(line)) {
          inSummarySection = true;
          tableLines = [];
          continue;
        }

        if (line.startsWith("|")) {
          tableLines.push(line);
        } else if (tableLines.length >= 2 && !inSummarySection && line !== "") {
          break;
        } else if (tableLines.length >= 2 && inSummarySection && line !== "" && !line.startsWith("|")) {
          if (/^#+|^[-*]\s+|\d+\.\s+/.test(line)) {
            break;
          }
        }
      }

      if (tableLines.length >= 2) {
        const rows = [];
        let maxCols = 1;

        for (const tLine of tableLines) {
          let normalized = tLine;
          if (!normalized.endsWith("|")) normalized += " |";
          const rawCells = normalized.slice(1, -1).split("|");
          const cells = rawCells.map(c => this.cleanCellText(c));

          // Check if this row is a separator row (e.g. | :--- | :--- |)
          const isSeparator = cells.every(c => /^[\s\-:|]+$/.test(c) && c.includes("-"));
          if (isSeparator) continue;

          if (cells.length > 0) {
            maxCols = Math.max(maxCols, cells.length);
            rows.push(cells);
          }
        }

        const paddedRows = rows.map(r => {
          const row = [...r];
          while (row.length < maxCols) row.push("");
          return row;
        });

        if (paddedRows.length >= 2) return paddedRows;
      }

      // Fallback: parse any markdown table
      const fallbackTable = this.parseMarkdownTableFromResponse(text);
      if (fallbackTable && fallbackTable.length >= 2) return fallbackTable;
    }

    // 2. HTML table fallback
    if (htmlContent && htmlContent.includes("<table")) {
      try {
        if (typeof document !== "undefined") {
          const tempDiv = document.createElement("div");
          tempDiv.innerHTML = htmlContent;
          const table = tempDiv.querySelector("table");
          if (table) {
            const rows = [];
            const trs = Array.from(table.querySelectorAll("tr"));
            trs.forEach(tr => {
              const cells = Array.from(tr.querySelectorAll("th, td")).map(td => this.cleanCellText(td.innerText || td.textContent || ""));
              if (cells.length > 0) rows.push(cells);
            });
            if (rows.length >= 2) return rows;
          }
        }
      } catch (e) {
        // Ignore
      }
    }

    return null;
  }

  /**
   * Insert Structured Insights Summary Table directly at the user's currently selected cell location.
   * @param {string} responseText - AI response markdown
   * @param {string} htmlContent - Optional parsed HTML
   */
  static async insertInsightsSummaryTable(responseText = "", htmlContent = "") {
    try {
      if (typeof Excel !== 'undefined') {
        return await Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          const selectedRange = context.workbook.getSelectedRange();
          selectedRange.load(["rowIndex", "columnIndex"]);
          await context.sync();

          const rawTable = this.extractInsightsSummaryTable(responseText, htmlContent);
          if (!rawTable || rawTable.length < 2) {
            throw new Error("No Structured Insights Summary Table found in response.");
          }

          const normalized = this.normalizeInsightsTableData(rawTable);
          const matrix = normalized.cleanedMatrix;
          const numRows = matrix.length;
          const numCols = matrix[0].length;

          const startRow = (selectedRange.rowIndex !== undefined) ? selectedRange.rowIndex : 0;
          const startCol = (selectedRange.columnIndex !== undefined) ? selectedRange.columnIndex : 0;

          // 1. Insert structured table data into sheet
          const targetRange = sheet.getRangeByIndexes(startRow, startCol, numRows, numCols);
          targetRange.values = matrix;

          // 2. Format Table Header
          const headerRange = sheet.getRangeByIndexes(startRow, startCol, 1, numCols);
          headerRange.format.font.bold = true;
          headerRange.format.font.name = "Segoe UI";
          headerRange.format.font.size = 11;
          headerRange.format.fill.color = "#0078D4";
          headerRange.format.font.color = "#FFFFFF";
          headerRange.format.horizontalAlignment = "Center";

          // 3. Format Data Rows & Apply Numeric Formats
          if (numRows > 1) {
            const dataRange = sheet.getRangeByIndexes(startRow + 1, startCol, numRows - 1, numCols);
            dataRange.format.font.name = "Segoe UI";
            dataRange.format.font.size = 10;
            dataRange.format.fill.color = "#FAFAFA";

            // Apply specific number formats column-by-column
            normalized.columnFormats.forEach((fmt, cIdx) => {
              if (fmt) {
                try {
                  const colDataRange = sheet.getRangeByIndexes(startRow + 1, startCol + cIdx, numRows - 1, 1);
                  colDataRange.numberFormat = fmt;
                } catch (fmtErr) {
                  console.warn("Number format assignment notice:", fmtErr);
                }
              }
            });

            // Format Section Subheaders
            if (normalized.sectionRows && normalized.sectionRows.length > 0) {
              normalized.sectionRows.forEach(secRIdx => {
                try {
                  const secRange = sheet.getRangeByIndexes(startRow + secRIdx, startCol, 1, numCols);
                  secRange.format.font.bold = true;
                  secRange.format.font.color = "#004E8C";
                  secRange.format.fill.color = "#F0F4F8";
                } catch (secErr) {}
              });
            }
          }

          targetRange.format.autofitColumns();
          await context.sync();

          return true;
        });
      }
    } catch (err) {
      console.error("Excel insertInsightsSummaryTable error:", err);
      throw err;
    }
    return false;
  }

  /**
   * Insert a live, interactive Native Excel Chart linked to the concise AI visual dataset
   * at the user's selected cell location.
   */
  static async insertNativeInteractiveChart(chartTitle = "Executive Summary Analysis", customType = null, responseText = "") {
    try {
      if (typeof Excel !== 'undefined') {
        return await Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          const selectedRange = context.workbook.getSelectedRange();
          selectedRange.load(["rowIndex", "columnIndex", "left", "top", "width", "height", "rowCount", "columnCount", "values"]);
          await context.sync();

          let targetRange = null;
          let selectedChartType = Excel.ChartType.columnClustered;

          const startRow = (selectedRange.rowIndex !== undefined) ? selectedRange.rowIndex : 0;
          const startCol = (selectedRange.columnIndex !== undefined) ? selectedRange.columnIndex : 0;

          // 1. First priority: Check if response has a Structured Insights Summary Table or markdown table
          const rawTable = this.extractInsightsSummaryTable(responseText);
          if (rawTable && rawTable.length >= 2) {
            const normalized = this.normalizeInsightsTableData(rawTable);
            const matrix = normalized.cleanedMatrix;
            const tableRows = matrix.length;
            const tableCols = matrix[0].length;

            const summaryRange = sheet.getRangeByIndexes(startRow, startCol, tableRows, tableCols);
            summaryRange.values = matrix;

            // Professional header styling
            const headerRange = sheet.getRangeByIndexes(startRow, startCol, 1, tableCols);
            headerRange.format.font.bold = true;
            headerRange.format.font.name = "Segoe UI";
            headerRange.format.font.size = 11;
            headerRange.format.fill.color = "#0078D4";
            headerRange.format.font.color = "#FFFFFF";

            // Clean data cell styling and number formats
            if (tableRows > 1) {
              const dataCellsRange = sheet.getRangeByIndexes(startRow + 1, startCol, tableRows - 1, tableCols);
              dataCellsRange.format.font.name = "Segoe UI";
              dataCellsRange.format.font.size = 10;
              dataCellsRange.format.fill.color = "#F8F9FA";

              normalized.columnFormats.forEach((fmt, cIdx) => {
                if (fmt) {
                  try {
                    const colDataRange = sheet.getRangeByIndexes(startRow + 1, startCol + cIdx, tableRows - 1, 1);
                    colDataRange.numberFormat = fmt;
                  } catch (fmtErr) {}
                }
              });

              // Format Section Subheaders
              if (normalized.sectionRows && normalized.sectionRows.length > 0) {
                normalized.sectionRows.forEach(secRIdx => {
                  try {
                    const secRange = sheet.getRangeByIndexes(startRow + secRIdx, startCol, 1, tableCols);
                    secRange.format.font.bold = true;
                    secRange.format.font.color = "#004E8C";
                    secRange.format.fill.color = "#F0F4F8";
                  } catch (secErr) {}
                });
              }
            }

            summaryRange.format.autofitColumns();
            summaryRange.load(["left", "top", "width", "height"]);
            await context.sync();

            // Set target data range and chart type for the native chart
            if (normalized.chartRangeInfo && normalized.chartRangeInfo.isSubSection) {
              const subStartRow = startRow + normalized.chartRangeInfo.startRowOffset;
              targetRange = sheet.getRangeByIndexes(subStartRow, startCol + normalized.chartRangeInfo.startColOffset, normalized.chartRangeInfo.rowCount, normalized.chartRangeInfo.colCount);
              chartTitle = normalized.chartRangeInfo.title || chartTitle;
            } else {
              let chartCols = 2;
              if (normalized.numericColumns.length > 0) {
                let consecutiveNumCols = 0;
                for (let c = 1; c < tableCols; c++) {
                  if (normalized.numericColumns.includes(c)) consecutiveNumCols++;
                  else break;
                }
                chartCols = 1 + Math.max(1, consecutiveNumCols);
              }
              targetRange = sheet.getRangeByIndexes(startRow, startCol, tableRows, Math.min(chartCols, tableCols));
            }
            targetRange.load(["left", "top", "width", "height"]);

            if (normalized.chartType === "pie") {
              selectedChartType = Excel.ChartType.pie;
            } else if (normalized.chartType === "line") {
              selectedChartType = Excel.ChartType.lineMarkers || Excel.ChartType.line;
            } else if (normalized.chartType === "barClustered") {
              selectedChartType = Excel.ChartType.barClustered;
            } else {
              selectedChartType = Excel.ChartType.columnClustered;
            }
          }

          // 2. Second priority: Extract from visual directive / bullet points
          if (!targetRange) {
            const datasetInfo = this.extractChartDataset(responseText);
            if (datasetInfo && datasetInfo.matrix && datasetInfo.matrix.length >= 2) {
              const matrix = datasetInfo.matrix;
              const tableRows = matrix.length;
              const tableCols = matrix[0].length;

              const summaryRange = sheet.getRangeByIndexes(startRow, startCol, tableRows, tableCols);
              summaryRange.values = matrix;

              const headerRange = sheet.getRangeByIndexes(startRow, startCol, 1, tableCols);
              headerRange.format.font.bold = true;
              headerRange.format.font.name = "Segoe UI";
              headerRange.format.font.size = 11;
              headerRange.format.fill.color = "#0078D4";
              headerRange.format.font.color = "#FFFFFF";

              if (tableRows > 1) {
                const dataCellsRange = sheet.getRangeByIndexes(startRow + 1, startCol, tableRows - 1, tableCols);
                dataCellsRange.format.font.name = "Segoe UI";
                dataCellsRange.format.font.size = 10;
                dataCellsRange.format.fill.color = "#F8F9FA";
              }

              summaryRange.format.autofitColumns();
              summaryRange.load(["left", "top", "width", "height"]);
              await context.sync();

              targetRange = summaryRange;

              if (datasetInfo.chartType === "pie") {
                selectedChartType = Excel.ChartType.pie;
              } else if (datasetInfo.chartType === "line") {
                selectedChartType = Excel.ChartType.lineMarkers || Excel.ChartType.line;
              } else if (datasetInfo.chartType === "barClustered") {
                selectedChartType = Excel.ChartType.barClustered;
              } else {
                selectedChartType = Excel.ChartType.columnClustered;
              }

              if (datasetInfo.title) {
                chartTitle = datasetInfo.title;
              }
            }
          }

          // 3. Fallback: If no AI dataset, use safe data bounds from the active worksheet
          if (!targetRange) {
            if (selectedRange.rowCount > 1 && selectedRange.rowCount < 200 && selectedRange.columnCount > 1 && selectedRange.columnCount < 20) {
              targetRange = selectedRange;
            } else {
              const tables = sheet.tables;
              tables.load("items");
              await context.sync();

              if (tables.items && tables.items.length > 0) {
                targetRange = tables.items[0].getRange();
              } else {
                const usedRange = sheet.getUsedRange();
                usedRange.load(["rowCount", "columnCount", "left", "top", "width", "height", "values", "address", "rowIndex", "columnIndex"]);
                await context.sync();
                if (usedRange.rowCount > 1 && usedRange.columnCount > 1) {
                  const safeRows = Math.min(usedRange.rowCount, 12);
                  const safeCols = Math.min(usedRange.columnCount, 3);
                  targetRange = sheet.getRangeByIndexes(usedRange.rowIndex || 0, usedRange.columnIndex || 0, safeRows, safeCols);
                }
              }
            }
          }

          if (!targetRange) {
            throw new Error("Could not locate concise data to chart.");
          }

          targetRange.load(["left", "top", "width", "height", "values"]);
          await context.sync();

          if (customType) {
            selectedChartType = customType;
          }

          // Create native Excel interactive chart
          const chart = sheet.charts.add(selectedChartType, targetRange, Excel.ChartSeriesBy.columns);
          chart.name = "Gemini_Interactive_Chart_" + Date.now();
          chart.title.text = chartTitle;
          chart.title.format.font.name = "Segoe UI";
          chart.title.format.font.size = 13;
          chart.title.format.font.bold = true;

          chart.legend.position = (selectedChartType === Excel.ChartType.pie)
            ? Excel.ChartLegendPosition.right
            : Excel.ChartLegendPosition.top;
          chart.legend.visible = (selectedChartType === Excel.ChartType.pie);
          
          chart.dataLabels.showValue = true;
          chart.format.fill.setSolidColor("#FFFFFF");

          // Position chart in-line, one row below the data table
          let leftPos = (targetRange.left !== undefined) ? targetRange.left : 10;
          let topPos = 200;
          try {
            const chartPlacementRow = startRow + (rawTable ? rawTable.length : (targetRange.rowCount || 10)) + 1;
            const belowRange = sheet.getRangeByIndexes(chartPlacementRow, startCol, 1, 1);
            belowRange.load(["left", "top"]);
            await context.sync();
            if (belowRange.left !== undefined) leftPos = belowRange.left;
            if (belowRange.top !== undefined) topPos = belowRange.top + 5;
          } catch (_) {
            if (targetRange.top !== undefined && targetRange.height !== undefined) {
              topPos = targetRange.top + targetRange.height + 20;
            }
          }

          chart.left = leftPos;
          chart.top = topPos;
          chart.width = 540;
          chart.height = 300;

          await context.sync();
          return true;
        });
      }
    } catch (err) {
      console.error("Excel insertNativeInteractiveChart error:", err);
      throw err;
    }
    return false;
  }
}
