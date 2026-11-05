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
 * Gemini for Microsoft Excel - Response Actions Toolbar
 */

import { TaskpaneShell } from '../../shell/TaskpaneShell.js';
import { extractExcelFormula, extractExcelChartSpec, extractExcelChartEditSpec, extractExcelRangeFormatSpec, extractExcelTaxonomySpec, extractInsightsSummaryTable } from '../../adapters/ExcelAdapter.js';
import { copyRichContentToClipboard } from '../../core/clipboardUtils.js';

/**
 * Renders response action buttons for Excel.
 * @param {Object} ctx - Response actions context
 * @returns {HTMLElement}
 */
export function renderResponseActions(ctx) {
  const { rawText, formattedHtml, textDiv, adapter, chatLog, prompt = "", userPrompt = "", chip = null } = ctx;

  const primaryActions = document.createElement("div");
  primaryActions.className = "primary-actions";

  // Check intent contexts
  const isInsightsContext = Boolean(
    ctx.isInsights ||
    (chip && chip.isInsights) ||
    /insights|table patterns|anomalies.*outliers/i.test(userPrompt) ||
    /insights|table patterns|anomalies.*outliers/i.test(prompt) ||
    /###\s*(?:Structured\s+)?Insights\s+Summary\s+Table|###\s*Key\s+Table\s+Patterns|###\s*Anomalies,\s*Outliers/i.test(rawText)
  );

  const isTaxonomyContext = Boolean(
    ctx.isTaxonomy ||
    (chip && (chip.id === "chipTaxonomy" || chip.id === "selTaxonomy")) ||
    /taxonomy|categoriz|categories\s+column/i.test(userPrompt) ||
    /taxonomy|categoriz|categories\s+column/i.test(prompt) ||
    /```(?:excel-taxonomy|excel-categorize|taxonomy|categorize)/i.test(rawText)
  );

  // 1. Chart Edit has top precedence when user asks to update/modify a chart or chart edit spec is returned
  const detectedChartEdit = (!isInsightsContext && !isTaxonomyContext)
    ? extractExcelChartEditSpec(rawText, formattedHtml, { userPrompt, prompt })
    : null;

  // 2. Chart Creation
  const detectedChart = (!detectedChartEdit && !isInsightsContext && !isTaxonomyContext)
    ? extractExcelChartSpec(rawText, formattedHtml)
    : null;

  // 3. Formula: Extract formula FIRST before range formatting
  const detectedFormula = (!detectedChartEdit && !detectedChart && !isTaxonomyContext && !isInsightsContext)
    ? extractExcelFormula(rawText, formattedHtml)
    : null;

  // 4. Taxonomy Categorization (only if taxonomy context, never when chart/insights/formula)
  const detectedTaxonomy = (isTaxonomyContext && !isInsightsContext && !detectedChartEdit && !detectedChart && !detectedFormula)
    ? extractExcelTaxonomySpec(rawText, formattedHtml, { isTaxonomy: true, userPrompt, prompt })
    : null;

  // 5. Insights Table (if in insights context and not chart/taxonomy/formula)
  const detectedInsightsTable = (isInsightsContext && !isTaxonomyContext && !detectedChartEdit && !detectedChart && !detectedFormula && !detectedTaxonomy)
    ? extractInsightsSummaryTable(rawText, formattedHtml)
    : null;

  // 6. Range Formatting: ONLY when NOT chart, NOT formula, NOT taxonomy, NOT insights
  const detectedRangeFormat = (!detectedChartEdit && !detectedChart && !detectedFormula && !detectedTaxonomy && !isInsightsContext)
    ? extractExcelRangeFormatSpec(rawText, formattedHtml)
    : null;

  const insertBtn = document.createElement("button");

  if (detectedChartEdit) {
    insertBtn.className = "action-btn chart-edit";
    insertBtn.innerHTML = `🔄 Update Chart`;
    insertBtn.title = `Apply changes to chart (${detectedChartEdit.targetChart || detectedChartEdit.chartId || 'active'})`;
    insertBtn.onclick = async () => {
      try {
        insertBtn.disabled = true;
        insertBtn.innerHTML = `⚡ Updating Chart...`;
        if (adapter && typeof adapter.editChart === 'function') {
          await adapter.editChart(detectedChartEdit, formattedHtml);
          insertBtn.innerHTML = `✅ Chart Updated!`;
        } else {
          const htmlToInsert = textDiv ? textDiv.innerHTML : "";
          await TaskpaneShell.performDocumentInsertion(adapter, htmlToInsert, rawText, "chart_edit", "Excel");
          insertBtn.innerHTML = `✅ Chart Updated!`;
        }
      } catch (err) {
        console.error("Update Chart Error:", err);
        insertBtn.innerHTML = `❌ Failed`;
      } finally {
        setTimeout(() => {
          insertBtn.disabled = false;
          insertBtn.innerHTML = `🔄 Update Chart`;
        }, 3000);
      }
    };
    primaryActions.appendChild(insertBtn);
    const pill = document.createElement("span");
    pill.className = "chart-edit-preview-pill";
    const editSummary = detectedChartEdit.chartType ? ` ➔ ${detectedChartEdit.chartType}` : (detectedChartEdit.title ? ` ➔ "${detectedChartEdit.title}"` : '');
    pill.innerText = `🔄 Edit ${detectedChartEdit.targetChart || 'Chart'}${editSummary}`;
    pill.title = `Update existing chart in-place`;
    primaryActions.appendChild(pill);
  } else if (detectedChart) {
    insertBtn.className = "action-btn chart";
    insertBtn.innerHTML = `📊 Insert Native Chart`;
    insertBtn.title = `Insert live interactive Excel chart (${detectedChart.type}) into worksheet`;
    insertBtn.onclick = async () => {
      try {
        insertBtn.disabled = true;
        insertBtn.innerHTML = `⚡ Inserting Chart...`;
        const htmlToInsert = textDiv ? textDiv.innerHTML : "";
        await TaskpaneShell.performDocumentInsertion(adapter, htmlToInsert, rawText, "chart", "Excel");
        insertBtn.innerHTML = `✅ Chart Inserted!`;
      } catch (err) {
        console.error("Insert Native Chart Error:", err);
        insertBtn.innerHTML = `❌ Failed`;
      } finally {
        setTimeout(() => {
          insertBtn.disabled = false;
          insertBtn.innerHTML = `📊 Insert Native Chart`;
        }, 3000);
      }
    };
    primaryActions.appendChild(insertBtn);

    const pill = document.createElement("span");
    pill.className = "chart-preview-pill";
    const rangeLabel = detectedChart.dataRange || (detectedChart.categoriesRange ? `${detectedChart.categoriesRange} / ${detectedChart.seriesRange}` : 'Auto Range');
    pill.innerText = `📊 ${detectedChart.type}: ${rangeLabel}`;
    pill.title = `Native Chart: ${detectedChart.type}`;
    primaryActions.appendChild(pill);
  } else if (isTaxonomyContext || detectedTaxonomy) {
    insertBtn.className = "action-btn taxonomy";
    insertBtn.innerHTML = `🏷️ Insert Categories Column`;
    insertBtn.title = `Insert Column ${detectedTaxonomy ? detectedTaxonomy.targetColumn : 'C'} adjacent to selected data and populate taxonomy categories`;
    insertBtn.onclick = async () => {
      try {
        insertBtn.disabled = true;
        insertBtn.innerHTML = `⚡ Inserting Column...`;
        if (adapter && typeof adapter.insertTaxonomyColumn === 'function') {
          await adapter.insertTaxonomyColumn(rawText, formattedHtml, { userPrompt, prompt, isTaxonomy: true, spec: detectedTaxonomy });
          insertBtn.innerHTML = `✅ Column Inserted!`;
        } else {
          const htmlToInsert = textDiv ? textDiv.innerHTML : "";
          await TaskpaneShell.performDocumentInsertion(adapter, htmlToInsert, rawText, "taxonomy", "Excel");
          insertBtn.innerHTML = `✅ Inserted!`;
        }
      } catch (err) {
        console.error("Insert Taxonomy Column Error:", err);
        insertBtn.innerHTML = `❌ Failed`;
      } finally {
        setTimeout(() => {
          insertBtn.disabled = false;
          insertBtn.innerHTML = `🏷️ Insert Categories Column`;
        }, 3000);
      }
    };
    primaryActions.appendChild(insertBtn);

    const pill = document.createElement("span");
    pill.className = "taxonomy-preview-pill";
    if (detectedTaxonomy) {
      pill.innerText = `🏷️ Col ${detectedTaxonomy.sourceColumn} ➔ Col ${detectedTaxonomy.targetColumn} (${detectedTaxonomy.count || ''} rows)`;
      pill.title = `Taxonomy: ${detectedTaxonomy.header} (${detectedTaxonomy.count || ''} items)`;
    } else {
      pill.innerText = `🏷️ Categories Column`;
      pill.title = `Taxonomy Categorization`;
    }
    primaryActions.appendChild(pill);
  } else if (detectedInsightsTable) {
    insertBtn.className = "action-btn insert";
    insertBtn.innerHTML = `📊 Insert Insights Summary Table`;
    insertBtn.title = `Extract and insert Structured Insights Summary Table at selected cell position`;
    insertBtn.onclick = async () => {
      try {
        insertBtn.disabled = true;
        insertBtn.innerHTML = `⚡ Inserting Table...`;
        if (adapter && typeof adapter.insertInsightsSummaryTable === 'function') {
          await adapter.insertInsightsSummaryTable(rawText, formattedHtml);
          insertBtn.innerHTML = `✅ Table Inserted!`;
        } else {
          await TaskpaneShell.performDocumentInsertion(adapter, formattedHtml, rawText, "insights_table", "Excel");
          insertBtn.innerHTML = `✅ Inserted!`;
        }
      } catch (err) {
        console.error("Insert Insights Table Error:", err);
        insertBtn.innerHTML = `❌ Failed`;
      } finally {
        setTimeout(() => {
          insertBtn.disabled = false;
          insertBtn.innerHTML = `📊 Insert Insights Summary Table`;
        }, 3000);
      }
    };
    primaryActions.appendChild(insertBtn);

    const pill = document.createElement("span");
    pill.className = "taxonomy-preview-pill";
    const rowCount = detectedInsightsTable.length - 1;
    const colCount = detectedInsightsTable[0].length;
    pill.innerText = `📊 Summary Table (${rowCount} rows × ${colCount} cols)`;
    pill.title = `Structured Insights Summary Table (${rowCount} rows × ${colCount} columns)`;
    primaryActions.appendChild(pill);
  } else if (detectedFormula) {
    insertBtn.className = "action-btn formula";
    insertBtn.innerHTML = `➕ Insert Formula`;
    insertBtn.title = `Insert evaluated formula ${detectedFormula} into active cell`;
    insertBtn.onclick = async () => {
      const htmlToInsert = textDiv ? textDiv.innerHTML : "";
      await TaskpaneShell.performDocumentInsertion(adapter, htmlToInsert, rawText, "formula", "Excel");
    };
    primaryActions.appendChild(insertBtn);

    const pill = document.createElement("span");
    pill.className = "formula-preview-pill";
    pill.innerText = detectedFormula;
    pill.title = `Formula: ${detectedFormula}`;
    primaryActions.appendChild(pill);
  } else if (detectedRangeFormat) {
    insertBtn.className = "action-btn range-format";
    insertBtn.innerHTML = `🎨 Apply Range Format`;
    insertBtn.title = `Apply formatting styles to range (${detectedRangeFormat.targetRange || 'active selection'})`;
    insertBtn.onclick = async () => {
      const htmlToInsert = textDiv ? textDiv.innerHTML : "";
      await TaskpaneShell.performDocumentInsertion(adapter, htmlToInsert, rawText, "range_format", "Excel");
    };
    primaryActions.appendChild(insertBtn);

    const pill = document.createElement("span");
    pill.className = "range-format-preview-pill";
    const formatDetails = [];
    if (detectedRangeFormat.fill && detectedRangeFormat.fill.color) formatDetails.push(`Fill: ${detectedRangeFormat.fill.color}`);
    if (detectedRangeFormat.font && detectedRangeFormat.font.bold) formatDetails.push("Bold");
    if (detectedRangeFormat.numberFormat) formatDetails.push(`Num: ${detectedRangeFormat.numberFormat}`);
    if (detectedRangeFormat.autofitColumns) formatDetails.push("Autofit");
    const summaryStr = formatDetails.length > 0 ? ` (${formatDetails.join(", ")})` : "";
    pill.innerText = `🎨 Format: ${detectedRangeFormat.targetRange || 'Selection'}${summaryStr}`;
    pill.title = `Formatting properties for ${detectedRangeFormat.targetRange || 'selection'}`;
    primaryActions.appendChild(pill);
  } else {
    insertBtn.className = "action-btn insert";
    insertBtn.innerHTML = `➕ Insert into Sheet`;
    insertBtn.title = "Insert tabular data or text into active spreadsheet";
    insertBtn.onclick = async () => {
      const htmlToInsert = textDiv ? textDiv.innerHTML : "";
      await TaskpaneShell.performDocumentInsertion(adapter, htmlToInsert, rawText, "smart", "Excel");
    };
    primaryActions.appendChild(insertBtn);
  }

  // Excel Specific: Create Summary Tab Button (Insights / Executive Summary)
  if (!isTaxonomyContext && !detectedFormula && (isInsightsContext || (!detectedChart && !detectedChartEdit && !detectedRangeFormat && rawText.includes("#") && rawText.length > 150))) {
    const createSheetBtn = document.createElement("button");
    createSheetBtn.className = "action-btn outlined insert";
    createSheetBtn.innerHTML = `📊 New Summary Tab`;
    createSheetBtn.title = "Insert generated summary into a new Excel worksheet tab";
    createSheetBtn.onclick = async () => {
      try {
        createSheetBtn.disabled = true;
        createSheetBtn.innerHTML = `⚡ Creating Tab...`;
        if (adapter && typeof adapter.insertExecutiveSummaryWorksheet === 'function') {
          await adapter.insertExecutiveSummaryWorksheet("Executive Summary", rawText);
          createSheetBtn.innerHTML = `✅ Tab Created!`;
        }
      } catch (err) {
        console.error("Create Summary Tab Error:", err);
        createSheetBtn.innerHTML = `❌ Failed`;
      } finally {
        setTimeout(() => {
          createSheetBtn.disabled = false;
          createSheetBtn.innerHTML = `📊 New Summary Tab`;
        }, 3000);
      }
    };
    primaryActions.appendChild(createSheetBtn);
  }

  // Excel Specific: Fallback Interactive Chart for tabular/insights text when no declarative chart block is present
  if (!isTaxonomyContext && !detectedFormula && !detectedChart && !detectedChartEdit && !detectedRangeFormat && (isInsightsContext || (rawText.includes("|") && rawText.length > 200 && /chart|graph|distribution|breakdown/i.test(rawText)))) {
    const insertNativeChartBtn = document.createElement("button");
    insertNativeChartBtn.className = "action-btn outlined insert";
    insertNativeChartBtn.innerHTML = `📊 Insert Interactive Chart`;
    insertNativeChartBtn.title = "Insert summary table and interactive Excel chart into worksheet";
    insertNativeChartBtn.onclick = async () => {
      try {
        insertNativeChartBtn.disabled = true;
        insertNativeChartBtn.innerHTML = `⚡ Creating Table & Chart...`;
        if (adapter && typeof adapter.insertNativeInteractiveChart === 'function') {
          await adapter.insertNativeInteractiveChart("Executive Analysis Chart", null, rawText);
          insertNativeChartBtn.innerHTML = `✅ Table & Chart Inserted!`;
        }
      } catch (err) {
        console.error("Insert Native Chart Error:", err);
        insertNativeChartBtn.innerHTML = `❌ Failed`;
      } finally {
        setTimeout(() => {
          insertNativeChartBtn.disabled = false;
          insertNativeChartBtn.innerHTML = `📊 Insert Interactive Chart`;
        }, 3000);
      }
    };
    primaryActions.appendChild(insertNativeChartBtn);
  }

  // Excel Specific: AI Image insertion if an image is rendered in the response
  if (textDiv) {
    const imgElement = textDiv.querySelector ? textDiv.querySelector("img") : null;
    if (imgElement && imgElement.src) {
      const insertImgBtn = document.createElement("button");
      insertImgBtn.className = "action-btn outlined insert";
      insertImgBtn.innerHTML = `🖼️ Insert AI Image`;
      insertImgBtn.title = "Insert AI visual infographic image into worksheet";
      insertImgBtn.onclick = async () => {
        try {
          insertImgBtn.disabled = true;
          insertImgBtn.innerHTML = `⚡ Inserting...`;
          if (adapter && typeof adapter.insertVisualChart === 'function') {
            await adapter.insertVisualChart(imgElement.src);
            insertImgBtn.innerHTML = `✅ Image Inserted!`;
          }
        } catch (err) {
          console.error("Insert AI Image Error:", err);
          insertImgBtn.innerHTML = `❌ Failed`;
        } finally {
          setTimeout(() => {
            insertImgBtn.disabled = false;
            insertImgBtn.innerHTML = `🖼️ Insert AI Image`;
          }, 3000);
        }
      };
      primaryActions.appendChild(insertImgBtn);
    }
  }

  // If Executive Summary or briefing is detected, provide dedicated Summary Tab button
  const isExecutiveSummary = !detectedFormula && /(?:Executive Summary|Performance Highlights|Leadership Actions|> \[!NOTE\])/i.test(rawText);
  if (isExecutiveSummary && adapter && typeof adapter.insertExecutiveSummaryWorksheet === 'function') {
    const summaryTabBtn = document.createElement("button");
    summaryTabBtn.className = "action-btn summary-tab";
    summaryTabBtn.innerHTML = `📑 New Summary Tab`;
    summaryTabBtn.title = "Create a dedicated Executive Summary worksheet tab with this briefing";
    summaryTabBtn.onclick = async () => {
      try {
        await adapter.insertExecutiveSummaryWorksheet("Executive Summary", rawText);
        if (chatLog) chatLog.appendSystem("✓ Created new 'Executive Summary' worksheet tab.");
      } catch (err) {
        if (chatLog) chatLog.appendSystem("Error creating summary tab: " + err.message);
      }
    };
    primaryActions.appendChild(summaryTabBtn);
  }

  // Copy to Clipboard Button (omitted when host taskpane provides a dedicated footer copy icon)
  if (!ctx.hideCopyButton) {
    if (chatLog && typeof chatLog.createCopyButton === 'function') {
      const copyBtn = chatLog.createCopyButton(rawText, textDiv);
      primaryActions.appendChild(copyBtn);
    } else {
      const copyBtn = document.createElement("button");
      copyBtn.className = "action-btn copy";
      copyBtn.innerHTML = `<svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor" style="vertical-align: -2px; margin-right: 4px;"><path d="M16 1H4c-1.1 0-2 .9-2 2v14h2V3h12V1zm3 4H8c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h11c1.1 0 2-.9 2-2V7c0-1.1-.9-2-2-2zm0 16H8V7h11v14z"/></svg>Copy to Clipboard`;
      copyBtn.title = "Copy rich formatted content to clipboard";
      copyBtn.onclick = async () => {
        try {
          const success = await copyRichContentToClipboard(rawText, textDiv);
          if (success) {
            const origHtml = copyBtn.innerHTML;
            copyBtn.innerHTML = `✓ Copied!`;
            copyBtn.style.backgroundColor = "#e8f0fe";
            copyBtn.style.color = "#0b57d0";
            copyBtn.style.borderColor = "#0b57d0";
            setTimeout(() => {
              copyBtn.innerHTML = origHtml;
              copyBtn.style.backgroundColor = "";
              copyBtn.style.color = "";
              copyBtn.style.borderColor = "";
            }, 2000);
          }
        } catch (err) {
          console.warn("Clipboard copy notice:", err);
        }
      };
      primaryActions.appendChild(copyBtn);
    }
  }

  return primaryActions;
}
