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
 * Gemini for Microsoft Excel - Host Module Descriptor
 */

/* global Excel */

import { ExcelAdapter } from '../../adapters/ExcelAdapter.js';
import { docChips, selectionChips } from './chips.js';
import { renderResponseActions } from './responseActions.js';

export default {
  hostName: "Excel",

  strings: {
    welcome: "Type a prompt below, request formulas or native charts, or select cells/charts in Excel!",
    readyStatus: "Excel Ready",
    toolsTitle: "📊 <strong>Spreadsheet Tools & Analysis:</strong>"
  },

  createAdapter: () => new ExcelAdapter(),

  docChips,
  selectionChips,
  renderResponseActions,

  _lockedChart: null,
  _lastSelectionSig: null,
  _defaultPlaceholder: "Ask Gemini anything, request formulas, or type instructions...",

  /**
   * Updates user guidance in the input placeholder based on whether a target chart is in context.
   */
  updatePromptGuidance() {
    const promptInput = document.getElementById("promptText");
    if (!promptInput) return;

    if (this._lockedChart) {
      promptInput.placeholder = `Target chart '${this._lockedChart.title}' included. Explicitly reference the chart in your prompt to modify it...`;
    } else {
      promptInput.placeholder = this._defaultPlaceholder;
    }
  },

  /**
   * Reset session hook called when "New Chat" is clicked.
   * Clears the target chart object while preserving the active target range.
   */
  onSessionReset(ctx) {
    this._lockedChart = null;
    this._lastSelectionSig = null;
    const adapter = (ctx && ctx.adapter) ? ctx.adapter : (this._adapter = this._adapter || new ExcelAdapter());
    if (adapter && typeof adapter.setLockedChart === 'function') {
      adapter.setLockedChart(null);
    }

    const targetChartBadge = document.getElementById("targetChartBadge");
    if (targetChartBadge) {
      targetChartBadge.style.display = "none";
    }
    const targetChartRow = document.getElementById("targetChartRow");
    if (targetChartRow) {
      targetChartRow.style.display = "none";
    }

    const chatHistory = document.getElementById("chatHistory");
    if (chatHistory) {
      chatHistory.scrollTop = 0;
    }
    if (ctx && ctx.chatLog && typeof ctx.chatLog.scrollToTop === 'function') {
      ctx.chatLog.scrollToTop();
    }

    this.updatePromptGuidance();
    this.onSelectionChange("", ctx).catch(() => {});
  },

  /**
   * Selection change hook to update target range badge (Green), target chart badge (Blue),
   * and schema indicator pill.
   */
  async onSelectionChange(selectionText, ctx) {
    const targetRangeBadge = document.getElementById("targetRangeBadge") || document.getElementById("targetCellBadge");
    const targetChartBadge = document.getElementById("targetChartBadge");
    const targetChartText = document.getElementById("targetChartText");
    const targetChartRow = document.getElementById("targetChartRow");
    const schemaBadge = document.getElementById("sheetSchemaBadge");
    if (!targetRangeBadge) return;

    const adapter = (ctx && ctx.adapter) ? ctx.adapter : (this._adapter = this._adapter || new ExcelAdapter());

    try {
      if (typeof Excel !== 'undefined' && Excel.run) {
        await Excel.run(async (context) => {
          const targetInfo = await adapter.getSelectionTargetInfo(context, null, this._lockedChart);

          // 1. Update Green Target Range Chip (Always Active)
          if (targetInfo.type === "table" && targetInfo.table) {
            const tbl = targetInfo.table;
            targetRangeBadge.innerText = `📑 Target Table: ${tbl.name} (${targetInfo.cleanAddress})`;
            targetRangeBadge.title = `Active Table: ${tbl.name} (${targetInfo.rowCount} rows x ${targetInfo.columnCount} cols)`;
            targetRangeBadge.className = "excel-pill target-table table-selected";
            targetRangeBadge.dataset.selectionType = "table";
            targetRangeBadge.dataset.tableName = tbl.name;
            targetRangeBadge.dataset.target = targetInfo.cleanAddress;
          } else if (targetInfo.type === "range") {
            targetRangeBadge.innerText = `📊 Target Range: ${targetInfo.cleanAddress}`;
            targetRangeBadge.title = `Active Range: ${targetInfo.cleanAddress} (${targetInfo.rowCount} rows x ${targetInfo.columnCount} cols)`;
            targetRangeBadge.className = "excel-pill target-range";
            targetRangeBadge.dataset.selectionType = "range";
            targetRangeBadge.dataset.range = targetInfo.cleanAddress;
            targetRangeBadge.dataset.target = targetInfo.cleanAddress;
          } else {
            targetRangeBadge.innerText = `🎯 Target Cell: ${targetInfo.cleanAddress}`;
            targetRangeBadge.title = `Cursor Cell: ${targetInfo.cleanAddress}`;
            targetRangeBadge.className = "excel-pill target-cell";
            targetRangeBadge.dataset.selectionType = "cell";
            targetRangeBadge.dataset.cell = targetInfo.cleanAddress;
            targetRangeBadge.dataset.target = targetInfo.cleanAddress;
          }

          // 2. Update Blue Target Chart Chip (Sticky / Dual Target)
          if (targetInfo.chart) {
            this._lockedChart = targetInfo.chart;
            const ch = targetInfo.chart;
            if (targetChartBadge) {
              targetChartBadge.style.display = "inline-flex";
              if (targetChartText) {
                targetChartText.innerText = `📊 Chart: ${ch.title} (${ch.chartType})`;
              }
              targetChartBadge.title = `Active Chart Target: ${ch.name} (ID: ${ch.id}). Explicitly reference this chart in your prompt to modify it, or click 'x' to remove from context.`;
              targetChartBadge.dataset.chartId = ch.id;
              targetChartBadge.dataset.chartName = ch.name;
              targetChartBadge.dataset.chartType = ch.chartType;
            }
            if (targetChartRow) {
              targetChartRow.style.display = "flex";
            }
          } else if (this._lockedChart) {
            if (targetChartBadge) {
              targetChartBadge.style.display = "inline-flex";
              if (targetChartText) {
                targetChartText.innerText = `📊 Chart: ${this._lockedChart.title} (${this._lockedChart.chartType})`;
              }
              targetChartBadge.title = `Active Chart Target: ${this._lockedChart.name} (ID: ${this._lockedChart.id}). Explicitly reference this chart in your prompt to modify it, or click 'x' to remove from context.`;
            }
            if (targetChartRow) {
              targetChartRow.style.display = "flex";
            }
          } else {
            if (targetChartBadge) {
              targetChartBadge.style.display = "none";
            }
            if (targetChartRow) {
              targetChartRow.style.display = "none";
            }
          }

          this.updatePromptGuidance();

          // 3. Update Schema Pill
          if (schemaBadge) {
            schemaBadge.innerText = `📄 ${targetInfo.sheetName} (${targetInfo.rowCount}x${targetInfo.columnCount})`;
          }

          // 4. Notify selection toolbar listener when selected cell/range or its text content changes
          const selectionSig = `${targetInfo.sheetName}!${targetInfo.cleanAddress}::${targetInfo.rawText || ""}`;
          const isInitialBoot = this._lastSelectionSig === null;
          if (this._lastSelectionSig !== selectionSig) {
            this._lastSelectionSig = selectionSig;
            if (ctx && typeof ctx.onSelectionChanged === "function") {
              // On initial boot, a single default cursor cell (e.g. A1) is not an explicit user selection
              const effectiveText = (isInitialBoot && (targetInfo.cellCount || 1) <= 1)
                ? ""
                : (targetInfo.rawText || "");
              ctx.onSelectionChanged(effectiveText);
            }
          }
        });
      }
    } catch (err) {
      console.debug("Excel selection preview lookup notice:", err);
    }
  },

  /**
   * Registers native Excel event handlers for chart clicks, worksheet selections, and tab switches
   */
  async registerEventHandlers(ctx) {
    if (this._handlersRegistered) return;
    this._handlersRegistered = true;
    try {
      if (typeof Excel !== 'undefined' && Excel.run) {
        await Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();

          // 1. Chart collection activation / deactivation events (ExcelApi 1.8+)
          try {
            if (sheet.charts && sheet.charts.onActivated) {
              sheet.charts.onActivated.add(async () => {
                // When a chart is clicked, inspect and lock onto it
                try {
                  await Excel.run(async (c) => {
                    const activeChart = c.workbook.getActiveChartOrNullObject();
                    activeChart.load(["id", "name", "chartType", "width", "height", "top", "left", "isNullObject"]);
                    await c.sync();
                    if (!activeChart.isNullObject) {
                      let chartTitle = activeChart.name || "Chart";
                      try {
                        if (activeChart.title) {
                          activeChart.title.load(["text", "visible"]);
                          await c.sync();
                          if (activeChart.title.visible && activeChart.title.text) {
                            chartTitle = activeChart.title.text;
                          }
                        }
                      } catch (tErr) {}

                      this._lockedChart = {
                        id: activeChart.id,
                        name: activeChart.name,
                        chartType: activeChart.chartType,
                        title: chartTitle,
                        width: Math.round(activeChart.width || 0),
                        height: Math.round(activeChart.height || 0),
                        top: Math.round(activeChart.top || 0),
                        left: Math.round(activeChart.left || 0)
                      };
                      if (ctx && ctx.adapter) {
                        ctx.adapter.setLockedChart(this._lockedChart);
                      }
                    }
                  });
                } catch (e) {}

                await this.onSelectionChange("", ctx);
              });
            }
            if (sheet.charts && sheet.charts.onDeactivated) {
              sheet.charts.onDeactivated.add(async () => {
                // Retain sticky chart focus when focus moves off chart (e.g. clicking side panel)
              });
            }
          } catch (cErr) {
            console.debug("Chart event registration notice:", cErr);
          }

          // 2. Worksheet selection change events (ExcelApi 1.7+)
          try {
            if (sheet.onSelectionChanged) {
              sheet.onSelectionChanged.add(async () => {
                // When range selection changes on the worksheet, check if a new chart was clicked:
                try {
                  await Excel.run(async (c) => {
                    let activeChart = null;
                    try {
                      activeChart = c.workbook.getActiveChartOrNullObject();
                      activeChart.load(["id", "name", "chartType", "width", "height", "top", "left", "isNullObject"]);
                      await c.sync();
                    } catch (e) {}

                    if (activeChart && !activeChart.isNullObject) {
                      let chartTitle = activeChart.name || "Chart";
                      try {
                        if (activeChart.title) {
                          activeChart.title.load(["text", "visible"]);
                          await c.sync();
                          if (activeChart.title.visible && activeChart.title.text) {
                            chartTitle = activeChart.title.text;
                          }
                        }
                      } catch (tErr) {}

                      this._lockedChart = {
                        id: activeChart.id,
                        name: activeChart.name,
                        chartType: activeChart.chartType,
                        title: chartTitle,
                        width: Math.round(activeChart.width || 0),
                        height: Math.round(activeChart.height || 0),
                        top: Math.round(activeChart.top || 0),
                        left: Math.round(activeChart.left || 0)
                      };
                      if (ctx && ctx.adapter) {
                        ctx.adapter.setLockedChart(this._lockedChart);
                      }
                    }
                    // If no chart is active on cell change, DO NOT clear the target chart!
                    // Target chart persists alongside target range until explicitly removed or new chart selected.
                  });
                } catch (sErr) {}

                await this.onSelectionChange("", ctx);
              });
            }
          } catch (sErr) {
            console.debug("Sheet selection event registration notice:", sErr);
          }

          // 3. Worksheet tab activation event (ExcelApi 1.7+)
          try {
            if (context.workbook.worksheets && context.workbook.worksheets.onActivated) {
              context.workbook.worksheets.onActivated.add(async () => {
                this._lockedChart = null; // Reset chart cache on sheet switch
                if (ctx && ctx.adapter) {
                  ctx.adapter.setLockedChart(null);
                }
                await this.onSelectionChange("", ctx);
              });
            }
          } catch (wErr) {
            console.debug("Worksheet activation event registration notice:", wErr);
          }

          await context.sync();
        });
      }
    } catch (err) {
      console.debug("Excel event handler registration notice:", err);
    }
  },

  /**
   * Lifecycle ready hook to initialize worksheet indicators, event handlers, and polling
   */
  async onReady(ctx) {
    this._adapter = (ctx && ctx.adapter) ? ctx.adapter : new ExcelAdapter();

    // Wire Remove Chart Close Button
    const removeTargetChartBtn = document.getElementById("removeTargetChartBtn");
    if (removeTargetChartBtn) {
      removeTargetChartBtn.onclick = (e) => {
        if (e) e.stopPropagation();
        this._lockedChart = null;
        if (ctx && ctx.adapter) {
          ctx.adapter.setLockedChart(null);
        }
        const targetChartBadge = document.getElementById("targetChartBadge");
        if (targetChartBadge) {
          targetChartBadge.style.display = "none";
        }
        const targetChartRow = document.getElementById("targetChartRow");
        if (targetChartRow) {
          targetChartRow.style.display = "none";
        }
        this.updatePromptGuidance();
        this.onSelectionChange("", ctx).catch(() => {});
      };
    }

    // Wire New Chat Button explicitly if present
    const clearSessionBtn = document.getElementById("clearSession");
    if (clearSessionBtn) {
      clearSessionBtn.addEventListener("click", () => {
        this.onSessionReset(ctx);
      });
    }

    await this.onSelectionChange("", ctx);
    await this.registerEventHandlers(ctx);

    // Polling interval (1000ms) ensures instant feedback when charts are clicked
    if (this._selectionPoll) {
      clearInterval(this._selectionPoll);
    }
    this._selectionPoll = setInterval(() => {
      if (typeof Excel !== 'undefined' && Excel.run) {
        Excel.run(async (context) => {
          let activeChart = null;
          try {
            activeChart = context.workbook.getActiveChartOrNullObject();
            activeChart.load(["id", "name", "chartType", "width", "height", "top", "left", "isNullObject"]);
            await context.sync();
          } catch (e) {}

          if (activeChart && !activeChart.isNullObject) {
            if (!this._lockedChart || this._lockedChart.id !== activeChart.id) {
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

              this._lockedChart = {
                id: activeChart.id,
                name: activeChart.name,
                chartType: activeChart.chartType,
                title: chartTitle,
                width: Math.round(activeChart.width || 0),
                height: Math.round(activeChart.height || 0),
                top: Math.round(activeChart.top || 0),
                left: Math.round(activeChart.left || 0)
              };
              if (ctx && ctx.adapter) {
                ctx.adapter.setLockedChart(this._lockedChart);
              }
              await this.onSelectionChange("", ctx);
            }
          } else {
            // Always keep target range selection up to date
            await this.onSelectionChange("", ctx);
          }
        }).catch(() => {});
      }
    }, 1000);
  }
};

