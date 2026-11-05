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
 * Gemini for Microsoft 365 - Taskpane Shell Orchestrator
 * 
 * Shared lifecycle engine that initializes Office.js, mounts host modules,
 * renders declarative chips, coordinates chat workflows, and manages document I/O.
 */

/* global Office */

import { ChatSession } from './ChatSession.js';
import { ChatLog } from './ChatLog.js';
import { SelectionWatcher } from './SelectionWatcher.js';
import { HostAdapterFactory } from '../adapters/HostAdapterFactory.js';
import { setProxyUrlOverride } from '../core/geminiClient.js';

export class TaskpaneShell {
  /**
   * Boots the taskpane shell with the given host module descriptor.
   * @param {Object} hostModule - Declarative host module descriptor
   */
  static async boot(hostModule) {
    if (!hostModule) {
      console.error("TaskpaneShell.boot called without hostModule");
      return;
    }

    Office.onReady(async (info) => {
      try {
        // Validate host environment if running inside Office client
        if (info && info.host && hostModule.hostName) {
          const expectedHostType = Office.HostType[hostModule.hostName];
          if (expectedHostType && info.host !== expectedHostType) {
            console.warn(`Host mismatch: running in ${info.host}, expected ${hostModule.hostName}`);
          }
        }

        // 1. Initialize Document Adapter
        const adapter = typeof hostModule.createAdapter === 'function'
          ? hostModule.createAdapter()
          : HostAdapterFactory.getAdapter();

        // 2. Initialize Core Subsystems
        const chatSession = new ChatSession();
        const chatLog = new ChatLog("#chatHistory");
        const selectionWatcher = new SelectionWatcher({
          adapter,
          hostModule,
          chatLog,
          chatSession
        });

        // 3. Populate Host-Specific Localized Strings
        if (hostModule.strings) {
          const tipBannerText = document.getElementById("tipBannerText");
          if (tipBannerText && hostModule.strings.tipBanner) {
            tipBannerText.innerHTML = hostModule.strings.tipBanner;
          }

          const docToolsTitle = document.getElementById("docToolsTitle");
          if (docToolsTitle && hostModule.strings.toolsTitle) {
            docToolsTitle.innerHTML = hostModule.strings.toolsTitle;
          }

          const welcomeBubble = document.getElementById("welcomeSystemBubble");
          if (welcomeBubble && hostModule.strings.welcome) {
            welcomeBubble.innerHTML = hostModule.strings.welcome;
          }

          const debugStatus = document.getElementById("debugStatus");
          if (debugStatus && hostModule.strings.readyStatus) {
            debugStatus.innerText = hostModule.strings.readyStatus;
          }
        }

        // Dynamically update model badge from parameterized environment configuration
        const activeModelLabel = (typeof process !== 'undefined' && process.env && process.env.GEMINI_MODEL_LABEL)
          || (typeof process !== 'undefined' && process.env && process.env.GEMINI_MODEL ? process.env.GEMINI_MODEL.replace(/^gemini-/, '').replace(/-/g, ' ') : 'Enterprise');
        document.querySelectorAll('.header-badge, .model-badge').forEach((badge) => {
          if (badge && !badge.classList.contains('coach-option-badge')) {
            badge.innerText = activeModelLabel;
          }
        });

        // 4. Render Document Intelligence Chips
        TaskpaneShell.renderDocChips(hostModule.docChips, {
          adapter,
          chatSession,
          chatLog,
          hostModule
        });

        // 5. Render Selection Quick Toolbar Chips
        TaskpaneShell.renderSelectionChips(hostModule.selectionChips, {
          adapter,
          selectionWatcher,
          chatSession,
          chatLog,
          hostModule
        });

        // 6. Wire Input Controls & Buttons
        TaskpaneShell.wireControls({
          adapter,
          chatSession,
          chatLog,
          selectionWatcher,
          hostModule
        });

        // 7. Initialize Selection Listener
        selectionWatcher.init();

        // 8. Invoke Per-Host onReady Hook if Provided
        if (typeof hostModule.onReady === 'function') {
          await hostModule.onReady({
            adapter,
            chatSession,
            chatLog,
            selectionWatcher,
            shell: TaskpaneShell
          });
        }
      } catch (err) {
        console.error("TaskpaneShell boot error:", err);
      }
    });
  }

  /**
   * Renders document intelligence chips into the container.
   */
  static renderDocChips(chips, ctx) {
    const container = document.getElementById("docToolsChipsContainer") || document.querySelector(".doc-tools-bar .chips-container");
    if (!container || !Array.isArray(chips)) return;

    container.innerHTML = "";
    chips.forEach(chip => {
      const btn = document.createElement("button");
      btn.className = "quick-chip";
      if (chip.id) btn.id = chip.id;
      btn.innerHTML = chip.label || "Chip";
      if (chip.title) btn.title = chip.title;

      btn.onclick = async () => {
        const loadingText = document.getElementById("loading");
        if (loadingText) {
          loadingText.innerText = (chip.isInsights || chip.isExecSummary)
            ? "⚡ Pre-parsing Excel table data & calculating statistics..."
            : `⚡ Analyzing ${ctx.hostModule.hostName || 'document'} context...`;
          loadingText.style.display = "block";
        }

        try {
          let fullPrompt = "";
          let displayBubble = "";

          if (typeof chip.handler === 'function') {
            const result = await chip.handler(ctx);
            fullPrompt = (result && result.fullPrompt) ? result.fullPrompt : result;
            displayBubble = (result && result.displayBubble) ? result.displayBubble : `📊 ${chip.label}`;
          } else if (chip.isInsights && typeof ctx.adapter.getTableInsightsPrompt === 'function') {
            fullPrompt = await ctx.adapter.getTableInsightsPrompt(chip.scope || "auto");
            const scopeLabel = chip.scope === "selection" ? "Highlighted Cells" : (chip.scope === "sheet" ? "Entire Sheet" : "Active Range / Sheet");
            displayBubble = `📊 [Table Insights] Scope: ${scopeLabel}`;
          } else if (chip.isExecSummary && typeof ctx.adapter.getExecutiveSummaryPrompt === 'function') {
            fullPrompt = await ctx.adapter.getExecutiveSummaryPrompt(chip.scope || "auto");
            const scopeLabel = chip.scope === "selection" ? "Highlighted Cells" : (chip.scope === "sheet" ? "Entire Sheet" : "Active Range / Sheet");
            displayBubble = `📑 [Executive Summary] Scope: ${scopeLabel}`;
          } else if (chip.isTaxonomy && typeof ctx.adapter.getTaxonomyPrompt === 'function') {
            fullPrompt = await ctx.adapter.getTaxonomyPrompt(chip.scope || "auto");
            const scopeLabel = chip.scope === "selection" ? "Highlighted Cells" : (chip.scope === "sheet" ? "Entire Sheet" : "Active Range / Sheet");
            displayBubble = `🏷️ [Taxonomy Categorizer] Scope: ${scopeLabel}`;
          } else {
            let selectionText = "";
            if (ctx.selectionWatcher && typeof ctx.selectionWatcher.getCurrentSelection === 'function') {
              selectionText = ctx.selectionWatcher.getCurrentSelection();
            }
            if (!selectionText && ctx.adapter && typeof ctx.adapter.getSelectedText === 'function') {
              try {
                selectionText = await ctx.adapter.getSelectedText();
              } catch (_) {}
            }
            const fullDocText = await ctx.adapter.getFullDocumentText();
            const contextParts = [];
            if (selectionText && selectionText.trim().length > 0) {
              contextParts.push(`Selected Data / Range Context:\n"""\n${selectionText.substring(0, 50000)}\n"""`);
            }
            if (fullDocText && fullDocText.trim().length > 0) {
              contextParts.push(`Active Document / Sheet Context:\n"""\n${fullDocText.substring(0, 100000)}\n"""`);
            }

            if (contextParts.length > 0) {
              fullPrompt = `${contextParts.join("\n\n")}\n\nUser Instruction: ${chip.prompt}`;
            } else {
              fullPrompt = chip.prompt;
            }
            displayBubble = `📊 [${ctx.hostModule.hostName || 'Document'} Analysis] ${(chip.label || chip.prompt || "").substring(0, 55)}...`;
          }
          await ctx.chatSession.executeWorkflow(fullPrompt, displayBubble, {
            chatLog: ctx.chatLog,
            renderResponseActions: ctx.hostModule.renderResponseActions,
            adapter: ctx.adapter,
            chip: chip,
            isInsights: chip.isInsights,
            isTaxonomy: chip.isTaxonomy || chip.id === "chipTaxonomy"
          });
        } catch (err) {
          console.error("Doc intelligence chip error:", err);
          if (ctx.chatLog) ctx.chatLog.appendSystem("Error reading context: " + err.message);
        } finally {
          if (loadingText) loadingText.style.display = "none";
        }
      };

      container.appendChild(btn);
    });
  }

  static populateDocChips(chips, ctx) {
    return TaskpaneShell.renderDocChips(chips, ctx);
  }

  /**
   * Renders selection chips into the selection toolbar.
   */
  static renderSelectionChips(chips, ctx) {
    const container = document.getElementById("selectionChipsContainer") || document.querySelector("#selectionToolbar .chips-container");
    if (!container || !Array.isArray(chips)) return;

    container.innerHTML = "";
    chips.forEach(chip => {
      const btn = document.createElement("button");
      btn.className = "quick-chip";
      if (chip.id) btn.id = chip.id;
      btn.innerHTML = chip.label || "Chip";
      if (chip.title) btn.title = chip.title;

      btn.onclick = async () => {
        const loadingText = document.getElementById("loading");
        if (loadingText && (chip.isInsights || chip.isExecSummary || chip.isTaxonomy)) {
          loadingText.innerText = "⚡ Pre-parsing Excel table data & calculating statistics...";
          loadingText.style.display = "block";
        }

        try {
          if (chip.isInsights && typeof ctx.adapter.getTableInsightsPrompt === 'function') {
            const fullPrompt = await ctx.adapter.getTableInsightsPrompt("selection");
            const displayBubble = `📊 [Table Insights] Scope: Highlighted Cells`;
            await ctx.chatSession.executeWorkflow(fullPrompt, displayBubble, {
              chatLog: ctx.chatLog,
              renderResponseActions: ctx.hostModule.renderResponseActions,
              adapter: ctx.adapter,
              chip: chip,
              isInsights: true
            });
            return;
          }

          if (chip.isExecSummary && typeof ctx.adapter.getExecutiveSummaryPrompt === 'function') {
            const fullPrompt = await ctx.adapter.getExecutiveSummaryPrompt("selection");
            const displayBubble = `💡 [Executive Summary] Scope: Highlighted Cells`;
            await ctx.chatSession.executeWorkflow(fullPrompt, displayBubble, {
              chatLog: ctx.chatLog,
              renderResponseActions: ctx.hostModule.renderResponseActions,
              adapter: ctx.adapter,
              chip: chip,
              isExecSummary: true
            });
            return;
          }

          if (chip.isTaxonomy && typeof ctx.adapter.getTaxonomyPrompt === 'function') {
            const fullPrompt = await ctx.adapter.getTaxonomyPrompt("selection");
            const displayBubble = `🏷️ [Taxonomy Categorizer] Scope: Highlighted Cells`;
            await ctx.chatSession.executeWorkflow(fullPrompt, displayBubble, {
              chatLog: ctx.chatLog,
              renderResponseActions: ctx.hostModule.renderResponseActions,
              adapter: ctx.adapter,
              chip: chip,
              isTaxonomy: true
            });
            return;
          }

          let currentSelectedText = ctx.selectionWatcher.getCurrentSelection();
          if (!currentSelectedText) {
            currentSelectedText = await ctx.adapter.getSelectedText();
          }

          if (!currentSelectedText || currentSelectedText.trim().length === 0) {
            if (ctx.chatLog) {
              ctx.chatLog.appendSystem(`Please highlight content in ${ctx.hostModule.hostName || 'Office'} first.`);
            }
            return;
          }

          const fullPrompt = `Selected Document Text:\n"${currentSelectedText}"\n\nTask: ${chip.instruction}`;
          const displayBubble = `✨ ${chip.instruction}\n📌 Context: "${currentSelectedText.substring(0, 70)}..."`;

          await ctx.chatSession.executeWorkflow(fullPrompt, displayBubble, {
            chatLog: ctx.chatLog,
            renderResponseActions: ctx.hostModule.renderResponseActions,
            adapter: ctx.adapter,
            chip: chip,
            isTaxonomy: chip.isTaxonomy || chip.id === "selTaxonomy"
          });
        } catch (err) {
          console.error("Selection chip error:", err);
          if (ctx.chatLog) ctx.chatLog.appendSystem("Error: " + err.message);
        } finally {
          if (loadingText) loadingText.style.display = "none";
        }
      };

      container.appendChild(btn);
    });
  }

  /**
   * Wires interactive controls and listeners.
   */
  static wireControls(ctx) {
    const runBtn = document.getElementById("run");
    const promptInput = document.getElementById("promptText");
    const clearSessionBtn = document.getElementById("clearSession");
    const scanBtn = document.getElementById("scanInDoc");
    const endpointSelect = document.getElementById("endpointSelect");

    // Send Button Workflow
    if (runBtn) {
      runBtn.onclick = async () => {
        const userText = promptInput ? promptInput.value.trim() : "";
        let selectedText = await ctx.adapter.getSelectedText();
        if (selectedText) selectedText = selectedText.trim();

        if (!userText && !selectedText) {
          if (ctx.chatLog) {
            ctx.chatLog.appendSystem(`Please type a prompt or select content in ${ctx.hostModule.hostName || 'Office'} first.`);
          }
          return;
        }

        let fullPrompt = "";
        let displayUserBubble = "";

        if (selectedText && userText) {
          fullPrompt = `Selected Document Context:\n"${selectedText}"\n\nUser Instruction: ${userText}`;
          displayUserBubble = `📌 Context: "${selectedText.substring(0, 70)}${selectedText.length > 70 ? '...' : ''}"\n\n${userText}`;
        } else if (selectedText && !userText) {
          fullPrompt = `Please analyze, summarize, or explain the following selected text:\n"${selectedText}"`;
          displayUserBubble = `📌 Selected Text:\n"${selectedText.substring(0, 90)}${selectedText.length > 90 ? '...' : ''}"`;
        } else {
          fullPrompt = userText;
          displayUserBubble = userText;
        }

        if (promptInput) promptInput.value = "";

        const loadingText = document.getElementById("loading");
        await ctx.chatSession.executeWorkflow(fullPrompt, displayUserBubble, {
          chatLog: ctx.chatLog,
          renderResponseActions: ctx.hostModule.renderResponseActions,
          adapter: ctx.adapter,
          prompt: fullPrompt,
          userPrompt: userText || displayUserBubble,
          uiCallbacks: {
            onLoading: (isLoading, msg) => {
              if (runBtn) runBtn.disabled = isLoading;
              if (loadingText) {
                if (isLoading) {
                  loadingText.innerText = msg || "⚡ Gemini is thinking...";
                  loadingText.style.display = "block";
                } else {
                  loadingText.style.display = "none";
                }
              }
            }
          }
        });
      };
    }

    // Form Keyboard Shortcuts:
    // - Enter: Submit / Execute prompt
    // - Shift + Enter or Ctrl + Enter: Insert newline in editor
    if (promptInput) {
      promptInput.addEventListener("keydown", (e) => {
        if (e.key === "Enter") {
          if (e.shiftKey) {
            // Shift + Enter: Default textarea behavior inserts a newline
            return;
          }
          if (e.ctrlKey || e.metaKey) {
            // Ctrl + Enter or Cmd + Enter: Explicitly insert newline at cursor
            e.preventDefault();
            const start = promptInput.selectionStart;
            const end = promptInput.selectionEnd;
            const val = promptInput.value;
            promptInput.value = val.substring(0, start) + "\n" + val.substring(end);
            promptInput.selectionStart = promptInput.selectionEnd = start + 1;
            return;
          }
          // Enter without modifiers: Submit prompt
          e.preventDefault();
          if (runBtn && !runBtn.disabled) {
            runBtn.click();
          }
        }
      });
    }

    // Clear Session Button
    if (clearSessionBtn) {
      clearSessionBtn.onclick = () => {
        ctx.chatSession.reset();
        if (ctx.chatLog) {
          ctx.chatLog.clear('<div class="chat-bubble system">Chat session reset. Ready for a new topic!</div>');
          ctx.chatLog.scrollToTop();
        }
        const chatHistory = document.getElementById("chatHistory");
        if (chatHistory) {
          chatHistory.scrollTop = 0;
        }
        const debugStatus = document.getElementById("debugStatus");
        if (debugStatus) {
          debugStatus.innerText = `${ctx.hostModule.hostName || 'Office'} Ready (Session Reset)`;
        }
        if (typeof ctx.hostModule.onSessionReset === 'function') {
          ctx.hostModule.onSessionReset(ctx);
        }
      };
    }

    // In-Document @gemini Scan Button
    if (scanBtn) {
      scanBtn.onclick = () => {
        if (typeof ctx.adapter.checkInDocumentCommands === 'function') {
          const debugStatus = document.getElementById("debugStatus");
          ctx.adapter.checkInDocumentCommands(true, {
            onStatus: (msg) => {
              if (debugStatus) debugStatus.innerText = msg;
            },
            executePrompt: async (userPrompt) => {
              return ctx.selectionWatcher.executeInDocPrompt(userPrompt);
            }
          });
        }
      };
    }

    // Endpoint Override Selector
    if (endpointSelect) {
      const savedUrl = window.localStorage ? window.localStorage.getItem('gemini_proxy_url') : "";
      if (savedUrl) {
        endpointSelect.value = savedUrl;
      }
      endpointSelect.onchange = (e) => {
        const selectedUrl = e.target.value;
        setProxyUrlOverride(selectedUrl);
        const debugStatus = document.getElementById("debugStatus");
        if (debugStatus) {
          debugStatus.innerText = selectedUrl ? `Target: Override Active` : `${ctx.hostModule.hostName || 'Office'} Ready`;
        }
      };
    }

    // Collapsible Document Intelligence Quick Bar
    const docToolsBar = document.getElementById("docToolsBar") || document.querySelector(".doc-tools-bar");
    const docToolsToggle = document.getElementById("docToolsToggle") || (docToolsBar ? docToolsBar.querySelector(".doc-tools-title") : null);
    const docToolsContainer = document.getElementById("docToolsChipsContainer") || (docToolsBar ? docToolsBar.querySelector(".chips-container") : null);
    const docToolsIcon = document.getElementById("docToolsToggleIcon");

    if (docToolsToggle && docToolsContainer && docToolsBar) {
      const toggleSection = (forceExpand) => {
        const isCollapsed = docToolsBar.classList.contains("collapsed") || docToolsContainer.style.display === "none";
        const willExpand = forceExpand !== undefined ? forceExpand : isCollapsed;

        if (willExpand) {
          docToolsBar.classList.remove("collapsed");
          docToolsBar.classList.add("expanded");
          docToolsContainer.style.display = "flex";
          if (docToolsIcon) docToolsIcon.textContent = "▼";
          docToolsToggle.setAttribute("aria-expanded", "true");
        } else {
          docToolsBar.classList.add("collapsed");
          docToolsBar.classList.remove("expanded");
          docToolsContainer.style.display = "none";
          if (docToolsIcon) docToolsIcon.textContent = "▶";
          docToolsToggle.setAttribute("aria-expanded", "false");
        }
      };

      docToolsToggle.addEventListener("click", () => toggleSection());
      docToolsToggle.addEventListener("keydown", (e) => {
        if (e.key === "Enter" || e.key === " ") {
          e.preventDefault();
          toggleSection();
        }
      });
    }
  }

  /**
   * Universal helper for performing document insertion via the host adapter.
   */
  static async performDocumentInsertion(adapter, htmlContent, rawText, mode = "smart", hostName = "Office") {
    const runButton = document.getElementById("run");
    const loadingText = document.getElementById("loading");

    if (runButton) runButton.disabled = true;
    if (loadingText) {
      loadingText.innerText = `⚡ Updating ${hostName}...`;
      loadingText.style.display = "block";
    }

    try {
      await adapter.insertContent(htmlContent, rawText, mode);
      const debugStatus = document.getElementById("debugStatus");
      if (debugStatus) {
        debugStatus.innerText = `Updated in ${hostName} (${mode === 'replace_draft' ? 'Replaced' : 'Inserted'})`;
      }
      const chatHistory = document.getElementById("chatHistory");
      if (chatHistory) {
        const successBubble = document.createElement("div");
        successBubble.className = "chat-bubble system";
        successBubble.style.borderLeft = "4px solid #107c10";
        successBubble.style.backgroundColor = "#f0fdf4";
        successBubble.style.color = "#15803d";
        successBubble.innerText = `✅ Successfully inserted into ${hostName}!`;
        chatHistory.appendChild(successBubble);
        chatHistory.scrollTop = chatHistory.scrollHeight;
      }
    } catch (err) {
      console.error("Document insertion error:", err);
      const chatHistory = document.getElementById("chatHistory");
      if (chatHistory) {
        const errorBubble = document.createElement("div");
        errorBubble.className = "chat-bubble system";
        errorBubble.innerText = `🔴 Insertion Error: ${err.message || err}`;
        chatHistory.appendChild(errorBubble);
        chatHistory.scrollTop = chatHistory.scrollHeight;
      }
    } finally {
      if (runButton) runButton.disabled = false;
      if (loadingText) loadingText.style.display = "none";
    }
  }
}
