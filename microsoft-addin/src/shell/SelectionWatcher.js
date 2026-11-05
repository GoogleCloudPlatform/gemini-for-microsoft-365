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
 * Gemini for Microsoft 365 - Document Selection Watcher
 * 
 * Listens to Office selection events, manages selection toolbar UI visibility,
 * and notifies host-specific selection handlers.
 */

/* global Office */

export class SelectionWatcher {
  /**
   * @param {Object} options
   * @param {Object} options.adapter - Host document adapter
   * @param {Object} options.hostModule - Active host module descriptor
   * @param {Object} options.chatLog - Active ChatLog instance
   * @param {Object} options.chatSession - Active ChatSession instance
   */
  constructor(options = {}) {
    this.adapter = options.adapter;
    this.hostModule = options.hostModule || {};
    this.chatLog = options.chatLog;
    this.chatSession = options.chatSession;
    this.currentSelectedText = "";
    this.selectionRequestId = 0;
  }

  /**
   * Initializes selection event listener.
   */
  init() {
    try {
      if (typeof Office !== 'undefined' && Office.context && Office.context.document && Office.context.document.addHandlerAsync) {
        Office.context.document.addHandlerAsync(
          Office.EventType.DocumentSelectionChanged,
          () => this.handleSelectionChanged()
        );
      }
    } catch (e) {
      console.warn("Could not attach DocumentSelectionChanged handler:", e);
    }
  }

  /**
   * Returns latest cached selected text.
   * @returns {string}
   */
  getCurrentSelection() {
    return this.currentSelectedText;
  }

  /**
   * Sets cached selection text.
   * @param {string} text 
   */
  setCurrentSelection(text) {
    this.currentSelectedText = text || "";
  }

  /**
   * Handles document selection change event.
   */
  async handleSelectionChanged() {
    if (!this.adapter) return;

    const currentReq = ++this.selectionRequestId;
    const selToolbar = document.getElementById("selectionToolbar");

    try {
      const text = typeof this.adapter.getSelectedRawText === "function"
        ? await this.adapter.getSelectedRawText()
        : await this.adapter.getSelectedText();
      if (this.selectionRequestId !== currentReq) return; // Stale request

      this.currentSelectedText = text ? text.trim() : "";

      if (selToolbar) {
        selToolbar.style.display = this.currentSelectedText.length > 0 ? "block" : "none";
      }

      // Invoke per-host hook if defined
      if (typeof this.hostModule.onSelectionChange === 'function') {
        await this.hostModule.onSelectionChange(this.currentSelectedText, {
          adapter: this.adapter,
          chatLog: this.chatLog,
          chatSession: this.chatSession,
          selectionToolbar: selToolbar
        });
      }

      // Check for in-document @gemini command
      if (typeof this.adapter.checkInDocumentCommands === 'function') {
        const debugStatus = document.getElementById("debugStatus");
        await this.adapter.checkInDocumentCommands(false, {
          onStatus: (msg) => {
            if (debugStatus) debugStatus.innerText = msg;
          },
          executePrompt: async (userPrompt) => {
            return this.executeInDocPrompt(userPrompt);
          }
        });
      }
    } catch (err) {
      console.warn("Selection change handler warning:", err);
    }
  }

  /**
   * Executes an in-document @gemini prompt.
   * @param {string} userPrompt 
   * @returns {Promise<string|null>}
   */
  async executeInDocPrompt(userPrompt) {
    if (!this.chatSession) return null;

    const debugStatus = document.getElementById("debugStatus");
    const scanBtn = document.getElementById("scanInDoc");
    const loadingText = document.getElementById("loading");

    if (scanBtn) scanBtn.disabled = true;
    if (loadingText) {
      loadingText.innerText = "⚡ Generating @gemini response...";
      loadingText.style.display = "block";
    }

    if (this.chatLog) {
      this.chatLog.appendUserBubble(`@gemini ${userPrompt}`);
    }

    try {
      let resultText = null;
      await this.chatSession.executeWorkflow(userPrompt, null, {
        chatLog: this.chatLog,
        renderResponseActions: this.hostModule.renderResponseActions,
        adapter: this.adapter,
        uiCallbacks: {
          onSuccess: (data) => {
            resultText = data.result || "";
          },
          onError: (err) => {
            if (debugStatus) debugStatus.innerText = "Error: " + err.message;
          }
        }
      });
      return resultText;
    } catch (err) {
      console.error("In-document execution error:", err);
      return null;
    } finally {
      if (scanBtn) scanBtn.disabled = false;
      if (loadingText) loadingText.style.display = "none";
    }
  }
}
