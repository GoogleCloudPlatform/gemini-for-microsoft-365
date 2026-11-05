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
 * Gemini for Microsoft 365 - Chat Session State & Workflow Kernel
 * 
 * Manages conversational multi-turn state (sessionId, message history),
 * API requests via askGeminiEnterprise, and workflow execution.
 */

import { askGeminiEnterprise } from '../core/geminiClient.js';

export class ChatSession {
  constructor() {
    this.sessionId = null;
    this.history = [];
    this.isProcessing = false;
  }

  /**
   * Resets conversational history and session ID.
   */
  reset() {
    this.sessionId = null;
    this.history = [];
    this.isProcessing = false;
  }

  /**
   * Returns active session ID.
   * @returns {string|null}
   */
  getSessionId() {
    return this.sessionId;
  }

  /**
   * Sets active session ID.
   * @param {string|null} id 
   */
  setSessionId(id) {
    this.sessionId = id;
  }

  /**
   * Returns conversation history array.
   * @returns {Array}
   */
  getHistory() {
    return this.history;
  }

  /**
   * Sets conversation history array.
   * @param {Array} history 
   */
  setHistory(history) {
    this.history = Array.isArray(history) ? history : [];
  }

  /**
   * Executes a standard Gemini query workflow.
   * @param {string} fullPrompt - Complete prompt including context sent to the API
   * @param {string} displayUserBubble - Text displayed in the user chat bubble
   * @param {Object} options - Context containing chatLog, renderResponseActions, and uiCallbacks
   */
  async executeWorkflow(fullPrompt, displayUserBubble, options = {}) {
    const { chatLog, renderResponseActions, uiCallbacks = {}, adapter } = options;

    if (this.isProcessing) return;
    this.isProcessing = true;

    if (uiCallbacks.onLoading) {
      uiCallbacks.onLoading(true, "⚡ Gemini 2.5 is thinking...");
    }

    try {
      if (chatLog && displayUserBubble) {
        chatLog.appendUserBubble(displayUserBubble);
      }

      const data = await askGeminiEnterprise(fullPrompt, this.history, this.sessionId);

      if (data.sessionId) {
        this.sessionId = data.sessionId;
      }
      if (Array.isArray(data.history)) {
        this.history = data.history;
      }

      const aiResponse = data.result || "No content returned.";
      if (chatLog) {
        chatLog.appendAssistantBubble(aiResponse, renderResponseActions, {
          adapter,
          chatSession: this,
          prompt: fullPrompt,
          userPrompt: displayUserBubble,
          chip: options.chip || null,
          isInsights: options.isInsights || (options.chip && options.chip.isInsights),
          isTaxonomy: options.isTaxonomy || (options.chip && (options.chip.id === 'chipTaxonomy' || options.chip.id === 'selTaxonomy'))
        });
      }

      if (uiCallbacks.onSuccess) {
        uiCallbacks.onSuccess(data);
      }
    } catch (error) {
      console.error("ChatSession workflow error:", error);
      if (chatLog) {
        chatLog.appendSystem("Error: " + (error.message || error));
      }
      if (uiCallbacks.onError) {
        uiCallbacks.onError(error);
      }
    } finally {
      this.isProcessing = false;
      if (uiCallbacks.onLoading) {
        uiCallbacks.onLoading(false);
      }
      if (chatLog) {
        chatLog.scrollToBottom();
      }
    }
  }
}
