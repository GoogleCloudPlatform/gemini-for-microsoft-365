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
 * Gemini for Microsoft 365 - Chat Log DOM Renderer
 * 
 * Manages the chat history view, appending user, system, and formatted
 * assistant bubbles with markdown rendering and action buttons.
 */

import { parseMarkdown } from '../core/markdownParser.js';
import { copyRichContentToClipboard } from '../core/clipboardUtils.js';

export class ChatLog {
  /**
   * @param {HTMLElement|string} container - Container element or selector (defaults to #chatHistory)
   */
  constructor(container = "#chatHistory") {
    this.container = typeof container === 'string' ? document.querySelector(container) : container;
  }

  /**
   * Appends a user message bubble.
   * @param {string} text 
   * @returns {HTMLElement}
   */
  appendUserBubble(text) {
    if (!this.container) return null;
    const bubble = document.createElement("div");
    bubble.className = "chat-bubble user";
    bubble.innerText = text;
    this.container.appendChild(bubble);
    this.scrollToBottom();
    return bubble;
  }

  /**
   * Appends a system notification bubble.
   * @param {string} textOrHtml 
   * @param {boolean} isHtml 
   * @returns {HTMLElement}
   */
  appendSystem(textOrHtml, isHtml = false) {
    if (!this.container) return null;
    const bubble = document.createElement("div");
    bubble.className = "chat-bubble system";
    if (isHtml) {
      bubble.innerHTML = textOrHtml;
    } else {
      bubble.innerText = textOrHtml;
    }
    this.container.appendChild(bubble);
    this.scrollToBottom();
    return bubble;
  }

  /**
   * Appends an assistant response bubble with markdown parsing and host response actions.
   * @param {string} rawMarkdownText 
   * @param {Function} renderResponseActionsFn 
   * @param {Object} actionContext 
   * @returns {HTMLElement}
   */
  appendAssistantBubble(rawMarkdownText, renderResponseActionsFn = null, actionContext = {}) {
    if (!this.container) return null;

    const bubble = document.createElement("div");
    bubble.className = "chat-bubble assistant";

    // 1. Format raw markdown to executive HTML
    const formattedHtml = parseMarkdown(rawMarkdownText);
    const textDiv = document.createElement("div");
    textDiv.innerHTML = formattedHtml;
    bubble.appendChild(textDiv);

    // 2. Action Toolbar & Status Footer Container
    const actionsContainer = document.createElement("div");
    actionsContainer.className = "response-actions-container";

    if (typeof renderResponseActionsFn === 'function') {
      const customActions = renderResponseActionsFn({
        rawText: rawMarkdownText,
        formattedHtml,
        textDiv,
        bubble,
        chatLog: this,
        ...actionContext
      });
      if (customActions) {
        actionsContainer.appendChild(customActions);
      }
    } else {
      // Fallback copy action
      const primaryActions = document.createElement("div");
      primaryActions.className = "primary-actions";
      const copyBtn = this.createCopyButton(rawMarkdownText, textDiv);
      primaryActions.appendChild(copyBtn);
      actionsContainer.appendChild(primaryActions);
    }

    // 3. Status footer
    const footerStatus = document.createElement("div");
    footerStatus.className = "response-footer-status";
    footerStatus.innerHTML = `<span>✨ Response complete • Awaiting your next request</span>`;
    actionsContainer.appendChild(footerStatus);

    bubble.appendChild(actionsContainer);
    this.container.appendChild(bubble);
    this.scrollToBottom();

    return bubble;
  }

  /**
   * Helper: Creates a rich clipboard copy button for an AI message.
   * Preserves rich images in HTML clipboard while keeping clean text without base64 noise.
   * @param {string} text - Plain text / markdown content
   * @param {HTMLElement} textDiv - Rendered message DOM element
   * @returns {HTMLButtonElement}
   */
  createCopyButton(text, textDiv) {
    const copyBtn = document.createElement("button");
    copyBtn.className = "action-btn copy";
    copyBtn.innerHTML = `<svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor" style="vertical-align: -2px; margin-right: 4px;"><path d="M16 1H4c-1.1 0-2 .9-2 2v14h2V3h12V1zm3 4H8c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h11c1.1 0 2-.9 2-2V7c0-1.1-.9-2-2-2zm0 16H8V7h11v14z"/></svg>Copy to Clipboard`;
    copyBtn.title = "Copy rich formatted content with tables and images to clipboard";
    copyBtn.onclick = async () => {
      const success = await copyRichContentToClipboard(text, textDiv);

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
      } else {
        this.appendSystem("Unable to copy to clipboard in this environment. Please select and copy text manually.");
      }
    };
    return copyBtn;
  }

  /**
   * Resets chat history and renders an optional greeting bubble,
   * while preserving the top Quick Actions component and resetting scroll to top.
   * @param {string} welcomeHtml 
   */
  clear(welcomeHtml = null) {
    if (!this.container) return;
    const selectionToolbar = (typeof this.container.querySelector === 'function')
      ? this.container.querySelector("#selectionToolbar, .quick-actions-toolbar")
      : null;
    this.container.innerHTML = "";
    if (selectionToolbar) {
      this.container.appendChild(selectionToolbar);
    }
    if (welcomeHtml) {
      const bubble = this.appendSystem(welcomeHtml, true);
      if (bubble) {
        bubble.id = "welcomeSystemBubble";
      }
    }
    this.scrollToTop();
  }

  /**
   * Scrolls chat history viewport to the top.
   */
  scrollToTop() {
    if (this.container) {
      this.container.scrollTop = 0;
    }
  }

  /**
   * Scrolls chat history viewport to the bottom.
   */
  scrollToBottom() {
    if (this.container) {
      this.container.scrollTop = this.container.scrollHeight;
    }
  }
}
