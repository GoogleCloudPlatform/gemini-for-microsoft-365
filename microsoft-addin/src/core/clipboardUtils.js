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
 * Enterprise Multimodal Rich Clipboard Export Utility
 * 
 * Provides rich multi-MIME (text/html and text/plain) clipboard copying for
 * Microsoft Office add-ins, preserving embedded tables, typography, callouts,
 * and base64-encoded visual chart images (e.g., bar charts and infographics)
 * for native rendering when pasted into Microsoft Word, Google Docs, Apple Pages,
 * and Microsoft Excel.
 * 
 * Implements a robust 4-tier fallback architecture:
 * - Tier 1: Modern Async Clipboard API with ClipboardItem (text/html + text/plain)
 * - Tier 2: Synthetic 'copy' event interception with e.clipboardData.setData
 * - Tier 3: Hidden contenteditable container with DOM Range Selection copy
 * - Tier 4: Sanitized Plain Text fallback (navigator.clipboard.writeText / textarea)
 * 
 * @author Sathya AG, Principal Architect, Google
 */

/**
 * Wraps rich HTML body content in a complete, self-contained HTML document envelope
 * with standard Office styling for high-fidelity rendering across Word and Google Docs.
 * 
 * @param {string} bodyHtml - Inner HTML fragment
 * @returns {string} Fully styled HTML document string
 */
export function formatHtmlForClipboard(bodyHtml) {
  if (!bodyHtml) return "";

  // Normalize image tags to ensure explicit inline display and size attributes
  let processedHtml = bodyHtml.replace(
    /<img([^>]*?)src=["'](data:image\/[^;]+;base64,[^"']+)["']([^>]*?)>/gi,
    (match, p1, src, p2) => {
      // Check if style already contains max-width
      const combinedAttrs = `${p1} ${p2}`;
      if (!combinedAttrs.includes("style=")) {
        return `<img ${p1} src="${src}" ${p2} style="max-width: 100%; height: auto; display: block; margin: 14px auto; border: 1px solid #c7e0f4; border-radius: 4px;" />`;
      }
      return match;
    }
  );

  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <title>Gemini Export</title>
  <style>
    body {
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Calibri, sans-serif;
      font-size: 11pt;
      color: #201f1e;
      line-height: 1.5;
      background-color: #ffffff;
      margin: 8px;
    }
    table {
      border-collapse: collapse;
      width: 100%;
      margin: 14px 0;
      font-family: 'Segoe UI', Calibri, sans-serif;
      font-size: 10pt;
    }
    th {
      background-color: #0f6cbd;
      color: #ffffff;
      font-weight: 600;
      border: 1px solid #0f6cbd;
      padding: 8px 12px;
      text-align: left;
    }
    td {
      border: 1px solid #d2d0ce;
      padding: 8px 12px;
      color: #201f1e;
      vertical-align: top;
    }
    tr:nth-child(even) td {
      background-color: #f8f9fa;
    }
    img {
      max-width: 100%;
      height: auto;
      display: block;
      margin: 14px auto;
    }
    h1 {
      font-size: 16pt;
      font-weight: 700;
      color: #004e8c;
      margin-top: 18px;
      margin-bottom: 8px;
      border-bottom: 2px solid #0078d4;
      padding-bottom: 4px;
    }
    h2 {
      font-size: 13pt;
      font-weight: 600;
      color: #0f6cbd;
      margin-top: 14px;
      margin-bottom: 6px;
      border-bottom: 1px solid #c7e0f4;
      padding-bottom: 3px;
    }
    h3 {
      font-size: 11.5pt;
      font-weight: 600;
      color: #115ea3;
      margin-top: 10px;
      margin-bottom: 4px;
    }
    blockquote {
      border-left: 4px solid #0078d4;
      background-color: #eff6fc;
      padding: 10px 14px;
      margin: 12px 0;
      color: #004e8c;
    }
    code {
      font-family: Consolas, monospace;
      font-size: 10pt;
      background-color: #f3f2f1;
      padding: 2px 4px;
      border-radius: 3px;
    }
    pre {
      font-family: Consolas, monospace;
      font-size: 10pt;
      background-color: #201f1e;
      color: #f3f2f1;
      padding: 10px;
      border-radius: 4px;
      overflow-x: auto;
    }
  </style>
</head>
<body>
${processedHtml}
</body>
</html>`;
}

/**
 * Sanitizes plain-text / markdown representations by replacing massive raw
 * base64 data URIs with human-readable placeholders (e.g. [Image: Chart Title]).
 * Prevents 500KB+ base64 text dumps when pasting into plain text editors.
 * 
 * @param {string} rawText - Raw markdown text or assistant output
 * @param {HTMLElement|null} textDiv - Optional rendered DOM container
 * @returns {string} Clean plain text suitable for text/plain clipboard
 */
export function sanitizePlainTextForClipboard(rawText = "", textDiv = null) {
  let text = rawText || (textDiv ? textDiv.innerText : "") || "";

  // 1. Replace HTML <img> tags with clean [Image: Alt] placeholders
  text = text.replace(
    /<img[^>]+(?:alt=["']([^"']*)["'][^>]*src=["']data:image\/[^;]+;base64,[^"']+["']|src=["']data:image\/[^;]+;base64,[^"']+["'][^>]*alt=["']([^"']*)["'])[^>]*>/gi,
    (_, alt1, alt2) => {
      const alt = (alt1 || alt2 || "Visual Chart").trim();
      return `\n[Image: ${alt}]\n`;
    }
  );

  // 2. Replace standalone data:image <img> tags without alt
  text = text.replace(/<img[^>]+src=["']data:image\/[^;]+;base64,[^"']+["'][^>]*>/gi, '\n[Image: Visual Chart]\n');

  // 3. Replace Markdown format images ![alt](data:image/...)
  text = text.replace(/!\[([^\]]*)\]\(data:image\/[^\)]+\)/gi, (_, alt) => {
    const label = alt ? alt.trim() : "Visual Chart";
    return `[Image: ${label}]`;
  });

  // 4. Replace Imagen / Visual code blocks with clean placeholder
  text = text.replace(/```(?:image|imagen|visual)\s*([\s\S]*?)```/gi, (_, prompt) => {
    const cleanPrompt = prompt.trim().replace(/\s+/g, ' ').substring(0, 80);
    return `[Visual Chart: ${cleanPrompt}]`;
  });

  // 5. Strip remaining non-structural HTML tags (div, span, p tags that might linger)
  text = text.replace(/<\/div>/gi, '\n');
  text = text.replace(/<\/p>/gi, '\n\n');
  text = text.replace(/<br\s*\/?>/gi, '\n');
  text = text.replace(/<[^>]+>/g, '');

  // 6. Normalize whitespace and trailing empty lines
  text = text.replace(/\n{3,}/g, '\n\n').trim();

  return text;
}

/**
 * Copies rich content (HTML + Plain Text) to the system clipboard using
 * a resilient 4-tier fallback architecture.
 * 
 * @param {string} rawMarkdownText - Original markdown response
 * @param {HTMLElement|null} textDiv - Rendered DOM node containing styled HTML
 * @returns {Promise<boolean>} True if content was successfully copied
 */
export async function copyRichContentToClipboard(rawMarkdownText = "", textDiv = null) {
  const bodyHtml = textDiv ? textDiv.innerHTML : "";
  const fullHtmlDocument = formatHtmlForClipboard(bodyHtml);
  const cleanPlainText = sanitizePlainTextForClipboard(rawMarkdownText, textDiv);

  if (!fullHtmlDocument && !cleanPlainText) {
    return false;
  }

  // Tier 1: Modern Async Clipboard API with ClipboardItem (MIME text/html + text/plain)
  if (typeof ClipboardItem !== "undefined" && typeof navigator !== "undefined" && navigator.clipboard && navigator.clipboard.write) {
    try {
      const clipboardData = {};
      if (fullHtmlDocument) {
        clipboardData["text/html"] = new Blob([fullHtmlDocument], { type: "text/html" });
      }
      if (cleanPlainText) {
        clipboardData["text/plain"] = new Blob([cleanPlainText], { type: "text/plain" });
      }

      const item = new ClipboardItem(clipboardData);
      await navigator.clipboard.write([item]);
      return true;
    } catch (tier1Err) {
      console.warn("Clipboard Tier 1 (navigator.clipboard.write) fallback needed:", tier1Err);
    }
  }

  // Tier 2: Synthetic 'copy' event listener with e.clipboardData.setData
  if (typeof document !== "undefined") {
    try {
      let tier2Success = false;
      const onCopy = (e) => {
        e.preventDefault();
        if (e.clipboardData) {
          if (fullHtmlDocument) {
            e.clipboardData.setData("text/html", fullHtmlDocument);
          }
          if (cleanPlainText) {
            e.clipboardData.setData("text/plain", cleanPlainText);
          }
          tier2Success = true;
        }
      };

      document.addEventListener("copy", onCopy);
      try {
        tier2Success = document.execCommand("copy");
      } finally {
        document.removeEventListener("copy", onCopy);
      }

      if (tier2Success) {
        return true;
      }
    } catch (tier2Err) {
      console.warn("Clipboard Tier 2 (copy event interception) fallback needed:", tier2Err);
    }
  }

  // Tier 3: Hidden contenteditable container with DOM Range selection + execCommand('copy')
  if (typeof document !== "undefined" && fullHtmlDocument) {
    try {
      const tempDiv = document.createElement("div");
      tempDiv.contentEditable = "true";
      tempDiv.style.position = "fixed";
      tempDiv.style.left = "-9999px";
      tempDiv.style.top = "0";
      tempDiv.style.opacity = "0";
      tempDiv.style.pointerEvents = "none";
      tempDiv.innerHTML = fullHtmlDocument;
      document.body.appendChild(tempDiv);

      const range = document.createRange();
      range.selectNodeContents(tempDiv);
      const sel = window.getSelection ? window.getSelection() : null;
      if (sel) {
        sel.removeAllRanges();
        sel.addRange(range);
        const tier3Success = document.execCommand("copy");
        sel.removeAllRanges();
        document.body.removeChild(tempDiv);

        if (tier3Success) {
          return true;
        }
      } else {
        document.body.removeChild(tempDiv);
      }
    } catch (tier3Err) {
      console.warn("Clipboard Tier 3 (Range Selection copy) fallback needed:", tier3Err);
    }
  }

  // Tier 4: Baseline Plain Text write (navigator.clipboard.writeText / textarea fallback)
  if (typeof navigator !== "undefined" && navigator.clipboard && navigator.clipboard.writeText && cleanPlainText) {
    try {
      await navigator.clipboard.writeText(cleanPlainText);
      return true;
    } catch (tier4Err) {
      console.warn("Clipboard Tier 4 (writeText) fallback needed:", tier4Err);
    }
  }

  if (typeof document !== "undefined" && cleanPlainText) {
    try {
      const tempArea = document.createElement("textarea");
      tempArea.value = cleanPlainText;
      tempArea.setAttribute("readonly", "");
      tempArea.style.position = "fixed";
      tempArea.style.left = "-9999px";
      tempArea.style.top = "0";
      tempArea.style.opacity = "0";
      document.body.appendChild(tempArea);
      tempArea.focus();
      tempArea.select();
      tempArea.setSelectionRange(0, 99999);
      const success = document.execCommand("copy");
      document.body.removeChild(tempArea);
      return success;
    } catch (execErr) {
      console.error("Clipboard Tier 4 textarea copy failed:", execErr);
    }
  }

  return false;
}
