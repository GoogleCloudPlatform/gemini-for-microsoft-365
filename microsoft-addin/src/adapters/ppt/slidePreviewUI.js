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
 * Slide Preview UI & Interactive Deck Handler for PowerPoint
 * 
 * Automatically attaches to the Add-in Taskpane in PowerPoint, detects multi-slide
 * content in Gemini responses, and renders an executive Slide Deck Outline Preview.
 * Enhances the native "+ Insert into Slides" toolbar button to generate the complete deck.
 * Image responses show one preview per picture and no "Replace with Image" action.
 * 
 * @author Sathya AG, Principal Architect, Google
 */

import { parseSlides } from './slideParser.js';

let stylesInjected = false;

/**
 * Injects scoped styling for the PowerPoint Slide Deck Preview into the DOM.
 */
export function injectPowerPointStyles() {
  if (stylesInjected || typeof document === "undefined") return;
  stylesInjected = true;

  const styleEl = document.createElement("style");
  styleEl.id = "ppt-preview-styles";
  styleEl.textContent = `
    .ppt-deck-preview-container {
      margin: 10px 0 8px 0;
      background: #ffffff;
      border: 1px solid #d3e3fd;
      border-left: 4px solid #0b57d0;
      border-radius: 8px;
      padding: 10px;
      box-shadow: 0 2px 8px rgba(11, 87, 208, 0.08);
    }
    .ppt-deck-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 8px;
      padding-bottom: 6px;
      border-bottom: 1px solid #f0f4f9;
    }
    .ppt-deck-title {
      font-size: 12px;
      font-weight: 700;
      color: #0b57d0;
      display: flex;
      align-items: center;
      gap: 6px;
    }
    .ppt-deck-badge {
      font-size: 10.5px;
      background: #d3e3fd;
      color: #0b57d0;
      padding: 2px 8px;
      border-radius: 9999px;
      font-weight: 600;
    }
    .ppt-slides-list {
      display: flex;
      flex-direction: column;
      gap: 6px;
      max-height: 220px;
      overflow-y: auto;
      margin-bottom: 6px;
      padding-right: 2px;
    }
    .ppt-slide-card {
      background: #f8fafd;
      border: 1px solid #e1e3e1;
      border-radius: 6px;
      padding: 6px 8px;
      font-size: 11.5px;
      transition: all 0.15s ease;
    }
    .ppt-slide-card:hover {
      border-color: #0b57d0;
      background: #f0f4f9;
    }
    .ppt-slide-card-header {
      font-weight: 600;
      color: #0b57d0;
      display: flex;
      align-items: baseline;
      gap: 6px;
      margin-bottom: 3px;
    }
    .ppt-slide-num {
      font-size: 10px;
      background: #0b57d0;
      color: #ffffff;
      padding: 1px 6px;
      border-radius: 9999px;
      font-weight: 700;
    }
    .ppt-slide-preview-body {
      color: #444746;
      font-size: 11px;
      line-height: 1.4;
      white-space: pre-line;
      padding-top: 4px;
    }
    .ppt-deck-footer-hint {
      font-size: 11px;
      color: #0b57d0;
      margin-top: 6px;
      text-align: center;
      font-weight: 500;
      background: #f0f4f9;
      padding: 4px 6px;
      border-radius: 6px;
    }
  `;
  document.head.appendChild(styleEl);
}

/**
 * The parts of a PNG data URI that decide what it looks like: the IHDR, PLTE
 * and tRNS chunks plus the compressed pixel stream (every IDAT chunk, joined).
 * All other chunks are metadata (text, EXIF, timestamps, colour profiles, C2PA
 * content credentials), so two files of one picture whose metadata differs,
 * even in length, return the same string. Returns null for anything that is
 * not a complete PNG data URI.
 */
export function pngPixelData(src) {
  if (typeof src !== "string" || typeof atob !== "function") return null;
  const head = /^data:image\/png;base64,/i.exec(src);
  if (!head) return null;
  let bytes;
  try {
    bytes = atob(src.slice(head[0].length));
  } catch (err) {
    return null;
  }
  if (bytes.slice(0, 8) !== "\x89PNG\r\n\x1a\n") return null;
  let header = "";
  const pixels = [];
  let pos = 8;
  while (pos + 12 <= bytes.length) {
    const length = ((bytes.charCodeAt(pos) << 24) | (bytes.charCodeAt(pos + 1) << 16) |
      (bytes.charCodeAt(pos + 2) << 8) | bytes.charCodeAt(pos + 3)) >>> 0;
    const type = bytes.slice(pos + 4, pos + 8);
    const start = pos + 8;
    const end = start + length;
    if (end + 4 > bytes.length) return null;
    if (type === "IDAT") {
      pixels.push(bytes.slice(start, end));
    } else if (type === "IHDR" || type === "PLTE" || type === "tRNS") {
      header += `${type}:${length}:${bytes.slice(start, end)};`;
    } else if (type === "IEND") {
      break;
    }
    pos = end + 4;
  }
  if (header.indexOf("IHDR:") !== 0 || pixels.length === 0) return null;
  return header + pixels.join("");
}

/**
 * True when two image sources show the same picture: the same string, or two
 * PNG data URIs with identical pixel data (see pngPixelData).
 *
 * Gemini Enterprise can return one generated image as two session files that
 * differ only in embedded metadata, and that metadata can differ in length,
 * so neither the strings nor their lengths can be compared directly. Remote
 * URLs must match exactly.
 */
export function isSameImageSource(a, b) {
  if (!a || !b) return false;
  if (a === b) return true;
  const pixelsA = pngPixelData(a);
  return pixelsA !== null && pixelsA === pngPixelData(b);
}

/**
 * Image preview cards in a response bubble, leaving out deck preview
 * thumbnails and wrapper cards that contain other cards.
 */
function imagePreviewCards(bubbleEl) {
  if (!bubbleEl || typeof bubbleEl.querySelectorAll !== "function") return [];
  return Array.from(bubbleEl.querySelectorAll(".office-visual-image-card")).filter(
    (card) => !card.closest(".ppt-deck-preview-container") && !card.querySelector(".office-visual-image-card")
  );
}

function previewImageSource(card) {
  const img = card.querySelector("img");
  return img ? img.getAttribute("src") || "" : "";
}

/**
 * Keeps one preview card per picture in a response bubble. The last copy is
 * the one kept, because the Insert buttons place the last image on the slide.
 * Returns how many duplicate cards were removed.
 *
 * This is the instant check (see isSameImageSource);
 * removeDuplicateImagePreviewsByPixels then catches copies that were
 * re-encoded or resized.
 */
export function removeDuplicateImagePreviews(bubbleEl) {
  const cards = imagePreviewCards(bubbleEl);
  if (cards.length < 2) return 0;
  const pixelsOf = (entry) => {
    if (entry.pixels === undefined) entry.pixels = pngPixelData(entry.src);
    return entry.pixels;
  };
  const kept = [];
  let removed = 0;
  for (let i = cards.length - 1; i >= 0; i--) {
    const src = previewImageSource(cards[i]);
    if (!src) continue;
    const entry = { src, pixels: undefined };
    const isCopy = kept.some(
      (k) => k.src === src || (pixelsOf(k) !== null && pixelsOf(k) === pixelsOf(entry))
    );
    if (isCopy) {
      cards[i].remove();
      removed++;
    } else {
      kept.push(entry);
    }
  }
  if (removed > 0) {
    console.log(`[PPT] Image response: removed ${removed} duplicate image preview(s).`);
  }
  return removed;
}

// Decoded-pixel check: both pictures are drawn at most PIXEL_CHECK_MAX_SIDE
// pixels on their longer side. They count as one picture when no more than
// PIXEL_CHECK_MAX_MISMATCH of those pixels differ by more than
// PIXEL_CHECK_TOLERANCE levels (of 255) in any channel. Measured on a 1344x768
// infographic: re-encoded, half-size and JPEG (quality 0.3 to 0.9) copies
// differ in 0.00-0.07% of pixels; one changed 70x24 figure gives 0.19%, new
// text on the same layout 12%, a different picture over 50%.
const PIXEL_CHECK_MAX_SIDE = 256;
const PIXEL_CHECK_TOLERANCE = 16;
const PIXEL_CHECK_MAX_MISMATCH = 0.001;
const PIXEL_CHECK_LOAD_TIMEOUT_MS = 5000;
const pixelChecks = new WeakMap();

/** Resolves true once the image has loaded with real dimensions, false if it fails or times out. */
function whenImageLoaded(img) {
  if (img.complete) return Promise.resolve(img.naturalWidth > 0 && img.naturalHeight > 0);
  return new Promise((resolve) => {
    let timer = null;
    const done = () => {
      clearTimeout(timer);
      img.removeEventListener("load", done);
      img.removeEventListener("error", done);
      resolve(img.complete && img.naturalWidth > 0 && img.naturalHeight > 0);
    };
    img.addEventListener("load", done);
    img.addEventListener("error", done);
    timer = setTimeout(done, PIXEL_CHECK_LOAD_TIMEOUT_MS);
  });
}

function readPixels(img, width, height) {
  try {
    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    const ctx = canvas.getContext("2d", { willReadFrequently: true });
    if (!ctx) return null;
    ctx.imageSmoothingEnabled = true;
    ctx.imageSmoothingQuality = "high";
    ctx.drawImage(img, 0, 0, width, height);
    return ctx.getImageData(0, 0, width, height).data;
  } catch (err) {
    return null;
  }
}

/**
 * Compares two images by what the browser decodes. Resolves to
 * { same, detail }, where detail is a short reason for the console.
 */
async function comparePictures(imgA, imgB) {
  const [loadedA, loadedB] = await Promise.all([whenImageLoaded(imgA), whenImageLoaded(imgB)]);
  if (!loadedA || !loadedB) return { same: false, detail: "an image did not load" };
  const sizeA = `${imgA.naturalWidth}x${imgA.naturalHeight}`;
  const sizeB = `${imgB.naturalWidth}x${imgB.naturalHeight}`;
  const ratioA = imgA.naturalWidth / imgA.naturalHeight;
  const ratioB = imgB.naturalWidth / imgB.naturalHeight;
  if (Math.abs(ratioA - ratioB) > ratioA * 0.01) {
    return { same: false, detail: `different shapes, ${sizeA} vs ${sizeB}` };
  }
  const baseWidth = Math.min(imgA.naturalWidth, imgB.naturalWidth);
  const baseHeight = Math.min(imgA.naturalHeight, imgB.naturalHeight);
  const scale = Math.min(1, PIXEL_CHECK_MAX_SIDE / Math.max(baseWidth, baseHeight));
  const width = Math.max(1, Math.round(baseWidth * scale));
  const height = Math.max(1, Math.round(baseHeight * scale));
  const a = readPixels(imgA, width, height);
  const b = readPixels(imgB, width, height);
  if (!a || !b) return { same: false, detail: "pixels could not be read" };
  let mismatched = 0;
  for (let p = 0; p < a.length; p += 4) {
    if (Math.abs(a[p] - b[p]) > PIXEL_CHECK_TOLERANCE ||
        Math.abs(a[p + 1] - b[p + 1]) > PIXEL_CHECK_TOLERANCE ||
        Math.abs(a[p + 2] - b[p + 2]) > PIXEL_CHECK_TOLERANCE ||
        Math.abs(a[p + 3] - b[p + 3]) > PIXEL_CHECK_TOLERANCE) {
      mismatched++;
    }
  }
  const share = mismatched / (width * height);
  return {
    same: share <= PIXEL_CHECK_MAX_MISMATCH,
    detail: `${(share * 100).toFixed(2)}% of pixels differ (${sizeA} vs ${sizeB}, compared at ${width}x${height})`
  };
}

/**
 * Second pass for copies that removeDuplicateImagePreviews cannot match from
 * the file alone, such as a re-encoded or resized copy: compares the decoded
 * pixels. Only data URI previews are compared, because remote images cannot
 * be read back. The last card stays on screen. Earlier cards are hidden while
 * they are checked, so a duplicate never flashes up, and each one is then
 * either removed or shown again. Runs once per bubble and resolves to how many
 * duplicate cards were removed.
 */
export function removeDuplicateImagePreviewsByPixels(bubbleEl) {
  if (!bubbleEl || typeof document === "undefined") return Promise.resolve(0);
  if (pixelChecks.has(bubbleEl)) return pixelChecks.get(bubbleEl);
  const cards = imagePreviewCards(bubbleEl).filter((card) => /^data:image\//i.test(previewImageSource(card)));
  if (cards.length < 2) return Promise.resolve(0);

  const hidden = new Map();
  cards.slice(0, -1).forEach((card) => {
    hidden.set(card, card.style.display);
    card.style.display = "none";
  });
  const reveal = (card) => {
    if (!hidden.has(card)) return;
    card.style.display = hidden.get(card);
    hidden.delete(card);
  };

  const check = (async () => {
    const kept = [cards[cards.length - 1].querySelector("img")];
    let removed = 0;
    for (let i = cards.length - 2; i >= 0; i--) {
      const card = cards[i];
      const img = card.querySelector("img");
      const reasons = [];
      let match = null;
      for (const keptImg of kept) {
        const result = await comparePictures(keptImg, img);
        if (result.same) {
          match = result;
          break;
        }
        reasons.push(result.detail);
      }
      if (match) {
        hidden.delete(card);
        card.remove();
        removed++;
        console.log(`[PPT] Image response: removed a duplicate image preview (${match.detail}).`);
      } else {
        reveal(card);
        kept.push(img);
        console.log(`[PPT] Image response: showing a different picture (${reasons.join("; ")}).`);
      }
    }
    return removed;
  })();

  const settled = check.then(
    (removed) => {
      Array.from(hidden.keys()).forEach(reveal);
      return removed;
    },
    (err) => {
      Array.from(hidden.keys()).forEach(reveal);
      console.warn("PowerPoint image pixel check error:", err);
      return 0;
    }
  );
  pixelChecks.set(bubbleEl, settled);
  return settled;
}

/**
 * Image responses offer "Insert Image as New Slide" and "Insert Image on
 * Slide" only, so their "Replace with Image" button is removed. Text responses
 * keep "Replace Draft". Returns true when a button was removed.
 */
export function removeReplaceWithImageButton(bubbleEl) {
  if (!bubbleEl || typeof bubbleEl.querySelector !== "function") return false;
  const actions = bubbleEl.querySelector(".response-actions-container");
  if (!actions) return false;
  const button = Array.from(actions.querySelectorAll("button")).find(
    (btn) => (btn.textContent || "").trim() === "Replace with Image"
  );
  if (!button) return false;
  button.remove();
  return true;
}

/**
 * Attaches the Slide Deck Outline Preview to an assistant response bubble.
 */
export function enhanceBubbleWithSlideDeck(bubbleEl, htmlContent, rawText, adapter) {
  if (!bubbleEl) return;

  // Image responses: one preview per picture, and no "Replace with Image".
  try {
    removeDuplicateImagePreviews(bubbleEl);
    removeDuplicateImagePreviewsByPixels(bubbleEl);
    removeReplaceWithImageButton(bubbleEl);
  } catch (err) {
    console.warn("PowerPoint image response tidy-up error:", err);
  }

  if (bubbleEl.querySelector(".ppt-deck-preview-container")) return;

  try {
    const isSummarizeSlides = bubbleEl.dataset?.isSummarizeSlides === "true" || (typeof window !== "undefined" && window.__isSummarizeSlidesAction);
    // Same cap the insert buttons will use, read from the same stamp, so the
    // preview can never advertise slides the build is going to merge away.
    const maxSlides = bubbleEl.dataset?.forceSingleSlide === "true" ? 1 : 0;
    let slides = parseSlides(htmlContent, rawText, { isSummarizeSlides, maxSlides });
    if (!slides || slides.length < 2) return;
    if (isSummarizeSlides && slides.length > 5) {
      slides = slides.slice(0, 5);
    }

    injectPowerPointStyles();

    const container = document.createElement("div");
    container.className = "ppt-deck-preview-container";

    // Header
    const header = document.createElement("div");
    header.className = "ppt-deck-header";
    header.innerHTML = `
      <div class="ppt-deck-title">📊 <span>Presentation Deck Ready</span></div>
      <div class="ppt-deck-badge">${slides.length} Slides</div>
    `;
    container.appendChild(header);

    // Slides List
    const slidesList = document.createElement("div");
    slidesList.className = "ppt-slides-list";

    slides.forEach((slide, idx) => {
      const card = document.createElement("div");
      card.className = "ppt-slide-card";

      let previewContent = "";
      if (slide.visualType === "metric_grid_3col" && slide.visualData && slide.visualData.cards) {
        const cardsHtml = slide.visualData.cards.map(c => `
          <div style="flex: 1; min-width: 70px; background: #f8fafc; border: 1px solid #cbd5e1; border-radius: 4px; padding: 6px; text-align: center;">
            <div style="font-size: 16px; font-weight: 700; color: #0078d4;">${escapeHtml(c.metric || "")}</div>
            <div style="font-size: 10px; font-weight: 700; color: #0f172a; margin-top: 2px;">${escapeHtml(c.title || "")}</div>
          </div>
        `).join("");
        previewContent = `<div style="display: flex; gap: 6px; margin: 4px 0;">${cardsHtml}</div>`;
        if (slide.body) {
          previewContent += `<div style="white-space: pre-wrap; font-size: 10.5px; color: #475569; margin-top: 4px;">${escapeHtml(slide.body.trim())}</div>`;
        }
      } else if (slide.visualType === "before_after" && slide.visualData) {
        previewContent = `
          <div style="display: flex; gap: 6px; margin: 4px 0;">
            <div style="flex: 1; background: #fff5f5; border: 1px solid #fed7d7; border-radius: 4px; padding: 6px;">
              <div style="font-size: 10.5px; font-weight: 700; color: #c53030;">🔴 ${escapeHtml(slide.visualData.before?.title || "Before")}</div>
            </div>
            <div style="flex: 1; background: #f0fff4; border: 1px solid #c6f6d5; border-radius: 4px; padding: 6px;">
              <div style="font-size: 10.5px; font-weight: 700; color: #276749;">🟢 ${escapeHtml(slide.visualData.after?.title || "After")}</div>
            </div>
          </div>
        `;
        if (slide.body) {
          previewContent += `<div style="white-space: pre-wrap; font-size: 10.5px; color: #475569; margin-top: 4px;">${escapeHtml(slide.body.trim())}</div>`;
        }
      } else if (slide.visualType === "three_column_cards" && slide.visualData && slide.visualData.cards) {
        const cardsHtml = slide.visualData.cards.map(c => `
          <div style="flex: 1; min-width: 60px; background: #f8fafc; border: 1px solid #cbd5e1; border-radius: 4px; padding: 6px;">
            <div style="font-size: 8.5px; font-weight: 700; color: #0284c7; text-transform: uppercase;">${escapeHtml(c.pill || "PILLAR")}</div>
            <div style="font-size: 10px; font-weight: 700; color: #0f172a; margin-top: 2px;">${escapeHtml(c.title || "")}</div>
          </div>
        `).join("");
        previewContent = `<div style="display: flex; gap: 6px; margin: 4px 0;">${cardsHtml}</div>`;
        if (slide.body) {
          previewContent += `<div style="white-space: pre-wrap; font-size: 10.5px; color: #475569; margin-top: 4px;">${escapeHtml(slide.body.trim())}</div>`;
        }
      } else if (slide.visualType === "process_timeline" && slide.visualData && slide.visualData.steps) {
        const stepsHtml = slide.visualData.steps.map(s => `
          <div style="flex: 1; min-width: 60px; background: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 4px; padding: 6px;">
            <div style="font-size: 8.5px; font-weight: 700; color: #16a34a; text-transform: uppercase;">STEP ${escapeHtml(s.step || "•")}</div>
            <div style="font-size: 10px; font-weight: 700; color: #0f172a; margin-top: 2px;">${escapeHtml(s.title || "")}</div>
          </div>
        `).join("");
        previewContent = `<div style="display: flex; gap: 6px; margin: 4px 0;">${stepsHtml}</div>`;
        if (slide.body) {
          previewContent += `<div style="white-space: pre-wrap; font-size: 10.5px; color: #475569; margin-top: 4px;">${escapeHtml(slide.body.trim())}</div>`;
        }
      } else if (slide.structuredCards && slide.structuredCards.length >= 2) {
        const cardsHtml = slide.structuredCards.map((c, i) => `
          <div style="flex: 1; min-width: 60px; background: #f8fafc; border: 1px solid #cbd5e1; border-radius: 4px; padding: 6px;">
            <div style="font-size: 8.5px; font-weight: 700; color: #0284c7; text-transform: uppercase;">PILLAR 0${i + 1}</div>
            <div style="font-size: 10px; font-weight: 700; color: #0f172a; margin-top: 2px;">${escapeHtml(c.title || "")}</div>
          </div>
        `).join("");
        previewContent = `<div style="display: flex; gap: 6px; margin: 4px 0;">${cardsHtml}</div>`;
        if (slide.body) {
          previewContent += `<div style="white-space: pre-wrap; font-size: 10.5px; color: #475569; margin-top: 4px;">${escapeHtml(slide.body.trim())}</div>`;
        }
      } else if (slide.tableData && slide.tableData.rows && slide.tableData.rows.length > 0) {
        const headersHtml = (slide.tableData.headers || []).map(h => `<th style="padding: 4px 6px; background: #ffffff; color: #000000; font-weight: 700; text-align: left; font-size: 10px; border: 1px solid #165b7d; border-bottom: 2px solid #165b7d;">${escapeHtml(h)}</th>`).join("");
        const rowsHtml = slide.tableData.rows.map((r, rIdx) => {
          const bg = rIdx % 2 === 0 ? '#e1edf5' : '#ffffff';
          const cells = r.map((c, cIdx) => `<td style="padding: 4px 6px; font-size: 10px; border: 1px solid #165b7d; color: #000000; ${cIdx === 0 ? 'font-weight: 600;' : ''}">${escapeHtml(c)}</td>`).join("");
          return `<tr style="background: ${bg};">${cells}</tr>`;
        }).join("");

        previewContent = `
          <div style="overflow-x: auto; margin: 4px 0;">
            <table style="width: 100%; border-collapse: collapse; border: 1px solid #165b7d; font-size: 10px; line-height: 1.3; background: #ffffff;">
              ${headersHtml ? `<thead><tr>${headersHtml}</tr></thead>` : ''}
              <tbody>${rowsHtml}</tbody>
            </table>
          </div>
        `;
        if (slide.additionalBody) {
          previewContent += `<div style="white-space: pre-wrap; margin-top: 6px; font-size: 10.5px; color: #323130;">${escapeHtml(slide.additionalBody.trim())}</div>`;
        }
      } else {
        const previewBody = slide.body
          ? slide.body.trim()
          : "Full slide content & visual layout prepared.";
        previewContent = `<div style="white-space: pre-wrap;">${escapeHtml(previewBody)}</div>`;
      }

      if (slide.base64Images && slide.base64Images.length > 0 && !previewContent.includes("<img")) {
        previewContent += `<div style="margin-top: 4px; text-align: center;"><img src="${escapeHtml(slide.base64Images[0])}" style="max-height: 70px; border-radius: 4px; border: 1px solid #cbd5e1;" /></div>`;
      }

      card.innerHTML = `
        <div class="ppt-slide-card-header">
          <span class="ppt-slide-num">Slide ${idx + 1}</span>
          <span>${escapeHtml(slide.title)}</span>
        </div>
        ${slide.subtitle ? `<div style="font-size: 10.5px; font-weight: 600; color: #475569; margin: 2px 0 4px 0; padding-left: 2px;">${escapeHtml(slide.subtitle)}</div>` : ''}
        <div class="ppt-slide-preview-body">${previewContent}</div>
      `;
      slidesList.appendChild(card);
    });
    container.appendChild(slidesList);

    // Footer instruction
    const footerHint = document.createElement("div");
    footerHint.className = "ppt-deck-footer-hint";
    footerHint.innerHTML = `👉 Click <b>"+ Insert into Slides"</b> below to create all ${slides.length} slides.`;
    container.appendChild(footerHint);

    // Optional collapsible raw text outline
    const rawDetails = document.createElement("details");
    rawDetails.className = "ppt-raw-details";
    rawDetails.style.cssText = "margin-top: 6px; font-size: 10.5px; color: #444746;";
    rawDetails.innerHTML = `
      <summary style="cursor: pointer; color: #0b57d0; user-select: none; font-size: 10.5px; font-weight: 500;">📄 View Raw Markdown Text</summary>
      <div style="padding: 6px 8px; background: #f8fafd; border: 1px solid #e1e3e1; border-radius: 6px; margin-top: 4px; font-size: 11px; line-height: 1.4; max-height: 150px; overflow-y: auto;">
        ${htmlContent}
      </div>
    `;
    container.appendChild(rawDetails);

    // Update existing action buttons in the bubble & hide raw duplicate text
    const actionsContainer = bubbleEl.querySelector(".response-actions-container");
    // The bubble's own "Verified Sources" panel (direct child), if Gemini cited any.
    const sourcesPanel = Array.from(bubbleEl.children).find(
      (child) => child.classList && child.classList.contains("citation-container")
    ) || null;
    if (actionsContainer) {
      const insertBtn = actionsContainer.querySelector(".insert-btn") || actionsContainer.querySelector(".action-btn.insert") || actionsContainer.querySelector("button.insert");
      if (insertBtn) {
        insertBtn.innerHTML = slides.length === 1 ? `+ Insert as New Slide` : `+ Insert as ${slides.length} New Slides`;
        insertBtn.title = `Insert all ${slides.length} slides into presentation`;
      }
      const replaceBtn = actionsContainer.querySelector(".replace-btn") || actionsContainer.querySelector(".action-btn.replace") || actionsContainer.querySelector("button.replace");
      if (replaceBtn) {
        replaceBtn.remove();
      }

      // Hide all raw markdown text siblings to eliminate duplicate visual text.
      // The "Verified Sources" panel is not duplicate text - keep it, as Word does.
      Array.from(bubbleEl.children).forEach(child => {
        if (child !== actionsContainer && child !== container && child !== sourcesPanel) {
          child.style.display = "none";
        }
      });

      bubbleEl.insertBefore(container, actionsContainer);
      // Sources read after the deck they back, directly above the action buttons.
      if (sourcesPanel) bubbleEl.insertBefore(sourcesPanel, actionsContainer);
    } else {
      Array.from(bubbleEl.children).forEach(child => {
        if (child !== container && child !== sourcesPanel) child.style.display = "none";
      });
      bubbleEl.appendChild(container);
      if (sourcesPanel) bubbleEl.appendChild(sourcesPanel);
    }
    if (sourcesPanel) {
      console.log(`[PPT] Deck preview: showing ${sourcesPanel.querySelectorAll(".citation-card").length} verified source(s) below the deck.`);
    }
  } catch (err) {
    console.warn("enhanceBubbleWithSlideDeck error:", err);
  }
}

/**
 * Initializes a MutationObserver on #chatHistory to automatically enhance
 * all incoming assistant responses with the Slide Deck Preview.
 */
export function initSlidePreviewObserver(adapter) {
  if (typeof document === "undefined") return;

  const chatHistory = document.getElementById("chatHistory");
  if (!chatHistory) return;
  if (chatHistory.__slidePreviewObserverInitialized) return;
  chatHistory.__slidePreviewObserverInitialized = true;

  injectPowerPointStyles();

  const observer = new MutationObserver((mutations) => {
    mutations.forEach((mutation) => {
      mutation.addedNodes.forEach((node) => {
        if (node.nodeType === 1 && node.classList.contains("assistant")) {
          setTimeout(() => {
            const rawText = node.innerText || "";
            const htmlContent = node.innerHTML || "";
            enhanceBubbleWithSlideDeck(node, htmlContent, rawText, adapter);
          }, 50);
        }
      });
    });
  });

  observer.observe(chatHistory, { childList: true, subtree: true });
}

function escapeHtml(str) {
  if (!str) return "";
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
