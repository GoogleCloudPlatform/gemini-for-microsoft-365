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
 * Slide Parser for Microsoft PowerPoint Adapter
 * 
 * Intelligently extracts slide structures (Title, Body Bullets, Visuals, Metadata)
 * from various LLM response formats:
 *   1. Markdown / HTML Headings (## Slide 1: Title, ## Title)
 *   2. Structured Outline Tables (| Slide 1 | Title | Content |)
 *   3. Bulleted / Numbered Outline Lists (• Slide 1: Title — Details)
 *   4. Text delimiter blocks
 *   5. Single slide fallback
 * 
 * @author Sathya AG, Principal Architect, Google
 */
 import { parseChartSpec, renderChartToDataUrl, buildChartPalette } from '../../core/chartRenderer.js';
import { getCachedDeckTheme } from './deckTheme.js';
import { beginPhase } from './buildTimings.js';
import { applyCardLayoutBudget } from './cardBudget.js';

/**
 * Renders a chart using the open deck's accent palette when one has been sampled,
 * so generated charts match the customer's template.
 */
let themedChartCounter = 0;

function renderThemedChart(spec) {
  // Charts rasterize synchronously and land inside the parse phase; counting
  // them separately is the only way to tell parsing from PNG encoding.
  themedChartCounter += 1;
  const endChartPhase = beginPhase(`chart render #${themedChartCounter}`);
  try {
    const theme = getCachedDeckTheme();
    const palette = theme && theme.accentColor ? buildChartPalette(theme.accentColor) : null;
    return renderChartToDataUrl(spec, palette ? { palette } : {});
  } finally {
    endChartPhase(spec && spec.type ? String(spec.type) : "");
  }
}
 import { normalizeEscapedNewlines, parseMarkdown } from '../../core/markdownParser.js';

/**
 * Normalizes visual and layout JSON payloads into standardized archetype structures.
 */
function findPayloadArray(obj, preferredKeys) {
  if (!obj || typeof obj !== "object") return [];
  for (const k of preferredKeys) {
    if (Array.isArray(obj[k]) && obj[k].length > 0) return obj[k];
  }
  for (const k of Object.keys(obj)) {
    if (Array.isArray(obj[k]) && obj[k].length > 0 && obj[k][0] && typeof obj[k][0] === "object") {
      return obj[k];
    }
  }
  return [];
}

export function normalizeVisualPayload(parsed) {
  if (!parsed || typeof parsed !== "object") return null;
  const rawType = (parsed.visualType || parsed.layoutType || "").toLowerCase().trim();
  if (!rawType) return null;

  if (rawType.includes("metric") || rawType.includes("kpi") || rawType.includes("grid")) {
    const rawCards = findPayloadArray(parsed, ["cards", "metrics", "items", "kpis", "stats", "columns", "data", "values"]);
    return {
      visualType: "metric_grid_3col",
      title: parsed.title || "",
      subtitle: parsed.subtitle || "",
      cards: rawCards.map(c => {
        const desc = c.description || c.body || c.text || c.detail || "";
        return {
          metric: c.metric || c.stat || c.value || c.number || c.figure || c.kpi || "",
          title: c.title || c.headline || c.label || c.name || c.header || "",
          subtitle: c.subtitle || c.pill || c.category || "",
          footnote: c.footnote || c.source || c.citation || "",
          bullets: Array.isArray(c.bullets) ? c.bullets : (desc ? [desc] : [])
        };
      })
    };
  }

  if (rawType.includes("before") || rawType.includes("comparison") || rawType.includes("two_col")) {
    return {
      visualType: "before_after",
      title: parsed.title || "",
      subtitle: parsed.subtitle || "",
      before: parsed.before || { title: "Current State", bullets: [] },
      after: parsed.after || { title: "Target State", bullets: [] }
    };
  }

  if (rawType.includes("three_col") || rawType.includes("card") || rawType.includes("pillar")) {
    const rawCards = findPayloadArray(parsed, ["cards", "columns", "pillars", "items", "data"]);
    return {
      visualType: "three_column_cards",
      title: parsed.title || "",
      subtitle: parsed.subtitle || "",
      cards: rawCards.map((c, idx) => {
        const body = c.body || c.description || c.text || (Array.isArray(c.bullets) ? c.bullets.join("; ") : "");
        return {
          pill: c.pill || c.badge || c.subtitle || `PILLAR 0${idx + 1}`,
          title: c.title || c.headline || c.label || c.name || "",
          body: body,
          footnote: c.footnote || c.source || c.citation || "",
          bullets: Array.isArray(c.bullets) ? c.bullets : (body ? [body] : [])
        };
      })
    };
  }

  if (rawType.includes("timeline") || rawType.includes("process") || rawType.includes("roadmap") || rawType.includes("step")) {
    const rawSteps = findPayloadArray(parsed, ["steps", "milestones", "phases", "timeline", "cards", "items", "data"]);
    return {
      visualType: "process_timeline",
      title: parsed.title || "",
      subtitle: parsed.subtitle || "",
      steps: rawSteps.map((s, idx) => ({
        step: s.step || s.number || `0${idx + 1}`,
        title: s.title || s.headline || s.label || s.name || "",
        footnote: s.footnote || s.source || s.citation || "",
        description: s.description || s.body || s.text || (Array.isArray(s.bullets) ? s.bullets.join("; ") : "")
      }))
    };
  }

  if (rawType.includes("summary") || rawType.includes("exec_summary")) {
    return {
      visualType: "executive_summary",
      title: parsed.title || "",
      subtitle: parsed.subtitle || "",
      summary: parsed.summary || parsed.description || parsed.body || parsed.text || ""
    };
  }

  if (rawType.includes("split") || rawType.includes("media") || rawType.includes("insight")) {
    const rawInsights = findPayloadArray(parsed, ["insights", "cards", "items", "points"]);
    return {
      visualType: "split_media_and_insights",
      title: parsed.title || "",
      subtitle: parsed.subtitle || "",
      media: parsed.media || {},
      insights: rawInsights.map(i => ({
        pill: i.pill || i.badge || "KEY INSIGHT",
        title: i.title || i.headline || i.label || "",
        description: i.description || i.body || i.text || ""
      }))
    };
  }

  return null;
}

/**
 * Ensures a visual payload actually contains renderable items. If the model
 * emitted a layout tag in JSON (e.g. {"visualType":"metric_grid_3col"}) with
 * no array items and put the content in markdown bullets below it, hydrates
 * the payload from structured bullets or drops the empty shell so the markdown
 * body is never overwritten with an empty string.
 */
export function hydrateOrDropVisualPayload(visualType, visualData, bodyText) {
  if (!visualType || !visualData) return { visualType: null, visualData: null };

  const structured = extractStructuredBullets(bodyText) || [];

  if (visualType === "metric_grid_3col") {
    if (!Array.isArray(visualData.cards) || visualData.cards.length === 0) {
      if (structured.length > 0) {
        visualData.cards = structured.map(c => {
          const cleanTitle = String(c.title || "").replace(/:$/, "").trim();
          const m = cleanTitle.match(/^([+$#€£¥]?\d[\d.,%xX+\-]*\s*(?:[BMKTbmkt]|GW|MW|%|yrs|x)?|#[0-9]+)\s+(.*)$/);
          return {
            metric: m ? m[1].trim() : (c.pill || cleanTitle || "•"),
            title: m ? m[2].trim() : cleanTitle,
            subtitle: "",
            footnote: "",
            bullets: Array.isArray(c.bullets) && c.bullets.length ? c.bullets : (c.body ? [c.body] : [])
          };
        });
      }
    }
    if (!Array.isArray(visualData.cards) || visualData.cards.length === 0) {
      return { visualType: null, visualData: null };
    }
  } else if (visualType === "three_column_cards") {
    if (!Array.isArray(visualData.cards) || visualData.cards.length === 0) {
      if (structured.length > 0) {
        visualData.cards = structured.map((c, idx) => ({
          pill: c.pill || `PILLAR 0${idx + 1}`,
          title: String(c.title || "").replace(/:$/, "").trim(),
          body: c.body || "",
          footnote: "",
          bullets: Array.isArray(c.bullets) ? c.bullets : (c.body ? [c.body] : [])
        }));
      }
    }
    if (!Array.isArray(visualData.cards) || visualData.cards.length === 0) {
      return { visualType: null, visualData: null };
    }
  } else if (visualType === "process_timeline") {
    if (!Array.isArray(visualData.steps) || visualData.steps.length === 0) {
      if (structured.length > 0) {
        visualData.steps = structured.map((c, idx) => ({
          step: `0${idx + 1}`,
          title: String(c.title || "").replace(/:$/, "").trim(),
          footnote: "",
          description: c.body || ""
        }));
      }
    }
    if (!Array.isArray(visualData.steps) || visualData.steps.length === 0) {
      return { visualType: null, visualData: null };
    }
  } else if (visualType === "before_after") {
    const hasBefore = visualData.before && Array.isArray(visualData.before.bullets) && visualData.before.bullets.length > 0;
    const hasAfter = visualData.after && Array.isArray(visualData.after.bullets) && visualData.after.bullets.length > 0;
    if (!hasBefore && !hasAfter) {
      return { visualType: null, visualData: null };
    }
  } else if (visualType === "executive_summary") {
    if (!visualData.summary || !String(visualData.summary).trim()) {
      if (bodyText && bodyText.trim()) {
        visualData.summary = bodyText.trim();
      } else {
        return { visualType: null, visualData: null };
      }
    }
  }

  return { visualType, visualData };
}

/**
 * JSON.parse that tolerates the damage a rewrite pass typically does to a
 * visual block: smart quotes, trailing commas, comments and stray unescaped
 * quotes inside string values. As a last resort the item objects are salvaged
 * field by field so the slide keeps its steps/cards instead of going blank.
 * @returns {object|null}
 */
function lenientParseVisualJson(rawBlock) {
  const block = String(rawBlock || "").trim();
  if (!block) return null;
  try {
    return JSON.parse(block);
  } catch (_) {}

  const repaired = block
    .replace(/[\u201C\u201D\u201E\u201F\u2033]/g, '"')
    .replace(/[\u2018\u2019\u201A\u201B\u2032]/g, "'")
    .replace(/^\s*\/\/.*$/gm, "")
    .replace(/,\s*([}\]])/g, "$1");
  try {
    return JSON.parse(repaired);
  } catch (_) {}

  const typeMatch = repaired.match(/"(?:visualType|layoutType)"\s*:\s*"([^"]+)"/);
  if (!typeMatch) return null;

  // Each item is a flat object; read its "key": "value" pairs one by one. A
  // value runs to the quote that is followed by a comma, brace or line end, so
  // an unescaped quote inside the text does not end it early.
  const items = [];
  const objectRe = /\{([^{}]*)\}/g;
  let objMatch;
  while ((objMatch = objectRe.exec(repaired)) !== null) {
    const body = objMatch[1];
    if (/"(?:visualType|layoutType)"\s*:/.test(body)) continue;
    const item = {};
    const pairRe = /"([A-Za-z_]+)"\s*:\s*"([\s\S]*?)"\s*(?=,\s*"[A-Za-z_]+"\s*:|,?\s*$)/g;
    let pair;
    while ((pair = pairRe.exec(body)) !== null) {
      item[pair[1]] = pair[2].replace(/\\"/g, '"').replace(/\s+/g, " ").trim();
    }
    if (Object.keys(item).length > 0) items.push(item);
  }
  if (items.length === 0) return null;
  console.warn(`[SlideParser] Visual JSON did not parse; salvaged ${items.length} item(s) from it.`);
  return { visualType: typeMatch[1], items };
}

/**
 * Extracts and normalizes visual JSON payload from raw text.
 */
export function extractVisualPayload(text) {
  if (!text) return null;
  const normalizedText = normalizeEscapedNewlines(text);
  const jsonBlocks = normalizedText.match(/`{2,3}(?:json)?\s*(\{[\s\S]*?\})\s*`{2,3}/gi) || [];
  for (const block of jsonBlocks) {
    try {
      const cleanBlock = block.replace(/^`{2,3}(?:json)?/i, '').replace(/`{2,3}$/, '').trim();
      const parsed = lenientParseVisualJson(cleanBlock);
      const normalized = normalizeVisualPayload(parsed);
      if (normalized) return normalized;
    } catch (_) {}
  }
  const rawJsonMatch = normalizedText.match(/\{[\s\S]*?"(?:visualType|layoutType)"\s*:\s*"(?:metric_grid_3col|before_after|three_column_cards|process_timeline|split_media_and_insights|executive_summary)"[\s\S]*?\}/);
  if (rawJsonMatch) {
    try {
      const parsed = JSON.parse(rawJsonMatch[0]);
      const normalized = normalizeVisualPayload(parsed);
      if (normalized) return normalized;
    } catch (_) {}
  }
  // Last resort: a damaged or unclosed block. Salvage it rather than letting
  // the slide fall back to an empty body.
  const openFence = normalizedText.search(/`{2,3}(?:json)?\s*\{/i);
  const startIdx = openFence >= 0 ? normalizedText.indexOf("{", openFence) : normalizedText.search(/\{\s*"(?:visualType|layoutType)"/);
  if (startIdx >= 0) {
    const tail = normalizedText.slice(startIdx).replace(/`{2,3}[\s\S]*$/, "");
    const salvaged = normalizeVisualPayload(lenientParseVisualJson(tail));
    if (salvaged) return salvaged;
  }
  return null;
}

/**
 * Generates formatted markdown fallback text from structured visual payload.
 */
export function formatVisualPayloadBody(payload) {
  if (!payload) return "";
  if (payload.visualType === "metric_grid_3col") {
    return (payload.cards || []).map(c => `• ${c.metric || ''} ${c.title || ''}: ${(c.bullets || []).join('; ')}`).join('\n');
  }
  if (payload.visualType === "before_after") {
    const beforeList = (payload.before?.bullets || []).map(b => `  - ${b}`).join('\n');
    const afterList = (payload.after?.bullets || []).map(b => `  - ${b}`).join('\n');
    return `• 🔴 BEFORE: ${payload.before?.title || 'Current State'}\n${beforeList}\n\n• 🟢 AFTER: ${payload.after?.title || 'Target State'}\n${afterList}`;
  }
  if (payload.visualType === "three_column_cards") {
    return (payload.cards || []).map(c => `• **${c.title || ''}**: ${c.body || ''}`).join('\n');
  }
  if (payload.visualType === "process_timeline") {
    return (payload.steps || []).map(s => `• **Phase ${s.step}: ${s.title}**: ${s.description}`).join('\n');
  }
  if (payload.visualType === "executive_summary") {
    return payload.summary || "";
  }
  if (payload.visualType === "split_media_and_insights") {
    return (payload.insights || []).map(i => `• **${i.title}**: ${i.description}`).join('\n');
  }
  return "";
}

/**
 * Detects if a slide body has 2-4 structured bullets with bold lead-ins that can be rendered as modern cards.
 */
export function extractStructuredBullets(bodyText) {
  if (!bodyText) return null;

  // 1. Try paragraph-separated blocks (e.g. separated by \n\n)
  const blocks = bodyText.split(/\n\s*\n+/).map(b => b.trim()).filter(Boolean);
  if (blocks.length >= 2 && blocks.length <= 4) {
    const cards = [];
    for (let idx = 0; idx < blocks.length; idx++) {
      const block = blocks[idx];
      const lines = block.split(/\n+/).map(l => l.trim()).filter(Boolean);
      if (lines.length === 0) continue;

      // Case A: Single line with bold title and description: • **Title**: Description
      const singleLineMatch = lines[0].match(/^[•\-*]?\s*(?:\*\*(.*?)\*\*|__(.*?)__|([A-Z0-9+$#€£¥][A-Za-z0-9\s&/%#+$€£¥.,()\-]{1,45}):)\s*[:\-–—]?\s*(.*)$/);
      if (singleLineMatch) {
        const title = (singleLineMatch[1] || singleLineMatch[2] || singleLineMatch[3] || "").trim();
        let body = (singleLineMatch[4] || "").trim();
        if (lines.length > 1) {
          const rest = lines.slice(1).join(" ").trim();
          body = body ? `${body} ${rest}` : rest;
        }
        if (title && body) {
          cards.push({
            pill: `PILLAR 0${idx + 1}`,
            title: title,
            body: body,
            bullets: [body]
          });
          continue;
        }
      }

      // Case B: First line is title (e.g. **ALU API** or ALU API), subsequent lines are body
      if (lines.length >= 2) {
        const firstLineClean = lines[0].replace(/^[•\-*]\s*/, "").replace(/^[#*_`\s]+|[#*_`\s]+$/g, "").trim();
        const restBody = lines.slice(1).join(" ").replace(/^[•\-*]\s*/, "").trim();
        if (firstLineClean && firstLineClean.length <= 45 && restBody) {
          cards.push({
            pill: `PILLAR 0${idx + 1}`,
            title: firstLineClean,
            body: restBody,
            bullets: [restBody]
          });
          continue;
        }
      }
    }

    if (cards.length === blocks.length && cards.length >= 2) {
      return cards;
    }
  }

  // 2. Try single-line bullets (e.g. separated by single newline)
  const lines = bodyText.split(/\n+/).map(l => l.trim()).filter(Boolean);
  if (lines.length >= 2 && lines.length <= 4) {
    const cards = [];
    for (let idx = 0; idx < lines.length; idx++) {
      const line = lines[idx];
      const match = line.match(/^[•\-*]?\s*(?:\*\*(.*?)\*\*|__(.*?)__|([A-Z0-9+$#€£¥][A-Za-z0-9\s&/%#+$€£¥.,()\-]{1,45}):)\s*[:\-–—]?\s*(.*)$/);
      if (match) {
        const title = (match[1] || match[2] || match[3] || "").trim();
        const body = (match[4] || "").trim();
        if (title && body) {
          cards.push({
            pill: `PILLAR 0${idx + 1}`,
            title: title,
            body: body,
            bullets: [body]
          });
        }
      }
    }

    if (cards.length === lines.length && cards.length >= 2) {
      return cards;
    }
  }

  return null;
}

/**
 * Filters out duplicate bullet points that merely repeat rows or headers from an adjacent native table.
 */
export function filterDuplicateTableBullets(bullets, tableData) {
  if (!tableData || !tableData.rows || tableData.rows.length === 0 || !bullets || bullets.length === 0) {
    return bullets;
  }

  const exactTableCells = new Set();
  (tableData.headers || []).forEach(h => {
    const norm = String(h).toLowerCase().trim();
    if (norm.length > 2) exactTableCells.add(norm);
  });
  tableData.rows.forEach(row => {
    row.forEach(cell => {
      const norm = String(cell).toLowerCase().trim();
      if (norm.length > 2) exactTableCells.add(norm);
    });
  });

  return bullets.filter(b => {
    const raw = String(b).trim();
    if (!raw) return false;
    // Remove if it has pipe separators (markdown table remnants or synthetic bullets)
    if (raw.includes(" | ") || raw.startsWith("|") || raw.endsWith("|")) return false;
    // Remove if it is purely table formatting or dashes
    if (/^[-—\s|:]+$/.test(raw)) return false;

    // Remove only if the bullet is an exact verbatim duplicate of a table header or cell
    const cleanNorm = raw.toLowerCase().replace(/[*_`•\-–—]/g, "").trim();
    if (exactTableCells.has(cleanNorm)) {
      return false;
    }
    return true;
  });
}

/**
 * Extracts clean text from a DOM element, ensuring block elements and list items
 * maintain separating whitespace/newlines instead of concatenating adjacent text.
 */
export function getElementMarkdownText(el) {
  if (!el) return "";
  try {
    const clone = el.cloneNode(true);
    clone.querySelectorAll("strong, b").forEach(b => {
      const t = (b.textContent || "").trim();
      if (t) b.replaceWith(`**${t}**`);
    });
    return getElementCleanText(clone);
  } catch (_) {
    return getElementCleanText(el);
  }
}

export function getElementCleanText(el) {
  if (!el) return "";
  if (el.innerText) return el.innerText;
  try {
    const clone = el.cloneNode(true);
    const blockEls = clone.querySelectorAll("p, div, li, br, tr, td, th");
    blockEls.forEach(b => {
      if (b.tagName === "BR") {
        b.replaceWith("\n");
      } else {
        b.insertAdjacentText("afterend", "\n");
      }
    });
    return (clone.textContent || "").trim();
  } catch (_) {
    return (el.textContent || "").trim();
  }
}

export function extractSlideMetadataAndBullets(rawLines, initialSubtitle = "") {
  let subtitle = initialSubtitle || "";
  let takeaway = "";
  let visualConcept = "";
  let color = null;
  let titleSize = 36;
  let subtitleSize = 18;
  const contentBullets = [];

  for (const rawLine of rawLines) {
    if (!rawLine) continue;
    let line = rawLine.trim();
    if (!line) continue;

    // Check for Subtitle:
    // Matches "#### Subtitle", "### Subtitle", "Subtitle: ...", "*Subtitle:* ...", "**Subtitle:** ..."
    const subMatch = line.match(/^(?:#{3,6}\s*|(?:Sub-?title|Subtitle\s*Text):\s*)(.*)$/i);
    if (subMatch && !subtitle) {
      const cleanSub = subMatch[1].replace(/^[#*_`\s]+|[#*_`\s]+$/g, "").trim();
      if (cleanSub) {
        subtitle = cleanSub;
        continue;
      }
    }

    // Check for Takeaway / Key Takeaway:
    const takeMatch = line.match(/^(?:(?:Key\s*)?Takeaway|Takeaway\s*Text):\s*(.*)$/i);
    if (takeMatch && !takeaway) {
      const cleanTake = takeMatch[1].replace(/^[#*_`\s]+|[#*_`\s]+$/g, "").trim();
      if (cleanTake) {
        takeaway = cleanTake;
        continue;
      }
    }

    const colorMatch = line.match(/^(?:Color|Colour|Color\s*Scheme|Palette|Theme\s*Color):\s*(.*)$/i);
    if (colorMatch) {
      const colorVal = colorMatch[1].trim();
      const hex = colorVal.match(/#[A-Fa-f0-9]{6}/);
      if (hex) {
        color = hex[0];
      } else {
        const lower = colorVal.toLowerCase();
        if (lower.includes("blue")) color = "#004E8C";
        else if (lower.includes("green")) color = "#107C10";
        else if (lower.includes("red") || lower.includes("crimson")) color = "#A80000";
        else if (lower.includes("gold") || lower.includes("yellow")) color = "#C19C00";
        else if (lower.includes("purple") || lower.includes("violet")) color = "#5C2D91";
        else if (lower.includes("teal") || lower.includes("cyan")) color = "#008272";
        else if (lower.includes("orange") || lower.includes("brown") || lower.includes("sepia")) color = "#D83B01";
      }
      continue;
    }

    const tSizeMatch = line.match(/^Title\s*(?:Font\s*)?Size(?:\s*\(pt\))?:\s*(\d+)/i);
    if (tSizeMatch) {
      titleSize = parseInt(tSizeMatch[1], 10);
      continue;
    }

    const sSizeMatch = line.match(/^Subtitle\s*(?:Font\s*)?Size(?:\s*\(pt\))?:\s*(\d+)/i);
    if (sSizeMatch) {
      subtitleSize = parseInt(sSizeMatch[1], 10);
      continue;
    }

    const visMatch = line.match(/^(?:Visual(?:\s*Concept|\s*Description|\s*Prompt|\s*Idea)?|Image(?:\s*Prompt|\s*Concept|\s*Description)?):\s*(.*)$/i);
    if (visMatch) {
      visualConcept = visMatch[1].trim();
      continue;
    }

    if (/^(?:Main\s*)?Content:?$/i.test(line)) {
      continue;
    }

    if (/^(?:Layout|Slide\s*Layout|Template|Design\s*Theme):/i.test(line)) {
      continue;
    }

    // Ignore lines that are only backticks, markdown markers, quotes, or whitespace
    if (/^[`'"*#_~>|\-\s•]+$/.test(line) || line.startsWith("```") || line.startsWith("{") || line.startsWith("}") || line.startsWith('"') || line.startsWith("|") || /^{.*}$/.test(line)) {
      continue;
    }

    // Ignore raw visual markers leaked from pseudo-formatting
    if (/metric\s*grid|comparison\s*card/i.test(line)) {
      continue;
    }

    // Clean bullet text (preserve leading **bold** and #1 rankings):
    let cleanBullet = line
      .replace(/^[-•*]\s+/, "")
      .replace(/^#{1,6}\s+/, "")
      .replace(/^[`'"]+|[`'"]+$/g, "")
      .trim();

    // Must contain substantive alphanumeric text
    if (!/[a-zA-Z0-9]/.test(cleanBullet)) {
      continue;
    }

    // Check if it's a Subtitle or Takeaway after stripping bullet marker
    if (/^(?:Sub-?title|Subtitle\s*Text):\s*/i.test(cleanBullet) && !subtitle) {
      subtitle = cleanBullet.replace(/^(?:Sub-?title|Subtitle\s*Text):\s*/i, "").replace(/^[#*_`\s]+|[#*_`\s]+$/g, "").trim();
      continue;
    }
    if (/^[_*`\s]*(?:Key\s*)?Takeaway[_*`\s]*:\s*/i.test(cleanBullet) && !takeaway) {
      takeaway = cleanBullet.replace(/^[_*`\s]*(?:Key\s*)?Takeaway[_*`\s]*:\s*/i, "").replace(/^[#*_`\s]+|[#*_`\s]+$/g, "").trim();
      continue;
    }

    // Detect section heading or subtitle line (e.g. ### Strategic Analysis and Key Metrics)
    // Only infer subtitle if one has not already been found AND it's not a narrative sentence with a period
    if (!subtitle && (/^#{2,4}\s+/.test(line) || (!line.startsWith("-") && !line.startsWith("*") && !line.startsWith("•") && cleanBullet.length < 50 && !/[.!?]$/.test(cleanBullet)))) {
      subtitle = cleanBullet.replace(/^[#*_`\s]+|[#*_`\s]+$/g, "").trim();
      continue;
    }

    // Strip wrapping markdown italic/bold underscores or asterisks
    if ((cleanBullet.startsWith("_") && cleanBullet.endsWith("_") && cleanBullet.length > 2) ||
        (cleanBullet.startsWith("*") && cleanBullet.endsWith("*") && cleanBullet.length > 2)) {
      cleanBullet = cleanBullet.slice(1, -1).trim();
    }

    // Strip any lingering unparsed markdown formatting markers
    cleanBullet = cleanBullet.replace(/\*\*/g, "").replace(/__/g, "").trim();

    if (cleanBullet) {
      contentBullets.push(`•  ${cleanBullet}`);
    }
  }

  return {
    subtitle,
    takeaway,
    visualConcept,
    color,
    titleSize: titleSize || 36,
    subtitleSize: subtitleSize || 18,
    body: contentBullets.length > 0 ? contentBullets.join("\n\n") : "• Executive slide content"
  };
}

/**
 * Extracts clean structured table content and formats it into executive comparison bullets.
 */
export function extractTableContent(tableEl) {
  if (!tableEl) return null;
  const rows = Array.from(tableEl.querySelectorAll("tr"));
  if (rows.length === 0) return null;

  const headers = [];
  const headerRow = tableEl.querySelector("thead tr") || rows[0];
  if (headerRow) {
    const ths = Array.from(headerRow.querySelectorAll("th, td"))
      .map(c => (c.innerText || c.textContent || "").trim());
    if (ths.length > 0) headers.push(...ths);
  }

  const dataRows = [];
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    if (row === headerRow && tableEl.querySelector("th")) continue;
    const cells = Array.from(row.querySelectorAll("td, th"))
      .map(c => (c.innerText || c.textContent || "").trim());
    if (cells.length > 0 && cells.some(c => c.length > 0)) {
      dataRows.push(cells);
    }
  }

  // Format into clean structured comparison bullets for presentation slides
  const formattedBullets = [];
  dataRows.forEach(row => {
    if (row.length === 1) {
      formattedBullets.push(`• ${row[0]}`);
    } else if (row.length === 2) {
      formattedBullets.push(`• ${row[0]}: ${row[1]}`);
    } else {
      const itemName = row[0];
      const details = [];
      for (let c = 1; c < row.length; c++) {
        const h = headers[c] ? `${headers[c]}: ` : "";
        details.push(`${h}${row[c]}`);
      }
      formattedBullets.push(`• ${itemName}: ${details.join("  |  ")}`);
    }
  });

  return {
    headers,
    dataRows,
    bullets: formattedBullets,
    bulletText: formattedBullets.join("\n\n")
  };
}

/**
 * Parses raw markdown table (| Header 1 | Header 2 |) into structured columns and rows.
 */
export function parseMarkdownTable(text) {
  if (!text || !text.includes("|")) return null;
  const lines = text.split(/\r?\n/).map(l => l.trim()).filter(l => l.startsWith("|") && l.endsWith("|"));
  if (lines.length < 2) return null;

  const validRows = [];
  for (const line of lines) {
    const cells = line.split("|").slice(1, -1).map(c => c.trim().replace(/\*\*/g, ""));
    // Filter out separator lines like |---|---|
    if (cells.length > 0 && !cells.every(c => /^[-:\s]+$/.test(c))) {
      validRows.push(cells);
    }
  }

  if (validRows.length < 2) return null;

  const headers = validRows[0];
  const dataRows = validRows.slice(1);

  const formattedBullets = [];
  dataRows.forEach(row => {
    if (row.length === 1) {
      formattedBullets.push(`• ${row[0]}`);
    } else if (row.length === 2) {
      formattedBullets.push(`• ${row[0]}: ${row[1]}`);
    } else {
      const itemName = row[0];
      const details = [];
      for (let c = 1; c < row.length; c++) {
        const h = headers[c] ? `${headers[c]}: ` : "";
        details.push(`${h}${row[c]}`);
      }
      formattedBullets.push(`• ${itemName}: ${details.join("  |  ")}`);
    }
  });

  return {
    headers,
    dataRows,
    bullets: formattedBullets,
    bulletText: formattedBullets.join("\n\n")
  };
}

/**
 * Selects only the genuine slide boundary heading elements from a DOM fragment.
 * Prevents H3/H4 subheadings, subtitles, or document titles from being mistakenly
 * promoted to standalone slides.
 */
function getSlideHeaderElements(tempDiv) {
  const allHeaders = Array.from(tempDiv.querySelectorAll("h1, h2, h3, h4"));
  if (allHeaders.length === 0) return [];

  const h2s = Array.from(tempDiv.querySelectorAll("h2"));
  const slideNumHeaders = allHeaders.filter(h => {
    const text = (h.innerText || h.textContent || "").trim();
    return /^(?:Slide\s*\d+|#+\s*Slide\s*\d+)/i.test(text) || /\bSlide\s*\d+\b/i.test(text);
  });

  // 1. If H2 headers exist and cover at least as many sections as slideNumHeaders,
  // prefer H2s so unnumbered conclusion, takeaway, or summary slides are never dropped!
  if (h2s.length >= 2 && h2s.length >= slideNumHeaders.length) {
    const h1 = tempDiv.querySelector("h1");
    if (h1) {
      let hasBody = false;
      let sib = h1.nextElementSibling;
      while (sib && sib !== h2s[0]) {
        const txt = (sib.innerText || sib.textContent || "").trim();
        if (txt.length > 0) {
          hasBody = true;
          break;
        }
        sib = sib.nextElementSibling;
      }
      const firstH2Text = (h2s[0].innerText || h2s[0].textContent || "").trim();
      const firstH2IsSlide1 = /^(?:Slide\s*1\b|#+\s*Slide\s*1\b)/i.test(firstH2Text);
      if (hasBody && !firstH2IsSlide1) {
        return [h1, ...h2s];
      }
    }
    return h2s;
  }

  // 2. Check if headers have explicit Slide numbering (e.g. "Slide 1", "Slide 2", "## Slide 1")
  if (slideNumHeaders.length >= 2) {
    return slideNumHeaders;
  }

  // 3. Fallback to H2 headers if present
  if (h2s.length >= 2) {
    return h2s;
  }

  // 3. Check for H1 headers
  const h1s = Array.from(tempDiv.querySelectorAll("h1"));
  if (h1s.length >= 2) {
    return h1s;
  }

  // 4. Check for H3 headers (only if no H1 or H2 are defining slides)
  const h3s = Array.from(tempDiv.querySelectorAll("h3"));
  if (h3s.length >= 2) {
    return h3s;
  }

  return allHeaders.slice(0, 1);
}

/**
 * Parses Summarize Slides content into up to 5 distinct slides:
 * - Table slides (clean canvas with native table)
 * - Category breakdown with visual chart & secondary table
 * - Dedicated Key Takeaways slide with strategic bullets
 * Capped at 5 slides maximum.
 */
function parseSummarizeDeck(tempDiv, allImages = [], rawText = "", options = {}) {
  const slides = [];
  const tables = Array.from(tempDiv.querySelectorAll("table"));
  const chartContainers = Array.from(tempDiv.querySelectorAll(".rendered-chart-container, [data-chart-title]"));
  const bulletLists = Array.from(tempDiv.querySelectorAll("ul, ol"));

  // If there are no tables and no bullet lists, fallback to standard parsing
  if (tables.length === 0 && bulletLists.length === 0) {
    return null;
  }

  // 1. Process Tables into slides
  for (let i = 0; i < tables.length; i++) {
    const tableEl = tables[i];
    const tbl = extractTableContent(tableEl);
    if (!tbl || tbl.dataRows.length === 0) continue;

    let title = "";
    let associatedImgs = [];
    
    // Check previous siblings for title and preceding chart
    let prev = tableEl.previousElementSibling;
    while (prev) {
      if (prev.classList && prev.classList.contains("rendered-chart-container")) {
        const cTitle = prev.getAttribute("data-chart-title");
        if (cTitle && !title) title = cTitle;
        const cImgs = Array.from(prev.querySelectorAll("img")).map(img => img.src || img.getAttribute("src") || "").filter(s => s && s.length > 50);
        associatedImgs.push(...cImgs);
      }
      const txt = (prev.innerText || prev.textContent || "").trim();
      if (txt && !txt.startsWith("Verified Sources") && !txt.includes("Zoom & Review") && !txt.includes("🔍") && !title) {
        title = txt;
        break;
      }
      prev = prev.previousElementSibling;
    }

    if (!title && tbl.headers.length > 0) {
      title = i === 0 ? "Executive Slide Summary" : `${tbl.headers[0]} Breakdown`;
    }

    // If this table is the second table and no chart images were attached yet, but allImages has a chart
    if (i === 1 && associatedImgs.length === 0 && allImages.length > 0) {
      associatedImgs = [allImages[0]];
    }

    slides.push({
      slideNumber: slides.length + 1,
      title: cleanSlideTitle(title || `Summary Slide ${slides.length + 1}`, slides.length + 1),
      subtitle: "",
      visualConcept: "",
      visualType: null,
      visualData: null,
      color: null,
      titleSize: 36,
      subtitleSize: 20,
      body: "",
      additionalBody: "",
      tableData: {
        headers: tbl.headers,
        rows: tbl.dataRows
      },
      base64Images: associatedImgs
    });
  }

  // 2. Process Key Takeaways into a dedicated single slide
  const takeawayBullets = [];
  bulletLists.forEach(listEl => {
    const items = Array.from(listEl.querySelectorAll("li")).map(li => {
      const txt = (li.innerText || li.textContent || "").trim();
      return txt.startsWith("•") ? txt : `• ${txt}`;
    }).filter(b => b.length > 10 && !b.startsWith("|"));
    takeawayBullets.push(...items);
  });

  // If no <ul> was found, extract bullet lines from paragraphs
  if (takeawayBullets.length === 0) {
    const paras = Array.from(tempDiv.querySelectorAll("p, div")).map(p => (p.innerText || p.textContent || "").trim());
    for (const p of paras) {
      if (/^(?:•|[-*]|\*\*|\b[A-Z][a-zA-Z\s]+:)/.test(p) && p.length > 25 && !p.startsWith("|")) {
        const clean = p.replace(/^[-*•]\s*/, "");
        takeawayBullets.push(`• ${clean}`);
      }
    }
  }

  if (takeawayBullets.length > 0) {
    let takeawayTitle = "Executive Summary: Key Takeaways";
    const headings = Array.from(tempDiv.querySelectorAll("h1, h2, h3, h4, strong, b"));
    for (const h of headings) {
      const hText = (h.innerText || h.textContent || "").trim();
      if (hText.toLowerCase().includes("takeaway") || hText.toLowerCase().includes("key takeaway")) {
        takeawayTitle = cleanSlideTitle(hText, slides.length + 1);
        break;
      }
    }

    // Cap at most 4 concise, impactful takeaways for a clean, professional executive slide
    const cappedBullets = takeawayBullets.slice(0, 4);
    slides.push({
      slideNumber: slides.length + 1,
      title: takeawayTitle,
      subtitle: "Strategic Highlights & Next Steps",
      visualConcept: "",
      visualType: null,
      visualData: null,
      color: null,
      titleSize: 36,
      subtitleSize: 20,
      body: cappedBullets.join("\n\n"),
      additionalBody: cappedBullets.join("\n\n"),
      tableData: null,
      base64Images: []
    });
  }

  // Hard cap summarize deck at strictly 5 slides max.
  // NOTE: `slides` is a const binding, so trim in place rather than reassigning
  // (reassignment threw "Assignment to constant variable" and aborted the parse).
  if (slides.length > 5) {
    slides.splice(5);
  }

  if (slides.length >= 2) {
    return slides;
  }
  return null;
}

/**
 * Parses HTML or raw Markdown text into an array of slide objects:
 * [{ title: string, subtitle: string, body: string, color: string, titleSize: number, subtitleSize: number, base64Images: string[], slideNumber: number }]
 */
export function parseSlides(htmlContent, rawText = "", options = {}) {
  if (!htmlContent && !rawText) return [];

  let cleanHtml = normalizeEscapedNewlines(htmlContent || "");
  let cleanRaw = normalizeEscapedNewlines(rawText || "");

  // If htmlContent is plain markdown (e.g. contains ## or --- without HTML tags)
  // or lacks HTML headings while cleanRaw contains slide headings, compile with parseMarkdown
  if (cleanRaw && (!cleanHtml || !/<[a-z][\s\S]*>/i.test(cleanHtml) || (!cleanHtml.includes("<h1") && !cleanHtml.includes("<h2") && /(?:^|\n)##\s+/i.test(cleanRaw)))) {
    cleanHtml = parseMarkdown(cleanRaw);
  }

  let highResChartSrc = "";
  const rawDomImages = [];
  let remainingText = "";

  const tempDiv = document.createElement("div");
  tempDiv.innerHTML = cleanHtml || cleanRaw;

  // Clean UI buttons, zoom controls, citation callouts, action toolbars, and preview containers
  const uiCallouts = tempDiv.querySelectorAll("button, .img-zoom-btn, .img-action-btn-zoom, .action-btn, [class*='zoom'], blockquote, .note, .ppt-deck-preview-container, .response-actions-container, [style*='background-color:#f0f6ff']");
  uiCallouts.forEach(n => n.remove());

  remainingText = (tempDiv.innerText || tempDiv.textContent || "").replace(/🔍|Zoom|Review|Visual/gi, "").trim();

  // Extract all images upfront from HTML DOM, markdown images, and data URIs
  const highResChartEl = tempDiv.querySelector(".rendered-chart-container img");
  highResChartSrc = highResChartEl ? (highResChartEl.src || highResChartEl.getAttribute("src") || "") : "";

  const domImgs = Array.from(tempDiv.querySelectorAll("img"))
    .map(img => img.src || img.getAttribute("src") || "")
    .filter(s => s && s.length > 50);
  rawDomImages.push(...domImgs);

  const combinedSearch = ((htmlContent || '') + ' ' + (rawText || ''));
  const hasDistinctNonChartImageIntent = /(?:photo|photograph|portrait|illustration|logo|camera|scenery|picture of|image of a)/i.test(combinedSearch);

  const allImages = [];
  if (highResChartSrc) {
    allImages.push(highResChartSrc);
    // If the request requested distinct non-chart images (e.g. photos, logos), include them as well
    if (hasDistinctNonChartImageIntent) {
      for (const img of rawDomImages) {
        if (img !== highResChartSrc && !allImages.includes(img)) {
          allImages.push(img);
        }
      }
    }
  } else {
    for (const img of rawDomImages) {
      if (!allImages.includes(img)) allImages.push(img);
    }
    const mdImgMatches = combinedSearch.match(/!\[.*?\]\(\s*<?(data:image\/[^;]+;base64,[A-Za-z0-9+/=]+|https?:\/\/[^\s\)>]+)/gi) || [];
    for (const m of mdImgMatches) {
      const u = m.replace(/^!\[.*?\]\(\s*<?/i, '').replace(/>?\s*$/i, '').trim();
      if (u && !allImages.includes(u)) {
        allImages.push(u);
      }
    }

    const rawDataMatches = combinedSearch.match(/data:image\/(?:png|jpeg|jpg|webp|gif);base64,[A-Za-z0-9+/=]{100,}/gi) || [];
    for (const d of rawDataMatches) {
      if (!allImages.includes(d)) {
        allImages.push(d);
      }
    }
  }

  // Handle explicit or detected image-only insertion (e.g. from Zoom modal, image card, or image generation prompt)
  const lastPrompt = typeof window !== "undefined" ? (window.__lastUserPrompt || "") : "";
  const isImageRequestPrompt = /(?:generate|create|make|draw|show|render|insert|add)\s+(?:an?\s+)?(?:image|picture|photo|visual|illustration|graphic)/i.test(lastPrompt) ||
                               /(?:image|picture|photo|visual|illustration)\s+(?:of|for|showing|depicting)/i.test(lastPrompt);
  const isImageOnly = options?.imageOnly || (allImages.length > 0 && isImageRequestPrompt) || (allImages.length > 0 && !rawText.includes("##") && remainingText.length < 5);
  if (isImageOnly && allImages.length > 0) {
    let imagesForSlide = allImages;
    if (options?.imageOnly) {
      // Never test a base64 payload for the word "chart". Lower-casing it
      // collapses the alphabet to ~38 symbols, so the literal "chart" turns up
      // by chance in a meaningful fraction of large images and silently
      // disqualified the very picture the user asked to insert. Only the chart
      // the renderer itself produced, or a URL that names one, counts.
      const nonChartImgs = allImages.filter(img =>
        img !== highResChartSrc && !/^https?:\/\/[^\s]*chart/i.test(img)
      );
      if (nonChartImgs.length > 0) {
        imagesForSlide = nonChartImgs;
      }
    }
    return [{
      slideNumber: 1,
      title: "",
      subtitle: "",
      visualConcept: "",
      color: null,
      titleSize: 36,
      subtitleSize: 20,
      body: "",
      additionalBody: "",
      tableData: null,
      base64Images: imagesForSlide,
      imageOnly: true
    }];
  }

  const isSummarizeSlides = Boolean(
    options?.isSummarizeSlides ||
    (typeof window !== "undefined" && window.__isSummarizeSlidesAction)
  );

  const headerEls = getSlideHeaderElements(tempDiv);

  // If Summarize Slides mode is active without explicit slide headers, check if multiple sections
  // (tables, chart, takeaways) are present. If so, parse into up to 5 comprehensive slides!
  // CRITICAL: When explicit slide headers (## Slide 1, ## Slide 2) exist, Strategy 1 takes precedence!
  if (isSummarizeSlides && headerEls.length < 2) {
    const summarizeDeck = parseSummarizeDeck(tempDiv, allImages, rawText, options);
    if (summarizeDeck && summarizeDeck.length >= 2) {
      return finalizeSlides(summarizeDeck, allImages, rawText, options);
    }
  }

  // -------------------------------------------------------------
  // Strategy 0: Executive Visual JSON (3-Column Metric Grid, Before/After)
  // -------------------------------------------------------------
  const fullSearchText = (rawText || "") + "\n" + (tempDiv.innerText || "");
  let visualPayload = extractVisualPayload(fullSearchText);

  // Only trigger single-slide Strategy 0 if this is NOT a multi-slide document
  const isMultiSlideDoc = (tempDiv.querySelectorAll("h1, h2, h3").length >= 2) ||
    ((rawText || "").match(/(?:^|\n)##\s+/g) || []).length >= 2;

  if (!isMultiSlideDoc && visualPayload) {
    const formattedBody = formatVisualPayloadBody(visualPayload);

    return finalizeSlides([{
      slideNumber: 1,
      title: visualPayload.title || "Executive Visual",
      subtitle: visualPayload.subtitle || "",
      visualType: visualPayload.visualType,
      visualData: visualPayload,
      body: formattedBody,
      base64Images: allImages
    }], allImages, rawText, options);
  }

  // -------------------------------------------------------------
  // Strategy 1: Explicit Slide Headings
  // -------------------------------------------------------------
  // (headerEls was already resolved above via getSlideHeaderElements)
  
  // Check if there is an explicit multi-slide outline table or list in the document
  const hasOutlineTable = Array.from(tempDiv.querySelectorAll("table")).some(t => {
    const rows = Array.from(t.querySelectorAll("tr"));
    if (rows.length < 2) return false;
    const hText = (rows[0].innerText || rows[0].textContent || "").toLowerCase();
    return hText.includes("slide #") || hText.includes("slide number") || hText.includes("slide title") ||
      rows.some(r => Array.from(r.querySelectorAll("td, th")).some(c => /^slide\s*\d+/i.test(c.textContent.trim())));
  });

  const hasOutlineList = Array.from(tempDiv.querySelectorAll("li")).filter(li => 
    /^(?:Slide\s*\d+|\*\*Slide\s*\d+\*\*)\s*[:\-–—]/i.test((li.innerText || li.textContent || "").trim())
  ).length >= 2;

  // Activate Strategy 1 if multiple headers exist, OR if 1 header exists without a multi-slide outline table/list
  const shouldRunStrategy1 = headerEls.length >= 2 || (headerEls.length === 1 && !hasOutlineTable && !hasOutlineList);

  if (shouldRunStrategy1) {
    const slides = [];

    // Check if there is an Executive Summary or Key Takeaways section preceding the first slide header
    let preSummarySlide = null;
    if (headerEls.length >= 1) {
      const firstHeaderText = (headerEls[0].innerText || headerEls[0].textContent || "").trim();
      const isFirstHeaderSlide1 = /^(?:Slide\s*1\b|#+\s*Slide\s*1\b)/i.test(firstHeaderText);

      if (isFirstHeaderSlide1 || isSummarizeSlides || headerEls.length >= 2) {
        const preElements = [];
        let pNode = headerEls[0].previousElementSibling;
        while (pNode) {
          preElements.unshift(pNode);
          pNode = pNode.previousElementSibling;
        }

        let preTitle = "";
        let preSubtitle = "";
        const preBullets = [];
        const preImgs = [];

        for (const el of preElements) {
          const imgs = Array.from(el.querySelectorAll("img")).map(img => img.src || img.getAttribute("src") || "").filter(s => s && s.length > 50);
          if (el.tagName === "IMG") {
            const s = el.src || el.getAttribute("src") || "";
            if (s.length > 50) imgs.push(s);
          }
          if (imgs.length > 0) preImgs.push(...imgs);

          const txt = (el.innerText || el.textContent || "").trim();
          if (!txt || txt.startsWith("Verified Sources") || txt.includes("Zoom & Review") || txt.includes("🔍")) continue;

          // Check if this element is an Executive Summary or Key Takeaways heading
          if (/takeaways?|executive\s+summary/i.test(txt) && !preTitle) {
            preTitle = cleanSlideTitle(txt, 1);
            continue;
          }

          // Check if this element is a deck title/subtitle
          if (/^[📊📈⚡🎯💡]|\bdecisions?\b|\banalysis\b|\bstewardship\b|\breport\b/i.test(txt) && !preTitle && !preSubtitle && txt.length < 120 && !txt.includes(":")) {
            if (!preSubtitle && preTitle) {
              preSubtitle = txt;
            }
            continue;
          }

          // Check for bullet lists (ul/ol)
          if (el.tagName === "UL" || el.tagName === "OL") {
            Array.from(el.querySelectorAll("li")).forEach(li => {
              const bTxt = (li.innerText || li.textContent || "").trim();
              if (bTxt.length > 5) {
                preBullets.push(bTxt.startsWith("•") ? bTxt : `• ${bTxt}`);
              }
            });
            continue;
          }

          // Check for bullet-like paragraphs
          if (/^(?:•|[-*]|\*\*|\b[A-Z][a-zA-Z\s]+:)/.test(txt) && txt.length > 20 && !txt.startsWith("|")) {
            const cleanBullet = txt.replace(/^[-*•]\s*/, "");
            preBullets.push(`• ${cleanBullet}`);
            continue;
          }

          // If we already have a preTitle and no subtitle yet, a short descriptive line is the subtitle
          if (preTitle && !preSubtitle && txt.length < 150 && !txt.includes(":")) {
            preSubtitle = txt;
          }
        }

        // If we found an explicit Executive Summary title or substantive bullets (>= 2)
        if (preBullets.length >= 2 || (preTitle && preBullets.length >= 1)) {
          preSummarySlide = {
            slideNumber: 1,
            title: cleanSlideTitle(preTitle || "Executive Summary: Key Takeaways", 1),
            subtitle: preSubtitle || "Strategic Highlights & Takeaways",
            visualConcept: "",
            visualType: null,
            visualData: null,
            color: null,
            titleSize: 36,
            subtitleSize: 20,
            body: preBullets.slice(0, 5).join("\n\n"),
            additionalBody: preBullets.length > 5 ? preBullets.slice(5).join("\n\n") : "",
            tableData: null,
            base64Images: preImgs
          };
          slides.push(preSummarySlide);
        }
      }
    }

    for (let i = 0; i < headerEls.length; i++) {
      const h = headerEls[i];
      const nextHeader = headerEls[i + 1] || null;
      const slideNum = slides.length + 1;
      const rawTitle = (h.innerText || h.textContent || "").trim();
      const title = cleanSlideTitle(rawTitle, slideNum);

      const allBodyLines = [];
      const additionalBodyLines = [];
      const sectionImgs = [];
      const sectionRawLines = [];
      let sectionSubtitle = "";
      let sectionTableData = null;
      // CRITICAL: These must be declared BEFORE the element walk below, which assigns
      // them when it encounters a `.rendered-visual-container` payload. Declaring them
      // afterwards silently discarded every decoded visual payload.
      let sectionVisualType = null;
      let sectionVisualData = null;

      // If this is the first header, collect any chart images or illustrations that appear before it
      if (i === 0 && !preSummarySlide) {
        let prev = h.previousElementSibling;
        while (prev) {
          const prevImgs = Array.from(prev.querySelectorAll("img"))
            .map(img => img.src || img.getAttribute("src") || "")
            .filter(s => s && s.length > 50);
          if (prev.tagName === "IMG") {
            const pSrc = prev.src || prev.getAttribute("src") || "";
            if (pSrc.length > 50) prevImgs.push(pSrc);
          }
          if (prevImgs.length > 0) sectionImgs.unshift(...prevImgs);
          prev = prev.previousElementSibling;
        }
      }

      let curr = h.nextElementSibling;

      while (curr && curr !== nextHeader && !headerEls.includes(curr)) {
        // Skip verified source footers or hr elements
        if ((curr.innerText && curr.innerText.includes("Verified Sources")) || curr.tagName === "HR") {
          curr = curr.nextElementSibling;
          continue;
        }

        // Process rendered visual containers (metric_grid_3col, three_column_cards, etc.)
        // IMPORTANT: Extract payload and DO NOT include the taskpane card's text into allBodyLines/sectionRawLines,
        // because that causes helper headers (e.g. "Strategic Pillars: 🏛️ Core Pillars") to be parsed as slide content.
        if (curr.classList && curr.classList.contains("rendered-visual-container") || curr.querySelector && curr.querySelector(".rendered-visual-container")) {
          const visEl = curr.classList && curr.classList.contains("rendered-visual-container") ? curr : curr.querySelector(".rendered-visual-container");
          const payloadAttr = visEl.getAttribute("data-visual-payload");
          if (payloadAttr) {
            try {
              const rawDecoded = JSON.parse(decodeURIComponent(payloadAttr));
              const decoded = normalizeVisualPayload(rawDecoded) || rawDecoded;
              if (decoded && decoded.visualType) {
                sectionVisualType = decoded.visualType;
                sectionVisualData = decoded;
              }
            } catch (_) {}
          }
          curr = curr.nextElementSibling;
          continue;
        }

        const rawElText = getElementCleanText(curr);
        if (rawElText) {
          sectionRawLines.push(rawElText);
        }

        // Collect images inside this section
        const currImgs = Array.from(curr.querySelectorAll("img"))
          .map(img => img.src || img.getAttribute("src") || "")
          .filter(s => s && s.length > 50);

        if (curr.tagName === "IMG") {
          const imgSrc = curr.src || curr.getAttribute("src") || "";
          if (imgSrc.length > 50) currImgs.push(imgSrc);
        }
        if (currImgs.length > 0) sectionImgs.push(...currImgs);

        // Process H3, H4, H5, H6 immediately as Subtitle or lead-in
        if (["H3", "H4", "H5", "H6"].includes(curr.tagName)) {
          const subText = (curr.innerText || curr.textContent || "")
            .replace(/^#{1,6}\s*/, "")
            .replace(/^(?:Sub-?title|Subtitle\s*Text):\s*/i, "")
            .replace(/^[#*_`\s]+|[#*_`\s]+$/g, "")
            .trim();
          if (subText && !sectionSubtitle) {
            sectionSubtitle = subText;
            curr = curr.nextElementSibling;
            continue;
          } else if (subText) {
            allBodyLines.push(`• **${subText}:**`);
            additionalBodyLines.push(`• **${subText}:**`);
            curr = curr.nextElementSibling;
            continue;
          }
        }

        // Process tables inside a slide section
        if (curr.tagName === "TABLE") {
          const tbl = extractTableContent(curr);
          if (tbl && tbl.bullets.length > 0) {
            allBodyLines.push(...tbl.bullets);
            sectionTableData = {
              headers: tbl.headers,
              rows: tbl.dataRows
            };
          }
        } else if (curr.tagName === "UL" || curr.tagName === "OL") {
          Array.from(curr.querySelectorAll("li")).forEach(li => {
            const txt = getElementCleanText(li);
            if (txt) {
              const b = `• ${txt.replace(/^[-•*]\s+/, "")}`;
              allBodyLines.push(b);
              additionalBodyLines.push(b);
            }
          });
        } else {
          const txt = getElementCleanText(curr);
          if (txt) {
            allBodyLines.push(txt);
            if (!/^(?:Subtitle|Color|Visual|Title Size|Subtitle Size):/i.test(txt)) {
              additionalBodyLines.push(txt);
            }
          }
        }
        curr = curr.nextElementSibling;
      }

      const sectionCombinedText = sectionRawLines.join("\n");

      // Check for visual payload JSON (metric_grid_3col, three_column_cards, process_timeline, before_after, etc.)
      // Only run the text-based extractor when the DOM walk above did not already decode a payload.
      if (!sectionVisualType) {
        const secPayload = extractVisualPayload(sectionCombinedText);
        if (secPayload) {
          sectionVisualType = secPayload.visualType;
          sectionVisualData = secPayload;
        }
      }

      // Fallback: Check for Markdown-style Before / After Comparison
      if (!sectionVisualType && /BEFORE\s*:/i.test(sectionCombinedText) && /AFTER\s*:/i.test(sectionCombinedText)) {
        try {
          const beforeMatch = sectionCombinedText.match(/(?:🔴\s*)?BEFORE\s*:\s*([^\n]+)([\s\S]*?)(?:(?:🟢\s*)?AFTER\s*:\s*([^\n]+)([\s\S]*?))(?=(?:💡\s*Strategic Takeaway|Takeaway:|$))/i);
          if (beforeMatch) {
            const beforeTitle = beforeMatch[1].replace(/^[#*_`\s]+|[#*_`\s]+$/g, '').trim();
            const beforeRaw = beforeMatch[2].trim();
            const afterTitle = beforeMatch[3].replace(/^[#*_`\s]+|[#*_`\s]+$/g, '').trim();
            const afterRaw = beforeMatch[4].trim();

            const parseBullets = (raw) => {
              return raw
                .split(/\n+/)
                .map(l => l.replace(/^[-•*]\s*/, '').trim())
                .filter(l => l.length > 5 && !/^(?:Takeaway|💡|Verified)/i.test(l));
            };

            const beforeBullets = parseBullets(beforeRaw);
            const afterBullets = parseBullets(afterRaw);

            if (beforeBullets.length > 0 || afterBullets.length > 0) {
              sectionVisualType = "before_after";
              sectionVisualData = {
                title: title,
                subtitle: sectionSubtitle,
                before: {
                  title: beforeTitle || "Current State / Challenges",
                  bullets: beforeBullets
                },
                after: {
                  title: afterTitle || "Target State / Transformation",
                  bullets: afterBullets
                }
              };
            }
          }
        } catch (e) {
          console.warn("Error parsing markdown before/after visual in Strategy 1:", e);
        }
      }

      // Check for Chart specification in section
      const chartSpec = parseChartSpec(sectionCombinedText);
      if (chartSpec) {
        try {
          const chartDataUri = renderThemedChart(chartSpec);
          if (chartDataUri && !sectionImgs.includes(chartDataUri)) {
            sectionImgs.unshift(chartDataUri);
          }
        } catch (cErr) {
          console.warn("Chart rendering warning in Strategy 1:", cErr);
        }
      }

      // Check for markdown table in section text if not already found in DOM
      if (!sectionTableData) {
        const sMdTable = parseMarkdownTable(sectionCombinedText);
        if (sMdTable && sMdTable.dataRows.length > 0) {
          sectionTableData = {
            headers: sMdTable.headers,
            rows: sMdTable.dataRows
          };
        }
      }

      const parsedAll = extractSlideMetadataAndBullets(allBodyLines, sectionSubtitle);
      const parsedAdditional = extractSlideMetadataAndBullets(additionalBodyLines, sectionSubtitle);

      if (sectionVisualType && sectionVisualData) {
        const hydrated = hydrateOrDropVisualPayload(sectionVisualType, sectionVisualData, parsedAll.body);
        sectionVisualType = hydrated.visualType;
        sectionVisualData = hydrated.visualData;
      }

      if (sectionVisualType && sectionVisualData) {
        const formatted = formatVisualPayloadBody(sectionVisualData);
        if (formatted && formatted.trim()) {
          parsedAll.body = formatted;
        }
      }

      // Check if slide body contains 2-4 structured bullets with bold lead-ins for modern card layout
      // (Only compute if no explicit visualType payload was provided, preventing payload conflict)
      const structuredCards = (!sectionVisualType || !sectionVisualData) ? extractStructuredBullets(parsedAll.body) : null;

      // Deduplicate narrative bullets against native table if present
      let cleanAdditionalBody = "";
      if (sectionTableData && parsedAdditional.body !== "• Executive slide content") {
        const rawBullets = parsedAdditional.body.split(/\n\n+/).filter(b => b.trim() && b.trim() !== "• Executive slide content");
        const filteredBullets = filterDuplicateTableBullets(rawBullets, sectionTableData);
        cleanAdditionalBody = filteredBullets.join("\n\n");
      }

      // Slide visual images (chart or companion illustrations)
      let slideImgs = [];
      if (sectionImgs.length > 0) {
        slideImgs = sectionImgs;
      } else if (allImages.length > 0 && headerEls.length === 1) {
        slideImgs = allImages;
      }

      slides.push({
        slideNumber: i + 1,
        title: title,
        subtitle: sectionSubtitle || parsedAll.subtitle,
        takeaway: parsedAll.takeaway,
        visualConcept: parsedAll.visualConcept,
        visualType: sectionVisualType,
        visualData: sectionVisualData,
        structuredCards: structuredCards,
        color: parsedAll.color,
        titleSize: parsedAll.titleSize,
        subtitleSize: parsedAll.subtitleSize,
        body: parsedAll.body,
        additionalBody: cleanAdditionalBody,
        tableData: sectionTableData,
        base64Images: slideImgs
      });
    }

    if (slides.length >= 1) {
      return finalizeSlides(slides, allImages, rawText, options);
    }
  }

  // -------------------------------------------------------------
  // Strategy 2: Outline Table Unpacker (e.g. | Slide # | Title | Content |)
  // ONLY activates when table explicitly describes a presentation outline!
  // -------------------------------------------------------------
  const tableEls = Array.from(tempDiv.querySelectorAll("table"));
  for (const table of tableEls) {
    const rows = Array.from(table.querySelectorAll("tr"));
    if (rows.length < 2) continue;

    // Check if this table is explicitly a slide outline table
    const headerText = (rows[0].innerText || rows[0].textContent || "").toLowerCase();
    const isExplicitOutlineHeader = 
      headerText.includes("slide #") || 
      headerText.includes("slide number") || 
      headerText.includes("slide title") || 
      headerText.includes("deck outline") ||
      headerText.includes("presentation outline");

    let explicitSlideRowCount = 0;
    for (let r = 0; r < rows.length; r++) {
      const cells = Array.from(rows[r].querySelectorAll("td, th"))
        .map(c => (c.innerText || c.textContent || "").trim());
      if (cells.length >= 2 && (/^slide\s*\d+/i.test(cells[0]) || /^slide\s*\d+/i.test(cells[1]))) {
        explicitSlideRowCount++;
      }
    }

    // ONLY treat as slide outline if headers or multiple cells explicitly identify slides!
    if (!isExplicitOutlineHeader && explicitSlideRowCount < 2) {
      // This is a data/comparison table, NOT a slide outline! Skip Strategy 2.
      continue;
    }

    const tableSlides = [];
    for (let r = 0; r < rows.length; r++) {
      const row = rows[r];
      const cells = Array.from(row.querySelectorAll("td, th"))
        .map(c => (c.innerText || c.textContent || "").trim());

      if (cells.length < 2) continue;

      // Skip header row if it contains generic labels like "Slide #", "Slide Title"
      const firstCell = cells[0].toLowerCase();
      const secondCell = cells[1].toLowerCase();
      if (firstCell.includes("slide #") || firstCell.includes("slide number") || secondCell === "slide title") {
        continue;
      }

      // Check if this row represents a slide (e.g. "Slide 1", "Slide 1: Intro")
      const isSlideRow = /^slide\s*\d+/i.test(cells[0]) || 
                         /^slide\s*\d+/i.test(cells[1]) ||
                         cells.some(c => /^slide\s*\d+[:\-]/i.test(c));

      if (isSlideRow || (isExplicitOutlineHeader && r > 0)) {
        let slideTitle = "";
        let contentParts = [];

        if (cells.length >= 3 && /^slide\s*\d+/i.test(cells[0])) {
          // Format: [Slide 1] | [Title] | [Content / Visuals]
          slideTitle = cells[1];
          contentParts = cells.slice(2);
        } else if (cells.length >= 2 && /^slide\s*\d+[:\-]/i.test(cells[0])) {
          // Format: [Slide 1: Title] | [Content]
          slideTitle = cells[0].replace(/^slide\s*\d+[:\-]?\s*/i, "").trim();
          contentParts = cells.slice(1);
        } else {
          // Format: [Title] | [Content]
          slideTitle = cells[0];
          contentParts = cells.slice(1);
        }

        slideTitle = cleanSlideTitle(slideTitle, tableSlides.length + 1);

        // Format content parts into clean bullet points
        const bodyBullets = [];
        contentParts.forEach(part => {
          if (!part) return;
          const cleanPart = part
            .replace(/\\rightarrow|\$\\rightarrow\$/g, "→")
            .replace(/\*\*(.*?)\*\*/g, "$1");

          const lines = cleanPart.split(/\r?\n|•\s*/).map(l => l.trim()).filter(Boolean);
          lines.forEach(line => {
            if (line.includes(" — ")) {
              const subSplit = line.split(" — ");
              bodyBullets.push(`• **${subSplit[0].trim()}:** ${subSplit.slice(1).join(" — ").trim()}`);
            } else if (line.includes(": ") && !line.startsWith("http")) {
              const subSplit = line.split(": ");
              bodyBullets.push(`• **${subSplit[0].trim()}:** ${subSplit.slice(1).join(": ").trim()}`);
            } else {
              bodyBullets.push(`• ${line}`);
            }
          });
        });

        tableSlides.push({
          slideNumber: tableSlides.length + 1,
          title: slideTitle,
          body: bodyBullets.join("\n\n"),
          base64Images: allImages[tableSlides.length] ? [allImages[tableSlides.length]] : []
        });
      }
    }

    if (tableSlides.length >= 2) {
      return finalizeSlides(tableSlides, allImages, rawText, options);
    }
  }

  // -------------------------------------------------------------
  // Strategy 3: Bulleted / Numbered Outline List Unpacker
  // -------------------------------------------------------------
  const listItems = Array.from(tempDiv.querySelectorAll("li"));
  const outlineSlides = [];

  for (let i = 0; i < listItems.length; i++) {
    const text = (listItems[i].innerText || listItems[i].textContent || "").trim();
    // Match "Slide 1: Title — Description" or "Slide 1 - Title: Description"
    const slideMatch = text.match(/^(?:Slide\s*(\d+)|\*\*Slide\s*(\d+)\*\*)\s*[:\-–—]\s*(.*)$/i);

    if (slideMatch) {
      const rest = slideMatch[3].trim();
      let title = "";
      let bodyText = "";

      if (rest.includes(" — ")) {
        const parts = rest.split(" — ");
        title = parts[0].trim();
        bodyText = parts.slice(1).join(" — ").trim();
      } else if (rest.includes(" - ")) {
        const parts = rest.split(" - ");
        title = parts[0].trim();
        bodyText = parts.slice(1).join(" - ").trim();
      } else if (rest.includes(": ")) {
        const parts = rest.split(": ");
        title = parts[0].trim();
        bodyText = parts.slice(1).join(": ").trim();
      } else {
        title = rest;
      }

      title = cleanSlideTitle(title, outlineSlides.length + 1);

      // Clean body text into structured bullet points
      const bodyBullets = [];
      if (bodyText) {
        const subItems = bodyText
          .replace(/\\rightarrow|\$\\rightarrow\$/g, "→")
          .split(/\s*;\s*|\.\s+(?=[A-Z])/)
          .filter(Boolean);

        subItems.forEach(item => {
          const cleanItem = item.replace(/^[-•*]\s*/, "").trim();
          if (cleanItem) bodyBullets.push(`• ${cleanItem}`);
        });
      }

      outlineSlides.push({
        slideNumber: outlineSlides.length + 1,
        title: title,
        body: bodyBullets.length > 0 ? bodyBullets.join("\n\n") : (bodyText ? `• ${bodyText}` : ""),
        base64Images: allImages[outlineSlides.length] ? [allImages[outlineSlides.length]] : []
      });
    }
  }

  if (outlineSlides.length >= 2) {
    return finalizeSlides(outlineSlides, allImages, rawText, options);
  }

  // -------------------------------------------------------------
  // Strategy 4: Standalone Data Table (HTML <table> or Markdown |...|)
  // When a table is present without multiple explicit slide markers,
  // produce a dedicated slide with full tabular layout data.
  // -------------------------------------------------------------
  const allTables = Array.from(tempDiv.querySelectorAll("table"));
  if (allTables.length >= 2 && isSummarizeSlides) {
    const summarizeDeck = parseSummarizeDeck(tempDiv, allImages, rawText, options);
    if (summarizeDeck && summarizeDeck.length >= 2) {
      return finalizeSlides(summarizeDeck, allImages, rawText, options);
    }
  }

  const standaloneTable = tempDiv.querySelector("table");
  if (standaloneTable) {
    const tbl = extractTableContent(standaloneTable);
    if (tbl && tbl.dataRows.length > 0) {
      let tableTitle = "";
      
      // Check if there is a rendered chart container or chart image in tempDiv
      const chartEl = tempDiv.querySelector(".rendered-chart-container, [data-chart-title]");
      if (chartEl && chartEl.getAttribute("data-chart-title")) {
        tableTitle = chartEl.getAttribute("data-chart-title");
      } else {
        const chartImg = tempDiv.querySelector("img[alt*='chart' i], img[alt]");
        if (chartImg && chartImg.getAttribute("alt")) {
          const altText = chartImg.getAttribute("alt").replace(/\s*\([^)]*chart\)/i, "").trim();
          if (altText && altText !== "Image" && altText !== "Generated Chart") {
            tableTitle = altText;
          }
        }
      }

      if (!tableTitle) {
        let prevEl = standaloneTable.previousElementSibling;
        while (prevEl) {
          const txt = (prevEl.innerText || prevEl.textContent || "").trim();
          if (txt && !txt.startsWith("Verified Sources")) {
            tableTitle = txt;
            break;
          }
          prevEl = prevEl.previousElementSibling;
        }
      }

      if (!tableTitle) {
        const firstLine = (tempDiv.innerText || tempDiv.textContent || "").split("\n").map(l => l.trim()).find(l => l.length > 0);
        if (firstLine && !firstLine.startsWith("|")) tableTitle = firstLine;
      }

      if (!tableTitle && tbl.headers.length > 0) {
        tableTitle = `${tbl.headers[0]} Comparison`;
      }

      const cleanTitle = cleanSlideTitle(tableTitle || "Comparison Table", 1);

      // Extract any extra commentary/bullets before or after standalone table
      const extraLines = [];
      Array.from(tempDiv.querySelectorAll("h1, h2, h3, h4, p, li")).forEach(el => {
        const txt = (el.innerText || el.textContent || "").trim();
        if (txt && !txt.startsWith("Verified Sources") && txt !== tableTitle && txt !== cleanTitle && !/^(?:Subtitle|Color|Visual):/i.test(txt)) {
          extraLines.push(txt);
        }
      });
      const parsedExtra = extractSlideMetadataAndBullets(extraLines);
      const tableData = {
        headers: tbl.headers,
        rows: tbl.dataRows
      };
      let cleanAdditionalBody = "";
      if (parsedExtra.body !== "• Executive slide content") {
        const rawBullets = parsedExtra.body.split(/\n\n+/).filter(b => b.trim() && b.trim() !== "• Executive slide content");
        const filteredBullets = filterDuplicateTableBullets(rawBullets, tableData);
        cleanAdditionalBody = filteredBullets.join("\n\n");
      }

      return finalizeSlides([{
        slideNumber: 1,
        title: cleanTitle,
        subtitle: parsedExtra.subtitle || "",
        visualConcept: "",
        color: null,
        titleSize: 36,
        subtitleSize: 20,
        body: tbl.bulletText,
        additionalBody: cleanAdditionalBody,
        tableData: tableData,
        base64Images: allImages
      }], allImages, rawText, options);
    }
  }

  // Check for raw markdown table in text
  const textContent = tempDiv.innerText || tempDiv.textContent || rawText;
  const mdTable = parseMarkdownTable(textContent);
  if (mdTable && mdTable.dataRows.length > 0) {
    const textLines = textContent.split("\n").map(l => l.trim()).filter(Boolean);
    const firstLine = textLines.find(l => !l.startsWith("|"));
    const cleanTitle = cleanSlideTitle(firstLine || "Comparison Table", 1);
    const nonTableLines = textLines.slice(1).filter(l => !l.startsWith("|") && !l.endsWith("|") && l !== firstLine);
    const parsedExtra = extractSlideMetadataAndBullets(nonTableLines);
    const tableData = {
      headers: mdTable.headers,
      rows: mdTable.dataRows
    };
    let cleanAdditionalBody = "";
    if (parsedExtra.body !== "• Executive slide content") {
      const rawBullets = parsedExtra.body.split(/\n\n+/).filter(b => b.trim() && b.trim() !== "• Executive slide content");
      const filteredBullets = filterDuplicateTableBullets(rawBullets, tableData);
      cleanAdditionalBody = filteredBullets.join("\n\n");
    }
    return finalizeSlides([{
      slideNumber: 1,
      title: cleanTitle,
      subtitle: parsedExtra.subtitle || "",
      visualConcept: "",
      color: null,
      titleSize: 36,
      subtitleSize: 20,
      body: mdTable.bulletText,
      additionalBody: cleanAdditionalBody,
      tableData: tableData,
      base64Images: allImages
    }], allImages, rawText, options);
  }

  // -------------------------------------------------------------
  // Strategy 5: Raw Text Block Splitting
  // -------------------------------------------------------------
  let splitRegex = /(?:^|\n)(?=(?:#{1,6}\s*(?:Slide\s*\d+|[^\n]+)|(?:Slide|SLIDE)\s*\d+(?:\s*[:\-–—]|\s*\n)|(?:\d+[\.\)]\s*(?:\*\*)?Slide\s*\d+)))/gi;
  if (!/(?:Slide\s*\d+|#+\s*Slide\s*\d+)/i.test(textContent)) {
    if ((textContent.match(/(?:^|\n)##\s+/g) || []).length >= 2) {
      splitRegex = /(?:^|\n)(?=##\s+)/g;
    } else {
      splitRegex = /(?:^|\n)(?=(?:#{1,2}\s+|Slide\s*\d+[:\-]|(?:\d+\.\s+\*\*Slide)))/gi;
    }
  }
  const blocks = textContent
    .split(splitRegex)
    .map(b => b.trim())
    .filter(b => {
      if (b.length < 5) return false;
      // Skip preamble / presentation metadata cards if they precede Slide 1
      if (/^(?:presentation\s+deck\s+ready|\d+\s+slides)/i.test(b) && !b.includes("Slide 1")) return false;
      // Skip conversational intro preambles like "Here is a comparison..." if they precede real headings
      if (/^(?:here\s+is\s+|below\s+is\s+|sure|certainly|i've\s+prepared)/i.test(b) && !b.includes("\n#") && !b.includes("\n•") && !b.includes("Slide 1")) {
        return false;
      }
      return true;
    });

  if (blocks.length >= 2) {
    const textSlides = [];
    blocks.forEach((block, idx) => {
      const lines = block.split("\n").map(l => l.trim()).filter(Boolean);
      if (lines.length > 0) {
        let titleLine = lines[0];
        let contentStartIndex = 1;
        // If line 0 is just "Slide X" (or "Slide X:" with nothing else), take line 1 as the title!
        if (/^Slide\s*\d+[:\-–—]?$/i.test(lines[0]) && lines.length > 1) {
          titleLine = lines[1];
          contentStartIndex = 2;
        }

        const rawTitle = titleLine
          .replace(/^[#*\s:]+/, "")
          .replace(/^\d+\.\s*/, "")
          .replace(/^Slide\s*\d+[:\-]?\s*/i, "")
          .trim();

        const title = cleanSlideTitle(rawTitle, idx + 1);
        const blockMdTable = parseMarkdownTable(block);
        const nonTableLines = lines.slice(contentStartIndex).filter(l => !l.startsWith("|") && !l.endsWith("|"));
        const parsed = extractSlideMetadataAndBullets(lines.slice(contentStartIndex));
        const parsedNonTable = extractSlideMetadataAndBullets(nonTableLines);

        let blockVisualType = null;
        let blockVisualData = null;
        const blockImgs = allImages[idx] ? [allImages[idx]] : [];

        // Check for visual payload JSON (metric_grid_3col, three_column_cards, process_timeline, before_after, etc.)
        const blockPayload = extractVisualPayload(block);
        if (blockPayload) {
          blockVisualType = blockPayload.visualType;
          blockVisualData = blockPayload;
        }

        // Fallback: Check for Markdown-style Before / After Comparison
        if (!blockVisualType && /BEFORE\s*:/i.test(block) && /AFTER\s*:/i.test(block)) {
          try {
            const beforeMatch = block.match(/(?:🔴\s*)?BEFORE\s*:\s*([^\n]+)([\s\S]*?)(?:(?:🟢\s*)?AFTER\s*:\s*([^\n]+)([\s\S]*?))(?=(?:💡\s*Strategic Takeaway|Takeaway:|$))/i);
            if (beforeMatch) {
              const beforeTitle = beforeMatch[1].replace(/^[#*_`\s]+|[#*_`\s]+$/g, '').trim();
              const beforeRaw = beforeMatch[2].trim();
              const afterTitle = beforeMatch[3].replace(/^[#*_`\s]+|[#*_`\s]+$/g, '').trim();
              const afterRaw = beforeMatch[4].trim();

              const parseBullets = (raw) => {
                return raw
                  .split(/\n+/)
                  .map(l => l.replace(/^[-•*]\s*/, '').trim())
                  .filter(l => l.length > 5 && !/^(?:Takeaway|💡|Verified)/i.test(l));
              };

              const beforeBullets = parseBullets(beforeRaw);
              const afterBullets = parseBullets(afterRaw);

              if (beforeBullets.length > 0 || afterBullets.length > 0) {
                blockVisualType = "before_after";
                blockVisualData = {
                  title: title,
                  subtitle: parsed.subtitle,
                  before: {
                    title: beforeTitle || "Current State / Challenges",
                    bullets: beforeBullets
                  },
                  after: {
                    title: afterTitle || "Target State / Transformation",
                    bullets: afterBullets
                  }
                };
              }
            }
          } catch (e) {
            console.warn("Error parsing markdown before/after visual in Strategy 5:", e);
          }
        }

        // Check for chart spec in block
        const chartSpec = parseChartSpec(block);
        if (chartSpec) {
          try {
            const chartDataUri = renderThemedChart(chartSpec);
            if (chartDataUri && !blockImgs.includes(chartDataUri)) {
              blockImgs.unshift(chartDataUri);
            }
          } catch (cErr) {
            console.warn("Chart rendering warning in Strategy 5:", cErr);
          }
        }

        if (blockVisualType && blockVisualData) {
          const hydrated = hydrateOrDropVisualPayload(blockVisualType, blockVisualData, parsed.body);
          blockVisualType = hydrated.visualType;
          blockVisualData = hydrated.visualData;
        }

        if (blockVisualType && blockVisualData) {
          const formatted = formatVisualPayloadBody(blockVisualData);
          if (formatted && formatted.trim()) {
            parsed.body = formatted;
          }
        }

        // Check if slide body contains 2-4 structured bullets with bold lead-ins for modern card layout
        // (Only compute if no explicit visualType payload was provided, preventing payload conflict)
        const structuredCards = (!blockVisualType || !blockVisualData) ? extractStructuredBullets(parsed.body) : null;

        let cleanAdditionalBody = "";
        const tableData = blockMdTable ? { headers: blockMdTable.headers, rows: blockMdTable.dataRows } : null;
        if (tableData && parsedNonTable.body !== "• Executive slide content") {
          const rawBullets = parsedNonTable.body.split(/\n\n+/).filter(b => b.trim() && b.trim() !== "• Executive slide content");
          const filtered = filterDuplicateTableBullets(rawBullets, tableData);
          cleanAdditionalBody = filtered.join("\n\n");
        }

        textSlides.push({
          slideNumber: idx + 1,
          title: title,
          subtitle: parsed.subtitle,
          takeaway: parsed.takeaway,
          visualConcept: parsed.visualConcept,
          visualType: blockVisualType,
          visualData: blockVisualData,
          structuredCards: structuredCards,
          color: parsed.color,
          titleSize: parsed.titleSize,
          subtitleSize: parsed.subtitleSize,
          body: blockMdTable ? blockMdTable.bulletText : parsed.body,
          additionalBody: cleanAdditionalBody,
          tableData: tableData,
          base64Images: blockImgs
        });
      }
    });

    if (textSlides.length >= 2) return finalizeSlides(textSlides, allImages, rawText, options);
  }

  // -------------------------------------------------------------
  // Strategy 6: Single Slide Fallback
  // -------------------------------------------------------------
  const cloneDiv = tempDiv.cloneNode(true);
  cloneDiv.querySelectorAll("h1, h2, h3, p, li, tr, div, blockquote").forEach(el => {
    el.insertAdjacentText("afterend", "\n");
  });
  const fallbackRaw = (cloneDiv.textContent || cloneDiv.innerText || rawText || "").trim();
  const sanitizedFallback = sanitizeAiResponse(fallbackRaw);
  const lines = sanitizedFallback.split("\n").map(l => l.trim()).filter(Boolean);

  let singleTitle = "🎯 Key Highlights";
  let contentLines = lines;

  if (lines.length > 0) {
    const candidateTitle = cleanSlideTitle(lines[0], 1);
    // A paragraph is never a title. Promoting one used to leave a cut-off
    // first sentence in the title box and nothing at all in the body.
    if (candidateTitle && !isParagraphLikeLine(lines[0]) && !isConversationalPreamble(lines[0]) && !/^[-•*]/.test(lines[0])) {
      singleTitle = candidateTitle;
      contentLines = lines.slice(1);
    } else {
      contentLines = lines;
    }
  }

  // A lone long paragraph reads as a wall of text on a slide: one sentence per line.
  if (contentLines.length === 1 && contentLines[0].length > 240) {
    const sentences = splitParagraphIntoSentences(contentLines[0]);
    if (sentences.length >= 2) contentLines = sentences;
  }

  // Filter out any conversational preamble lines from slide body content
  contentLines = contentLines.filter(l => !isConversationalPreamble(l));

  // If this slide has an image, derive a clean subject title if title is generic or preamble
  if (allImages.length > 0) {
    const combinedPromptContext = `${(typeof window !== "undefined" && window.__lastUserPrompt) || ""} ${fallbackRaw}`;
    const subjectMatch = combinedPromptContext.match(/(?:image|picture|photo|photograph|portrait|illustration|drawing)\s+(?:of|showing|featuring|with)?\s+(?:a|an|the)?\s*([a-zA-Z0-9\s]+?)(?:\s+you\s+requested|[.:!\n]|$)/i);
    if (subjectMatch && subjectMatch[1]) {
      const subj = subjectMatch[1].trim();
      if (subj.length > 1 && subj.length < 35 && !isConversationalPreamble(subj)) {
        singleTitle = `🎨 ${subj.charAt(0).toUpperCase() + subj.slice(1)}`;
      }
    } else if (singleTitle === "🎯 Key Highlights" || isConversationalPreamble(singleTitle) || singleTitle.startsWith("Slide 1")) {
      singleTitle = "🎨 Visual Concept";
    }
  }

  const parsed = extractSlideMetadataAndBullets(contentLines);
  const slideBody = contentLines.length > 0 ? parsed.body : "";

  return finalizeSlides([{
    slideNumber: 1,
    title: singleTitle,
    subtitle: parsed.subtitle,
    takeaway: parsed.takeaway,
    visualConcept: parsed.visualConcept,
    color: parsed.color,
    titleSize: parsed.titleSize,
    subtitleSize: parsed.subtitleSize,
    body: slideBody,
    base64Images: allImages
  }], allImages, rawText, options);
}

/**
 * Finalizes parsed slides:
 * 1. Consolidates executive summary slides into a single slide if applicable.
 * 2. Enforces explicit slide count capping if the user explicitly requested N slides (e.g. "create 5 slides").
 * 3. Normalizes slide numbering.
 */
function finalizeSlides(slides, allImages = [], rawText = "", options = {}) {
  if (!slides || slides.length === 0) return [];

  const isSummarizeSlides = Boolean(
    options?.isSummarizeSlides ||
    (typeof window !== "undefined" && window.__isSummarizeSlidesAction)
  );

  let finalSlides = isSummarizeSlides ? slides : consolidateExecutiveSummarySlides(slides, allImages, options);

  // Executive Presentation Rule:
  // When a table is present, substantive bullet points beneath the table become slide notes (slide.notes)
  // instead of creating an unwanted separate slide.
  const processedSlides = [];
  for (const s of finalSlides) {
    const hasTableData = Boolean(s.tableData && s.tableData.rows && s.tableData.rows.length > 0);
    const hasVisualImages = Boolean(s.base64Images && s.base64Images.length > 0);

    // If slide contains BOTH a table AND an image/chart, keep them together on a unified side-by-side slide!
    if (hasTableData && hasVisualImages) {
      const rawBullets = (s.additionalBody || s.body || "")
        .split(/\r?\n\r?\n+/)
        .map(b => b.trim())
        .filter(b => b && b !== "• Executive slide content" && !b.startsWith("|") && !b.startsWith("---"));
      const filteredBullets = filterDuplicateTableBullets(rawBullets, s.tableData);

      const combinedSlide = {
        ...s,
        tableData: s.tableData,
        base64Images: s.base64Images,
        compressedImages: s.compressedImages || [],
        notes: filteredBullets.length > 0 ? filteredBullets.join("\n\n") : (s.notes || ""),
        slideNumber: processedSlides.length + 1
      };
      processedSlides.push(combinedSlide);
      continue;
    }

    if (hasTableData) {
      const rawBullets = (s.additionalBody || s.body || "")
        .split(/\r?\n\r?\n+/)
        .map(b => b.trim())
        .filter(b => b && b !== "• Executive slide content" && !b.startsWith("|") && !b.startsWith("---"));
      const filteredBullets = filterDuplicateTableBullets(rawBullets, s.tableData);

      // Slide: Main Content & Table ONLY (clean canvas without overlapping text), takeaways saved into slide.notes
      const tableSlide = {
        ...s,
        base64Images: [],
        compressedImages: [],
        body: "",
        additionalBody: "",
        tableData: s.tableData,
        notes: filteredBullets.length > 0 ? filteredBullets.join("\n\n") : (s.notes || "")
      };
      processedSlides.push(tableSlide);

      // In Summarize Slides mode: if there are takeaways attached to the table slide, and no separate
      // Key Takeaways slide exists yet, break down the key takeaways into dedicated slide(s) (max 5 bullets per slide)!
      if (isSummarizeSlides && filteredBullets.length > 0) {
        const hasExistingTakeawaysSlide = finalSlides.some(other =>
          other !== s && (
            (other.title || "").toLowerCase().includes("takeaway") ||
            (other.title || "").toLowerCase().includes("key takeaways")
          )
        );
        if (!hasExistingTakeawaysSlide) {
          const maxBulletsPerSlide = 5;
          for (let c = 0; c < filteredBullets.length; c += maxBulletsPerSlide) {
            const chunk = filteredBullets.slice(c, c + maxBulletsPerSlide);
            const isContinuation = c > 0;
            const takeawaySlide = {
              slideNumber: processedSlides.length + 1,
              title: isContinuation ? "Executive Summary: Key Takeaways (Cont.)" : "Executive Summary: Key Takeaways",
              subtitle: isContinuation ? "Continued Highlights" : "Strategic Highlights & Next Steps",
              visualConcept: "",
              visualType: null,
              visualData: null,
              color: null,
              titleSize: 36,
              subtitleSize: 20,
              body: chunk.join("\n\n"),
              additionalBody: chunk.join("\n\n"),
              tableData: null,
              base64Images: []
            };
            processedSlides.push(takeawaySlide);
          }
        }
      }
      continue;
    }

    processedSlides.push(s);
  }
  finalSlides = processedSlides;

  // In Summarize Slides mode: enforce maximum 5 slides, cap bullets, and re-index slide numbers
  if (isSummarizeSlides) {
    if (finalSlides.length > 5) {
      finalSlides = finalSlides.slice(0, 5);
    }
    for (const s of finalSlides) {
      if (s.body && !s.tableData) {
        const bulletLines = s.body.split(/\n\s*\n/).filter(b => b.trim().length > 0);
        if (bulletLines.length > 4) {
          s.body = bulletLines.slice(0, 4).join("\n\n");
          if (s.additionalBody) s.additionalBody = s.body;
        }
      }
    }
    finalSlides.forEach((s, idx) => {
      s.slideNumber = idx + 1;
    });
    // Caps how many slides may be card slides; see cardBudget.js.
    return applyCardLayoutBudget(finalSlides);
  }

  const wordToNumber = {
    "one": 1, "two": 2, "three": 3, "four": 4, "five": 5,
    "six": 6, "seven": 7, "eight": 8, "nine": 9, "ten": 10
  };
  let requestedCount = null;

  if (typeof window !== "undefined" && window.__lastUserPrompt) {
    const match = window.__lastUserPrompt.match(/\b(?:create|generate|make|build|provide|give\s+me)?\s*(\d+|one|two|three|four|five|six|seven|eight|nine|ten)\s+slides?\b/i);
    if (match && match[1]) {
      const token = match[1].toLowerCase();
      requestedCount = wordToNumber[token] || parseInt(token, 10);
    }
  }

  if (!requestedCount && rawText) {
    const match = rawText.match(/\b(?:create|generate|make|build|provide|give\s+me)?\s*(\d+|one|two|three|four|five|six|seven|eight|nine|ten)\s+slides?\b/i);
    if (match && match[1]) {
      const token = match[1].toLowerCase();
      requestedCount = wordToNumber[token] || parseInt(token, 10);
    }
  }

  if (requestedCount && requestedCount > 0 && finalSlides.length > requestedCount) {
    finalSlides = finalSlides.slice(0, requestedCount);
  }

  // On multi-slide decks, sanitize Slide 1 (Title slide) to strictly contain Title, Subtitle, and Executive Summary
  if (!isSummarizeSlides && finalSlides.length >= 2 && finalSlides[0].body) {
    let body = finalSlides[0].body;
    // Strip "Deck Scope & Outline" and all subsequent lines
    body = body.replace(/(?:•\s*)?Deck Scope\s*&\s*Outline[\s\S]*$/i, "").trim();
    // Strip any remaining upcoming slide references like "• Slide 2: ...", "• Slide 3: ..."
    body = body.replace(/(?:^|\n)\s*•?\s*Slide\s*\d+\s*:?[^\n]*/gi, "").trim();
    // If the body contains "• Executive Summary: ...", keep that clean executive summary paragraph
    const execSummaryMatch = body.match(/(?:•\s*)?(?:Executive Summary|Summary):\s*([\s\S]*?)(?=(?:\n\s*•|\n\s*#{1,6}|$))/i);
    if (execSummaryMatch && execSummaryMatch[1]) {
      const summaryText = execSummaryMatch[1].trim();
      if (summaryText.length > 20) {
        body = `• Executive Summary: ${summaryText}`;
      }
    }
    finalSlides[0].body = body;
  }

  const seenImageSignatures = new Set();
  finalSlides.forEach((s, idx) => {
    s.slideNumber = idx + 1;

    // Deduplicate images across all slides in the deck
    if (s.base64Images && s.base64Images.length > 0) {
      const uniqueImgs = [];
      for (const img of s.base64Images) {
        if (!img || typeof img !== "string" || img.length < 50) continue;
        const sig = `${img.length}_${img.substring(0, 60)}_${img.substring(Math.max(0, img.length - 60))}`;
        if (!seenImageSignatures.has(sig)) {
          seenImageSignatures.add(sig);
          uniqueImgs.push(img);
        }
      }
      s.base64Images = uniqueImgs;
    }
  });

  // Explicit single-visual actions (the Create Chart button, for example) must
  // never fan out into a deck just because the model emitted two headers while
  // a previous multi-slide request was still in its conversation context.
  // Deliberately NOT window.__forceSingleSlide. A session-wide global outlived
  // the chart request that set it and collapsed the next Doc to Deck run from
  // six slides to one. The caller that owns the response passes the cap.
  const slideCap = Number(options && options.maxSlides) || 0;

  if (slideCap === 1 && finalSlides.length > 1) {
    const primary = finalSlides[0];
    const extras = [];

    for (let i = 1; i < finalSlides.length; i++) {
      const extra = finalSlides[i];
      if (!primary.tableData && extra.tableData) primary.tableData = extra.tableData;
      if ((!primary.base64Images || primary.base64Images.length === 0) &&
          extra.base64Images && extra.base64Images.length > 0) {
        primary.base64Images = extra.base64Images;
      }
      if (!primary.visualType && extra.visualType) {
        primary.visualType = extra.visualType;
        primary.visualData = extra.visualData;
      }
      if (!primary.takeaway && extra.takeaway) primary.takeaway = extra.takeaway;

      for (const chunk of [extra.body, extra.additionalBody]) {
        const trimmed = (chunk || "").trim();
        if (trimmed && trimmed !== "\u2022 Executive slide content") extras.push(trimmed);
      }
    }

    if (extras.length > 0) {
      const merged = [primary.additionalBody, ...extras].filter(t => t && t.trim()).join("\n\n");
      primary.additionalBody = merged;
      const primaryBody = (primary.body || "").trim();
      if (!primaryBody || primaryBody === "\u2022 Executive slide content") {
        primary.body = extras.join("\n\n");
      }
    }

    primary.slideNumber = 1;
    return [primary];
  }

  // Caps how many slides may be card slides; see cardBudget.js.
  // Applied here so the chat preview and the build agree.
  return applyCardLayoutBudget(finalSlides);
}

/**
 * Consolidates multi-section executive summary slides into a single executive slide,
 * preserving all subtitles, bullets, tables, takeaways, and generated images.
 */
function consolidateExecutiveSummarySlides(slides, allImages = [], options = {}) {
  if (!slides || slides.length === 0) return slides;

  const isSummarizeSlides = Boolean(
    options?.isSummarizeSlides ||
    (typeof window !== "undefined" && window.__isSummarizeSlidesAction)
  );

  if (isSummarizeSlides) return slides;

  // CRITICAL GUARD: Never collapse a multi-slide deck where slides have distinct topics/titles!
  if (slides.length >= 2) {
    const distinctTitles = new Set(slides.map(s => (s.title || "").toLowerCase().trim()).filter(Boolean));
    if (distinctTitles.size >= 2) {
      return slides;
    }
  }

  const firstTitle = (slides[0].title || "").toLowerCase();
  const isExecutiveSummary = slides.length > 1 && (
    firstTitle.includes("executive slide summary") ||
    firstTitle.includes("executive summary") ||
    firstTitle.includes("slide summary") ||
    (firstTitle.includes("executive") && firstTitle.includes("summary")) ||
    (firstTitle.includes("summary") && !firstTitle.includes("slide 1") && !firstTitle.includes("slide #"))
  );

  if (!isExecutiveSummary) return slides;

  const merged = { ...slides[0] };
  const allBullets = [];

  if (merged.body && merged.body !== "• Executive slide content") {
    allBullets.push(merged.body);
  }
  if (merged.additionalBody && merged.additionalBody !== "• Executive slide content") {
    allBullets.push(merged.additionalBody);
  }

  for (let s = 1; s < slides.length; s++) {
    const sub = slides[s];
    if (!merged.subtitle && sub.title && !sub.title.toLowerCase().startsWith("slide")) {
      merged.subtitle = sub.title;
    }
    if (sub.body && sub.body !== "• Executive slide content") {
      allBullets.push(sub.body);
    }
    if (sub.additionalBody && sub.additionalBody !== "• Executive slide content") {
      allBullets.push(sub.additionalBody);
    }
    if (!merged.tableData && sub.tableData) {
      merged.tableData = sub.tableData;
    }
    if (sub.base64Images && sub.base64Images.length > 0) {
      if (!merged.base64Images) merged.base64Images = [];
      for (const img of sub.base64Images) {
        if (!merged.base64Images.includes(img)) {
          merged.base64Images.push(img);
        }
      }
    }
    if (!merged.takeaway && sub.takeaway) {
      merged.takeaway = sub.takeaway;
    }
  }

  if (allImages && allImages.length > 0) {
    if (!merged.base64Images) merged.base64Images = [];
    for (const img of allImages) {
      if (!merged.base64Images.includes(img)) {
        merged.base64Images.push(img);
      }
    }
  }

  // Deduplicate and structure bullets
  const uniqueBullets = [];
  const seen = new Set();
  for (const bBlock of allBullets) {
    const lines = bBlock.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
    for (const line of lines) {
      const cleanLine = line.replace(/^[-•*]\s*/, "").trim();
      if (cleanLine.startsWith("|") || cleanLine.startsWith("---") || /^(?:Subtitle|Color|Visual):/i.test(cleanLine)) {
        continue;
      }
      if (cleanLine.toLowerCase().startsWith("the chart and table below") || cleanLine.toLowerCase().startsWith("here is the")) {
        continue;
      }
      const normKey = cleanLine.toLowerCase().replace(/[*_`]/g, "");
      if (normKey.length > 5 && !seen.has(normKey)) {
        seen.add(normKey);
        uniqueBullets.push(line.startsWith("•") ? line : `• ${cleanLine}`);
      }
    }
  }

  // If a native PowerPoint table is present, exclude bullets that are merely converted table rows
  if (merged.tableData && merged.tableData.rows && merged.tableData.rows.length > 0) {
    const filtered = filterDuplicateTableBullets(uniqueBullets, merged.tableData);
    if (filtered.length > 0) {
      uniqueBullets.length = 0;
      uniqueBullets.push(...filtered);
    }

    if (uniqueBullets.length > 0) {
      // Single Executive Slide with Table ONLY; Key Takeaways attached as slide notes
      const slide1 = {
        ...merged,
        body: "",
        additionalBody: "",
        tableData: merged.tableData,
        notes: uniqueBullets.join("\n\n").trim()
      };
      return [slide1];
    }
  }

  merged.body = uniqueBullets.join("\n\n").trim() || "• Executive slide content";
  merged.additionalBody = merged.body;
  return [merged];
}

export function sanitizeAiResponse(text) {
  if (!text) return "";
  let clean = text.trim();

  // 1. Remove conversational preamble line(s) at start
  clean = clean.replace(/^(?:here\s+(?:is|are)|below\s+(?:is|are)|sure[.,!]?|certainly[.,!]?|of\s+course[.,!]?|as\s+requested[.,:]?|optimized\s+version[.,:]?|revised\s+version[.,:]?|i've\s+created|in\s+summary[.,:]?)[^\n]*\n+/i, "");

  // 2. Truncate alternative options (keep only Option 1)
  const altMatch = clean.match(/\n+\s*(?:(?:or,?\s+(?:for\s+)?(?:an\s+)?(?:even\s+)?(?:more\s+)?(?:minimalist|concise|alternative|compact)[^\n]*)|(?:option\s*2\b[^\n]*)|(?:alternative\s*(?:option|\d)?\b[^\n]*))\s*[:\n]/i);
  if (altMatch && altMatch.index !== undefined) {
    clean = clean.substring(0, altMatch.index).trim();
  }

  // 3. Remove conversational closing remarks at the end
  clean = clean.replace(/\n+(?:hope\s+this\s+helps|let\s+me\s+know|feel\s+free\s+to|let\s+me\s+know\s+if)[^\n]*$/i, "");

  return clean.trim();
}

export function isConversationalPreamble(line) {
  if (!line) return false;
  const l = line.toLowerCase().trim();
  if (/^(?:i\s+(?:can|have|will|am\s+happy\s+to)|let\s+me|happy\s+to|glad\s+to|feel\s+free|here\s+(?:is|are)|below\s+(?:is|are)|sure|certainly|of\s+course|as\s+requested|optimized|revised|concise|punchy|in\s+summary|to\s+make)/i.test(l)) return true;
  if (l.includes("help you with that") || l.includes("image you requested") || l.includes("image of a") || l.includes("here is the image") || l.includes("created an image") || l.includes("generated an image")) return true;
  if (l.includes("optimized for a presentation") || l.includes("version optimized") || l.includes("concise and punchy version") || l.includes("minimalist layout") || l.includes("presentation slide")) return true;
  return false;
}

export function extractCleanBulletPoints(htmlContent, rawText = "") {
  let text = (rawText || htmlContent || "").trim();
  if (text.includes("<") && text.includes(">")) {
    if (typeof document !== "undefined" && document.createElement) {
      const tmp = document.createElement("div");
      tmp.innerHTML = text;
      tmp.querySelectorAll("blockquote, .note, .ppt-deck-preview-container, .response-actions-container").forEach(el => el.remove());
      text = (tmp.textContent || tmp.innerText || "").trim();
    } else {
      text = text.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
    }
  }

  const sanitized = sanitizeAiResponse(text);
  const lines = sanitized.split("\n").map(l => l.trim()).filter(Boolean);
  const bullets = [];

  for (const line of lines) {
    if (isConversationalPreamble(line)) continue;
    const cleanLine = line.replace(/^[-•*]\s*/, "").replace(/\*\*/g, "").trim();
    if (cleanLine.length > 0) {
      bullets.push(`• ${cleanLine}`);
    }
  }

  return bullets.length > 0 ? bullets.join("\n") : sanitized;
}

export function cleanSlideTitle(rawTitle, defaultNum = 1) {
  if (!rawTitle) return `Slide ${defaultNum}`;

  let clean = rawTitle
    .replace(/^[#*\s:]+/, "")
    .replace(/^\d+[\.\)]\s*/, "")
    .replace(/^Slide\s*\d+[:\-–—]?\s*/i, "")
    .replace(/^(?:i\s+can\s+(?:certainly|definitely|gladly)?\s*help|happy\s+to\s+help|sure[.,]?\s*|certainly[.,]?\s*|absolutely[.,]?\s*|of\s+course[.,]?\s*)/i, "")
    .replace(/^(?:here\s+is\s+(?:a\s+)?(?:table|comparison|list|breakdown|summary|the\s+image|an\s+image)?\s+(?:of|comparing|for)?|below\s+is\s+(?:a\s+)?(?:table|comparison|list|breakdown|summary|the\s+image|an\s+image)?\s+(?:of|comparing|for)?)\s*/i, "")
    .replace(/^(?:a\s+table\s+of|a\s+comparison\s+of|table\s+of|comparison\s+of)\s*/i, "")
    .replace(/[:.]+$/, "")
    .replace(/\*\*/g, "")
    .trim();

  // If preamble was detected on the cleaned string, but it contains a colon, extract candidate after colon
  if (isConversationalPreamble(clean)) {
    if (clean.includes(":")) {
      const parts = clean.split(":");
      const candidate = parts.slice(1).join(":").trim();
      if (candidate && !isConversationalPreamble(candidate)) {
        clean = candidate;
      } else {
        return `Slide ${defaultNum}`;
      }
    } else {
      return `Slide ${defaultNum}`;
    }
  }

  // Preserve meaningful titles up to 8 words or 60 chars
  const words = clean.split(/\s+/);
  if (words.length > 8) {
    clean = words.slice(0, 8).join(" ");
  }
  if (clean.length > 60) {
    clean = clean.substring(0, 60).replace(/\s+\S*$/, "").trim();
  }

  return clean || `Slide ${defaultNum}`;
}

/**
 * True when a line reads as prose rather than a heading: long, or more than
 * one sentence. Used so a heading-less paragraph is never promoted to a title.
 */
function isParagraphLikeLine(line) {
  const text = String(line || "").trim();
  if (text.length > 100) return true;
  return text.length > 60 && /[a-z0-9%)][.!?]["')\]]*\s+[A-Z]/.test(text);
}

/**
 * Splits prose into sentences without breaking numbers ("$1.400 billion") or
 * initialisms ("U.S. History"): a boundary needs a lowercase letter, digit,
 * percent sign or closing bracket before the punctuation and whitespace after.
 */
function splitParagraphIntoSentences(text) {
  const source = String(text || "");
  const out = [];
  const boundary = /[a-z0-9%)\]][.!?]["')\]]*\s+(?=[A-Z0-9"'(])/g;
  let start = 0;
  let match;
  while ((match = boundary.exec(source)) !== null) {
    const end = match.index + match[0].length;
    const sentence = source.slice(start, end).trim();
    if (sentence) out.push(sentence);
    start = end;
  }
  const tail = source.slice(start).trim();
  if (tail) out.push(tail);
  return out;
}
