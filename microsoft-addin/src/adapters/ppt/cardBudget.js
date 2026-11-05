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
 * Card layout budget for generated decks.
 *
 * Card-style slides (three_column_cards, metric_grid_3col, process_timeline,
 * before_after, and 2-4 bold lead-in bullets that the parser turns into cards)
 * are by far the most expensive slides to build. Three 10-slide logs from
 * 09-28 had them on 55% of the slides, taking 79-94% of the slide-building
 * time, and a 4-step timeline was the slide that stalled PowerPoint Online in
 * two of those builds.
 *
 * So a deck gets a budget: at most 40% of its slides (rounded up) may be card
 * slides, and a deck under 3 slides may still have one. (30% in r48; raised
 * once slim cards made a card slide ~3x cheaper to build.) The cards beyond the
 * budget are drawn as plain bold lead-in bullets. The first card slides in the
 * deck keep their cards, so the result is predictable and the chat preview
 * (which runs the same parser) matches what gets built.
 *
 * PowerPoint-only: imported by slideParser, slideBuilder and promptEnhancer.
 */

export const CARD_LAYOUT_SHARE_PERCENT = 40;

export const CARD_VISUAL_TYPES = ["metric_grid_3col", "three_column_cards", "process_timeline", "before_after"];

/**
 * How many slides of a `totalSlides` deck may use a card layout.
 * Integer maths on purpose: 0.3 * 10 is 3.0000000000000004 in floating point,
 * and Math.ceil would turn that into 4. The same trap exists for other shares.
 */
export function cardLayoutBudget(totalSlides) {
  const n = Math.max(0, Math.floor(Number(totalSlides) || 0));
  if (n === 0) return 0;
  if (n < 3) return 1;
  return Math.ceil((n * CARD_LAYOUT_SHARE_PERCENT) / 100);
}

function hasRows(tableData) {
  return Boolean(tableData && Array.isArray(tableData.rows) && tableData.rows.length > 0);
}

/**
 * True when the builder would draw this slide as a row of cards. Table and
 * image slides take their own branches in the builder before any card layout
 * is considered, so they never count against the budget.
 */
export function isCardLayoutSlide(slide) {
  if (!slide) return false;
  if (hasRows(slide.tableData)) return false;
  if (Array.isArray(slide.base64Images) && slide.base64Images.length > 0) return false;
  if (slide.visualType && slide.visualData && CARD_VISUAL_TYPES.indexOf(slide.visualType) !== -1) return true;
  if (!slide.visualType && Array.isArray(slide.structuredCards) && slide.structuredCards.length >= 2) return true;
  return false;
}

function clean(value) {
  return String(value === null || value === undefined ? "" : value).replace(/\s+/g, " ").trim();
}

function leadIn(head, text, footnote) {
  const h = clean(head).replace(/:$/, "");
  let t = clean(text);
  const f = clean(footnote);
  if (f) t = t ? `${t} (${f})` : `(${f})`;
  if (h && t) return `\u2022 **${h}**: ${t}`;
  return `\u2022 ${h || t}`;
}

function joinBullets(list) {
  return Array.isArray(list) ? list.map(clean).filter(Boolean).join("; ") : "";
}

/** The card payload as bold lead-in bullets, or "" to keep the existing body. */
function cardsAsBullets(slide) {
  const vd = slide.visualData || {};
  let lines = [];
  if (slide.visualType === "metric_grid_3col") {
    lines = (vd.cards || []).map(c => leadIn(
      [clean(c.metric), clean(c.title)].filter(Boolean).join(" "),
      joinBullets(c.bullets) || c.body || c.subtitle,
      c.footnote
    ));
  } else if (slide.visualType === "three_column_cards") {
    lines = (vd.cards || []).map(c => leadIn(c.title || c.pill, c.body || joinBullets(c.bullets), c.footnote));
  } else if (slide.visualType === "process_timeline") {
    lines = (vd.steps || []).map(s => {
      const step = clean(s.step);
      const label = step && !/^\d+$/.test(step) ? `${step} \u2014 ${clean(s.title)}` : s.title;
      return leadIn(label, s.description, s.footnote);
    });
  }
  return lines.filter(l => l.replace(/^\u2022\s*/, "").trim()).join("\n\n");
}

function demoteCardSlide(slide, limit, total) {
  const layout = slide.visualType || "structuredCards";
  if (slide.visualType) {
    const text = cardsAsBullets(slide);
    if (text) slide.body = text;
    slide.visualType = null;
    slide.visualData = null;
  }
  slide.structuredCards = null;
  slide.cardBudgetDemoted = { layout: layout, limit: limit, total: total };
}

/**
 * Keeps the first `cardLayoutBudget(slides.length)` card slides and turns the
 * rest into bullet slides, in place. Returns the same array.
 */
export function applyCardLayoutBudget(slides) {
  if (!Array.isArray(slides) || slides.length === 0) return slides;
  const total = slides.length;
  const limit = cardLayoutBudget(total);
  let used = 0;
  for (const slide of slides) {
    if (!isCardLayoutSlide(slide)) continue;
    if (used < limit) {
      used += 1;
      continue;
    }
    demoteCardSlide(slide, limit, total);
  }
  return slides;
}

/**
 * Turns one card slide into bold lead-in bullets at build time, because the
 * host is too slow for a card row to commit safely. Unlike the budget this is
 * decided per slide from the measured speed, so the chat preview (parsed
 * earlier) may still show cards; the build log says which slides changed.
 */
export function demoteCardSlideForSlowHost(slide) {
  if (!isCardLayoutSlide(slide)) return false;
  const layout = slide.visualType || "structuredCards";
  demoteCardSlide(slide, 0, 0);
  slide.cardBudgetDemoted = null;
  slide.slowHostDemoted = { layout: layout };
  return true;
}

