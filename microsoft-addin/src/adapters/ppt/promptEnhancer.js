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
 * PowerPoint Prompt Enhancer
 * 
 * Cleanly handles PowerPoint prompt interactions without polluting the user's visible prompt text.
 * 
 * @author Sathya AG, Principal Architect, Google
 */

import { cardLayoutBudget } from './cardBudget.js';

/**
 * Word -> number map for explicit slide counts.
 */
const SLIDE_COUNT_WORDS = {
  "a": 1, "an": 1, "single": 1,
  "one": 1, "two": 2, "three": 3, "four": 4, "five": 5,
  "six": 6, "seven": 7, "eight": 8, "nine": 9, "ten": 10
};

/**
 * Extracts an explicitly requested slide count from the user's TYPED instruction.
 *
 * This deliberately requires a creation verb immediately before the number
 * ("create 5 slides", "give me seven slides", "a 5-slide deck") so that
 * editing requests such as "shorten these 5 slides" are NOT treated as a
 * request to generate a deck.
 *
 * Only ever call this with the user's own instruction text - never with the
 * fully assembled prompt, which can carry pasted slide content whose wording
 * would otherwise hijack the count.
 *
 * @param {string} text The user's typed instruction.
 * @returns {number|null} The requested slide count, or null if none was stated.
 */
export function extractRequestedSlideCount(text) {
  if (!text || typeof text !== "string") return null;
  const lower = text.toLowerCase();

  const verbMatch = lower.match(
    /\b(?:create|generate|make|build|produce|draft|prepare|write|add|provide|give\s+me|put\s+together)\s+(?:me\s+)?(?:a\s+set\s+of\s+)?(\d{1,2}|a|an|one|two|three|four|five|six|seven|eight|nine|ten)\s+(?:new\s+|more\s+|additional\s+|short\s+|quick\s+|executive\s+|simple\s+)*slides?\b/
  );
  const deckMatch = lower.match(
    /\b(\d{1,2}|one|two|three|four|five|six|seven|eight|nine|ten)[-\s]slide\s+(?:deck|presentation|summary|overview|story|narrative)\b/
  );
  // "summarize google sustainability efforts in 10 slides" had no verb right
  // before the number, so no count was found and the prompt fell through to
  // the single-slide Executive Summary rules: no title slide, no deck
  // structure, and an "Executive Slide Summary" block on the last slide.
  // A count after "in/into/across/over/as/using/within" is a deck request.
  // "shorten these 5 slides" still does not match - "these" is not one of them.
  const prepMatch = lower.match(
    /\b(?:in|into|across|over|as|using|within)\s+(?:just\s+|only\s+|exactly\s+)?(\d{1,2}|one|two|three|four|five|six|seven|eight|nine|ten)\s+(?:new\s+|short\s+|quick\s+|executive\s+|simple\s+)*slides?\b/
  );
  // "10 slides about X" / "10 slides on X" typed on its own.
  const leadMatch = lower.match(
    /^\s*(\d{1,2}|one|two|three|four|five|six|seven|eight|nine|ten)\s+(?:new\s+|short\s+|quick\s+|executive\s+|simple\s+)*slides?\s+(?:about|on|covering|for|summari[sz]ing)\b/
  );

  const token = (verbMatch && verbMatch[1]) || (deckMatch && deckMatch[1]) ||
    (prepMatch && prepMatch[1]) || (leadMatch && leadMatch[1]) || null;
  if (!token) return null;

  const parsed = SLIDE_COUNT_WORDS[token] || parseInt(token, 10);
  if (!parsed || isNaN(parsed) || parsed < 1 || parsed > 20) return null;
  return parsed;
}

/**
 * Cleanly returns the user prompt without injecting visible system rules into the chat input.
 *
 * @param {string} userPrompt The fully assembled prompt (may include slide context).
 * @param {{userInstruction?: string}} [options] Pass the user's TYPED instruction
 *        separately so an explicit slide count can be detected without being
 *        confused by pasted deck content.
 */
export function enhancePromptForPowerPoint(userPrompt, options = {}) {
  const lowerPrompt = userPrompt.toLowerCase();

  const isSummarizeSlidesAction = typeof window !== "undefined" && !!window.__isSummarizeSlidesAction;
  const isForcedSingleSlide = typeof window !== "undefined" && !!window.__forceSingleSlide;

  // EXPLICIT SLIDE COUNT OVERRIDE
  // "create 5 slides to summarize X" used to trip the summary gate below and be
  // told to emit EXACTLY ONE slide, because the slide-count parser lived in a
  // later, unreachable branch. An explicitly requested count now wins over the
  // executive-visual, chart and summary rule sets.
  // The dedicated Summarize Slides button and the Create Chart flow keep their
  // own contracts and are never overridden.
  const userInstruction = options && typeof options.userInstruction === "string" ? options.userInstruction : "";
  const explicitCount = (isSummarizeSlidesAction || isForcedSingleSlide)
    ? null
    : extractRequestedSlideCount(userInstruction);
  const forceMultiSlide = !!explicitCount && explicitCount >= 2;
  if (forceMultiSlide && typeof console !== "undefined" && console.info) {
    console.info("[PPT] Explicit slide count detected (" + explicitCount + ") - bypassing single-slide rule sets.");
  }

  // Rule set for executive visuals (metric grid, comparison card, 3-column cards, timeline)
  if (!forceMultiSlide && (lowerPrompt.includes("executive visual") || lowerPrompt.includes("metric grid") || lowerPrompt.includes("comparison card") || lowerPrompt.includes("before/after") || lowerPrompt.includes("before and after") || lowerPrompt.includes("timeline") || lowerPrompt.includes("pillar") || lowerPrompt.includes("cards"))) {
    const rules = `
CRITICAL INSTRUCTIONS FOR EXECUTIVE VISUAL GENERATION:
1. CRITICAL SCOPE CONTRACT: Provide EXACTLY ONE definitive visual design on a SINGLE slide. DO NOT output multiple options, and DO NOT generate or repeat previous presentation slides from this conversation under any circumstances.
2. DO NOT output conversational preamble, pleasantries, or conclusions.
3. Supported Visual Archetypes (output as a structured JSON code block):
   - Metric Grid / KPI Cards: "visualType": "metric_grid_3col" containing "title", "subtitle", and an array of 3 "cards" (each with "metric", "title", "subtitle", and "bullets").
   - Three Column Cards: "visualType": "three_column_cards" containing "title", "subtitle", and an array of 3 "cards" (each with "pill", "title", "body").
   - Process Timeline / Roadmap: "visualType": "process_timeline" containing "title", "subtitle", and an array of 3-4 "steps" (each with "step", "title", "description").
   - Before/After Comparison: "visualType": "before_after" containing "title", "subtitle", "before" ({ "title", "bullets" }), and "after" ({ "title", "bullets" }).
`;
    return `${userPrompt}\n\n${rules}`;
  }

  // Rule set for editing / shortening / making smaller / rewriting
  if (lowerPrompt.includes("shorten") || lowerPrompt.includes("smaller") || lowerPrompt.includes("concise") || lowerPrompt.includes("rewrite") || lowerPrompt.includes("punchy") || lowerPrompt.includes("fluff") || lowerPrompt.includes("trim")) {
    const rules = `
CRITICAL INSTRUCTIONS FOR SLIDE EDITING:
1. CRITICAL SCOPE CONTRACT: Output ONLY the finalized content for this single slide. DO NOT generate a multi-slide deck or repeat previous presentation slides from this conversation.
2. Provide EXACTLY ONE definitive, finalized version. NEVER provide multiple alternative options, variations, or choices (e.g. NEVER output "Option 1", "Option 2", or "Or, for an even more minimalist layout...").
3. DO NOT output conversational preamble, introduction, or pleasantries (e.g. NEVER output "Here is a concise and punchy version...", "Sure!", "Here are your revised bullets:").
4. DO NOT output conversational sign-offs or questions (e.g. "Let me know if you need changes").
5. Output ONLY the finalized slide content. Format each point as a clean bullet point with a bold lead-in phrase:
   • **Key Theme**: Crisp, high-impact description.
`;
    return `${userPrompt}\n\n${rules}`;
  }

  // Rule set for chart generation (pie, bar, line, doughnut, column, graph, breakdown)
  if (!forceMultiSlide && (lowerPrompt.includes("chart") || lowerPrompt.includes("pie") || lowerPrompt.includes("bar") || lowerPrompt.includes("graph") || lowerPrompt.includes("visualization") || lowerPrompt.includes("visualize") || lowerPrompt.includes("plot") || lowerPrompt.includes("breakdown") || lowerPrompt.includes("doughnut") || lowerPrompt.includes("column"))) {
    const rules = `
CRITICAL INSTRUCTIONS FOR CHART GENERATION:
1. CRITICAL SCOPE CONTRACT: Generate EXACTLY ONE single slide containing this chart and its accompanying data/takeaways. DO NOT generate, repeat, or expand upon previous presentation slides from earlier in this conversation under any circumstances.
2. Output a structured JSON code block with the exact data metrics:
\`\`\`json
{
  "chartType": "pie",
  "title": "Chart Title",
  "data": [
    { "label": "Category / Label", "value": 12345 }
  ]
}
\`\`\`
Supported chartType values: "pie", "doughnut", "bar", "column", "line". Use exact numeric values (not strings).
3. STRICT SINGLE-SLIDE OUTPUT FORMAT:
   - Your ENTIRE response must contain AT MOST ONE markdown heading ("##"). NEVER emit "## Slide 1" and "## Slide 2", and NEVER add a separate "Strategic Insights", "Key Takeaways", or "Summary" heading - those bullets belong on the same slide, directly beneath the chart.
   - NEVER output a horizontal rule ("---"); it is interpreted as a slide break.
4. NEVER output a "visualType" JSON block (for example "three_column_cards", "metric_grid_3col", "process_timeline"). The ONLY JSON block permitted in this response is the chart specification above.
5. Provide 2-3 executive bullet points with bold lead-ins highlighting strategic insights from the chart.
6. DO NOT output conversational preamble or pleasantries.
`;
    return `${userPrompt}\n\n${rules}`;
  }

  // Rule set for image generation
  if (lowerPrompt.includes("image") || lowerPrompt.includes("picture") || lowerPrompt.includes("illustration") || lowerPrompt.includes("drawing") || lowerPrompt.includes("visual art")) {
    const rules = `
CRITICAL INSTRUCTIONS FOR IMAGE GENERATION:
1. TRIGGER TOOL DIRECTLY: Invoke the native image generation tool (imageGenerationSpec) directly to produce and return the visual image.
2. ABSOLUTELY NO JSON SCHEMAS OR TEXT SPECIFICATIONS: NEVER output raw JSON code blocks, JSON schemas, or "visual_request" objects in your response text. Output the generated image directly.
3. DO NOT output conversational preamble, pleasantries, slide outline text, or image descriptions.
`;
    return `${userPrompt}\n\n${rules}`;
  }

  // Rule set for Executive Summary (Summarize Slides allows up to 5 slides with dedicated takeaways slide; other summaries remain single-slide)
  if (!forceMultiSlide && (lowerPrompt.includes("summarize") || lowerPrompt.includes("executive summary") || lowerPrompt.includes("slide summary") || lowerPrompt.includes("key takeaway"))) {
    const isSummarizeSlides = (typeof window !== "undefined" && window.__isSummarizeSlidesAction) ||
      lowerPrompt.includes("summarize slides") ||
      lowerPrompt.includes("up to 5 slides") ||
      lowerPrompt.includes("[summarize slides]");

    if (isSummarizeSlides) {
      const rules = `
CRITICAL INSTRUCTIONS FOR SUMMARIZE SLIDES GENERATION (STRICTLY 3 TO 5 SLIDES MAXIMUM):
1. Provide a comprehensive Executive Summary presentation across multiple slides (STRICTLY BETWEEN 3 AND 5 SLIDES MAXIMUM):
   - You must generate AT MOST 5 slides total (strictly 3 to 5 slides). NEVER generate 6 or more slides.
   - Separate distinct topics, tables, and visual charts into their own slides (e.g. ## Slide 1: [Executive Overview / Main Metrics Table], ## Slide 2: [Category Breakdown / Visual Chart / Details Table], etc.).
   - If there are multiple tables or data sets, place each table on its own appropriate slide.
   - If a visual chart represents data, output a structured JSON code block with the exact data metrics (chartType: "doughnut" or "bar", title: "...", data: [...]) so our client presentation engine can render a crisp chart.
   - Break down the key takeaways into a dedicated single slide titled "## 📊 Executive Summary: Key Takeaways" with at most 3 to 4 concise executive bullet points (max 15 words per bullet). CRITICAL: Generate strictly at most 4 bullet points; do NOT exceed 4 bullet points or create continuation slides.
   - Format each slide with a clear markdown header (## Slide 1: [Title], ## Slide 2: [Title], etc.) so each section generates its own slide.
2. STRICT CLOSED-BOOK GROUNDING CONTRACT:
   - You are operating in 100% STRICT CLOSED-BOOK MODE based SOLELY on the provided slide context.
   - You must synthesize information ONLY AND EXCLUSIVELY from the text and data present in the selected slides.
   - NEVER extrapolate, bring in external industry knowledge, or introduce topics, facts, or assumptions that do not appear in the selected slides.
   - If a concept, fact, or metric is not explicitly stated in the selected slides, DO NOT mention it.
3. DO NOT output conversational preamble or pleasantries.
`;
      return `${userPrompt}\n\n${rules}`;
    }

    const rules = `
CRITICAL INSTRUCTIONS FOR EXECUTIVE SUMMARY GENERATION:
1. Provide a professional Executive Summary with EXACTLY ONE single slide:
   - "## 📊 Executive Slide Summary" containing the high-level context and structured Markdown Table (| Metric / Focus Area | FY Progress Status | Target Benchmark |) summarizing key data directly from the slides.
   - Beneath the table, provide concise bullet points / key takeaways synthesizing the core findings.
   - STRICT SINGLE-SLIDE CONSTRAINT: Do NOT generate multiple slides or a separate "Key Takeaways" slide.
2. STRICT CLOSED-BOOK GROUNDING CONTRACT:
   - You are operating in 100% STRICT CLOSED-BOOK MODE based SOLELY on the provided slide context.
   - You must synthesize information ONLY AND EXCLUSIVELY from the text and data present in the selected slides.
   - NEVER extrapolate, bring in external industry knowledge, or introduce topics, facts, or assumptions that do not appear in the selected slides.
   - If a concept, fact, or metric is not explicitly stated in the selected slides, DO NOT mention it.
3. Visual Chart: If the summary contains quantitative breakdown, comparisons, or metrics directly stated in the slides, output a structured JSON code block with the exact data metrics (chartType: "doughnut", title: "...", data: [...]) so our client presentation engine can render a crisp, high-resolution chart.
4. Content layout:
   - If the slides contain structured comparisons or metrics, provide a clean Markdown Table (| Category | Metric | Share / Value |) summarizing data directly from the slides.
   - Place all executive takeaway bullets (2 to 4 concise bullets with bold lead-ins) beneath the table.
5. DO NOT output conversational preamble or pleasantries.
`;
    return `${userPrompt}\n\n${rules}`;
  }

  // Rule set for slide generation
  if (lowerPrompt.includes("slide") || lowerPrompt.includes("presentation") || lowerPrompt.includes("deck") || lowerPrompt.includes("table") || lowerPrompt.includes("pitch")) {
    // An explicit count taken from the user's own instruction always wins; the
    // legacy scan of the whole prompt stays as a fallback for flows that do not
    // pass the instruction separately.
    let requestedCount = explicitCount || null;
    if (!requestedCount) {
      const countMatch = lowerPrompt.match(/\b(?:create|generate|make|build|provide|give\s+me|add)?\s*(an?|single|\d+|one|two|three|four|five|six|seven|eight|nine|ten)\s+(?:new\s+|executive\s+)?slides?(?!\s+deck)\b/i);
      if (countMatch && countMatch[1]) {
        const token = countMatch[1].toLowerCase();
        requestedCount = SLIDE_COUNT_WORDS[token] || parseInt(token, 10);
      }
    }

    let countConstraint = "";
    if (requestedCount && requestedCount > 0) {
      const slideWord = requestedCount === 1 ? "slide" : "slides";
      const slideList = Array.from({ length: requestedCount }, (_, idx) => `- ## Slide ${idx + 1}: [Title]`).join("\n");
      countConstraint = `
CRITICAL CONSTRAINT - EXACT SLIDE COUNT:
The user explicitly requested EXACTLY ${requestedCount} ${slideWord}. You MUST generate ALL ${requestedCount} ${slideWord} from ## Slide 1 through ## Slide ${requestedCount}:
${slideList}
NEVER output fewer or more than ${requestedCount} ${slideWord} under any circumstances.
`;
      if (requestedCount >= 2) {
        const cardBudget = cardLayoutBudget(requestedCount);
        countConstraint += `
CARD LAYOUT BUDGET: At most ${cardBudget} of the ${requestedCount} slides may use a card layout ("three_column_cards", "metric_grid_3col", "process_timeline", "before_after", or 2-4 bullets with bold lead-ins, which are also drawn as cards). Spend them on the slides where a visual adds the most. Give every other content slide a markdown table, a chart, or plain bullets with bold lead-ins - any card layouts beyond the budget are drawn as plain bullets.
`;
      }
    }

    // A 1-2 slide request is a quick answer, not a deck: spending half of it
    // on a title slide leaves one slide (or none) for the actual content.
    const slideOneRule = (requestedCount && requestedCount < 3)
      ? `- NO TITLE SLIDE: The user asked for only ${requestedCount} ${requestedCount === 1 ? "slide" : "slides"}. Do NOT create a title, cover, or executive-summary slide. Every slide, including Slide 1, MUST be a content slide with a table, chart, card layout, or 3-4 bullets with bold lead-ins.`
      : `- SLIDE 1 MUST BE A TITLE & SUMMARY SLIDE: Slide 1 must contain ONLY the presentation title, subtitle, and ONE single executive summary small paragraph written as plain prose (no bullet points). Save specific operational details, findings, metrics, data tables, and pillars for subsequent slides (Slide 2 onwards).`;

    const rules = `
CRITICAL STRUCTURE CONTRACT FOR SLIDE GENERATION:${countConstraint}
1. Provide EXACTLY ONE definitive presentation version. DO NOT output multiple alternatives, conversational preamble, pleasantries, or filler. Output the presentation slides directly.
2. For EVERY slide, strictly follow this standardized structure:

---
## Slide {N}: [Relevant Emoji] [Title: STRICTLY 3 TO 5 WORDS MAXIMUM]
### [Contextual Subtitle / One-sentence Takeaway]

[CONTENT AREA - CHOOSE APPROPRIATE VISUAL OR CONTENT FORMAT]:
- For narrative, strategy, or core pillars: 3 to 4 crisp executive bullet points with bold lead-ins (• **Key Pillar**: Impactful description) OR a structured JSON code block with "visualType": "three_column_cards" (with array of 3 cards: { "pill", "title", "body", "footnote" }). Empty space inside boxes should be filled with concrete regional/operational details or footnote citations.
- For KPI metrics & growth drivers: A structured JSON code block with "visualType": "metric_grid_3col" (with array of 3 cards: { "metric", "title", "subtitle", "bullets", "footnote" }).
- For roadmap, stages, or implementation milestones: A structured JSON code block with "visualType": "process_timeline" (with array of 3-4 steps: { "step", "title", "description" }).
- For quantitative data / financial metrics: An Executive Markdown Table with 3 to 4 columns (| Dimension | Metric | Strategic Impact |). You MAY also include a companion Data Chart (\`\`\`json { "chartType": "doughnut|bar|pie", "title": "...", "data": [...] } \`\`\`) to render a side-by-side table + chart presentation.
- For distribution / breakdown / share metrics: A structured JSON code block (\`\`\`json { "chartType": "doughnut|bar|pie", "title": "...", "data": [...] } \`\`\`) accompanied by an Executive Markdown Table or 2-3 narrative bullet points.

3. CRITICAL VISUAL & CONTENT RULES:
- PRESENTATION DIVERSITY & THEMATIC SEPARATION (Decks with 3+ slides): NEVER make slides repetitive or redundant! Every slide MUST explore a completely distinct strategic angle:
  * Slide 1: Dedicated Title and Executive Summary Slide containing ONLY a clean Presentation Title, Subtitle, and ONE small Executive Summary paragraph (2-4 sentences max) explaining the strategic scope. NEVER include "Deck Scope & Outline", upcoming slide lists (Slide 2, Slide 3, etc.), operational bullets, metrics, charts, or tables on Slide 1.
  * Slide 2 (Key Pillars / Infrastructure / Regional Scope): Use 3-column cards ("three_column_cards") or structured bullets with bold lead-ins. Include substantive details or footnote source citations inside each pillar box.
  * Slide 3 (Market Leadership & Production Data): At least 1 slide MUST feature an Executive Markdown Table (3-4 columns) or Metric Grid ("metric_grid_3col") or Table + Chart showing core metrics and export share.
  * Slide 4 (NO REDUNDANCY WITH SLIDE 3): Slide 4 MUST NEVER cover the same topic as Slide 3. If Slide 3 covered production or export metrics, Slide 4 MUST pivot to a different domain such as Agronomic Innovation & Tech, Supply Chain & Logistics, or Quality & Variety differentiation.
  * Slide 5 (Strategic Outlook / Recommendations / Roadmap): Slide 5 MUST be rich and substantive! NEVER output an empty or near-empty slide. Include at least 3-4 concrete strategic recommendations, future market forecasts, or an implementation roadmap ("process_timeline").
- STRICT TITLE LENGTH: Every slide title MUST be concise, strictly between 3 to 5 words (e.g. "☕ Brazilian Coffee Leadership"). NEVER write long multi-line sentences as slide titles.
- NEVER OUTPUT PSEUDO-TEXT LABELS: NEVER output labels like "Metric Grid", "Comparison Card", or raw placeholder text in the slide content.
- EACH CHART MUST BE 100% UNIQUE: NEVER repeat or duplicate the same chart, metrics, or title across multiple slides. If a deck has 5 to 10 slides, at most 2 should contain a data chart, and each MUST cover a completely different topic and metric.
- CRITICAL CHART TITLE RULE: Every chart title MUST be short and punchy (maximum 2 to 5 words). NEVER create long, rambling titles or include parenthetical details in the title.
${slideOneRule}
4. Do NOT output internal design metadata, font sizes (like "Title Size: 44"), hex colors (like "Color: #..."), or raw "Visual Concept:" labels. Keep the output clean, executive-ready presentation content.
5. Separate every slide cleanly with a horizontal rule "---".
`;
    return `${userPrompt}\n\n${rules}`;
  }
  return userPrompt;
}

/**
 * Initializes prompt enhancer (no-op to prevent polluting UI input text).
 */
export function initPromptEnhancer() {
  // Kept clean to avoid displaying bracketed formatting rules to the user
}
