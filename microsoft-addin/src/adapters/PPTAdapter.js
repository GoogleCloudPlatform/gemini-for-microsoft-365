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
 * Host Adapter Implementation for Microsoft PowerPoint (PowerPoint.run + Office Common API)
 * 
 * Provides complete multi-slide presentation generation, intelligent outline/table
 * unpacking, visual slide deck preview, executive PowerPoint styling, and in-taskpane diagnostics.
 * 
 * @author Sathya AG, Principal Architect, Google
 */

import { parseSlides, extractCleanBulletPoints } from './ppt/slideParser.js';
import { buildPresentation, compressImageForPowerPoint, insertOnCurrentSlide } from './ppt/slideBuilder.js';
import { initSlidePreviewObserver, injectPowerPointStyles } from './ppt/slidePreviewUI.js';
import { initPromptEnhancer, enhancePromptForPowerPoint } from './ppt/promptEnhancer.js';
import { initPowerPointDiagnostics } from './ppt/pptDiagnostics.js';
import { resetBuildTimings, beginPhase, logBuildTimings } from './ppt/buildTimings.js';

/* ------------------------------------------------------------------ */
/* Slide-build lease                                                   */
/*                                                                     */
/* This used to be a bare `window.__isGeneratingSlides` boolean. Any    */
/* await inside insertContent() that never settled left it true for the*/
/* life of the page, and every later insert then failed with "a slide  */
/* build is already in progress" - which is exactly what the unguarded */
/* Common-API picture fallback did.                                    */
/*                                                                     */
/* A lease is reclaimable. The build has to keep ticking it, and a     */
/* lease that has not ticked recently is treated as abandoned and      */
/* taken over. The tick comes from real build output (see              */
/* pptDiagnostics.logToPPTConsole), never from a timer: a pending      */
/* promise does not block the event loop, so an interval would happily */
/* tick straight through a hang and the lease would never go stale.    */
/*                                                                     */
/* The lease also carries an owner token, so two callers that stop     */
/* waiting in the same tick cannot both believe they hold it, and a    */
/* caller that finishes early cannot release a lease someone else has  */
/* since taken over.                                                   */
/* ------------------------------------------------------------------ */

// No build output for this long => the build is considered abandoned.
// Must exceed the longest legitimately silent stretch, which is the image
// decode budget in compressImageForPowerPoint (12s).
const BUILD_LEASE_STALE_MS = 20000;

// Never queue behind a live build indefinitely.
const BUILD_LEASE_MAX_WAIT_MS = 90000;

let buildLeaseSeq = 0;

function buildLeaseIsLive() {
  const lease = (typeof window !== "undefined") ? window.__slideBuildLease : null;
  if (!lease) return false;
  return Date.now() - (lease.tickedAt || 0) < BUILD_LEASE_STALE_MS;
}

export function touchSlideBuildLease() {
  if (typeof window === "undefined" || !window.__slideBuildLease) return;
  window.__slideBuildLease.tickedAt = Date.now();
}

if (typeof window !== "undefined") {
  // pptDiagnostics reaches this through the global to avoid an import cycle.
  window.__touchSlideBuildLease = touchSlideBuildLease;
}

export async function acquireSlideBuildLease(onWaiting) {
  const waitStarted = Date.now();
  let announced = false;

  while (buildLeaseIsLive() && Date.now() - waitStarted < BUILD_LEASE_MAX_WAIT_MS) {
    if (!announced && typeof onWaiting === "function") {
      announced = true;
      onWaiting();
    }
    await new Promise((r) => setTimeout(r, 150));
  }

  if (buildLeaseIsLive()) {
    throw new Error("A slide build is still running. Please try again in a moment.");
  }

  const abandoned = window.__slideBuildLease;
  if (abandoned) {
    const silentFor = Math.round((Date.now() - (abandoned.tickedAt || 0)) / 1000);
    if (typeof window.logToDiagBox === "function") {
      window.logToDiagBox(
        `[PPT] The previous slide build produced no output for ${silentFor}s - reclaiming it and continuing.`,
        true
      );
    }
  }

  const token = `build-${Date.now()}-${++buildLeaseSeq}`;
  window.__slideBuildLease = { token, tickedAt: Date.now() };
  // Kept in step for anything still reading the original flag.
  window.__isGeneratingSlides = true;
  return token;
}

export function releaseSlideBuildLease(token) {
  if (typeof window === "undefined") return;
  // Only the holder may release. Without this, a caller that finished early
  // would clear a lease a later caller had already taken over.
  if (window.__slideBuildLease && window.__slideBuildLease.token !== token) return;
  window.__slideBuildLease = null;
  window.__isGeneratingSlides = false;
}

export class PPTAdapter {
  constructor() {
    this.name = "PowerPoint";
    
    try {
      injectPowerPointStyles();
      initSlidePreviewObserver(this);
      initPromptEnhancer();
      initPowerPointDiagnostics();
    } catch (e) {
      console.warn("PPTAdapter initialization warning:", e);
    }
  }

  /**
   * Safely extracts all text and table contents from a PowerPoint slide object.
   * Completely resilient against non-text shapes, groups, and API version differences.
   */
  async extractSlideText(slide, slideNum, context) {
    const slideLines = [];
    try {
      const shapes = slide.shapes;
      shapes.load("items/id,items/type,items/name");
      await context.sync();

      if (!shapes.items || shapes.items.length === 0) {
        return { slideNumber: slideNum, id: slide.id || `slide-${slideNum}`, text: "" };
      }

      // Track text ranges and tables safely based on verified shape type
      const textItemTrackers = [];
      const tableItemTrackers = [];

      for (const shape of shapes.items) {
        try {
          const type = shape.type;
          // Check for Table shape (starting in PowerPointApi 1.8)
          if (type === "Table" || (typeof PowerPoint !== "undefined" && PowerPoint.ShapeType && type === PowerPoint.ShapeType.table)) {
            if (typeof shape.getTable === "function") {
              const table = shape.getTable();
              table.load("values");
              tableItemTrackers.push(table);
            }
          } else if (
            type === "TextBox" || 
            type === "GeometricShape" || 
            type === "Placeholder" || 
            type === "Callout" ||
            !type // Fallback if type property was not populated
          ) {
            // Text-bearing shapes
            if (shape.textFrame) {
              const tr = shape.textFrame.textRange;
              tr.load("text");
              textItemTrackers.push(tr);
            }
          }
        } catch (shapeErr) {
          console.warn(`[PPTAdapter] Shape pre-check skipped on slide ${slideNum}:`, shapeErr);
        }
      }

      // Try batch sync for efficiency
      let batchSuccess = false;
      if (textItemTrackers.length > 0 || tableItemTrackers.length > 0) {
        try {
          await context.sync();
          batchSuccess = true;
        } catch (batchErr) {
          console.warn(`[PPTAdapter] Batch shape sync failed on slide ${slideNum}, falling back to per-shape sync:`, batchErr);
        }
      }

      if (batchSuccess) {
        for (const tr of textItemTrackers) {
          try {
            const val = (tr.text || "").trim();
            if (val) slideLines.push(val);
          } catch (_) {}
        }
        for (const tb of tableItemTrackers) {
          try {
            if (tb.values && Array.isArray(tb.values)) {
              for (const row of tb.values) {
                if (Array.isArray(row)) {
                  const rowStr = row.map(c => String(c ?? "").trim()).join(" | ");
                  if (rowStr.replace(/[|\s]/g, "")) {
                    slideLines.push(`| ${rowStr} |`);
                  }
                }
              }
            }
          } catch (_) {}
        }
      } else {
        // Fallback: Per-shape isolated extraction so a problematic shape never blocks other shapes
        for (const shape of shapes.items) {
          try {
            if (shape.type === "Table" && typeof shape.getTable === "function") {
              const table = shape.getTable();
              table.load("values");
              await context.sync();
              if (table.values && Array.isArray(table.values)) {
                for (const row of table.values) {
                  if (Array.isArray(row)) {
                    const rowStr = row.map(c => String(c ?? "").trim()).join(" | ");
                    if (rowStr.replace(/[|\s]/g, "")) slideLines.push(`| ${rowStr} |`);
                  }
                }
              }
            } else if (shape.textFrame) {
              const tr = shape.textFrame.textRange;
              tr.load("text");
              await context.sync();
              const val = (tr.text || "").trim();
              if (val) slideLines.push(val);
            }
          } catch (innerErr) {
            // Silently swallow per-shape errors so other shapes succeed
          }
        }
      }
    } catch (slideErr) {
      console.warn(`[PPTAdapter] Error extracting text from slide ${slideNum}:`, slideErr);
    }

    const slideText = slideLines.join("\n").trim();
    console.log(`📊 [PPTAdapter] Extracted ${slideLines.length} text/table blocks from Slide ${slideNum} (${slideText.length} chars)`);
    return {
      slideNumber: slideNum,
      id: slide.id || `slide-${slideNum}`,
      text: slideText
    };
  }

  // On-demand selection extraction (strictly user-triggered on button click, NO background event listeners)
  async getSelectedSlidesText() {
    const selectedSlidesData = [];
    try {
      if (typeof PowerPoint !== 'undefined') {
        await PowerPoint.run(async (context) => {
          // Map all slide IDs in presentation to determine real 1-based slide numbers
          const allSlides = context.presentation.slides;
          allSlides.load("items/id");

          let selectedSlides = null;
          if (context.presentation.getSelectedSlides) {
            selectedSlides = context.presentation.getSelectedSlides();
            selectedSlides.load("items/id");
          }
          await context.sync();

          const slideIndexMap = {};
          if (allSlides.items) {
            allSlides.items.forEach((s, idx) => {
              if (s.id) slideIndexMap[s.id] = idx + 1;
            });
          }

          if (selectedSlides && selectedSlides.items && selectedSlides.items.length > 0) {
            for (let i = 0; i < selectedSlides.items.length; i++) {
              try {
                const slide = selectedSlides.items[i];
                const slideNum = (slide.id && slideIndexMap[slide.id]) ? slideIndexMap[slide.id] : (i + 1);
                const slideData = await this.extractSlideText(slide, slideNum, context);
                selectedSlidesData.push(slideData);
              } catch (slideLoopErr) {
                console.warn(`[PPTAdapter] Error in selected slide loop item ${i}:`, slideLoopErr);
              }
            }
          }
        });
      }
    } catch (err) {
      console.warn("PowerPoint getSelectedSlidesText warning:", err);
    }
    return selectedSlidesData;
  }

  // Read currently highlighted shape(s), text frame, or table on the active slide on demand
  async getSelectedShapeText() {
    let selectedText = "";
    try {
      if (typeof PowerPoint !== 'undefined') {
        await PowerPoint.run(async (context) => {
          // 1. Try reading actively highlighted text range first (PowerPointApi 1.5+)
          if (context.presentation && typeof context.presentation.getSelectedTextRangeOrNullObject === 'function') {
            try {
              const textRange = context.presentation.getSelectedTextRangeOrNullObject();
              textRange.load("text");
              await context.sync();
              if (!textRange.isNullObject && textRange.text && textRange.text.trim().length > 0) {
                selectedText = textRange.text.trim();
                return;
              }
            } catch (trErr) {
              // Ignore if no text range or not supported in this host version
            }
          }

          // 2. Read selected shapes (text boxes, shapes, or tables)
          if (context.presentation.getSelectedShapes) {
            const selection = context.presentation.getSelectedShapes();
            selection.load("items");
            await context.sync();

            if (selection.items && selection.items.length > 0) {
              const textTrackers = [];
              const tableTrackers = [];

              for (const shape of selection.items) {
                try {
                  const type = shape.type;
                  const isTable = type === "Table" || 
                                  (typeof PowerPoint !== "undefined" && PowerPoint.ShapeType && type === PowerPoint.ShapeType.table);

                  if (isTable) {
                    if (typeof shape.getTable === "function") {
                      const table = shape.getTable();
                      table.load("values");
                      tableTrackers.push(table);
                    }
                  } else {
                    if (shape.textFrame) {
                      const tr = shape.textFrame.textRange;
                      tr.load("text");
                      textTrackers.push(tr);
                    }
                  }
                } catch (_) {}
              }

              if (textTrackers.length > 0 || tableTrackers.length > 0) {
                let syncSuccess = false;
                try {
                  await context.sync();
                  syncSuccess = true;
                } catch (batchErr) {
                  console.warn("[PPTAdapter] Batch shape sync failed in getSelectedShapeText, falling back to per-shape sync:", batchErr);
                }

                const texts = [];
                if (syncSuccess) {
                  for (const tr of textTrackers) {
                    try {
                      const val = (tr.text || "").trim();
                      if (val) texts.push(val);
                    } catch (_) {}
                  }
                  for (const tb of tableTrackers) {
                    try {
                      if (tb.values && Array.isArray(tb.values)) {
                        const tableRows = [];
                        for (let r = 0; r < tb.values.length; r++) {
                          const row = tb.values[r];
                          if (Array.isArray(row)) {
                            const rowStr = row.map(c => String(c ?? "").trim()).join(" | ");
                            if (rowStr.replace(/[|\s]/g, "")) {
                              tableRows.push(`| ${rowStr} |`);
                              if (r === 0 && tb.values.length > 1) {
                                const sep = row.map(() => "---").join(" | ");
                                tableRows.push(`| ${sep} |`);
                              }
                            }
                          }
                        }
                        if (tableRows.length > 0) {
                          texts.push(tableRows.join("\n"));
                        }
                      }
                    } catch (_) {}
                  }
                } else {
                  // Fallback per-shape isolated extraction
                  for (const shape of selection.items) {
                    try {
                      const type = shape.type;
                      const isTable = type === "Table" || 
                                      (typeof PowerPoint !== "undefined" && PowerPoint.ShapeType && type === PowerPoint.ShapeType.table);

                      if (isTable && typeof shape.getTable === "function") {
                        const table = shape.getTable();
                        table.load("values");
                        await context.sync();
                        if (table.values && Array.isArray(table.values)) {
                          const tableRows = [];
                          for (let r = 0; r < table.values.length; r++) {
                            const row = table.values[r];
                            if (Array.isArray(row)) {
                              const rowStr = row.map(c => String(c ?? "").trim()).join(" | ");
                              if (rowStr.replace(/[|\s]/g, "")) {
                                tableRows.push(`| ${rowStr} |`);
                                if (r === 0 && table.values.length > 1) {
                                  const sep = row.map(() => "---").join(" | ");
                                  tableRows.push(`| ${sep} |`);
                                }
                              }
                            }
                          }
                          if (tableRows.length > 0) texts.push(tableRows.join("\n"));
                        }
                      } else if (shape.textFrame) {
                        const tr = shape.textFrame.textRange;
                        tr.load("text");
                        await context.sync();
                        const val = (tr.text || "").trim();
                        if (val) texts.push(val);
                      }
                    } catch (_) {}
                  }
                }

                selectedText = texts.join("\n\n").trim();
              }
            }
          }
        });
      }
    } catch (err) {
      console.warn("PowerPoint getSelectedShapeText error:", err);
    }
    return selectedText;
  }

  // Read currently highlighted text or selected text shape on demand via PowerPoint.run
  async getSelectedText() {
    // Read selected shape(s) text cleanly through PowerPoint Rich API
    // Avoids Office.context.document.getSelectedDataAsync which triggers Chrome Enterprise SDP/DLP popups
    const shapeText = await this.getSelectedShapeText();
    if (shapeText && shapeText.trim().length > 0) {
      return shapeText.trim();
    }

    // Return empty string if no specific text shape is selected
    // (Multi-slide text extraction is handled cleanly on demand by getSelectedSlidesText)
    return "";
  }

  // Read full presentation text across all slides and shapes safely using index-based traversal
  async getFullDocumentText() {
    let fullText = "";
    try {
      if (typeof PowerPoint !== 'undefined') {
        await PowerPoint.run(async (context) => {
          const slides = context.presentation.slides;
          const countResult = slides.getCount();
          await context.sync();

          const total = countResult.value || 0;
          const slideTexts = [];

          for (let i = 0; i < total; i++) {
            try {
              const slide = slides.getItemAt(i);
              const slideData = await this.extractSlideText(slide, i + 1, context);
              if (slideData.text && slideData.text.length > 0) {
                slideTexts.push(`[Slide ${i + 1}]:\n${slideData.text}`);
              }
            } catch (slideErr) {
              console.warn(`[PPTAdapter] Error in getFullDocumentText slide ${i + 1}:`, slideErr);
            }
          }

          fullText = slideTexts.join("\n\n---\n\n");
        });
      }
    } catch (e) {
      console.warn("PPT full text read error:", e);
    }
    return fullText.trim();
  }

  // Parse HTML or Markdown content into executive slide structures
  async parseSlidesFromHtml(htmlContent, rawText = "", options = {}) {
    return parseSlides(htmlContent, rawText, options);
  }

  // Insert AI content as executive PowerPoint slides with exact positioning & visuals
  async insertContent(htmlContent, rawText = "", options = {}) {
    const hasImages = options && Array.isArray(options.base64Images) && options.base64Images.length > 0;
    if (!htmlContent && !rawText && !hasImages) return;
    const debugStatus = typeof document !== "undefined" ? document.getElementById("debugStatus") : null;
    const loadingText = typeof document !== "undefined" ? document.getElementById("loading") : null;
    const isReplace = options.mode === "replace" || options.mode === "replace_draft";

    // Take the build lease before touching the deck. This sits OUTSIDE the try
    // on purpose: the finally below releases it, and it must never release a
    // lease this call did not take.
    //
    // A lease that has stopped producing output is reclaimed rather than
    // respected, so a hang inside one build can no longer disable every
    // subsequent insert for the life of the page.
    const buildLeaseToken = await acquireSlideBuildLease(() => {
      if (loadingText) {
        loadingText.innerText = "⏳ Waiting for the current slide build to finish...";
        loadingText.style.display = "block";
      }
    });

    // Instrumentation only. The clock starts after the lease is held so that
    // time spent queued behind someone else's build is not blamed on ours.
    resetBuildTimings();
    const buildStartedAt = (typeof performance !== "undefined" && performance.now)
      ? performance.now()
      : Date.now();
    const elapsedMs = () => ((typeof performance !== "undefined" && performance.now)
      ? performance.now()
      : Date.now()) - buildStartedAt;
    let timingsReported = false;

    try {
      if (typeof PowerPoint === 'undefined') {
        throw new Error("PowerPoint Office.js environment is not available.");
      }

      // 1. If replacing and an active shape/text is selected, perform in-place text replacement in the shape
      // UNLESS options.imageOnly is true, where we want to insert/replace with the picture
      if (isReplace && !options.imageOnly) {
        let shapeReplaced = false;
        try {
          await PowerPoint.run(async (context) => {
            if (context.presentation.getSelectedShapes) {
              const selection = context.presentation.getSelectedShapes();
              selection.load("items/textFrame");
              await context.sync();
              if (selection.items && selection.items.length > 0) {
                const shape = selection.items[0];
                if (shape.textFrame) {
                  const cleanBullets = extractCleanBulletPoints(htmlContent, rawText);
                  shape.textFrame.textRange.text = cleanBullets;
                  try {
                    shape.textFrame.textRange.font.size = 14;
                  } catch (_) {}
                  await context.sync();

                  // Bold lead-in phrases before colons
                  try {
                    const lines = cleanBullets.split("\n");
                    let charOffset = 0;
                    for (const l of lines) {
                      const colonIdx = l.indexOf(":");
                      const dashIdx = l.indexOf("—");
                      const sepIdx = colonIdx > 0 ? colonIdx : (dashIdx > 0 ? dashIdx : -1);
                      if (sepIdx > 0 && sepIdx < 45 && typeof shape.textFrame.textRange.getSubstring === "function") {
                        try {
                          const leadIn = shape.textFrame.textRange.getSubstring(charOffset, sepIdx + 1);
                          leadIn.font.bold = true;
                        } catch (_) {}
                      }
                      charOffset += l.length + 1;
                    }
                  } catch (_) {}

                  shapeReplaced = true;
                }
              }
            }
          });
        } catch (shapeErr) {
          console.warn("Shape replacement check:", shapeErr);
        }

        if (shapeReplaced) {
          if (debugStatus) debugStatus.innerText = "✅ Replaced text in slide!";
          if (loadingText) loadingText.style.display = "none";
          // No slides were built; an empty breakdown would just be noise.
          timingsReported = true;
          return [{ title: "Updated Shape", body: rawText }];
        }
      }

      // 1b. If inserting on current slide
      options.layoutChoice = options.layoutChoice || "auto";
      const isInsertCurrent = options.mode === "insert_current" || options.mode === "insert_current_slide";
      if (isInsertCurrent) {
        if (debugStatus) debugStatus.innerText = "Targeting current slide...";
        if (loadingText) {
          loadingText.innerText = "⚡ Inserting onto current slide...";
          loadingText.style.display = "block";
        }
        const endParseCurrent = beginPhase("parse + chart render");
        const slideStructures = await this.parseSlidesFromHtml(htmlContent, rawText, options);
        endParseCurrent(`${(slideStructures || []).length} structure(s)`);
        if (!slideStructures || slideStructures.length === 0) {
          throw new Error("Slide parser returned 0 slide structures.");
        }
        const endInsertCurrent = beginPhase("insert on current slide");
        await insertOnCurrentSlide(slideStructures, options);
        endInsertCurrent();
        if (debugStatus) debugStatus.innerText = "✅ Inserted on current slide!";
        if (loadingText) loadingText.style.display = "none";
        timingsReported = true;
        logBuildTimings(elapsedMs());
        return slideStructures;
      }

      // 2. Otherwise parse slide structures and build / replace slide(s)
      if (debugStatus) debugStatus.innerText = "Parsing presentation structure...";
      // Sample the deck ONCE, here, for the whole build.
      //
      // The reset used to live in buildPresentation, which meant the deck was
      // read twice per build: once here to give parse-time charts a palette,
      // and again in the builder purely to apply the canvas hint from the
      // chosen layout. The hint is applied to an already-collected sample, so
      // the second read was 10 seconds of I/O for arithmetic. The builder now
      // re-derives its tokens from this sample instead; resetting here is what
      // makes that sample this build's, and not the previous build's.
      const endScan = beginPhase("deck scan (theme sample)");
      try {
        const { getDeckTheme, resetDeckTheme } = await import("./ppt/deckTheme.js");
        resetDeckTheme();
        await getDeckTheme();
      } catch (themeErr) {
        console.warn("[PPTAdapter] Deck theme pre-sample notice:", themeErr);
      }
      endScan();
      const endParse = beginPhase("parse + chart render");
      const slideStructures = await this.parseSlidesFromHtml(htmlContent, rawText, options);
      endParse(`${(slideStructures || []).length} structure(s)`);

      if (!slideStructures || slideStructures.length === 0) {
        throw new Error("Slide parser returned 0 slide structures.");
      }

      const onProgress = (prog) => {
        // Per-slide proof of life, on top of the per-log-line tick.
        touchSlideBuildLease();
        const msg = `⚡ ${isReplace ? 'Replacing' : 'Creating'} slide ${prog.current}/${prog.total}: "${prog.title}"...`;
        if (debugStatus) debugStatus.innerText = msg;
        if (loadingText) {
          loadingText.innerText = msg;
          loadingText.style.display = "block";
        }
        if (typeof options.onProgress === "function") {
          options.onProgress(prog);
        }
      };

      const buildResult = await buildPresentation(slideStructures, options, onProgress);

      const createdCount = (buildResult && typeof buildResult.created === "number")
        ? buildResult.created
        : slideStructures.length;
      const skipped = (buildResult && buildResult.failures) ? buildResult.failures.length : 0;
      const stoppedAt = (buildResult && buildResult.abortedAtSlide) ? buildResult.abortedAtSlide : 0;

      if (debugStatus && stoppedAt > 0) {
        // Not "1 skipped": every slide after this one was never attempted.
        debugStatus.innerText =
          `⚠️ Stopped at slide ${stoppedAt} of ${slideStructures.length} - PowerPoint stopped responding. ` +
          `${createdCount} slide(s) kept. Reload the task pane before trying again.`;
      } else if (debugStatus) {
        debugStatus.innerText = skipped > 0
          ? `⚠️ ${isReplace ? 'Replaced' : 'Created'} ${createdCount}/${slideStructures.length} slide(s) — ${skipped} skipped (see log).`
          : `✅ ${isReplace ? 'Replaced' : 'Created'} ${createdCount} slide(s) in PowerPoint!`;
      }
      if (loadingText) {
        loadingText.style.display = "none";
      }

      return slideStructures;
    } catch (err) {
      console.error("PPTAdapter Exception:", err);
      const errDetail = err.message || JSON.stringify(err);
      if (debugStatus) {
        debugStatus.innerHTML = `<span style="color:#d13438; font-weight:bold;">🔴 PPT Error: ${errDetail}</span>`;
      }
      if (loadingText) {
        loadingText.innerText = `🔴 Error: ${errDetail}`;
      }
      throw err;
    } finally {
      // Report from here so a build that threw is measured too - a slow failure
      // is exactly the case worth seeing.
      if (!timingsReported) {
        timingsReported = true;
        logBuildTimings(elapsedMs());
      }
      releaseSlideBuildLease(buildLeaseToken);
    }
  }

  // Scan in-slide / in-shape @gemini commands for PowerPoint
  async checkInDocumentCommands(forceRun = false, callbacks = {}) {
    if (callbacks.onStatus) callbacks.onStatus("PowerPoint Adapter Ready (@gemini in shape)");
  }
}
