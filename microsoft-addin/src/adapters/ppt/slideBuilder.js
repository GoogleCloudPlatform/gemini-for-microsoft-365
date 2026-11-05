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
 * PowerPoint Slide Builder
 *
 * Standard Office.js Slide Generation:
 * - Samples the OPEN presentation's theme (fonts, colors, canvas size) so generated
 *   slides visually inherit the customer's own template instead of a hard-coded look
 * - Creates each slide with the official pattern: slides.add() -> sync -> slides.getItemAt(count - 1)
 *   (SlideCollection.add() returns void, so the new slide MUST be resolved by index)
 * - Phased, fault-isolated rendering: the slide canvas is committed first, then core
 *   content, then purely cosmetic polish. A failure in a later phase can never discard
 *   an earlier one.
 * - Strictly isolates tables: when a slide has a table, only the native table is rendered
 * - Auto-fits type so content never overflows the slide canvas
 *
 * @author Sathya AG, Principal Architect, Google
 */

import { getDeckTheme, probeThemeColorScheme, tint, shade, readableTextOn, ADDIN_SLIDE_TAG, ADDIN_PALETTE_TAG, encodePalette } from "./deckTheme.js";
import { beginPhase, addPhase, recordSync, recordBackoff, nowMs, syncsSoFar, recordShapesDrawn } from "./buildTimings.js";
import { cardLayoutBudget, CARD_LAYOUT_SHARE_PERCENT, isCardLayoutSlide, demoteCardSlideForSlowHost } from "./cardBudget.js";

export function isSubstantiveSlideBody(text) {
  if (!text || typeof text !== "string") return false;
  const trimmed = text.trim();
  if (!trimmed || trimmed === "• Executive slide content") return false;
  // If text contains only backticks, punctuation, quotes, or markdown fence characters
  if (/^[`'"*#_~>|\-\s•]+$/.test(trimmed)) return false;

  // Clean markdown markers and bullets
  const cleaned = trimmed
    .replace(/[`*#_~>|\-•]/g, " ")
    .replace(/\b(?:here\s+is\s+|below\s+is\s+|sure|certainly|image\s+of|the\s+image\s+you\s+requested|executive\s+slide\s+content)\b/gi, " ")
    .trim();

  // Must have at least 15 characters and contain words with letters
  return /[a-zA-Z]{3,}/.test(cleaned) && cleaned.length >= 15;
}

/**
 * Writes to the task pane's Troubleshooting & Diagnostics panel.
 *
 * `level` is `false` for information, `"warn"` for something the host refused
 * but the builder recovered from, and `true` (or `"error"`) for content that
 * was actually lost. The distinction matters: this host refuses batches it
 * goes on to commit, so a wall of red for writes that all landed is worse than
 * no logging at all -- it hides the two or three lines that mean something.
 */
function logToPPTConsole(msg, level = false) {
  const isError = level === true || level === "error";
  const isWarn = level === "warn";
  const prefix = isError ? "❌ [PPT] " : isWarn ? "⚠️ [PPT] " : "ℹ️ [PPT] ";
  console.log(`${prefix}${msg}`);

  // The task pane's "Troubleshooting & Diagnostics" panel owns the visible log.
  // This used to target `debugConsole`, an element that does not exist in
  // taskpane.html, so every diagnostic written here was invisible unless the
  // user had a web inspector attached -- which made "I saw no errors"
  // impossible to act on. `debugConsole` is still honoured as a fallback.
  const box = document.getElementById("diagLogBox") || document.getElementById("debugConsole");
  if (!box) return;

  const time = new Date().toLocaleTimeString();
  const line = `[${time}] ${prefix}${msg}`;
  if (box.tagName === "PRE") {
    box.textContent = `${box.textContent}\n${line}`;
  } else {
    const entry = document.createElement("div");
    entry.style.color = isError ? "#a4262c" : isWarn ? "#8a6d00" : "#242424";
    entry.style.marginBottom = "2px";
    entry.innerText = line;
    box.appendChild(entry);
  }
  box.scrollTop = box.scrollHeight;
}

/** Yields to the host for `ms`. */
function delay(ms) {
  // Every call site is retry backoff or a deliberate yield to the host, so the
  // total is worth reporting apart from time spent doing actual work.
  recordBackoff(ms);
  return new Promise((resolve) => setTimeout(resolve, ms));
}

/* ------------------------------------------------------------------ */
/* Host capability detection                                           */
/* ------------------------------------------------------------------ */

const apiSupportCache = {};

/** True when the host supports at least the given PowerPointApi requirement set. */
function apiAtLeast(version) {
  if (apiSupportCache[version] !== undefined) return apiSupportCache[version];
  let supported = false;
  try {
    supported = Boolean(
      typeof Office !== "undefined" &&
      Office.context &&
      Office.context.requirements &&
      Office.context.requirements.isSetSupported("PowerPointApi", version)
    );
  } catch (_) {
    supported = false;
  }
  apiSupportCache[version] = supported;
  return supported;
}

/* ------------------------------------------------------------------ */
/* Layout geometry (derived from the real slide canvas)                */
/* ------------------------------------------------------------------ */

function getLayout(theme) {
  const W = theme?.slideWidth || 960;
  const H = theme?.slideHeight || 540;
  const margin = Math.round(W * 0.052);
  return {
    W,
    H,
    margin,
    contentWidth: W - margin * 2,
    titleTop: Math.round(H * 0.052),
    // Lowest y-coordinate body content may occupy.
    contentBottom: H - Math.round(H * 0.055)
  };
}

/* ------------------------------------------------------------------ */
/* Typography helpers                                                  */
/* ------------------------------------------------------------------ */

// applyFont and addText swallow every write failure, which is how a narrative
// box could come out in PowerPoint's default 18pt black while the title on the
// same slide carried the deck theme. Record what the host refused so the log
// can name the specific write instead of reporting a clean slide.
let TEXT_WRITE_REFUSALS = [];

function takeTextWriteRefusals() {
  const notes = TEXT_WRITE_REFUSALS;
  TEXT_WRITE_REFUSALS = [];
  return notes;
}

function noteTextWriteRefused(what, err) {
  // Capped: one pathological slide must not bury the rest of the log.
  if (TEXT_WRITE_REFUSALS.length < 12) {
    TEXT_WRITE_REFUSALS.push(`${what} (${(err && (err.message || err.code)) || String(err)})`);
  }
}

function applyFont(textRange, style = {}) {
  if (!textRange) {
    noteTextWriteRefused("textRange", "the shape exposed no text range");
    return;
  }
  let font = null;
  try {
    font = textRange.font;
  } catch (err) {
    noteTextWriteRefused("textRange.font", err);
    return;
  }
  if (!font) {
    noteTextWriteRefused("textRange.font", "the text range exposed no font");
    return;
  }
  try { if (style.size) font.size = style.size; } catch (err) { noteTextWriteRefused("font.size", err); }
  try { if (style.color) font.color = style.color; } catch (err) { noteTextWriteRefused("font.color", err); }
  try { if (style.bold !== undefined) font.bold = style.bold; } catch (err) { noteTextWriteRefused("font.bold", err); }
  try { if (style.italic !== undefined) font.italic = style.italic; } catch (err) { noteTextWriteRefused("font.italic", err); }
  // Only set the typeface when we actually sampled one from the deck; otherwise
  // leave it untouched so PowerPoint applies the presentation theme font.
  try { if (style.font) font.name = style.font; } catch (err) { noteTextWriteRefused("font.name", err); }
}

/* ------------------------------------------------------------------ */
/* Shape accounting                                                     */
/*                                                                      */
/* Answers the one question the host's shape count cannot: did anything */
/* this slide drew actually reach the document?                         */
/*                                                                      */
/* SHAPES_QUEUED counts shapes handed to the host but not yet           */
/* acknowledged. SHAPES_COMMITTED counts the ones a sync accepted.      */
/* Office.js discards a refused batch WHOLE, so a failed sync rolls the */
/* queue back to the committed mark - those shapes never existed.       */
/* SHAPES_ATTEMPTED never rolls back; it only records that we tried.    */
/*                                                                      */
/* Counting accepted SYNCS instead would not do. A slide whose every    */
/* write is refused still issues read-only syncs - shape counts, font   */
/* read-backs - that the host answers happily, and any one of them      */
/* would hide the failure.                                              */
/* ------------------------------------------------------------------ */

let SHAPES_QUEUED = 0;
let SHAPES_COMMITTED = 0;
let SHAPES_ATTEMPTED = 0;

/* SHAPES_DISCARDED counts shapes a refused sync threw away. SHAPES_REDRAWN
 * counts the ones a retry deliberately queued again afterwards.
 *
 * Without the pair, `landed=24/27` is unreadable: ATTEMPTED never rolls back,
 * so a header that was refused once and redrawn successfully is counted twice
 * and a complete slide looks like it lost three shapes - while a slide that
 * genuinely lost three reads exactly the same. discarded-minus-redrawn is the
 * number that tells those two apart. */
let SHAPES_DISCARDED = 0;
let SHAPES_REDRAWN = 0;

/** Records shapes re-queued after a refusal, so they are not read as losses. */
function noteShapesRedrawn(n) {
  SHAPES_REDRAWN += Math.max(0, Number(n) || 0);
}

/**
 * Runs a redraw and counts whatever it queued.
 *
 * Measuring the delta beats annotating each site with a number: the card
 * committer redraws a whole row through a closure it cannot count, and the
 * first version of this accounting declared eighteen shapes missing from a
 * slide that had just redrawn all eighteen of them.
 */
function redrawCounted(fn) {
  const before = SHAPES_ATTEMPTED;
  try {
    return fn();
  } finally {
    noteShapesRedrawn(SHAPES_ATTEMPTED - before);
  }
}

function noteShapeQueued() {
  SHAPES_QUEUED += 1;
  SHAPES_ATTEMPTED += 1;
  recordShapesDrawn(1);
}

/**
 * Adds a word-wrapped text box with the supplied theme styling.
 * `style.align` accepts a PowerPoint ParagraphHorizontalAlignment string and is
 * only applied on hosts that support requirement set 1.4.
 */
function addText(slide, text, bounds, style = {}) {
  const value = text === null || text === undefined ? "" : String(text);
  const box = slide.shapes.addTextBox(value, bounds);
  noteShapeQueued();
  try { box.textFrame.wordWrap = true; } catch (err) { noteTextWriteRefused("textFrame.wordWrap", err); }
  if (style.align && apiAtLeast("1.4")) {
    try {
      box.textFrame.textRange.paragraphFormat.horizontalAlignment = style.align;
    } catch (_) {}
  }
  applyFont(box.textFrame ? box.textFrame.textRange : null, style);
  return box;
}

/**
 * Commits the queued batch without letting a host rejection abort the caller.
 *
 * Office.js discards the WHOLE queued batch when a sync fails, so the builder
 * syncs in small units: one bad property write then costs a single card rather
 * than the entire slide.
 */
let SAFE_SYNC_OK = 0;

// Sync failures used to go only to console.warn, so a slide could fail every
// single write and still report nothing but "Created". The budget keeps one
// pathological slide from burying the rest of the log.
let SYNC_ERROR_BUDGET = 15;

function resetSyncDiagnostics() {
  SYNC_ERROR_BUDGET = 15;
}

/**
 * `context.sync()` with a stopwatch.
 *
 * The first timing run could only see the 18 safeSync call sites and still
 * attributed 49% of a 108-second build to host round trips, at 2.5 seconds
 * each. The other 38 sync sites were invisible. Measuring all of them is the
 * difference between knowing where the time goes and guessing.
 *
 * Rethrows on failure: callers that want a refusal swallowed use safeSync.
 */
/**
 * How long to wait for one host round trip before giving up on the build.
 *
 * The base, for a sync that carries no shapes. syncDeadlineFor() adds an
 * allowance per queued shape on top: a big card row on a slow host has been
 * measured needing ~20s on its own, so a flat 20s caught "slow", not "never".
 *
 * Not a constant only so the offline harness can shorten it; nothing in the
 * add-in changes it.
 */
let SYNC_TIMEOUT_MS = 20000;

function setSyncTimeoutForTests(ms) {
  SYNC_TIMEOUT_MS = ms;
}

/**
 * The deadline for one sync, sized to what it carries.
 *
 * A flat 20s was wrong in the other direction. The 09-28 11:57 build ran at
 * ~650ms per shape from start to finish, and slide 7's timeline committed 33
 * shapes in one "4 cards" batch - about 20s of honest work. The flat watchdog
 * called that a hang and abandoned slides 7-10 on a host that was still
 * answering. Cost on this host scales with shape count, so the allowance does
 * too: base + 1.5s per queued shape (about twice the slowest rate measured),
 * capped at 6x base so a real hang is still caught inside two minutes.
 *
 * Expressed as fractions of SYNC_TIMEOUT_MS so the offline harness, which
 * shortens the base, scales every part of this with it.
 */
function syncDeadlineFor(pendingShapes) {
  // The 09-29 14:50 build ran at ~3.2s per shape - 6x the usual - and a
  // 6-shape card batch got the fixed 29s (20s + 6 x 1.5s). It was still
  // committing when the watchdog stopped the deck at slide 6 of 10. The
  // allowance per shape is now 3x what this host has actually been taking,
  // whenever that is more than the fixed 1.5s. Same 6x cap, so a real hang is
  // still caught inside two minutes.
  const perShape = Math.max(Math.round(SYNC_TIMEOUT_MS * 0.075), Math.round(measuredMsPerShape() * 3));
  const sized = SYNC_TIMEOUT_MS + Math.max(0, pendingShapes) * perShape;
  return Math.min(SYNC_TIMEOUT_MS * 6, sized);
}

/**
 * Set by the first sync that never came back. Latched, never cleared mid-build.
 *
 * Once this is set every meteredSync throws before issuing anything. The
 * builder is full of catch blocks that survive a refusal - which is correct
 * for a refusal and wrong for a host that has stopped answering, because they
 * would let the build keep drawing into a void. Short-circuiting at the one
 * place every sync passes through means no catch site has to know about this.
 */
let BUILD_ABORTED = null;   // { label, ms }

/**
 * How long this host takes to commit one shape, learned from this build.
 *
 * The 09-28 12:12 build committed 21-shape rows in ~15s and then lost a
 * 29-shape row outright: no answer in 64s, and nothing from it ever landed.
 * The same 28-32 shape rows had committed on faster days. What kills a commit
 * is how long it runs, and that depends on the host's speed that day - so the
 * speed is measured rather than assumed.
 */
let MS_PER_SHAPE_EST = 0;
let ROW_BUDGET_ANNOUNCED = false;

/** Declared after the estimate so syncDeadlineFor can read it lint-clean. */
function measuredMsPerShape() {
  return MS_PER_SHAPE_EST || 0;
}

/**
 * Above this measured rate a card slide is drawn as styled bullets instead:
 * 2s per shape at the production 20s base (0.1x, so the harness scales). A
 * slim card row is 9+ shapes - 20s+ of commit at that speed - while the
 * bullet version is 1-2 shapes.
 */
function slowHostMsPerShape() {
  return SYNC_TIMEOUT_MS * 0.1;
}

function resetBuildAbort() {
  BUILD_ABORTED = null;
  MS_PER_SHAPE_EST = 0;
  ROW_BUDGET_ANNOUNCED = false;
}

function noteCommitRate(shapes, ms) {
  if (shapes < 3 || !(ms > 0)) return;
  const per = ms / shapes;
  MS_PER_SHAPE_EST = MS_PER_SHAPE_EST ? MS_PER_SHAPE_EST * 0.5 + per * 0.5 : per;
}

/**
 * Most shapes one card-row commit may carry: whatever fits in ~8s at the
 * measured rate (0.4x the base timeout, so the harness scales with it).
 * Clamped: never below 6 (about one card), never above 24 (a fast host keeps
 * whole rows in one round trip, as before). Unmeasured: 24, the pre-r47 row.
 */
function rowShapeBudget() {
  // Unmeasured only when a renderer runs outside a build: every real slide
  // commits its header first, which measures the host before any card row.
  if (!MS_PER_SHAPE_EST) return 24;
  const target = SYNC_TIMEOUT_MS * 0.4;
  return Math.max(6, Math.min(24, Math.floor(target / MS_PER_SHAPE_EST)));
}

function abortedWhere() {
  return BUILD_ABORTED && BUILD_ABORTED.label
    ? ` while committing "${BUILD_ABORTED.label}"`
    : "";
}

function buildAbortError() {
  const ms = BUILD_ABORTED ? BUILD_ABORTED.ms : SYNC_TIMEOUT_MS;
  const err = new Error(
    `PowerPoint stopped responding${abortedWhere()} (no answer in ${Math.round(ms / 1000)}s)`
  );
  // Read by safeSync and the slide loop. A refusal is survivable; this is not.
  err.isBuildAbort = true;
  return err;
}

async function meteredSync(ctx, label) {
  // The build is already over. Issuing more work would only queue shapes into
  // a context nobody is going to commit.
  if (BUILD_ABORTED) throw buildAbortError();

  const startedAt = nowMs();
  // Measured before the sync: this is what the host is being asked to commit,
  // and it is what the deadline has to cover.
  const pendingShapes = Math.max(0, SHAPES_QUEUED - SHAPES_COMMITTED);
  const deadlineMs = syncDeadlineFor(pendingShapes);
  // Whatever was queued is about to be committed - or refused and, on this
  // host, possibly committed anyway. Either way the cached count is stale.
  invalidateShapeCount();
  let timer = null;
  let heartbeat = null;
  try {
    // A long commit used to be invisible: nothing was logged between "Preparing"
    // and the verdict. Each line here also ticks the build lease, which goes
    // stale after 20s of silence - without it a second build could reclaim one
    // that is still working. Bounded by the deadline, so it cannot keep a real
    // hang alive past it.
    heartbeat = setInterval(() => {
      const waited = Math.round((nowMs() - startedAt) / 1000);
      logToPPTConsole(
        `PowerPoint is still committing${label ? ` "${label}"` : ""}` +
        `${pendingShapes > 0 ? ` (${pendingShapes} shape(s))` : ""} - ${waited}s so far, ` +
        `waiting up to ${Math.round(deadlineMs / 1000)}s.`,
        "warn"
      );
    }, Math.max(50, Math.round(SYNC_TIMEOUT_MS / 2)));

    await Promise.race([
      ctx.sync(),
      new Promise((_resolve, reject) => {
        timer = setTimeout(() => {
          BUILD_ABORTED = { label: label || "", ms: deadlineMs };
          reject(buildAbortError());
        }, deadlineMs);
      })
    ]);
    // Everything queued up to this point is now in the document.
    SHAPES_COMMITTED = SHAPES_QUEUED;
    noteCommitRate(pendingShapes, nowMs() - startedAt);
    recordSync(nowMs() - startedAt, true);
  } catch (err) {
    if (err && err.isBuildAbort) {
      // Deliberately leaves the shape counters alone. A refused batch is known
      // to be gone; a batch that never answered might still land. Recording a
      // discard here would be a guess dressed up as a measurement, and the
      // whole point of this accounting is that it never does that.
      recordSync(nowMs() - startedAt, false);
      throw err;
    }
    // Office.js discards a refused batch whole, so the shapes in it are gone.
    SHAPES_DISCARDED += SHAPES_QUEUED - SHAPES_COMMITTED;
    SHAPES_QUEUED = SHAPES_COMMITTED;
    recordSync(nowMs() - startedAt, false);
    throw err;
  } finally {
    if (timer) clearTimeout(timer);
    if (heartbeat) clearInterval(heartbeat);
  }
}

async function safeSync(context, label) {
  try {
    await meteredSync(context, label);
    // Tallied so a caller can ask whether anything it queued actually
    // committed. Renderers swallow their own sync results, which is how a card
    // layout that committed nothing could still report success.
    SAFE_SYNC_OK += 1;
    return true;
  } catch (err) {
    // Every safeSync call site is written to survive a refusal by carrying on.
    // That is exactly the wrong response to a host that has stopped answering,
    // so this one error is not swallowed.
    if (err && err.isBuildAbort) throw err;
    console.warn(`[PPTBuilder] Sync notice (${label}):`, err);
    if (SYNC_ERROR_BUDGET > 0) {
      SYNC_ERROR_BUDGET -= 1;
      const detail = (err && (err.message || err.code)) || String(err);
      // A warning, not an error: every safeSync call site is written to
      // survive a refusal, and this host refuses plenty of batches it then
      // commits anyway.
      logToPPTConsole(
        `host refused "${label}": ${detail}` +
        (SYNC_ERROR_BUDGET === 0 ? " (further sync errors suppressed)" : ""),
        "warn"
      );
    }
    // The refusals come in windows: once the host starts answering
    // "InvalidParam passed to GetItem(id)" it refuses everything for a moment
    // and then recovers. Firing the next call immediately just spends it
    // inside the same window.
    await delay(200);
    return false;
  }
}

/**
 * Commits a row of cards one at a time, remembering any the host refused so
 * they can be drawn again once the rest of the row is down.
 *
 * A single refused card leaves a hole: slide 9 of the user's deck lost
 * "timeline step 1" and the row began at step 2. Retrying the whole layout is
 * not an option - the cards that committed would be drawn twice - so the retry
 * belongs around one card. A refused batch is discarded whole, so the card
 * really is absent when it is queued again.
 *
 * Two immediate attempts are not always enough: the host's refusal windows run
 * for seconds (four, on slide 8 of the same log), and both attempts can land
 * inside one. Drawing the rest of the row takes about as long, so a card that
 * was refused at the start of it usually commits at the end - and a card drawn
 * late is invisible to the reader, where a missing one is not.
 */
function createCardCommitter(context, slide, redrawPrologue) {
  const missed = [];
  // Drawn into the batch but not yet committed. Queued until the batch reaches
  // rowShapeBudget(), then synced: one round trip per row on a fast host, a
  // few short ones on a slow host, where one long commit is simply dropped.
  const pending = [];
  // Whatever the renderer drew before its first card (the timeline rail) rides
  // in the first chunk only. A refusal of any later chunk must not redraw it.
  let prologueInFlight = true;

  async function commitOne(label, draw) {
    // Only ever reached after the row batch was refused, so both of these are
    // redraws of shapes the host already threw away.
    redrawCounted(draw);
    if (await safeSync(context, label)) return true;
    await delay(400);
    redrawCounted(draw);
    if (await safeSync(context, `${label} (retry)`)) return true;
    missed.push({ label: label, draw: draw });
    return false;
  }

  /**
   * Commits the queued row in one batch.
   *
   * Office.js discards a refused batch whole, which is what makes the per-card
   * fallback safe to run afterwards - the cards really are absent. But this
   * host has been seen to report a refusal for a batch it then kept, so before
   * redrawing anything we check whether the slide grew. Redrawing over cards
   * the host quietly committed is the worse failure of the two.
   */
  async function commitPending() {
    if (pending.length === 0) return;
    const row = pending.splice(0, pending.length);

    // Free when createSingleSlide counted the slide just before the renderer
    // ran, which is the normal case. Cold only on a retry after a refusal.
    let before = knownShapeCount(slide);
    if (!before.ok) before = await countSlideShapes(context, slide);

    const label = row.length === 1 ? row[0].label : `${row.length} cards`;
    if (await safeSync(context, label)) return;

    const after = await countSlideShapes(context, slide);
    if (before.ok && after.ok && after.count > before.count) {
      logToPPTConsole(
        `${label}: the host refused the batch but kept it - not redrawing.`,
        "warn"
      );
      return;
    }

    logToPPTConsole(`${label}: batch refused, drawing them one at a time.`, "warn");

    // Anything the renderer queued before its first card went down with the
    // batch. The per-card loop below only knows about cards, so whatever came
    // first has to be put back by the renderer that drew it.
    if (prologueInFlight && typeof redrawPrologue === "function") {
      try {
        redrawCounted(redrawPrologue);
      } catch (prologueErr) {
        console.warn("[PPTBuilder] Could not redraw the row prologue:", prologueErr);
      }
    }

    for (let i = 0; i < row.length; i++) {
      await commitOne(row[i].label, row[i].draw);
    }
  }

  async function commitChunk() {
    await commitPending();
    prologueInFlight = false;
  }

  return {
    async commit(label, draw) {
      draw();
      pending.push({ label: label, draw: draw });
      const budget = rowShapeBudget();
      if (SHAPES_QUEUED - SHAPES_COMMITTED >= budget) {
        if (!ROW_BUDGET_ANNOUNCED) {
          ROW_BUDGET_ANNOUNCED = true;
          logToPPTConsole(
            `Card rows commit in batches of up to ${budget} shapes` +
            (MS_PER_SHAPE_EST ? ` (host measured at ~${Math.round(MS_PER_SHAPE_EST)}ms per shape)` : "") +
            ` - one long commit is dropped by this host, several short ones are not.`
          );
        }
        await commitChunk();
      }
      return true;
    },
    async flush() {
      await commitChunk();
      for (let i = 0; i < missed.length; i++) {
        await delay(300);
        redrawCounted(missed[i].draw);
        const ok = await safeSync(context, `${missed[i].label} (late)`);
        logToPPTConsole(
          ok
            ? `${missed[i].label}: refused twice, drawn after the rest of the row.`
            : `${missed[i].label}: refused three times - this card is missing from the slide.`,
          ok ? "warn" : true
        );
      }
      missed.length = 0;
    }
  };
}

/**
 * Estimates how many wrapped lines `text` occupies in a `width`-point column at
 * `fontSize`. Word-aware, so long words are not silently split mid-token.
 */
function estimateWrappedLines(text, width, fontSize, padding = 12) {
  const usable = Math.max(16, width - padding);
  const charsPerLine = Math.max(4, Math.floor(usable / (fontSize * 0.52)));
  const words = String(text === null || text === undefined ? "" : text).split(/\s+/).filter(Boolean);
  if (words.length === 0) return 1;

  let lines = 1;
  let used = 0;
  for (const word of words) {
    const add = used === 0 ? word.length : word.length + 1;
    if (used > 0 && used + add > charsPerLine) {
      lines += 1;
      used = word.length;
    } else {
      used += add;
    }
  }
  return lines;
}

// PowerPoint insets text roughly 0.1 inch on the left and right of every text
// box, so that much of the declared width never holds glyphs. Measuring against
// the raw width is what let card text wrap one line further than planned and
// hang out of the bottom of the bubble.
const TEXT_INSET = 14;

// A single-spaced line box is a little taller than the point size. 1.34 leaves
// a small safety margin over PowerPoint's ~1.2 without wasting obvious space.
const LINE_FACTOR = 1.34;

/**
 * Vertical space a multi-paragraph block really needs, measured the way the
 * host lays it out: word-aware wrapping inside the usable (inset) width.
 */
function measureTextHeight(text, width, size, inset = TEXT_INSET, lineFactor = LINE_FACTOR) {
  const content = String(text || "");
  if (!content.trim()) return 0;
  const usable = Math.max(16, width - inset);
  let lines = 0;
  for (const para of content.split("\n")) {
    lines += para.trim() ? estimateWrappedLines(para, usable, size, 0) : 0.5;
  }
  return Math.ceil(lines * size * lineFactor);
}

/** Vertical space a multi-paragraph block needs at `size` inside `width` points. */
function textBlockHeight(text, width, size, lineFactor = LINE_FACTOR) {
  return measureTextHeight(text, width, size, TEXT_INSET, lineFactor);
}

/**
 * Estimates the largest font size at which `text` still fits inside the box.
 * Purely arithmetic (no host round-trip) so it is safe to call for every shape.
 *
 * This shares one measurement with textBlockHeight. When the two disagreed, the
 * layout pass reserved a box using one estimate and the type was sized with the
 * other, and the difference showed up as text overflowing its card.
 */
export function fitFontSize(text, width, height, { max = 18, min = 10, inset = TEXT_INSET } = {}) {
  const content = String(text || "");
  if (!content.trim()) return max;

  let size = max;
  while (size > min) {
    if (measureTextHeight(content, width, size, inset) <= height) break;
    size -= 0.5;
  }

  return Math.max(min, Math.round(size * 2) / 2);
}

/**
 * Trims `text` to the number of lines the box can actually show at `size`.
 *
 * Only used once shrinking has bottomed out. PowerPoint text boxes do not clip,
 * so the alternative is not "smaller text" but text drawn across whatever sits
 * below the card.
 */
function clampTextToBox(text, width, height, size, inset = TEXT_INSET, lineFactor = LINE_FACTOR) {
  const content = String(text || "");
  if (!content.trim()) return content;
  if (measureTextHeight(content, width, size, inset, lineFactor) <= height) return content;

  const usable = Math.max(16, width - inset);
  const maxLines = Math.max(1, Math.floor(height / (size * lineFactor)));

  const kept = [];
  let used = 0;

  for (const para of content.split("\n")) {
    if (used >= maxLines) break;

    if (!para.trim()) {
      used += 0.5;
      kept.push(para);
      continue;
    }

    const lines = estimateWrappedLines(para, usable, size, 0);
    if (used + lines <= maxLines) {
      kept.push(para);
      used += lines;
      continue;
    }

    // Partial paragraph. Drop trailing words and re-measure after each cut --
    // a character budget is not sufficient, because wrapping can push a long
    // token onto a line the budget never counted. That assumption is what put
    // the text outside the card in the first place.
    const remaining = Math.floor(maxLines - used);
    if (remaining >= 1) {
      const words = para.split(/\s+/).filter(Boolean);
      let take = words.length;
      while (take > 0) {
        const candidate = `${words.slice(0, take).join(" ")}\u2026`;
        if (estimateWrappedLines(candidate, usable, size, 0) <= remaining) {
          kept.push(candidate);
          break;
        }
        take -= 1;
      }
    }
    used = maxLines;
    break;
  }

  // Final guarantee: keep cutting until the measured height really is inside
  // the box. The per-paragraph pass above works in lines, and rounding between
  // lines and points can still leave it a line over.
  let out = kept.join("\n");
  let guard = 0;
  while (out.trim() && measureTextHeight(out, width, size, inset, lineFactor) > height && guard < 600) {
    guard += 1;
    const stripped = out.replace(/\u2026\s*$/, "").trimEnd();
    const cut = stripped.lastIndexOf(" ");
    if (cut > 0) {
      out = `${stripped.substring(0, cut).trimEnd()}\u2026`;
    } else if (stripped.length > 1) {
      out = `${stripped.substring(0, stripped.length - 1)}\u2026`;
    } else {
      out = "";
    }
  }

  return out;
}

/**
 * Returns the type size AND the text that genuinely fit `width` x `height`.
 * Shrinks to `min` first; trims only if that is still not enough.
 *
 * @returns {{size: number, text: string, clipped: boolean}}
 */
function fitTextToBox(text, width, height, { max = 13, min = 8.5 } = {}) {
  const content = String(text || "");
  const size = fitFontSize(content, width, height, { max, min });
  const fitted = clampTextToBox(content, width, height, size);
  return { size, text: fitted, clipped: fitted !== content };
}

/** Strips a leading bullet glyph from every line. */
function stripBulletGlyphs(text) {
  return String(text || "")
    .split("\n")
    .map(l => l.replace(/^\s*[•\-*]\s+/, "").trimEnd())
    .join("\n");
}

/**
 * Cosmetic pass: removes the literal "•" prefixes, bolds "Label:" lead-ins and
 * adds paragraph spacing. Runs in its own sync so a failure leaves the
 * committed literal-bullet text untouched (every line keeps its glyph).
 *
 * It used to rewrite the whole text and switch native bullets on. The host
 * keeps "\n" as line breaks inside ONE paragraph, so the slide got a single
 * bullet in front of the first item and none on the rest; the rewrite also
 * wiped every run style applied before it. Bullets are now switched off, and
 * the glyphs are cut in place, last first, so earlier offsets and styling stay
 * where they were.
 */
async function enhanceBulletsCosmetic(context, box, rawText, theme) {
  if (!box || !apiAtLeast("1.4")) return false;
  const lines = String(rawText || "").split("\n");
  const cuts = [];
  const leads = [];
  let offset = 0;
  for (const line of lines) {
    const m = line.match(/^\s*[\u2022\-*]\s+/);
    if (m) {
      cuts.push({ start: offset, length: m[0].length });
      const lead = styledBulletColonLead(line.slice(m[0].length));
      if (lead) leads.push({ start: offset + m[0].length, length: lead.lead.length + 1 });
    }
    offset += line.length + 1;
  }
  if (cuts.length === 0) return false;

  try {
    const textRange = box.textFrame.textRange;
    // Queued before the cuts, so these offsets are still the original ones.
    for (const l of leads) textRange.getSubstring(l.start, l.length).font.bold = true;
    for (let i = cuts.length - 1; i >= 0; i--) {
      textRange.getSubstring(cuts[i].start, cuts[i].length).text = "";
    }
    const pf = textRange.paragraphFormat;
    if (pf && pf.bulletFormat) {
      pf.bulletFormat.visible = false;
    }
    await meteredSync(context);
  } catch (bulletErr) {
    console.warn("[PPTBuilder] Bullet clean-up notice (keeping literal bullets):", bulletErr);
    return false;
  }

  // Paragraph spacing is a newer requirement set; attempt separately so hosts that
  // only support 1.4 still keep their native bullets.
  if (apiAtLeast("1.8")) {
    try {
      const pf = box.textFrame.textRange.paragraphFormat;
      if (pf) {
        pf.spaceAfter = 8;
        pf.lineSpacing = 1.15;
      }
      await meteredSync(context);
    } catch (spacingErr) {
      console.warn("[PPTBuilder] Paragraph spacing notice:", spacingErr);
    }
  }

  return true;
}

/* ------------------------------------------------------------------ */
/* Image helpers                                                       */
/* ------------------------------------------------------------------ */

/**
 * Compresses base64 images down to fit comfortably within PowerPoint memory limits.
 */
export function compressImageForPowerPoint(base64Str, maxWidth = 1600, maxHeight = 1120, meta = null) {
  return new Promise((resolve) => {
    let resolved = false;
    // A multi-megabyte image can take several seconds to decode the first time
    // the WebView sees it. The old 1500ms budget lost that race on a cold decode
    // and resolved with the RAW payload, which PowerPoint then refused - leaving
    // the user with an empty slide. Once the same image had been shown full size
    // in the zoom modal the decode was warm, the compressor won the race, and
    // the identical click worked. Give the decode room, and record whether it
    // actually ran so the caller can say so in the log.
    const timer = setTimeout(() => {
      if (!resolved) {
        resolved = true;
        if (meta) meta.compressed = false;
        resolve(base64Str ? base64Str.replace(/^data:image\/[^;]+;base64,/i, "").replace(/[\r\n\s]+/g, "").trim() : "");
      }
    }, 12000);

    try {
      const cleanRaw = base64Str ? base64Str.replace(/^data:image\/[^;]+;base64,/i, "").replace(/[\r\n\s]+/g, "").trim() : "";
      if (!cleanRaw) {
        clearTimeout(timer);
        resolve("");
        return;
      }
      const img = new Image();
      img.onload = () => {
        if (resolved) return;
        resolved = true;
        clearTimeout(timer);
        let w = img.width || 800;
        let h = img.height || 600;
        // Record the native geometry so the caller can place the picture without
        // distorting it (charts are 900x560; photos are arbitrary).
        if (meta) {
          meta.width = w;
          meta.height = h;
          meta.aspect = h > 0 ? w / h : null;
        }
        if (w > maxWidth || h > maxHeight) {
          const ratio = Math.min(maxWidth / w, maxHeight / h);
          w = Math.round(w * ratio);
          h = Math.round(h * ratio);
        }
        const canvas = document.createElement("canvas");
        canvas.width = w;
        canvas.height = h;
        const ctx = canvas.getContext("2d");
        ctx.clearRect(0, 0, w, h);
        ctx.drawImage(img, 0, 0, w, h);
        const dataUrl = canvas.toDataURL("image/png");
        const compressedBase64 = dataUrl.replace(/^data:image\/[^;]+;base64,/, "");
        if (meta) {
          meta.compressed = true;
          meta.rawBytes = Math.round(cleanRaw.length * 0.75);
          meta.bytes = Math.round(compressedBase64.length * 0.75);
        }
        resolve(compressedBase64);
      };
      img.onerror = () => {
        if (resolved) return;
        resolved = true;
        clearTimeout(timer);
        if (meta) meta.compressed = false;
        resolve(cleanRaw);
      };
      img.src = base64Str.startsWith("data:") ? base64Str : `data:image/png;base64,${cleanRaw}`;
      // decode() resolves once the bitmap is genuinely ready. onload alone can
      // fire before a large PNG is decodable on some WebView builds, and when it
      // does not fire at all the timer above is the only way out.
      if (typeof img.decode === "function") {
        img.decode().then(() => {
          if (!resolved && typeof img.onload === "function") img.onload();
        }).catch(() => {});
      }
    } catch (e) {
      if (!resolved) {
        resolved = true;
        clearTimeout(timer);
        resolve(base64Str ? base64Str.replace(/^data:image\/[^;]+;base64,/i, "").replace(/[\r\n\s]+/g, "").trim() : "");
      }
    }
  });
}

// Ceiling for the Common API picture fallback. Comfortably longer than a
// healthy insert, short enough that a hang never outlives the build lease.
const COMMON_API_INSERT_TIMEOUT_MS = 15000;

/**
 * Inserts a picture via Office Common API setSelectedDataAsync (works universally across platforms).
 */
export async function insertPictureViaCommonApi(rawImg, { left, top, width, height }) {
  if (typeof Office === "undefined" || !Office.context?.document?.setSelectedDataAsync) {
    return false;
  }
  const clean = rawImg.replace(/^data:image\/[^;]+;base64,/i, "").replace(/[\r\n\s]+/g, "").trim();
  if (clean.length <= 50) return false;

  // setSelectedDataAsync does not reliably call back. With no valid selection
  // to insert into, or when handed a multi-megabyte payload, it can simply
  // never return. This promise used to have no way out, so one hang left
  // insertContent() pending forever and wedged the whole add-in: the build flag
  // never cleared and every later insert reported "a slide build is already in
  // progress". It must always settle.
  return new Promise((resolve) => {
    let settled = false;
    let timer = null;

    const finish = (value, note) => {
      if (settled) return;
      settled = true;
      if (timer) clearTimeout(timer);
      if (note) logToPPTConsole(note);
      resolve(value);
    };

    timer = setTimeout(() => {
      finish(false, "The Common API picture insert never responded (15s) - abandoning it so the build can finish.");
    }, COMMON_API_INSERT_TIMEOUT_MS);

    try {
      Office.context.document.setSelectedDataAsync(
        clean,
        {
          coercionType: Office.CoercionType.Image,
          imageLeft: left,
          imageTop: top,
          imageWidth: width,
          imageHeight: height
        },
        (asyncResult) => {
          if (asyncResult && asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            finish(true, `Added picture via Office Common API setSelectedDataAsync.`);
          } else {
            console.warn("[PPTBuilder] setSelectedDataAsync notice:", asyncResult ? asyncResult.error : null);
            finish(false);
          }
        }
      );
    } catch (err) {
      console.warn("[PPTBuilder] setSelectedDataAsync threw:", err);
      finish(false);
    }
  });
}

/**
 * Inserts a picture onto a slide using official Office.js PowerPoint APIs.
 * 1. Geometric shape with fill.setImage(clean) (PowerPointApi 1.8+ official method)
 * 2. shapes.addPicture(clean) (PowerPointApi 1.10+ / preview)
 */
function fitBoxToAspect(bounds, aspect) {
  const box = {
    left: Math.round(bounds.left),
    top: Math.round(bounds.top),
    width: Math.round(bounds.width),
    height: Math.round(bounds.height)
  };
  if (!aspect || !Number.isFinite(aspect) || aspect <= 0 || box.width <= 0 || box.height <= 0) {
    return box;
  }

  let width = box.width;
  let height = box.height;
  if (box.width / box.height > aspect) {
    width = Math.round(box.height * aspect);
  } else {
    height = Math.round(box.width / aspect);
  }

  return {
    left: Math.round(box.left + (box.width - width) / 2),
    top: Math.round(box.top + (box.height - height) / 2),
    width,
    height
  };
}

export async function insertPictureOnSlide(slide, rawImg, bounds, slideNum = 1, aspect = null, context = null) {
  if (!rawImg || !slide) return false;
  const clean = rawImg.replace(/^data:image\/[^;]+;base64,/i, "").replace(/[\r\n\s]+/g, "").trim();
  if (clean.length <= 50) return false;

  const sizeKb = Math.round((clean.length * 0.75) / 1024);

  // Letterbox inside the requested region so charts keep their native
  // proportions instead of being squashed to fill the box.
  const { left, top, width, height } = fitBoxToAspect(bounds, aspect);

  // A picture is only inserted once the HOST has accepted the batch. This used
  // to return true the moment setImage() was queued, so a rejected sync - which
  // is what an oversized payload produces - discarded the picture while the
  // caller recorded success. That suppressed both the non-blank guard and the
  // Common-API fallback, and the user got an empty slide.
  const commit = async (label) => {
    if (!context || typeof context.sync !== "function") return true;
    try {
      await meteredSync(context);
      return true;
    } catch (syncErr) {
      console.warn(`[PPTBuilder] Slide ${slideNum} picture sync refused (${label}):`, syncErr);
      logToPPTConsole(
        `Slide ${slideNum}: the host refused the picture via ${label} (${sizeKb} KB) - ${syncErr.message || syncErr}`
      );
      return false;
    }
  };

  // 1. Primary: Geometric Shape with picture fill (PowerPointApi 1.8+ official method)
  try {
    if (apiAtLeast("1.8") && typeof slide.shapes.addGeometricShape === "function" && typeof PowerPoint !== "undefined" && PowerPoint.GeometricShapeType) {
      const rect = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
        left,
        top,
        width,
        height
      });
      noteShapeQueued();
      if (rect && rect.fill && typeof rect.fill.setImage === "function") {
        try {
          if (rect.lineFormat) {
            rect.lineFormat.visible = false;
          }
        } catch (_) {}
        rect.fill.setImage(clean);
        if (await commit("shape.fill.setImage")) {
          logToPPTConsole(
            `Slide ${slideNum}: Added picture via shape.fill.setImage (${sizeKb} KB at ${width}x${height}pt).`
          );
          return true;
        }
      }
    }
  } catch (fillErr) {
    console.warn("[PPTBuilder] Shape fill.setImage attempt notice:", fillErr);
  }

  // 2. Secondary: shapes.addPicture (PowerPointApi 1.10+ / preview)
  try {
    if (typeof slide.shapes.addPicture === "function") {
      const pic = slide.shapes.addPicture(clean, { left, top, width, height });
      noteShapeQueued();
      if (pic && await commit("shapes.addPicture")) {
        logToPPTConsole(
          `Slide ${slideNum}: Added picture via shapes.addPicture (${sizeKb} KB at ${width}x${height}pt).`
        );
        return true;
      }
    }
  } catch (e) {
    console.warn("[PPTBuilder] shapes.addPicture notice:", e);
  }

  // Return false so caller safely falls back to insertPictureViaCommonApi
  logToPPTConsole(`Slide ${slideNum}: could not place the ${sizeKb} KB picture - trying again.`);
  return false;
}

/**
 * Re-encodes a picture as a smaller JPEG (longest side <= maxDim). Returns the
 * new data URL, or null when there is no canvas, it fails, or it is not
 * actually smaller. Oversized payloads are the usual reason a host refuses a
 * picture batch, so this is the second attempt before any fallback.
 */
export function shrinkImagePayload(rawImg, maxDim = 1600, quality = 0.85) {
  return new Promise((resolve) => {
    if (!rawImg || typeof document === "undefined" || typeof Image === "undefined") {
      resolve(null);
      return;
    }
    let done = false;
    let timer = null;
    const finish = (value) => {
      if (done) return;
      done = true;
      if (timer) clearTimeout(timer);
      resolve(value);
    };
    timer = setTimeout(() => finish(null), 5000);
    try {
      const img = new Image();
      img.onload = () => {
        try {
          const w = img.naturalWidth || img.width;
          const h = img.naturalHeight || img.height;
          if (!w || !h) { finish(null); return; }
          const scale = Math.min(1, maxDim / Math.max(w, h));
          const canvas = document.createElement("canvas");
          canvas.width = Math.max(1, Math.round(w * scale));
          canvas.height = Math.max(1, Math.round(h * scale));
          const g = canvas.getContext && canvas.getContext("2d");
          if (!g) { finish(null); return; }
          g.fillStyle = "#FFFFFF";
          g.fillRect(0, 0, canvas.width, canvas.height);
          g.drawImage(img, 0, 0, canvas.width, canvas.height);
          const out = canvas.toDataURL("image/jpeg", quality);
          finish(out && out.length < rawImg.length ? out : null);
        } catch (_) {
          finish(null);
        }
      };
      img.onerror = () => finish(null);
      img.src = /^data:image\//i.test(rawImg) ? rawImg : `data:image/png;base64,${rawImg}`;
    } catch (_) {
      finish(null);
    }
  });
}

/**
 * Places a picture on the slide at `slideIndex` in a context of its own.
 * Used after the build context refused it: a new context is the first moment a
 * write to that slide can succeed again.
 */
async function retryPictureOnSlideIndex(slideIndex, rawImg, bounds, slideNum, aspect) {
  if (slideIndex < 0 || !rawImg) return false;
  let placed = false;
  try {
    await PowerPoint.run(async (ctx) => {
      const slide = ctx.presentation.slides.getItemAt(slideIndex);
      placed = await insertPictureOnSlide(slide, rawImg, bounds, slideNum, aspect, ctx);
    });
  } catch (err) {
    console.warn(`[PPTBuilder] Slide ${slideNum} picture retry notice:`, err);
    placed = false;
  }
  return placed;
}

/**
 * Selects the slide at `slideIndex` and confirms the host actually moved the
 * selection there. setSelectedDataAsync inserts into whatever is selected, so
 * an unconfirmed selection is how a picture meant for a new slide ended up on
 * the user's current one.
 *
 * Returns "verified", "unverifiable" (host has no getSelectedSlides) or
 * "failed".
 */
async function selectSlideForCommonApi(slideIndex, slideNum) {
  if (slideIndex < 0) return "failed";
  let outcome = "failed";
  try {
    await PowerPoint.run(async (ctx) => {
      const pres = ctx.presentation;
      if (typeof pres.setSelectedSlides !== "function") return;
      const target = pres.slides.getItemAt(slideIndex);
      target.load("id");
      await meteredSync(ctx);
      const targetId = target.id;
      if (!targetId) return;

      pres.setSelectedSlides([targetId]);
      await meteredSync(ctx);

      if (typeof pres.getSelectedSlides !== "function") {
        outcome = "unverifiable";
        return;
      }
      for (let attempt = 0; attempt < 4; attempt++) {
        const selected = pres.getSelectedSlides();
        selected.load("items/id");
        await meteredSync(ctx);
        const items = selected.items || [];
        if (items.length === 1 && items[0].id === targetId) {
          outcome = "verified";
          return;
        }
        await delay(250 * (attempt + 1));
        if (attempt === 1) {
          pres.setSelectedSlides([targetId]);
          await meteredSync(ctx);
        }
      }
    });
  } catch (err) {
    console.warn(`[PPTBuilder] Slide ${slideNum} selection notice:`, err);
  }
  return outcome;
}

/* ------------------------------------------------------------------ */
/* Theme-aware surfaces                                                */
/* ------------------------------------------------------------------ */

function resolveShapeType(preferred) {
  try {
    if (typeof PowerPoint !== "undefined" && PowerPoint.GeometricShapeType) {
      const t = PowerPoint.GeometricShapeType[preferred];
      if (t) return t;
      if (PowerPoint.GeometricShapeType.rectangle) return PowerPoint.GeometricShapeType.rectangle;
    }
  } catch (_) {}
  return null;
}

/**
 * Draws a filled panel (rounded when the host supports it). Returns the shape or null.
 */
function drawPanel(slide, bounds, fillColor, borderColor = null, preferred = "roundRectangle") {
  try {
    if (typeof slide.shapes.addGeometricShape !== "function") return null;
    const shapeType = resolveShapeType(preferred);
    if (!shapeType) return null;

    const panel = slide.shapes.addGeometricShape(shapeType, bounds);
    if (!panel) return null;
    noteShapeQueued();

    try {
      if (panel.fill && typeof panel.fill.setSolidColor === "function" && fillColor) {
        panel.fill.setSolidColor(fillColor);
      }
    } catch (_) {}

    // lineFormat arrived in PowerPointApi 1.4; skip entirely on older hosts so we
    // never poison the batch with an unsupported property write.
    if (apiAtLeast("1.4")) {
      try {
        if (panel.lineFormat) {
          if (borderColor) {
            panel.lineFormat.color = borderColor;
            panel.lineFormat.visible = true;
            panel.lineFormat.weight = 1;
          } else {
            panel.lineFormat.visible = false;
          }
        }
      } catch (_) {}
    }

    return panel;
  } catch (panelErr) {
    console.warn("[PPTBuilder] Panel shape creation notice:", panelErr);
    return null;
  }
}

/**
 * Draws a card: an offset soft shadow behind, then the card surface on top.
 */
function drawCard(slide, bounds, theme, { fill, border, shadow = true, shape = "roundRectangle" } = {}) {
  if (shadow) {
    drawPanel(
      slide,
      { left: bounds.left + 3, top: bounds.top + 4, width: bounds.width, height: bounds.height },
      theme.shadowColor,
      null,
      shape
    );
  }
  return drawPanel(slide, bounds, fill || theme.cardFill, border || theme.cardBorder, shape);
}

/** Thin accent rule used to anchor the slide title to the deck's brand color. */
function drawAccentRule(slide, layout, theme, top) {
  drawPanel(
    slide,
    { left: layout.margin, top, width: Math.round(layout.contentWidth * 0.075), height: 4 },
    theme.accentColor,
    null,
    "rectangle"
  );
}

/* ------------------------------------------------------------------ */
/* Native tables                                                       */
/* ------------------------------------------------------------------ */

/**
 * Distributes the available width across columns in proportion to their content
 * volume, with the dynamic range compressed so a single prose column cannot
 * starve the short numeric ones.
 */
function computeColumnWidths(values, totalWidth, colCount) {
  const demand = new Array(colCount).fill(3);
  for (const row of values) {
    for (let c = 0; c < colCount; c++) {
      const len = String(row[c] === null || row[c] === undefined ? "" : row[c]).length;
      if (len > demand[c]) demand[c] = len;
    }
  }

  const weights = demand.map(d => Math.pow(d, 0.62));
  const weightSum = weights.reduce((a, b) => a + b, 0) || 1;
  const minWidth = Math.max(46, Math.floor(totalWidth / (colCount * 2.8)));

  const widths = weights.map(w => Math.max(minWidth, Math.round((w / weightSum) * totalWidth)));

  // Normalize back to exactly totalWidth by trimming the widest column.
  let drift = widths.reduce((a, b) => a + b, 0) - totalWidth;
  while (drift > 0) {
    const widestIdx = widths.indexOf(Math.max(...widths));
    const take = Math.min(drift, widths[widestIdx] - minWidth);
    if (take <= 0) break;
    widths[widestIdx] -= take;
    drift -= take;
  }
  if (drift < 0) {
    const widestIdx = widths.indexOf(Math.max(...widths));
    widths[widestIdx] -= drift;
  }

  return widths;
}

/**
 * Picks the largest font size at which the entire table still fits `maxHeight`,
 * measuring real wrapped-line counts per cell. If even the minimum size
 * overflows, cell text is clipped so the table can never run off the slide.
 */
function planTableLayout(values, colWidths, maxHeight, maxFont = 12.5, minFont = 7.5) {
  const rowPad = 7;

  for (let size = maxFont; size >= minFont; size -= 0.5) {
    const lineHeight = size * 1.26;
    const rowHeights = [];
    let total = 0;
    for (const row of values) {
      let lines = 1;
      for (let c = 0; c < colWidths.length; c++) {
        const l = estimateWrappedLines(row[c], colWidths[c], size);
        if (l > lines) lines = l;
      }
      const h = Math.max(size * 1.85, lines * lineHeight + rowPad);
      rowHeights.push(h);
      total += h;
    }
    if (total <= maxHeight) {
      return { fontSize: size, rowHeights, totalHeight: Math.round(total), values, clipped: false };
    }
  }

  // Minimum size still overflows: budget an equal share of the height to every
  // row and clip the cells that do not fit their share.
  const size = minFont;
  const lineHeight = size * 1.26;
  const rowCount = values.length || 1;
  const budgetPerRow = Math.max(size * 1.85, maxHeight / rowCount);
  const maxLines = Math.max(1, Math.floor((budgetPerRow - rowPad) / lineHeight));

  const clippedValues = values.map(row => row.map((cell, c) => {
    const str = String(cell === null || cell === undefined ? "" : cell);
    const usable = Math.max(16, colWidths[c] - 12);
    const charsPerLine = Math.max(4, Math.floor(usable / (size * 0.52)));
    const budget = charsPerLine * maxLines;
    return str.length > budget ? `${str.substring(0, Math.max(1, budget - 1)).trimEnd()}\u2026` : str;
  }));

  const rowHeights = clippedValues.map(row => {
    let lines = 1;
    for (let c = 0; c < colWidths.length; c++) {
      const l = Math.min(maxLines, estimateWrappedLines(row[c], colWidths[c], size));
      if (l > lines) lines = l;
    }
    return Math.max(size * 1.85, lines * lineHeight + rowPad);
  });

  const total = rowHeights.reduce((a, b) => a + b, 0);
  return {
    fontSize: size,
    rowHeights,
    totalHeight: Math.round(Math.min(total, maxHeight)),
    values: clippedValues,
    clipped: true
  };
}

/**
 * Populates a native Microsoft PowerPoint table, sized and typeset so it always
 * fits inside the region it was given.
 *
 * PowerPoint grows rows to fit wrapped text regardless of the height requested
 * on the shape, so the only reliable way to keep a table on the slide is to
 * plan the type size, column widths and row heights up front.
 */
/* ------------------------------------------------------------------ */
/* Shape-count cache                                                   */
/*                                                                     */
/* A count is only reusable while nothing can have changed the slide,  */
/* and nothing can change a slide without a sync. So every sync clears */
/* this - see meteredSync - which makes a cache hit provably equal to  */
/* the number a fresh read would return. A miss just costs the read.   */
/*                                                                     */
/* This exists so the card batch can ask "did the row land after all?" */
/* without spending a round trip to establish the baseline.            */
/* ------------------------------------------------------------------ */

let countedSlide = null;
let countedShapes = -1;

function invalidateShapeCount() {
  countedSlide = null;
  countedShapes = -1;
}

/** The slide's shape count if it is known for free, else {ok:false}. */
function knownShapeCount(slide) {
  if (slide && countedSlide === slide && countedShapes >= 0) {
    return { ok: true, count: countedShapes };
  }
  return { ok: false, count: -1 };
}

/**
 * Reads how many shapes a slide currently holds.
 *
 * `ShapeCollection.load("items/...")` is not dependable across PowerPoint
 * hosts: some acknowledge the load and still hand back an empty `items` array.
 * Acting on that produced slides with the body text stamped over the cards, so
 * this helper uses `getCount()` and reports whether the answer is usable at
 * all rather than defaulting to a number the caller might believe.
 *
 * @returns {Promise<{ok: boolean, count: number}>}
 */
async function countSlideShapes(context, slide) {
  try {
    // Re-resolving the slide first was tried (r37) on the theory that deleting
    // through a throwaway proxy breaks the caller's collection. It does not:
    // the next run read through a fresh proxy on all eight slides and six still
    // answered 0 against landed counts of 4 to 27. getCount() on this host is
    // unreliable full stop, so callers must treat it as advisory.
    if (slide && slide.shapes && typeof slide.shapes.getCount === "function") {
      const res = slide.shapes.getCount();
      await meteredSync(context);
      if (res && typeof res.value === "number" && res.value >= 0) {
        countedSlide = slide;
        countedShapes = res.value;
        return { ok: true, count: res.value };
      }
    }
  } catch (countErr) {
    console.warn("[PPTBuilder] Shape count unavailable:", countErr);
  }
  return { ok: false, count: -1 };
}

/**
 * Commits a block of text, retrying without styling if the styled write is
 * refused. A slide carrying plain bullets beats a slide carrying nothing.
 *
 * The retry is guarded by a shape count: a host can report a failed sync and
 * still have committed the shapes, and blindly retrying in that case is what
 * produced the same text box twice.
 *
 * @returns {Promise<object|null>} the committed shape, or null.
 */
async function commitTextBlock(context, slide, text, bounds, style, label) {
  const before = await countSlideShapes(context, slide);

  let box = null;
  try {
    box = addText(slide, text, bounds, style);
  } catch (queueErr) {
    console.warn(`[PPTBuilder] ${label} could not be queued:`, queueErr);
  }

  if (box && await safeSync(context, label)) return box;

  // Did it land despite the rejection? If so, leave well alone. Redrawing on
  // top of a shape the host quietly kept is what produced the duplicated text
  // boxes the user saw.
  const after = await countSlideShapes(context, slide);
  if (before.ok && after.ok && after.count > before.count) return null;

  // Rung 2: same styling, fresh batch. A rejection is often nothing to do with
  // the style, and dropping straight to unstyled would leave the slide in
  // default Calibri while every other slide carries the deck theme.
  try {
    const restyled = redrawCounted(() => addText(slide, text, bounds, style));
    if (restyled && await safeSync(context, `${label} (styled retry)`)) return restyled;
    const afterRestyled = await countSlideShapes(context, slide);
    if (before.ok && afterRestyled.ok && afterRestyled.count > before.count) return null;
  } catch (restyleErr) {
    console.warn(`[PPTBuilder] ${label} styled retry failed:`, restyleErr);
  }

  // Rung 3: the style is what the host objects to. Unstyled text on the slide
  // beats a correctly themed slide with nothing on it.
  try {
    const plain = redrawCounted(() => addText(slide, text, bounds, {}));
    if (plain && await safeSync(context, `${label} (unstyled retry)`)) return plain;
  } catch (plainErr) {
    console.warn(`[PPTBuilder] ${label} unstyled retry failed:`, plainErr);
  }
  return null;
}

/** Flattens a table to text so its data survives even when no shape could be drawn. */
function tableToText(tableData) {
  if (!tableData) return "";
  const lines = [];
  if (tableData.headers && tableData.headers.length) {
    lines.push(tableData.headers.join("  |  "));
  }
  for (const row of (tableData.rows || [])) {
    lines.push((row || []).join("  |  "));
  }
  return lines.join("\n");
}

// Bumped whenever the builder ships. Logged on every generation so the task
// pane console identifies the bundle in use -- "is the fix actually loaded?"
// is otherwise unanswerable from a screenshot.
export const BUILDER_REVISION = "r59-deck-verified-sources";

// Below this many points there is no room for a readable narrative block.
const MIN_NARRATIVE_HEIGHT = 38;

/**
 * How long this host needs after addTable before anything about the table can
 * be read back. Measured from the read-back trace: answers arrived 1.0-2.3s
 * after creation, and no first attempt ever succeeded.
 */
const TABLE_SETTLE_MS = 900;

/**
 * Backoff between read-back attempts, in ms. Was 250/500, which expired well
 * before the host was ready and then spent a third attempt - plus a full proxy
 * refresh - discovering that. With the 900ms settle above the total budget is
 * 3.3s, comfortably past the 2.3s worst case the trace measured.
 */
const TABLE_READBACK_BACKOFF = [500, 800, 1100];
const TABLE_READBACK_ATTEMPTS = TABLE_READBACK_BACKOFF.length + 1;

/**
 * Applies content-weighted column widths to a table that already exists.
 *
 * Sizing after creation (PowerPointApi 1.9) rather than through
 * `TableAddOptions` is deliberate: a rejected add option takes the whole table
 * with it and invites a retry, and retries are what stacked three copies of the
 * same table on one slide. A rejection here costs only the widths.
 *
 * @returns {Promise<boolean>} true when the widths were accepted.
 */
async function applyTableColumnWidths(context, table, widths) {
  if (!apiAtLeast("1.9")) return false;
  try {
    if (!table.columns || typeof table.columns.getItemAt !== "function") return false;
    for (let c = 0; c < widths.length; c++) {
      const column = table.columns.getItemAt(c);
      if (!column) return false;
      column.width = Math.round(widths[c]);
    }
  } catch (queueErr) {
    console.warn("[PPTBuilder] Column width write could not be queued:", queueErr);
    return false;
  }
  return await safeSync(context, "table: column widths");
}

/**
 * Applies planned row heights to a table that already exists. Best effort:
 * PowerPoint grows a row that is too short for its text regardless.
 */
async function applyTableRowHeights(context, table, heights) {
  if (!apiAtLeast("1.9")) return false;
  try {
    if (!table.rows || typeof table.rows.getItemAt !== "function") return false;
    for (let r = 0; r < heights.length; r++) {
      const row = table.rows.getItemAt(r);
      if (!row) return false;
      row.height = Math.round(heights[r]);
    }
  } catch (queueErr) {
    console.warn("[PPTBuilder] Row height write could not be queued:", queueErr);
    return false;
  }
  return await safeSync(context, "table: row heights");
}

/**
 * True when the slide carries an EXPLICIT card layout: a `visualType` +
 * `visualData` payload the model emitted on purpose, and that the archetype
 * dispatch below knows how to draw.
 *
 * `structuredCards` is deliberately excluded. Those are inferred from the body
 * by `extractStructuredBullets()`, and on a table-only slide the table's own
 * rows satisfy the pattern -- so treating them as a card layout would turn
 * every reference table into bubbles. Only a payload the model wrote counts.
 */
function hasExplicitCardLayout(slideData) {
  if (!slideData || !slideData.visualData) return false;
  const vt = slideData.visualType;
  const vd = slideData.visualData;
  if (vt === "metric_grid_3col" || vt === "three_column_cards") {
    return Array.isArray(vd.cards) && vd.cards.length > 0;
  }
  if (vt === "process_timeline") {
    return Array.isArray(vd.steps) && vd.steps.length > 0;
  }
  if (vt === "before_after") {
    const hasB = vd.before && Array.isArray(vd.before.bullets) && vd.before.bullets.length > 0;
    const hasA = vd.after && Array.isArray(vd.after.bullets) && vd.after.bullets.length > 0;
    return Boolean(hasB || hasA);
  }
  if (vt === "executive_summary") {
    return Boolean(vd.summary && String(vd.summary).trim());
  }
  return false;
}

/** Resolves the Table object behind a table shape across host variations. */
function tableFromShape(shape) {
  if (!shape) return null;
  try {
    if (typeof shape.getTable === "function") return shape.getTable();
  } catch (_) {}
  try {
    return shape.table || null;
  } catch (_) {}
  return null;
}

/* ------------------------------------------------------------------ */
/* Deferred table finishing                                            */
/* ------------------------------------------------------------------ */

/**
 * Tables whose verification and styling were postponed to the end of the build.
 *
 * Everything after `addTable` -- reading the values back, re-resolving the
 * shape proxy, applying column widths, row heights and header/zebra fills --
 * needs the host to have finished committing the table, and on this host that
 * takes anywhere from 0.9s to more than 7s. Waiting for it inline produced two
 * bad outcomes at once:
 *
 *   Slow:  ~5.7s of `await delay()` per unlucky table with the host idle
 *          (900ms settle + 2.4s verify backoff + 2.4s repair backoff).
 *   Ugly:  when the wait ran out first, the styling pass inherited a proxy the
 *          host would not resolve and every write was refused, so the table
 *          shipped in PowerPoint's default style -- wrong columns, no header
 *          fill, no zebra. That hit roughly one table in five: run #2 slide 2
 *          and run #4 slide 7 both ended `proxy as created` with widths,
 *          heights and fills all refused.
 *
 * Both problems only ever needed time, and the rest of the build supplies it
 * for free. By the time the last slide is drawn the host has had tens of
 * seconds and resolves the table on the first try.
 *
 * The text is deliberately NOT deferred -- it goes in through
 * `specificCellProperties` at creation, atomically. So a table whose finish
 * pass never runs is still a populated native table, which is the same "keep
 * it" outcome the inline path reached on this host anyway, minus the wait.
 */
let deferredTables = [];

function resetDeferredTables() {
  deferredTables = [];
}

/**
 * Records the deck position of the slide a pending table was drawn on.
 *
 * The finish pass runs in a fresh PowerPoint.run where the slide proxy from the
 * build context is dead, so the index is the only handle back to it. The caller
 * only learns it from takeAppendedSlideIndex() after createSingleSlide returns,
 * which is why this is a second step rather than part of the record.
 */
function noteDeferredTableSlideIndex(slideNum, slideIndex) {
  try {
    if (!(slideIndex >= 0)) return;
    for (let i = 0; i < deferredTables.length; i++) {
      if (deferredTables[i].slideNum === slideNum && deferredTables[i].slideIndex < 0) {
        deferredTables[i].slideIndex = slideIndex;
      }
    }
  } catch (_) {}
}

/** Resolves a cell through whichever accessor the host exposes. */
function tableCellAt(table, r, c) {
  if (!table) return null;
  if (typeof table.getCell === "function") return table.getCell(r, c);
  if (typeof table.getCellOrNullObject === "function") return table.getCellOrNullObject(r, c);
  return null;
}

/**
 * Queues font and fill for every populated cell.
 *
 * Shared by the inline and deferred paths on purpose: two copies would drift,
 * and the difference would surface as a table styled differently depending on
 * how fast the host happened to be that run.
 */
function queueTableCellStyling(table, spec) {
  try {
    for (let r = 0; r < spec.rowCount; r++) {
      const isHeaderRow = r === 0 && spec.hasHeader;
      for (let c = 0; c < spec.colCount; c++) {
        if (!spec.cellTextAt(r, c)) continue;
        const cell = tableCellAt(table, r, c);
        if (!cell) continue;
        try {
          if (cell.font) {
            cell.font.name = spec.bodyFont;
            cell.font.size = isHeaderRow ? spec.headerSize : spec.fontSize;
            cell.font.bold = isHeaderRow || c === 0;
            cell.font.color = isHeaderRow ? spec.headerText : spec.bodyText;
          }
        } catch (_) {}
        try {
          if (cell.fill && typeof cell.fill.setSolidColor === "function") {
            if (isHeaderRow) {
              cell.fill.setSolidColor(spec.headerFill);
            } else if ((r - (spec.hasHeader ? 1 : 0)) % 2 === 0) {
              cell.fill.setSolidColor(spec.zebraFill);
            }
          }
        } catch (_) {}
      }
    }
  } catch (styleErr) {
    console.warn("[PPTBuilder] Table styling could not be queued:", styleErr);
  }
}

/**
 * Reads a table's own copy of its text back. "populated" | "empty" | "unknown".
 *
 * A non-empty answer is proof. An exception or silence only means the question
 * could not be asked, which is emphatically not a "no".
 */
async function readDeferredTableState(context, table, label) {
  try {
    if (typeof table.load !== "function") return "unknown";
    table.load("values");
    if (!(await safeSync(context, label))) return "unknown";
    const values = table.values;
    if (!values || !values.length) return "unknown";
    for (const row of values) {
      if (!row) continue;
      for (const v of row) {
        if (String(v === null || v === undefined ? "" : v).trim()) return "populated";
      }
    }
    return "empty";
  } catch (_) {
    return "unknown";
  }
}

/**
 * Finds the table shape on a slide from a cold context.
 *
 * Two strategies, because hosts disagree about which one answers. Listing
 * `shapes.items` costs one round trip for the whole slide and is preferred, but
 * this host routinely returns an empty list for a slide that demonstrably has
 * shapes -- the same behaviour that makes the placeholder sweep report "the
 * host listed no shapes" -- so the indexed walk has to exist as well.
 *
 * The walk loads every candidate's type in a single batched sync rather than
 * one sync per shape: a card slide carries ~27 shapes, and asking 27 times
 * would cost more than the styling it is trying to enable.
 */
async function resolveTableOnSlide(context, slide, rec, label) {
  // Strategy A: by the id captured in the create batch. This is the route the
  // production host answers -- run #4 re-resolved slides 6 and 9 this way -- and
  // it is one round trip.
  try {
    if (rec && rec.tableShapeId && typeof slide.shapes.getItem === "function") {
      const shape = slide.shapes.getItem(rec.tableShapeId);
      const cand = tableFromShape(shape);
      if (cand) {
        shape.load("id");
        if (await safeSync(context, `${label}: by id`)) {
          return { table: cand, shape, via: "shape id" };
        }
      }
    }
  } catch (_) {}

  try {
    if (typeof slide.shapes.load === "function") {
      slide.shapes.load("items/id, items/type");
      if (await safeSync(context, `${label}: list`)) {
        const items = slide.shapes.items || [];
        for (let i = items.length - 1; i >= 0; i--) {
          const cand = tableFromShape(items[i]);
          if (cand) return { table: cand, shape: items[i], via: "the shape list" };
        }
      }
    }
  } catch (_) {}

  try {
    if (typeof slide.shapes.getCount !== "function" ||
        typeof slide.shapes.getItemAt !== "function") {
      return null;
    }
    const countRes = slide.shapes.getCount();
    if (!(await safeSync(context, `${label}: count`))) return null;
    const count = countRes && typeof countRes.value === "number" ? countRes.value : 0;
    if (count <= 0) return null;

    const proxies = [];
    for (let i = 0; i < count; i++) {
      const shape = slide.shapes.getItemAt(i);
      if (!shape) continue;
      try {
        if (typeof shape.load === "function") shape.load("type");
      } catch (_) {}
      proxies.push(shape);
    }
    if (!proxies.length) return null;
    if (!(await safeSync(context, `${label}: types`))) return null;

    // Newest first: the table is added before the picture and the narrative, so
    // it is not reliably the last shape, but it is nearer the end than the
    // header is.
    for (let i = proxies.length - 1; i >= 0; i--) {
      const cand = tableFromShape(proxies[i]);
      if (cand) return { table: cand, shape: proxies[i], via: "an indexed walk" };
    }
  } catch (_) {}

  return null;
}

/**
 * Backoff between deferred finish attempts, in ms.
 *
 * Deliberately shorter than the inline cascade it replaces: by the time this
 * runs the host has usually had the whole build to settle, so these are for the
 * short-deck case rather than the common one.
 */
const DEFERRED_FINISH_BACKOFF = [500, 900, 1500];

/**
 * Runs the verify-repair-style sequence against an already-resolved table.
 *
 * @returns {Promise<boolean>} true when there is nothing left to try: the table
 *   was fully styled, or it was confirmed empty and replaced with text.
 */
async function finishTableInContext(context, slide, rec, table, tableShape, via) {
  let state = await readDeferredTableState(context, table, `table finish ${rec.slideNum}: values`);

  // Repair on anything that is not provably populated, which is what the
  // inline path always did.
  //
  // Tempting to repair only on a definite "empty" -- on this host "unknown"
  // means the read-back lagged, and seven traces running show the write had
  // landed all along, so the repair is wasted. But a host that exposes no
  // `table.load`, no `table.values` and no per-cell `load` can never answer
  // anything BUT "unknown", and skipping the repair there ships a table with
  // empty cells. The write is idempotent -- the same text
  // specificCellProperties already carried -- so being wrong costs one sync,
  // while being right the other way costs the whole table.
  if (state !== "populated") {
    let wrote = false;
    try {
      for (let r = 0; r < rec.spec.rowCount; r++) {
        for (let c = 0; c < rec.spec.colCount; c++) {
          const strVal = rec.spec.cellTextAt(r, c);
          if (!strVal) continue;
          const cell = tableCellAt(table, r, c);
          if (!cell) continue;
          try {
            cell.text = strVal;
            wrote = true;
          } catch (_) {}
        }
      }
    } catch (textErr) {
      console.warn("[PPTBuilder] Deferred cell text could not be queued:", textErr);
    }
    if (wrote) {
      await safeSync(context, `table finish ${rec.slideNum}: repair`);
      state = await readDeferredTableState(context, table, `table finish ${rec.slideNum}: re-read`);
    }
  }

  // Confirmed empty even after a repair: remove it rather than leave an empty
  // grid with the text fallback showing through from behind, which is exactly
  // the artefact an earlier build produced.
  if (state === "empty") {
    try {
      if (tableShape && typeof tableShape.delete === "function") {
        tableShape.delete();
        await safeSync(context, `table finish ${rec.slideNum}: remove empty`);
      }
    } catch (delErr) {
      console.warn("[PPTBuilder] Empty table could not be removed:", delErr);
    }
    try {
      const fallbackText = rec.fallbackValues.map(r => r.join("  |  ")).join("\n");
      addText(slide, fallbackText, {
        left: rec.tableLeft,
        top: rec.tableTop,
        width: rec.tableWidth,
        height: rec.usedHeight
      }, {
        font: rec.spec.bodyFont,
        size: fitFontSize(fallbackText, rec.tableWidth, rec.usedHeight, { max: 13, min: 8 }),
        color: rec.spec.bodyText
      });
      await safeSync(context, `table finish ${rec.slideNum}: text fallback`);
    } catch (fErr) {
      console.warn("[PPTBuilder] Deferred table fallback failed:", fErr);
    }
    logToPPTConsole(
      `Slide ${rec.slideNum}: the table read back empty even after a repair, so it was ` +
      `replaced with text.`,
      true
    );
    return true;
  }

  const widthsApplied = await applyTableColumnWidths(context, table, rec.colWidths);
  const heightsApplied = await applyTableRowHeights(
    context, table, widthsApplied ? rec.sizedRowHeights : rec.equalRowHeights
  );
  queueTableCellStyling(table, rec.spec);
  const styled = await safeSync(context, `table finish ${rec.slideNum}: styling`);

  if (!widthsApplied || !heightsApplied || !styled) return false;

  logToPPTConsole(
    `Slide ${rec.slideNum}: table finished - ${state}, content-weighted columns and ` +
    `header/zebra styling applied (via ${via}).`
  );
  return true;
}

/**
 * Verifies and styles one deferred table.
 *
 * Two attempts, each in a context of its own:
 *
 *   1. The proxy `addTable` handed back. Invalid across a PowerPoint.run
 *      boundary on a real host, but it is the ONLY handle on hosts that expose
 *      neither shapes.getItem nor shapes.getItemAt and return an empty items
 *      list -- which is most of them.
 *   2. Cold resolution by id, then by listing, then by an indexed walk. This is
 *      the route that will actually run in production, where attempt 1 is dead.
 *
 * Separate contexts matter: this host refuses every subsequent sync in a
 * context once one object path has gone bad, so trying a rejected proxy in the
 * same context as the fallback would cost the fallback too -- the same failure
 * that made the batched placeholder sweep silently drop cleanups.
 */
async function finishOneDeferredTable(rec) {
  if (rec.table && rec.tableShape) {
    try {
      const done = await PowerPoint.run(async (context) => {
        const slide = rec.slideIndex >= 0
          ? context.presentation.slides.getItemAt(rec.slideIndex)
          : null;
        return await finishTableInContext(
          context, slide, rec, rec.table, rec.tableShape, "the original proxy"
        );
      });
      if (done) return true;
    } catch (_) {
      // Expected on a real host: the proxy belongs to a context that has closed.
    }
  }

  if (!(rec.slideIndex >= 0)) return false;

  // A ten-slide deck gives the host tens of seconds to settle before this runs,
  // and one attempt is plenty. A three-slide deck does not: the last table can
  // be created seconds before the build ends, putting this pass right back
  // inside the settle window. Without a retry the table loses its column widths
  // permanently, which is the defect this whole change set out to fix.
  //
  // Fresh context per attempt, because this host refuses every later sync in a
  // context once one object path has gone bad.
  let resolvedEver = false;
  for (let attempt = 0; attempt <= DEFERRED_FINISH_BACKOFF.length; attempt++) {
    if (attempt > 0) await delay(DEFERRED_FINISH_BACKOFF[attempt - 1]);

    let done = false;
    try {
      done = await PowerPoint.run(async (context) => {
        const slide = context.presentation.slides.getItemAt(rec.slideIndex);
        const found = await resolveTableOnSlide(
          context, slide, rec, `table finish ${rec.slideNum}`
        );
        if (!found) return false;
        resolvedEver = true;
        return await finishTableInContext(
          context, slide, rec, found.table, found.shape, found.via
        );
      });
    } catch (err) {
      console.warn(`[PPTBuilder] Deferred finish attempt failed on slide ${rec.slideNum}:`, err);
    }
    if (done) return true;
  }

  logToPPTConsole(
    resolvedEver
      ? `Slide ${rec.slideNum}: the host would not style the table even after the deck was ` +
        `finished - it is there and populated, but in PowerPoint's default style.`
      : `Slide ${rec.slideNum}: the table could not be found when finishing the deck, so it ` +
        `keeps PowerPoint's default style.`,
    "warn"
  );
  return false;
}

/**
 * Verifies and styles every deferred table once the deck is built.
 *
 * Best effort by construction: a table that cannot be resolved is left exactly
 * as `addTable` created it -- present and populated, just unstyled -- which is
 * strictly no worse than the inline path's own failure mode, and is reported
 * rather than hidden.
 */
async function finishDeferredTables() {
  const pending = deferredTables.slice();
  deferredTables = [];
  let styled = 0;

  for (const rec of pending) {
    if (!(rec.slideIndex >= 0)) {
      logToPPTConsole(
        `Slide ${rec.slideNum}: the host never said where the slide landed, so the table ` +
        `keeps PowerPoint's default style.`,
        "warn"
      );
      continue;
    }
    try {
      if (await finishOneDeferredTable(rec)) styled++;
    } catch (err) {
      console.warn(`[PPTBuilder] Deferred table finish failed on slide ${rec.slideNum}:`, err);
      logToPPTConsole(
        `Slide ${rec.slideNum}: finishing the table failed outright, so it keeps ` +
        `PowerPoint's default style.`,
        "warn"
      );
    }
  }
  return { total: pending.length, styled };
}

async function populateSlideTable(context, newSlide, tableData, slideNum, theme, bounds, opts) {
  const deferFinish = Boolean(opts && opts.deferFinish);
  const { left: tableLeft, top: tableTop, width: tableWidth, maxHeight } = bounds;
  const headers = tableData.headers || [];
  const rows = tableData.rows || [];
  const colCount = Math.max(headers.length, ...rows.map(r => r.length), 1);
  const hasHeader = headers.length > 0;

  const rawValues = [];
  if (hasHeader) {
    rawValues.push(Array.from({ length: colCount }, (_, c) =>
      (headers[c] === null || headers[c] === undefined ? "" : String(headers[c]))));
  }
  for (const r of rows) {
    rawValues.push(Array.from({ length: colCount }, (_, c) =>
      (r[c] === null || r[c] === undefined ? "" : String(r[c]))));
  }
  if (rawValues.length === 0) return 0;

  // Two layout plans, because the table may or may not get the column widths
  // it asks for. `sizedPlan` assumes content-weighted widths; `equalPlan`
  // assumes the equal widths PowerPoint applies when it refuses the `columns`
  // option. Planning only the first would let a prose column wrap far past its
  // budgeted height and push the table off the slide.
  const colWidths = computeColumnWidths(rawValues, tableWidth, colCount);
  const sizedPlan = planTableLayout(rawValues, colWidths, maxHeight);

  const equalWidths = Array.from({ length: colCount }, () => tableWidth / colCount);
  const equalPlan = planTableLayout(rawValues, equalWidths, maxHeight);

  const rowCount = rawValues.length;

  const headerFill = theme.accentColor;
  const headerText = readableTextOn(headerFill);
  const zebraFill = theme.isDark ? shade(theme.accentColor, 0.78) : tint(theme.accentColor, 0.94);
  const bodyText = theme.isDark ? "#E5E7EB" : "#1F2937";

  // ONE addTable, ever -- CREATED ALREADY POPULATED, then verified against the
  // host's own copy of the data.
  //
  // Two earlier designs failed here. Inferring success from the sync result
  // produced three stacked tables, because this host reports a rejection for a
  // batch it actually commits. Inferring it from the shape count then produced
  // an empty table, because a table that exists but holds no text satisfies a
  // count just as well as a good one does.
  //
  // So the text is no longer written into the table after the fact: it is
  // handed to `addTable` as `specificCellProperties`, which makes creation and
  // population a single atomic operation that cannot half-happen. Verification
  // reads `table.values` -- a read-only property the host computes. A cell
  // proxy cannot serve here, because assigning `cell.text = x` and reading it
  // back returns x whether or not the host ever accepted the write.
  const plan = equalPlan;            // width-agnostic: correct whether or not
  const tableValues = plan.values;   // the requested column widths are honoured
  const fontSize = plan.fontSize;
  const headerSize = fontSize;
  let usedHeight = Math.min(maxHeight, plan.totalHeight);
  let table = null;
  let tableShape = null;
  let tableVerified = false;
  let widthsApplied = false;
  let createSynced = false;
  // Declared up here because the create batch below captures it, which is ~40
  // lines before refreshTableProxy - the function that used to own it - is
  // defined.
  let knownTableShapeId = null;

  const cellAt = (r, c) => {
    if (typeof table.getCell === "function") return table.getCell(r, c);
    if (typeof table.getCellOrNullObject === "function") return table.getCellOrNullObject(r, c);
    return null;
  };

  const cellTextAt = (r, c) => {
    const v = tableValues[r] ? tableValues[r][c] : "";
    return String(v === null || v === undefined ? "" : v).trim();
  };

  // Built once and handed to both the inline and the deferred styling pass, so
  // a table cannot end up styled differently depending on which path ran.
  const stylingSpec = {
    rowCount,
    colCount,
    hasHeader,
    cellTextAt,
    bodyFont: theme.bodyFont,
    fontSize,
    headerSize,
    headerFill,
    headerText,
    zebraFill,
    bodyText
  };

  // Formatting travels with the text, so no separate styling pass is needed on
  // the atomic path.
  const buildCellProperties = () => {
    const grid = [];
    for (let r = 0; r < rowCount; r++) {
      const isHeaderRow = r === 0 && hasHeader;
      const row = [];
      for (let c = 0; c < colCount; c++) {
        const props = {
          text: cellTextAt(r, c),
          font: {
            name: theme.bodyFont,
            size: isHeaderRow ? headerSize : fontSize,
            bold: isHeaderRow || c === 0,
            color: isHeaderRow ? headerText : bodyText
          },
          verticalAlignment: "Middle"
        };
        if (isHeaderRow) {
          props.fill = { color: headerFill };
        } else if ((r - (hasHeader ? 1 : 0)) % 2 === 0) {
          props.fill = { color: zebraFill };
        }
        row.push(props);
      }
      grid.push(row);
    }
    return grid;
  };

  // "populated" | "empty" | "unknown". A non-empty answer is proof. An
  // A short record of how the verify/repair dance actually went on this host:
  // which channel could answer, how many attempts it took, and whether the
  // repair pass populated the table or merely coincided with the host becoming
  // readable. Those two have completely different fixes, and the log could not
  // previously tell them apart.
  const tableTrace = [];
  const tableStartedSyncs = syncsSoFar();
  const tableStartedAt = nowMs();
  let lastChannel = "none";

  // exception only means the question could not be asked right now, which is
  // emphatically not the same as a "no".
  const readBackState = async () => {
    lastChannel = "none";
    if (!table) return "unknown";

    // Channel 1: `values`, computed by the host over the whole grid.
    try {
      if (typeof table.load === "function") {
        table.load("values");
        await meteredSync(context);
        const values = table.values;
        if (values && values.length) {
          lastChannel = "values";
          for (const row of values) {
            if (!row) continue;
            for (const v of row) {
              if (String(v === null || v === undefined ? "" : v).trim()) return "populated";
            }
          }
          return "empty";
        }
      }
    } catch (_) {}

    // Channel 2: read a few cells back individually. Weaker, because a cell
    // proxy echoes back whatever was assigned to it -- but these are *fresh*
    // proxies that were never written through, so a blank answer really does
    // come from the host. That is the case that matters: without this channel a
    // host that hides `values` would force every table into "unknown", and an
    // empty one would be kept.
    try {
      const probes = [];
      for (let r = 0; r < rowCount && probes.length < 3; r++) {
        for (let c = 0; c < colCount && probes.length < 3; c++) {
          if (!cellTextAt(r, c)) continue;
          const cell = cellAt(r, c);
          if (cell && typeof cell.load === "function") {
            cell.load("text");
            probes.push(cell);
          }
        }
      }
      if (probes.length) {
        await meteredSync(context);
        let readable = false;
        for (const probe of probes) {
          const t = probe.text;
          if (t === null || t === undefined) continue;
          readable = true;
          if (String(t).trim()) {
            lastChannel = "cells";
            return "populated";
          }
        }
        if (readable) {
          lastChannel = "cells";
          return "empty";
        }
      }
    } catch (_) {}

    return "unknown";
  };

  // Snapshot for the blind case below: a host that refuses addTable still
  // returns a Table object, so "we got an object" is not proof of a shape.
  const shapeCountBefore = await countSlideShapes(context, newSlide);

  if (typeof newSlide.shapes.addTable === "function") {
    // `width` is undocumented for tables (the typings derive width from the
    // column widths) but this host honours it, and the requested column
    // widths are not always accepted -- so it stays.
    const baseOptions = { left: tableLeft, top: tableTop, width: tableWidth, height: usedHeight };
    let queued = false;
    try {
      // Deliberately no `columns` here. The widths are applied after creation
      // through the 1.9 column API, and `columns` is the option most likely to
      // be refused -- a host that refuses it outright would cost us the native
      // table for no gain.
      tableShape = newSlide.shapes.addTable(rowCount, colCount, Object.assign({}, baseOptions, {
        specificCellProperties: buildCellProperties()
      }));
      queued = true;
      // One shape, not one per cell: the host draws a table as a single
      // object and the per-cell writes ride along inside it.
      noteShapeQueued();
    } catch (queueErr) {
      // A synchronous throw means nothing reached the queue, so falling back to
      // a plain table here cannot duplicate anything. Only a *sync-time*
      // rejection is ambiguous, and that is never retried.
      console.warn("[PPTBuilder] Populated addTable could not be queued:", queueErr);
      try {
        tableShape = newSlide.shapes.addTable(rowCount, colCount, baseOptions);
        queued = true;
        noteShapeQueued();
      } catch (plainErr) {
        console.warn("[PPTBuilder] Plain addTable could not be queued either:", plainErr);
      }
    }
    if (queued) {
      table = tableFromShape(tableShape);
      // A clean sync proves the shape landed. A failed one proves nothing at
      // all on this host, which commits rejected batches -- so this is only
      // ever read as positive evidence, never as negative.
      // Ask for the id in the same batch that creates the table: it costs no
      // extra round trip, and it is the handle the deferred finish pass needs
      // to re-resolve the shape from a cold context later. Reading it on its
      // own after the fact would land inside the refusal window and fail --
      // which is exactly what happened to slide 7.
      try {
        if (typeof tableShape.load === "function") tableShape.load("id");
      } catch (_) {}
      createSynced = await safeSync(context, "table: create");
      try {
        if (tableShape.id !== null && tableShape.id !== undefined && String(tableShape.id)) {
          knownTableShapeId = String(tableShape.id);
        }
      } catch (_) {}
    }
  }

  // Everything below needs the host to have finished committing the table, and
  // that is the wait that costs 0.9-7s and loses the styling when it runs out.
  // Hand it to the end of the build instead, where the time is already spent.
  if (deferFinish && table && tableShape) {
    deferredTables.push({
      slideNum,
      slideIndex: -1,              // filled in by noteDeferredTableSlideIndex
      // The live proxy, tried first. A real host invalidates it across a
      // PowerPoint.run boundary, but hosts that expose neither getItem nor
      // getItemAt have no other handle at all, and it costs nothing to try.
      table,
      tableShape,
      tableShapeId: knownTableShapeId,
      colWidths,
      sizedRowHeights: sizedPlan.rowHeights,
      equalRowHeights: plan.rowHeights,
      spec: stylingSpec,
      fallbackValues: equalPlan.values,
      tableLeft,
      tableTop,
      tableWidth,
      usedHeight
    });
    return usedHeight;
  }

  // The Table handed back by `addTable` is resolved against a shape the host
  // has not necessarily committed yet, and on this host that proxy goes stale:
  // every later cell, column and row call answers "InvalidParam passed to
  // GetItem(id)", so the table keeps PowerPoint's equal columns and none of the
  // theme styling. Re-resolving it from the slide by shape id is the same move
  // that stopped slides being written onto their predecessor -- and it can be
  // run again later, once the host has had a moment, which is what the retries
  // below rely on. Entirely best-effort: on any refusal the proxy is kept.
  let proxyRefreshed = false;
  const refreshTableProxy = async () => {
    if (!newSlide || !newSlide.shapes) return false;

    // Strategy A: resolve by shape ID if known or readable
    try {
      if (!knownTableShapeId && tableShape && typeof tableShape.load === "function") {
        tableShape.load("id");
        await meteredSync(context);
        if (tableShape.id !== null && tableShape.id !== undefined && String(tableShape.id)) {
          knownTableShapeId = String(tableShape.id);
        }
      }
      if (knownTableShapeId && typeof newSlide.shapes.getItem === "function") {
        const freshShape = newSlide.shapes.getItem(knownTableShapeId);
        freshShape.load("id");
        await meteredSync(context);
        if (freshShape.id !== null && freshShape.id !== undefined) {
          const freshTable = tableFromShape(freshShape);
          if (freshTable) {
            tableShape = freshShape;
            table = freshTable;
            proxyRefreshed = true;
            return true;
          }
        }
      }
    } catch (_) {}

    // Strategy B: scan shapes.items backward for a table shape
    try {
      if (typeof newSlide.shapes.load === "function") {
        newSlide.shapes.load("items/id, items/type");
        await meteredSync(context);
        const items = newSlide.shapes.items || [];
        for (let i = items.length - 1; i >= 0; i--) {
          const cand = items[i];
          const candTable = tableFromShape(cand);
          if (candTable) {
            tableShape = cand;
            table = candTable;
            if (cand.id) knownTableShapeId = String(cand.id);
            proxyRefreshed = true;
            return true;
          }
        }
      }
    } catch (_) {}

    // Strategy C: resolve the last shape on the slide by index (getCount - 1)
    try {
      if (typeof newSlide.shapes.getCount === "function" && typeof newSlide.shapes.getItemAt === "function") {
        const cRes = newSlide.shapes.getCount();
        await meteredSync(context);
        const cnt = cRes && typeof cRes.value === "number" ? cRes.value : 0;
        if (cnt > 0) {
          const cand = newSlide.shapes.getItemAt(cnt - 1);
          if (cand && typeof cand.load === "function") {
            cand.load("id");
            await meteredSync(context);
          }
          const candTable = tableFromShape(cand);
          if (candTable) {
            tableShape = cand;
            table = candTable;
            if (cand && cand.id) knownTableShapeId = String(cand.id);
            proxyRefreshed = true;
            return true;
          }
        }
      }
    } catch (refreshErr) {
      console.warn("[PPTBuilder] Table proxy could not be re-resolved:", refreshErr);
    }
    return false;
  };

  // Nothing about this table is resolvable for a beat after it is created:
  // not `values`, not the cells, not even the shape by id. The read-back trace
  // measured it at 1.0-2.3s across five tables, with the first attempt failing
  // 5 times out of 5.
  //
  // The loop below was therefore not verifying anything on its first pass, it
  // was waiting - at ~330ms of round trip per question asked. Waiting off the
  // wire is free by comparison. Everything after this is unchanged, so a host
  // that needs longer still gets the same retries and the same fallbacks.
  await delay(TABLE_SETTLE_MS);

  await refreshTableProxy();

  if (table) {
    // Retried: this host rejects a short burst of syncs immediately after
    // addTable and then recovers, so a lone attempt can land inside the burst
    // and answer "unknown" for a perfectly good table.
    let state = "unknown";
    for (let attempt = 1; attempt <= TABLE_READBACK_ATTEMPTS; attempt++) {
      state = await readBackState();
      tableTrace.push(`v${attempt}:${state}/${lastChannel}`);
      if (state !== "unknown") break;
      if (attempt === TABLE_READBACK_ATTEMPTS) break;
      // Backing off matters more than retrying: several immediate attempts all
      // land inside the same window in which the host cannot resolve the
      // table, and all of them answer "unknown" for a perfectly good table.
      await delay(TABLE_READBACK_BACKOFF[attempt - 1] || 1100);
      // Deliberately NOT refreshing after the first failure. Re-resolving the
      // shape costs 2-5 round trips and is refused by the very window that
      // just refused the read, so it cannot help yet - another wait can. Only
      // from the second failure is a genuinely dead proxy the better
      // explanation.
      if (attempt >= 2) await refreshTableProxy();
    }

    // Repair pass, for a host that accepts addTable but ignores
    // specificCellProperties. One pass only, then re-verify.
    if (state !== "populated") {
      let wrote = false;
      try {
        for (let r = 0; r < rowCount; r++) {
          for (let c = 0; c < colCount; c++) {
            const strVal = cellTextAt(r, c);
            if (!strVal) continue;
            const cell = cellAt(r, c);
            if (!cell) continue;
            try {
              cell.text = strVal;
              wrote = true;
            } catch (_) {}
          }
        }
      } catch (textErr) {
        console.warn("[PPTBuilder] Table cell text could not be queued:", textErr);
      }
      if (wrote) {
        tableTrace.push("repair:wrote");
        await safeSync(context, "table: cell text repair");
        for (let attempt = 1; attempt <= TABLE_READBACK_ATTEMPTS; attempt++) {
          state = await readBackState();
          tableTrace.push(`r${attempt}:${state}/${lastChannel}`);
          if (state !== "unknown") break;
          if (attempt === TABLE_READBACK_ATTEMPTS) break;
          await delay(TABLE_READBACK_BACKOFF[attempt - 1] || 1100);
          if (attempt >= 2) await refreshTableProxy();
        }
        if (state === "populated") {
          logToPPTConsole(`Slide ${slideNum}: table text needed a second pass, but the table is populated.`);
        }
      }
    }

    logToPPTConsole(
      `Slide ${slideNum}: table read-back - ${tableTrace.join(" ") || "none"} ` +
      `-> ${state} | ${syncsSoFar() - tableStartedSyncs} syncs, ` +
      `${((nowMs() - tableStartedAt) / 1000).toFixed(2)}s | proxy ` +
      `${proxyRefreshed ? "re-resolved" : "as created"}.`
    );

    if (state === "empty") {
      // Confirmed empty. Remove it, rather than leave an empty grid with the
      // text fallback showing through from behind -- precisely the artefact an
      // earlier build produced.
      try {
        if (tableShape && typeof tableShape.delete === "function") {
          tableShape.delete();
          await safeSync(context, "table: remove empty");
        }
      } catch (delErr) {
        console.warn("[PPTBuilder] Empty table could not be removed:", delErr);
      }
      table = null;
      tableShape = null;
      logToPPTConsole(
        `Slide ${slideNum}: the table read back empty, so it was removed before falling back to text.`,
        true
      );
    } else if (state === "populated") {
      tableVerified = true;
    } else {
      // Unreadable by either channel. Decide whether the shape actually landed
      // before abandoning it -- but weight the evidence honestly. A host that
      // refuses addTable outright still hands back a Table object, so the
      // question is worth asking; on the other hand THIS host commits batches
      // it reports as rejected, so any single negative signal is more likely to
      // be a lie than a genuine absence. A stale `getCount()` alone once cost
      // two good tables, which were then covered with pipe-delimited text.
      //
      // Hence: three independent signals, any one of which is proof the shape
      // is there, and the fallback only when all three come back empty.
      let landedBy = null;
      // Set only when the host actually ANSWERED the question and the answer
      // was "no new shape". Silence is not an answer -- and the host is silent
      // precisely when it is refusing everything, which is also precisely when
      // the table is most likely to have landed anyway.
      let countedAbsent = false;

      if (createSynced) landedBy = "create sync";

      if (!landedBy) {
        for (let attempt = 1; attempt <= 3 && !landedBy; attempt++) {
          const after = await countSlideShapes(context, newSlide);
          if (after.ok && shapeCountBefore.ok) {
            if (after.count > shapeCountBefore.count) {
              landedBy = "shape count";
            } else {
              countedAbsent = true;
            }
          }
          if (!landedBy) await delay(attempt * 250);
        }
      }

      if (!landedBy && tableShape && typeof tableShape.load === "function") {
        // A shape from a batch the host truly discarded has no valid object
        // path, so loading a property throws. One success is enough.
        for (let attempt = 1; attempt <= 3 && !landedBy; attempt++) {
          try {
            tableShape.load("id");
            await meteredSync(context);
            if (tableShape.id !== null && tableShape.id !== undefined) landedBy = "shape id";
          } catch (_) {}
          if (!landedBy) await delay(attempt * 250);
        }
      }

      // Nothing said yes, but nothing said no either: the host simply would
      // not answer. Keep the table. Two slides in the user's deck were given
      // pipe-delimited text drawn straight over a table that was there all
      // along, because this case used to count as proof of absence.
      if (!landedBy && !countedAbsent) landedBy = "no answer from the host (assumed present)";

      if (landedBy) {
        // Kept: native tables are a hard requirement, and discarding one that is
        // probably fine is the worse error. No text is drawn either, so the
        // "text behind an empty grid" artefact stays unreachable on this path.
        tableVerified = true;
        logToPPTConsole(
          `Slide ${slideNum}: the table could not be read back, but it is on the slide ` +
          `(confirmed by ${landedBy}); keeping it.`,
          "warn"
        );
      } else {
        table = null;
        tableShape = null;
        logToPPTConsole(
          `Slide ${slideNum}: no evidence the table shape landed (create sync, shape count ` +
          `and shape id all silent) - falling back to text.`,
          true
        );
      }
    }
  }

  // Sizing and styling apply to ANY table being kept, not only one that could
  // be read back. Gating them on "populated" meant that on a host where the
  // read-back is unavailable -- which is every table on the user's host -- the
  // columns stayed equal and no row height was ever set, so PowerPoint grew the
  // rows to fit their text until the table ran off the bottom of the slide.
  if (tableVerified && table) {
    // `table` is re-read on every call because refreshTableProxy may have
    // swapped it for a freshly resolved one since the last pass.
    const queueCellStyling = () => queueTableCellStyling(table, stylingSpec);

    widthsApplied = await applyTableColumnWidths(context, table, colWidths);
    let heightsApplied = await applyTableRowHeights(
      context, table, widthsApplied ? sizedPlan.rowHeights : plan.rowHeights
    );
    queueCellStyling();
    let styled = await safeSync(context, "table: cell styling");

    // Host refusal windows on Mac PowerPoint have been observed lasting up to
    // 4 seconds (e.g. Slide 8: 12:25:28 -> 12:25:32). A single 500ms retry
    // fires inside the window and leaves the table with equal columns and no
    // header/zebra fill. Retry up to 4 times with progressive backoff
    // (600ms + 1000ms + 1400ms + 1800ms = 4.8s), re-resolving the table proxy
    // before each pass.
    const backoffSchedule = [600, 1000, 1400, 1800];
    for (let pass = 0; pass < backoffSchedule.length; pass++) {
      if (widthsApplied && heightsApplied && styled) break;
      await delay(backoffSchedule[pass]);
      const reResolved = await refreshTableProxy();

      // Every property write goes through the table proxy, so if the host will
      // not hand one back there is nothing left that can succeed. Slide 2 of
      // the 10-slide run spent 15 refused syncs and 4.8s proving that four
      // times over. One pass with no progress and no proxy is enough.
      if (!reResolved && !widthsApplied && !heightsApplied && !styled && pass > 0) {
        break;
      }

      if (!widthsApplied) {
        widthsApplied = await applyTableColumnWidths(context, table, colWidths);
      }
      if (!heightsApplied) {
        heightsApplied = await applyTableRowHeights(
          context, table, widthsApplied ? sizedPlan.rowHeights : plan.rowHeights
        );
      }
      if (!styled) {
        queueCellStyling();
        styled = await safeSync(context, `table: cell styling (retry ${pass + 1})`);
      }
    }

    logToPPTConsole(
      `Slide ${slideNum}: Added native table (${rowCount}x${colCount} @ ${fontSize}pt, ` +
      `${widthsApplied ? "content-weighted columns" : "default equal columns"}` +
      `${plan.clipped ? ", text clipped to fit" : ""}).`
    );

    // Say plainly what the user is looking at, rather than leaving them to
    // infer it from a run of refusals.
    if (!widthsApplied || !heightsApplied || !styled) {
      const missing = [];
      if (!widthsApplied) missing.push("column widths");
      if (!heightsApplied) missing.push("row heights");
      if (!styled) missing.push("header/zebra styling");
      logToPPTConsole(
        `Slide ${slideNum}: the host would not apply ${missing.join(", ")} - the table is ` +
        `there and populated, but in PowerPoint's default style` +
        `${proxyRefreshed ? "" : " (the table could not be re-resolved by id either)"}.`,
        "warn"
      );
    }
  }

  const tableSuccess = tableVerified;

  // Last resort: a plain text frame, so the data is never lost outright.
  let fallbackRendered = false;
  if (!tableSuccess) {
    try {
      const fallbackText = equalPlan.values.map(r => r.join("  |  ")).join("\n");
      const tableBox = addText(newSlide, fallbackText, {
        left: tableLeft,
        top: tableTop,
        width: tableWidth,
        height: usedHeight
      }, {
        font: theme.bodyFont,
        size: fitFontSize(fallbackText, tableWidth, usedHeight, { max: 13, min: 8 }),
        color: bodyText
      });
      if (tableBox) {
        fallbackRendered = await safeSync(context, "table fallback");
      }
      logToPPTConsole(`Slide ${slideNum}: every native table request was refused - rendered as text.`, true);
    } catch (fErr) {
      console.warn("[PPTBuilder] Table fallback textbox failed:", fErr);
    }
  }

  // 0 tells the caller nothing was drawn, so it can reclaim the space.
  return (tableSuccess || fallbackRendered) ? usedHeight : 0;
}

/**
 * Parses markdown bold and bullet markers from text, returning cleanText and bold character ranges.
 */
export function parseMarkdownFormatting(rawContent) {
  if (!rawContent) return { cleanText: "", parsedParagraphs: [] };

  const lines = String(rawContent).split(/\r?\n/);
  const cleanLines = [];
  const parsedParagraphs = [];

  for (const rawLine of lines) {
    if (!rawLine.trim()) {
      cleanLines.push("");
      parsedParagraphs.push({ cleanText: "", boldRanges: [] });
      continue;
    }

    let cleanText = "";
    const boldRanges = [];

    // Match **bold**, __bold__, or text
    const mdRegex = /(\*\*(.*?)\*\*|__([^_]+)__|[^*_]+|[*_])/g;
    let match;
    while ((match = mdRegex.exec(rawLine)) !== null) {
      if (match[2] !== undefined) {
        const boldText = match[2];
        const start = cleanText.length;
        cleanText += boldText;
        boldRanges.push({ start, length: boldText.length });
      } else if (match[3] !== undefined) {
        const boldText = match[3];
        const start = cleanText.length;
        cleanText += boldText;
        boldRanges.push({ start, length: boldText.length });
      } else if (match[0]) {
        cleanText += match[0];
      }
    }

    cleanLines.push(cleanText);
    parsedParagraphs.push({ cleanText, boldRanges });
  }

  return {
    cleanText: cleanLines.join("\n"),
    parsedParagraphs
  };
}

/** Applies bold runs to lead-in phrases. Safe to fail. */
async function applyBoldLeadIns(context, box, cleanText, parsedParagraphs) {
  if (!box || !parsedParagraphs || !parsedParagraphs.some(p => p.boldRanges && p.boldRanges.length > 0)) {
    return;
  }
  try {
    let charOffset = 0;
    for (const p of parsedParagraphs) {
      for (const b of (p.boldRanges || [])) {
        if (b.start >= 0 && b.length > 0 && (charOffset + b.start + b.length) <= cleanText.length) {
          const sub = box.textFrame.textRange.getSubstring(charOffset + b.start, b.length);
          sub.font.bold = true;
        }
      }
      charOffset += p.cleanText.length + 1; // +1 for \n
    }
    await meteredSync(context);
  } catch (boldErr) {
    console.warn("[PPTBuilder] Notice applying bold lead-in:", boldErr);
  }
}

/* ------------------------------------------------------------------ */
/* Modern card archetypes (theme-aware)                                */
/* ------------------------------------------------------------------ */

/**
 * Column geometry for a row of `count` cards.
 *
 * `neededHeight` lets a renderer size the row to the content it actually has, so
 * a short card no longer stretches to the bottom of the slide leaving a large
 * empty box, and a long card is not clipped.
 */
function cardGeometry(layout, count, contentTop, contentBottom, neededHeight = null) {
  const gap = count === 2 ? 30 : count >= 4 ? 16 : 22;
  const cardWidth = (layout.contentWidth - gap * (count - 1)) / count;
  const cardTop = contentTop + 8;
  const available = Math.max(110, contentBottom - cardTop);
  const cardHeight = neededHeight
    ? Math.max(120, Math.min(available, Math.round(neededHeight)))
    : Math.max(140, Math.min(340, available));
  return { gap, cardWidth, cardTop, cardHeight, available };
}

/** Normalizes a card's body into displayable text. */
function cardBodyText(card) {
  if (card.body) return String(card.body);
  if (Array.isArray(card.bullets) && card.bullets.length > 0) {
    return card.bullets.map(b => `\u2022 ${b}`).join("\n\n");
  }
  if (card.description) return String(card.description);
  return "";
}

function cardFootnote(card) {
  return card.footnote || card.source || card.citation || "";
}

/* ------------------------------------------------------------------ */
/* Slim cards: 3 shapes per card                                        */
/* ------------------------------------------------------------------ */

/**
 * Build time on PowerPoint Online tracks shape count (~0.5s per shape in the
 * 09-28 logs), and a card used to be seven shapes: shadow, surface, accent
 * strip, eyebrow, title, body and footnote - eight or nine for a timeline step
 * with its badge. Card slides were 55% of the slides and 79-94% of the build.
 *
 * A slim card is three: the surface, the accent strip, and ONE text box whose
 * runs carry the eyebrow / metric / title / body / footnote styling. A 3-card
 * row drops from 21 shapes to 9 and a 4-step timeline from 29 to 13, with the
 * same colours, type sizes and order on the slide.
 *
 * Run styling uses TextRange.getSubstring (PowerPointApi 1.4); older hosts
 * keep the classic multi-shape renderers below.
 */
// A variable only so the offline harness can exercise the classic renderers;
// nothing in the add-in changes it.
let SLIM_CARDS_ENABLED = true;

function setSlimCardsForTests(on) {
  SLIM_CARDS_ENABLED = Boolean(on);
}

function useSlimCards() {
  return SLIM_CARDS_ENABLED && apiAtLeast("1.4");
}

const SLIM_GAP_PT = 5;

/**
 * Type scales tried, largest first, when a card row has room to spare. Each
 * role is capped (see scaleSlimSize) so labels and footnotes stay small.
 * r49 went up to 1.35 (+5pt): 13pt body text became 17-18pt and the user
 * found it too big next to the r48 cards, so growth is now a light touch.
 */
const SLIM_TYPE_SCALES = [1.1, 1.05];

/** A spacer paragraph: one space at a small size, so it only adds a gap. */
function slimGap(size = SLIM_GAP_PT) {
  return { spacer: true, size: size };
}

function slimRunsHeight(runs, width) {
  let h = 0;
  for (const r of runs) {
    h += r.spacer ? Math.ceil(r.size * LINE_FACTOR) : textBlockHeight(r.text, width, r.style.size);
  }
  return h;
}

/**
 * One text box, several styled paragraphs. `base` styles the whole box; each
 * run then restyles its own paragraph through getSubstring. A refused run
 * write is recorded like any other text write and costs only that run's look.
 */
function addRichText(slide, runs, bounds, base) {
  const parts = runs.map(r => (r.spacer ? " " : String(r.text || "")));
  const box = addText(slide, parts.join("\n"), bounds, base);
  let range = null;
  try {
    range = box.textFrame.textRange;
  } catch (err) {
    noteTextWriteRefused("textFrame.textRange", err);
    return box;
  }
  let offset = 0;
  for (let i = 0; i < runs.length; i++) {
    const len = parts[i].length;
    if (len > 0) {
      try {
        applyFont(range.getSubstring(offset, len), runs[i].spacer ? { size: runs[i].size } : runs[i].style);
      } catch (err) {
        noteTextWriteRefused("textRange.getSubstring", err);
      }
    }
    // Every paragraph break is one character in the host's text.
    offset += len + 1;
  }
  return box;
}

/**
 * Draws a row of slim cards. `opts.plan(item, i, accent, innerWidth)` returns
 * { accent, head: runs[], body: {text, max, min, color} | null, foot: runs[] };
 * the body is the only part that is shrunk (then trimmed) to fit the card.
 */
async function renderSlimCardRow(context, newSlide, items, contentTop, theme, layout, contentBottom, opts) {
  const count = items.length;
  if (count === 0) return false;

  const probe = cardGeometry(layout, count, contentTop, contentBottom);
  const pad = Math.max(12, Math.round(probe.cardWidth * 0.07));
  const innerWidth = probe.cardWidth - pad * 2;
  const basePlans = items.map((item, i) => opts.plan(item, i, theme.accentRamp[i % theme.accentRamp.length], innerWidth));

  const rowNeeds = (list) => {
    let needed = 0;
    for (const p of list) {
      const h = pad * 2 + 6
        + slimRunsHeight(p.head, innerWidth)
        + (p.body ? textBlockHeight(p.body.text, innerWidth, p.body.max) : 0)
        + slimRunsHeight(p.foot, innerWidth);
      if (h > needed) needed = h;
    }
    return needed;
  };

  // The 09-28 screenshots showed three half-empty cards over an empty bottom
  // 40% of the slide: 13pt body text sized for the worst case. Take the
  // largest type scale the row still fits at.
  let plans = basePlans;
  for (const k of SLIM_TYPE_SCALES) {
    const scaled = basePlans.map(p => scaleSlimPlan(p, k));
    if (rowNeeds(scaled) <= probe.available) {
      plans = scaled;
      break;
    }
  }

  const geom = cardGeometry(layout, count, contentTop, contentBottom, rowNeeds(plans));
  // A short row still leaves room; give part of it to the space above, so the
  // row sits in the slide instead of hanging under the title.
  const spare = geom.available - geom.cardHeight;
  if (spare > 24) geom.cardTop += Math.round(Math.min(spare * 0.35, 48));
  const { gap, cardWidth, cardTop, cardHeight } = geom;

  const prologue = typeof opts.prologue === "function" ? () => opts.prologue(geom, pad) : null;
  if (prologue) prologue();
  const committer = createCardCommitter(context, newSlide, prologue);

  for (let i = 0; i < count; i++) {
    const drawOne = () => {
      const plan = plans[i];
      const left = layout.margin + i * (cardWidth + gap);

      // Square corners: the accent strip is a straight bar, and on a rounded
      // card its ends overhung the curve.
      drawCard(newSlide, { left: left, top: cardTop, width: cardWidth, height: cardHeight }, theme, { shadow: false, shape: "rectangle" });
      drawPanel(newSlide, { left: left, top: cardTop, width: cardWidth, height: 5 }, plan.accent, null, "rectangle");

      const boxTop = cardTop + pad;
      const boxHeight = Math.max(24, cardTop + cardHeight - pad - boxTop);

      let runs = plan.head.slice();
      if (plan.body) {
        const room = boxHeight - slimRunsHeight(plan.head, innerWidth) - slimRunsHeight(plan.foot, innerWidth);
        if (room > 18) {
          const fitted = fitTextToBox(plan.body.text, innerWidth, room, { max: plan.body.max, min: plan.body.min });
          if (fitted.clipped && typeof opts.clipNote === "function") {
            logToPPTConsole(opts.clipNote(i), true);
          }
          if (fitted.text && fitted.text.trim()) {
            runs.push({
              text: fitted.text,
              style: { font: theme.bodyFont, size: fitted.size, bold: false, italic: false, color: plan.body.color }
            });
          }
        }
      }
      runs = runs.concat(plan.foot);
      // A trailing spacer would only pad the bottom of the box.
      while (runs.length > 0 && runs[runs.length - 1].spacer) runs.pop();

      if (runs.length > 0) {
        addRichText(newSlide, runs, { left: left + pad, top: boxTop, width: innerWidth, height: boxHeight }, {
          font: theme.bodyFont,
          size: plan.body ? plan.body.max : 12,
          color: plan.body ? plan.body.color : theme.onCardColor
        });
      }
    };

    await committer.commit(opts.label(i), drawOne);
  }

  await committer.flush();
  return true;
}

function slimFootRuns(footnoteText, theme) {
  if (!footnoteText) return [];
  return [
    slimGap(),
    { text: String(footnoteText), style: { font: theme.bodyFont, size: 8.5, bold: false, italic: true, color: theme.mutedColor } }
  ];
}

async function renderSlimKpiCards(context, newSlide, cards, contentTop, theme, layout, contentBottom) {
  const metricSize = cards.length >= 4 ? 26 : 32;
  return renderSlimCardRow(context, newSlide, cards, contentTop, theme, layout, contentBottom, {
    label: i => `kpi card ${i + 1}`,
    plan: (card, i, accent, innerWidth) => {
      const head = [];
      const metric = String(card.metric || "").trim();
      if (metric) {
        head.push({
          text: metric,
          // Already sized to the card width; growing it would wrap it.
          fixed: true,
          style: {
            font: theme.headingFont,
            size: fitFontSize(metric, innerWidth, Math.round(metricSize * 1.5), { max: metricSize, min: 16 }),
            bold: true,
            italic: false,
            color: accent
          }
        });
        head.push(slimGap(3));
      }
      const header = (card.subtitle ? `${card.title || ""}\n${card.subtitle}` : String(card.title || "")).trim();
      if (header) {
        head.push({ text: header, style: { font: theme.headingFont, size: 14, bold: true, italic: false, color: theme.onCardColor } });
        head.push(slimGap());
      }
      const body = cardBodyText(card);
      return {
        accent: accent,
        head: head,
        body: body ? { text: body, max: 12.5, min: 8.5, color: theme.isDark ? "#CBD5E1" : "#475569" } : null,
        foot: slimFootRuns(cardFootnote(card), theme)
      };
    }
  });
}

async function renderSlimColumnCards(context, newSlide, cards, contentTop, theme, layout, contentBottom) {
  const titleMax = cards.length >= 4 ? 14 : 16;
  return renderSlimCardRow(context, newSlide, cards, contentTop, theme, layout, contentBottom, {
    label: i => `column card ${i + 1}`,
    clipNote: i => `Card ${i + 1} text was trimmed to fit its card.`,
    plan: (card, i, accent) => {
      const head = [{
        text: String(card.pill || `PILLAR 0${i + 1}`).toUpperCase(),
        style: { font: theme.bodyFont, size: 9.5, bold: true, italic: false, color: accent }
      }, slimGap(3)];
      const title = String(card.title || "").trim();
      if (title) {
        head.push({ text: title, style: { font: theme.headingFont, size: titleMax, bold: true, italic: false, color: theme.onCardColor } });
        head.push(slimGap());
      }
      const body = cardBodyText(card);
      return {
        accent: accent,
        head: head,
        body: body ? { text: body, max: 13, min: 8.5, color: theme.isDark ? "#CBD5E1" : "#334155" } : null,
        foot: slimFootRuns(cardFootnote(card), theme)
      };
    }
  });
}

async function renderSlimTimeline(context, newSlide, steps, contentTop, theme, layout, contentBottom) {
  const numberSize = 22;
  return renderSlimCardRow(context, newSlide, steps, contentTop, theme, layout, contentBottom, {
    label: i => `timeline step ${i + 1}`,
    clipNote: i => `Timeline step ${i + 1} text was trimmed to fit its card.`,
    // The rail runs behind the cards at the height of the step numbers, so it
    // shows as a connector in each gap. Redrawn by the committer only if the
    // chunk it rode in was refused.
    prologue: (geom, pad) => {
      drawPanel(
        newSlide,
        {
          left: layout.margin + geom.cardWidth * 0.5,
          top: geom.cardTop + pad + Math.round((numberSize * LINE_FACTOR) / 2),
          width: layout.contentWidth - geom.cardWidth,
          height: 2
        },
        theme.isDark ? shade(theme.accentColor, 0.35) : tint(theme.accentColor, 0.65),
        null,
        "rectangle"
      );
    },
    plan: (step, i, accent) => {
      const rawStep = String(step.step === undefined || step.step === null ? i + 1 : step.step).trim();
      const digits = rawStep.match(/\d+/);
      const number = digits ? digits[0] : String(i + 1);
      const eyebrow = /^\d+$/.test(rawStep) ? "" : rawStep.toUpperCase();

      const head = [{
        text: number.length === 1 ? `0${number}` : number,
        style: { font: theme.headingFont, size: numberSize, bold: true, italic: false, color: accent }
      }];
      if (eyebrow) {
        head.push({ text: eyebrow, style: { font: theme.bodyFont, size: 9.5, bold: true, italic: false, color: accent } });
      }
      head.push(slimGap(3));
      const title = String(step.title || "").trim();
      if (title) {
        head.push({ text: title, style: { font: theme.headingFont, size: 15, bold: true, italic: false, color: theme.onCardColor } });
        head.push(slimGap(3));
      }
      const desc = String(step.description || "");
      return {
        accent: accent,
        head: head,
        body: desc.trim() ? { text: desc, max: 12.5, min: 8.5, color: theme.isDark ? "#CBD5E1" : "#334155" } : null,
        foot: slimFootRuns(cardFootnote(step), theme)
      };
    }
  });
}

function scaleSlimSize(size, k) {
  const cap = size < 10 ? size + 0.5 : size + 1.5;
  return Math.min(cap, Math.round(size * k * 2) / 2);
}

function scaleSlimRuns(runs, k) {
  return runs.map(r => {
    if (r.spacer) return { spacer: true, size: scaleSlimSize(r.size, k) };
    if (r.fixed) return r;
    return { text: r.text, style: Object.assign({}, r.style, { size: scaleSlimSize(r.style.size, k) }) };
  });
}

function scaleSlimPlan(p, k) {
  return {
    accent: p.accent,
    head: scaleSlimRuns(p.head, k),
    body: p.body ? Object.assign({}, p.body, { max: Math.min(p.body.max + 4, Math.round(p.body.max * k * 2) / 2) }) : null,
    foot: scaleSlimRuns(p.foot, k)
  };
}

/* ------------------------------------------------------------------ */
/* Styled bullet slides                                                 */
/* ------------------------------------------------------------------ */

/**
 * With the card budget, a deck can have more bullet slides than card slides,
 * and plain black bullets read as unfinished next to the cards. Every bullet
 * slide now gets its bold lead-ins in the accent ramp, plus one treatment in
 * rotation, each costing 0-1 extra shapes (a card row costs 8-9):
 *   bar       - a thin accent bar down the left of the list
 *   panel     - one tinted panel behind the whole list
 *   numbered  - large 01 / 02 / 03 in the accent colours, no extra shape
 * Everything is one text box styled by runs, like the slim cards.
 */
let STYLED_BULLET_SLIDES = 0;
/** Space between a table's bottom edge and the notes drawn under it. */
const TABLE_NOTES_GAP_PT = 28;
const STYLED_BULLET_VARIANTS = ["bar", "panel", "numbered"];

/** Items as {lead, rest}, or null when the body is not a flat 2-8 item list. */
function styledBulletItems(parsedParagraphs) {
  const items = [];
  for (const p of parsedParagraphs || []) {
    const raw = String((p && p.cleanText) || "");
    if (!raw.trim()) continue;
    // A nested list (before/after panels, sub-bullets) keeps plain rendering.
    if (/^\s+[-*\u2022\u25AA]\s/.test(raw)) return null;
    const m = raw.match(/^\s*(?:[\u2022\-*\u25AA]|\d{1,2}[.)])\s*/);
    const skip = m ? m[0].length : 0;
    const text = raw.slice(skip).trim();
    if (!text) continue;
    const bold = (p.boldRanges || [])[0];
    let lead = "";
    let rest = text;
    if (bold && bold.length > 0 && bold.start <= skip + 1) {
      lead = raw.substr(bold.start, bold.length).replace(/[:\-\u2013\u2014]\s*$/, "").trim();
      rest = raw.slice(bold.start + bold.length).replace(/^\s*[:\-\u2013\u2014]\s*/, "").trim();
      if (!lead) rest = text;
    }
    if (!lead) {
      // slideParser strips the ** markers from ordinary bullets, so by the
      // time the body gets here "**Legal Desegregation**: ..." is plain text
      // and has no bold range. A short label before the first colon is the
      // same lead-in, so it gets the same accent colour.
      const colon = styledBulletColonLead(text);
      if (colon) {
        lead = colon.lead;
        rest = colon.rest;
      }
    }
    items.push({ lead: lead, rest: rest });
  }
  return items.length >= 2 && items.length <= 8 ? items : null;
}

/** "Label: rest" -> {lead, rest} when the label is short enough to be a lead-in. */
function styledBulletColonLead(text) {
  const m = String(text || "").match(/^([^:\n]{2,60}?):\s+(\S[\s\S]*)$/);
  if (!m) return null;
  const lead = m[1].trim();
  if (!lead || /[.!?]$/.test(lead) || /https?$/i.test(lead)) return null;
  if (lead.split(/\s+/).length > 7) return null;
  return { lead: lead, rest: m[2].trim() };
}

function styledBulletNumber(i) {
  return (i + 1 < 10 ? "0" : "") + String(i + 1);
}

/** Paragraphs of segments for one list at body size `size`. */
function styledBulletParas(items, theme, variant, size) {
  const paras = [];
  items.forEach((it, i) => {
    if (i > 0) paras.push([{ text: " ", style: { size: Math.max(4, Math.round(size * 0.45)) } }]);
    const accent = theme.accentRamp[i % theme.accentRamp.length];
    const segs = [];
    if (variant === "numbered") {
      segs.push({ text: `${styledBulletNumber(i)}   `, style: { font: theme.headingFont, size: size + 4, bold: true, italic: false, color: accent } });
    }
    if (it.lead) {
      segs.push({ text: it.lead, style: { font: theme.headingFont, size: size, bold: true, italic: false, color: accent } });
      if (it.rest) segs.push({ text: `: ${it.rest}`, style: { font: theme.bodyFont, size: size, bold: false, italic: false, color: theme.bodyColor } });
    } else {
      segs.push({ text: it.rest, style: { font: theme.bodyFont, size: size, bold: false, italic: false, color: theme.bodyColor } });
    }
    paras.push(segs);
  });
  return paras;
}

function styledBulletHeight(items, variant, width, size) {
  let h = 0;
  items.forEach((it, i) => {
    const line = (variant === "numbered" ? `${styledBulletNumber(i)}   ` : "") + (it.lead ? `${it.lead}: ${it.rest}` : it.rest);
    h += measureTextHeight(line, width, variant === "numbered" ? size * 1.08 : size);
    if (i > 0) h += Math.ceil(Math.max(4, Math.round(size * 0.45)) * LINE_FACTOR);
  });
  return h;
}

/** One text box, paragraphs of inline runs styled through getSubstring. */
function addRichParagraphs(slide, paras, bounds, base) {
  const lines = paras.map(segs => segs.map(s => String(s.text || "")).join(""));
  const box = addText(slide, lines.join("\n"), bounds, base);
  let range = null;
  try {
    range = box.textFrame.textRange;
  } catch (err) {
    noteTextWriteRefused("textFrame.textRange", err);
    return box;
  }
  let offset = 0;
  paras.forEach((segs, pi) => {
    for (const seg of segs) {
      const len = String(seg.text || "").length;
      if (len > 0) {
        try {
          applyFont(range.getSubstring(offset, len), seg.style);
        } catch (err) {
          noteTextWriteRefused("textRange.getSubstring", err);
        }
      }
      offset += len;
    }
    if (pi < paras.length - 1) offset += 1;
  });
  return box;
}

/**
 * Draws a bullet slide's body with the next treatment in rotation.
 * Returns { box, text, variant } once it committed (box is null when the host
 * refused the batch but kept it), or null so the caller draws plain text.
 */
async function renderStyledBullets(context, slide, parsedParagraphs, area, theme, label) {
  if (!useSlimCards()) return null;
  const items = styledBulletItems(parsedParagraphs);
  if (!items) return null;

  const variant = STYLED_BULLET_VARIANTS[STYLED_BULLET_SLIDES % STYLED_BULLET_VARIANTS.length];
  const pad = variant === "panel" ? 20 : 0;
  const indent = variant === "bar" ? 18 : 0;
  const textWidth = area.width - pad * 2 - indent;
  const maxText = area.height - pad * 2 - 8;

  let size = 0;
  for (let s = 18; s >= 11; s -= 0.5) {
    if (styledBulletHeight(items, variant, textWidth, s) <= maxText) {
      size = s;
      break;
    }
  }
  // Too long for the treatment at a readable size: the plain box shrinks
  // further and trims, which is the better outcome for a wall of text.
  if (!size) return null;

  const textHeight = styledBulletHeight(items, variant, textWidth, size) + 8;
  const blockHeight = textHeight + pad * 2;
  const top = area.top + Math.round(Math.min(Math.max(0, area.height - blockHeight) * 0.3, 40));
  const paras = styledBulletParas(items, theme, variant, size);
  const plain = paras.map(segs => segs.map(s => s.text).join("")).join("\n");

  let before = knownShapeCount(slide);
  if (!before.ok) before = await countSlideShapes(context, slide);

  let box = null;
  let decor = null;
  // A host can refuse addTextBox synchronously. That throw used to escape to
  // createSingleSlide's catch-all, whose plain-layout recovery and the bare-
  // slide rescue both redrew the body: two copies of the same list on top of
  // each other. Refusals are handled here instead - the decoration is removed
  // and the caller draws the plain body, once.
  try {
    if (variant === "panel") {
      decor = drawPanel(slide, { left: area.left, top: top, width: area.width, height: blockHeight }, theme.cardFill, theme.cardBorder, "roundRectangle");
    } else if (variant === "bar") {
      decor = drawPanel(slide, { left: area.left, top: top + 3, width: 4, height: Math.max(20, textHeight - 6) }, theme.accentColor, null, "rectangle");
    }
    box = addRichParagraphs(slide, paras, {
      left: area.left + pad + indent,
      top: top + pad,
      width: textWidth,
      height: textHeight
    }, { font: theme.bodyFont, size: size, color: theme.bodyColor });
  } catch (drawErr) {
    if (decor && typeof decor.delete === "function") {
      try { decor.delete(); } catch (_) { /* the plain body still draws */ }
    }
    logToPPTConsole(`${label}: the host refused the styled list (${(drawErr && drawErr.message) || drawErr}) - drawing it as plain text.`, "warn");
    return null;
  }

  if (await safeSync(context, label)) {
    STYLED_BULLET_SLIDES += 1;
    return { box: box, text: plain, variant: variant };
  }
  const after = await countSlideShapes(context, slide);
  if (before.ok && after.ok && after.count > before.count) {
    STYLED_BULLET_SLIDES += 1;
    return { box: null, text: plain, variant: variant };
  }
  return null;
}

/** One line per deck: the card budget and which card style this host gets. */
function logCardLayoutPlan(slides) {
  const total = slides.length;
  if (total < 2) return;
  const limit = cardLayoutBudget(total);
  const demoted = [];
  slides.forEach((s, idx) => {
    if (s && s.cardBudgetDemoted) demoted.push(idx + 1);
  });
  logToPPTConsole(
    `Card layouts: up to ${limit} of ${total} slides (${CARD_LAYOUT_SHARE_PERCENT}%, rounded up)` +
    (demoted.length > 0
      ? ` - slide(s) ${demoted.join(", ")} drawn as bullets instead of cards`
      : " - within budget") +
    `; ${useSlimCards() ? "slim cards (3 shapes each)" : "classic cards (host is below PowerPointApi 1.4)"}.`
  );
}

async function renderKpiHeroCards(context, newSlide, visualData, contentTop, theme, layout, contentBottom) {
  const cards = (visualData.cards || []).slice(0, 4);
  if (cards.length === 0) return false;
  if (useSlimCards()) return await renderSlimKpiCards(context, newSlide, cards, contentTop, theme, layout, contentBottom);

  const count = cards.length;
  const probe = cardGeometry(layout, count, contentTop, contentBottom);
  const pad = Math.max(12, Math.round(probe.cardWidth * 0.07));
  const innerWidth = probe.cardWidth - pad * 2;
  const metricSize = count >= 4 ? 26 : 32;

  // Size the row to its tallest card instead of always filling the slide.
  let needed = 0;
  for (const card of cards) {
    const header = card.subtitle ? `${card.title || ""}\n${card.subtitle}` : (card.title || "");
    const h = pad * 2 + 12
      + (card.metric ? Math.round(metricSize * 1.5) + 6 : 0)
      + textBlockHeight(header, innerWidth, 14) + 10
      + textBlockHeight(cardBodyText(card), innerWidth, 12)
      + (cardFootnote(card) ? 28 : 0);
    if (h > needed) needed = h;
  }

  const { gap, cardWidth, cardTop, cardHeight } = cardGeometry(layout, count, contentTop, contentBottom, needed);

  const committer = createCardCommitter(context, newSlide);

  for (let i = 0; i < count; i++) {
    const drawOne = () => {
      const card = cards[i];
      const cardLeft = layout.margin + i * (cardWidth + gap);
      const accentColor = theme.accentRamp[i % theme.accentRamp.length];

      drawCard(newSlide, { left: cardLeft, top: cardTop, width: cardWidth, height: cardHeight }, theme);
      // Colored top strip ties each card to the deck accent ramp.
      drawPanel(newSlide, { left: cardLeft, top: cardTop, width: cardWidth, height: 5 }, accentColor, null, "rectangle");

      let cursor = cardTop + pad + 8;

      const metricText = card.metric || "";
      if (metricText) {
        const metricHeight = Math.round(metricSize * 1.5);
        addText(newSlide, metricText, { left: cardLeft + pad, top: cursor, width: innerWidth, height: metricHeight }, {
          font: theme.headingFont,
          size: fitFontSize(metricText, innerWidth, metricHeight, { max: metricSize, min: 16 }),
          bold: true,
          color: accentColor
        });
        cursor += metricHeight + 6;
      }

      const titleText = card.title || "";
      const subtitleText = card.subtitle || "";
      const headerContent = subtitleText ? `${titleText}\n${subtitleText}` : titleText;
      if (headerContent.trim()) {
        const headerHeight = Math.max(24, textBlockHeight(headerContent, innerWidth, 14) + 6);
        addText(newSlide, headerContent, { left: cardLeft + pad, top: cursor, width: innerWidth, height: headerHeight }, {
          font: theme.headingFont,
          size: fitFontSize(headerContent, innerWidth, headerHeight, { max: 14, min: 10 }),
          bold: true,
          color: theme.onCardColor
        });
        cursor += headerHeight + 8;
      }

      const footnoteText = cardFootnote(card);
      const bodyHeight = cardTop + cardHeight - pad - (footnoteText ? 26 : 0) - cursor;
      const bodyContent = cardBodyText(card);

      if (bodyContent && bodyHeight > 20) {
        const fitted = fitTextToBox(bodyContent, innerWidth, bodyHeight, { max: 12.5, min: 8.5 });
        addText(newSlide, fitted.text, { left: cardLeft + pad, top: cursor, width: innerWidth, height: bodyHeight }, {
          font: theme.bodyFont,
          size: fitted.size,
          color: theme.isDark ? "#CBD5E1" : "#475569"
        });
      }

      if (footnoteText) {
        addText(newSlide, footnoteText, {
          left: cardLeft + pad,
          top: cardTop + cardHeight - pad - 18,
          width: innerWidth,
          height: 22
        }, {
          font: theme.bodyFont,
          size: 8.5,
          italic: true,
          color: theme.mutedColor
        });
      }

      // Commit each card individually: a rejected property write then costs one
      // card instead of wiping every shape queued for the slide.
    };

    await committer.commit(`kpi card ${i + 1}`, drawOne);
  }

  await committer.flush();

  return true;
}

async function renderBeforeAfterCards(context, newSlide, visualData, contentTop, theme, layout, contentBottom) {
  const before = visualData.before || { title: "Current State", bullets: [] };
  const after = visualData.after || { title: "Target State", bullets: [] };

  const probe = cardGeometry(layout, 2, contentTop, contentBottom);
  const pad = 20;
  const innerWidth = probe.cardWidth - pad * 2;

  const bodyOf = (panel) => (panel.bullets || []).map(b => `\u2022 ${b}`).join("\n\n");
  const needed = Math.max(
    pad * 2 + 54 + textBlockHeight(bodyOf(before), innerWidth, 13),
    pad * 2 + 54 + textBlockHeight(bodyOf(after), innerWidth, 13)
  );

  const { gap, cardWidth, cardTop, cardHeight } = cardGeometry(layout, 2, contentTop, contentBottom, needed);

  const panels = [
    {
      left: layout.margin,
      label: before.title || "Current State",
      bullets: before.bullets || [],
      fill: theme.isDark ? "#3F1D22" : "#FFF1F2",
      border: theme.isDark ? "#7F1D2A" : "#FECDD3",
      accent: theme.isDark ? "#FDA4AF" : "#9F1239",
      glyph: "\u{1F534}"
    },
    {
      left: layout.margin + cardWidth + gap,
      label: after.title || "Target State / Transformation",
      bullets: after.bullets || [],
      fill: theme.isDark ? "#123024" : "#F0FDF4",
      border: theme.isDark ? "#166534" : "#BBF7D0",
      accent: theme.isDark ? "#86EFAC" : "#166534",
      glyph: "\u{1F7E2}"
    }
  ];

  for (const panel of panels) {
    drawCard(newSlide, { left: panel.left, top: cardTop, width: cardWidth, height: cardHeight }, theme, {
      fill: panel.fill,
      border: panel.border
    });

    addText(newSlide, `${panel.glyph}  ${panel.label}`, {
      left: panel.left + pad,
      top: cardTop + 16,
      width: cardWidth - pad * 2,
      height: 40
    }, {
      font: theme.headingFont,
      size: fitFontSize(panel.label, cardWidth - pad * 2, 38, { max: 16, min: 11 }),
      bold: true,
      color: panel.accent
    });

    const body = (panel.bullets || []).map(b => `\u2022 ${b}`).join("\n\n");
    if (body.trim()) {
      const bodyHeight = cardHeight - 72;
      const fitted = fitTextToBox(body, cardWidth - pad * 2, bodyHeight, { max: 13, min: 9 });
      addText(newSlide, fitted.text, {
        left: panel.left + pad,
        top: cardTop + 60,
        width: cardWidth - pad * 2,
        height: bodyHeight
      }, {
        font: theme.bodyFont,
        size: fitted.size,
        color: theme.isDark ? "#E2E8F0" : "#374151"
      });
    }

    await safeSync(context, "before/after panel");
  }

  return true;
}

async function renderMultiColumnCards(context, newSlide, cards, contentTop, theme, layout, contentBottom) {
  if (!Array.isArray(cards) || cards.length === 0) return false;
  const displayCards = cards.slice(0, 4);
  if (useSlimCards()) return await renderSlimColumnCards(context, newSlide, displayCards, contentTop, theme, layout, contentBottom);
  const count = displayCards.length;

  const probe = cardGeometry(layout, count, contentTop, contentBottom);
  const pad = Math.max(12, Math.round(probe.cardWidth * 0.07));
  const innerWidth = probe.cardWidth - pad * 2;
  const titleMax = count >= 4 ? 14 : 16;

  let needed = 0;
  for (const card of displayCards) {
    const h = pad * 2 + 26
      + textBlockHeight(card.title || "", innerWidth, titleMax) + 10
      + textBlockHeight(cardBodyText(card), innerWidth, 12.5)
      + (cardFootnote(card) ? 28 : 0);
    if (h > needed) needed = h;
  }

  const { gap, cardWidth, cardTop, cardHeight } = cardGeometry(layout, count, contentTop, contentBottom, needed);

  const committer = createCardCommitter(context, newSlide);

  for (let i = 0; i < count; i++) {
    const drawOne = () => {
      const card = displayCards[i];
      const cardLeft = layout.margin + i * (cardWidth + gap);
      const accentColor = theme.accentRamp[i % theme.accentRamp.length];

      drawCard(newSlide, { left: cardLeft, top: cardTop, width: cardWidth, height: cardHeight }, theme);
      drawPanel(newSlide, { left: cardLeft, top: cardTop, width: cardWidth, height: 5 }, accentColor, null, "rectangle");

      const pillText = String(card.pill || `PILLAR 0${i + 1}`).toUpperCase();
      addText(newSlide, pillText, { left: cardLeft + pad, top: cardTop + pad + 2, width: innerWidth, height: 20 }, {
        font: theme.bodyFont,
        size: 9.5,
        bold: true,
        color: accentColor
      });

      let cursor = cardTop + pad + 26;

      const titleText = card.title || "";
      if (titleText) {
        const titleHeight = Math.max(22, textBlockHeight(titleText, innerWidth, titleMax) + 6);
        addText(newSlide, titleText, { left: cardLeft + pad, top: cursor, width: innerWidth, height: titleHeight }, {
          font: theme.headingFont,
          size: fitFontSize(titleText, innerWidth, titleHeight, { max: titleMax, min: 10 }),
          bold: true,
          color: theme.onCardColor
        });
        cursor += titleHeight + 8;
      }

      const footnoteText = cardFootnote(card);
      const bodyContent = cardBodyText(card);
      const bodyHeight = cardTop + cardHeight - pad - (footnoteText ? 26 : 0) - cursor;

      if (bodyContent && bodyHeight > 20) {
        const fitted = fitTextToBox(bodyContent, innerWidth, bodyHeight, { max: 13, min: 8.5 });
        if (fitted.clipped) {
          logToPPTConsole(`Card ${i + 1} text was trimmed to fit its card.`, true);
        }
        addText(newSlide, fitted.text, { left: cardLeft + pad, top: cursor, width: innerWidth, height: bodyHeight }, {
          font: theme.bodyFont,
          size: fitted.size,
          color: theme.isDark ? "#CBD5E1" : "#334155"
        });
      }

      if (footnoteText) {
        addText(newSlide, footnoteText, {
          left: cardLeft + pad,
          top: cardTop + cardHeight - pad - 18,
          width: innerWidth,
          height: 22
        }, {
          font: theme.bodyFont,
          size: 8.5,
          italic: true,
          color: theme.mutedColor
        });
      }

    };

    await committer.commit(`column card ${i + 1}`, drawOne);
  }

  await committer.flush();

  return true;
}

async function renderProcessTimeline(context, newSlide, steps, contentTop, theme, layout, contentBottom) {
  if (!Array.isArray(steps) || steps.length === 0) return false;
  const displaySteps = steps.slice(0, 4);
  if (useSlimCards()) return await renderSlimTimeline(context, newSlide, displaySteps, contentTop, theme, layout, contentBottom);
  const count = displaySteps.length;

  const probe = cardGeometry(layout, count, contentTop, contentBottom);
  const pad = Math.max(12, Math.round(probe.cardWidth * 0.07));
  const innerWidth = probe.cardWidth - pad * 2;
  const badgeSize = 30;

  let needed = 0;
  for (const step of displaySteps) {
    const h = pad * 2 + badgeSize + 12
      + textBlockHeight(step.title || "", innerWidth, 15) + 8
      + textBlockHeight(step.description || "", innerWidth, 12.5)
      + (cardFootnote(step) ? 28 : 0);
    if (h > needed) needed = h;
  }

  const { gap, cardWidth, cardTop, cardHeight } = cardGeometry(layout, count, contentTop, contentBottom, needed);

  // Connecting rail behind the step cards conveys sequence; aligned to badge
  // centers. Kept as a function because the cards now commit as one batch: if
  // the host refuses that batch it discards the rail too, and the per-card
  // fallback has no other way to put it back.
  const drawRail = () => {
    drawPanel(
      newSlide,
      {
        left: layout.margin + cardWidth * 0.5,
        top: cardTop + pad + Math.round(badgeSize / 2),
        width: layout.contentWidth - cardWidth,
        height: 2
      },
      theme.isDark ? shade(theme.accentColor, 0.35) : tint(theme.accentColor, 0.65),
      null,
      "rectangle"
    );
  };
  drawRail();

  const committer = createCardCommitter(context, newSlide, drawRail);

  for (let i = 0; i < count; i++) {
    const drawOne = () => {
      const step = displaySteps[i];
      const stepLeft = layout.margin + i * (cardWidth + gap);
      const accentColor = theme.accentRamp[i % theme.accentRamp.length];

      drawCard(newSlide, { left: stepLeft, top: cardTop, width: cardWidth, height: cardHeight }, theme);

      // Step badge. Only a very short ordinal fits inside a 30pt circle: a label
      // such as "Phase 1" wraps to three lines and spills out of the badge as
      // white-on-white text, so the full label is shown as an eyebrow instead.
      const rawStep = String(step.step === undefined || step.step === null ? i + 1 : step.step).trim();
      const digits = rawStep.match(/\d+/);
      const badgeLabel = digits ? digits[0] : String(i + 1);
      const eyebrow = /^\d+$/.test(rawStep) ? "" : rawStep.toUpperCase();

      drawPanel(newSlide, { left: stepLeft + pad, top: cardTop + pad, width: badgeSize, height: badgeSize }, accentColor, null, "ellipse");
      addText(newSlide, badgeLabel, {
        left: stepLeft + pad,
        top: cardTop + pad + 7,
        width: badgeSize,
        height: 18
      }, {
        font: theme.headingFont,
        size: badgeLabel.length > 2 ? 10 : 13,
        bold: true,
        align: "Center",
        color: readableTextOn(accentColor)
      });

      if (eyebrow) {
        addText(newSlide, eyebrow, {
          left: stepLeft + pad + badgeSize + 8,
          top: cardTop + pad + 8,
          width: Math.max(30, innerWidth - badgeSize - 8),
          height: 18
        }, {
          font: theme.bodyFont,
          size: 9.5,
          bold: true,
          color: accentColor
        });
      }

      let cursor = cardTop + pad + badgeSize + 12;

      const titleText = step.title || "";
      if (titleText) {
        const titleHeight = Math.max(22, textBlockHeight(titleText, innerWidth, 15) + 6);
        addText(newSlide, titleText, { left: stepLeft + pad, top: cursor, width: innerWidth, height: titleHeight }, {
          font: theme.headingFont,
          size: fitFontSize(titleText, innerWidth, titleHeight, { max: 15, min: 10 }),
          bold: true,
          color: theme.onCardColor
        });
        cursor += titleHeight + 6;
      }

      const footnoteText = cardFootnote(step);
      const descHeight = cardTop + cardHeight - pad - (footnoteText ? 26 : 0) - cursor;
      const descText = step.description || "";
      if (descText && descHeight > 18) {
        const fitted = fitTextToBox(descText, innerWidth, descHeight, { max: 12.5, min: 8.5 });
        if (fitted.clipped) {
          logToPPTConsole(`Timeline step ${i + 1} text was trimmed to fit its card.`, true);
        }
        addText(newSlide, fitted.text, { left: stepLeft + pad, top: cursor, width: innerWidth, height: descHeight }, {
          font: theme.bodyFont,
          size: fitted.size,
          color: theme.isDark ? "#CBD5E1" : "#334155"
        });
      }

      if (footnoteText) {
        addText(newSlide, footnoteText, {
          left: stepLeft + pad,
          top: cardTop + cardHeight - pad - 18,
          width: innerWidth,
          height: 22
        }, {
          font: theme.bodyFont,
          size: 8.5,
          italic: true,
          color: theme.mutedColor
        });
      }

    };

    await committer.commit(`timeline step ${i + 1}`, drawOne);
  }

  await committer.flush();

  return true;
}

async function renderExecutiveSummaryCard(context, newSlide, visualData, contentTop, theme, layout, contentBottom) {
  const summaryText = visualData.summary || visualData.description || "";
  if (!summaryText) return false;

  const cardTop = contentTop + 10;
  const cardHeight = Math.max(120, Math.min(320, contentBottom - cardTop));
  const cardWidth = layout.contentWidth;
  const pad = 26;

  drawCard(newSlide, { left: layout.margin, top: cardTop, width: cardWidth, height: cardHeight }, theme);
  drawPanel(newSlide, { left: layout.margin, top: cardTop, width: 5, height: cardHeight }, theme.accentColor, null, "rectangle");

  addText(newSlide, "EXECUTIVE BRIEF", {
    left: layout.margin + pad,
    top: cardTop + 18,
    width: cardWidth - pad * 2,
    height: 24
  }, {
    font: theme.bodyFont,
    size: 10.5,
    bold: true,
    color: theme.accentColor
  });

  const bodyHeight = cardHeight - 70;
  addText(newSlide, summaryText, {
    left: layout.margin + pad,
    top: cardTop + 48,
    width: cardWidth - pad * 2,
    height: bodyHeight
  }, {
    font: theme.bodyFont,
    size: fitFontSize(summaryText, cardWidth - pad * 2, bodyHeight, { max: 16, min: 10 }),
    color: theme.onCardColor
  });

  return true;
}

/** Highlighted takeaway band pinned to the bottom of the slide. */
function renderTakeawayBanner(newSlide, takeaway, theme, layout) {
  const text = String(takeaway || "").trim();
  if (!text) return 0;

  const bandHeight = 44;
  const top = layout.contentBottom - bandHeight;
  drawPanel(newSlide, { left: layout.margin, top, width: layout.contentWidth, height: bandHeight }, theme.accentSoft, theme.accentColor);
  drawPanel(newSlide, { left: layout.margin, top, width: 4, height: bandHeight }, theme.accentColor, null, "rectangle");

  const label = `💡  ${text}`;
  addText(newSlide, label, {
    left: layout.margin + 18,
    top: top + 11,
    width: layout.contentWidth - 36,
    height: bandHeight - 16
  }, {
    font: theme.bodyFont,
    size: fitFontSize(label, layout.contentWidth - 36, bandHeight - 16, { max: 12.5, min: 9 }),
    bold: true,
    color: theme.isDark ? "#F1F5F9" : shade(theme.accentColor, 0.35)
  });

  return bandHeight + 12;
}

/**
 * Flattens a structured visual payload into plain bullet text.
 *
 * Used as the last-resort slide body: if every styled shape was rejected by the
 * host, the slide still carries its content instead of being left empty.
 */
/**
 * Renders the narrative bullets that accompany a table or a chart.
 *
 * `slideParser` parks these on `slideData.additionalBody` (the chat preview
 * already shows them), but the builder used to drop them entirely, which is
 * why table/chart slides arrived looking almost empty. Only draws when there
 * is real vertical room left under the visual.
 *
 * @returns {number} the height consumed, or 0 when nothing was drawn.
 */
async function renderSupportingNarrative(context, slide, rawText, theme, bounds, label) {
  const text = (rawText || "").trim();
  if (!isSubstantiveSlideBody(text)) return 0;

  const available = Math.floor(bounds.bottom - bounds.top);
  if (available < MIN_NARRATIVE_HEIGHT || bounds.width < 80) return 0;

  const parsed = parseMarkdownFormatting(text);
  const fullClean = (parsed.cleanText || "").trim();
  if (!fullClean) return 0;

  const size = fitFontSize(fullClean, bounds.width, available, { max: 13, min: 8.5 });
  // Even at the minimum size, seven notes under a table ran past the bottom of
  // the slide on the 09-29 16:34 deck. Keep the notes that fit and drop the
  // rest, rather than printing over the slide edge.
  const clean = clampTextToBox(fullClean, bounds.width, available, size).trim();
  if (!clean) return 0;
  if (clean !== fullClean) {
    logToPPTConsole(`${label || "narrative"}: notes trimmed to the space under the visual (${fullClean.length} -> ${clean.length} chars).`);
  }
  const needed = Math.min(available, Math.max(24, Math.ceil(textBlockHeight(clean, bounds.width, size))));

  // Clear anything left over from an earlier shape so the notes below belong to
  // this box alone.
  takeTextWriteRefusals();

  const box = addText(slide, clean, {
    left: bounds.left,
    top: bounds.top,
    width: bounds.width,
    height: needed
  }, {
    font: theme.bodyFont,
    size,
    color: theme.bodyColor
  });

  // This used to be the whole function. The box was queued and its wordWrap and
  // font writes rode along on whatever sync came next - on a table+chart slide
  // that is the generic end-of-content sync, two awaits and a picture later. On
  // the user's deck the box landed and every property write was dropped: the
  // bullets came out 18pt black (PowerPoint's text-box default) while the title
  // above them carried the deck theme. The header gets this right because it is
  // committed by its own safeSync, so the narrative now gets one too.
  if (box && context) {
    await safeSync(context, label || "narrative");
    await verifyNarrativeStyling(context, box, {
      size,
      color: theme.bodyColor,
      label: label || "narrative"
    });
  }

  return box ? needed : 0;
}


/**
 * Reads the narrative box's font back and repairs it once if the host kept
 * something other than what was asked for.
 *
 * A refused property write is indistinguishable from a successful one: the
 * batch still syncs cleanly and every set sits behind a catch. Asking the host
 * what it actually stored is the only way to tell the difference, and it is the
 * measurement that identified this defect in the first place.
 *
 * Diagnostic and self-healing only - never throws, so a host that cannot read
 * font properties back just leaves the slide as it was.
 */
async function verifyNarrativeStyling(context, box, want) {
  const refused = takeTextWriteRefusals();

  let range = null;
  let frame = null;
  try {
    frame = box.textFrame;
    range = frame.textRange;
    range.font.load("size,color");
    frame.load("wordWrap");
    await meteredSync(context);
  } catch (readErr) {
    logToPPTConsole(
      `${want.label}: styling could not be read back (${readErr.message || readErr})` +
      (refused.length ? ` | host refused ${refused.join(", ")}` : ""),
      "warn"
    );
    return;
  }

  let gotSize = null;
  let gotColor = null;
  let gotWrap = null;
  try { gotSize = range.font.size; } catch (_) {}
  try { gotColor = range.font.color; } catch (_) {}
  try { gotWrap = frame.wordWrap; } catch (_) {}

  const sizeOk = gotSize === null || Math.abs(Number(gotSize) - want.size) < 0.51;
  const colorOk = gotColor === null || !want.color ||
    String(gotColor).toUpperCase().replace("#", "") === String(want.color).toUpperCase().replace("#", "");
  // Only a positive `true` counts. A host that will not report wordWrap is
  // exactly the host that cannot be trusted to have applied it, and an
  // unwrapped narrative runs off the side of the slide rather than looking
  // merely unstyled.
  const wrapOk = gotWrap === true;

  if (sizeOk && colorOk && wrapOk && refused.length === 0) {
    logToPPTConsole(
      `${want.label}: styling verified - ${want.size}pt ${want.color || "(theme colour)"}, wrapped.`
    );
    return;
  }

  logToPPTConsole(
    `${want.label}: the host kept ` +
    `${gotSize === null ? "?" : gotSize}pt ${gotColor === null ? "?" : gotColor} ` +
    `wrap=${gotWrap === null ? "?" : gotWrap ? "on" : "OFF"} ` +
    `instead of ${want.size}pt ${want.color || "(theme colour)"} wrap=on` +
    (refused.length ? ` | refused ${refused.join(", ")}` : "") +
    ` - reapplying.`,
    "warn"
  );

  // One repair pass, in its own batch. Writing to an existing shape is a very
  // different operation from writing to one created in the same batch, and it
  // is the one the host has consistently honoured elsewhere.
  try {
    if (!sizeOk) range.font.size = want.size;
    if (!colorOk && want.color) range.font.color = want.color;
    if (!wrapOk) frame.wordWrap = true;
    const repaired = await safeSync(context, `${want.label} (restyle)`);
    logToPPTConsole(
      repaired
        ? `${want.label}: styling reapplied.`
        : `${want.label}: the host refused the restyle - these bullets keep PowerPoint's defaults.`,
      repaired ? "warn" : true
    );
  } catch (fixErr) {
    logToPPTConsole(`${want.label}: restyle failed (${fixErr.message || fixErr}).`, "warn");
  }
}


/**
 * Picks the narrative that belongs under a visual: the parser's dedicated
 * `additionalBody` first, otherwise the plain body (which the table and image
 * branches would otherwise never render).
 */
function supportingNarrativeFor(slideData) {
  const extra = (slideData.additionalBody || "").trim();
  if (isSubstantiveSlideBody(extra)) return extra;
  const body = (slideData.body || "").trim();
  return isSubstantiveSlideBody(body) ? body : "";
}

function tableEchoKey(s) {
  return String(s || "")
    .toLowerCase()
    .replace(/\*\*|__/g, "")
    .replace(/[\u2022\u25AA]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

/**
 * The parser turns a table into bullets ("Row: Header: value  |  Header:
 * value") and those bullets often come back as the slide's narrative, so a
 * table slide printed its own rows a second time underneath. A line is an
 * echo when it names a row's first cell and at least half of that row's other
 * cells; echoes are dropped and anything else in the narrative is kept.
 */
function withoutTableEcho(text, tableData) {
  const raw = String(text || "");
  const rows = (tableData && Array.isArray(tableData.rows)) ? tableData.rows : [];
  if (!raw.trim() || rows.length === 0) return raw;
  const keyed = rows.map(r => (Array.isArray(r) ? r : []).map(tableEchoKey));
  const headerKey = tableEchoKey(((tableData && tableData.headers) || []).join(" "));
  const isEcho = (line) => {
    const k = tableEchoKey(line);
    if (!k) return false;
    if (headerKey && k === headerKey) return true;
    for (const cells of keyed) {
      if (!cells[0] || k.indexOf(cells[0]) === -1) continue;
      const others = cells.slice(1).filter(Boolean);
      if (others.length === 0) return k.length <= cells[0].length + 4;
      const hits = others.filter(c => k.indexOf(c) !== -1).length;
      if (hits >= Math.ceil(others.length / 2)) return true;
    }
    return false;
  };
  const lines = raw.split(/\r?\n/);
  let dropped = 0;
  const kept = lines.filter(line => {
    if (isEcho(line)) {
      dropped += 1;
      return false;
    }
    return true;
  });
  if (dropped === 0) return raw;
  return kept.join("\n").replace(/\n{3,}/g, "\n\n").trim();
}

/** The narrative for a slide that draws a table, minus lines that repeat it. */
function tableNarrativeFor(slideData, tableData, label) {
  const narrative = supportingNarrativeFor(slideData);
  if (!narrative) return narrative;
  const kept = withoutTableEcho(narrative, tableData);
  if (kept !== narrative) {
    const left = isSubstantiveSlideBody(kept) ? kept : "";
    logToPPTConsole(
      `${label}: notes under the table repeated its rows - ` +
      (left ? "kept only the lines that add something." : "dropped them; the table gets the room.")
    );
    return left;
  }
  return narrative;
}

function flattenVisualData(slideData) {
  const vd = slideData.visualData || {};
  const lines = [];

  const pushCard = (c) => {
    if (!c) return;
    const head = [c.metric, c.pill, c.title].filter(Boolean).join(" \u2014 ");
    const body = c.body || c.subtitle || c.description ||
      (Array.isArray(c.bullets) ? c.bullets.join(" ") : "");
    if (head || body) lines.push(`\u2022 ${[head, body].filter(Boolean).join(": ")}`);
  };

  (vd.cards || []).forEach(pushCard);
  (vd.steps || []).forEach((s, i) => pushCard({
    pill: s.step || `Step ${i + 1}`,
    title: s.title,
    body: s.description
  }));
  (slideData.structuredCards || []).forEach(pushCard);

  if (vd.before && Array.isArray(vd.before.bullets) && vd.before.bullets.length > 0) {
    lines.push(`\u2022 ${vd.before.title || "Current State"}: ${vd.before.bullets.join("; ")}`);
  }
  if (vd.after && Array.isArray(vd.after.bullets) && vd.after.bullets.length > 0) {
    lines.push(`\u2022 ${vd.after.title || "Target State"}: ${vd.after.bullets.join("; ")}`);
  }
  if (vd.summary) lines.push(String(vd.summary));

  return lines.join("\n\n");
}

/* ------------------------------------------------------------------ */
/* Slide creation                                                      */
/* ------------------------------------------------------------------ */

/**
 * Discovers the Blank layout (or best clean layout) from the presentation's active theme/master.
 */
const BLANK_LAYOUT_NAME_HINTS = [
  "blank", "branco", "blanco", "vide", "leer", "vuoto", "empty", "clean",
  "tom", "tyhjä", "tomt", "pusty", "prázdn", "üres", "boş", "空白", "빈 화면"
];

/** Reads a property the host may never have loaded, without throwing. */
function safeRead(obj, prop) {
  try {
    const value = obj[prop];
    return value === undefined || value === null ? "" : value;
  } catch (err) {
    return "";
  }
}

/**
 * Turns a host error into something worth putting in front of a user.
 *
 * `debugInfo.errorLocation` is the valuable part: it names the exact call the
 * host rejected, which is the difference between "layouts failed" and
 * "SlideLayoutCollection.load refused items/type".
 */
function describeHostError(err) {
  if (!err) return "unknown";
  const code = err.code ? `${err.code}: ` : "";
  const message = err.message || String(err);
  let where = "";
  try {
    if (err.debugInfo && err.debugInfo.errorLocation) {
      where = ` @ ${err.debugInfo.errorLocation}`;
    }
  } catch (ignored) {
    where = "";
  }
  return `${code}${message}${where}`.slice(0, 200);
}

/**
 * Chooses the emptiest layout from plain descriptors.
 *
 * Takes data rather than live proxies so that every discovery route below can
 * share one decision, and so the decision can be tested without a host.
 *
 * @param candidates [{ masterId, layoutId, name, type, shapeCount }]
 */
function pickLayoutFromCandidates(candidates, route) {
  if (!candidates || candidates.length === 0) return null;

  const summary = candidates
    .slice(0, 10)
    .map(c => `${c.name || "(unnamed)"}${c.type ? `/${c.type}` : ""}${c.shapeCount === null || c.shapeCount === undefined ? "" : `:${c.shapeCount}`}`)
    .join(", ");

  const describe = (chosen, matchedBy) => ({
    options: { slideMasterId: chosen.masterId, layoutId: chosen.layoutId },
    name: chosen.name,
    type: chosen.type,
    shapeCount: chosen.shapeCount === undefined ? null : chosen.shapeCount,
    matchedBy,
    route,
    summary
  });

  // 1. The language-independent answer.
  const byType = candidates.find(c => String(c.type).toLowerCase() === "blank");
  if (byType) return describe(byType, "type");

  // 2. The old name list, kept for hosts that will not report a type.
  const byName = candidates.find(c => {
    const n = String(c.name || "").toLowerCase();
    return BLANK_LAYOUT_NAME_HINTS.some(hint => n.includes(hint));
  });
  if (byName) return describe(byName, "name");

  // 3. No blank layout exists - take the one with the fewest placeholders.
  const counted = candidates.filter(c => typeof c.shapeCount === "number");
  if (counted.length > 0) {
    const leanest = counted.reduce((best, c) => (c.shapeCount < best.shapeCount ? c : best), counted[0]);
    return describe(leanest, leanest.shapeCount === 0 ? "no placeholders" : "fewest placeholders");
  }

  return { options: null, route, summary };
}

/**
 * Best-effort placeholder counts for a set of layout proxies.
 *
 * A host that refuses `getCount()` must not cost us the type and name rungs,
 * whose data is already sitting in local strings by the time we get here.
 */
async function attachPlaceholderCounts(context, layouts) {
  try {
    const counters = layouts.map(l => l.shapes.getCount());
    await meteredSync(context);
    return counters.map(c => (c && typeof c.value === "number" ? c.value : null));
  } catch (err) {
    console.warn("[PPTBuilder] Layout placeholder counts unavailable:", err);
    return layouts.map(() => null);
  }
}

/**
 * Reads the background colour of the layout we are about to build on.
 *
 * This is the canvas, by definition - more authoritative than the slide master,
 * which on one real deck reported Solid #FFFFFF for a layout that renders dark
 * maroon. Returns null for a gradient or picture fill, which cannot yield a
 * single colour.
 */
async function readLayoutBackground(context, layout) {
  try {
    const bg = layout && layout.background;
    if (!bg || !bg.fill) return null;
    bg.fill.load("type");
    const solid = typeof bg.fill.getSolidFillOrNullObject === "function"
      ? bg.fill.getSolidFillOrNullObject()
      : null;
    if (solid) solid.load("color,isNullObject");
    await meteredSync(context);

    const type = safeRead(bg.fill, "type") || "Unknown";
    let color = null;
    if (solid && safeRead(solid, "isNullObject") !== true) {
      const raw = safeRead(solid, "color");
      color = raw ? normalizeThemeHex(raw) : null;
    }
    return { type, color };
  } catch (err) {
    console.warn("[PPTBuilder] Layout background unavailable:", err);
    return null;
  }
}

/** Normalises a host colour that may or may not carry a leading '#'. */
function normalizeThemeHex(raw) {
  const s = String(raw || "").trim();
  if (!s) return null;
  const hex = s.charAt(0) === "#" ? s.slice(1) : s;
  return /^[0-9a-fA-F]{6}$/.test(hex) ? `#${hex.toUpperCase()}` : null;
}

/** Route A/B: walk the slideMasters collection. */
async function layoutsFromSlideMasters(withType) {
  return await PowerPoint.run(async (context) => {
    const slideMasters = context.presentation.slideMasters;
    slideMasters.load("items/id, items/name");
    await meteredSync(context);

    if (!slideMasters.items || slideMasters.items.length === 0) return null;

    for (const master of slideMasters.items) {
      master.layouts.load(withType ? "items/id, items/name, items/type" : "items/id, items/name");
    }
    await meteredSync(context);

    const proxies = [];
    const candidates = [];
    for (const master of slideMasters.items) {
      const items = (master.layouts && master.layouts.items) || [];
      for (const layout of items) {
        proxies.push(layout);
        candidates.push({
          masterId: safeRead(master, "id"),
          layoutId: safeRead(layout, "id"),
          name: safeRead(layout, "name"),
          type: withType ? safeRead(layout, "type") : "",
          shapeCount: null
        });
      }
    }

    if (candidates.length === 0) return null;

    const counts = await attachPlaceholderCounts(context, proxies);
    counts.forEach((value, i) => { candidates[i].shapeCount = value; });

    const picked = pickLayoutFromCandidates(candidates, withType ? "masters+type" : "masters");
    if (picked && picked.options) {
      const chosen = proxies[candidates.findIndex(c => c.layoutId === picked.options.layoutId)];
      picked.background = await readLayoutBackground(context, chosen);
    }
    return picked;
  });
}

/**
 * Route C: read the layouts the deck's existing slides are already using.
 *
 * `Slide.layout` and `Slide.slideMaster` (1.3) reach a SlideLayout without
 * going through the slideMasters collection at all. On the deck that produced
 * "the host would not describe its slide layouts", that collection is exactly
 * what failed - so this route is worth having even though it can only see the
 * layouts already in play.
 */
async function layoutsFromExistingSlides(withType) {
  return await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await meteredSync(context);

    const items = (slides.items || []).slice(0, 12);
    if (items.length === 0) return null;

    const probes = items.map(slide => {
      const layout = slide.layout;
      layout.load(withType ? "id, name, type" : "id, name");
      const master = slide.slideMaster;
      master.load("id");
      return { layout, master };
    });
    await meteredSync(context);

    const seen = {};
    const proxies = [];
    const candidates = [];
    for (const probe of probes) {
      const layoutId = safeRead(probe.layout, "id");
      if (!layoutId || seen[layoutId]) continue;
      seen[layoutId] = true;
      proxies.push(probe.layout);
      candidates.push({
        masterId: safeRead(probe.master, "id"),
        layoutId,
        name: safeRead(probe.layout, "name"),
        type: withType ? safeRead(probe.layout, "type") : "",
        shapeCount: null
      });
    }

    if (candidates.length === 0) return null;

    const counts = await attachPlaceholderCounts(context, proxies);
    counts.forEach((value, i) => { candidates[i].shapeCount = value; });

    const picked = pickLayoutFromCandidates(candidates, "existing slides");
    if (picked && picked.options) {
      const chosen = proxies[candidates.findIndex(c => c.layoutId === picked.options.layoutId)];
      picked.background = await readLayoutBackground(context, chosen);
    }
    return picked;
  });
}

/**
 * Finds the emptiest layout the deck offers, and explains which rung matched.
 *
 * Matching on `name` alone - which is all this used to do before r24 - fails on
 * any deck whose theme renames its layouts. When the search fails we pass no
 * layout to `slides.add()`, and the host then instantiates its own "Title and
 * Content" default, which is where the "Click to add title" / "Click to add
 * subtitle" ghosts come from.
 *
 * Each route runs in its own PowerPoint.run: this host refuses whole contexts
 * at a time, so a failure in one must not cost us the next.
 *
 * Returns `null` when the deck has no layouts at all, otherwise a descriptor
 * `{ options, name, type, shapeCount, matchedBy, route, summary, error }`.
 * `options` is a clean `{ slideMasterId, layoutId }` - nothing else is ever
 * handed to `slides.add()`, because unknown members on that payload are not
 * worth the risk on a host this particular.
 */
/**
 * What each discovery route has done so far this session.
 *
 * Retiring a route is a bet that the host cannot do something, and the bet is
 * only worth making against a route that has never once worked. The plain
 * `masters` route is the one that finds "Blank"; retiring it on a bad run
 * would drop every later build back to a five-placeholder layout, a prompt
 * cleanup on every slide, and the repaint bug behind that. Refusal rates of
 * 30-50% have been measured, so two refusals in a row are ordinary luck.
 *
 * `masters+type` is the opposite case: 6 attempts, 6 refusals, no successes,
 * while the same call without `items/type` succeeds immediately beside it.
 *
 * Session-scoped, never persisted. A reload re-tests every assumption here.
 */
const LAYOUT_ROUTE_HISTORY = new Map();   // name -> { ok: bool, badBuilds: int }

function layoutRouteRecord(name) {
  let rec = LAYOUT_ROUTE_HISTORY.get(name);
  if (!rec) {
    rec = { ok: false, badBuilds: 0 };
    LAYOUT_ROUTE_HISTORY.set(name, rec);
  }
  return rec;
}

/** A route is retired only after two consecutive hopeless builds. */
function layoutRouteIsRetired(name) {
  const rec = LAYOUT_ROUTE_HISTORY.get(name);
  return Boolean(rec) && !rec.ok && rec.badBuilds >= 2;
}

/** Test seam: forget what the routes have done. */
export function resetLayoutRouteMemo() {
  LAYOUT_ROUTE_HISTORY.clear();
}

export async function resolveThemeLayout() {
  const wantType = apiAtLeast("1.8");

  const routes = [];
  if (wantType) routes.push({ name: "masters+type", run: () => layoutsFromSlideMasters(true) });
  routes.push({ name: "masters", run: () => layoutsFromSlideMasters(false) });
  routes.push({ name: "existing slides", run: () => layoutsFromExistingSlides(wantType) });

  let lastError = null;
  let lastEmpty = null;

  for (const route of routes) {
    // Two attempts, because a refusal is not a verdict on the route.
    //
    // meteredSync rethrows, so one refused batch anywhere inside a route kills
    // its whole PowerPoint.run and the route is abandoned. That is how the
    // 12:57 build came to log `layout discovery 9 syncs (2 refused)` and then
    // pick a five-placeholder "Title Slide": both masters routes were refused
    // once each, and discovery fell through to the layouts already in the deck.
    //
    // The bill for that arrives on every slide - a prompt read and a delete
    // each, three of which were themselves refused - and every prompt the
    // per-slide clear misses ends up in the deck-wide sweep, which is the
    // repaint bug. The run before it kept the masters route, matched "Blank"
    // and skipped the whole business.
    // Written off earlier in this session. `masters+type` was refused on four
    // consecutive attempts across two builds while the identical route without
    // `items/type` succeeded immediately both times, which is not how a
    // transient refusal behaves - and a retry cannot recover a host that
    // simply will not serve the property.
    if (layoutRouteIsRetired(route.name)) {
      logToPPTConsole(
        `Layout: skipping the "${route.name}" route - this host refused every attempt ` +
        `at it in two builds running.`
      );
      continue;
    }

    for (let attempt = 1; attempt <= 2; attempt++) {
      try {
        const info = await route.run();
        if (info && info.options) {
          // Proof the route works on this host. From here it is never retired,
          // however badly a later build goes.
          const rec = layoutRouteRecord(route.name);
          rec.ok = true;
          rec.badBuilds = 0;
          return info;
        }
        // A route that answers "no usable layouts" has answered. Only a
        // refusal is worth asking twice.
        if (info) lastEmpty = info;
        break;
      } catch (err) {
        lastError = err;
        console.warn(`[PPTBuilder] Layout discovery route "${route.name}" failed:`, err);
        // The reason only ever reached console.warn, so five builds went by
        // without it being visible in a screenshot. A refusal and an
        // unsupported property need completely different answers, and the
        // host's own words are what tell them apart.
        const why = describeHostError(err);
        if (attempt === 1) {
          logToPPTConsole(
            `Layout: the "${route.name}" route was refused - trying it once more. (${why})`,
            "warn"
          );
          await delay(500);
        } else {
          const rec = layoutRouteRecord(route.name);
          rec.badBuilds += 1;
          const retiring = !rec.ok && rec.badBuilds >= 2;
          logToPPTConsole(
            `Layout: the "${route.name}" route was refused twice - moving on.` +
            (retiring
              ? " It has never worked in this session, so it will not be tried again."
              : "") +
            ` (${why})`,
            "warn"
          );
        }
      }
    }
  }

  if (lastEmpty) return lastEmpty;
  if (lastError) return { options: null, summary: "", error: describeHostError(lastError) };
  return null;
}

/**
 * Shapes PowerPoint creates from a layout's placeholders. Our own shapes are
 * named "TextBox 4" / "Rectangle 9" / "Oval 2" by the host and never match.
 */
const PLACEHOLDER_SHAPE_NAME = /(placeholder|^title\s*\d*$|^subtitle\s*\d*$|^content\s*\d*$)/i;

/**
 * Removes the leftover "Click to add title" placeholders once the deck is built.
 *
 * r17 did this BEFORE adding content, in the SAME context, and every write
 * afterwards was refused - ten slides came out blank. The ordering is the whole
 * difference: this runs in its own PowerPoint.run after every slide is
 * committed, so even a total failure costs nothing but the ghosts.
 *
 * Only shapes that BOTH carry a placeholder name AND have no text are removed.
 * A placeholder we filled is skipped, and so is anything we drew ourselves.
 */
/** The presentation's slide count, or null when the host will not say. */
async function readSlideCount() {
  try {
    return await PowerPoint.run(async (context) => {
      const counter = context.presentation.slides.getCount();
      await meteredSync(context);
      return typeof counter.value === "number" ? counter.value : null;
    });
  } catch (err) {
    console.warn("[PPTBuilder] Could not read the slide count:", err);
    return null;
  }
}

/**
 * One slide's empty placeholder prompts, read fresh.
 *
 * Always re-gets the slide from the presentation rather than reusing a
 * collection proxy: after a delete, a held `shapes` proxy has been seen to
 * replay its old items. The r18 lesson, applied to reads.
 */
async function collectEmptyPlaceholders(context, index) {
  const shapes = context.presentation.slides.getItemAt(index).shapes;
  shapes.load("items/name");
  await meteredSync(context);

  const items = (shapes && shapes.items) || [];
  const named = [];
  for (const shape of items) {
    const name = String(safeRead(shape, "name") || "");
    if (!name || !PLACEHOLDER_SHAPE_NAME.test(name)) continue;
    let frame = null;
    try {
      frame = shape.textFrame || null;
      if (frame) frame.load("hasText");
    } catch (frameErr) {
      frame = null;
    }
    named.push({ name, shape, frame });
  }

  if (named.length > 0) {
    // A shape with no text frame can reject the hasText load at sync time.
    // Losing the answer only costs us precision, so it must not kill the pass.
    try {
      await meteredSync(context);
    } catch (textErr) {
      console.warn("[PPTBuilder] hasText unavailable for this slide:", textErr);
    }
  }

  // A placeholder holding real content is the deck's, not ours.
  const candidates = named.filter(
    (entry) => !(entry.frame && safeRead(entry.frame, "hasText") === true)
  );
  return { listed: items.length, candidates };
}

/** Drops the names the final read can still see from the removed tally. */
function discountSurvivors(removed, survivors) {
  if (survivors.length === 0) return removed;
  return removed.filter((name) => !survivors.includes(name));
}

/**
 * Sweeps ONE slide, in its own context, and verifies the result.
 *
 * Isolation matters: in r28 a single batch covered the whole deck, so one
 * refused slide took the other four with it and the log could only report a
 * deck-wide total. Per slide we can say exactly which deck position still has
 * a ghost - and re-read to find out whether the delete actually landed.
 */
async function sweepOneSlideIn(context, result) {
  const index = result.index;

  let pass = await collectEmptyPlaceholders(context, index);
  // A slide we just populated always has shapes. An empty list means the
  // host answered before it was ready, not that the slide is bare.
  if (pass.listed === 0) pass = await collectEmptyPlaceholders(context, index);
  result.listed = pass.listed;
  if (pass.candidates.length === 0) return;

  for (const candidate of pass.candidates) {
    try {
      candidate.shape.delete();
      result.removed.push(candidate.name);
    } catch (deleteErr) {
      console.warn("[PPTBuilder] Could not remove placeholder:", deleteErr);
    }
  }
  if (result.removed.length === 0) return;
  await meteredSync(context);

  // The delete is only believable once the slide is read back. This also
  // catches a ghost the first enumeration never listed.
  const verify = await collectEmptyPlaceholders(context, index);
  if (verify.candidates.length === 0) return;

  const second = [];
  for (const candidate of verify.candidates) {
    try {
      candidate.shape.delete();
      second.push(candidate.name);
    } catch (deleteErr) {
      console.warn("[PPTBuilder] Could not remove placeholder on retry:", deleteErr);
    }
  }
  if (second.length === 0) {
    result.survivors = verify.candidates.map((entry) => entry.name);
    result.removed = discountSurvivors(result.removed, result.survivors);
    return;
  }
  await meteredSync(context);

  const final = await collectEmptyPlaceholders(context, index);
  result.survivors = final.candidates.map((entry) => entry.name);
  // The second pass may have caught a ghost the first enumeration never
  // listed, so its names count too - but anything the final read can still
  // see was never removed, whatever the proxy said when we asked.
  for (const name of second) result.removed.push(name);
  result.removed = discountSurvivors(result.removed, result.survivors);
}

function blankSweepResult(index) {
  return { index, listed: 0, removed: [], survivors: [], error: null };
}

/* ------------------------------------------------------------------ */
/* Does this build's layout carry prompt placeholders at all?          */
/*                                                                     */
/* "unknown" until a slide manages to read its shape list. Every       */
/* generated slide in a build uses the same layout, so the first       */
/* readable answer holds for all of them.                              */
/*                                                                     */
/* Worth memoising because the read is not free and not safe: it costs */
/* a second slide proxy resolved inside the slide's own drawing        */
/* context, and after that proxy loads a shape collection the host     */
/* stops counting shapes through the builder's own proxy - every slide */
/* of the 11:16 run reported shapes=0 while committing up to 24. On a  */
/* "Blank" layout that price bought nothing at all: no slide had a     */
/* single prompt to clear.                                             */
/* ------------------------------------------------------------------ */
const LAYOUT_PROMPTS_UNKNOWN = 0;
const LAYOUT_PROMPTS_PRESENT = 1;
const LAYOUT_PROMPTS_ABSENT = 2;

let layoutPromptState = LAYOUT_PROMPTS_UNKNOWN;

function resetLayoutPromptState() {
  layoutPromptState = LAYOUT_PROMPTS_UNKNOWN;
}

/** Whether it is still worth reading a new slide's placeholders. */
function shouldProbeLayoutPrompts() {
  return layoutPromptState !== LAYOUT_PROMPTS_ABSENT;
}


/**
 * Queues a read of a freshly created slide's shape names.
 *
 * Split from the delete so the load can ride along on a sync the build was
 * going to make anyway (the slide-tag write), which is what keeps clearing the
 * prompts free on the normal path.
 */
function queuePromptRead(context, index) {
  if (!apiAtLeast("1.4")) return null;   // Shape.name
  if (typeof index !== "number" || index < 0) return null;
  // Already established that this build's layout has nothing to clear. Skipping
  // spares every later slide the second proxy - and the shape count it costs.
  if (!shouldProbeLayoutPrompts()) return null;
  try {
    // Deliberately a throwaway proxy rather than the caller's slide. Loading a
    // collection and then deleting through it leaves that collection stale for
    // everyone still holding it, and the builder holds `newSlide.shapes` for
    // every shape it is about to add and every count it later reads. Resolving
    // the slide again keeps the damage inside this function.
    const shapes = context.presentation.slides.getItemAt(index).shapes;
    shapes.load("items/name");
    return shapes;
  } catch (err) {
    console.warn("[PPTBuilder] Could not queue the prompt-placeholder read:", err);
    return null;
  }
}

/**
 * Deletes the layout's prompt placeholders from a slide we just created, before
 * anything is drawn on it.
 *
 * Why here, when removeEmptyPlaceholders() already sweeps the whole deck: on
 * PowerPoint for the web the document was never wrong - reloading the browser
 * showed clean slides - but the client only repaints a slide when something
 * writes to it. The deck-wide sweep is deliberately the LAST thing that touches
 * the deck, so for every generated slide the final operation was a delete and
 * the canvas went on painting "Click to add title" over our content until the
 * user reloaded. Nothing followed the delete to trigger a repaint.
 *
 * Sweeping first inverts that: the header and body are written after the
 * delete, so the slide repaints in the ordinary course of being built. It is
 * also the cheapest read available: a slide this new holds only the layout's
 * placeholders, so there is nothing else to enumerate.
 *
 * The collection comes from a throwaway slide proxy, never from the one the
 * builder is about to draw on - see queuePromptRead().
 *
 * Never throws. The end-of-deck sweep still runs behind this as a safety net.
 */
async function stripQueuedPrompts(context, shapes, alreadySynced, label) {
  const nothingToDo = { removed: 0, pending: [] };
  if (!shapes) return nothingToDo;

  if (!alreadySynced) {
    try {
      await meteredSync(context);
    } catch (syncErr) {
      // Second refusal of the same queued read. Worth its own line: it says the
      // host is not merely having a bad moment but is refusing this particular
      // request, which is the pattern that precedes a wholly refused context.
      console.warn("[PPTBuilder] Could not read the new slide's placeholders:", syncErr);
      logToPPTConsole(
        `${label}: the host refused the prompt-placeholder read a second time - ` +
        `skipping the early clear, the end-of-deck sweep still runs. ` +
        `${(syncErr && (syncErr.message || syncErr.code)) || syncErr}`,
        "warn"
      );
      return nothingToDo;
    }
  }

  let items = [];
  try {
    items = (shapes && shapes.items) || [];
  } catch (readErr) {
    console.warn("[PPTBuilder] The new slide listed no shapes:", readErr);
    return nothingToDo;
  }

  // Collect first, delete second. Deleting while iterating the collection the
  // host handed back shifts the remaining entries down, so the loop steps over
  // whatever moved into the slot it just left - and on a title+subtitle layout
  // that skipped shape is always the subtitle. Nothing has been drawn on this
  // slide yet, so every placeholder here is an empty prompt by construction:
  // no hasText probe is needed, and with it goes the second round trip that
  // made this too expensive to do per slide before.
  const doomed = [];
  for (const shape of items) {
    const name = String(safeRead(shape, "name") || "");
    if (!name || !PLACEHOLDER_SHAPE_NAME.test(name)) continue;
    doomed.push({ name, shape });
  }

  // The read succeeded, so this is a real answer about the layout rather than
  // a failure to find out - which is why it is recorded here and not in the
  // error paths above. Every generated slide uses the same layout, so one
  // answer settles the whole build.
  if (layoutPromptState === LAYOUT_PROMPTS_UNKNOWN) {
    layoutPromptState = doomed.length > 0 ? LAYOUT_PROMPTS_PRESENT : LAYOUT_PROMPTS_ABSENT;
    if (doomed.length === 0) {
      logToPPTConsole(
        `${label}: this layout has no prompt placeholders - skipping the per-slide ` +
        `check for the rest of the build.`
      );
    }
  }

  const removed = [];
  for (const entry of doomed) {
    try {
      entry.shape.delete();
      removed.push(entry.name);
    } catch (delErr) {
      console.warn("[PPTBuilder] Could not remove a prompt placeholder:", delErr);
    }
  }
  if (removed.length === 0) return nothingToDo;

  // Committed here rather than left to ride on the next write. A dropped delete
  // is indistinguishable from one that landed, and the point of moving this
  // earlier is precisely that the writes which follow repaint the slide.
  const ok = await safeSync(context, `${label} prompt cleanup`);
  if (!ok) {
    // Handed back rather than abandoned. A refused batch is discarded whole, so
    // these prompts really are still on the slide, and leaving them to the
    // end-of-deck sweep is what paints "Click to add title" over the finished
    // slide until the user reloads - see retryQueuedPromptDeletes.
    return { removed: 0, pending: doomed };
  }

  logToPPTConsole(`${label}: cleared ${removed.length} layout prompt(s) before drawing - ${removed.join(", ")}.`);
  return { removed: removed.length, pending: [] };
}

/**
 * Deletes prompt placeholders the first attempt could not.
 *
 * Called after the header has been drawn, for the same reason the header's own
 * third attempt is made after the content: this host's refusal windows are
 * short, and simply trying again straight away tends to spend the retry inside
 * the window that caused the failure.
 *
 * Position matters as much as the retry. Running here keeps the order that
 * makes the fix work - header, then delete, then everything else - so a write
 * always follows the delete and PowerPoint for the web repaints the slide on
 * its own. Deferring to the end-of-deck sweep is what loses that.
 *
 * Takes its own sync. Re-deleting a shape the host quietly committed the first
 * time throws at commit, and letting that ride on the content batch would cost
 * the slide its content.
 *
 * Never throws.
 */
async function retryQueuedPromptDeletes(context, pending, label) {
  if (!pending || pending.length === 0) return 0;

  const removed = [];
  for (const entry of pending) {
    try {
      entry.shape.delete();
      removed.push(entry.name);
    } catch (delErr) {
      console.warn("[PPTBuilder] Could not remove a prompt placeholder on retry:", delErr);
    }
  }
  if (removed.length === 0) return 0;

  const ok = await safeSync(context, `${label} prompt cleanup (retry)`);
  logToPPTConsole(
    ok
      ? `${label}: cleared ${removed.length} layout prompt(s) on the second attempt - ${removed.join(", ")}.`
      : `${label}: the host refused the prompt cleanup twice - the end-of-deck sweep will ` +
        `remove them, but this slide may show the prompts until it is reloaded.`,
    ok ? undefined : "warn"
  );
  return ok ? removed.length : 0;
}

async function sweepSlidePlaceholders(index) {
  const result = blankSweepResult(index);

  try {
    await PowerPoint.run(async (context) => {
      await sweepOneSlideIn(context, result);
    });
  } catch (err) {
    result.error = describeHostError(err);
  }

  return result;
}

/**
 * Sweeps every slide from ONE context, or returns null if the host refuses the
 * context outright.
 *
 * This is not the r28 mistake. r28 put the whole deck in a single queued
 * operation, so one refused slide discarded the other four and the log could
 * only report a deck-wide total. Here the slides are still swept one after
 * another, each with its own try/catch and its own result - what is shared is
 * the context, not the batch. Five contexts cost 8.5 seconds on a 17-slide
 * deck, and most of that was setup rather than the reads.
 *
 * Re-sweeping is idempotent, so the caller is free to fall back to the
 * per-slide path even if this got partway through.
 */
async function sweepSlidesInOneContext(slideIndexes) {
  const results = slideIndexes.map(blankSweepResult);

  try {
    await PowerPoint.run(async (context) => {
      for (const result of results) {
        try {
          await sweepOneSlideIn(context, result);
        } catch (err) {
          result.error = describeHostError(err);
        }
      }
    });
  } catch (err) {
    console.warn("[PPTBuilder] Shared placeholder sweep context failed:", err);
    return null;
  }

  return results;
}

/**
 * The placeholder candidates for MANY slides, in two syncs instead of two per
 * slide. Returns null - meaning "use the per-slide path" - if the host refuses
 * either read.
 */
async function collectEmptyPlaceholdersBatched(context, indexes) {
  // Re-get every slide from the presentation rather than reusing a held
  // collection proxy: after a delete, a held `shapes` proxy has been seen to
  // replay its old items. The r18 lesson, applied to the batch.
  const perSlide = indexes.map((index) => ({
    index,
    shapes: context.presentation.slides.getItemAt(index).shapes,
    listed: 0,
    candidates: []
  }));

  for (const entry of perSlide) entry.shapes.load("items/name");
  try {
    await meteredSync(context);
  } catch (namesErr) {
    console.warn("[PPTBuilder] Batched placeholder name read refused:", namesErr);
    return null;
  }

  let named = 0;
  for (const entry of perSlide) {
    const items = (entry.shapes && entry.shapes.items) || [];
    entry.listed = items.length;
    for (const shape of items) {
      const name = String(safeRead(shape, "name") || "");
      if (!name || !PLACEHOLDER_SHAPE_NAME.test(name)) continue;
      let frame = null;
      try {
        frame = shape.textFrame || null;
        if (frame) frame.load("hasText");
      } catch (frameErr) {
        frame = null;
      }
      entry.candidates.push({ name, shape, frame });
      named += 1;
    }
  }

  if (named > 0) {
    try {
      await meteredSync(context);
    } catch (textErr) {
      // Per slide, an unreadable hasText only costs precision. Batched it
      // would mean deleting every named placeholder in the DECK on the
      // strength of an answer we never got - including ones the user filled.
      // hasText is the only thing protecting their content, so hand back.
      console.warn("[PPTBuilder] Batched hasText read refused:", textErr);
      return null;
    }
    for (const entry of perSlide) {
      // A placeholder holding real content is the deck's, not ours.
      entry.candidates = entry.candidates.filter(
        (entry2) => !(entry2.frame && safeRead(entry2.frame, "hasText") === true)
      );
    }
  }

  return perSlide;
}

/**
 * Sweeps the whole deck in a FLAT four syncs instead of four per slide.
 *
 * This is not the r28 mistake either. r28 batched the deck and had nowhere to
 * go when the host said no: one refused slide silently cost the other four
 * their cleanup, and the log could only report a deck-wide total. Here every
 * stage checks, and a refusal anywhere abandons the batch and returns null so
 * the caller can re-run the proven per-slide path. Re-sweeping is idempotent,
 * so a batch that got halfway costs nothing but the syncs it already spent.
 *
 * @returns {Promise<Array|null>} results, or null to ask for the slower path.
 */
async function sweepSlidesBatched(slideIndexes) {
  const results = slideIndexes.map(blankSweepResult);
  const byIndex = new Map(results.map((r) => [r.index, r]));

  try {
    return await PowerPoint.run(async (context) => {
      let passes = await collectEmptyPlaceholdersBatched(context, slideIndexes);
      if (!passes) return null;

      // A slide we just populated always has shapes. An empty list means the
      // host answered before it was ready, not that the slide is bare.
      const notReady = passes.filter((p) => p.listed === 0).map((p) => p.index);
      if (notReady.length > 0) {
        const retry = await collectEmptyPlaceholdersBatched(context, notReady);
        if (!retry) return null;
        const fresh = new Map(retry.map((p) => [p.index, p]));
        passes = passes.map((p) => fresh.get(p.index) || p);
      }

      for (const pass of passes) byIndex.get(pass.index).listed = pass.listed;

      const withCandidates = passes.filter((p) => p.candidates.length > 0);
      if (withCandidates.length === 0) return results;

      for (const pass of withCandidates) {
        const result = byIndex.get(pass.index);
        for (const candidate of pass.candidates) {
          try {
            candidate.shape.delete();
            result.removed.push(candidate.name);
          } catch (deleteErr) {
            console.warn("[PPTBuilder] Could not remove placeholder:", deleteErr);
          }
        }
      }

      const touched = results.filter((r) => r.removed.length > 0);
      if (touched.length === 0) return results;

      try {
        await meteredSync(context);
      } catch (deleteErr) {
        // The deletes were queued together so they failed together: nothing
        // was removed and nothing is half-done.
        console.warn("[PPTBuilder] Batched placeholder delete refused:", deleteErr);
        return null;
      }

      // A delete is only believable once the slide is read back.
      const verify = await collectEmptyPlaceholdersBatched(
        context,
        touched.map((r) => r.index)
      );
      if (!verify) return null;

      const stillThere = verify.filter((p) => p.candidates.length > 0);
      if (stillThere.length === 0) return results;

      // Second attempt, mirroring the per-slide path. Falling back here
      // instead would be worse than useless: the deletes above have already
      // been committed, so a re-sweep would find nothing left and report that
      // nothing was removed.
      //
      // The verify pass can also surface a ghost the opening enumeration never
      // listed, so its names count too.
      const second = [];
      for (const pass of stillThere) {
        const result = byIndex.get(pass.index);
        for (const candidate of pass.candidates) {
          try {
            candidate.shape.delete();
            result.removed.push(candidate.name);
            second.push(candidate.name);
          } catch (retryErr) {
            console.warn("[PPTBuilder] Could not remove placeholder on retry:", retryErr);
          }
        }
      }

      const markSurvivors = (passes) => {
        for (const pass of passes) {
          const result = byIndex.get(pass.index);
          result.survivors = pass.candidates.map((candidate) => candidate.name);
          result.removed = discountSurvivors(result.removed, result.survivors);
        }
      };

      if (second.length === 0) {
        markSurvivors(stillThere);
        return results;
      }

      try {
        await meteredSync(context);
      } catch (retrySyncErr) {
        console.warn("[PPTBuilder] Batched placeholder retry refused:", retrySyncErr);
        return null;
      }

      // Anything a final read can still see was never removed, whatever the
      // proxy said when we asked.
      const final = await collectEmptyPlaceholdersBatched(
        context,
        stillThere.map((pass) => pass.index)
      );
      if (!final) return null;
      markSurvivors(final);

      return results;
    });
  } catch (err) {
    console.warn("[PPTBuilder] Batched placeholder sweep failed:", err);
    return null;
  }
}

/**
 * Writes something harmless to slides the deck-wide sweep just cleaned, so the
 * web client repaints them.
 *
 * The sweep leaves the document correct and the screen wrong. It is the last
 * thing that touches the deck, and PowerPoint for the web only repaints a slide
 * when something writes to it, so a slide whose final operation was a delete
 * goes on showing "Click to add title" over finished content until the user
 * reloads the browser.
 *
 * Re-writing the add-in tag is the cheapest write available and changes nothing
 * a user can see. Whether the host treats a tag write as enough to invalidate
 * the rendered slide is not documented anywhere, so the outcome is logged:
 * a build that still shows prompts after this line will say so.
 *
 * Best effort. Never throws, and the deck is already correct without it.
 */
async function nudgeSlidesToRepaint(slideIndexes) {
  if (!Array.isArray(slideIndexes) || slideIndexes.length === 0) return false;

  let committed = false;
  try {
    await PowerPoint.run(async (ctx) => {
      let queued = 0;
      for (const index of slideIndexes) {
        try {
          const slide = ctx.presentation.slides.getItemAt(index);
          if (slide && slide.tags) {
            slide.tags.add(ADDIN_SLIDE_TAG, BUILDER_REVISION);
            queued += 1;
          }
        } catch (slideErr) {
          console.warn("[PPTBuilder] Could not queue a repaint nudge:", slideErr);
        }
      }
      if (queued === 0) return;
      committed = await safeSync(ctx, "repaint nudge after the placeholder sweep");
    });
  } catch (err) {
    console.warn("[PPTBuilder] Repaint nudge notice:", err);
  }

  logToPPTConsole(
    committed
      ? `Cleanup: nudged ${slideIndexes.length} swept slide(s) so the canvas repaints.`
      : `Cleanup: could not nudge the swept slide(s) - they may show their prompts ` +
        `until the deck is reloaded.`,
    committed ? undefined : "warn"
  );
  return committed;
}

// Exported for the offline harness: this is the riskiest part of the build
// to change blind, since it is what removed the "Click to add title"
// prompts the user reported. Nothing else imports it.
export async function removeEmptyPlaceholders(slideIndexes) {
  if (!Array.isArray(slideIndexes) || slideIndexes.length === 0) return 0;
  if (!apiAtLeast("1.4")) return 0;   // Shape.name and TextFrame.hasText

  // Fastest first. Each rung is slower and more defensive than the last, and
  // every one of them is idempotent, so falling through costs only round trips.
  let results = await sweepSlidesBatched(slideIndexes);

  if (!results) {
    // Worth saying out loud: if this starts firing on every build, the batch
    // is not earning its place.
    logToPPTConsole("Cleanup: the batched sweep was refused; falling back.", "warn");
    results = await sweepSlidesInOneContext(slideIndexes);
  }

  let alreadyIsolated = false;

  if (!results) {
    // The shared context failed as a whole. Fall back to the proven path so a
    // single bad slide cannot cost the others their cleanup.
    logToPPTConsole("Cleanup: shared sweep context refused; retrying slide by slide.", "warn");
    results = [];
    for (const index of slideIndexes) {
      // Sequential on purpose: two PowerPoint.run contexts racing over the same
      // deck is how r17 started.
      results.push(await sweepSlidePlaceholders(index));
    }
    alreadyIsolated = true;
  }

  // A slide can fail inside a SHARED context because an earlier slide poisoned
  // it, not because anything is wrong with it - this host refuses every sync in
  // a context once one getItemAt has gone bad. Sharing the context was the
  // whole point of Phase 1, but it quietly gave up the isolation that made a
  // single bad slide survivable, so every slide after it lost its cleanup and
  // the prompts came back.
  //
  // Retry just the casualties, each in its own context. Nothing failed on the
  // normal path, so this costs nothing there.
  if (!alreadyIsolated) {
    for (let i = 0; i < results.length; i++) {
      if (!results[i].error) continue;
      results[i] = await sweepSlidePlaceholders(results[i].index);
    }
  }

  const removedTotal = results.reduce((sum, r) => sum + r.removed.length, 0);
  const listedTotal = results.reduce((sum, r) => sum + r.listed, 0);
  const refused = results.filter((r) => r.error);
  const stuck = results.filter((r) => r.survivors.length > 0);

  if (removedTotal > 0) {
    // Deck position, not array index - so a screenshot can be matched to a line.
    const detail = results
      .filter((r) => r.removed.length > 0)
      .map((r) => `slide ${r.index + 1} [${r.removed.join(", ")}]`)
      .join("; ");
    logToPPTConsole(`Cleanup: removed ${removedTotal} empty placeholder prompt(s) - ${detail}.`);

    // These slides ended on a delete, which leaves the web client painting the
    // prompts it just removed. Give each of them a write to repaint from.
    await nudgeSlidesToRepaint(
      results.filter((r) => r.removed.length > 0).map((r) => r.index)
    );
  } else if (listedTotal === 0 && refused.length === 0) {
    logToPPTConsole("Cleanup: the host listed no shapes, so placeholder prompts were left alone.");
  }

  if (stuck.length > 0) {
    const detail = stuck
      .map((r) => `slide ${r.index + 1} [${r.survivors.join(", ")}]`)
      .join("; ");
    logToPPTConsole(
      `Cleanup: ${detail} still shows a prompt after two passes. They are edit-view prompts and do not appear in a slideshow.`,
      "warn"
    );
  }

  if (refused.length > 0) {
    const detail = refused.map((r) => `slide ${r.index + 1} (${r.error})`).join("; ");
    logToPPTConsole(`Cleanup: the host refused the placeholder sweep on ${detail}.`, "warn");
  }

  return removedTotal;
}

/**
 * Test seam for removeEmptyPlaceholders().
 *
 * The cleanup pass is the one piece of this file that DELETES user content, so
 * it is worth exercising directly rather than only through buildPresentation.
 */
export function __testRemoveEmptyPlaceholders(slideIndexes) {
  return removeEmptyPlaceholders(slideIndexes);
}

async function getThemeBlankLayoutOptions() {
  let info = null;
  try {
    info = await resolveThemeLayout();
  } catch (err) {
    console.warn("[PPTBuilder] Could not query slide masters/layouts:", err);
    logToPPTConsole(`Layout: discovery failed outright (${describeHostError(err)}) - using PowerPoint's default.`, "warn");
    return null;
  }

  if (!info || !info.options) {
    const reason = info && info.error
      ? `host refused every route (${info.error})`
      : `no blank layout in this theme`;
    const seen = info && info.summary ? info.summary : "none reported";
    logToPPTConsole(
      `Layout: ${reason} - PowerPoint's default will be used, so placeholder prompts may appear. Layouts seen: ${seen}`,
      "warn"
    );
    return null;
  }

  const placeholders = info.shapeCount === null ? "?" : info.shapeCount;
  const bg = info.background;
  const bgNote = bg
    ? ` | background ${bg.type}${bg.color ? ` ${bg.color}` : " (no flat colour)"}`
    : "";
  logToPPTConsole(
    `Layout: "${info.name || "(unnamed)"}" type=${info.type || "?"} placeholders=${placeholders} - matched by ${info.matchedBy} via ${info.route}.${bgNote}`
  );
  return info;
}

/**
 * Reads the ids of every slide currently in the presentation.
 *
 * Returns `null` - not an empty list - when the host will not hand them over,
 * so a caller can tell "no slides" apart from "cannot see the slides".
 */
async function readSlideIds(context) {
  try {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await meteredSync(context);
    const items = slides.items;
    if (!items || !items.length) return null;
    const ids = [];
    for (let i = 0; i < items.length; i++) {
      const id = items[i] && items[i].id;
      if (!id) return null;
      ids.push(String(id));
    }
    return ids;
  } catch (err) {
    console.warn("[PPTBuilder] Could not read slide ids:", err);
    return null;
  }
}

/**
 * Appends one slide to the tail of the presentation and returns its resolved proxy.
 *
 * CRITICAL: PowerPoint.SlideCollection.add() returns void - the new Slide has to
 * be looked up after a sync. Assigning add()'s return value yields `undefined`
 * and every subsequent `.shapes` access throws.
 *
 * It used to be looked up with `getItemAt(getCount() - 1)`, which is unsafe:
 * `getCount()` lags badly on this host. A count that is one behind resolves the
 * PREVIOUS slide, so two slides' worth of content lands on one slide; a count
 * that is ahead resolves nothing, and every later write on that proxy fails
 * with "InvalidParam passed to GetItem(id)". Both were observed in the wild.
 *
 * So the slide is identified by ID: snapshot the ids, add, snapshot again, and
 * take the id that is new. The index path stays only as a last resort for when
 * the host refuses to hand over the ids.
 */
/**
 * Where appendSlide() last put a slide, for the placeholder sweep to collect.
 *
 * A module-level hand-off rather than a return value: appendSlide's proxy is
 * consumed inside createSingleSlide's own PowerPoint.run, and threading an
 * index back out would mean changing four signatures on the hot path. Only one
 * build runs at a time, and takeAppendedSlideIndex() clears it, so a slide that
 * failed to append can never donate its index to the next one.
 */
let lastAppendedSlideIndex = -1;

// The slide ids as of the end of the last successful append. Null means
// "unknown", which forces a fresh read. Only ever set from a list the host
// handed over AFTER a slide was added.
let cachedSlideIdsAfterAppend = null;

function resetSlideIdCache() {
  cachedSlideIdsAfterAppend = null;
}

function recordAppendedSlideIndex(index) {
  lastAppendedSlideIndex = typeof index === "number" && index >= 0 ? index : -1;
}

function takeAppendedSlideIndex() {
  const index = lastAppendedSlideIndex;
  lastAppendedSlideIndex = -1;
  return index;
}

/**
 * The same index without consuming it.
 *
 * The prompt clear needs to name the slide it just appended while
 * createSingleSlide is still running, but the build loop reads the index after
 * createSingleSlide returns and would find -1 if we took it here.
 */
function peekAppendedSlideIndex() {
  return lastAppendedSlideIndex;
}

async function appendSlide(context, layoutOptions) {
  const slides = context.presentation.slides;
  recordAppendedSlideIndex(-1);

  // The id list the PREVIOUS append ended with is still accurate: between then
  // and now the builder only added shapes to a slide, which cannot change the
  // slide list. Reading it again costs a marshal of every id in the deck, and
  // that was the single most deck-size-sensitive call in the build.
  //
  // Taken, not peeked: a failure below must never leave a stale list behind.
  // The authoritative read is still the one taken AFTER the add.
  let beforeIds = cachedSlideIdsAfterAppend;
  cachedSlideIdsAfterAppend = null;
  let startCount = 0;

  if (beforeIds) {
    startCount = beforeIds.length;
  } else {
    // Queued BEFORE readSlideIds so its sync resolves both, rather than
    // spending a second round trip on the count alone.
    const beforeCount = slides.getCount();
    beforeIds = await readSlideIds(context);
    try {
      startCount = beforeCount.value || 0;
    } catch (_) {
      startCount = 0;
    }
  }

  let added = false;
  if (layoutOptions && layoutOptions.slideMasterId && layoutOptions.layoutId) {
    try {
      slides.add(layoutOptions);
      await meteredSync(context);
      added = true;
    } catch (layoutErr) {
      console.warn("[PPTBuilder] Themed layout add failed, retrying with the default layout:", layoutErr);
    }
  }

  if (!added) {
    slides.add();
    await meteredSync(context);
  }

  // ---- preferred: find the id that was not there before --------------
  if (beforeIds) {
    const known = {};
    for (let i = 0; i < beforeIds.length; i++) known[beforeIds[i]] = true;
    for (let attempt = 1; attempt <= 3; attempt++) {
      const afterIds = await readSlideIds(context);
      if (afterIds) {
        let newIndex = -1;
        for (let i = afterIds.length - 1; i >= 0; i--) {
          if (!known[afterIds[i]]) { newIndex = i; break; }
        }
        if (newIndex >= 0) {
          logToPPTConsole(
            `Slide resolved by id: index ${newIndex} of ${afterIds.length} ` +
            `(was ${beforeIds.length} slides).`
          );
          recordAppendedSlideIndex(newIndex);
          // This read IS authoritative, so the next append can start from it.
          cachedSlideIdsAfterAppend = afterIds;
          return context.presentation.slides.getItemAt(newIndex);
        }
      }
      // The id list can lag exactly like the count does; give it a moment.
      await meteredSync(context);
    }
    logToPPTConsole("Slide ids did not show the new slide - falling back to the index.", true);
  }

  // ---- fallback: the old, count-based path ----------------------------
  const afterCount = slides.getCount();
  await meteredSync(context);

  if ((afterCount.value || 0) <= startCount) {
    throw new Error("PowerPoint did not append a new slide to the presentation.");
  }

  logToPPTConsole(
    `Slide resolved by index ${afterCount.value - 1} (count ${startCount} -> ${afterCount.value}).`
  );
  recordAppendedSlideIndex(afterCount.value - 1);
  return slides.getItemAt(afterCount.value - 1);
}

/**
 * Creates a single slide: title, subtitle, and either a table, images, modern cards
 * or bullets - styled with the deck's own theme tokens.
 */
/**
 * Redraws a slide's essentials in a brand-new PowerPoint.run.
 *
 * Once this host answers "InvalidParam passed to GetItem(id)" it refuses every
 * remaining sync in that context. A slide can queue its title, its subtitle and
 * thirty content shapes in 1.4 seconds and land none of them - which is exactly
 * how slide 8 of the user's deck arrived blank while the build log called it
 * "Created (content)". The in-context rescue cannot help there: it would join
 * the same refused batch. Only a new context can, and the slide itself is still
 * there to draw on because an earlier sync committed the canvas.
 *
 * Deliberately plain - one title box, one subtitle, one text box. No cards, no
 * tables, no read-backs, nothing that could fail halfway. The point is that the
 * user gets their content, not that it matches the other nine slides.
 *
 * @param {number} slideIndex zero-based index of the slide to rebuild.
 * @param {number} slideNum   the slide's number, for logging.
 * @param {{title: string, subtitle: string, color: ?string, text: string}} content
 * @param {object} theme
 * @returns {Promise<boolean>} whether the redraw committed.
 */
async function rescueSlideInFreshContext(slideIndex, slideNum, content, theme) {
  if (!(slideIndex >= 0)) return false;

  const layout = getLayout(theme);
  let committed = false;

  try {
    await PowerPoint.run(async (ctx) => {
      const slide = ctx.presentation.slides.getItemAt(slideIndex);

      // Resolve the handle on its own sync before queueing anything onto it. An
      // unresolvable getItemAt is what poisons a context in the first place, and
      // finding that out now costs one round trip instead of the whole rebuild.
      slide.load("id");
      if (!(await safeSync(ctx, `slide ${slideNum} rescue handle`))) return;

      let cursor = layout.titleTop;

      if (content.title) {
        const titleHeight = 52;
        addText(slide, content.title, {
          left: layout.margin,
          top: cursor,
          width: layout.contentWidth,
          height: titleHeight
        }, {
          font: theme.headingFont,
          size: fitFontSize(content.title, layout.contentWidth, titleHeight, { max: 30, min: 18 }),
          bold: true,
          color: content.color || theme.titleColor
        });
        cursor += titleHeight + 10;
      }

      if (content.subtitle) {
        const subtitleHeight = 34;
        addText(slide, content.subtitle, {
          left: layout.margin,
          top: cursor,
          width: layout.contentWidth,
          height: subtitleHeight
        }, {
          font: theme.bodyFont,
          size: fitFontSize(content.subtitle, layout.contentWidth, subtitleHeight, { max: 14.5, min: 10 }),
          color: theme.mutedColor
        });
        cursor += subtitleHeight + 8;
      }

      if (content.text) {
        const bodyHeight = Math.max(90, layout.contentBottom - cursor);
        addText(slide, content.text, {
          left: layout.margin,
          top: cursor,
          width: layout.contentWidth,
          height: bodyHeight
        }, {
          font: theme.bodyFont,
          size: fitFontSize(content.text, layout.contentWidth, bodyHeight, { max: 16, min: 10 }),
          color: theme.bodyColor
        });
      }

      committed = await safeSync(ctx, `slide ${slideNum} rescue draw`);
    });
  } catch (rescueErr) {
    console.warn(`[PPTBuilder] Slide ${slideNum} fresh-context rescue notice:`, rescueErr);
  }

  return committed;
}

/* ------------------------------------------------------------------ */
/* Cover slide                                                          */
/* ------------------------------------------------------------------ */

/**
 * Slide 1 of a 3+ slide deck is the title and executive summary. It used to be
 * drawn like any content slide: a 30pt title in the corner and the summary as
 * a single bullet under it, with two thirds of the slide left empty. It now
 * gets a cover: a large title in the middle of the slide, the subtitle, and
 * the summary as a lede paragraph with no bullet. Same shape count as before
 * plus one accent band.
 */
const COVER_MAX_SUMMARY_CHARS = 900;

/** The summary as plain prose: bullet glyphs, numbering and ** removed. */
function coverLedeText(text) {
  return String(text || "")
    .split(/\r?\n/)
    .map(l => l.replace(/^\s*(?:[\u2022\u25AA\-*]|\d{1,2}[.)])\s*/, "").replace(/\*\*|__/g, "").trim())
    .filter(Boolean)
    .join("\n");
}

/** True when this slide is the deck's title-and-summary slide. */
function isCoverSlide(slideData, slideNum, deckSize) {
  if (slideNum !== 1 || !(deckSize >= 3) || !slideData) return false;
  if (slideData.tableData && slideData.tableData.rows && slideData.tableData.rows.length > 0) return false;
  if ((slideData.base64Images || []).length > 0 || (slideData.compressedImages || []).length > 0) return false;
  if (slideData.visualType || slideData.imageOnly) return false;
  if (Array.isArray(slideData.structuredCards) && slideData.structuredCards.length >= 2) return false;
  const lede = coverLedeText(slideData.body || slideData.additionalBody || "");
  if (lede.length > COVER_MAX_SUMMARY_CHARS) return false;
  return lede.split("\n").length <= 3;
}

/** Cover geometry: centred title block with room for the lede under it. */
function coverPlan(layout, title, subtitle, lede) {
  const titleWidth = Math.round(layout.contentWidth * 0.88);
  const titleSize = fitFontSize(title, titleWidth, 112, { max: 44, min: 28 });
  const titleHeight = Math.ceil(textBlockHeight(title, titleWidth, titleSize)) + 8;
  const subtitleSize = subtitle ? fitFontSize(subtitle, titleWidth, 52, { max: 20, min: 13 }) : 0;
  const subtitleHeight = subtitle ? Math.ceil(textBlockHeight(subtitle, titleWidth, subtitleSize)) + 6 : 0;
  const ledeWidth = Math.round(layout.contentWidth * 0.8);
  const ledeSize = lede ? fitFontSize(lede, ledeWidth, 150, { max: 17, min: 12 }) : 0;
  const ledeHeight = lede ? Math.ceil(textBlockHeight(lede, ledeWidth, ledeSize)) + 8 : 0;
  const block = titleHeight + 22 + (subtitle ? subtitleHeight + 18 : 0) + ledeHeight;
  const top = Math.max(layout.titleTop, Math.round((layout.H - block) / 2) - 12);
  return { titleWidth, titleSize, titleHeight, subtitleSize, subtitleHeight, ledeWidth, ledeSize, ledeHeight, top };
}

async function createSingleSlide(slideData, slideNum, layoutOptions, options, theme) {
  const layout = getLayout(theme);

  const cleanTitle = (slideData.title || `Slide ${slideNum}`).replace(/\*\*/g, "").trim();
  const subtitle = slideData.subtitle || "";
  const takeaway = (slideData.takeaway || "").trim();
  const color = slideData.color || null;
  const bodyTextContent = (slideData.body || "").trim();
  const hasMeaningfulBody = isSubstantiveSlideBody(bodyTextContent);

  const tableData = slideData.tableData || null;
  const hasTable = Boolean(tableData && tableData.rows && tableData.rows.length > 0);

  const imagesToInsert = (slideData.compressedImages && slideData.compressedImages.length > 0)
    ? slideData.compressedImages
    : (slideData.base64Images || []);
  const hasImages = imagesToInsert.length > 0;
  const imageMeta = slideData.compressedImageMeta || [];
  const primaryAspect = (imageMeta[0] && imageMeta[0].aspect) || null;

  // When a slide carries BOTH a markdown table and a card layout, the chat
  // preview shows the cards -- it tests the archetypes first and only falls
  // back to the table. The builder used to test the table first, so the user
  // was promised bubbles and handed a grid. The preview wins.
  //
  // A slide that also carries a chart is left alone: the table+chart pairing
  // uses the space well, and giving the chart up to make room for cards loses
  // more than it gains.
  const isAutoDesign = !options.layoutChoice || options.layoutChoice === "auto";
  const preferArchetype = isAutoDesign && !hasImages && hasExplicitCardLayout(slideData);
  const useTable = hasTable && !preferArchetype;

  const deckSize = slideData.deckSize || 0;
  const coverSlide = isCoverSlide(slideData, slideNum, deckSize);
  const cover = coverSlide
    ? coverPlan(getLayout(theme), cleanTitle, subtitle, coverLedeText(bodyTextContent || supportingNarrativeFor(slideData)))
    : null;

  logToPPTConsole(`Slide ${slideNum}: Preparing "${cleanTitle.substring(0, 32)}..."${coverSlide ? " (cover)" : ""}`);
  if (hasTable && preferArchetype) {
    logToPPTConsole(
      `Slide ${slideNum}: has both a table and a "${slideData.visualType}" layout - ` +
      `drawing the layout, to match the preview.`
    );
  }

  // Hoisted out of the PowerPoint.run below: the fresh-context rebuild runs
  // after that context has closed and still needs to know whether this slide
  // was ever supposed to carry a heading.
  const isImageOnlySlide = slideData.imageOnly || (!cleanTitle && hasImages && !hasMeaningfulBody && !hasTable);

  let imageInserted = false;

  // Index of a slide that landed nothing because the host refused its context
  // wholesale. The rebuild has to wait for that context to close - see
  // rescueSlideInFreshContext.
  let deferredRescueIndex = -1;
  // Index of the slide this call appended. The picture fallback targets it by
  // index - never "the last slide" and never "whatever is selected".
  let builtSlideIndex = -1;

  await PowerPoint.run(async (context) => {
    // ---- PHASE 1: commit the slide canvas -----------------------------
    // Isolated + synced so that a later styling failure can never discard it.
    const newSlide = await appendSlide(context, layoutOptions);

    // Stamp the slide as ours before anything else can fail. The deck-theme
    // sampler reads the deck's first slides to learn its colours, and after one
    // run those are our slides - untagged, it would be sampling its own output.
    // Queue the prompt-placeholder read alongside the tag write: the tag sync
    // below fulfils both, so clearing the prompts costs no extra round trip on
    // the normal path. See stripQueuedPrompts() for why the prompts have to go
    // before we draw rather than in the end-of-deck sweep.
    const appendedIndex = peekAppendedSlideIndex();
    builtSlideIndex = appendedIndex;
    const promptShapes = queuePromptRead(context, appendedIndex);
    let promptReadSynced = false;
    // Whether the host refused the sync carrying this slide's prompt-placeholder
    // read. Reported in the bare-slide check so a blank slide can be traced back
    // to the first thing that went wrong on it.
    let firstSyncRefused = false;

    if (newSlide.tags) {
      try {
        newSlide.tags.add(ADDIN_SLIDE_TAG, BUILDER_REVISION);
        // Record the palette so the next run can match this slide rather than
        // re-deriving from a deck whose only content is our own output - which
        // the sampler is built not to look at. The fingerprint comes off the
        // theme because buildTokens computed it from the sample; recomputing it
        // here from resolved token values would hash different inputs.
        if (theme) {
          try {
            newSlide.tags.add(ADDIN_PALETTE_TAG, encodePalette(theme, theme.themeFingerprint || ""));
          } catch (paletteErr) {
            console.warn("[PPTBuilder] Could not record the palette:", paletteErr);
          }
        }
        if (slideData.notes) {
          const cleanNotes = String(slideData.notes).replace(/[\r\n]+/g, " | ").trim();
          newSlide.tags.add("SpeakerNotes", cleanNotes.substring(0, 250));
        }
        await meteredSync(context);
        promptReadSynced = true;
      } catch (tErr) {
        // This sync carries two things: the slide tag, and the prompt-placeholder
        // read queued just above - the only getItemAt() this context makes. When
        // the host refuses a slide's whole context it answers "InvalidParam
        // passed to GetItem(id)", so whether the refusals begin HERE is the
        // difference between a root cause and a guess. Until now this only
        // reached console.warn, which is invisible in a screenshot of the task
        // pane and so never made it into a bug report.
        firstSyncRefused = true;
        console.warn("[PPTBuilder] Notice setting slide tag:", tErr);
        logToPPTConsole(
          `Slide ${slideNum}: the host refused this slide's first sync, which carried the ` +
          `slide tag and the prompt-placeholder read for index ${appendedIndex}. ` +
          `${(tErr && (tErr.message || tErr.code)) || tErr}`,
          "warn"
        );
      }
    }

    const promptCleanup = await stripQueuedPrompts(context, promptShapes, promptReadSynced, `Slide ${slideNum}`);

    // The canvas is committed; every shape from here on has to survive a sync
    // to exist. These two marks turn "the host threw this whole slide away"
    // into something the builder can test, instead of something only the user
    // can see.
    const shapesAttemptedAtStart = SHAPES_ATTEMPTED;
    const shapesCommittedAtStart = SHAPES_COMMITTED;
    const shapesDiscardedAtStart = SHAPES_DISCARDED;
    const shapesRedrawnAtStart = SHAPES_REDRAWN;

    // ---- PHASE 2: title + subtitle ------------------------------------
    // Queued through a function so the whole header can be re-queued if the
    // host refuses the batch. It used to be a bare context.sync() in a
    // try/catch that only reached console.warn, so a refusal silently cost the
    // title AND the subtitle, and the log said nothing at all - which is
    // exactly what produced the untitled slides in the user's deck.
    const drawHeader = () => {
      if (isImageOnlySlide || !cleanTitle) return layout.titleTop;

      if (cover) {
        // A band down the left edge marks the cover without a picture.
        drawPanel(newSlide, { left: 0, top: 0, width: 12, height: layout.H }, theme.accentColor, null, "rectangle");
        addText(newSlide, cleanTitle, {
          left: layout.margin,
          top: cover.top,
          width: cover.titleWidth,
          height: cover.titleHeight
        }, {
          font: theme.headingFont,
          size: cover.titleSize,
          bold: true,
          color: color || theme.titleColor
        });
        let coverCursor = cover.top + cover.titleHeight;
        drawAccentRule(newSlide, layout, theme, coverCursor + 4);
        coverCursor += 22;
        if (subtitle) {
          addText(newSlide, subtitle, {
            left: layout.margin,
            top: coverCursor,
            width: cover.titleWidth,
            height: cover.subtitleHeight
          }, {
            font: theme.bodyFont,
            size: cover.subtitleSize,
            color: theme.mutedColor
          });
          coverCursor += cover.subtitleHeight + 18;
        }
        return coverCursor;
      }

      const titleHeight = cleanTitle.length <= 32 ? 46 : cleanTitle.length <= 52 ? 56 : 64;
      const titleSize = fitFontSize(cleanTitle, layout.contentWidth, titleHeight, { max: 30, min: 18 });

      addText(newSlide, cleanTitle, {
        left: layout.margin,
        top: layout.titleTop,
        width: layout.contentWidth,
        height: titleHeight
      }, {
        font: theme.headingFont,
        size: titleSize,
        bold: true,
        color: color || theme.titleColor
      });

      let cursor = layout.titleTop + titleHeight;
      drawAccentRule(newSlide, layout, theme, cursor + 2);
      cursor += 12;

      if (subtitle) {
        const subtitleHeight = 34;
        addText(newSlide, subtitle, {
          left: layout.margin,
          top: cursor,
          width: layout.contentWidth,
          height: subtitleHeight
        }, {
          font: theme.bodyFont,
          size: fitFontSize(subtitle, layout.contentWidth, subtitleHeight, { max: 14.5, min: 10 }),
          color: theme.mutedColor
        });
        cursor += subtitleHeight + 2;
      }

      return cursor + 6;
    };

    const contentTop = drawHeader();
    let headerCommitted = await safeSync(context, `slide ${slideNum} header`);
    if (!headerCommitted) {
      // A refused batch is discarded whole, so the title and subtitle have to
      // be queued again. Re-queueing cannot double them up on a host that
      // genuinely dropped the batch, and on one that committed it anyway the
      // duplicate is drawn at identical geometry with identical text - which
      // is invisible, and far better than a slide with no title.
      //
      // Straight back in, with no backoff. The 500ms wait below was picked
      // when a refusal was thought to open a window seconds wide; then the
      // 13:08 run refused six headers out of ten, recovered every one, and
      // spent 7.20s of a 63.76s build waiting on purpose. Whether any of that
      // wait is load-bearing has never been tested, so this tests it: the
      // immediate attempt goes first and the log says which one won.
      redrawCounted(drawHeader);
      headerCommitted = await safeSync(context, `slide ${slideNum} header (immediate retry)`);

      if (headerCommitted) {
        logToPPTConsole(
          `Slide ${slideNum}: the header was refused and committed straight away on ` +
          `a second attempt - no backoff needed.`,
          "warn"
        );
      } else {
        await delay(500);
        redrawCounted(drawHeader);
        headerCommitted = await safeSync(context, `slide ${slideNum} header (retry after 500ms)`);
        logToPPTConsole(
          headerCommitted
            ? `Slide ${slideNum}: the header needed the 500ms backoff before it committed.`
            : `Slide ${slideNum}: the header was refused twice - trying again after the content.`,
          "warn"
        );
      }
      // Not given up on yet - see the late attempt after the content below.
    }

    // Deliberately here: after the header has committed and before any content
    // is drawn. That ordering is the whole mechanism - a write has to follow the
    // delete or the web client goes on painting the prompts over the slide.
    // Costs a sync only on the slides whose first attempt was refused.
    await retryQueuedPromptDeletes(context, promptCleanup.pending, `Slide ${slideNum}`);

    // Reserve room at the bottom for the takeaway banner when present.
    const takeawayReserve = takeaway ? 58 : 0;
    const contentBottom = layout.contentBottom - takeawayReserve;
    const contentHeight = Math.max(90, contentBottom - contentTop);

    // ---- PHASE 3: main content ----------------------------------------
    let bodyBox = null;
    let bodyClean = "";
    let bodyParagraphs = null;
    // Set when the body was drawn by runs (cover lede, styled list). Those
    // boxes are finished; the cosmetic pass below would only undo them.
    let bodyRunStyled = false;
    // Tracks whether anything at all was drawn into the content area. The
    // recovery paths below must never draw a second time on top of it.
    let contentCommitted = false;

    try {
      if (isImageOnlySlide && hasImages) {
        imageInserted = await insertPictureOnSlide(newSlide, imagesToInsert[0], {
          left: layout.margin + 10,
          top: layout.titleTop,
          width: layout.contentWidth - 20,
          height: layout.contentBottom - layout.titleTop
        }, slideNum, primaryAspect, context);
        if (imageInserted) contentCommitted = true;
      } else if (hasTable && hasImages) {
        const narrative = tableNarrativeFor(slideData, tableData, `Slide ${slideNum}`);
        // Keep a slice of the slide for the narrative so the visuals do not
        // expand into the space the bullets need.
        const visualBottom = narrative
          ? contentBottom - Math.round(Math.min(120, contentHeight * 0.3))
          : contentBottom;
        const visualHeight = Math.max(120, visualBottom - contentTop);
        // A 4+ column table squeezed into half the slide leaves ~55pt per
        // column, which wraps every heading onto three lines. Wide tables take
        // more of the width and the chart gives it up.
        const tableCols = (tableData.headers || []).length || (tableData.rows[0] || []).length;
        // Share the width between the two instead of handing the table a fixed
        // slice. The picture is letterboxed to its native aspect, so a narrow
        // slot wasted the slot's own height: a 900x560 chart dropped into a
        // 317x266 box rendered 317x197 and left 69pt empty underneath it. Width
        // was the binding constraint, never height.
        //
        // Size the chart off the height it is allowed to use, then let the table
        // keep whatever is left above a readable floor. On a 4-column table that
        // is 395x246 for the chart (was 317x197) and 447pt for the table.
        const visualGutter = 18;
        const minTableWidth = Math.round(layout.contentWidth * (tableCols >= 4 ? 0.52 : 0.44));
        const maxChartWidth = layout.contentWidth - minTableWidth - visualGutter;
        const idealChartWidth = (primaryAspect && Number.isFinite(primaryAspect))
          ? Math.round(visualHeight * primaryAspect)
          : Math.round(layout.contentWidth * 0.40);
        const chartWidth = Math.max(
          Math.round(layout.contentWidth * 0.30),
          Math.min(idealChartWidth, maxChartWidth)
        );
        const tableWidth = layout.contentWidth - chartWidth - visualGutter;
        const tableHeight = await populateSlideTable(context, newSlide, tableData, slideNum, theme, {
          left: layout.margin,
          top: contentTop,
          width: tableWidth,
          maxHeight: visualHeight
        }, { deferFinish: true });
        // Match the box to what the picture will actually occupy, so the
        // narrative below starts at the true bottom edge rather than under a
        // band of letterbox.
        const chartHeight = (primaryAspect && Number.isFinite(primaryAspect))
          ? Math.min(visualHeight, Math.round(chartWidth / primaryAspect))
          : Math.min(visualHeight, 330);
        imageInserted = await insertPictureOnSlide(newSlide, imagesToInsert[0], {
          left: layout.margin + tableWidth + visualGutter,
          top: contentTop + 4,
          width: chartWidth,
          height: chartHeight
        }, slideNum, primaryAspect, context);

        // If the table could not be drawn, hand its space to the narrative
        // rather than leaving half the slide empty.
        const usedBottom = tableHeight
          ? Math.max(contentTop + tableHeight, contentTop + 4 + chartHeight)
          : contentTop + 4 + chartHeight;
        const narrativeBounds = tableHeight
          ? { left: layout.margin, top: usedBottom + TABLE_NOTES_GAP_PT, width: layout.contentWidth, bottom: contentBottom }
          : { left: layout.margin, top: contentTop, width: tableWidth, bottom: contentBottom };
        const narrativeHeight = await renderSupportingNarrative(
          context, newSlide, narrative, theme, narrativeBounds, `slide ${slideNum} narrative`
        );
        if (imageInserted || tableHeight || narrativeHeight) contentCommitted = true;
      } else if (useTable) {
        const narrative = tableNarrativeFor(slideData, tableData, `Slide ${slideNum}`);
        const tableCeiling = narrative
          ? contentHeight - Math.round(Math.min(150, contentHeight * 0.36))
          : contentHeight;
        const tableHeight = await populateSlideTable(context, newSlide, tableData, slideNum, theme, {
          left: layout.margin,
          top: contentTop,
          width: layout.contentWidth,
          maxHeight: Math.max(110, tableCeiling)
        }, { deferFinish: true });
        const narrativeHeight = await renderSupportingNarrative(context, newSlide, narrative, theme, {
          left: layout.margin,
          top: tableHeight ? contentTop + tableHeight + TABLE_NOTES_GAP_PT : contentTop,
          width: layout.contentWidth,
          bottom: contentBottom
        }, `slide ${slideNum} narrative`);
        if (tableHeight || narrativeHeight) contentCommitted = true;
      } else if (hasImages && hasMeaningfulBody) {
        const textWidth = Math.round(layout.contentWidth * 0.48);
        const parsed = parseMarkdownFormatting(bodyTextContent);
        bodyClean = parsed.cleanText;
        bodyParagraphs = parsed.parsedParagraphs;
        bodyBox = await commitTextBlock(context, newSlide, bodyClean, {
          left: layout.margin,
          top: contentTop,
          width: textWidth,
          height: contentHeight
        }, {
          font: theme.bodyFont,
          size: fitFontSize(bodyClean, textWidth, contentHeight, { max: 16, min: 10 }),
          color: theme.bodyColor
        }, `slide ${slideNum} body`);
        imageInserted = await insertPictureOnSlide(newSlide, imagesToInsert[0], {
          left: layout.margin + textWidth + 24,
          top: contentTop + 6,
          width: layout.contentWidth - textWidth - 24,
          height: Math.min(contentHeight, 340)
        }, slideNum, primaryAspect, context);
        if (bodyBox || imageInserted) contentCommitted = true;
      } else if (hasImages) {
        const narrative = supportingNarrativeFor(slideData);
        const imgWidth = Math.round(layout.contentWidth * 0.72);
        const imgHeight = Math.min(
          narrative ? Math.max(180, contentHeight - 110) : contentHeight,
          360
        );
        imageInserted = await insertPictureOnSlide(newSlide, imagesToInsert[0], {
          left: Math.round((layout.W - imgWidth) / 2),
          top: contentTop + 6,
          width: imgWidth,
          height: imgHeight
        }, slideNum, primaryAspect, context);
        const narrativeHeight = await renderSupportingNarrative(context, newSlide, narrative, theme, {
          left: layout.margin,
          top: contentTop + 6 + imgHeight + 12,
          width: layout.contentWidth,
          bottom: contentBottom
        }, `slide ${slideNum} narrative`);
        if (imageInserted || narrativeHeight) contentCommitted = true;
      } else {
        let renderedArchetype = false;

        if (isAutoDesign && (slideData.visualType || slideData.structuredCards)) {
          const beforeCards = await countSlideShapes(context, newSlide);
          const syncOkBefore = SAFE_SYNC_OK;

          // Extracted so the whole layout can be drawn again if the host
          // refused every single card. On slide 3 of the user's deck all three
          // "kpi card N" syncs were refused, the renderer reported success
          // anyway, and the slide came out with nothing but its heading.
          const drawArchetype = async () => {
            if (slideData.visualType === "metric_grid_3col" && slideData.visualData) {
              return await renderKpiHeroCards(context, newSlide, slideData.visualData, contentTop, theme, layout, contentBottom);
            } else if (slideData.visualType === "before_after" && slideData.visualData) {
              return await renderBeforeAfterCards(context, newSlide, slideData.visualData, contentTop, theme, layout, contentBottom);
            } else if (slideData.visualType === "three_column_cards" && slideData.visualData) {
              return await renderMultiColumnCards(context, newSlide, slideData.visualData.cards, contentTop, theme, layout, contentBottom);
            } else if (slideData.visualType === "process_timeline" && slideData.visualData) {
              return await renderProcessTimeline(context, newSlide, slideData.visualData.steps, contentTop, theme, layout, contentBottom);
            } else if (slideData.visualType === "executive_summary" && slideData.visualData) {
              return await renderExecutiveSummaryCard(context, newSlide, slideData.visualData, contentTop, theme, layout, contentBottom);
            } else if (slideData.structuredCards && slideData.structuredCards.length >= 2) {
              return await renderMultiColumnCards(context, newSlide, slideData.structuredCards, contentTop, theme, layout, contentBottom);
            }
            return false;
          };

          try {
            renderedArchetype = await drawArchetype();

            // The ONE negative signal on this host that can be trusted: the
            // renderers commit through safeSync, so if the tally has not moved
            // then not a single batch was accepted. (The shape count is no use
            // here - it answers "?" in exactly these moments.) Wait for the
            // refusal window to pass and draw the layout again.
            if (renderedArchetype && SAFE_SYNC_OK === syncOkBefore) {
              logToPPTConsole(
                `Slide ${slideNum}: the host refused every card - waiting for it to settle ` +
                `and drawing the layout again.`,
                "warn"
              );
              await delay(700);
              renderedArchetype = await drawArchetype();
              if (SAFE_SYNC_OK === syncOkBefore) {
                renderedArchetype = false;
                logToPPTConsole(
                  `Slide ${slideNum}: the card layout was refused twice - falling back to text.`,
                  true
                );
              }
            }
          } catch (cardErr) {
            console.warn("[PPTBuilder] Archetype card render notice:", cardErr);
            // Card renderers commit one card at a time, so a failure part-way
            // through still leaves real shapes on the slide. Printing the
            // bullets over them is far worse than losing the tail of the
            // layout, so only fall back if nothing was committed.
            const afterCards = await countSlideShapes(context, newSlide);
            renderedArchetype = Boolean(
              beforeCards.ok && afterCards.ok && afterCards.count > beforeCards.count
            );
            if (renderedArchetype) {
              logToPPTConsole(
                `Slide ${slideNum}: card layout stopped early - keeping what rendered rather than overprinting bullets.`,
                true
              );
            }
          }

          // NO CALLER-SIDE OVERRIDE OF THE RENDERER'S VERDICT.
          //
          // Two previous rounds tried to second-guess the card renderers with
          // the slide's shape count, and both printed bullets on top of cards
          // that had in fact rendered. The count is not evidence on this host:
          // slide 7 read `2 -> 2` across the renderer and `23` seconds later,
          // and slide 5 reported 2 shapes for a 3-shape header. It is wrong in
          // both directions, so it cannot veto anything.
          //
          // The count and the sync tally are still logged below, purely as
          // diagnostics. Overprinting real cards is the worse failure, so the
          // renderer is believed.
          void syncOkBefore;

          const archetypeItems =
            (slideData.visualData && (slideData.visualData.cards || slideData.visualData.steps)) ||
            slideData.structuredCards || [];
          // Was a second countSlideShapes, i.e. a host round trip per card slide
          // spent on a diagnostic that printed "shapes 0 -> 0" while 27 shapes
          // committed. The landed delta is accurate, and free.
          logToPPTConsole(
            `Slide ${slideNum}: layout "${slideData.visualType || "structuredCards"}" with ` +
            `${archetypeItems.length} item(s) -> ${renderedArchetype ? "rendered" : "declined"}; ` +
            `${SHAPES_COMMITTED - shapesCommittedAtStart} shape(s) committed so far.`
          );
        }

        // Some parser strategies park the narrative on `additionalBody`; the
        // chat preview shows it, so the slide must too.
        const plainBody = hasMeaningfulBody ? bodyTextContent : supportingNarrativeFor(slideData);
        let styledBullets = null;
        if (!renderedArchetype && cover && isSubstantiveSlideBody(plainBody)) {
          const lede = coverLedeText(plainBody);
          const ledeBox = await commitTextBlock(context, newSlide, lede, {
            left: layout.margin,
            top: contentTop,
            width: cover.ledeWidth,
            height: Math.max(cover.ledeHeight, Math.min(contentHeight, cover.ledeHeight + 40))
          }, {
            font: theme.bodyFont,
            size: cover.ledeSize,
            color: theme.bodyColor
          }, `slide ${slideNum} cover summary`);
          if (ledeBox) {
            bodyBox = ledeBox;
            bodyClean = lede;
            bodyParagraphs = null;
            styledBullets = { box: ledeBox, text: lede, variant: "cover" };
            bodyRunStyled = true;
            logToPPTConsole(`Slide ${slideNum}: drawn as the cover (title ${cover.titleSize}pt, summary ${cover.ledeSize}pt, no bullet).`);
          }
        }
        // Slide 1 of a 1-2 slide deck is content, not a cover, so it is styled too.
        if (!styledBullets && !renderedArchetype && isSubstantiveSlideBody(plainBody) && (slideNum > 1 || deckSize < 3)) {
          // Slide 1 is the title and summary; every other bullet list gets the
          // styled treatment when it is a flat list that fits.
          const parsedList = parseMarkdownFormatting(plainBody);
          styledBullets = await renderStyledBullets(context, newSlide, parsedList.parsedParagraphs, {
            left: layout.margin,
            top: contentTop,
            width: layout.contentWidth,
            height: contentHeight
          }, theme, `slide ${slideNum} styled bullets`);
          if (styledBullets) {
            bodyBox = styledBullets.box;
            bodyClean = styledBullets.text;
            bodyParagraphs = null;
            bodyRunStyled = true;
            logToPPTConsole(`Slide ${slideNum}: bullets styled (${styledBullets.variant}).`);
          }
        }
        if (!renderedArchetype && !styledBullets && isSubstantiveSlideBody(plainBody)) {
          const parsed = parseMarkdownFormatting(plainBody);
          bodyClean = parsed.cleanText;
          bodyParagraphs = parsed.parsedParagraphs;
          bodyBox = await commitTextBlock(context, newSlide, bodyClean, {
            left: layout.margin,
            top: contentTop,
            width: layout.contentWidth,
            height: contentHeight
          }, {
            font: theme.bodyFont,
            size: fitFontSize(bodyClean, layout.contentWidth, contentHeight, { max: 17, min: 10 }),
            color: theme.bodyColor
          }, `slide ${slideNum} body`);
        }

        if (!renderedArchetype && !bodyBox && !styledBullets) {
          logToPPTConsole(
            `Slide ${slideNum}: no card layout and no body text was drawn ` +
            `(body ${isSubstantiveSlideBody(plainBody) ? "substantive" : "empty/too short"}, ` +
            `${String(plainBody || "").length} chars).`,
            true
          );
        }

        if (renderedArchetype || bodyBox || styledBullets) contentCommitted = true;
      }

      // The heading gets one more chance here. Drawing the content takes
      // seconds, and the host's refusal windows are far shorter than that, so
      // a header that was refused twice in a row at the top of the slide will
      // very often commit now. The geometry is computed rather than measured,
      // so a late header lands exactly where the early one would have.
      if (!headerCommitted && !isImageOnlySlide && cleanTitle) {
        redrawCounted(drawHeader);
        headerCommitted = await safeSync(context, `slide ${slideNum} header (late)`);
        logToPPTConsole(
          headerCommitted
            ? `Slide ${slideNum}: the heading was refused twice but committed after the content.`
            : `Slide ${slideNum}: the host refused the title and subtitle three times - this ` +
              `slide has no heading.`,
          headerCommitted ? "warn" : true
        );
      }

      if (takeaway) {
        renderTakeawayBanner(newSlide, takeaway, theme, layout);
      }

      await meteredSync(context);
    } catch (contentErr) {
      // Degrading to plain text needs a host that answers. After an abort every
      // sync throws before issuing anything, so announcing a fallback here would
      // promise something that cannot happen - and absorbing the error would let
      // this slide be counted as created when it holds only its header.
      if (contentErr && contentErr.isBuildAbort) throw contentErr;
      // The slide itself is already committed, so degrade to a plain text body
      // rather than losing the slide entirely.
      console.warn(`[PPTBuilder] Slide ${slideNum} content notice, applying plain fallback:`, contentErr);
      logToPPTConsole(`Slide ${slideNum}: content styling failed (${contentErr.message || contentErr}), using plain layout.`, true);
      bodyBox = null;
      try {
        // Only degrade to plain text if the content area is genuinely empty.
        // Cards and tables commit through their own syncs, so a late failure
        // here would otherwise print the bullets on top of them.
        if (hasMeaningfulBody && !contentCommitted) {
          const parsed = parseMarkdownFormatting(bodyTextContent);
          addText(newSlide, parsed.cleanText, {
            left: layout.margin,
            top: contentTop,
            width: layout.contentWidth,
            height: contentHeight
          }, { size: 16 });
          await meteredSync(context);
        }
      } catch (fallbackErr) {
        console.warn(`[PPTBuilder] Slide ${slideNum} plain fallback notice:`, fallbackErr);
      }
    }

    // ---- PHASE 3b: guarantee the slide is not blank ---------------------
    // A rejected sync discards every shape queued in that batch. Verify what
    // actually landed and, if only the header survived, write the content out
    // as plain text so the user never gets an empty slide.
    {
      // Read before the count below, which syncs: a host coming back out of a
      // refusal window can serve that one read even though it threw away every
      // shape the slide drew.
      const shapesDrawn = SHAPES_ATTEMPTED - shapesAttemptedAtStart;
      const shapesLanded = SHAPES_COMMITTED - shapesCommittedAtStart;
      const shapesDiscarded = SHAPES_DISCARDED - shapesDiscardedAtStart;
      const shapesRedrawn = SHAPES_REDRAWN - shapesRedrawnAtStart;
      // What a refusal destroyed and nobody queued again. Everything else in
      // this check is about a slide that landed NOTHING; this is the partial
      // loss that has been invisible all along.
      const shapesLost = Math.max(0, shapesDiscarded - shapesRedrawn);

      // Free: whatever a previous read happened to leave cached, or "unknown".
      // This used to spend a whole round trip per slide on a number that turned
      // out to be noise - 91% of a build is host syncs at ~428ms each, and the
      // decisions below are driven by `landed`, which costs nothing.
      const { ok: countOk, count: shapeCount } = knownShapeCount(newSlide);

      // Header shapes: title, accent rule, subtitle.
      const headerShapeCount = (cleanTitle ? 2 : 0) + (subtitle ? 1 : 0);

      // A count lower than the header we just drew means the host is not
      // reporting shapes correctly, so the reading is unusable.
      const countTrustworthy = countOk && shapeCount >= headerShapeCount;

      // The shape count lags badly on this host. Slide 1 of a real deck
      // reported shapes=3 against header=3 immediately after the body text box
      // had committed cleanly, and a count-authoritative rule then redrew the
      // same paragraph on top of itself.
      //
      // Neither signal can be trusted on its own, so BOTH must say the content
      // area is empty before anything is redrawn:
      //   - contentCommitted        the builder's own record that a sync for
      //                             the body, table or card layout succeeded.
      //                             A text batch that syncs cleanly has landed;
      //                             a refused text batch genuinely loses its
      //                             shapes. This signal is reliable for text.
      //   - countSaysContentPresent more shapes than the header we drew. This
      //                             covers the opposite failure, where the host
      //                             reports a rejection but keeps the shapes.
      // A stale count can now only ever suppress a rescue, never cause one.
      const countSaysContentPresent = countTrustworthy && shapeCount > headerShapeCount;

      // Still gated on !imageInserted. This branch prints the body into the
      // content area on a slide that may already hold a picture, so it has to
      // stay off image slides - which is why the whole check used to be gated,
      // leaving image slides with no protection and no diagnostics at all.
      const needsRescue = !imageInserted && !contentCommitted && !countSaysContentPresent;

      // The slide drew shapes and not one of them survived a sync. That is not
      // a judgement call the way the two signals above are: a refused batch is
      // discarded whole, so nothing reached the document - whatever
      // contentCommitted claims, since that flag only records that a renderer
      // QUEUED something. It also means a redraw cannot duplicate anything,
      // which is the risk every other branch here is written around.
      // Deliberately NOT ANDed with the shape count any more. The count reads
      // 0 on slides holding 27 shapes and non-zero on slides holding almost
      // none, so letting it veto here would leave a genuinely empty slide blank
      // on the strength of a number known to be wrong. Zero committed shapes is
      // provable; the count is not.
      //
      // Safe on an image slide too, unlike the text rescue above: nothing
      // committed means the picture did not land either, so there is nothing on
      // the slide for the rebuild to cover up.
      const slideLandedNothing = shapesDrawn > 0 && shapesLanded === 0;

      logToPPTConsole(
        `Slide ${slideNum}: bare-slide check - shapes=${countOk ? shapeCount : "?"} (advisory; ` +
        `this host miscounts), header=${headerShapeCount}, committed=${contentCommitted}, ` +
        `landed=${shapesLanded}/${shapesDrawn}` +
        (shapesDiscarded
          ? ` (${shapesDiscarded} discarded by a refused batch, ${shapesRedrawn} re-drawn)`
          : "") + `, ` +
        `firstSync=${firstSyncRefused ? "REFUSED" : "ok"}, ` +
        `${imageInserted ? "image=yes, " : ""}` +
        `countSeesContent=${countSaysContentPresent ? "yes" : "no"}, ` +
        `rescue=${slideLandedNothing ? "fresh context" : needsRescue ? "yes" : "no"}.`
      );

      // Deliberately quotes the two raw tallies instead of claiming a count of
      // missing shapes. Once a renderer falls back to substitute content the
      // difference stops being a headcount: a persistent refusal drove it to
      // 36 on a slide that was short 18. The fact worth reporting is that a
      // refusal destroyed something nothing put back.
      // Not on a slide that landed nothing: that one is about to be rebuilt in
      // a fresh context and says so in its own words. Two warnings for one
      // slide, the first of which is about to stop being true, is worse than
      // one.
      if (shapesLost > 0 && !slideLandedNothing) {
        logToPPTConsole(
          `Slide ${slideNum}: shapes discarded by a refused batch were never re-drawn ` +
          `(${shapesDiscarded} discarded, ${shapesRedrawn} re-drawn) - this slide is ` +
          `missing content.`,
          "warn"
        );
      }

      if (slideLandedNothing) {
        // Redrawing in THIS context would be pointless - it would join the same
        // refused batch. The rebuild is deferred until the context closes.
        deferredRescueIndex = appendedIndex;
      } else if (needsRescue) {
        const rescueParts = [];
        const narrative = hasMeaningfulBody
          ? parseMarkdownFormatting(bodyTextContent).cleanText
          : parseMarkdownFormatting(supportingNarrativeFor(slideData)).cleanText;
        if (narrative) rescueParts.push(narrative);
        if (hasTable) rescueParts.push(tableToText(tableData));
        if (rescueParts.length === 0) rescueParts.push(flattenVisualData(slideData));
        const rescueText = rescueParts.filter(Boolean).join("\n\n");

        if (rescueText && rescueText.trim()) {
          logToPPTConsole(
            `Slide ${slideNum}: styled content did not commit - restoring content as text.`,
            true
          );
          try {
            bodyBox = addText(newSlide, rescueText, {
              left: layout.margin,
              top: contentTop,
              width: layout.contentWidth,
              height: contentHeight
            }, {
              font: theme.bodyFont,
              size: fitFontSize(rescueText, layout.contentWidth, contentHeight, { max: 16, min: 10 }),
              color: theme.bodyColor
            });
            bodyClean = rescueText;
            bodyParagraphs = null;
            await meteredSync(context);
          } catch (rescueErr) {
            console.warn(`[PPTBuilder] Slide ${slideNum} rescue notice:`, rescueErr);
            bodyBox = null;
          }
        }
      }
    }

    // ---- PHASE 4: cosmetics (safe to fail) -----------------------------
    // Purely cosmetic, and the slide is already correct without it. A throw
    // here used to abort createSingleSlide and lose the whole slide.
    // A styled list went through here too: its spacer lines made the text
    // differ from the glyph-stripped copy, so the old pass rewrote the whole
    // box - every run style gone (all bold) and one bullet on the first line.
    if (bodyBox && !bodyRunStyled) {
      try {
        await applyBoldLeadIns(context, bodyBox, bodyClean, bodyParagraphs);
        await enhanceBulletsCosmetic(context, bodyBox, bodyClean, theme);
      } catch (cosmeticErr) {
        console.warn(`[PPTBuilder] Slide ${slideNum} cosmetic pass notice:`, cosmeticErr);
      }
    }

    if (deferredRescueIndex < 0) {
      logToPPTConsole(`Slide ${slideNum}: ✅ Created (${useTable ? "table" : hasImages ? "visual" : "content"}${takeaway ? " + takeaway" : ""}).`);
    }
  });

  // ---- PHASE 3c: rebuild a slide that landed nothing -------------------
  // Outside the PowerPoint.run on purpose: that context refused every batch,
  // so this is the first moment a write to this slide can succeed.
  if (deferredRescueIndex >= 0) {
    const rescueParts = [];
    const narrative = hasMeaningfulBody
      ? parseMarkdownFormatting(bodyTextContent).cleanText
      : parseMarkdownFormatting(supportingNarrativeFor(slideData)).cleanText;
    if (narrative) rescueParts.push(narrative);
    if (hasTable) rescueParts.push(tableToText(tableData));
    if (rescueParts.length === 0) rescueParts.push(flattenVisualData(slideData));
    const rescueText = rescueParts.filter(Boolean).join("\n\n");

    logToPPTConsole(
      `Slide ${slideNum}: the host refused every write in this slide's context - ` +
      `rebuilding it in a new one.`,
      true
    );

    const rescueContent = {
      title: isImageOnlySlide ? "" : cleanTitle,
      subtitle: isImageOnlySlide ? "" : subtitle,
      color,
      text: rescueText
    };

    let rebuilt = await rescueSlideInFreshContext(deferredRescueIndex, slideNum, rescueContent, theme);
    if (!rebuilt) {
      // The refusal windows are short - hundreds of milliseconds, not seconds.
      // One more attempt a beat later is far cheaper than the blank slide it
      // prevents.
      await delay(800);
      rebuilt = await rescueSlideInFreshContext(deferredRescueIndex, slideNum, rescueContent, theme);
    }

    logToPPTConsole(
      rebuilt
        ? `Slide ${slideNum}: ✅ Created (rebuilt as text - the host refused the styled slide).`
        : `Slide ${slideNum}: ❌ the host refused the rebuild too, so this slide is blank.`,
      rebuilt ? "warn" : true
    );
  }

  // Picture fallback. The build context refused the picture, so:
  //   1. retry on THIS slide in a fresh context,
  //   2. retry with a downscaled copy (oversized payloads are the usual cause),
  //   3. only then the Common API - and only once the host confirms this slide
  //      is the selected one. setSelectedDataAsync inserts into the current
  //      selection; unconfirmed, it put "Insert Image as New Slide" pictures on
  //      the user's existing slide and left the new one blank.
  if (hasImages && !imageInserted && imagesToInsert.length > 0) {
    const imgWidth = (hasMeaningfulBody || hasTable) ? 420 : 620;
    const imgLeft = (hasMeaningfulBody || hasTable) ? 490 : 170;
    const fbLayout = getLayout(theme);
    const retryBounds = isImageOnlySlide
      ? {
          left: fbLayout.margin + 10,
          top: fbLayout.titleTop,
          width: fbLayout.contentWidth - 20,
          height: fbLayout.contentBottom - fbLayout.titleTop
        }
      : { left: imgLeft, top: 95, width: imgWidth, height: 350 };
    const firstImg = imagesToInsert[0];

    let placed = await retryPictureOnSlideIndex(builtSlideIndex, firstImg, retryBounds, slideNum, primaryAspect);
    if (placed) {
      logToPPTConsole(`Slide ${slideNum}: picture placed on the new slide on a second attempt.`, "warn");
    } else {
      const smaller = await shrinkImagePayload(firstImg);
      if (smaller) {
        placed = await retryPictureOnSlideIndex(builtSlideIndex, smaller, retryBounds, slideNum, primaryAspect);
        if (placed) {
          logToPPTConsole(
            `Slide ${slideNum}: picture placed on the new slide after downscaling it ` +
            `(${Math.round((firstImg.length * 0.75) / 1024)} KB -> ${Math.round((smaller.length * 0.75) / 1024)} KB).`,
            "warn"
          );
        }
      }
    }

    if (!placed) {
      const selection = await selectSlideForCommonApi(builtSlideIndex, slideNum);
      if (selection === "verified" || selection === "unverifiable") {
        if (selection === "unverifiable") {
          logToPPTConsole(
            `Slide ${slideNum}: this PowerPoint cannot confirm the selection - inserting the picture via the Common API anyway.`,
            "warn"
          );
        }
        try {
          await insertPictureViaCommonApi(firstImg, retryBounds);
        } catch (commonApiErr) {
          console.warn("[PPTBuilder] Notice inserting picture via Common API:", commonApiErr);
        }
      } else {
        logToPPTConsole(
          `Slide ${slideNum}: PowerPoint refused the picture on the new slide and would not select it, ` +
          `so it was not inserted anywhere else. Select the new slide and use "Insert Image on Slide".`,
          true
        );
      }
    }
  }
}

/**
 * Inserts content (image, visual, table, or text) directly onto the user's currently active slide.
 */
/**
 * Insert-on-current used to start the content 17% down the slide and draw no
 * header at all, on the assumption that the user's slide already had a title.
 * On a blank slide - or one showing only empty "Click to add title" prompts -
 * "Insert Slide 1 on Current" therefore put the bullets on the slide and the
 * title and subtitle the chat preview promised never came over.
 *
 * Runs in its OWN PowerPoint.run so a refused read cannot poison the context
 * that does the drawing. A title placeholder that already holds text is the
 * user's and is kept; empty prompt placeholders are removed so the generated
 * header is not drawn on top of a "Click to add title" ghost.
 *
 * @returns {Promise<{hasTitleText: boolean, removed: number}>}
 */
async function prepareCurrentSlideHeader() {
  const state = { hasTitleText: false, removed: 0 };
  if (!apiAtLeast("1.4")) return state;   // Shape.name and TextFrame.hasText
  try {
    await PowerPoint.run(async (context) => {
      let slide = null;
      try {
        if (context.presentation.getSelectedSlides) {
          const selected = context.presentation.getSelectedSlides();
          selected.load("items");
          await meteredSync(context, "current slide (header check)");
          if (selected.items && selected.items.length > 0) slide = selected.items[0];
        }
      } catch (_) {}
      if (!slide) {
        const countRes = context.presentation.slides.getCount();
        await meteredSync(context, "slide count (header check)");
        if (countRes.value > 0) slide = context.presentation.slides.getItemAt(0);
      }
      if (!slide) return;

      const shapes = slide.shapes;
      shapes.load("items/name");
      await meteredSync(context, "current slide shape names");
      const named = (shapes.items || []).filter((s) => PLACEHOLDER_SHAPE_NAME.test(String(s.name || "")));
      if (named.length === 0) return;

      named.forEach((s) => s.textFrame.load("hasText"));
      await meteredSync(context, "current slide placeholder text");

      const empty = [];
      for (const s of named) {
        const name = String(s.name || "");
        if (s.textFrame && s.textFrame.hasText) {
          if (/title/i.test(name) && !/subtitle/i.test(name)) state.hasTitleText = true;
        } else {
          empty.push(s);
        }
      }
      if (state.hasTitleText) return;   // the user's slide; leave it exactly as it is

      for (const s of empty) {
        try {
          s.delete();
          state.removed += 1;
        } catch (delErr) {
          console.warn("[PPTBuilder] Could not clear a placeholder on the current slide:", delErr);
        }
      }
      if (state.removed > 0) await meteredSync(context, "clear current slide placeholders");
    });
  } catch (err) {
    console.warn("[PPTBuilder] Current slide header check failed; drawing the header anyway:", err);
  }
  return state;
}

/**
 * Title, accent rule and subtitle in the same geometry a generated slide uses,
 * so a slide inserted onto the current one matches "Insert as New Slides".
 * @returns {number} the y position where content may start
 */
function drawCurrentSlideHeader(slide, layout, theme, title, subtitle, color) {
  const titleHeight = title.length <= 32 ? 46 : title.length <= 52 ? 56 : 64;
  addText(slide, title, {
    left: layout.margin,
    top: layout.titleTop,
    width: layout.contentWidth,
    height: titleHeight
  }, {
    font: theme.headingFont,
    size: fitFontSize(title, layout.contentWidth, titleHeight, { max: 30, min: 18 }),
    bold: true,
    color: color || theme.titleColor
  });

  let cursor = layout.titleTop + titleHeight;
  drawAccentRule(slide, layout, theme, cursor + 2);
  cursor += 12;

  if (subtitle) {
    const subtitleHeight = 34;
    addText(slide, subtitle, {
      left: layout.margin,
      top: cursor,
      width: layout.contentWidth,
      height: subtitleHeight
    }, {
      font: theme.bodyFont,
      size: fitFontSize(subtitle, layout.contentWidth, subtitleHeight, { max: 14.5, min: 10 }),
      color: theme.mutedColor
    });
    cursor += subtitleHeight + 2;
  }

  return cursor + 6;
}

export async function insertOnCurrentSlide(slideStructures, options = {}) {
  if (typeof PowerPoint === "undefined" || !PowerPoint.run) {
    throw new Error("PowerPoint Office.js runtime is not available.");
  }

  if (!slideStructures || slideStructures.length === 0) {
    throw new Error("No slide structures found to insert.");
  }

  logToPPTConsole(`=== Starting Insert on Current Slide ===`);

  const theme = await getDeckTheme();
  const layout = getLayout(theme);

  // 1. Pre-process and compress images
  for (const slide of slideStructures) {
    slide.compressedImages = [];
    slide.compressedImageMeta = [];
    const rawImages = slide.base64Images || [];
    for (const rawImg of rawImages) {
      try {
        const meta = {};
        const comp = await compressImageForPowerPoint(rawImg, 1600, 1120, meta);
        if (comp && comp.length > 50) {
          slide.compressedImages.push(comp);
          slide.compressedImageMeta.push(meta);
          logToPPTConsole(
            meta.compressed
              ? `Image re-encoded ${meta.width || "?"}x${meta.height || "?"} -> ${Math.round((comp.length * 0.75) / 1024)} KB.`
              : `Image could NOT be re-encoded (decode timed out) - sending ${Math.round((comp.length * 0.75) / 1024)} KB as-is.`
          );
        }
      } catch (cErr) {
        logToPPTConsole(`Image compression notice: ${cErr.message}`);
      }
    }
  }

  const slideData = slideStructures[0];
  const imagesToInsert = (slideData.compressedImages && slideData.compressedImages.length > 0)
    ? slideData.compressedImages
    : (slideData.base64Images || []);
  const hasImages = imagesToInsert.length > 0;
  const currentMeta = slideData.compressedImageMeta || [];
  const currentAspect = (currentMeta[0] && currentMeta[0].aspect) || null;
  const tableData = slideData.tableData || null;
  const hasTable = Boolean(tableData && tableData.rows && tableData.rows.length > 0);
  const rawBody = (slideData.body || "").trim();
  const hasMeaningfulBody = isSubstantiveSlideBody(rawBody);

  // Same preview-parity rule as the multi-slide path above.
  const isAutoDesign = !options.layoutChoice || options.layoutChoice === "auto";
  const preferArchetype = isAutoDesign && !hasImages && hasExplicitCardLayout(slideData);
  const useTable = hasTable && !preferArchetype;

  // Content starts below whatever title area already exists on the user's slide.
  // Moved down again below the generated header when one is drawn (see below).
  let contentTop = Math.round(layout.H * 0.17);
  const contentBottom = layout.contentBottom;
  let contentHeight = contentBottom - contentTop;

  // Bring the slide's title and subtitle over too - unless the user's slide
  // already has a title of its own, which is never overwritten.
  const headerTitle = String(slideData.title || "").replace(/\*\*/g, "").trim();
  const headerSubtitle = String(slideData.subtitle || "").trim();
  let drawGeneratedHeader = false;
  if (!slideData.imageOnly && headerTitle) {
    const headerState = await prepareCurrentSlideHeader();
    drawGeneratedHeader = !headerState.hasTitleText;
    logToPPTConsole(
      headerState.hasTitleText
        ? `Current slide already has a title - keeping it; inserting the content below it.`
        : `Current slide: drawing title "${headerTitle.substring(0, 60)}"` +
          `${headerSubtitle ? " + subtitle" : ""}` +
          `${headerState.removed ? ` (cleared ${headerState.removed} empty placeholder(s))` : ""}.`
    );
  }

  let imageInserted = false;

  await PowerPoint.run(async (context) => {
    let activeSlide = null;
    try {
      if (context.presentation.getSelectedSlides) {
        const selected = context.presentation.getSelectedSlides();
        selected.load("items");
        await meteredSync(context);
        if (selected.items && selected.items.length > 0) {
          activeSlide = selected.items[0];
        }
      }
    } catch (_) {}

    if (!activeSlide) {
      const slides = context.presentation.slides;
      const countRes = slides.getCount();
      await meteredSync(context);
      if (countRes.value > 0) {
        activeSlide = slides.getItemAt(0);
      }
    }

    if (!activeSlide) {
      throw new Error("No slide available in presentation to insert content onto.");
    }

    if (drawGeneratedHeader) {
      contentTop = drawCurrentSlideHeader(activeSlide, layout, theme, headerTitle, headerSubtitle, slideData.color || null);
      contentHeight = contentBottom - contentTop;
    }

    // 1. Image Only
    if (hasImages && slideData.imageOnly) {
      for (const rawImg of imagesToInsert) {
        const added = await insertPictureOnSlide(activeSlide, rawImg, {
          left: layout.margin + 10,
          top: layout.titleTop,
          width: layout.contentWidth - 20,
          height: contentBottom - layout.titleTop
        }, 1, currentAspect, context);
        if (added) imageInserted = true;
      }
      logToPPTConsole(`Inserted full-size image centered onto current slide.`);
    }
    // 2. Chart/Image + Table
    else if (hasImages && hasTable) {
      const narrative = tableNarrativeFor(slideData, tableData, "Current slide");
      const visualHeight = narrative
        ? Math.max(120, contentHeight - Math.round(Math.min(120, contentHeight * 0.3)))
        : contentHeight;
      const tableCols = (tableData.headers || []).length || (tableData.rows[0] || []).length;
      // Same balanced split as the new-slide path: the chart is sized off the
      // height it may occupy, and the table keeps the remainder above a floor.
      const visualGutter = 18;
      const minTableWidth = Math.round(layout.contentWidth * (tableCols >= 4 ? 0.52 : 0.44));
      const maxChartWidth = layout.contentWidth - minTableWidth - visualGutter;
      const idealChartWidth = (currentAspect && Number.isFinite(currentAspect))
        ? Math.round(visualHeight * currentAspect)
        : Math.round(layout.contentWidth * 0.40);
      const chartWidth = Math.max(
        Math.round(layout.contentWidth * 0.30),
        Math.min(idealChartWidth, maxChartWidth)
      );
      const tableWidth = layout.contentWidth - chartWidth - visualGutter;
      const tableHeight = await populateSlideTable(context, activeSlide, tableData, 1, theme, {
        left: layout.margin,
        top: contentTop,
        width: tableWidth,
        maxHeight: visualHeight
      });
      const chartHeight = (currentAspect && Number.isFinite(currentAspect))
        ? Math.min(visualHeight, Math.round(chartWidth / currentAspect))
        : Math.min(visualHeight, 330);
      imageInserted = await insertPictureOnSlide(activeSlide, imagesToInsert[0], {
        left: layout.margin + tableWidth + visualGutter,
        top: contentTop,
        width: chartWidth,
        height: chartHeight
      }, 1, currentAspect, context);
      await renderSupportingNarrative(context, activeSlide, narrative, theme, {
        left: layout.margin,
        top: contentTop + Math.max(tableHeight || 0, chartHeight) + 14,
        width: layout.contentWidth,
        bottom: contentBottom
      }, "narrative");
      logToPPTConsole(`Inserted side-by-side table and chart onto current slide.`);
    }
    // 3. Images present with bullets or single image
    else if (hasImages) {
      const textWidth = hasMeaningfulBody ? Math.round(layout.contentWidth * 0.48) : 0;
      const imgLeft = hasMeaningfulBody
        ? layout.margin + textWidth + 24
        : Math.round((layout.W - layout.contentWidth * 0.72) / 2);
      const imgWidth = hasMeaningfulBody ? layout.contentWidth - textWidth - 24 : Math.round(layout.contentWidth * 0.72);

      for (const rawImg of imagesToInsert) {
        const added = await insertPictureOnSlide(activeSlide, rawImg, {
          left: imgLeft,
          top: contentTop,
          width: imgWidth,
          height: Math.min(contentHeight, 350)
        }, 1, currentAspect, context);
        if (added) imageInserted = true;
      }
      if (hasMeaningfulBody) {
        const { cleanText } = parseMarkdownFormatting(rawBody);
        addText(activeSlide, cleanText, {
          left: layout.margin,
          top: contentTop,
          width: textWidth,
          height: contentHeight
        }, {
          font: theme.bodyFont,
          size: fitFontSize(cleanText, textWidth, contentHeight, { max: 15, min: 9 }),
          color: theme.bodyColor
        });
      }
      logToPPTConsole(`Inserted image onto current slide.`);
    }
    // 4. Table presentation
    else if (useTable) {
      const narrative = tableNarrativeFor(slideData, tableData, "Current slide");
      const tableCeiling = narrative
        ? Math.max(110, contentHeight - Math.round(Math.min(150, contentHeight * 0.36)))
        : contentHeight;
      const tableHeight = await populateSlideTable(context, activeSlide, tableData, 1, theme, {
        left: layout.margin,
        top: contentTop,
        width: layout.contentWidth,
        maxHeight: tableCeiling
      });
      await renderSupportingNarrative(context, activeSlide, narrative, theme, {
        left: layout.margin,
        top: contentTop + (tableHeight ? tableHeight + TABLE_NOTES_GAP_PT : 0),
        width: layout.contentWidth,
        bottom: contentBottom
      }, "narrative");
      logToPPTConsole(`Inserted table onto current slide.`);
    }
    // 5. Bullets or Modern Card Archetypes
    else {
      let renderedArchetype = false;

      if (isAutoDesign && (slideData.visualType || slideData.structuredCards)) {
        try {
          if (slideData.visualType === "metric_grid_3col" && slideData.visualData) {
            renderedArchetype = await renderKpiHeroCards(context, activeSlide, slideData.visualData, contentTop, theme, layout, contentBottom);
          } else if (slideData.visualType === "before_after" && slideData.visualData) {
            renderedArchetype = await renderBeforeAfterCards(context, activeSlide, slideData.visualData, contentTop, theme, layout, contentBottom);
          } else if (slideData.visualType === "three_column_cards" && slideData.visualData) {
            renderedArchetype = await renderMultiColumnCards(context, activeSlide, slideData.visualData.cards, contentTop, theme, layout, contentBottom);
          } else if (slideData.visualType === "process_timeline" && slideData.visualData) {
            renderedArchetype = await renderProcessTimeline(context, activeSlide, slideData.visualData.steps, contentTop, theme, layout, contentBottom);
          } else if (slideData.visualType === "executive_summary" && slideData.visualData) {
            renderedArchetype = await renderExecutiveSummaryCard(context, activeSlide, slideData.visualData, contentTop, theme, layout, contentBottom);
          } else if (slideData.structuredCards && slideData.structuredCards.length >= 2) {
            renderedArchetype = await renderMultiColumnCards(context, activeSlide, slideData.structuredCards, contentTop, theme, layout, contentBottom);
          }
        } catch (cardErr) {
          console.warn("[PPTBuilder] Archetype card render notice on current slide:", cardErr);
          renderedArchetype = false;
        }
      }

      if (!renderedArchetype && hasMeaningfulBody) {
        const { cleanText } = parseMarkdownFormatting(rawBody);
        addText(activeSlide, cleanText, {
          left: layout.margin,
          top: contentTop,
          width: layout.contentWidth,
          height: contentHeight
        }, {
          font: theme.bodyFont,
          size: fitFontSize(cleanText, layout.contentWidth, contentHeight, { max: 17, min: 10 }),
          color: theme.bodyColor
        });
        logToPPTConsole(`Inserted text box onto current slide.`);
      } else if (renderedArchetype) {
        logToPPTConsole(`Inserted modern card layout onto current slide.`);
      }
    }

    if (slideData.takeaway) {
      try {
        renderTakeawayBanner(activeSlide, slideData.takeaway, theme, layout);
      } catch (_) {}
    }

    await meteredSync(context);
    logToPPTConsole(`🎉 Successfully inserted content onto current slide!`);
  });

  // Universal Fallback: If shape picture insertion failed, inject via Office Common API
  if (hasImages && !imageInserted && imagesToInsert.length > 0) {
    const imgWidth = (hasMeaningfulBody || hasTable) ? 420 : 620;
    const imgLeft = (hasMeaningfulBody || hasTable) ? 490 : 170;
    for (const rawImg of imagesToInsert) {
      await insertPictureViaCommonApi(rawImg, {
        left: imgLeft,
        top: 80,
        width: imgWidth,
        height: 370
      });
    }
  }
}

/**
 * Builds all parsed slides sequentially, appending them to the end of the presentation.
 * @param {Array} slideStructures - Array of { title, body, base64Images, slideNumber }
 * @param {Object} options - { mode: 'insert' | 'replace' }
 * @param {Function} onProgress - Callback with { current, total, title }
 */
/** Diagnostic helper: true when a hex colour reads as a dark canvas. */
function luminanceIsDark(hex) {
  const h = String(hex || "").replace("#", "");
  if (h.length !== 6) return false;
  const r = parseInt(h.substring(0, 2), 16);
  const g = parseInt(h.substring(2, 4), 16);
  const b = parseInt(h.substring(4, 6), 16);
  return (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255 < 0.5;
}

/**
 * DIAGNOSTIC ONLY. Logs the real Office theme colour scheme alongside the values
 * the text sampler chose, so the two can be compared in the Troubleshooting &
 * Diagnostics console. Nothing here feeds back into rendering.
 */
async function logThemeProbe(theme) {
  try {
    // getDeckTheme() already probed; reuse it rather than opening a second
    // PowerPoint.run just to log.
    const probe = theme.probe || await probeThemeColorScheme();

    if (probe.available) {
      const c = probe.colors || {};
      const accents = ["Accent1", "Accent2", "Accent3", "Accent4", "Accent5", "Accent6"]
        .filter((k) => c[k])
        .map((k) => `${k}=${c[k]}`)
        .join(" ");
      const neutrals = ["Dark1", "Dark2", "Light1", "Light2"]
        .filter((k) => c[k])
        .map((k) => `${k}=${c[k]}`)
        .join(" ");
      logToPPTConsole(`Theme probe: real Office theme | ${accents}`);
      if (neutrals) logToPPTConsole(`Theme probe: real Office theme | ${neutrals}`);

      const bg = probe.background;
      if (bg) {
        // A solid background tells us definitively whether the canvas is dark.
        // Anything else (gradient, picture) means we cannot know from the API
        // alone, which is itself worth recording.
        const verdict = bg.color
          ? `${luminanceIsDark(bg.color) ? "DARK canvas" : "light canvas"}`
          : "canvas brightness unknown from this fill type";
        logToPPTConsole(
          `Theme probe: master background | type=${bg.type}` +
          `${bg.color ? ` color=${bg.color}` : " (no solid colour)"} -> ${verdict}`
        );
      } else {
        logToPPTConsole(`Theme probe: master background | could not be read.`);
      }

      // The master is not authoritative - a layout can override it. One real
      // deck reported Solid #FFFFFF here while rendering black slides, which
      // is how dark text ended up on a dark canvas.
      const evidence = theme && theme.backgroundEvidence;
      if (evidence && evidence.slides && evidence.slides.length > 0) {
        logToPPTConsole(
          `Theme probe: slide backgrounds | ${evidence.slides.join(", ")} -> ` +
          (evidence.agreed
            ? `all agree on ${evidence.agreed}`
            : "they disagree, so the master decides the canvas")
        );
      }

      const slideBg = probe.slideBackground;
      if (slideBg) {
        const slideVerdict = slideBg.color
          ? `${luminanceIsDark(slideBg.color) ? "DARK canvas" : "light canvas"}`
          : "no flat colour - the master value may not be what is rendered";
        logToPPTConsole(
          slideBg.error
            ? `Theme probe: slide background | could not be read (${slideBg.error})`
            : `Theme probe: slide ${slideBg.index + 1} background | type=${slideBg.type}` +
              `${slideBg.color ? ` color=${slideBg.color}` : " (no solid colour)"} -> ${slideVerdict}`
        );
      }
    } else {
      logToPPTConsole(
        `Theme probe: the real theme colour scheme is unavailable (${probe.reason || "unknown"}).`,
        "warn"
      );
    }

    const p = theme.provenance;
    if (p) {
      const where = (idx) =>
        typeof idx === "number" && idx >= 0 ? `deck slide ${idx + 1}` : "not found";
      const skipped = p.skippedGenerated
        ? `, skipped ${p.skippedGenerated} slide(s) this add-in made`
        : "";
      const neutral = typeof p.neutralRatio === "number"
        ? ` | ${Math.round(p.neutralRatio * 100)}% of that text is neutral`
        : "";
      logToPPTConsole(
        `Theme probe: sampler scanned ${p.scannedSlides} slide(s)${skipped}, ${p.textShapes} text shape(s)${neutral} | ` +
        `accent ${theme.accentColor} from ${where(p.accentFromSlide)} | ` +
        `heading font ${theme.headingFont || "(theme default)"} from ${where(p.headingFromSlide)}.`
      );
    } else {
      logToPPTConsole(
        `Theme probe: the sampler found nothing to read, so neutral defaults are in use.`
      );
    }

    logToPPTConsole(
      `Theme palette in use: accent=${theme.accentColor} title=${theme.titleColor} ` +
      `body=${theme.bodyColor} muted=${theme.mutedColor} card=${theme.cardFill} | ` +
      `ramp ${(theme.accentRamp || []).join(" ")} | ${theme.isDark ? "dark" : "light"} canvas.`
    );

    if (theme.paletteNote) {
      logToPPTConsole(`Theme palette: ${theme.paletteNote}.`);
    }
  } catch (probeErr) {
    console.warn("[PPTBuilder] Theme probe notice:", probeErr);
  }
}

export async function buildPresentation(slideStructures, options = {}, onProgress = null) {
  if (typeof PowerPoint === "undefined" || !PowerPoint.run) {
    throw new Error("PowerPoint Office.js runtime is not available.");
  }

  if (!slideStructures || slideStructures.length === 0) {
    throw new Error("No slide structures found to build.");
  }

  const totalSlides = slideStructures.length;
  resetSyncDiagnostics();
  resetBuildAbort();
  // Per build, not per module: the user can run against a different deck, and
  // so a different layout, without reloading the task pane.
  resetLayoutPromptState();
  resetSlideIdCache();
  resetDeferredTables();
  logToPPTConsole(`=== Starting Generation of ${totalSlides} Slide(s) === [builder ${BUILDER_REVISION}]`);
  logCardLayoutPlan(slideStructures);
  STYLED_BULLET_SLIDES = 0;

  // 1. Resolve the layout FIRST. Its background is the canvas the slides will
  //    actually sit on, and the theme's contrast maths is only as good as the
  //    canvas it is given - one deck's master claimed Solid #FFFFFF for a
  //    layout that renders dark maroon.
  let layoutOptions = null;
  let layoutInfo = null;
  const userLayoutChoice = options.layoutChoice || "auto";
  const endLayoutPhase = beginPhase("layout discovery");

  if (userLayoutChoice.includes("::")) {
    const parts = userLayoutChoice.split("::");
    layoutOptions = { slideMasterId: parts[0], layoutId: parts[1] };
    logToPPTConsole(`ℹ️ Using custom slide layout: Master ${parts[0]}, Layout ${parts[1]}`);
  } else {
    try {
      // getThemeBlankLayoutOptions() logs its own outcome, including the list
      // of layouts it had to choose between when nothing matched.
      layoutInfo = await getThemeBlankLayoutOptions();
      layoutOptions = layoutInfo ? layoutInfo.options : null;
    } catch (lErr) {
      console.warn("[PPTBuilder] Notice discovering theme layout:", lErr);
    }
  }

  endLayoutPhase(layoutOptions ? "matched a layout" : "no layout match");

  const canvasHint = layoutInfo && layoutInfo.background ? layoutInfo.background.color : null;

  // 2. Sample the customer's theme BEFORE we add any slides of our own, so the
  //    tokens reflect their template rather than our generated content.
  // Re-derive the tokens with the canvas hint from the layout we just chose.
  //
  // This used to call resetDeckTheme() and sample the whole deck a second
  // time - 10.6 seconds on a 17-slide deck - even though the hint is applied
  // to data the first sample already collected. The adapter now owns the
  // reset, and the scan it performs is the only one in the build.
  //
  // reuseSample is a request, not an assumption: if no sample was primed (a
  // direct caller, or a first scan the host refused), getDeckTheme falls back
  // to reading the deck exactly as before.
  const endDerivePhase = beginPhase("theme derive (canvas hint)");
  const theme = await getDeckTheme(true, { canvasHint, reuseSample: true });
  endDerivePhase(theme && theme.themeSource ? theme.themeSource : "");
  logToPPTConsole(
    `Theme: ${theme.themeSource || (theme.sampled ? "sampled from deck" : "neutral defaults")} | font=${theme.headingFont || "(theme default)"} | accent=${theme.accentColor} | canvas=${theme.canvasColor}${canvasHint ? " (from the layout)" : ""} ${theme.slideWidth}x${theme.slideHeight}`
  );
  const endProbePhase = beginPhase("theme probe");
  await logThemeProbe(theme);
  endProbePhase();

  // 3. Pre-process images
  const endCompressPhase = beginPhase("image compression");
  for (let idx = 0; idx < slideStructures.length; idx++) {
    const slide = slideStructures[idx];
    slide.compressedImages = [];
    slide.compressedImageMeta = [];
    const rawImages = slide.base64Images || [];
    for (const rawImg of rawImages) {
      try {
        const meta = {};
        const comp = await compressImageForPowerPoint(rawImg, 1600, 1120, meta);
        if (comp && comp.length > 50) {
          slide.compressedImages.push(comp);
          slide.compressedImageMeta.push(meta);
          logToPPTConsole(
            meta.compressed
              ? `Slide ${idx + 1}: image re-encoded ${meta.width || "?"}x${meta.height || "?"} -> ${Math.round((comp.length * 0.75) / 1024)} KB.`
              : `Slide ${idx + 1}: image could NOT be re-encoded (decode timed out) - sending ${Math.round((comp.length * 0.75) / 1024)} KB as-is.`
          );
        }
      } catch (cErr) {
        logToPPTConsole(`Image compression notice: ${cErr.message}`);
      }
    }
  }

  endCompressPhase(
    `${slideStructures.reduce((n, s) => n + ((s.compressedImages || []).length), 0)} image(s)`
  );

  // 4. Build each slide sequentially, appending to end of presentation
  let interSlideYieldMs = 0;
  let createdCount = 0;
  let abortedAtSlide = 0;      // the slide the host went quiet on, if any
  const createdIndexes = [];   // where the host actually put each slide
  const failures = [];

  for (let i = 0; i < totalSlides; i++) {
    const slideData = slideStructures[i];
    const slideNum = i + 1;

    if (typeof onProgress === "function") {
      onProgress({
        current: slideNum,
        total: totalSlides,
        title: slideData.title
      });
    }

    const endSlidePhase = beginPhase(`slide ${slideNum}`);
    try {
      slideData.deckSize = totalSlides;
      // Decided per slide from the speed measured so far, so a host that
      // recovers gets its cards back on the next card slide.
      if (MS_PER_SHAPE_EST > slowHostMsPerShape() && isCardLayoutSlide(slideData)) {
        const layoutName = slideData.visualType || "structuredCards";
        if (demoteCardSlideForSlowHost(slideData)) {
          logToPPTConsole(
            `Slide ${slideNum}: PowerPoint is slow right now (~${Math.round(MS_PER_SHAPE_EST)}ms per shape) - ` +
            `drawing this "${layoutName}" slide as styled bullets so the deck finishes.`,
            "warn"
          );
        }
      }
      await createSingleSlide(slideData, slideNum, layoutOptions, options, theme);
      createdCount++;
      const placedAt = takeAppendedSlideIndex();
      if (placedAt >= 0 && !createdIndexes.includes(placedAt)) createdIndexes.push(placedAt);
      // The deferred finish pass runs in a fresh context where this slide's
      // proxy is dead, so its deck position is the only handle back to it.
      noteDeferredTableSlideIndex(slideNum, placedAt);
    } catch (slideErr) {
      const detail = slideErr && slideErr.message ? slideErr.message : String(slideErr);
      failures.push(`Slide ${slideNum}: ${detail}`);
      logToPPTConsole(`Slide ${slideNum} failed: ${detail}`, true);
      console.error(`[PPTBuilder] Slide ${slideNum} failed:`, slideErr);
    }

    endSlidePhase(slideData && slideData.layout ? String(slideData.layout) : "");

    if (BUILD_ABORTED) {
      abortedAtSlide = slideNum;
      logToPPTConsole(
        `Build stopped at slide ${slideNum}: PowerPoint did not answer within ` +
        `${Math.round(BUILD_ABORTED.ms / 1000)}s${abortedWhere()}. ` +
        (createdCount > 0
          ? `The ${createdCount} slide(s) finished before it have been kept; slide ${slideNum} ` +
            `is partly drawn. `
          : `No slide was finished. `) +
        `Reload the task pane before trying again - re-running without a reload races the ` +
        `stalled build.`,
        true
      );
      break;
    }

    // Yield event loop briefly so the PowerPoint host can finalize layout.
    // Measured rather than hidden: five of these are 1.25s, which is exactly
    // the "(unattributed)" remainder the first timing runs reported.
    const yieldStartedAt = nowMs();
    await delay(250);
    interSlideYieldMs += nowMs() - yieldStartedAt;
  }

  if (interSlideYieldMs > 0) {
    addPhase("inter-slide yields", interSlideYieldMs, `${totalSlides} x 250ms`);
  }

  // 4b. Verify and style every table now that the host has had the whole build
  //     to commit them. Before the placeholder sweep, so the sweep sees the
  //     deck in its final shape.
  if (deferredTables.length > 0 && !BUILD_ABORTED) {
    const endTableFinish = beginPhase("table finish");
    let summary = { total: 0, styled: 0 };
    try {
      summary = await finishDeferredTables();
    } catch (finishErr) {
      console.warn("[PPTBuilder] Deferred table pass failed:", finishErr);
    }
    endTableFinish(`${summary.styled}/${summary.total} styled`);
  }

  // 5. Safety net for the layout's prompt placeholders. Each slide already
  //    cleared its own before drawing - see stripQueuedPrompts() - which is
  //    what makes the web client repaint them away. This pass only catches a
  //    slide whose early cleanup the host refused, and on a healthy build it
  //    finds nothing. It stays LAST because a prompt removed here still will
  //    not repaint until the user reloads, so it is a correctness backstop
  //    rather than the mechanism. Each slide is swept and then re-read in its
  //    own context, so a ghost the host failed to delete is reported by deck
  //    position instead of being counted as done.
  if (createdCount > 0 && !BUILD_ABORTED) {
    try {
      let indexes = createdIndexes.slice().sort((a, b) => a - b);
      if (indexes.length === 0) {
        // Only when the host never told us where a slide landed. getCount() is
        // the number this builder trusts least, so a window derived from it is
        // a last resort, not the default.
        const finalCount = await readSlideCount();
        if (typeof finalCount === "number" && finalCount >= createdCount) {
          for (let k = finalCount - createdCount; k < finalCount; k++) indexes.push(k);
        }
      }
      if (indexes.length > 0) {
        const endSweepPhase = beginPhase("placeholder sweep");
        await removeEmptyPlaceholders(indexes);
        endSweepPhase(`${indexes.length} slide(s)`);
      }
    } catch (cleanupErr) {
      console.warn("[PPTBuilder] Placeholder cleanup skipped:", cleanupErr);
    }
  }

  // 6. Never silently report success when nothing was actually inserted.
  if (createdCount === 0) {
    throw new Error(
      `No slides could be created in PowerPoint. First error: ${failures[0] || "unknown error"}`
    );
  }

  if (BUILD_ABORTED) {
    logToPPTConsole(
      `⚠️ Stopped at slide ${abortedAtSlide} of ${totalSlides} - PowerPoint stopped responding. ` +
      `${createdCount} slide(s) completed.`,
      true
    );
  } else if (failures.length > 0) {
    logToPPTConsole(`⚠️ Created ${createdCount}/${totalSlides} slides. Skipped: ${failures.join(" | ")}`, true);
  } else {
    logToPPTConsole(`🎉 All ${totalSlides} slides created successfully!`);
  }

  return {
    created: createdCount,
    total: totalSlides,
    failures,
    // 0 when the build ran to the end. The adapter needs this to tell "stopped
    // at slide 7" apart from "one slide skipped": both leave one failure.
    abortedAtSlide: BUILD_ABORTED ? abortedAtSlide : 0
  };
}

/**
 * Offline-harness hooks.
 *
 * The card committer's failure paths - a refused batch, a refusal the host did
 * not honour, a discarded row prologue - cannot be reached through
 * buildPresentation with a realistic mock, and they are exactly the paths that
 * decide whether a card goes missing or gets drawn twice. Nothing in the
 * add-in imports this.
 */
export const __test = {
  selectSlideForCommonApi,
  retryPictureOnSlideIndex,
  enhanceBulletsCosmetic,
  syncDeadlineFor,
  noteCommitRate,
  withoutTableEcho,
  isCoverSlide,
  coverLedeText,
  renderProcessTimeline,
  renderMultiColumnCards,
  renderKpiHeroCards,
  createCardCommitter,
  countSlideShapes,
  knownShapeCount,
  meteredSync,
  resetBuildAbort,
  setSyncTimeoutForTests,
  setSlimCardsForTests
};
