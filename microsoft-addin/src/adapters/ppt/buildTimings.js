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
 * Slide-build timing instrumentation.
 *
 * Measurement only. Every entry point swallows its own errors so that a bug in
 * the instrumentation can never take down a build. Output goes to the
 * Troubleshooting & Diagnostics console, never to the chat.
 *
 * Phases are reported in execution order rather than sorted by cost: knowing
 * that the deck was scanned twice is the whole point, and sorting would hide it.
 */

import { logToPPTConsole } from "./pptDiagnostics.js";

export function nowMs() {
  try {
    if (typeof performance !== "undefined" && typeof performance.now === "function") {
      return performance.now();
    }
  } catch (_) {}
  return Date.now();
}

function now() {
  return nowMs();
}

let runStart = 0;
let phases = [];
let syncCount = 0;
let syncMs = 0;
let syncFailures = 0;
let backoffMs = 0;
let shapeCount = 0;

export function resetBuildTimings() {
  try {
    runStart = now();
    phases = [];
    syncCount = 0;
    syncMs = 0;
    syncFailures = 0;
    backoffMs = 0;
    shapeCount = 0;
  } catch (_) {}
}

/**
 * Start a phase. Returns a function that ends it.
 * The returned ender is safe to call more than once; only the first counts.
 *
 *   const end = beginPhase("deck scan");
 *   ...
 *   end("7 slides sampled");
 *
 * Each phase also records the host round trips that happened inside it. Time
 * alone says which phase is slow; the sync count says whether it is slow
 * because of chatter or because of one expensive call, and those have
 * completely different fixes.
 */
export function beginPhase(name) {
  let started = 0;
  let syncsAtStart = 0;
  let shapesAtStart = 0;
  let failuresAtStart = 0;
  let done = false;
  try {
    started = now();
    syncsAtStart = syncCount;
    shapesAtStart = shapeCount;
    failuresAtStart = syncFailures;
  } catch (_) {}
  return function endPhase(detail) {
    try {
      if (done) return;
      done = true;
      phases.push({
        name: String(name),
        ms: now() - started,
        syncs: syncCount - syncsAtStart,
        shapes: shapeCount - shapesAtStart,
        refused: syncFailures - failuresAtStart,
        detail: detail ? String(detail) : "",
      });
    } catch (_) {}
  };
}

/** Records a phase measured by the caller, e.g. an accumulated total. */
export function addPhase(name, ms, detail) {
  try {
    phases.push({
      name: String(name),
      ms: Number(ms) || 0,
      syncs: 0,
      shapes: 0,
      detail: detail ? String(detail) : "",
    });
  } catch (_) {}
}

/**
 * Counts a shape this builder asked the host to create.
 *
 * Deliberately counts our own `add*` calls rather than reading the slide's
 * shape count back: this host miscounts shapes badly enough that the bare-slide
 * check has to label the number "advisory", but it cannot miscount instructions
 * we issued ourselves. A table counts as one shape, not one per cell.
 *
 * The point of the measurement: across three runs the host spent ~70s doing the
 * work regardless of whether it was packeted into 136 syncs or 212, which says
 * the cost follows the shapes drawn rather than the round trips. That is either
 * true or it is not, and ms/shape is the number that settles it.
 */
export function recordShapesDrawn(n) {
  try {
    shapeCount += Number(n) || 0;
  } catch (_) {}
}

export function recordSync(ms, ok) {
  try {
    syncCount += 1;
    syncMs += Number(ms) || 0;
    if (!ok) syncFailures += 1;
  } catch (_) {}
}

/**
 * The running sync tally.
 *
 * The per-phase counts say which phase is expensive; this lets one code path
 * inside a phase report what it cost on its own, which is how the table
 * read-back loop gets measured without guessing.
 */
export function syncsSoFar() {
  return syncCount;
}

/** Time spent deliberately doing nothing: retry backoff, inter-slide yields. */
export function recordBackoff(ms) {
  try {
    backoffMs += Number(ms) || 0;
  } catch (_) {}
}

function secs(ms) {
  return (Math.max(0, ms) / 1000).toFixed(2) + "s";
}

function pct(part, whole) {
  if (!whole || whole <= 0) return "--";
  return Math.round((part / whole) * 100) + "%";
}

function pad(str, width) {
  let s = String(str);
  while (s.length < width) s += " ";
  return s;
}

/**
 * @param {number} totalMs authoritative end-to-end time measured by the caller.
 *   Phases will not sum to this, and the gap is the interesting part, so it is
 *   reported explicitly as "(unattributed)" rather than quietly absorbed.
 */
export function logBuildTimings(totalMs) {
  try {
    const total = Number(totalMs) > 0 ? Number(totalMs) : now() - runStart;

    let widest = 14;
    for (let i = 0; i < phases.length; i++) {
      if (phases[i].name.length > widest) widest = phases[i].name.length;
    }

    const lines = [];
    lines.push(`--- build timings -- total ${secs(total)} ---`);

    let attributed = 0;
    for (let i = 0; i < phases.length; i++) {
      const p = phases[i];
      attributed += p.ms;
      const detail = p.detail ? `   ${p.detail}` : "";
      // A refusal is attributed to the phase it happened in. The deck-wide
      // total at the foot of the table has reported as many as 11 refusals
      // while only two reached the log, and there was no way to tell whether
      // the rest came from a slide, from layout discovery or from the table
      // read-back - three problems with nothing in common.
      const syncs = p.syncs > 0
        ? `${p.syncs} sync` + (p.syncs === 1 ? "" : "s") +
          (p.refused > 0 ? ` (${p.refused} refused)` : "")
        : "";
      // ms/shape is the whole point of the column: sync counts have already
      // been driven down twice with no effect on the clock, so the question is
      // whether the remaining time tracks the shapes instead.
      const shapes = p.shapes > 0
        ? `${p.shapes} shape` + (p.shapes === 1 ? "" : "s") +
          ` @ ${Math.round(p.ms / p.shapes)}ms`
        : "";
      lines.push(
        `  ${pad(p.name, widest)}  ${pad(secs(p.ms), 8)} ${pad(pct(p.ms, total), 4)} ` +
        `${pad(syncs, 20)}${pad(shapes, 20)}${detail}`
      );
    }

    const gap = total - attributed;
    if (gap > 0) {
      lines.push(
        `  ${pad("(unattributed)", widest)}  ${pad(secs(gap), 8)} ${pad(pct(gap, total), 4)}`
      );
    }

    if (syncCount > 0) {
      const avg = Math.round(syncMs / syncCount);
      lines.push(
        `  -- ${syncCount} host sync(s), ${secs(syncMs)} total ${pct(syncMs, total)}, ` +
          `avg ${avg}ms, ${syncFailures} refused`
      );
    }
    if (shapeCount > 0) {
      lines.push(
        `  -- ${shapeCount} shape(s) drawn, avg ${Math.round(total / shapeCount)}ms/shape`
      );
    }
    if (backoffMs > 0) {
      lines.push(
        `  -- deliberate waiting (backoff + yields): ${secs(backoffMs)} ${pct(backoffMs, total)}`
      );
    }

    logToPPTConsole(lines.join("\n"));
  } catch (_) {}
}
