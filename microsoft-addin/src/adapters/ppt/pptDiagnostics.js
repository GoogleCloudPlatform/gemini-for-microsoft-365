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
 * PowerPoint Diagnostics & Live Logger
 * 
 * Routes PowerPoint slide generation logs and test calls directly
 * into the unified Troubleshooting & Diagnostics panel.
 */

export function logToPPTConsole(msg, isError = false) {
  // Every line of build output doubles as a liveness tick for the slide-build
  // lease in PPTAdapter. It has to be driven by real progress like this: a
  // plain setInterval would keep ticking straight through a hung await, since
  // a pending promise does not block the event loop, and the lease would never
  // be recognised as abandoned.
  if (typeof window !== "undefined" && typeof window.__touchSlideBuildLease === "function") {
    try {
      window.__touchSlideBuildLease();
    } catch (_) {}
  }

  const logBox = document.getElementById("diagLogBox");
  if (!logBox) {
    console.log(`[PPT] ${msg}`);
    return;
  }
  const ts = new Date().toLocaleTimeString();
  const prefix = isError ? "❌ [PPT]" : "ℹ️ [PPT]";
  logBox.textContent += `\n[${ts}] ${prefix} ${msg}`;
  logBox.scrollTop = logBox.scrollHeight;
}

export function initPowerPointDiagnostics() {
  // Integrated into unified taskpane troubleshooting panel
}
