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
 * Deck Theme Sampler for the PowerPoint Adapter
 *
 * Inspects the user's OPEN presentation (fonts, text colors, shape fills and slide
 * geometry) and derives a set of design tokens so that AI-generated slides visually
 * blend into the customer's own template instead of imposing a hard-coded look.
 *
 * Everything in here is best-effort and defensive: PowerPoint on Mac / Web / Windows
 * expose different requirement sets, so every load is guarded and every failure
 * degrades silently back to neutral defaults that respect the deck theme
 * (i.e. we simply do not set a property and let PowerPoint apply the theme value).
 *
 * @author Sathya AG, Principal Architect, Google
 */

/* ------------------------------------------------------------------ */
/* Color utilities                                                     */
/* ------------------------------------------------------------------ */

/**
 * Tag key stamped on every slide this add-in creates.
 *
 * The sampler below learns the deck's colours from its first few slides. After
 * one generation run those are OUR slides, so without a way to recognise them
 * the palette samples its own output: run 1 read a monochrome deck and painted
 * it black, run 2 read run 1's slides and painted the next deck pink.
 *
 * PowerPoint upper-cases tag keys, so every comparison must be case-folded.
 */
export const ADDIN_SLIDE_TAG = "GEMINIADDINSLIDE";

/**
 * Tag key holding the palette a previous run actually used.
 *
 * The sampler cannot see our own slides, and that is correct - deriving a
 * palette from colours we chose last run is how one deck went monochrome at
 * 09:15 and hot pink at 09:17. But it leaves a deck whose only content IS our
 * slides with nothing to read at all. One real deck reported "scanned 3
 * slide(s), skipped 10 slide(s) this add-in made, 0 text shape(s)", fell
 * through to the theme's six accents, and put pink, green and olive cards
 * beside ten grey ones.
 *
 * Re-inferring was the mistake. We knew the palette when we chose it, so it is
 * recorded here and read back. Sampling stays the way a palette is FOUND; this
 * is how it is KEPT.
 */
export const ADDIN_PALETTE_TAG = "GEMINIADDINPALETTE";

const PALETTE_TAG_VERSION = "1";

/**
 * The theme values a palette is derived from.
 *
 * If any of them move, the user restyled the deck on purpose and the recorded
 * palette is stale - which is the whole escape hatch. Without it, one recorded
 * palette would outlive every future design change.
 */
const FINGERPRINT_KEYS = [
  "Accent1", "Accent2", "Accent3", "Accent4", "Accent5", "Accent6",
  "Dark1", "Dark2", "Light1", "Light2"
];

// FNV-1a. Not security, just a short stable string that fits in a tag value.
function hash32(str) {
  let h = 0x811c9dc5;
  for (let i = 0; i < str.length; i++) {
    h ^= str.charCodeAt(i);
    h = (h + ((h << 1) + (h << 4) + (h << 7) + (h << 8) + (h << 24))) >>> 0;
  }
  return h.toString(36);
}

/**
 * Fingerprints the deck design a palette was derived from.
 *
 * Deliberately built from the SAMPLE, not from the finished tokens: the tokens
 * carry resolved values (a canvas that fell back to Light1, an accent that was
 * clamped for contrast), so fingerprinting those would hash a different thing
 * on write than on read and never match.
 */
export function themeFingerprint(sample = {}) {
  const real = sample.themeColors || null;
  const scheme = real
    ? FINGERPRINT_KEYS.map((k) => normalizeHex(real[k]) || "-").join("")
    : "none";
  return hash32(`${scheme}|${normalizeHex(sample.canvasColor) || "-"}`);
}

function paletteHex(color) {
  return (normalizeHex(color) || "#000000").slice(1);
}

function paletteFont(name) {
  // Pipes would break the field split; the cap keeps the whole value well
  // inside what a tag value will hold.
  return String(name || "").replace(/\|/g, " ").trim().slice(0, 24);
}

export function encodePalette(theme = {}, fingerprint = "") {
  const ramp = (Array.isArray(theme.accentRamp) ? theme.accentRamp : []).slice(0, 4);
  while (ramp.length < 4) ramp.push(theme.accentColor || "#000000");
  return [
    PALETTE_TAG_VERSION,
    fingerprint || "",
    paletteHex(theme.canvasColor),
    paletteHex(theme.accentColor),
    paletteHex(theme.titleColor),
    paletteHex(theme.bodyColor),
    paletteHex(theme.mutedColor),
    paletteHex(theme.cardFill),
    ramp.map(paletteHex).join(""),
    paletteFont(theme.headingFont),
    paletteFont(theme.bodyFont)
  ].join("|");
}

export function decodePalette(value) {
  if (!value || typeof value !== "string") return null;
  const parts = value.split("|");
  if (parts.length !== 11 || parts[0] !== PALETTE_TAG_VERSION) return null;

  const ramp = [];
  for (let i = 0; i < 4; i++) {
    const c = normalizeHex(`#${String(parts[8]).slice(i * 6, i * 6 + 6)}`);
    if (!c) return null;
    ramp.push(c);
  }

  const out = {
    fingerprint: parts[1],
    canvasColor: normalizeHex(`#${parts[2]}`),
    accentColor: normalizeHex(`#${parts[3]}`),
    titleColor: normalizeHex(`#${parts[4]}`),
    bodyColor: normalizeHex(`#${parts[5]}`),
    mutedColor: normalizeHex(`#${parts[6]}`),
    cardFill: normalizeHex(`#${parts[7]}`),
    accentRamp: ramp,
    headingFont: parts[9] || null,
    bodyFont: parts[10] || null
  };

  // Half a palette is worse than none: it would paint some elements from the
  // record and derive the rest, matching neither the old slides nor the theme.
  const required = ["canvasColor", "accentColor", "titleColor", "bodyColor", "mutedColor", "cardFill"];
  for (const key of required) {
    if (!out[key]) return null;
  }
  return out;
}

export function normalizeHex(value) {
  if (!value || typeof value !== "string") return null;
  let v = value.trim();
  if (!v) return null;
  if (v[0] !== "#") v = `#${v}`;
  if (/^#[0-9a-fA-F]{3}$/.test(v)) {
    v = `#${v[1]}${v[1]}${v[2]}${v[2]}${v[3]}${v[3]}`;
  }
  return /^#[0-9a-fA-F]{6}$/.test(v) ? v.toUpperCase() : null;
}

function toRgb(hex) {
  const h = normalizeHex(hex);
  if (!h) return null;
  return {
    r: parseInt(h.substring(1, 3), 16),
    g: parseInt(h.substring(3, 5), 16),
    b: parseInt(h.substring(5, 7), 16)
  };
}

function toHex({ r, g, b }) {
  const clamp = (n) => Math.max(0, Math.min(255, Math.round(n)));
  const part = (n) => clamp(n).toString(16).padStart(2, "0");
  return `#${part(r)}${part(g)}${part(b)}`.toUpperCase();
}

/** Relative luminance (0 = black, 1 = white). */
export function luminance(hex) {
  const rgb = toRgb(hex);
  if (!rgb) return 0.5;
  return (0.2126 * rgb.r + 0.7152 * rgb.g + 0.0722 * rgb.b) / 255;
}

/** HSL-style saturation (0 = grey, 1 = fully saturated). */
export function saturation(hex) {
  const rgb = toRgb(hex);
  if (!rgb) return 0;
  const r = rgb.r / 255;
  const g = rgb.g / 255;
  const b = rgb.b / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return 0;
  const d = max - min;
  return l > 0.5 ? d / (2 - max - min) : d / (max + min);
}

/** Linearly blends two colors. ratio = 0 returns `from`, ratio = 1 returns `to`. */
export function mix(from, to, ratio) {
  const a = toRgb(from);
  const b = toRgb(to);
  if (!a || !b) return normalizeHex(from) || normalizeHex(to) || "#000000";
  const t = Math.max(0, Math.min(1, ratio));
  return toHex({
    r: a.r + (b.r - a.r) * t,
    g: a.g + (b.g - a.g) * t,
    b: a.b + (b.b - a.b) * t
  });
}

export function tint(hex, ratio) {
  return mix(hex, "#FFFFFF", ratio);
}

export function shade(hex, ratio) {
  return mix(hex, "#000000", ratio);
}

/**
 * Picks black or white text for maximum contrast against `bg`.
 *
 * Decided by measured WCAG contrast rather than a luminance threshold: a
 * mid-tone accent such as #E97132 sits just under a 0.6 cutoff and would be
 * given white text at 3.06:1, where black reads at 6.86:1.
 */
export function readableTextOn(bg) {
  var dark = "#111827";
  var light = "#FFFFFF";
  try {
    return contrastRatio(dark, bg) >= contrastRatio(light, bg) ? dark : light;
  } catch (err) {
    return luminance(bg) > 0.6 ? dark : light;
  }
}

function srgbChannel(c) {
  const v = c / 255;
  return v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4);
}

/**
 * True WCAG relative luminance. `luminance()` above is a fast approximation
 * that is fine for ranking colours but wrong for contrast decisions.
 */
export function wcagLuminance(hex) {
  const rgb = toRgb(hex);
  if (!rgb) return 0.5;
  return 0.2126 * srgbChannel(rgb.r) + 0.7152 * srgbChannel(rgb.g) + 0.0722 * srgbChannel(rgb.b);
}

/** WCAG contrast ratio: 1 = identical, 21 = black on white. */
export function contrastRatio(a, b) {
  const la = wcagLuminance(a);
  const lb = wcagLuminance(b);
  return (Math.max(la, lb) + 0.05) / (Math.min(la, lb) + 0.05);
}

/**
 * Walks `fg` toward black or white until it reads cleanly against `bg`.
 *
 * Text nobody can read is worse than text that is slightly off-palette. This is
 * the last line of defence: whatever the theme says, body copy must be legible.
 */
export function ensureContrast(fg, bg, minRatio = 4.5) {
  const start = normalizeHex(fg);
  const canvas = normalizeHex(bg);
  if (!start) return null;
  if (!canvas || contrastRatio(start, canvas) >= minRatio) return start;

  const target = wcagLuminance(canvas) > 0.5 ? "#000000" : "#FFFFFF";
  for (let step = 1; step <= 20; step++) {
    const candidate = mix(start, target, step / 20);
    if (contrastRatio(candidate, canvas) >= minRatio) return candidate;
  }
  return target;
}

/**
 * Picks the accent a designer would: the earliest theme accent that still reads
 * clearly against the canvas.
 *
 * Templates routinely set Accent1-3 to pale background washes. One real deck
 * shipped Accent1=#E6E3DD, Accent2=#DCE4E7, Accent3=#F2F0ED - all near-white -
 * with the usable colours at Accent4-6. Taking Accent1 on faith there would
 * paint metric numbers and card rules in something nobody can see, so order of
 * preference yields to legibility.
 */
function pickAccentForCanvas(accents, canvas, minRatio = 3.5) {
  if (!accents || accents.length === 0) return null;
  for (const candidate of accents) {
    if (contrastRatio(candidate, canvas) >= minRatio) return candidate;
  }
  // Nothing clears the bar, so take the best of a bad set rather than none.
  return accents.reduce(
    (best, c) => (contrastRatio(c, canvas) > contrastRatio(best, canvas) ? c : best),
    accents[0]
  );
}

/**
 * A colour is only usable as an accent if it is neither near-black nor
 * near-white. The cream #F6F2E6 that washed out an entire deck scores 0.47
 * saturation but 0.95 luminance, which is exactly what this rejects.
 */
function usableAccent(hex) {
  const h = normalizeHex(hex);
  if (!h) return null;
  const l = luminance(h);
  return l > 0.06 && l < 0.82 ? h : null;
}

/* ------------------------------------------------------------------ */
/* Neutral fallback theme (used when nothing can be sampled)           */
/* ------------------------------------------------------------------ */

const NEUTRAL_ACCENT = "#1A73E8";

/**
 * Rebuilds the full token set from a recorded palette.
 *
 * Only the independent colours are stored; everything else is recomputed with
 * the SAME expressions buildTokens uses, so a reused palette is identical to
 * the one that was recorded rather than merely similar. Keep the two in step.
 */
function tokensFromPalette(rec, sample) {
  const canvas = rec.canvasColor;
  const isDark = luminance(canvas) < 0.5;
  const accent = rec.accentColor;
  const cardFill = rec.cardFill;

  return {
    themeSource: "reusing this deck's recorded palette",
    canvasColor: canvas,

    headingFont: rec.headingFont || null,
    bodyFont: rec.bodyFont || rec.headingFont || null,

    isDark,
    accentColor: accent,
    accentSoft: isDark ? shade(accent, 0.6) : tint(accent, 0.88),
    onAccentColor: readableTextOn(accent),

    titleColor: rec.titleColor,
    bodyColor: rec.bodyColor,
    mutedColor: rec.mutedColor,

    cardFill,
    cardBorder: isDark ? mix(cardFill, "#FFFFFF", 0.14) : mix(cardFill, rec.titleColor, 0.14),
    shadowColor: isDark ? "#000000" : "#D7DDE5",
    onCardColor: readableTextOn(cardFill),

    accentRamp: rec.accentRamp.slice(),

    slideWidth: sample.slideWidth || 960,
    slideHeight: sample.slideHeight || 540,
    sampled: Boolean(sample.sampled),

    themeFingerprint: rec.fingerprint,
    reusedPalette: true,
    provenance: sample.provenance || null,
    backgroundEvidence: sample.backgroundEvidence || null
  };
}

export function buildTokens(sample = {}) {
  const real = sample.themeColors || null;

  // A palette this add-in recorded on a slide of this very deck - see
  // ADDIN_PALETTE_TAG. Two conditions, both necessary:
  //
  //   1. The design it was derived from is unchanged. A different fingerprint
  //      means the deck was restyled on purpose and new slides must follow.
  //   2. The sampler could not form an opinion of its own. Below two sampled
  //      colours `monochromeDeck` cannot be true whatever the deck looks like,
  //      so the palette falls through to the theme's six accents - that is the
  //      gap the skip guard leaves, and the only gap this record fills.
  //
  // Condition 2 matters as much as condition 1. Without it a user who
  // recoloured their own slides could never move the palette again, because
  // the fingerprint only watches the theme, not the content.
  const recorded = decodePalette(sample.recordedPalette);
  const fingerprint = themeFingerprint(sample);
  const sampledColors = Array.isArray(sample.textColors) ? sample.textColors.length : 0;
  const samplerHasAnOpinion = sampledColors >= 2;

  let paletteNote = null;
  if (recorded) {
    if (recorded.fingerprint !== fingerprint) {
      paletteNote = "a recorded palette was ignored - this deck's theme has changed since it was written";
    } else if (samplerHasAnOpinion) {
      paletteNote =
        `a recorded palette was set aside - the deck's own slides carry ${sampledColors} colours to read`;
    } else {
      const reused = tokensFromPalette(recorded, sample);
      reused.paletteNote =
        typeof sample.recordedPaletteIndex === "number" && sample.recordedPaletteIndex >= 0
          ? `reused the palette recorded on deck slide ${sample.recordedPaletteIndex + 1} ` +
            `(nothing else in this deck carries text to sample)`
          : "reused the palette recorded on this deck (nothing else carries text to sample)";
      return reused;
    }
  }

  // ---- Canvas -------------------------------------------------------
  // The CANVAS decides light vs dark. Text colour never does: a deck whose
  // title slide is black is not a dark deck, and treating it as one paints
  // near-white body copy onto a cream background.
  //   1. the slide master's solid background, when the host reports one
  //   2. the theme's Light1, which is the background for every light theme
  //   3. the text heuristic, but only with enough samples to mean anything
  const canvas =
    normalizeHex(sample.canvasColor) ||
    (real ? normalizeHex(real.Light1) : null) ||
    (sample.isDark ? "#1A1A1A" : "#FFFFFF");
  const isDark = luminance(canvas) < 0.5;

  // ---- Accent -------------------------------------------------------
  // A theme's colour scheme is what the deck COULD use, not what it DOES use.
  // When the sampler saw real text on the deck's own slides and every one of
  // those colours was neutral, the deck is a monochrome design and injecting
  // Accent3 - which on one real theme is a hot pink - is simply wrong.
  const deckText = normalizeHex(sample.titleColor) || null;
  const monochrome = Boolean(sample.monochromeDeck);

  const realAccents = real
    ? [real.Accent1, real.Accent2, real.Accent3, real.Accent4, real.Accent5, real.Accent6]
        .map(usableAccent)
        .filter(Boolean)
    : [];

  // The restrained accent for a monochrome deck: its own ink on a light
  // canvas, its own paper on a dark one. usableAccent() is deliberately
  // bypassed here - it rejects near-black, which is exactly what we want.
  const monoAccent = isDark
    ? (real ? normalizeHex(real.Light1) : null) || "#F8FAFC"
    : (deckText && luminance(deckText) < 0.5 ? deckText : null) ||
      (real ? normalizeHex(real.Dark1) : null) ||
      "#1F2937";

  const accentColor =
    (monochrome ? monoAccent : null) ||
    pickAccentForCanvas(realAccents, canvas) ||
    (real ? usableAccent(real.Dark2) : null) ||
    usableAccent(sample.accentColor) ||
    // The sampled title colour is a last resort, and only when it is both
    // saturated AND mid-toned. Omitting the luminance half of this test is
    // what let a near-white cream become the accent.
    (saturation(sample.titleColor || "") > 0.15 ? usableAccent(sample.titleColor) : null) ||
    NEUTRAL_ACCENT;

  // On dark canvases the accent needs lifting so it stays legible.
  const effectiveAccent = isDark && luminance(accentColor) < 0.35 ? tint(accentColor, 0.45) : accentColor;

  // ---- Text ---------------------------------------------------------
  const realTitle = real
    ? (luminance(real.Dark2 || "") < 0.45 ? normalizeHex(real.Dark2) : normalizeHex(real.Dark1))
    : null;

  // On a monochrome deck the deck's own ink beats the theme's Dark2, which on
  // one real black-and-white deck is a navy blue nothing on screen uses.
  const titleSeed = isDark
    ? (real ? normalizeHex(real.Light1) : null) || "#F8FAFC"
    : (monochrome ? deckText : null) || realTitle || deckText || "#0F172A";

  // Every text colour is clamped to a WCAG floor against the canvas. Whatever
  // the theme or the sampler says, body copy must be readable.
  const titleColor = ensureContrast(titleSeed, canvas, 4.5) || titleSeed;
  const bodyColor = ensureContrast(mix(titleColor, canvas, 0.18), canvas, 4.5) || titleColor;
  // 4.5:1, not 3.2:1. Muted text is the slide subtitle and the card eyebrow
  // labels - small text, and often sitting on a themed background picture
  // rather than the flat master colour we measured against.
  const mutedColor = ensureContrast(mix(titleColor, canvas, 0.45), canvas, 4.5) || bodyColor;

  // ---- Cards --------------------------------------------------------
  // The theme's own neutrals, not a shade of the accent. A deck whose theme
  // ships a warm off-white Light2 was getting near-black #2D2D26 bubbles.
  // A neutral that is indistinguishable from the canvas is no use as a card,
  // so it has to clear a visible-difference bar before it is accepted.
  const themeCardSeed = real
    ? (isDark ? normalizeHex(real.Dark2) : normalizeHex(real.Light2))
    : null;
  const cardSeedUsable =
    !monochrome &&
    themeCardSeed &&
    contrastRatio(themeCardSeed, canvas) >= 1.04 &&
    (isDark ? luminance(themeCardSeed) < 0.5 : luminance(themeCardSeed) > 0.5);

  const cardFill = cardSeedUsable
    ? themeCardSeed
    : (isDark ? shade(effectiveAccent, 0.72) : tint(effectiveAccent, 0.93));
  const cardBorder = isDark
    ? mix(cardFill, "#FFFFFF", 0.14)
    : mix(cardFill, titleColor, 0.14);
  const shadowColor = isDark ? "#000000" : "#D7DDE5";

  // ---- Card ramp ----------------------------------------------------
  // Always derived from the ONE chosen accent. r23 used the deck's raw
  // Accent1-4 here, which on a muted theme looked native and on a vivid one
  // put a pink, a green and a gold bar on the same slide.
  //
  // A monochrome deck gets tints of its own ink - four shades of one colour.
  // Everything else gets the muted synthetic mix, which is the ramp that was
  // in place for the run the user singled out as looking right.
  //   - a monochrome deck gets tints of its own ink
  //   - a deck with a real theme gets that theme's accents, which is the whole
  //     point of matching the deck. The synthetic mix below was putting teal
  //     and violet bars on a theme made entirely of warm greys.
  //   - only a deck with no theme at all falls back to the mix
  let accentRamp;
  if (monochrome) {
    accentRamp = [
      effectiveAccent,
      mix(effectiveAccent, canvas, 0.22),
      mix(effectiveAccent, canvas, 0.42),
      mix(effectiveAccent, canvas, 0.6)
    ];
  } else if (realAccents.length >= 3) {
    accentRamp = [effectiveAccent];
    for (const c of realAccents) {
      if (accentRamp.length >= 4) break;
      if (accentRamp.indexOf(c) === -1) accentRamp.push(c);
    }
    while (accentRamp.length < 4) {
      accentRamp.push(mix(effectiveAccent, canvas, 0.35));
    }
  } else {
    accentRamp = [
      effectiveAccent,
      mix(effectiveAccent, "#0D9488", 0.45),
      mix(effectiveAccent, "#7C3AED", 0.45),
      mix(effectiveAccent, "#EA580C", 0.45)
    ];
  }

  // Every ramp entry is a card top-bar a few points tall. A theme accent that
  // is a near-white wash - or a padded mix that drifted into one - is an
  // invisible bar, so each entry has to stay visible against the canvas.
  accentRamp = accentRamp.map((c) => ensureContrast(c, canvas, 1.6) || c);

  const themeSource = monochrome
    ? "deck reads monochrome"
    : realAccents.length > 0
      ? "office theme"
      : (sample.sampled ? "sampled from deck" : "neutral defaults");

  return {
    themeSource,
    paletteNote,
    themeFingerprint: fingerprint,
    canvasColor: canvas,

    // Fonts are intentionally nullable: when null we do NOT set font.name and
    // PowerPoint applies the presentation theme font (true theme inheritance).
    headingFont: sample.headingFont || null,
    bodyFont: sample.bodyFont || sample.headingFont || null,

    isDark,
    accentColor: effectiveAccent,
    accentSoft: isDark ? shade(effectiveAccent, 0.6) : tint(effectiveAccent, 0.88),
    onAccentColor: readableTextOn(effectiveAccent),

    titleColor,
    bodyColor,
    mutedColor,

    cardFill,
    cardBorder,
    shadowColor,
    onCardColor: readableTextOn(cardFill),

    // Accent ramp used to differentiate sibling cards while staying on-theme.
    accentRamp,

    slideWidth: sample.slideWidth || 960,
    slideHeight: sample.slideHeight || 540,
    sampled: Boolean(sample.sampled),

    // Diagnostic only - never read by any renderer. Records WHERE each sampled
    // value came from, so reading our own generated slides back is visible in
    // the troubleshooting console.
    provenance: sample.provenance || null,
    backgroundEvidence: sample.backgroundEvidence || null
  };
}

/* ------------------------------------------------------------------ */
/* Theme probe (diagnostic only)                                       */
/* ------------------------------------------------------------------ */

const PROBE_COLORS = [
  "Accent1", "Accent2", "Accent3", "Accent4", "Accent5", "Accent6",
  "Dark1", "Dark2", "Light1", "Light2"
];

/**
 * Reads the REAL Office theme colour scheme from the first slide master.
 *
 * `SlideMaster.themeColorScheme.getThemeColor()` is PowerPointApi 1.10 and is
 * the authoritative theme - the same values PowerPoint's Design tab shows. It
 * works on a completely empty deck and, unlike the text sampler, can never read
 * our own generated slides back.
 *
 * This is deliberately DIAGNOSTIC ONLY for now: nothing consumes the result.
 * It never throws and runs in its own PowerPoint.run so it cannot disturb slide
 * generation.
 *
 * @returns {Promise<{available: boolean, reason?: string, colors?: object}>}
 */
export async function probeThemeColorScheme() {
  try {
    const supported = Boolean(
      typeof Office !== "undefined" &&
      Office.context &&
      Office.context.requirements &&
      Office.context.requirements.isSetSupported("PowerPointApi", "1.10")
    );
    if (!supported) {
      return { available: false, reason: "PowerPointApi 1.10 is not supported by this host" };
    }
    if (typeof PowerPoint === "undefined" || !PowerPoint.run) {
      return { available: false, reason: "PowerPoint.run is unavailable" };
    }

    return await PowerPoint.run(async (context) => {
      const masters = context.presentation.slideMasters;
      if (!masters || typeof masters.getItemAt !== "function") {
        return { available: false, reason: "slideMasters.getItemAt is unavailable" };
      }

      const master = masters.getItemAt(0);
      const scheme = master && master.themeColorScheme;
      if (!scheme || typeof scheme.getThemeColor !== "function") {
        return { available: false, reason: "themeColorScheme is unavailable on the slide master" };
      }

      const pending = [];
      for (const name of PROBE_COLORS) {
        try {
          pending.push({ name, result: scheme.getThemeColor(name) });
        } catch (queueErr) {
          console.warn(`[DeckTheme] probe could not queue ${name}:`, queueErr);
        }
      }
      if (pending.length === 0) {
        return { available: false, reason: "no theme colours could be queued" };
      }

      await context.sync();

      const colors = {};
      for (const entry of pending) {
        try {
          const raw = entry.result && entry.result.value;
          const hex = normalizeHex(raw);
          if (hex) colors[entry.name] = hex;
        } catch (readErr) {
          console.warn(`[DeckTheme] probe could not read ${entry.name}:`, readErr);
        }
      }

      if (Object.keys(colors).length === 0) {
        return { available: false, reason: "the host returned no theme colours" };
      }

      // The colour scheme cannot tell us whether the deck is dark: Office keeps
      // Light1=#FFFFFF / Dark1=#000000 even for dark designs, because the
      // darkness lives in the master's background fill. Read it in a SEPARATE
      // sync so a failure here cannot cost us the colours we already have.
      let background = null;
      try {
        const bg = master.background;
        if (bg && bg.fill) {
          bg.fill.load("type");
          const solid = typeof bg.fill.getSolidFillOrNullObject === "function"
            ? bg.fill.getSolidFillOrNullObject()
            : null;
          if (solid) solid.load("color,isNullObject");
          await context.sync();

          background = { type: bg.fill.type || "Unknown", color: null };
          if (solid && !solid.isNullObject) {
            background.color = normalizeHex(solid.color);
          }
        }
      } catch (bgErr) {
        console.warn("[DeckTheme] probe could not read the master background:", bgErr);
      }

      // The master is not authoritative: a layout can override it, and one real
      // deck reported Solid #FFFFFF for a design that renders dark maroon.
      // Slide.background (1.10) is what the audience actually sees. Own sync
      // again, for the same reason as the master read above.
      let slideBackground = null;
      const slideBackgrounds = [];
      try {
        const slides = context.presentation.slides;
        const counter = slides.getCount();
        await context.sync();

        const total = Math.min(Number(counter.value) || 0, 3);
        const probes = [];
        for (let i = 0; i < total; i++) {
          const slide = slides.getItemAt(i);
          const slideBg = slide && slide.background;
          if (!slideBg || !slideBg.fill) continue;
          slideBg.fill.load("type");
          const slideSolid = typeof slideBg.fill.getSolidFillOrNullObject === "function"
            ? slideBg.fill.getSolidFillOrNullObject()
            : null;
          if (slideSolid) slideSolid.load("color,isNullObject");
          probes.push({ index: i, fill: slideBg.fill, solid: slideSolid });
        }

        if (probes.length > 0) {
          await context.sync();
          // Read them ALL. One slide is not evidence: a deck's cover slide
          // routinely has a background no other slide shares, and treating it
          // as the canvas inverted an entire palette.
          for (const entry of probes) {
            const type = entry.fill.type || "Unknown";
            const color = entry.solid && !entry.solid.isNullObject
              ? normalizeHex(entry.solid.color)
              : null;
            slideBackgrounds.push({ index: entry.index, type, color });
            if (!slideBackground || (!slideBackground.color && color)) {
              slideBackground = { index: entry.index, type, color };
            }
          }
        }
      } catch (slideBgErr) {
        console.warn("[DeckTheme] probe could not read a slide background:", slideBgErr);
        // Say so. Silence here left a real deck's dark canvas undiagnosed for
        // a whole round.
        slideBackground = {
          index: 0,
          type: "unavailable",
          color: null,
          error: String((slideBgErr && slideBgErr.message) || slideBgErr).slice(0, 120)
        };
      }

      return { available: true, colors, background, slideBackground, slideBackgrounds };
    });
  } catch (err) {
    return { available: false, reason: String((err && err.message) || err) };
  }
}

/* ------------------------------------------------------------------ */
/* Live sampling from the open presentation                            */
/* ------------------------------------------------------------------ */

let cachedTokens = null;
let cachedAt = 0;
const CACHE_TTL_MS = 60 * 1000;

/* ------------------------------------------------------------------ */
/* Raw sample cache                                                    */
/*                                                                     */
/* Sampling the deck is the expensive part: two PowerPoint.run          */
/* contexts and eight syncs, measured at 8-11 seconds on a 17-slide     */
/* deck. Deriving tokens from a sample is pure arithmetic.              */
/*                                                                     */
/* The builder used to force a SECOND full sample purely to apply the   */
/* canvas hint from the chosen layout - but the hint is applied after   */
/* the I/O, to data the first sample already collected. Keeping the raw */
/* sample lets the second call re-derive the same tokens for free.      */
/*                                                                     */
/* Reuse is deliberately conservative: it requires an explicit opt-in   */
/* from the caller AND a sample primed within this build. Anything else */
/* re-reads the deck.                                                   */
/* ------------------------------------------------------------------ */

let cachedRawSample = null;
let cachedProbe = null;
let cachedSampleAt = 0;
let samplePrimed = false;
const SAMPLE_REUSE_TTL_MS = 60 * 1000;

export function resetDeckTheme() {
  cachedTokens = null;
  cachedAt = 0;
  cachedRawSample = null;
  cachedProbe = null;
  cachedSampleAt = 0;
  samplePrimed = false;
}

/**
 * Synchronous read of the last sampled theme, or null if nothing is cached yet.
 *
 * Needed by the slide parser, which renders charts during a synchronous DOM walk
 * and therefore cannot await getDeckTheme().
 */
export function getCachedDeckTheme() {
  return cachedTokens;
}

/**
 * Reads fonts, colors and geometry from the first few existing slides.
 * Never throws - always resolves to a usable token set.
 */
/**
 * The colour every readable slide background agrees on, or null.
 *
 * Needs at least two slides: one is a coincidence, and on most decks slide 1
 * is a cover with a background nothing else shares.
 */
export function unanimousBackgroundColor(backgrounds) {
  const colors = (backgrounds || []).map((b) => b && b.color).filter(Boolean);
  if (colors.length < 2) return null;
  return colors.every((c) => c === colors[0]) ? colors[0] : null;
}

export async function getDeckTheme(force = false, options = {}) {
  const now = Date.now();
  if (!force && cachedTokens && now - cachedAt < CACHE_TTL_MS) {
    return cachedTokens;
  }

  // Can we answer from the sample we already paid for? Only the caller knows
  // whether its tokens must reflect a re-read of the deck, so this is opt-in.
  const canReuse = Boolean(
    options.reuseSample &&
    samplePrimed &&
    cachedRawSample &&
    now - cachedSampleAt < SAMPLE_REUSE_TTL_MS
  );

  let sample;
  let probe;

  if (canReuse) {
    // Shallow copy: the merge below writes onto the sample, and the cached
    // original has to stay pristine for the next derive.
    sample = Object.assign({}, cachedRawSample);
    probe = cachedProbe || { available: false, reason: "not attempted" };
  } else {
    sample = { sampled: false };
    try {
      sample = await sampleOpenPresentation();
    } catch (err) {
      console.warn("[DeckTheme] Falling back to neutral tokens:", err);
      sample = { sampled: false };
    }

    // The authoritative theme, when the host can give it to us. Probed here
    // rather than in the builder so the parser's cached tokens agree.
    probe = { available: false, reason: "not attempted" };
    try {
      probe = await probeThemeColorScheme();
    } catch (probeErr) {
      probe = { available: false, reason: String((probeErr && probeErr.message) || probeErr) };
    }

    // Stored BEFORE the merge below, so a reuse re-runs identical arithmetic
    // rather than compounding it.
    cachedRawSample = Object.assign({}, sample);
    cachedProbe = probe;
    cachedSampleAt = now;
    samplePrimed = true;
  }

  if (probe.available) {
    sample.themeColors = probe.colors;
    // Canvas priority:
    //   1. the layout we are about to build on (applied below, outranks all)
    //   2. the slide backgrounds, but ONLY when they agree with each other
    //   3. the slide master
    //   4. Light1, i.e. white, i.e. black text
    //
    // Rung 2 needs the agreement clause because a cover slide's background is
    // routinely unique to it. A maroon cover in front of white body slides was
    // read as "this deck is dark" and turned every title white on white.
    const agreed = unanimousBackgroundColor(probe.slideBackgrounds);
    sample.backgroundEvidence = {
      slides: (probe.slideBackgrounds || []).map((s) => s.color || s.type || "?"),
      agreed,
      master: (probe.background && probe.background.color) || null
    };

    if (agreed) {
      sample.canvasColor = agreed;
    } else if (probe.background && probe.background.color) {
      sample.canvasColor = probe.background.color;
    }
  }

  // The layout we are about to build on beats every other reading: it is the
  // surface the slides will literally sit on.
  const hint = normalizeHex(options.canvasHint || "");
  if (hint) sample.canvasColor = hint;

  cachedTokens = buildTokens(sample);
  cachedTokens.probe = probe;
  cachedAt = now;
  return cachedTokens;
}

async function sampleOpenPresentation() {
  if (typeof PowerPoint === "undefined" || !PowerPoint.run) {
    return { sampled: false };
  }

  return await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    const countRes = slides.getCount();
    await context.sync();

    const total = countRes.value || 0;
    if (total === 0) return { sampled: false };

    // ---- which slides are the customer's? ------------------------------
    // Ours are tagged. Reading them would mean deriving the deck's colours
    // from colours we chose last run, which is how the same deck produced a
    // monochrome palette at 09:15 and a hot pink one at 09:17. Look a little
    // way past our work to find the customer's own design.
    const lookAhead = Math.min(total, 12);
    // The head decides whose slides are whose. The TAIL is where the newest
    // recorded palette lives, because our slides are appended - in a 35-slide
    // deck, reading only the first twelve missed every palette we ever wrote.
    const tagIndexes = [];
    for (let i = 0; i < lookAhead; i++) tagIndexes.push(i);
    for (let i = Math.max(lookAhead, total - 6); i < total; i++) tagIndexes.push(i);

    const tagTrackers = [];
    for (const i of tagIndexes) {
      try {
        const tags = slides.getItemAt(i).tags;
        if (!tags || typeof tags.load !== "function") continue;
        tags.load("items/key,items/value");
        tagTrackers.push({ index: i, tags });
      } catch (_) {}
    }

    // Query presentation pageSetup (PowerPointApi 1.10) for true slide dimensions.
    let pageSetupWidth = null;
    let pageSetupHeight = null;
    try {
      const pageSetup = context.presentation.pageSetup;
      if (pageSetup && typeof pageSetup.load === "function") {
        pageSetup.load("slideWidth,slideHeight");
      }
    } catch (_) {}

    const generated = new Set();
    let recordedPalette = null;
    let recordedPaletteIndex = -1;
    if (tagTrackers.length > 0) {
      try {
        await context.sync();
        for (const tracker of tagTrackers) {
          const items = (tracker.tags && tracker.tags.items) || [];
          for (const item of items) {
            const key = String((item && item.key) || "").toUpperCase();
            if (key === ADDIN_SLIDE_TAG) {
              generated.add(tracker.index);
            } else if (key === ADDIN_PALETTE_TAG && tracker.index >= recordedPaletteIndex) {
              // Latest wins. The most recent run is the look a new slide has
              // to sit next to.
              recordedPalette = String((item && item.value) || "");
              recordedPaletteIndex = tracker.index;
            }
          }
        }
      } catch (tagErr) {
        // Pre-1.3 hosts have no tags. Fall back to reading from the top.
        console.warn("[DeckTheme] Slide tags unavailable:", tagErr);
        generated.clear();
        recordedPalette = null;
        recordedPaletteIndex = -1;
      }
    }

    try {
      const ps = context.presentation.pageSetup;
      if (ps && Number.isFinite(ps.slideWidth) && ps.slideWidth > 0 &&
          Number.isFinite(ps.slideHeight) && ps.slideHeight > 0) {
        pageSetupWidth = ps.slideWidth;
        pageSetupHeight = ps.slideHeight;
      }
    } catch (_) {}

    const scanIndexes = [];
    for (let i = 0; i < total && scanIndexes.length < 3; i++) {
      if (!generated.has(i)) scanIndexes.push(i);
    }
    // A deck made entirely of our slides still has to be read from something.
    if (scanIndexes.length === 0) {
      for (let i = 0; i < Math.min(total, 3); i++) scanIndexes.push(i);
    }

    const scanCount = scanIndexes.length;
    const skippedGenerated = generated.size;
    const fontTrackers = [];
    const geometryTrackers = [];
    // Diagnostic: remembers which slide each shape collection belongs to so the
    // probe can report that an accent came from a slide WE generated.
    const trackerSlideIndex = new Map();

    for (const i of scanIndexes) {
      try {
        const slide = slides.getItemAt(i);
        const shapes = slide.shapes;
        shapes.load("items/id,items/type,items/left,items/top,items/width,items/height");
        geometryTrackers.push(shapes);
        trackerSlideIndex.set(shapes, i);
      } catch (_) {}
    }

    try {
      await context.sync();
    } catch (geoErr) {
      console.warn("[DeckTheme] Shape geometry sync notice:", geoErr);
      return { sampled: false };
    }

    // Estimate the slide canvas from the widest shape extent found in the deck.
    // Exclude slides generated by this add-in to prevent feedback poisoning.
    let maxRight = 0;
    let maxBottom = 0;
    for (const shapes of geometryTrackers) {
      if (!shapes.items) continue;
      const slideIdx = trackerSlideIndex.get(shapes);
      if (typeof slideIdx === "number" && generated.has(slideIdx)) {
        // Skip measuring our own generated slides if customer slides exist.
        if (scanIndexes.some((idx) => !generated.has(idx))) continue;
      }
      for (const shape of shapes.items) {
        try {
          const right = (shape.left || 0) + (shape.width || 0);
          const bottom = (shape.top || 0) + (shape.height || 0);
          if (Number.isFinite(right)) maxRight = Math.max(maxRight, right);
          if (Number.isFinite(bottom)) maxBottom = Math.max(maxBottom, bottom);
        } catch (_) {}
      }
    }

    // Queue font reads for every text-bearing shape we found.
    for (const shapes of geometryTrackers) {
      if (!shapes.items) continue;
      for (const shape of shapes.items) {
        try {
          if (!shape.textFrame) continue;
          const tr = shape.textFrame.textRange;
          tr.load("text");
          tr.font.load("name,size,color,bold");
          fontTrackers.push({
            tr,
            height: shape.height || 0,
            slideIndex: trackerSlideIndex.has(shapes) ? trackerSlideIndex.get(shapes) : -1
          });
        } catch (_) {}
      }
    }

    if (fontTrackers.length === 0) {
      return finalizeSample({ maxRight, maxBottom, pageSetupWidth, pageSetupHeight, entries: [], scanCount, skippedGenerated, recordedPalette, recordedPaletteIndex });
    }

    try {
      await context.sync();
    } catch (fontErr) {
      console.warn("[DeckTheme] Font sync notice:", fontErr);
      return finalizeSample({ maxRight, maxBottom, pageSetupWidth, pageSetupHeight, entries: [], scanCount, skippedGenerated, recordedPalette, recordedPaletteIndex });
    }

    const entries = [];
    for (const tracker of fontTrackers) {
      try {
        const text = (tracker.tr.text || "").trim();
        if (!text) continue;
        entries.push({
          text,
          name: tracker.tr.font.name || null,
          size: Number(tracker.tr.font.size) || 0,
          color: normalizeHex(tracker.tr.font.color),
          bold: Boolean(tracker.tr.font.bold),
          slideIndex: typeof tracker.slideIndex === "number" ? tracker.slideIndex : -1
        });
      } catch (_) {}
    }

    return finalizeSample({ maxRight, maxBottom, pageSetupWidth, pageSetupHeight, entries, scanCount, skippedGenerated, recordedPalette, recordedPaletteIndex });
  });
}

export function finalizeSample({ maxRight = 0, maxBottom = 0, pageSetupWidth = null, pageSetupHeight = null, entries = [], scanCount = 0, skippedGenerated = 0,
                         recordedPalette = null, recordedPaletteIndex = -1 }) {
  const sample = { sampled: entries.length > 0 };
  // Read off a slide, not inferred from one: buildTokens decides whether it is
  // still current.
  sample.recordedPalette = recordedPalette;
  sample.recordedPaletteIndex = recordedPaletteIndex;
  sample.provenance = {
    scannedSlides: scanCount || 0,
    textShapes: entries.length,
    skippedGenerated: skippedGenerated || 0,
    neutralRatio: null,
    accentFromSlide: null,
    headingFromSlide: null
  };

  // ---- Slide canvas -------------------------------------------------
  // Snap the measured extent to the nearest standard PowerPoint canvas.
  // Standard PowerPoint canvas sizes (in points, 72 pt/in):
  // 1. 720 x 405 : 16:9 widescreen 10" x 5.625" (standard on Mac, Web, and Office 2013+)
  // 2. 960 x 540 : 16:9 widescreen 13.333" x 7.5" (standard on Windows Office 365 default)
  // 3. 720 x 540 : 4:3 standard 10" x 7.5"
  // 4. 1280 x 720 : 16:9 large high-res
  const candidates = [
    { w: 720, h: 405 }, // 16:9 widescreen 10in (Mac, Web, Office 2013+)
    { w: 960, h: 540 }, // 16:9 widescreen 13.333in
    { w: 720, h: 540 }, // 4:3 standard
    { w: 1280, h: 720 } // 16:9 large
  ];

  if (Number.isFinite(pageSetupWidth) && pageSetupWidth > 0 &&
      Number.isFinite(pageSetupHeight) && pageSetupHeight > 0) {
    // Definitive measurement from PowerPointApi 1.10 pageSetup
    sample.slideWidth = Math.round(pageSetupWidth);
    sample.slideHeight = Math.round(pageSetupHeight);
  } else {
    let best = candidates[0];
    if (maxRight > 40) {
      let bestDelta = Infinity;
      for (const c of candidates) {
        // Shapes never exceed the canvas, so penalise candidates smaller than what we saw.
        const delta = maxRight > c.w + 2 ? Infinity : c.w - maxRight;
        if (delta < bestDelta) {
          bestDelta = delta;
          best = c;
        }
      }
    }
    sample.slideWidth = best.w;
    sample.slideHeight = maxBottom > best.h + 2 ? Math.round(maxBottom) : best.h;
  }

  if (entries.length === 0) return sample;

  // ---- Fonts --------------------------------------------------------
  const sorted = [...entries].sort((a, b) => (b.size || 0) - (a.size || 0));
  const largest = sorted[0];
  if (largest && largest.name) {
    sample.headingFont = largest.name;
    sample.provenance.headingFromSlide = largest.slideIndex;
  }

  const nameCounts = new Map();
  for (const e of entries) {
    if (!e.name) continue;
    nameCounts.set(e.name, (nameCounts.get(e.name) || 0) + 1);
  }
  let bodyFont = null;
  let bodyFontCount = 0;
  for (const [name, count] of nameCounts.entries()) {
    if (count > bodyFontCount) {
      bodyFontCount = count;
      bodyFont = name;
    }
  }
  if (bodyFont) sample.bodyFont = bodyFont;

  // ---- Colors -------------------------------------------------------
  // Used for the dark-canvas heuristic. The accent search below walks `entries`
  // directly so it can record which slide the winning colour came from.
  const colors = entries.map((e) => e.color).filter(Boolean);
  if (largest && largest.color) sample.titleColor = largest.color;
  sample.textColors = colors;

  // A deck whose rendered text carries essentially no colour is a monochrome
  // design, and buildTokens must not hand it one of its theme's six accents.
  //
  // This used to demand that EVERY sampled colour be neutral, which made the
  // verdict hostage to a single shape: eighteen black runs and one pink one
  // turned a black-and-white deck into #EF3475. A deck is monochrome when its
  // text is overwhelmingly neutral. Two samples is still the floor - one shape
  // is a coincidence, not a design.
  const neutralCount = colors.filter((c) => saturation(c) < 0.15).length;
  sample.provenance.neutralRatio = colors.length > 0 ? neutralCount / colors.length : null;
  sample.monochromeDeck = colors.length >= 2 && neutralCount / colors.length >= 0.85;

  // Dark canvas heuristic, used ONLY when the host cannot tell us the real
  // background. One cream title on a black cover slide is not evidence that the
  // whole deck is dark, so this now needs at least three colour samples and a
  // strong majority before it will speak up.
  if (colors.length >= 3) {
    const lightCount = colors.filter((c) => luminance(c) > 0.75).length;
    sample.isDark = lightCount / colors.length > 0.75;
  }

  // Accent = the most saturated text color in the deck (brand color), if any.
  let accent = null;
  let accentSlide = null;
  let bestSat = 0.18;
  for (const entry of entries) {
    const c = entry.color;
    if (!c) continue;
    const s = saturation(c);
    const l = luminance(c);
    if (s > bestSat && l > 0.08 && l < 0.92) {
      bestSat = s;
      accent = c;
      accentSlide = entry.slideIndex;
    }
  }
  if (accent) {
    sample.accentColor = accent;
    sample.provenance.accentFromSlide = accentSlide;
  }

  return sample;
}
