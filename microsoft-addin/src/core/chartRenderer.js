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
 * Client-Side Canvas Chart Rendering Engine for Microsoft Office Add-in
 * 
 * Dynamically parses structured chart JSON, Vega-Lite specs, and tabular data
 * and renders high-DPI corporate charts (Pie, Doughnut, Bar, Column, Line)
 * to Base64 PNG images for taskpane display and native PowerPoint/Word/Excel insertion.
 * 
 * @author Sathya AG, Principal Architect, Google
 */

const PALETTE = [
  '#1a73e8', // Google Blue
  '#34a853', // Google Green
  '#fbbc04', // Google Yellow
  '#ea4335', // Google Red
  '#9334e6', // Purple
  '#12b5cb', // Teal / Cyan
  '#fa7b17', // Orange
  '#174ea6', // Navy Dark Blue
  '#e37400', // Amber
  '#a142f4'  // Light Purple
];

/**
 * Parses raw text or code block content to extract chart specification.
 * Supports:
 * 1. Standard chart JSON: { chartType: "pie", title: "...", data: [{ label, value }] }
 * 2. Variant keys: type, chart_type, labels/values, series/categories
 * 3. Array of objects: [{ label: "A", value: 10 }, ...]
 * @param {string} text - Raw string content
 * @returns {Object|null} Parsed chart spec or null
 */
export function parseChartSpec(text) {
  if (!text || typeof text !== 'string') return null;

  const trimmed = text.trim();

  // Try parsing direct JSON
  try {
    const parsed = JSON.parse(trimmed);
    return normalizeChartSpec(parsed);
  } catch (_) {}

  // Try extracting JSON from markdown code fence ```json ... ``` or ```chart ... ```
  const fenceMatch = trimmed.match(/```(?:json|chart|vega|vega-lite)?\s*([\s\S]*?)```/i);
  if (fenceMatch) {
    try {
      const parsed = JSON.parse(fenceMatch[1].trim());
      return normalizeChartSpec(parsed);
    } catch (_) {}
  }

  // Try searching for first { ... } block
  const jsonBlockMatch = trimmed.match(/\{[\s\S]*\}/);
  if (jsonBlockMatch) {
    try {
      const parsed = JSON.parse(jsonBlockMatch[0]);
      return normalizeChartSpec(parsed);
    } catch (_) {}
  }

  return null;
}

/**
 * Normalizes various JSON chart structures into a uniform schema:
 * {
 *   chartType: 'pie' | 'doughnut' | 'bar' | 'column' | 'line',
 *   title: string,
 *   subtitle: string,
 *   data: [{ label: string, value: number, color?: string }]
 * }
 */
export function normalizeChartSpec(obj) {
  if (!obj || typeof obj !== 'object') return null;

  // If top-level is an array: [{ label: "A", value: 10 }]
  if (Array.isArray(obj)) {
    const items = extractDataItems(obj);
    if (items.length > 0) {
      return {
        chartType: 'bar',
        title: 'Data Chart',
        subtitle: '',
        data: items
      };
    }
    return null;
  }

  const rawType = (obj.chartType || obj.chart_type || obj.type || obj.mark?.type || obj.mark || 'bar').toString().toLowerCase();
  let chartType = 'bar';
  if (rawType.includes('pie')) chartType = 'pie';
  else if (rawType.includes('doughnut') || rawType.includes('donut')) chartType = 'doughnut';
  else if (rawType.includes('line') || rawType.includes('trend')) chartType = 'line';
  else if (rawType.includes('column') || rawType.includes('vertical')) chartType = 'column';
  else if (rawType.includes('bar') || rawType.includes('horizontal')) chartType = 'bar';

  const title = obj.title?.text || obj.title || obj.name || 'Data Chart';
  const subtitle = obj.subtitle || obj.description || '';

  // Data can be in obj.data, obj.values, obj.items, obj.series
  let rawData = obj.data || obj.values || obj.items || obj.series || [];
  if (rawData && typeof rawData === 'object' && !Array.isArray(rawData)) {
    // Could be Vega-lite: { data: { values: [...] } }
    if (Array.isArray(rawData.values)) {
      rawData = rawData.values;
    } else if (Array.isArray(obj.labels) && Array.isArray(obj.values)) {
      // { labels: ["A", "B"], values: [10, 20] }
      rawData = obj.labels.map((l, i) => ({ label: l, value: obj.values[i] }));
    } else {
      // Key-value map: { "India": 1430000000, "China": 1410000000 }
      rawData = Object.entries(rawData).map(([k, v]) => ({ label: k, value: v }));
    }
  }

  // Handle { labels: [...], datasets: [{ data: [...] }] } (Chart.js format)
  if (Array.isArray(obj.labels) && Array.isArray(obj.datasets) && obj.datasets[0]?.data) {
    rawData = obj.labels.map((l, i) => ({ label: l, value: obj.datasets[0].data[i] }));
  }

  const dataItems = extractDataItems(rawData);
  if (dataItems.length === 0) return null;

  return {
    chartType,
    title: typeof title === 'string' ? title : 'Data Chart',
    subtitle: typeof subtitle === 'string' ? subtitle : '',
    data: dataItems
  };
}

/**
 * Extracts [{ label, value, color }] from diverse object shapes
 */
function extractDataItems(arr) {
  if (!Array.isArray(arr)) return [];

  const items = [];
  for (const item of arr) {
    if (!item) continue;
    if (typeof item !== 'object') continue;

    // Direct label and value
    const label = item.label || item.name || item.category || item.country || item.pillar || item.item || item.segment || item.x || Object.keys(item)[0];
    let rawVal = item.value !== undefined ? item.value : (item.val !== undefined ? item.val : (item.count !== undefined ? item.count : (item.amount !== undefined ? item.amount : (item.metric !== undefined ? item.metric : item.y))));
    if (rawVal === undefined) {
      const vals = Object.values(item);
      rawVal = vals.length > 1 ? vals[1] : vals[0];
    }


    const numVal = parseNumericValue(rawVal);
    if (label && !isNaN(numVal)) {
      items.push({
        label: String(label).trim(),
        value: numVal,
        color: item.color || null
      });
    }
  }
  return items;
}

function parseNumericValue(val) {
  if (typeof val === 'number') return val;
  if (!val) return NaN;
  const cleaned = String(val).replace(/[$,%BMk\s]/gi, '').trim();
  const num = parseFloat(cleaned);
  if (isNaN(num)) return NaN;
  if (String(val).toLowerCase().includes('b')) return num * 1e9;
  if (String(val).toLowerCase().includes('m')) return num * 1e6;
  if (String(val).toLowerCase().includes('k')) return num * 1e3;
  return num;
}

/**
 * Formats a number with friendly suffixes (e.g. 1.43B, 340M, 45K)
 */
export function formatValue(val) {
  if (val >= 1e9) return (val / 1e9).toFixed(val % 1e9 === 0 ? 0 : 2) + 'B';
  if (val >= 1e6) return (val / 1e6).toFixed(val % 1e6 === 0 ? 0 : 1) + 'M';
  if (val >= 1e3) return (val / 1e3).toFixed(val % 1e3 === 0 ? 0 : 1) + 'K';
  if (Number.isInteger(val)) return val.toLocaleString();
  return val.toFixed(1);
}

/** Native aspect ratio (width / height) of every chart produced by this module. */
export const CHART_ASPECT_RATIO = 900 / 560;

/**
 * Builds a series palette anchored on the deck's own accent color, so generated
 * charts match the customer's template instead of always using Google brand colors.
 * @param {string} accent - Hex accent sampled from the open presentation
 * @returns {string[]} Ordered palette (falls back to the default palette)
 */
export function buildChartPalette(accent) {
  const raw = typeof accent === 'string' ? accent.trim() : '';
  if (!/^#?[0-9a-fA-F]{6}$/.test(raw)) return PALETTE.slice();

  const base = raw[0] === '#' ? raw.toLowerCase() : `#${raw.toLowerCase()}`;
  const r = parseInt(base.substring(1, 3), 16);
  const g = parseInt(base.substring(3, 5), 16);
  const b = parseInt(base.substring(5, 7), 16);

  const part = (n) => Math.max(0, Math.min(255, Math.round(n))).toString(16).padStart(2, '0');
  const blend = (to, ratio) =>
    `#${part(r + (to[0] - r) * ratio)}${part(g + (to[1] - g) * ratio)}${part(b + (to[2] - b) * ratio)}`;

  // Harmonized ramp: the deck accent first, then hue-shifted siblings that stay on-brand.
  return [
    base,
    blend([13, 148, 136], 0.5),   // teal
    blend([234, 88, 12], 0.5),    // amber
    blend([124, 58, 237], 0.5),   // violet
    blend([220, 38, 38], 0.5),    // red
    blend([255, 255, 255], 0.45), // light tint of the accent
    blend([6, 182, 212], 0.55),   // cyan
    blend([15, 23, 42], 0.35)     // deep shade of the accent
  ];
}

/**
 * Renders a chart specification onto an HTML5 Canvas and returns a Base64 PNG data URL.
 *
 * Rendered at 900x560 logical units with a 3x device scale (2700x1680 real pixels).
 * PowerPoint draws a half-slide chart at roughly 1200 physical pixels on a HiDPI
 * display, so anything smaller than this is upscaled by the host and looks soft.
 *
 * @param {Object} spec - Normalized chart spec
 * @param {Object} [options] - width, height, scale, palette
 * @returns {string|null} data:image/png;base64,... or null if canvas unsupported
 */
export function renderChartToDataUrl(spec, options = {}) {
  if (!spec || !spec.data || spec.data.length === 0) return null;
  if (typeof document === 'undefined' || !document.createElement) return null;

  const width = options.width || 900;
  const height = options.height || 560;
  const scale = options.scale || 3;

  const canvas = document.createElement('canvas');
  canvas.width = Math.round(width * scale);
  canvas.height = Math.round(height * scale);

  const ctx = canvas.getContext('2d');
  if (!ctx) return null;

  ctx.scale(scale, scale);
  ctx.imageSmoothingEnabled = true;
  ctx.imageSmoothingQuality = 'high';
  if ('textRendering' in ctx) {
    ctx.textRendering = 'geometricPrecision';
  }

  // Resolve the palette once and bake it onto the data, so every renderer below
  // (bar / column / pie / line) stays consistent with the deck's accent color.
  const palette = (Array.isArray(options.palette) && options.palette.length > 0)
    ? options.palette
    : PALETTE;
  spec = {
    ...spec,
    accent: palette[0],
    data: spec.data.map((d, i) => ({ ...d, color: d.color || palette[i % palette.length] }))
  };

  // Background
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, width, height);

  // Header: Title & Subtitle with dynamic fitting to prevent overflow
  let titleText = (spec.title || 'Data Visualization').trim();
  const maxTitleWidth = width - 72; // 36px padding on left and right

  let titleFontSize = 26;
  ctx.font = `800 ${titleFontSize}px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif`;
  while (ctx.measureText(titleText).width > maxTitleWidth && titleFontSize > 16) {
    titleFontSize -= 2;
    ctx.font = `800 ${titleFontSize}px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif`;
  }
  if (ctx.measureText(titleText).width > maxTitleWidth) {
    while (titleText.length > 0 && ctx.measureText(titleText + '...').width > maxTitleWidth) {
      titleText = titleText.slice(0, -1).trim();
    }
    titleText += '...';
  }

  ctx.fillStyle = '#0f172a';
  ctx.textAlign = 'left';
  ctx.fillText(titleText, 36, 44);

  if (spec.subtitle) {
    let subText = spec.subtitle.trim();
    ctx.fillStyle = '#475569';
    let subFontSize = 15;
    ctx.font = `600 ${subFontSize}px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif`;
    while (ctx.measureText(subText).width > maxTitleWidth && subFontSize > 12) {
      subFontSize -= 1;
      ctx.font = `600 ${subFontSize}px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif`;
    }
    if (ctx.measureText(subText).width > maxTitleWidth) {
      while (subText.length > 0 && ctx.measureText(subText + '...').width > maxTitleWidth) {
        subText = subText.slice(0, -1).trim();
      }
      subText += '...';
    }
    ctx.fillText(subText, 36, 72);
  }

  // Draw chart according to type
  const chartType = (spec.chartType || 'bar').toLowerCase();
  if (chartType === 'pie' || chartType === 'doughnut') {
    drawPieOrDoughnutChart(ctx, spec, width, height, chartType === 'doughnut');
  } else if (chartType === 'line') {
    drawLineChart(ctx, spec, width, height);
  } else if (chartType === 'column') {
    drawColumnChart(ctx, spec, width, height);
  } else {
    drawBarChart(ctx, spec, width, height);
  }

  return canvas.toDataURL('image/png');
}

/**
 * Draws Pie or Doughnut Chart with clean non-overlapping legend and bold typography
 */
function drawPieOrDoughnutChart(ctx, spec, width, height, isDoughnut) {
  const data = spec.data;
  const total = data.reduce((sum, d) => sum + (d.value > 0 ? d.value : 0), 0) || 1;

  const startY = spec.subtitle ? 90 : 70;
  const availableHeight = height - startY - 20;
  const centerX = Math.round(width * 0.32);
  const centerY = Math.round(startY + availableHeight / 2);
  const radius = Math.min(centerX - 30, Math.round(availableHeight / 2) - 10);

  let startAngle = -Math.PI / 2;

  // Draw Slices
  data.forEach((item, idx) => {
    const val = item.value > 0 ? item.value : 0;
    const sliceAngle = (val / total) * (2 * Math.PI);
    const endAngle = startAngle + sliceAngle;
    const color = item.color || PALETTE[idx % PALETTE.length];

    ctx.beginPath();
    ctx.moveTo(centerX, centerY);
    ctx.arc(centerX, centerY, radius, startAngle, endAngle);
    ctx.closePath();
    ctx.fillStyle = color;
    ctx.fill();

    // Clean white slice separator
    ctx.lineWidth = 3;
    ctx.strokeStyle = '#ffffff';
    ctx.stroke();

    startAngle = endAngle;
  });

  // If Doughnut, cutout inner circle
  if (isDoughnut) {
    ctx.beginPath();
    ctx.arc(centerX, centerY, Math.round(radius * 0.54), 0, 2 * Math.PI);
    ctx.fillStyle = '#ffffff';
    ctx.fill();

    // Inner total label
    ctx.fillStyle = '#0f172a';
    ctx.font = '800 24px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
    ctx.textAlign = 'center';
    ctx.fillText('Total', centerX, centerY - 6);
    ctx.font = '700 18px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
    ctx.fillStyle = '#64748b';
    ctx.fillText(formatValue(total), centerX, centerY + 20);
  }

  // Draw Clean Legend on Right Side (Auto single-line or 2-column to prevent vertical crowding)
  const legendX = Math.round(width * 0.58);
  const isMultiCol = data.length > 6;
  const colCount = isMultiCol ? 2 : 1;
  const colWidth = isMultiCol ? Math.round((width - legendX - 16) / 2) : (width - legendX - 16);
  const rowsPerCol = Math.ceil(data.length / colCount);
  const itemHeight = Math.min(46, Math.floor((availableHeight - 10) / rowsPerCol));
  const legendStartY = Math.round(centerY - (rowsPerCol * itemHeight) / 2);

  ctx.textAlign = 'left';

  data.forEach((item, idx) => {
    const colIdx = isMultiCol ? Math.floor(idx / rowsPerCol) : 0;
    const rowIdx = isMultiCol ? (idx % rowsPerCol) : idx;
    const x = Math.round(legendX + colIdx * colWidth);
    const y = Math.round(legendStartY + rowIdx * itemHeight + itemHeight / 2);
    const color = item.color || PALETTE[idx % PALETTE.length];
    const pct = ((item.value / total) * 100).toFixed(1);

    // Color Swatch
    ctx.beginPath();
    ctx.arc(x + 8, y, 7, 0, 2 * Math.PI);
    ctx.fillStyle = color;
    ctx.fill();

    // Clean Single-Line Format with Bold High-Contrast Typography
    ctx.fillStyle = '#0f172a';
    const fontSize = isMultiCol ? 14 : 16;
    ctx.font = `800 ${fontSize}px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif`;
    const maxChars = isMultiCol ? 14 : 22;
    const labelText = item.label.length > maxChars ? item.label.substring(0, maxChars - 1) + '…' : item.label;
    
    if (isMultiCol) {
      ctx.fillText(`${labelText} (${pct}%)`, x + 20, y + 5);
    } else {
      ctx.fillText(labelText, x + 22, y + 5);
      ctx.fillStyle = '#334155';
      ctx.font = `700 ${fontSize}px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif`;
      ctx.fillText(`${formatValue(item.value)} (${pct}%)`, x + 22 + ctx.measureText(labelText + '  ').width, y + 5);
    }
  });
}

/**
 * Draws Horizontal Bar Chart with high-contrast, razor-sharp typography
 */
function drawBarChart(ctx, spec, width, height) {
  const data = spec.data;
  const maxVal = Math.max(...data.map(d => d.value), 1);
  const startY = spec.subtitle ? 92 : 72;
  const availableHeight = height - startY - 24;
  const itemGap = data.length > 8 ? 6 : 10;
  const barHeight = Math.min(36, Math.max(16, Math.floor((availableHeight - (data.length - 1) * itemGap) / data.length)));
  const fontSize = data.length > 8 ? 14 : 15;

  // Measure label widths to dynamically allocate sufficient left margin without clipping
  ctx.font = `800 ${fontSize}px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif`;
  let maxMeasuredLabelWidth = 0;
  data.forEach(item => {
    const rawLabel = String(item.label || '');
    const w = ctx.measureText(rawLabel).width;
    if (w > maxMeasuredLabelWidth) maxMeasuredLabelWidth = w;
  });

  // Keep leftMargin between 180 and 320 to leave ample room for bars and right-side values
  const leftMargin = Math.min(320, Math.max(180, Math.round(maxMeasuredLabelWidth + 30)));
  const maxBarWidth = width - leftMargin - 110;
  const maxLabelAllowedWidth = leftMargin - 28;

  data.forEach((item, idx) => {
    const y = Math.round(startY + idx * (barHeight + itemGap));
    const barWidth = Math.max(8, Math.round((item.value / maxVal) * maxBarWidth));
    const color = item.color || PALETTE[idx % PALETTE.length];

    // Label on Left - dynamically truncated if it exceeds allowed width
    ctx.fillStyle = '#0f172a';
    ctx.font = `800 ${fontSize}px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif`;
    ctx.textAlign = 'right';
    
    let labelText = String(item.label || '');
    if (ctx.measureText(labelText).width > maxLabelAllowedWidth) {
      while (labelText.length > 0 && ctx.measureText(labelText + '…').width > maxLabelAllowedWidth) {
        labelText = labelText.slice(0, -1).trim();
      }
      labelText += '…';
    }
    ctx.fillText(labelText, leftMargin - 14, Math.round(y + barHeight / 2 + 5));

    // Bar background track
    ctx.fillStyle = '#f1f5f9';
    drawRoundedRect(ctx, leftMargin, y, maxBarWidth, barHeight, 5);
    ctx.fill();

    // Bar fill
    ctx.fillStyle = color;
    drawRoundedRect(ctx, leftMargin, y, barWidth, barHeight, 5);
    ctx.fill();

    // Value label on Right
    ctx.fillStyle = '#0f172a';
    ctx.font = `900 ${fontSize}px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif`;
    ctx.textAlign = 'left';
    ctx.fillText(formatValue(item.value), leftMargin + barWidth + 10, Math.round(y + barHeight / 2 + 5));
  });
}

/**
 * Draws Vertical Column Chart
 */
function drawColumnChart(ctx, spec, width, height) {
  const data = spec.data;
  const maxVal = Math.max(...data.map(d => d.value), 1);
  const startY = spec.subtitle ? 90 : 70;
  const bottomY = height - 50;
  const availableHeight = bottomY - startY;
  const leftMargin = 70;
  const availableWidth = width - leftMargin - 50;
  const colWidth = Math.min(50, Math.floor(availableWidth / data.length) - 16);
  const gap = (availableWidth - colWidth * data.length) / (data.length + 1);

  // Horizontal gridlines + value axis give the eye a reference for column heights.
  ctx.lineWidth = 1;
  for (let i = 1; i <= 4; i++) {
    const gridY = Math.round(bottomY - (availableHeight * i) / 4) + 0.5;
    ctx.strokeStyle = '#EDF1F6';
    ctx.beginPath();
    ctx.moveTo(leftMargin - 10, gridY);
    ctx.lineTo(width - 40, gridY);
    ctx.stroke();

    ctx.fillStyle = '#94A3B8';
    ctx.font = '600 12px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
    ctx.textAlign = 'right';
    ctx.fillText(formatValue((maxVal * i) / 4), leftMargin - 18, gridY + 4);
  }

  // Baseline
  ctx.strokeStyle = '#CBD5E1';
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.moveTo(leftMargin - 10, bottomY + 0.5);
  ctx.lineTo(width - 40, bottomY + 0.5);
  ctx.stroke();

  data.forEach((item, idx) => {
    const x = leftMargin + gap + idx * (colWidth + gap);
    const colHeight = Math.max(8, (item.value / maxVal) * availableHeight);
    const y = bottomY - colHeight;
    const color = item.color || PALETTE[idx % PALETTE.length];

    // Column fill
    ctx.fillStyle = color;
    drawRoundedRect(ctx, x, y, colWidth, colHeight, 4);
    ctx.fill();

    // Value above column
    ctx.fillStyle = '#0f172a';
    ctx.font = '800 14px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif';
    ctx.textAlign = 'center';
    ctx.fillText(formatValue(item.value), Math.round(x + colWidth / 2), Math.round(y - 8));

    // Label below column
    ctx.fillStyle = '#0f172a';
    ctx.font = '700 13px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif';
    const labelText = item.label.length > 10 ? item.label.substring(0, 8) + '…' : item.label;
    ctx.fillText(labelText, Math.round(x + colWidth / 2), Math.round(bottomY + 22));
  });
}

/**
 * Draws Line Chart
 */
function drawLineChart(ctx, spec, width, height) {
  const data = spec.data;
  const maxVal = Math.max(...data.map(d => d.value), 1);
  const minVal = Math.min(...data.map(d => d.value), 0);
  const range = maxVal - minVal || 1;

  const startY = spec.subtitle ? 95 : 75;
  const bottomY = height - 55;
  const availableHeight = bottomY - startY;
  const leftMargin = 80;
  const rightMargin = 50;
  const availableWidth = width - leftMargin - rightMargin;
  const stepX = availableWidth / Math.max(data.length - 1, 1);

  // Horizontal Grid Lines
  ctx.strokeStyle = '#f1f3f4';
  ctx.lineWidth = 1;
  for (let i = 0; i <= 4; i++) {
    const gridY = bottomY - (availableHeight * i) / 4;
    ctx.beginPath();
    ctx.moveTo(leftMargin, gridY);
    ctx.lineTo(width - rightMargin, gridY);
    ctx.stroke();

    // Axis label
    ctx.fillStyle = '#80868b';
    ctx.font = '12px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
    ctx.textAlign = 'right';
    const axisVal = minVal + (range * i) / 4;
    ctx.fillText(formatValue(axisVal), leftMargin - 10, gridY + 4);
  }

  // Points coordinates
  const points = data.map((d, i) => {
    const x = leftMargin + i * stepX;
    const y = bottomY - ((d.value - minVal) / range) * availableHeight;
    return { x, y, value: d.value, label: d.label };
  });

  // Series color follows the deck accent (baked onto the spec by the caller).
  const seriesColor = spec.accent || spec.data[0]?.color || PALETTE[0];

  // Draw Area under line (soft vertical gradient in the series color)
  ctx.beginPath();
  ctx.moveTo(points[0].x, bottomY);
  points.forEach(p => ctx.lineTo(p.x, p.y));
  ctx.lineTo(points[points.length - 1].x, bottomY);
  ctx.closePath();
  const areaGradient = ctx.createLinearGradient(0, startY, 0, bottomY);
  areaGradient.addColorStop(0, hexToRgba(seriesColor, 0.22));
  areaGradient.addColorStop(1, hexToRgba(seriesColor, 0.02));
  ctx.fillStyle = areaGradient;
  ctx.fill();

  // Draw Line
  ctx.strokeStyle = seriesColor;
  ctx.lineWidth = 3;
  ctx.lineJoin = 'round';
  ctx.beginPath();
  points.forEach((p, i) => {
    if (i === 0) ctx.moveTo(p.x, p.y);
    else ctx.lineTo(p.x, p.y);
  });
  ctx.stroke();

  // Draw Dots & Labels
  points.forEach(p => {
    // Outer white dot
    ctx.beginPath();
    ctx.arc(p.x, p.y, 6, 0, 2 * Math.PI);
    ctx.fillStyle = '#ffffff';
    ctx.fill();
    ctx.strokeStyle = seriesColor;
    ctx.lineWidth = 2.5;
    ctx.stroke();

    // Value
    ctx.fillStyle = '#202124';
    ctx.font = 'bold 11px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
    ctx.textAlign = 'center';
    ctx.fillText(formatValue(p.value), p.x, p.y - 12);

    // X-Axis Label
    ctx.fillStyle = '#5f6368';
    ctx.font = '600 11px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
    ctx.fillText(p.label, p.x, bottomY + 20);
  });
}

/** Converts #rrggbb to an rgba() string at the requested alpha. */
function hexToRgba(hex, alpha) {
  const raw = typeof hex === 'string' ? hex.replace('#', '') : '';
  if (!/^[0-9a-fA-F]{6}$/.test(raw)) return `rgba(26, 115, 232, ${alpha})`;
  const r = parseInt(raw.substring(0, 2), 16);
  const g = parseInt(raw.substring(2, 4), 16);
  const b = parseInt(raw.substring(4, 6), 16);
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}

function drawRoundedRect(ctx, x, y, width, height, radius) {
  ctx.beginPath();
  ctx.moveTo(x + radius, y);
  ctx.lineTo(x + width - radius, y);
  ctx.quadraticCurveTo(x + width, y, x + width, y + radius);
  ctx.lineTo(x + width, y + height - radius);
  ctx.quadraticCurveTo(x + width, y + height, x + width - radius, y + height);
  ctx.lineTo(x + radius, y + height);
  ctx.quadraticCurveTo(x, y + height, x, y + height - radius);
  ctx.lineTo(x, y + radius);
  ctx.quadraticCurveTo(x, y, x + radius, y);
  ctx.closePath();
}

function escapeHtmlAttr(str) {
  if (!str) return '';
  return String(str).replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/**
 * Convenience method that parses raw text/JSON and renders an HTML image wrapper.
 * @param {string} text - Raw chart JSON string
 * @returns {string|null} HTML markup with <img src="data:image/png;base64,...">
 */
export function renderChartHtml(text) {
  const spec = parseChartSpec(text);
  if (!spec) return null;

  const dataUri = renderChartToDataUrl(spec);
  if (dataUri) {
    const chartTitle = spec.title || 'Generated Chart';
    const alt = spec.title ? `${spec.title} (${spec.chartType} chart)` : 'Generated Chart';
    return `<div class="rendered-chart-container office-visual-image-card" data-card-img-src="${dataUri}" data-chart-title="${escapeHtmlAttr(chartTitle)}" style="margin:18px 0; text-align:center; background:#f8fafc; border:1px solid #cbd5e1; border-radius:8px; padding:10px; box-shadow:0 2px 8px rgba(0,0,0,0.06);">
      <div style="position:relative; display:inline-block; max-width:100%;">
        <img src="${dataUri}" alt="${escapeHtmlAttr(alt)}" class="office-preview-img" style="max-width:100%; border-radius:6px; display:block; cursor:pointer;" title="Click to zoom / review chart" />
      </div>
      <div style="margin-top:8px; display:flex; justify-content:center; gap:6px; flex-wrap:wrap;">
        <button type="button" class="img-action-btn-zoom" data-img-src="${dataUri}" data-img-alt="${escapeHtmlAttr(alt)}" style="background:#ffffff; color:#0b57d0; border:1px solid #e1e3e1; border-radius:9999px; padding:6px 14px; font-family:'Google Sans',sans-serif; font-size:12px; font-weight:600; cursor:pointer; display:inline-flex; align-items:center; gap:5px; box-shadow:0 1px 2px rgba(0,0,0,0.05);">🔍 Zoom & Review</button>
      </div>
    </div>`;
  }

  // Graceful fallback if Canvas is unavailable: render responsive visual HTML table
  const total = spec.data.reduce((sum, d) => sum + (d.value > 0 ? d.value : 0), 0) || 1;
  const rows = spec.data.map((d, i) => {
    const pct = ((d.value / total) * 100).toFixed(1);
    const color = PALETTE[i % PALETTE.length];
    return `<tr>
      <td style="padding:8px 12px; border-bottom:1px solid #e0e0e0;"><span style="display:inline-block; width:10px; height:10px; border-radius:50%; background:${color}; margin-right:8px;"></span>${escapeHtmlAttr(d.label)}</td>
      <td style="padding:8px 12px; border-bottom:1px solid #e0e0e0; font-weight:bold; text-align:right;">${formatValue(d.value)}</td>
      <td style="padding:8px 12px; border-bottom:1px solid #e0e0e0; color:#5f6368; text-align:right;">${pct}%</td>
    </tr>`;
  }).join('');

  return `<div class="chart-table-fallback" style="margin:16px 0; border:1px solid #d2e3fc; border-radius:8px; overflow:hidden;">
    <div style="background:#f8fafd; padding:10px 14px; font-weight:bold; color:#1a73e8; border-bottom:1px solid #d2e3fc;">📊 ${escapeHtmlAttr(spec.title)}</div>
    <table style="width:100%; border-collapse:collapse; font-size:13px;">
      <thead>
        <tr style="background:#fafafa; color:#5f6368; text-align:left;">
          <th style="padding:8px 12px;">Category</th>
          <th style="padding:8px 12px; text-align:right;">Metric</th>
          <th style="padding:8px 12px; text-align:right;">Share</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  </div>`;
}

