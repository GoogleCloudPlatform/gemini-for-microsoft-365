# PPT-02: Executive Visual Cards & Metric Grids Specification

## 1. Overview
Defines multimodal executive visual cards rendered natively within PowerPoint chat and dynamically inserted onto slides as styled shapes, tables, and metric cards.

## 2. Visual Component Types
1. **Metric Grid (3-Column)**: Displays 3 key executive KPI cards with values, titles, subtitles, and bullet points.
2. **Before / After Transformation**: High-contrast operational comparison cards displaying current vs target states.
3. **Strategic Pillars**: 3-column architecture cards displaying foundational pillars and descriptions.
4. **Process Timeline**: Sequential roadmap cards highlighting milestone stages with directional flow.

## 3. Rendering Pipeline
- JSON visual payloads (`visualType: "metric_grid_3col" | "before_after" | "three_column_cards" | "process_timeline"`) are extracted by `markdownParser.js`.
- `renderExecutiveVisualHtml` renders rich interactive previews in the taskpane.
- `PPTAdapter.insertContent` creates native PowerPoint text frames and grouped shapes on active or new slides.
