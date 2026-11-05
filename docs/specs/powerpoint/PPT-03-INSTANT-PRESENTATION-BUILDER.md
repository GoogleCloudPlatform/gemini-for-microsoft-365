# PPT-03: Instant Presentation Builder & Multi-Slide Deck Engine

## 1. Overview
Enables creating multi-slide presentations directly from a single prompt or executive document outline. The system automatically structures slides, applies clean corporate formatting, and inserts them into PowerPoint.

## 2. Multi-Slide Recognition & UI
- When `detectSlideCount(text) > 1`, `slidePreviewUI.js` generates slide outline cards in the taskpane.
- The primary action button dynamically switches to **➕ Create Presentation (N Slides)**.
- `PPTAdapter.js` iterates through each parsed slide section and creates corresponding PowerPoint slides with headers, content blocks, and accent styling.
