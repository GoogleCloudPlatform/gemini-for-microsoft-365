# PPT-05: AI Image Generation & Insertion Specification

## 1. Overview
Enables generating high-fidelity visual concept imagery using Gemini Enterprise / Imagen models and inserting them directly onto PowerPoint slides.

## 2. Key Features
- **Create Image Chip**: Dedicated quick-action chip (`🎨 Create Image`) triggers prompt-to-image workflows.
- **Visual Image Cards**: Rendered in taskpane with zoom modal and inspect capabilities.
- **Slide Insertion**: `PPTAdapter.insertContent` embeds images as native PowerPoint slide shapes with proper aspect ratios.
