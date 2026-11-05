# WORD-01: Word Document Chat & In-Document Authoring Specification

## 1. Overview
Provides native, context-aware AI authoring, document chat, summarization, and tone rewriting within Microsoft Word desktop and web.

## 2. Capabilities
- **In-Document Triggers**: Recognizes `@gemini <prompt>` typed directly into paragraphs.
- **Selection Toolbar**: Floating or docked chips to rewrite, professionalize, shorten, bulletize, or tabulate text.
- **Rich OOXML Insertion**: Inserts rich markdown, formatted headings, and HTML tables into Word body or cursor location via `WordAdapter.js`.
