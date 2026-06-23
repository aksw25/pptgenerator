# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

CLI tool that reads `.xlsx` and/or `.pdf` files and generates a polished `.pptx` styled to the Bear Nordic brand. No API key required — Claude Code itself reads the extracted content and writes the slide outline.

## Setup

```bash
# Python dependencies (Python 3.12+)
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt

# Node.js dependencies (Node.js 18+)
npm install
```

No API keys or credentials are needed.

## Two-step workflow

**Step 1 — extract content from source files:**
```bash
.venv/bin/python3 generate_deck.py --input report.xlsx --input notes.pdf --extract-only
# → writes extracted_content.json
```

**Step 2a — Claude Code reads extracted_content.json and writes outline.json**
Claude Code synthesises the content into a structured JSON slide outline
(title slide, agenda, themed sections, data highlights, closing slide) and
writes it to `outline.json`.

**Step 2b — render outline.json into a deck:**
```bash
.venv/bin/python3 generate_deck.py --outline outline.json --output deck.pptx
# → writes deck.pptx, runs text QA, converts to images for visual review
```

**Step 3 — visual QA**
The render step prints paths to slide JPEG images (if LibreOffice + pdftoppm
are available). Claude Code reads each image and verifies: no overflow, no
overlapping elements, no placeholder content, Bear Nordic brand applied.

## Architecture

```
generate_deck.py            CLI (--extract-only | --outline)
src/
  readers/
    xlsx_reader.py          pandas → sheet/column/row/summary dict
    pdf_reader.py           pdfplumber → page/text/table dict
  deck_renderer.py          pptxgenjs JS script → node → rezip
  qa.py                     text QA + LibreOffice image conversion
```

## outline.json schema

```json
{
  "deck_title": "string",
  "deck_subtitle": "string",
  "agenda_items": ["string"],
  "sections": [
    {
      "title": "string",
      "key_points": ["string (≤12 words each)"],
      "layout_hint": "bullets | two_column | stat_callout"
    }
  ],
  "data_highlights": [
    {"label": "string", "value": "string", "unit": "string", "context": "string"}
  ],
  "closing_message": "string"
}
```

## Bear Nordic palette
`2C2C2C` near-black · `4B5320` olive · `C26A4A` terracotta · `EAE4D7` off-white · `7A715C` warm taupe.
All hex values passed to pptxgenjs must omit the `#` prefix.

## pptxgenjs pitfalls
- Never prefix hex colours with `#`
- Never reuse option objects across `addText`/`addShape` calls (pptxgenjs mutates in-place)
- Use `bullet: true`, not Unicode bullet characters
