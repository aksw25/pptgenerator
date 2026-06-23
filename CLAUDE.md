# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

CLI tool that reads `.xlsx` and/or `.pdf` files and generates a polished `.pptx` presentation styled to the Bear Nordic brand. Content is summarised by the Claude API (claude-sonnet-4-6) before being rendered via pptxgenjs.

## Setup

```bash
# Python dependencies (requires Python 3.12+)
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt

# Node.js dependencies (requires Node.js 18+)
npm install

# API key — never commit this
export ANTHROPIC_API_KEY=sk-ant-...
```

## Common Commands

```bash
# Generate a deck from one or more inputs
.venv/bin/python3 generate_deck.py --input report.xlsx --output deck.pptx
.venv/bin/python3 generate_deck.py --input report.xlsx --input notes.pdf --output deck.pptx

# Run QA only on an existing pptx
.venv/bin/python3 -c "from src.qa import run_qa; print(run_qa('deck.pptx'))"
```

## Architecture

```
generate_deck.py            CLI entry point (argparse)
src/
  readers/
    xlsx_reader.py          pandas → sheet/column/row/summary dict
    pdf_reader.py           pdfplumber → page/text/table dict
  content_structurer.py     Claude API → JSON slide outline
  deck_renderer.py          pptxgenjs JS script → node → rezip
  qa.py                     text QA (python-pptx) + visual QA (LibreOffice + Claude)
```

**Data flow:** read inputs → format as text → Claude returns JSON outline → deck_renderer builds a temp `.js` script → `node` executes it via pptxgenjs → Python rezips the `.pptx` → QA runs and reports issues.

**Bear Nordic palette** (`2C2C2C` near-black, `4B5320` olive, `C26A4A` terracotta, `EAE4D7` off-white, `7A715C` warm taupe). All hex values passed to pptxgenjs must omit the `#` prefix.

**pptxgenjs pitfalls:** never reuse option objects across calls (pptxgenjs mutates in-place); never prefix hex with `#`; use `bullet: true` not Unicode bullets.
