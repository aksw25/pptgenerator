#!/usr/bin/env python3
"""
Bear Nordic deck generator — two-step workflow.

Step 1 — extraction:
    python generate_deck.py --input report.xlsx --input notes.pdf --extract-only
    → writes extracted_content.json

Step 2 — rendering (after Claude Code has written outline.json):
    python generate_deck.py --outline outline.json --output deck.pptx
    → renders the deck and runs QA
"""
import argparse
import json
import os
import sys


def cmd_extract(args):
    from src.readers.xlsx_reader import read_xlsx
    from src.readers.pdf_reader import read_pdf

    parts = []
    for path in args.input:
        if not os.path.isfile(path):
            print(f"Error: file not found: {path}", file=sys.stderr)
            sys.exit(1)
        ext = os.path.splitext(path)[1].lower()
        if ext not in (".xlsx", ".pdf"):
            print(f"Error: unsupported file type '{ext}': {path}", file=sys.stderr)
            sys.exit(1)
        print(f"Reading {path} …")
        if ext == ".xlsx":
            parts.append(read_xlsx(path))
        else:
            parts.append(read_pdf(path))

    out = "extracted_content.json"
    with open(out, "w", encoding="utf-8") as f:
        json.dump(parts, f, ensure_ascii=False, indent=2)
    print(f"Extracted content written to {out}")
    print()
    print("Next step: ask Claude Code to read that file and write outline.json, then run:")
    print("  python generate_deck.py --outline outline.json --output deck.pptx")


def cmd_render(args):
    if not os.path.isfile(args.outline):
        print(f"Error: outline file not found: {args.outline}", file=sys.stderr)
        sys.exit(1)

    with open(args.outline, encoding="utf-8") as f:
        outline = json.load(f)

    from src.deck_renderer import render_deck
    print("Rendering deck …")
    render_deck(outline, args.output)

    from src.qa import run_qa
    print("Running QA …")
    issues, slide_images = run_qa(args.output)

    if slide_images:
        print(f"\nSlide images for visual review ({len(slide_images)} slides):")
        for p in slide_images:
            print(f"  {p}")

    if issues:
        print("\nQA warnings:")
        for issue in issues:
            print(f"  • {issue}")
    else:
        print("Text QA passed.")

    print(f"\n{args.output}")


def main():
    parser = argparse.ArgumentParser(
        description="Bear Nordic deck generator.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument(
        "--input", action="append", metavar="FILE",
        help="Input .xlsx or .pdf file (repeatable). Used with --extract-only.",
    )
    parser.add_argument(
        "--extract-only", action="store_true",
        help="Extract content from --input files and write extracted_content.json. Stop there.",
    )
    parser.add_argument(
        "--outline", metavar="FILE",
        help="Path to outline.json (produced by Claude Code). Triggers rendering.",
    )
    parser.add_argument(
        "--output", default="deck.pptx", metavar="FILE",
        help="Output .pptx path (default: deck.pptx).",
    )
    args = parser.parse_args()

    if args.outline:
        cmd_render(args)
    elif args.input or args.extract_only:
        if not args.input:
            print("Error: --extract-only requires at least one --input file.", file=sys.stderr)
            sys.exit(1)
        cmd_extract(args)
    else:
        parser.print_help()
        sys.exit(1)


if __name__ == "__main__":
    main()
