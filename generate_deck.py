#!/usr/bin/env python3
import argparse
import os
import sys


def main():
    parser = argparse.ArgumentParser(
        description="Generate a Bear Nordic branded .pptx from .xlsx and/or .pdf inputs."
    )
    parser.add_argument(
        "--input",
        action="append",
        metavar="FILE",
        required=True,
        help="Input file (.xlsx or .pdf). May be passed multiple times.",
    )
    parser.add_argument(
        "--output",
        default="deck.pptx",
        metavar="FILE",
        help="Output .pptx path (default: deck.pptx)",
    )
    args = parser.parse_args()

    if not os.environ.get("ANTHROPIC_API_KEY"):
        print("Error: ANTHROPIC_API_KEY environment variable is not set.", file=sys.stderr)
        sys.exit(1)

    # Validate inputs
    for path in args.input:
        if not os.path.isfile(path):
            print(f"Error: input file not found: {path}", file=sys.stderr)
            sys.exit(1)
        ext = os.path.splitext(path)[1].lower()
        if ext not in (".xlsx", ".pdf"):
            print(f"Error: unsupported file type '{ext}' for {path}", file=sys.stderr)
            sys.exit(1)

    # Read inputs
    from src.readers.xlsx_reader import read_xlsx
    from src.readers.pdf_reader import read_pdf

    content_parts = []
    for path in args.input:
        ext = os.path.splitext(path)[1].lower()
        print(f"Reading {path} …")
        if ext == ".xlsx":
            content_parts.append(read_xlsx(path))
        else:
            content_parts.append(read_pdf(path))

    # Structure content via Claude
    from src.content_structurer import structure_content

    print("Generating slide outline via Claude …")
    outline = structure_content(content_parts)

    # Render deck
    from src.deck_renderer import render_deck

    print("Rendering deck …")
    render_deck(outline, args.output)

    # QA
    from src.qa import run_qa

    print("Running QA …")
    issues = run_qa(args.output)
    if issues:
        print("QA warnings:")
        for issue in issues:
            print(f"  • {issue}")
    else:
        print("QA passed.")

    print(args.output)


if __name__ == "__main__":
    main()
