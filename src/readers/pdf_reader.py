import os
import pdfplumber

MAX_CHARS_PER_FILE = 4000


def read_pdf(path: str) -> dict:
    pages = []
    total_chars = 0

    with pdfplumber.open(path) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            if total_chars >= MAX_CHARS_PER_FILE:
                break

            text = page.extract_text() or ""
            remaining = MAX_CHARS_PER_FILE - total_chars
            if len(text) > remaining:
                text = text[:remaining]
            total_chars += len(text)

            raw_tables = page.extract_tables() or []
            tables = []
            for tbl in raw_tables:
                cleaned = [[str(cell) if cell is not None else "" for cell in row] for row in tbl]
                tables.append(cleaned)

            pages.append({"page": i, "text": text, "tables": tables})

    return {"file": os.path.basename(path), "type": "pdf", "pages": pages}
