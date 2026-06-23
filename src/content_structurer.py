import json
import os
import anthropic

MODEL = "claude-sonnet-4-6"

SYSTEM_PROMPT = """\
You are a presentation strategist. Given extracted content from one or more source files, \
produce a single coherent slide outline for a polished business presentation.

Return ONLY a valid JSON object — no markdown fences, no prose, nothing else.

Required schema:
{
  "deck_title": "string",
  "deck_subtitle": "string — source/date context",
  "agenda_items": ["string", ...],   // 3–6 items matching sections below
  "sections": [
    {
      "title": "string",
      "key_points": ["string", "string", "string"],  // 3–5 concise points
      "layout_hint": "bullets" | "two_column" | "stat_callout"
    }
  ],
  "data_highlights": [
    {"label": "string", "value": "string", "unit": "string", "context": "string"}
  ],   // 2–4 key figures drawn from spreadsheet or PDF data
  "closing_message": "string"  // one strong takeaway sentence
}

Rules:
- Synthesise across ALL inputs into ONE coherent narrative — do not dump raw rows.
- Keep key_points to ≤12 words each.
- Alternate layout_hint so no two consecutive sections share the same layout.
- data_highlights must cite actual numbers from the source material.
- closing_message should be memorable and forward-looking.
"""


def _format_content(content_parts: list[dict]) -> str:
    parts = []
    for item in content_parts:
        if item["type"] == "xlsx":
            parts.append(f"=== SPREADSHEET: {item['file']} ===")
            for sheet in item["sheets"]:
                parts.append(f"\n-- Sheet: {sheet['name']} --")
                parts.append(f"Columns: {', '.join(sheet['columns'])}")
                parts.append(f"Row count: {sheet['summary'].get('row_count', '?')}")
                numeric_keys = [k for k in sheet["summary"] if k != "row_count"]
                if numeric_keys:
                    for k in numeric_keys[:8]:
                        s = sheet["summary"][k]
                        parts.append(f"  {k}: min={s['min']}, max={s['max']}, mean={s['mean']}")
                if sheet["rows"]:
                    header = " | ".join(sheet["columns"])
                    parts.append(f"\nSample rows (up to 20):\n{header}")
                    for row in sheet["rows"][:20]:
                        parts.append(" | ".join(str(c) for c in row))
        elif item["type"] == "pdf":
            parts.append(f"=== DOCUMENT: {item['file']} ===")
            for page in item["pages"]:
                if page["text"].strip():
                    parts.append(f"\n-- Page {page['page']} --\n{page['text']}")
                for tbl in page["tables"][:2]:
                    if tbl:
                        parts.append(f"\nTable on page {page['page']}:")
                        for row in tbl[:10]:
                            parts.append(" | ".join(row))
    return "\n".join(parts)


def structure_content(content_parts: list[dict]) -> dict:
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise EnvironmentError("ANTHROPIC_API_KEY environment variable is not set.")

    client = anthropic.Anthropic(api_key=api_key)
    user_content = _format_content(content_parts)

    message = client.messages.create(
        model=MODEL,
        max_tokens=4096,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_content}],
    )

    raw = message.content[0].text.strip()
    # Strip accidental markdown fences if present
    if raw.startswith("```"):
        raw = raw.split("\n", 1)[1]
        if raw.endswith("```"):
            raw = raw.rsplit("```", 1)[0]
        raw = raw.strip()

    try:
        outline = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise ValueError(f"Claude returned non-JSON response:\n{raw}") from exc

    return outline
