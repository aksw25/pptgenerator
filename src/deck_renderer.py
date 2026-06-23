import io
import json
import os
import shutil
import subprocess
import sys
import tempfile
import zipfile

# ---------------------------------------------------------------------------
# Bear Nordic palette — no '#' prefix (pptxgenjs requirement)
# ---------------------------------------------------------------------------
NEAR_BLACK = "2C2C2C"
OLIVE = "4B5320"
TERRACOTTA = "C26A4A"
OFF_WHITE = "EAE4D7"
WARM_TAUPE = "7A715C"

SLIDE_W = 10.0   # inches, LAYOUT_16x9
SLIDE_H = 5.625


def _esc(s: str) -> str:
    """Escape a string for embedding in a JS template literal or double-quoted string."""
    return s.replace("\\", "\\\\").replace('"', '\\"').replace("\n", "\\n")


def _js_text_opts(**kwargs) -> str:
    """Render a dict of pptxgenjs text options as a JS object literal."""
    pairs = []
    for k, v in kwargs.items():
        if isinstance(v, bool):
            pairs.append(f"{k}: {str(v).lower()}")
        elif isinstance(v, (int, float)):
            pairs.append(f"{k}: {v}")
        elif isinstance(v, str):
            pairs.append(f'{k}: "{v}"')
        elif isinstance(v, dict):
            pairs.append(f"{k}: {json.dumps(v)}")
        else:
            pairs.append(f"{k}: {v}")
    return "{ " + ", ".join(pairs) + " }"


# ---------------------------------------------------------------------------
# Slide builders — each returns a list of JS statement strings
# ---------------------------------------------------------------------------

def _slide_title(slide_var: str, title: str, subtitle: str) -> list[str]:
    lines = [
        f'let {slide_var} = pres.addSlide();',
        f'{slide_var}.background = {{ color: "{NEAR_BLACK}" }};',
        # Olive horizontal rule below title area
        f'{slide_var}.addShape(pres.ShapeType.rect, {{ x: 0.6, y: 2.55, w: 8.8, h: 0.035, fill: {{ color: "{OLIVE}" }}, line: {{ color: "{OLIVE}", pt: 0 }} }});',
        # Title
        f'{slide_var}.addText("{_esc(title)}", {{ x: 0.6, y: 1.3, w: 8.8, h: 1.1, fontFace: "Calibri", fontSize: 40, bold: true, color: "{OFF_WHITE}", align: "left", valign: "bottom", margin: 0 }});',
        # Subtitle
        f'{slide_var}.addText("{_esc(subtitle)}", {{ x: 0.6, y: 2.7, w: 8.8, h: 0.7, fontFace: "Calibri", fontSize: 18, color: "{WARM_TAUPE}", align: "left", valign: "top", margin: 0 }});',
        # Small terracotta accent square top-left
        f'{slide_var}.addShape(pres.ShapeType.rect, {{ x: 0.6, y: 0.65, w: 0.12, h: 0.12, fill: {{ color: "{TERRACOTTA}" }}, line: {{ color: "{TERRACOTTA}", pt: 0 }} }});',
    ]
    return lines


def _slide_agenda(slide_var: str, agenda_items: list[str], slide_num: int) -> list[str]:
    lines = [
        f'let {slide_var} = pres.addSlide();',
        f'{slide_var}.background = {{ color: "{OFF_WHITE}" }};',
        # Terracotta square + "AGENDA" label
        f'{slide_var}.addShape(pres.ShapeType.rect, {{ x: 0.6, y: 0.55, w: 0.14, h: 0.14, fill: {{ color: "{TERRACOTTA}" }}, line: {{ color: "{TERRACOTTA}", pt: 0 }} }});',
        f'{slide_var}.addText("AGENDA", {{ x: 0.85, y: 0.48, w: 3.0, h: 0.3, fontFace: "Calibri", fontSize: 13, bold: true, color: "{TERRACOTTA}", align: "left", charSpacing: 4 }});',
        # Thin olive rule below label
        f'{slide_var}.addShape(pres.ShapeType.rect, {{ x: 0.6, y: 0.88, w: 8.8, h: 0.025, fill: {{ color: "{OLIVE}" }}, line: {{ color: "{OLIVE}", pt: 0 }} }});',
    ]

    # Agenda items — two columns if > 4 items
    col2 = len(agenda_items) > 4
    left_items = agenda_items if not col2 else agenda_items[: len(agenda_items) // 2 + len(agenda_items) % 2]
    right_items = [] if not col2 else agenda_items[len(left_items):]

    def _item_block(var, items, x, y_start):
        stmts = []
        for j, item in enumerate(items):
            ypos = y_start + j * 0.65
            stmts += [
                f'{var}.addShape(pres.ShapeType.rect, {{ x: {x:.2f}, y: {ypos + 0.08:.2f}, w: 0.11, h: 0.11, fill: {{ color: "{TERRACOTTA}" }}, line: {{ color: "{TERRACOTTA}", pt: 0 }} }});',
                f'{var}.addText("{_esc(item)}", {{ x: {x + 0.22:.2f}, y: {ypos:.2f}, w: {(4.0 if col2 else 8.2):.2f}, h: 0.38, fontFace: "Calibri", fontSize: 16, color: "{NEAR_BLACK}", align: "left", valign: "middle" }});',
            ]
        return stmts

    lines += _item_block(slide_var, left_items, 0.6, 1.1)
    if right_items:
        lines += _item_block(slide_var, right_items, 5.2, 1.1)

    lines += _slide_num_motif(slide_var, slide_num, dark=False)
    return lines


def _slide_section_bullets(slide_var: str, section: dict, slide_num: int, dark: bool) -> list[str]:
    bg = NEAR_BLACK if dark else OFF_WHITE
    title_color = OFF_WHITE if dark else NEAR_BLACK
    body_color = OFF_WHITE if dark else NEAR_BLACK
    rule_color = OLIVE

    title = section["title"]
    points = section.get("key_points", [])

    lines = [
        f'let {slide_var} = pres.addSlide();',
        f'{slide_var}.background = {{ color: "{bg}" }};',
        # Terracotta square + section title
        f'{slide_var}.addShape(pres.ShapeType.rect, {{ x: 0.6, y: 0.52, w: 0.14, h: 0.14, fill: {{ color: "{TERRACOTTA}" }}, line: {{ color: "{TERRACOTTA}", pt: 0 }} }});',
        f'{slide_var}.addText("{_esc(title)}", {{ x: 0.85, y: 0.42, w: 8.5, h: 0.45, fontFace: "Calibri", fontSize: 28, bold: true, color: "{title_color}", align: "left", valign: "middle" }});',
        # Thin olive rule
        f'{slide_var}.addShape(pres.ShapeType.rect, {{ x: 0.6, y: 1.02, w: 8.8, h: 0.025, fill: {{ color: "{rule_color}" }}, line: {{ color: "{rule_color}", pt: 0 }} }});',
    ]

    # Bullet points
    bullet_rows = []
    for pt in points:
        bullet_rows.append(
            f'  {{ text: "{_esc(pt)}", options: {{ bullet: true, breakLine: true, fontSize: 15, fontFace: "Calibri", color: "{body_color}", paraSpaceAfter: 6 }} }}'
        )
    if bullet_rows:
        rows_js = ",\n".join(bullet_rows)
        lines.append(
            f'{slide_var}.addText([\n{rows_js}\n], {{ x: 0.6, y: 1.15, w: 8.8, h: 3.9, valign: "top", margin: 0 }});'
        )

    lines += _slide_num_motif(slide_var, slide_num, dark=dark)
    return lines


def _slide_section_two_column(slide_var: str, section: dict, slide_num: int, dark: bool) -> list[str]:
    bg = NEAR_BLACK if dark else OFF_WHITE
    title_color = OFF_WHITE if dark else NEAR_BLACK
    body_color = OFF_WHITE if dark else NEAR_BLACK

    title = section["title"]
    points = section.get("key_points", [])
    mid = (len(points) + 1) // 2
    left_pts, right_pts = points[:mid], points[mid:]

    lines = [
        f'let {slide_var} = pres.addSlide();',
        f'{slide_var}.background = {{ color: "{bg}" }};',
        f'{slide_var}.addShape(pres.ShapeType.rect, {{ x: 0.6, y: 0.52, w: 0.14, h: 0.14, fill: {{ color: "{TERRACOTTA}" }}, line: {{ color: "{TERRACOTTA}", pt: 0 }} }});',
        f'{slide_var}.addText("{_esc(title)}", {{ x: 0.85, y: 0.42, w: 8.5, h: 0.45, fontFace: "Calibri", fontSize: 28, bold: true, color: "{title_color}", align: "left", valign: "middle" }});',
        f'{slide_var}.addShape(pres.ShapeType.rect, {{ x: 0.6, y: 1.02, w: 8.8, h: 0.025, fill: {{ color: "{OLIVE}" }}, line: {{ color: "{OLIVE}", pt: 0 }} }});',
        # Vertical divider between columns
        f'{slide_var}.addShape(pres.ShapeType.rect, {{ x: 5.05, y: 1.1, w: 0.02, h: 4.0, fill: {{ color: "{OLIVE}" }}, line: {{ color: "{OLIVE}", pt: 0 }} }});',
    ]

    def _col_text(pts, x):
        rows = [
            f'  {{ text: "{_esc(p)}", options: {{ bullet: true, breakLine: true, fontSize: 14, fontFace: "Calibri", color: "{body_color}", paraSpaceAfter: 8 }} }}'
            for p in pts
        ]
        return (
            f'{slide_var}.addText([\n'
            + ",\n".join(rows)
            + f'\n], {{ x: {x:.2f}, y: 1.15, w: 4.2, h: 4.0, valign: "top", margin: 0 }});'
        )

    if left_pts:
        lines.append(_col_text(left_pts, 0.6))
    if right_pts:
        lines.append(_col_text(right_pts, 5.2))

    lines += _slide_num_motif(slide_var, slide_num, dark=dark)
    return lines


def _slide_data_highlight(slide_var: str, data_highlights: list[dict], slide_num: int) -> list[str]:
    lines = [
        f'let {slide_var} = pres.addSlide();',
        f'{slide_var}.background = {{ color: "{NEAR_BLACK}" }};',
        f'{slide_var}.addShape(pres.ShapeType.rect, {{ x: 0.6, y: 0.52, w: 0.14, h: 0.14, fill: {{ color: "{TERRACOTTA}" }}, line: {{ color: "{TERRACOTTA}", pt: 0 }} }});',
        f'{slide_var}.addText("KEY FIGURES", {{ x: 0.85, y: 0.44, w: 5.0, h: 0.3, fontFace: "Calibri", fontSize: 13, bold: true, color: "{TERRACOTTA}", charSpacing: 4 }});',
        f'{slide_var}.addShape(pres.ShapeType.rect, {{ x: 0.6, y: 0.88, w: 8.8, h: 0.025, fill: {{ color: "{OLIVE}" }}, line: {{ color: "{OLIVE}", pt: 0 }} }});',
    ]

    stats = data_highlights[:4]
    n = len(stats)
    card_w = 8.8 / n - 0.15
    y_card = 1.15

    for i, stat in enumerate(stats):
        x_card = 0.6 + i * (8.8 / n)
        label = stat.get("label", "")
        value = stat.get("value", "")
        unit = stat.get("unit", "")
        context = stat.get("context", "")
        value_str = f"{value} {unit}".strip()

        lines += [
            # Card background — olive rounded rect
            f'{slide_var}.addShape(pres.ShapeType.roundRect, {{ x: {x_card:.2f}, y: {y_card:.2f}, w: {card_w:.2f}, h: 3.8, fill: {{ color: "{OLIVE}", transparency: 75 }}, line: {{ color: "{OLIVE}", pt: 1 }}, rectRadius: 0.08 }});',
            # Large value
            f'{slide_var}.addText("{_esc(value_str)}", {{ x: {x_card + 0.1:.2f}, y: {y_card + 0.25:.2f}, w: {card_w - 0.2:.2f}, h: 1.1, fontFace: "Calibri", fontSize: 38, bold: true, color: "{OFF_WHITE}", align: "center" }});',
            # Label
            f'{slide_var}.addText("{_esc(label)}", {{ x: {x_card + 0.1:.2f}, y: {y_card + 1.4:.2f}, w: {card_w - 0.2:.2f}, h: 0.5, fontFace: "Calibri", fontSize: 13, bold: true, color: "{WARM_TAUPE}", align: "center" }});',
            # Context
            f'{slide_var}.addText("{_esc(context)}", {{ x: {x_card + 0.1:.2f}, y: {y_card + 1.95:.2f}, w: {card_w - 0.2:.2f}, h: 1.5, fontFace: "Calibri", fontSize: 11, color: "{WARM_TAUPE}", align: "center", valign: "top" }});',
        ]

    lines += _slide_num_motif(slide_var, slide_num, dark=True)
    return lines


def _slide_closing(slide_var: str, message: str) -> list[str]:
    return [
        f'let {slide_var} = pres.addSlide();',
        f'{slide_var}.background = {{ color: "{NEAR_BLACK}" }};',
        f'{slide_var}.addShape(pres.ShapeType.rect, {{ x: 1.5, y: 3.0, w: 7.0, h: 0.035, fill: {{ color: "{OLIVE}" }}, line: {{ color: "{OLIVE}", pt: 0 }} }});',
        f'{slide_var}.addText("{_esc(message)}", {{ x: 1.0, y: 1.5, w: 8.0, h: 2.0, fontFace: "Calibri", fontSize: 32, bold: false, color: "{OFF_WHITE}", align: "center", valign: "middle" }});',
        # Small terracotta accent
        f'{slide_var}.addShape(pres.ShapeType.rect, {{ x: 4.85, y: 3.2, w: 0.3, h: 0.3, fill: {{ color: "{TERRACOTTA}" }}, line: {{ color: "{TERRACOTTA}", pt: 0 }} }});',
    ]


def _slide_num_motif(slide_var: str, num: int, dark: bool) -> list[str]:
    color = WARM_TAUPE if dark else "B0A99A"
    return [
        f'{slide_var}.addText("{num}", {{ x: 8.5, y: 4.5, w: 1.4, h: 1.0, fontFace: "Calibri", fontSize: 110, bold: true, color: "{color}", transparency: 85, align: "right", valign: "bottom", margin: 0 }});',
    ]


# ---------------------------------------------------------------------------
# Layout cycling: ensure no two consecutive sections share the same layout
# ---------------------------------------------------------------------------
SECTION_LAYOUTS = ["bullets", "two_column", "bullets", "stat_callout", "two_column"]


def _assign_layouts(sections: list[dict]) -> list[str]:
    layouts = []
    for i, sec in enumerate(sections):
        hint = sec.get("layout_hint", "bullets")
        # Override if same as previous to avoid consecutive duplicates
        if layouts and hint == layouts[-1]:
            options = [l for l in ["bullets", "two_column"] if l != hint]
            hint = options[0]
        layouts.append(hint)
    return layouts


# ---------------------------------------------------------------------------
# Main JS builder
# ---------------------------------------------------------------------------

def _build_js(outline: dict, output_path: str) -> str:
    lines = [
        'const pptxgen = require("pptxgenjs");',
        "let pres = new pptxgen();",
        'pres.layout = "LAYOUT_16x9";',
        f'pres.title = "{_esc(outline.get("deck_title", "Presentation"))}";',
        "",
    ]

    slide_idx = 0
    content_slide_num = 0  # counter for the signature motif

    def sv():
        return f"s{slide_idx}"

    # 1. Title slide
    lines += _slide_title(sv(), outline.get("deck_title", ""), outline.get("deck_subtitle", ""))
    slide_idx += 1
    lines.append("")

    # 2. Agenda slide
    agenda_items = outline.get("agenda_items", [])
    if agenda_items:
        content_slide_num += 1
        lines += _slide_agenda(sv(), agenda_items, content_slide_num)
        slide_idx += 1
        lines.append("")

    sections = outline.get("sections", [])
    layouts = _assign_layouts(sections)
    data_highlights = outline.get("data_highlights", [])

    # Determine position to insert data highlights (after 2nd section or halfway)
    data_insert_after = max(1, len(sections) // 2)

    for i, (sec, layout) in enumerate(zip(sections, layouts)):
        dark = (i % 2 == 0)
        content_slide_num += 1

        if layout == "two_column":
            lines += _slide_section_two_column(sv(), sec, content_slide_num, dark=dark)
        else:
            lines += _slide_section_bullets(sv(), sec, content_slide_num, dark=dark)

        slide_idx += 1
        lines.append("")

        # Insert data highlight slide at the midpoint
        if i == data_insert_after - 1 and data_highlights:
            content_slide_num += 1
            lines += _slide_data_highlight(sv(), data_highlights, content_slide_num)
            slide_idx += 1
            lines.append("")

    # If data highlights were never inserted (no sections), add them now
    if data_highlights and len(sections) == 0:
        content_slide_num += 1
        lines += _slide_data_highlight(sv(), data_highlights, content_slide_num)
        slide_idx += 1
        lines.append("")

    # Closing slide
    closing_message = outline.get("closing_message", "Thank you.")
    lines += _slide_closing(sv(), closing_message)
    slide_idx += 1
    lines.append("")

    # Write file
    escaped_path = output_path.replace("\\", "\\\\")
    lines += [
        f'pres.writeFile({{ fileName: "{escaped_path}" }})',
        '  .then(() => { process.exit(0); })',
        '  .catch((err) => { console.error(err); process.exit(1); });',
    ]

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Rezip: re-compress the pptx as a proper Deflate ZIP
# ---------------------------------------------------------------------------

def _rezip(path: str) -> None:
    buf = io.BytesIO()
    with zipfile.ZipFile(path, "r") as src:
        with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as dst:
            for item in src.infolist():
                data = src.read(item.filename)
                dst.writestr(item, data)
    buf.seek(0)
    with open(path, "wb") as f:
        f.write(buf.read())


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def render_deck(outline: dict, output_path: str) -> None:
    # Resolve to absolute path so pptxgenjs (run from a temp dir) writes the right place
    output_path = os.path.abspath(output_path)

    node_bin = shutil.which("node")
    if not node_bin:
        print(
            "Error: 'node' not found in PATH. Install Node.js to generate the deck.",
            file=sys.stderr,
        )
        sys.exit(1)

    project_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    node_modules = os.path.join(project_dir, "node_modules")
    if not os.path.isdir(node_modules):
        print(
            "Error: node_modules not found. Run 'npm install' in the project directory.",
            file=sys.stderr,
        )
        sys.exit(1)

    js_code = _build_js(outline, output_path)

    with tempfile.NamedTemporaryFile(
        suffix=".js", mode="w", delete=False, dir=project_dir
    ) as tmp:
        tmp.write(js_code)
        tmp_path = tmp.name

    try:
        result = subprocess.run(
            [node_bin, tmp_path],
            capture_output=True,
            text=True,
            cwd=project_dir,
        )
        if result.returncode != 0:
            print(f"Error: pptxgenjs failed:\n{result.stderr}", file=sys.stderr)
            sys.exit(1)
    finally:
        os.unlink(tmp_path)

    _rezip(output_path)
