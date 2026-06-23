import os
import re
import shutil
import subprocess
import tempfile

BAD_PATTERNS = re.compile(
    r"\blorem\b|\bipsum\b|\bplaceholder\b|\bTODO\b|\[insert|\bx{4,}\b",
    re.IGNORECASE,
)


def _text_qa(path: str) -> list[str]:
    issues = []
    try:
        from pptx import Presentation
    except ImportError:
        return ["python-pptx not installed — text QA skipped"]

    prs = Presentation(path)
    for i, slide in enumerate(prs.slides, start=1):
        texts = []
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        texts.append(run.text)

        combined = " ".join(texts)
        if not combined.strip():
            issues.append(f"Slide {i}: no text content detected")
            continue

        match = BAD_PATTERNS.search(combined)
        if match:
            issues.append(f"Slide {i}: placeholder text near '{match.group()}'")

    return issues


def _convert_to_images(pptx_path: str) -> list[str]:
    """Convert pptx → PDF → JPEG slides. Returns sorted list of image paths."""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    pdftoppm_bin = shutil.which("pdftoppm")
    if not soffice or not pdftoppm_bin:
        return []

    # Use a persistent directory alongside the pptx so paths survive after return
    base = os.path.splitext(os.path.abspath(pptx_path))[0]
    slides_dir = base + "_slides"
    os.makedirs(slides_dir, exist_ok=True)

    # pptx → pdf
    result = subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf", "--outdir", slides_dir,
         os.path.abspath(pptx_path)],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        return []

    pdf_name = os.path.splitext(os.path.basename(pptx_path))[0] + ".pdf"
    pdf_path = os.path.join(slides_dir, pdf_name)
    if not os.path.exists(pdf_path):
        return []

    # pdf → jpegs
    prefix = os.path.join(slides_dir, "slide")
    result = subprocess.run(
        [pdftoppm_bin, "-jpeg", "-r", "120", pdf_path, prefix],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        return []

    images = sorted(
        os.path.join(slides_dir, f)
        for f in os.listdir(slides_dir)
        if f.endswith(".jpg")
    )
    return images


def run_qa(path: str) -> tuple[list[str], list[str]]:
    """Returns (text_issues, slide_image_paths)."""
    issues = _text_qa(path)
    slide_images = _convert_to_images(path)
    return issues, slide_images
