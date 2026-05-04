"""
Shared utilities for gen_svg_slide.

- Image extraction from PPTX shapes
- SVG validation (lightweight pre-check before full quality checker)
- File naming and path helpers
"""

from __future__ import annotations

import re
import base64
from pathlib import Path
from xml.etree import ElementTree as ET


def extract_images_from_parsed_slide(parsed, output_dir: Path) -> list[Path]:
    """Save all embedded images from a ParsedSlide to disk.

    Args:
        parsed: A ``ParsedSlide`` from ``pptx_parser.parse_pptx_slide()``.
        output_dir: Directory to write image files into.

    Returns:
        List of Paths to the saved image files.
    """
    from ..extractor.pptx_parser import ParsedSlide

    output_dir.mkdir(parents=True, exist_ok=True)
    saved: list[Path] = []

    for i, shape in enumerate(parsed.image_shapes):
        if shape.image_blob:
            ext = shape.image_ext or "png"
            fname = output_dir / f"image_{i + 1}.{ext}"
            fname.write_bytes(shape.image_blob)
            saved.append(fname)

    return saved


def validate_svg_quick(svg_content: str) -> tuple[bool, list[str]]:
    """Lightweight SVG validation before running the full quality checker.

    Checks for:
      - Valid XML parse
      - Presence of <svg> root
      - Banned features (mask, style, foreignObject)
      - XML entity issues

    Args:
        svg_content: SVG markup as a string.

    Returns:
        (is_valid, list_of_error_messages) tuple.
    """
    errors: list[str] = []

    # XML parse check
    try:
        root = ET.fromstring(svg_content)
    except ET.ParseError as e:
        return False, [f"XML parse error: {e}"]

    # Root element check
    if not root.tag.endswith("svg"):
        errors.append("Root element is not <svg>")

    # Banned features scan (quick regex)
    banned_patterns = [
        (r"<mask\b", "<mask> is banned — use gradient overlay or clipPath"),
        (r"<style\b", "<style> is banned — use inline attributes"),
        (r"<foreignObject\b", "<foreignObject> is banned"),
        (r"<textPath\b", "<textPath> is banned"),
        (r"<animate\b", "<animate> is banned"),
        (r"rgba?\(", "rgba() is banned — use hex + fill-opacity"),
        (r"&mdash;|&nbsp;|&copy;|&reg;|&rarr;|&middot;|&hellip;|&bull;",
         "HTML named entities found — use raw Unicode characters"),
    ]
    for pattern, msg in banned_patterns:
        if re.search(pattern, svg_content, re.IGNORECASE):
            errors.append(msg)

    return len(errors) == 0, errors


def embed_images_base64(svg_content: str, images_dir: Path) -> str:
    """Replace relative image hrefs with Base64 data URIs.

    Matches the pattern from ppt-master's ``svg_finalize/align_embed_images.py``.

    Args:
        svg_content: SVG markup string.
        images_dir: Directory containing image files.

    Returns:
        SVG string with images embedded as Base64.
    """
    def _replace(match: re.Match) -> str:
        filename = match.group(1)
        img_path = images_dir / filename
        if not img_path.exists():
            # Try without path prefix
            img_path = images_dir / Path(filename).name
        if img_path.exists():
            ext = img_path.suffix.lower().lstrip(".")
            mime_map = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg",
                         "gif": "image/gif", "webp": "image/webp", "svg": "image/svg+xml"}
            mime = mime_map.get(ext, "image/png")
            b64 = base64.b64encode(img_path.read_bytes()).decode()
            return f'href="data:{mime};base64,{b64}"'
        return match.group(0)

    return re.sub(
        r'href="[^"]*?([^/"]+\.(?:png|jpg|jpeg|gif|webp|svg))"',
        _replace, svg_content,
    )


def make_run_dir(run_id: str | None = None) -> Path:
    """Create a timestamped output directory for a generation run.

    Args:
        run_id: Optional run identifier. Auto-generated from timestamp if None.

    Returns:
        Path to the created directory.
    """
    from datetime import datetime
    from ..config import OUTPUT_ROOT

    if run_id is None:
        run_id = datetime.now().strftime("%Y%m%d_%H%M%S")

    run_dir = OUTPUT_ROOT / run_id
    run_dir.mkdir(parents=True, exist_ok=True)
    return run_dir
