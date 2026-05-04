"""
SVG → PPTX Converter — self-contained wrapper around the bundled svg_to_pptx package.

No longer imports from ppt-master; the entire svg_to_pptx pipeline lives at
``converter/svg_to_pptx/`` and ``converter/svg_finalize/``.

Corresponds to **Step 5**: "The SVG slide is turned into PPTX."
"""

from __future__ import annotations

import sys
from pathlib import Path

# ---------------------------------------------------------------------------
# Ensure the converter directory (containing config.py, project_utils.py,
# svg_to_pptx/, and svg_finalize/) is on sys.path so the bundled packages
# can resolve their internal imports (e.g. `from config import CANVAS_FORMATS`).
# ---------------------------------------------------------------------------
_CONVERTER_DIR = Path(__file__).resolve().parent
if str(_CONVERTER_DIR) not in sys.path:
    sys.path.insert(0, str(_CONVERTER_DIR))


# ---------------------------------------------------------------------------
# Public API — delegates to the bundled svg_to_pptx package
# ---------------------------------------------------------------------------

def svg_file_to_pptx(
    svg_path: Path,
    output_path: Path,
    canvas_format: str = "ppt169",
    use_native_shapes: bool = True,
    verbose: bool = True,
) -> bool:
    """Convert a single SVG file into a native-shapes PPTX.

    Uses the bundled ``svg_to_pptx.create_pptx_with_native_svg()`` which:
      1. Parses the SVG → DrawingML slide XML (via drawingml_converter.py)
      2. Builds a .pptx with python-pptx, injecting the DrawingML
      3. Handles embedded images (Base64 → ppt/media/ blips)

    Args:
        svg_path: Path to the SVG file to convert.
        output_path: Where to write the output .pptx.
        canvas_format: 'ppt169', 'ppt43', etc.
        use_native_shapes: If True, produce editable PowerPoint shapes.
        verbose: Print progress information.

    Returns:
        True on success, False on failure.
    """
    from svg_to_pptx import create_pptx_with_native_svg

    if not svg_path.exists():
        print(f"Error: SVG file not found: {svg_path}")
        return False

    return create_pptx_with_native_svg(
        svg_files=[svg_path],
        output_path=output_path,
        canvas_format=canvas_format,
        use_native_shapes=use_native_shapes,
        use_compat_mode=not use_native_shapes,
        verbose=verbose,
        transition=None,
        enable_notes=False,
    )


def svg_string_to_pptx(
    svg_content: str,
    output_path: Path,
    work_dir: Path | None = None,
    canvas_format: str = "ppt169",
    use_native_shapes: bool = True,
    verbose: bool = True,
) -> bool:
    """Convert an SVG string (in memory) into a native-shapes PPTX.

    Writes the SVG to a temporary file first, then calls ``svg_file_to_pptx``.

    Args:
        svg_content: The SVG markup as a string.
        output_path: Where to write the output .pptx.
        work_dir: Working directory for temp SVG (auto-created if None).
        canvas_format: 'ppt169', 'ppt43', etc.
        use_native_shapes: If True, produce editable PowerPoint shapes.
        verbose: Print progress information.

    Returns:
        True on success, False on failure.
    """
    if work_dir is None:
        work_dir = output_path.parent / ".gen_svg_work"
    work_dir.mkdir(parents=True, exist_ok=True)

    svg_temp = work_dir / "slide_01.svg"
    svg_temp.write_text(svg_content, encoding="utf-8")

    return svg_file_to_pptx(
        svg_path=svg_temp,
        output_path=output_path,
        canvas_format=canvas_format,
        use_native_shapes=use_native_shapes,
        verbose=verbose,
    )


def convert_svg_to_drawingml(svg_path: Path) -> str:
    """Low-level: convert an SVG file to DrawingML slide XML.

    Useful for inspection/debugging. Returns the raw DrawingML XML string
    that would be injected into a slide's <p:spTree>.

    Args:
        svg_path: Path to the SVG file.

    Returns:
        DrawingML XML string for the slide's shape tree.
    """
    from svg_to_pptx import convert_svg_to_slide_shapes

    shapes_xml = convert_svg_to_slide_shapes(svg_path, verbose=False)
    return shapes_xml
