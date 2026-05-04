"""
Minimal project_utils stub for svg_to_pptx/pptx_dimensions.py.

Provides the symbols that pptx_dimensions.py imports from project_utils.
The real project_utils in ppt-master does much more (validation, structure
checks), but the svg_to_pptx package only needs these two symbols.
"""

from __future__ import annotations

from pathlib import Path


def get_project_info(path: str) -> dict:
    """Return minimal project info dict for a path.

    The real version reads design_spec.md / spec_lock.md / svg files.
    This stub returns enough for svg_to_pptx to determine canvas format.
    """
    p = Path(path)
    return {
        "format": "unknown",
        "name": p.name,
        "path": str(p),
    }


def normalize_canvas_format(fmt: str | None) -> str:
    """Normalize a canvas format string to a known key."""
    from .config import CANVAS_FORMATS
    if fmt and fmt in CANVAS_FORMATS:
        return fmt
    return "ppt169"


def validate_project_structure(project_path: str) -> list[str]:
    """Stub — returns empty list (no validation errors)."""
    return []


def validate_svg_viewbox(svg_path: str, expected_viewbox: str | None = None) -> bool:
    """Stub — always returns True."""
    return True
