"""
Converter package — self-contained SVG → PPTX pipeline.

Re-exports the public API from wrapper.py so callers can do::

    from gen_svg_slide.converter import svg_string_to_pptx
"""

from .wrapper import (
    svg_file_to_pptx,
    svg_string_to_pptx,
    convert_svg_to_drawingml,
)

__all__ = [
    "svg_file_to_pptx",
    "svg_string_to_pptx",
    "convert_svg_to_drawingml",
]
