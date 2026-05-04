"""
Design Extractor — derives design tokens from a ``ParsedSlide``.

Analyzes font usage, colors, spacing, and layout to produce a structured
``DesignTokens`` object. This is the "AI parses the slide design" step (Step 3).

Strategy:
  - Inspect every shape's text runs → extract font families, sizes, weights.
  - Inspect fill / stroke colors → build a color palette.
  - Measure text positions relative to slide edges → infer margins.
  - Detect the title position and content area boundaries.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from collections import Counter

from .pptx_parser import ParsedSlide, ShapeInfo, ParagraphInfo, TextRun, emu_to_px


# ---------------------------------------------------------------------------
# Output data class
# ---------------------------------------------------------------------------

@dataclass
class DesignTokens:
    """Machine-readable design tokens extracted from a reference slide.

    Mirrors the format of ppt-master's ``spec_lock.md`` so downstream
    modules (generator, designer) can consume both without translation.
    """

    # Canvas
    viewbox: str = "0 0 1280 720"  # e.g. "0 0 1280 720"
    width_px: float = 1280
    height_px: float = 720

    # Colors — hex strings WITHOUT # (matching spec_lock format)
    colors: dict[str, str] = field(default_factory=dict)
    # Keys: bg, primary, secondary_bg, accent, text, text_secondary,
    #       text_tertiary, border, success, warning, etc.

    # Typography
    font_family: str = "Arial"
    title_family: str | None = None
    body_family: str | None = None
    code_family: str = "Consolas"

    # Font size ramp (in px, matching SVG viewBox convention)
    body: float = 18.0
    title: float = 32.0
    subtitle: float = 24.0
    annotation: float = 14.0
    cover_title: float = 60.0

    # Layout
    margin_left_px: float = 60.0
    margin_right_px: float = 60.0
    margin_top_px: float = 60.0
    margin_bottom_px: float = 60.0
    title_y_px: float = 60.0
    title_height_px: float = 50.0
    content_y_px: float = 140.0  # where body content starts

    # Deck-level preferences
    background_color: str = "FFFFFF"
    default_text_color: str = "333333"

    # Extra metadata
    font_size_ramp: dict[str, float] = field(default_factory=dict)
    font_weight_distribution: dict[str, int] = field(default_factory=dict)
    shape_type_distribution: dict[str, int] = field(default_factory=dict)

    # Original reference (for debugging / traceability)
    reference_pptx: str = ""


# ---------------------------------------------------------------------------
# Extraction logic
# ---------------------------------------------------------------------------

def extract_design_tokens(parsed: ParsedSlide, reference_path: str = "") -> DesignTokens:
    """Extract design tokens from a parsed slide.

    Args:
        parsed: A ``ParsedSlide`` from ``pptx_parser.parse_pptx_slide()``.
        reference_path: Path to the original PPTX file (for traceability).

    Returns:
        ``DesignTokens`` ready for ``spec_builder.build_spec()``.
    """
    tokens = DesignTokens(
        viewbox=f"0 0 {int(parsed.slide_width_px)} {int(parsed.slide_height_px)}",
        width_px=parsed.slide_width_px,
        height_px=parsed.slide_height_px,
        reference_pptx=reference_path,
    )

    _extract_colors(parsed, tokens)
    _extract_typography(parsed, tokens)
    _extract_layout(parsed, tokens)
    _extract_metadata(parsed, tokens)

    return tokens


def _extract_colors(parsed: ParsedSlide, tokens: DesignTokens) -> None:
    """Build a color palette from all shapes on the slide."""
    hex_counter: Counter[str] = Counter()
    text_colors: Counter[str] = Counter()
    bg_candidates: list[str] = []

    for shape in parsed.shapes:
        # Shape fills
        if shape.fill_color_hex and shape.fill_color_hex != "FFFFFF":
            # Large filled shapes may be backgrounds
            area = shape.width_emu * shape.height_emu
            slide_area = parsed.slide_width_emu * parsed.slide_height_emu
            if area / slide_area > 0.4:  # covers >40% of slide
                bg_candidates.append(shape.fill_color_hex)
            hex_counter[shape.fill_color_hex] += 1

        # Text colors
        for para in shape.paragraphs:
            for run in para.runs:
                if run.color_hex and run.color_hex != "000000":
                    text_colors[run.color_hex] += 1

    # Assign colors
    tokens.background_color = bg_candidates[0] if bg_candidates else "FFFFFF"

    # Most common non-background color → primary
    sorted_colors = hex_counter.most_common()
    tokens.colors["bg"] = tokens.background_color
    tokens.colors["secondary_bg"] = _lighten_hex(tokens.background_color) if tokens.background_color != "FFFFFF" else "F5F5F5"

    if sorted_colors:
        tokens.colors["primary"] = sorted_colors[0][0]
    if len(sorted_colors) > 1:
        tokens.colors["accent"] = sorted_colors[1][0]

    # Text colors
    sorted_text = text_colors.most_common()
    tokens.default_text_color = sorted_text[0][0] if sorted_text else "333333"
    tokens.colors["text"] = tokens.default_text_color

    if len(sorted_text) > 1:
        tokens.colors["text_secondary"] = sorted_text[1][0]
    else:
        tokens.colors["text_secondary"] = "888888"

    tokens.colors.setdefault("border", "DDDDDD")
    tokens.colors.setdefault("success", "4ADE80")
    tokens.colors.setdefault("warning", "F87171")


def _extract_typography(parsed: ParsedSlide, tokens: DesignTokens) -> None:
    """Extract font families, sizes, and weights from all text runs."""
    font_counter: Counter[str] = Counter()
    size_counter: Counter[float] = Counter()
    weight_counter: Counter[str] = Counter()

    for shape in parsed.shapes:
        for para in shape.paragraphs:
            for run in para.runs:
                if run.font_name:
                    font_counter[run.font_name] += 1
                if run.font_size_pt:
                    size_counter[run.font_size_pt] += 1
                weight = "bold" if run.bold else "regular"
                weight_counter[weight] += 1

    # Font family
    if font_counter:
        tokens.font_family = font_counter.most_common(1)[0][0]

    # Title font might differ — check the title shape
    title_shape = parsed.title_shape
    if title_shape and title_shape.paragraphs and title_shape.paragraphs[0].runs:
        title_run = title_shape.paragraphs[0].runs[0]
        if title_run.font_name and title_run.font_name != tokens.font_family:
            tokens.title_family = title_run.font_name
        if title_run.font_size_pt:
            tokens.title = title_run.font_size_pt

    # Body size — most common font size that isn't the title
    if size_counter:
        sizes = size_counter.most_common()
        tokens.body = sizes[0][0]
        if len(sizes) > 1:
            tokens.subtitle = sizes[1][0] if sizes[1][0] != tokens.body else sizes[1][0]

    tokens.font_size_ramp = dict(size_counter.most_common())
    tokens.font_weight_distribution = dict(weight_counter.most_common())


def _extract_layout(parsed: ParsedSlide, tokens: DesignTokens) -> None:
    """Infer margins and title position from shape positions."""
    text_shapes = sorted(parsed.text_shapes, key=lambda s: s.top_emu)

    if not text_shapes:
        return

    # Title = topmost text shape
    title_shape = text_shapes[0]
    tokens.title_y_px = emu_to_px(title_shape.top_emu)
    tokens.title_height_px = emu_to_px(title_shape.height_emu)

    # Left margin = smallest left edge among text shapes
    left_edges = [emu_to_px(s.left_emu) for s in text_shapes]
    tokens.margin_left_px = min(left_edges)

    # Right margin = slide_width - max(left+width) among text shapes
    right_edges = [emu_to_px(s.left_emu + s.width_emu) for s in text_shapes]
    tokens.margin_right_px = tokens.width_px - max(right_edges)

    # Top margin
    tokens.margin_top_px = tokens.title_y_px

    # Bottom margin
    bottom_edges = [emu_to_px(s.top_emu + s.height_emu) for s in text_shapes]
    raw_bottom = tokens.height_px - max(bottom_edges)
    tokens.margin_bottom_px = max(raw_bottom, 20.0)  # at least 20px

    # Content area starts after the title
    if len(text_shapes) > 1:
        second_shape = text_shapes[1]
        tokens.content_y_px = emu_to_px(second_shape.top_emu)
    else:
        tokens.content_y_px = tokens.title_y_px + tokens.title_height_px + 20


def _extract_metadata(parsed: ParsedSlide, tokens: DesignTokens) -> None:
    """Count shape types for reporting."""
    type_counter: Counter[str] = Counter()
    for shape in parsed.shapes:
        type_counter[shape.shape_type] += 1
    tokens.shape_type_distribution = dict(type_counter.most_common())


# ---------------------------------------------------------------------------
# Color utilities (matching ppt-master's drawingml_styles pattern)
# ---------------------------------------------------------------------------

def _lighten_hex(hex_color: str, factor: float = 0.15) -> str:
    """Lighten a hex color (RRGGBB, no #) by blending with white."""
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    r = int(r + (255 - r) * factor)
    g = int(g + (255 - g) * factor)
    b = int(b + (255 - b) * factor)
    return f"{r:02X}{g:02X}{b:02X}"
