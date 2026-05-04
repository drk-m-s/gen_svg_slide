"""
Design Spec Builder — produces design_spec.md and spec_lock.md from DesignTokens.

This is the bridge between "parsed slide design" and "AI SVG generator".
Outputs follow the ppt-master conventions so the generator can consume them
and the converter can validate against them.

Corresponds to **Step 3**: "AI formulates a design spec to be followed."
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from ..extractor.design_extractor import DesignTokens


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def build_design_spec(tokens: DesignTokens) -> str:
    """Build a human-readable ``design_spec.md`` from extracted tokens.

    Format follows ppt-master's ``templates/design_spec_reference.md``
    structure (sections I–XI), adapted for single-slide reference extraction.
    """
    return f"""# Design Specification (extracted from reference slide)

## I. Canvas Format
- **Format**: Derived from reference slide ({tokens.width_px:.0f}×{tokens.height_px:.0f} px)
- **viewBox**: `{tokens.viewbox}`

## II. Page Count
- **Range**: 1 page (single slide regeneration)
- **Structure**: Cover/content/ending as single slide

## III. Target Audience
- **Audience**: Matches the reference slide's intended audience
- **Tone**: Professional, consistent with reference design

## IV. Style Objective
- **Style**: Extracted from reference slide — maintain exact visual consistency
- **Layout rhythm**: Single page, following the reference layout pattern
- **Key principle**: Reproduce the design language of the uploaded slide

## V. Color Scheme
- **Background**: #{tokens.background_color}
- **Primary**: #{tokens.colors.get('primary', 'N/A')}
- **Accent**: #{tokens.colors.get('accent', 'N/A')}
- **Text**: #{tokens.colors.get('text', tokens.default_text_color)}
- **Text Secondary**: #{tokens.colors.get('text_secondary', '888888')}
- **Border**: #{tokens.colors.get('border', 'DDDDDD')}

## VI. Icon Usage
- **Approach**: Match the reference slide's icon style
- **Library**: To be determined from content needs

## VII. Typography
- **Title Font**: {tokens.title_family or tokens.font_family}
- **Body Font**: {tokens.font_family}
- **Code Font**: {tokens.code_family}
- **Body Size**: {tokens.body}pt
- **Title Size**: {tokens.title}pt
- **Subtitle Size**: {tokens.subtitle}pt
- **Annotation Size**: {tokens.annotation}pt

## VIII. Image Usage
- **Approach**: Match the reference slide's image placement pattern
- **Sources**: Extracted from reference slide, user-provided, or AI-generated

## IX. Page Outline
### Page 1 — Regenerated Slide
- **Content**: Regenerated version of the reference slide content
- **Layout Strategy**: Reproduce the spatial layout of the reference slide:
  - Title at y={tokens.title_y_px:.0f}px
  - Content area starts at y={tokens.content_y_px:.0f}px
  - Margins: L={tokens.margin_left_px:.0f} R={tokens.margin_right_px:.0f} T={tokens.margin_top_px:.0f} B={tokens.margin_bottom_px:.0f}
- **Rhythm tag**: anchor (single structural page)

## X. Notes
- Reference slide: `{tokens.reference_pptx}`

## XI. Constraints
- Follow all SVG constraints from ppt-master shared-standards.md
- No mask, no style/class, no foreignObject, no rgba()
- Use raw Unicode for typographic characters
"""


def build_spec_lock(tokens: DesignTokens) -> str:
    """Build a machine-readable ``spec_lock.md`` from extracted tokens.

    Format matches ppt-master's ``templates/spec_lock_reference.md`` exactly,
    so the generator and quality checker can consume it without modification.
    """
    lines = [
        "# Execution Lock",
        "",
        "## canvas",
        f"- viewBox: {tokens.viewbox}",
        f"- format: Derived {tokens.width_px:.0f}×{tokens.height_px:.0f}",
        "",
        "## colors",
    ]
    for key in ("bg", "secondary_bg", "primary", "accent", "text",
                 "text_secondary", "text_tertiary", "border", "success", "warning"):
        if key in tokens.colors:
            lines.append(f"- {key}: #{tokens.colors[key]}")

    lines += [
        "",
        "## typography",
        f"- font_family: \"{tokens.font_family}\", sans-serif",
    ]
    if tokens.title_family:
        lines.append(f"- title_family: \"{tokens.title_family}\"")
    if tokens.body_family:
        lines.append(f"- body_family: \"{tokens.body_family}\"")
    lines += [
        f"- code_family: \"{tokens.code_family}\", monospace",
        f"- body: {tokens.body:.0f}",
        f"- title: {tokens.title:.0f}",
        f"- subtitle: {tokens.subtitle:.0f}",
        f"- annotation: {tokens.annotation:.0f}",
        "",
        "## layout",
        f"- margin_left: {tokens.margin_left_px:.0f}",
        f"- margin_right: {tokens.margin_right_px:.0f}",
        f"- margin_top: {tokens.margin_top_px:.0f}",
        f"- margin_bottom: {tokens.margin_bottom_px:.0f}",
        f"- title_y: {tokens.title_y_px:.0f}",
        f"- title_height: {tokens.title_height_px:.0f}",
        f"- content_y: {tokens.content_y_px:.0f}",
        "",
        "## page_rhythm",
        "- P01: anchor",
        "",
        "## forbidden",
        "- Mixing icon libraries",
        "- rgba()",
        "- `<style>`, `class`, `<foreignObject>`, `textPath`, `@font-face`",
        "- `<animate*>`, `<script>`, `<iframe>`",
        "- `<g opacity>` (set opacity on each child element individually)",
    ]

    return "\n".join(lines)


def write_specs(tokens: DesignTokens, output_dir: Path) -> tuple[Path, Path]:
    """Write both design_spec.md and spec_lock.md to an output directory.

    Args:
        tokens: Extracted design tokens.
        output_dir: Directory to write the spec files into.

    Returns:
        (design_spec_path, spec_lock_path) tuple.
    """
    output_dir.mkdir(parents=True, exist_ok=True)

    design_spec_path = output_dir / "design_spec.md"
    design_spec_path.write_text(build_design_spec(tokens), encoding="utf-8")

    spec_lock_path = output_dir / "spec_lock.md"
    spec_lock_path.write_text(build_spec_lock(tokens), encoding="utf-8")

    return design_spec_path, spec_lock_path
