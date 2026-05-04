"""
Global configuration for gen_svg_slide.

Centralizes paths, defaults, and constants so the rest of the codebase
doesn't scatter magic values.
"""

from __future__ import annotations

from pathlib import Path

# ---------------------------------------------------------------------------
# Path to the ppt-master skill package (for reusing svg_to_pptx / finalize_svg)
# ---------------------------------------------------------------------------
_REPO_ROOT = Path(__file__).resolve().parent.parent
PPT_MASTER_SKILL_DIR = _REPO_ROOT / "skills" / "ppt-master"
PPT_MASTER_SCRIPTS = PPT_MASTER_SKILL_DIR / "scripts"

# ---------------------------------------------------------------------------
# Canvas / slide defaults
# ---------------------------------------------------------------------------
DEFAULT_CANVAS = "ppt169"  # 16:9 widescreen
CANVAS_VIEWBOXES = {
    "ppt169": "0 0 1280 720",
    "ppt43": "0 0 1024 768",
    "square": "0 0 1080 1080",
    "story": "0 0 1080 1920",
}
SLIDE_WIDTH_PX = 1280
SLIDE_HEIGHT_PX = 720

# ---------------------------------------------------------------------------
# Output structure (per run)
# ---------------------------------------------------------------------------
# Each generation run creates a timestamped working directory:
#   output/<run_id>/
#     input.pptx          # uploaded reference slide
#     design_spec.md      # extracted design spec
#     spec_lock.md        # machine-readable design contract
#     svg_output/         # AI-generated SVG
#     svg_final/          # post-processed SVG
#     images/             # extracted images
#     output.pptx         # final generated PPTX
OUTPUT_ROOT = _REPO_ROOT / "gen_svg_slide" / "output"

# ---------------------------------------------------------------------------
# SVG constraints (mirrors ppt-master shared-standards.md)
# ---------------------------------------------------------------------------
BANNED_SVG_FEATURES = {
    "mask", "foreignObject", "textPath", "animate", "set",
    "script", "iframe",
}
BANNED_SVG_ATTRS = {"style": "embedded stylesheets", "class": "CSS classes"}
XML_ILLEGAL_CHARS = {"&": "&amp;", "<": "&lt;", ">": "&gt;"}

# ---------------------------------------------------------------------------
# LLM / AI generation defaults
# ---------------------------------------------------------------------------
DEFAULT_MODEL = "gpt-4o"  # or claude-sonnet-4-20250514, etc.
SVG_GENERATION_SYSTEM_PROMPT = "You are an expert SVG designer. Generate clean, valid SVG that follows a provided design specification precisely."
