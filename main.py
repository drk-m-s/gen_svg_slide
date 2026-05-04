"""
main.py — the central pipeline orchestrator for gen_svg_slide.

Ties together all five steps of the workflow:
  1. Parse uploaded PPTX slide  → ParsedSlide
  2. Extract design tokens      → DesignTokens
  3. Build design spec + lock   → design_spec.md + spec_lock.md
  4. AI generates SVG           → slide_01.svg
  5. Convert SVG → PPTX         → output.pptx

Usage:
    # Programmatic API
    from main import SlideRegenerator
    regen = SlideRegenerator(backend=AnthropicSvgBackend(api_key="..."))
    regen.run("input.pptx", "output.pptx")

    # CLI
    python -m gen_svg_slide.main input.pptx -o output.pptx
"""

from __future__ import annotations

import sys
import argparse
from pathlib import Path

# Ensure the gen_svg_slide package root is on sys.path
_PKG_ROOT = Path(__file__).resolve().parent.parent
if str(_PKG_ROOT) not in sys.path:
    sys.path.insert(0, str(_PKG_ROOT))

from gen_svg_slide.config import DEFAULT_CANVAS
from gen_svg_slide.extractor.pptx_parser import parse_pptx_slide, ParsedSlide
from gen_svg_slide.extractor.design_extractor import extract_design_tokens, DesignTokens
from gen_svg_slide.designer.spec_builder import build_design_spec, build_spec_lock, write_specs
from gen_svg_slide.generator.svg_renderer import (
    generate_svg_slide,
    SvgGeneratorBackend,
    SvgGenerationResult,
    OpenAiSvgBackend,
    AnthropicSvgBackend,
    DeepSeekSvgBackend,
    QwenSvgBackend,
)
from gen_svg_slide.converter import svg_string_to_pptx
from gen_svg_slide.utils.image_utils import (
    extract_images_from_parsed_slide,
    validate_svg_quick,
    embed_images_base64,
    make_run_dir,
)


# ---------------------------------------------------------------------------
# Pipeline result
# ---------------------------------------------------------------------------

from dataclasses import dataclass, field


@dataclass
class SlideRegenerationResult:
    """Complete result of a slide regeneration run."""
    success: bool
    run_dir: Path
    output_pptx: Path | None = None
    svg_content: str = ""
    parsed_slide: ParsedSlide | None = None
    design_tokens: DesignTokens | None = None
    design_spec: str = ""
    spec_lock: str = ""
    errors: list[str] = field(default_factory=list)
    svg_validation_errors: list[str] = field(default_factory=list)
    generation_usage: dict = field(default_factory=dict)


# ---------------------------------------------------------------------------
# Orchestrator class
# ---------------------------------------------------------------------------

class SlideRegenerator:
    """End-to-end pipeline: upload PPTX → extract design → generate SVG → export PPTX.

    Usage::

        backend = AnthropicSvgBackend(api_key="sk-...")
        regen = SlideRegenerator(backend=backend)
        result = regen.run("reference.pptx", "output.pptx")
        print(f"Output: {result.output_pptx}")
    """

    def __init__(
        self,
        backend: SvgGeneratorBackend,
        canvas_format: str = DEFAULT_CANVAS,
        use_native_shapes: bool = True,
        verbose: bool = True,
    ):
        self.backend = backend
        self.canvas_format = canvas_format
        self.use_native_shapes = use_native_shapes
        self.verbose = verbose

    # ------------------------------------------------------------------
    # Step 1-2: Parse
    # ------------------------------------------------------------------

    def parse(self, pptx_path: str | Path, slide_index: int = 0) -> ParsedSlide:
        """Parse the uploaded PPTX slide (Steps 1-2)."""
        if self.verbose:
            print(f"[1/5] Parsing PPTX: {pptx_path}")
        return parse_pptx_slide(pptx_path, slide_index=slide_index)

    # ------------------------------------------------------------------
    # Step 3: Extract design
    # ------------------------------------------------------------------

    def extract_design(self, parsed: ParsedSlide, reference_path: str = "") -> tuple[DesignTokens, str, str]:
        """Extract design tokens and build spec files (Step 3)."""
        if self.verbose:
            print("[2/5] Extracting design tokens...")

        tokens = extract_design_tokens(parsed, reference_path=reference_path)

        if self.verbose:
            print(f"       Font: {tokens.font_family} ({tokens.body:.0f}pt body)")
            print(f"       Colors: bg=#{tokens.background_color}, text=#{tokens.default_text_color}")
            print(f"       Margins: L={tokens.margin_left_px:.0f} R={tokens.margin_right_px:.0f}")

        design_spec = build_design_spec(tokens)
        spec_lock = build_spec_lock(tokens)

        return tokens, design_spec, spec_lock

    # ------------------------------------------------------------------
    # Step 4: Generate SVG
    # ------------------------------------------------------------------

    def generate_svg(
        self,
        parsed: ParsedSlide,
        tokens: DesignTokens,
        design_spec: str,
        spec_lock: str,
        user_instruction: str = "",
    ) -> SvgGenerationResult:
        """AI generates SVG from the design spec (Step 4)."""
        if self.verbose:
            print("[3/5] Generating SVG via AI...")

        result = generate_svg_slide(
            parsed=parsed,
            tokens=tokens,
            design_spec=design_spec,
            spec_lock=spec_lock,
            backend=self.backend,
            user_instruction=user_instruction,
        )

        if self.verbose and result.usage:
            print(f"       Model: {result.model}, tokens: {result.usage}")

        return result

    # ------------------------------------------------------------------
    # Step 5: Convert to PPTX
    # ------------------------------------------------------------------

    def convert_to_pptx(
        self,
        svg_content: str,
        output_path: Path,
        work_dir: Path,
        images_dir: Path | None = None,
    ) -> bool:
        """Convert SVG to native-shapes PPTX (Step 5)."""
        if self.verbose:
            print("[4/5] Converting SVG → PPTX...")

        # Embed images as Base64 so the SVG is self-contained before conversion
        if images_dir and images_dir.exists():
            svg_content = embed_images_base64(svg_content, images_dir)

        success = svg_string_to_pptx(
            svg_content=svg_content,
            output_path=output_path,
            work_dir=work_dir,
            canvas_format=self.canvas_format,
            use_native_shapes=self.use_native_shapes,
            verbose=self.verbose,
        )

        if self.verbose:
            print(f"       {'✓' if success else '✗'} Output: {output_path}")

        return success

    # ------------------------------------------------------------------
    # Full pipeline
    # ------------------------------------------------------------------

    def run(
        self,
        input_pptx: str | Path,
        output_pptx: str | Path | None = None,
        user_instruction: str = "",
        run_id: str | None = None,
    ) -> SlideRegenerationResult:
        """Execute the full pipeline end-to-end.

        Args:
            input_pptx: Path to the uploaded reference PPTX slide.
            output_pptx: Desired output path. Auto-generated if None.
            user_instruction: Optional content changes from the user.
            run_id: Optional ID for the run directory.

        Returns:
            SlideRegenerationResult with all outputs and diagnostics.
        """
        errors: list[str] = []
        input_pptx = Path(input_pptx)

        # Create working directory
        run_dir = make_run_dir(run_id)

        try:
            # --- Step 1-2: Parse ---
            parsed = self.parse(input_pptx)
            if not parsed.shapes:
                return SlideRegenerationResult(
                    success=False, run_dir=run_dir,
                    errors=["No shapes found in the uploaded slide"],
                )

            # Save extracted images
            images_dir = run_dir / "images"
            extract_images_from_parsed_slide(parsed, images_dir)

            # --- Step 3: Extract design ---
            tokens, design_spec, spec_lock = self.extract_design(
                parsed, reference_path=str(input_pptx),
            )

            # Write specs to run dir
            write_specs(tokens, run_dir)

            # --- Step 4: Generate SVG ---
            gen_result = self.generate_svg(
                parsed, tokens, design_spec, spec_lock, user_instruction,
            )

            if not gen_result.svg_content:
                return SlideRegenerationResult(
                    success=False, run_dir=run_dir,
                    errors=["AI generated empty SVG content"],
                    parsed_slide=parsed,
                    design_tokens=tokens,
                    design_spec=design_spec,
                    spec_lock=spec_lock,
                )

            # Validate SVG before saving
            svg_valid, svg_errors = validate_svg_quick(gen_result.svg_content)

            # Save SVG to run dir
            svg_path = run_dir / "svg_output" / "slide_01.svg"
            svg_path.parent.mkdir(exist_ok=True)
            svg_path.write_text(gen_result.svg_content, encoding="utf-8")

            # --- Step 5: Convert to PPTX ---
            if output_pptx is None:
                output_path = run_dir / "output.pptx"
            else:
                output_path = Path(output_pptx)

            convert_success = self.convert_to_pptx(
                svg_content=gen_result.svg_content,
                output_path=output_path,
                work_dir=run_dir,
                images_dir=images_dir,
            )

            if self.verbose:
                print("[5/5] Done!")

            return SlideRegenerationResult(
                success=convert_success,
                run_dir=run_dir,
                output_pptx=output_path if convert_success else None,
                svg_content=gen_result.svg_content,
                parsed_slide=parsed,
                design_tokens=tokens,
                design_spec=design_spec,
                spec_lock=spec_lock,
                errors=errors + ([] if convert_success else ["PPTX conversion failed"]),
                svg_validation_errors=svg_errors,
                generation_usage=gen_result.usage,
            )

        except Exception as e:
            errors.append(f"Pipeline error: {e}")
            if self.verbose:
                import traceback
                traceback.print_exc()
            return SlideRegenerationResult(
                success=False, run_dir=run_dir, errors=errors,
            )


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(
        description="gen_svg_slide — regenerate a PPTX slide from a reference",
    )
    parser.add_argument("input", type=str, help="Path to reference PPTX file")
    parser.add_argument("-o", "--output", type=str, default=None, help="Output PPTX path")
    parser.add_argument("--api-key", type=str, default=None,
                        help="API key (or set OPENAI_API_KEY / ANTHROPIC_API_KEY env var)")
    parser.add_argument("--backend", type=str, default="openai",
                        choices=["openai", "anthropic", "deepseek", "qwen"], help="AI backend")
    parser.add_argument("--model", type=str, default=None,
                        help="Model name (default: gpt-4o or claude-sonnet-4-20250514)")
    parser.add_argument("--instruction", type=str, default="",
                        help="User instruction for content changes")
    parser.add_argument("--compat", action="store_true",
                        help="Use compatibility mode (SVG as image, not native shapes)")
    parser.add_argument("-q", "--quiet", action="store_true", help="Quiet mode")

    args = parser.parse_args()

    # Backend selection — API keys are loaded from .env or environment.
    # The --api-key flag overrides .env for convenience.
    backend_map = {
        "openai":    (OpenAiSvgBackend,    {}),
        "anthropic": (AnthropicSvgBackend, {}),
        "deepseek":  (DeepSeekSvgBackend,  {}),
        "qwen":      (QwenSvgBackend,      {}),
    }

    backend_cls, backend_kwargs = backend_map[args.backend]
    if args.api_key:
        backend_kwargs["api_key"] = args.api_key
    if args.model:
        backend_kwargs["model"] = args.model

    try:
        backend = backend_cls(**backend_kwargs)
    except ValueError as e:
        print(f"Error: {e}")
        print("Set the appropriate API key in .env or via --api-key.")
        sys.exit(1)

    # Run pipeline
    regen = SlideRegenerator(
        backend=backend,
        use_native_shapes=not args.compat,
        verbose=not args.quiet,
    )

    result = regen.run(
        input_pptx=args.input,
        output_pptx=args.output,
        user_instruction=args.instruction,
    )

    if result.success:
        print(f"\n✓ Slide regenerated: {result.output_pptx}")
        if not args.quiet:
            print(f"  Run directory: {result.run_dir}")
    else:
        print(f"\n✗ Failed: {'; '.join(result.errors)}")
        sys.exit(1)


if __name__ == "__main__":
    main()
