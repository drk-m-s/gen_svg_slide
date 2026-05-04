"""
End-to-end test for gen_svg_slide pipeline.

Validates the full workflow:
  1. Parse a PPTX slide          → ParsedSlide
  2. Extract design tokens       → DesignTokens
  3. Build design spec + lock    → design_spec.md + spec_lock.md
  4. Generate SVG via AI         → slide_01.svg
  5. Convert SVG → PPTX          → output.pptx

Usage:
    # Set your API key in .env first, then:
    cd gen_svg_slide
    python -m tests.test_pipeline <path_to_reference.pptx>

    # Or specify backend/model:
    python -m tests.test_pipeline reference.pptx --backend deepseek

    # Dry-run mode (skip AI call, test parser + designer + converter only):
    python -m tests.test_pipeline reference.pptx --dry-run [--skip-convert]
"""

from __future__ import annotations

import sys
import argparse
import tempfile
from pathlib import Path

# Ensure gen_svg_slide package is importable
_PKG_ROOT = Path(__file__).resolve().parent.parent
if str(_PKG_ROOT) not in sys.path:
    sys.path.insert(0, str(_PKG_ROOT))


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _section(title: str) -> None:
    print(f"\n{'='*60}")
    print(f"  {title}")
    print(f"{'='*60}")


def _ok(msg: str) -> None:
    print(f"  ✓ {msg}")


def _fail(msg: str) -> None:
    print(f"  ✗ FAIL: {msg}")


# ---------------------------------------------------------------------------
# Test cases
# ---------------------------------------------------------------------------

def test_parse_pptx(pptx_path: Path) -> bool:
    """Test Step 1-2: PPTX parsing."""
    _section("Test 1: PPTX Parsing")

    from gen_svg_slide.extractor.pptx_parser import parse_pptx_slide

    try:
        parsed = parse_pptx_slide(pptx_path)
    except Exception as e:
        _fail(f"Parse error: {e}")
        return False

    _ok(f"Slide dimensions: {parsed.slide_width_px:.0f}×{parsed.slide_height_px:.0f} px")
    _ok(f"Shapes found: {len(parsed.shapes)}")

    for s in parsed.shapes:
        tag = s.shape_type
        if s.paragraphs:
            text_preview = " ".join(p.plain_text for p in s.paragraphs if p.plain_text.strip())[:60]
            tag += f' — "{text_preview}..."' if len(text_preview) >= 60 else f' — "{text_preview}"'
        _ok(f"  [{s.name}] {tag}")

    # Store for later tests
    test_parse_pptx._parsed = parsed  # type: ignore[attr-defined]
    return True


def test_extract_design() -> bool:
    """Test Step 3: Design token extraction."""
    _section("Test 2: Design Extraction")

    from gen_svg_slide.extractor.design_extractor import extract_design_tokens

    parsed = getattr(test_parse_pptx, "_parsed", None)
    if parsed is None:
        _fail("No parsed slide available — run test_parse_pptx first")
        return False

    try:
        tokens = extract_design_tokens(parsed, reference_path="test_input.pptx")
    except Exception as e:
        _fail(f"Extraction error: {e}")
        return False

    _ok(f"Font family: {tokens.font_family}")
    _ok(f"Font sizes: body={tokens.body:.0f}, title={tokens.title:.0f}")
    _ok(f"Colors: bg=#{tokens.background_color}, text=#{tokens.default_text_color}")
    _ok(f"Palette: {list(tokens.colors.keys())}")
    _ok(f"Margins: L={tokens.margin_left_px:.0f} R={tokens.margin_right_px:.0f} "
         f"T={tokens.margin_top_px:.0f} B={tokens.margin_bottom_px:.0f}")
    _ok(f"ViewBox: {tokens.viewbox}")

    test_extract_design._tokens = tokens  # type: ignore[attr-defined]
    return True


def test_build_specs() -> bool:
    """Test Step 3b: Spec file generation."""
    _section("Test 3: Spec File Generation")

    from gen_svg_slide.designer.spec_builder import build_design_spec, build_spec_lock

    tokens = getattr(test_extract_design, "_tokens", None)
    if tokens is None:
        _fail("No design tokens available")
        return False

    try:
        design_spec = build_design_spec(tokens)
        spec_lock = build_spec_lock(tokens)
    except Exception as e:
        _fail(f"Spec build error: {e}")
        return False

    _ok(f"design_spec: {len(design_spec)} chars")
    _ok(f"spec_lock: {len(spec_lock)} chars")

    # Verify key fields present in spec_lock
    for required in ("## canvas", "## colors", "## typography", "## layout", "## page_rhythm", "## forbidden"):
        if required in spec_lock:
            _ok(f"  spec_lock contains: {required}")
        else:
            _fail(f"  spec_lock MISSING: {required}")

    test_build_specs._design_spec = design_spec  # type: ignore[attr-defined]
    test_build_specs._spec_lock = spec_lock  # type: ignore[attr-defined]
    return True


def test_generate_svg(backend_name: str = "openai", api_key: str | None = None) -> bool:
    """Test Step 4: AI SVG generation."""
    _section("Test 4: AI SVG Generation")

    from gen_svg_slide.generator.svg_renderer import (
        generate_svg_slide,
        OpenAiSvgBackend,
        AnthropicSvgBackend,
        DeepSeekSvgBackend,
        QwenSvgBackend,
    )

    parsed = getattr(test_parse_pptx, "_parsed", None)
    tokens = getattr(test_extract_design, "_tokens", None)
    design_spec = getattr(test_build_specs, "_design_spec", None)
    spec_lock = getattr(test_build_specs, "_spec_lock", None)

    if not all([parsed, tokens, design_spec, spec_lock]):
        _fail("Missing prerequisite data from earlier tests")
        return False

    # Select backend
    backend_map = {
        "openai":    (OpenAiSvgBackend,    {}),
        "anthropic": (AnthropicSvgBackend, {}),
        "deepseek":  (DeepSeekSvgBackend,  {}),
        "qwen":      (QwenSvgBackend,      {}),
    }

    if backend_name not in backend_map:
        _fail(f"Unknown backend: {backend_name}")
        return False

    backend_cls, kwargs = backend_map[backend_name]
    if api_key:
        kwargs["api_key"] = api_key

    try:
        backend = backend_cls(**kwargs)
    except ValueError as e:
        _fail(f"Backend init error: {e}")
        print("  Hint: Set the API key in .env or pass --api-key")
        return False

    _ok(f"Backend: {backend.__class__.__name__} (model: {backend.model})")

    try:
        result = generate_svg_slide(
            parsed=parsed,
            tokens=tokens,
            design_spec=design_spec,
            spec_lock=spec_lock,
            backend=backend,
            user_instruction="Keep the same design and layout as the reference slide.",
        )
    except Exception as e:
        _fail(f"Generation error: {e}")
        import traceback
        traceback.print_exc()
        return False

    if not result.svg_content:
        _fail("Empty SVG returned")
        return False

    _ok(f"SVG length: {len(result.svg_content)} chars")
    _ok(f"Model: {result.model}")
    if result.usage:
        _ok(f"Usage: {result.usage}")

    # Quick validation
    from gen_svg_slide.utils.image_utils import validate_svg_quick
    valid, errors = validate_svg_quick(result.svg_content)
    if valid:
        _ok("SVG validation: PASSED")
    else:
        for err in errors:
            _fail(f"SVG validation: {err}")

    # Check for <svg> tag
    if "<svg" in result.svg_content and "</svg>" in result.svg_content:
        _ok("SVG structure: valid <svg>...</svg>")
    else:
        _fail("SVG structure: missing <svg> tags")

    test_generate_svg._svg_content = result.svg_content  # type: ignore[attr-defined]
    return True


def test_convert_to_pptx(output_dir: Path | None = None) -> bool:
    """Test Step 5: SVG → PPTX conversion."""
    _section("Test 5: SVG → PPTX Conversion")

    from gen_svg_slide.converter import svg_string_to_pptx

    svg_content = getattr(test_generate_svg, "_svg_content", None)
    if svg_content is None:
        _fail("No SVG content available — run test_generate_svg first (or use --dry-run with a known SVG)")
        return False

    if output_dir is None:
        output_dir = Path(tempfile.mkdtemp())

    output_pptx = output_dir / "test_output.pptx"
    _ok(f"Output path: {output_pptx}")

    try:
        success = svg_string_to_pptx(
            svg_content=svg_content,
            output_path=output_pptx,
            canvas_format="ppt169",
            use_native_shapes=True,
            verbose=False,
        )
    except Exception as e:
        _fail(f"Conversion error: {e}")
        import traceback
        traceback.print_exc()
        return False

    if success and output_pptx.exists():
        size_kb = output_pptx.stat().st_size / 1024
        _ok(f"PPTX created: {output_pptx} ({size_kb:.1f} KB)")
        return True
    else:
        _fail("PPTX file not created or conversion reported failure")
        return False


# ---------------------------------------------------------------------------
# Main runner
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(
        description="gen_svg_slide — end-to-end pipeline test",
    )
    parser.add_argument("input", type=str, help="Path to reference PPTX file")
    parser.add_argument("--backend", type=str, default="openai",
                        choices=["openai", "anthropic", "deepseek", "qwen"],
                        help="AI backend to use (default: openai)")
    parser.add_argument("--api-key", type=str, default=None,
                        help="Override API key (otherwise reads from .env)")
    parser.add_argument("--dry-run", action="store_true",
                        help="Skip AI generation (test parser + designer only)")
    parser.add_argument("--skip-convert", action="store_true",
                        help="Skip PPTX conversion (test up to SVG generation)")
    parser.add_argument("-o", "--output-dir", type=str, default=None,
                        help="Output directory for generated files")

    args = parser.parse_args()

    pptx_path = Path(args.input)
    if not pptx_path.exists():
        print(f"Error: File not found: {pptx_path}")
        sys.exit(1)

    output_dir = Path(args.output_dir) if args.output_dir else Path(tempfile.mkdtemp())
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"Input:  {pptx_path}")
    print(f"Output: {output_dir}")
    print(f"Backend: {args.backend}{' (dry-run)' if args.dry_run else ''}")

    results: dict[str, bool] = {}

    # Always run parser + designer tests
    results["parse"] = test_parse_pptx(pptx_path)
    if not results["parse"]:
        print("\n✗ Parser failed — aborting.")
        sys.exit(1)

    results["design"] = test_extract_design()
    results["specs"] = test_build_specs()

    if args.dry_run:
        print("\n--- Dry-run complete (skipping AI call) ---")
        print(f"Output directory: {output_dir}")
        # Still save specs for inspection
        from gen_svg_slide.designer.spec_builder import write_specs
        tokens = getattr(test_extract_design, "_tokens", None)
        if tokens:
            write_specs(tokens, output_dir)
            print(f"Spec files written to: {output_dir}")
    else:
        results["generate"] = test_generate_svg(args.backend, args.api_key)
        if not results.get("generate"):
            print("\n✗ SVG generation failed — check API key / network.")
            sys.exit(1)

        if not args.skip_convert:
            results["convert"] = test_convert_to_pptx(output_dir)

    # Summary
    _section("Results")
    passed = sum(1 for v in results.values() if v)
    total = len(results)
    for name, ok in results.items():
        print(f"  {'✓' if ok else '✗'} {name}")
    print(f"\n  {passed}/{total} tests passed")

    if passed == total:
        print("\n✓ All tests passed!")
    else:
        print(f"\n✗ {total - passed} test(s) failed")
        sys.exit(1)


if __name__ == "__main__":
    main()
