# Full test (requires .env with API key):
python -m gen_svg_slide.tests.test_pipeline reference.pptx --backend deepseek

# Dry-run (parser + designer only, no AI call):
python -m gen_svg_slide.tests.test_pipeline reference.pptx --dry-run

# Skip conversion (test up to SVG generation):
python -m gen_svg_slide.tests.test_pipeline reference.pptx --backend qwen --skip-convert