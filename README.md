# gen_svg_slide — AI-Powered PPTX Slide Regeneration

> **Status**: Self-contained — zero runtime dependency on the parent ppt-master repo
>
> Takes an uploaded PPTX slide → extracts its content & design → AI regenerates as SVG → converts to native PowerPoint shapes.

---

## How It Maps to ppt-master

| Step | gen_svg_slide | ppt-master equivalent | Reuse? |
|------|--------------|----------------------|--------|
| **1. Parse PPTX** | `extractor/pptx_parser.py` | `source_to_md/ppt_to_md.py` | New (reverse direction — PPTX→structured data) |
| **2. Extract design** | `extractor/design_extractor.py` | Strategist phase (LLM) | New (rules-based extraction, not LLM) |
| **3. Build spec** | `designer/spec_builder.py` | `spec_lock.md` + `design_spec.md` | Same format, new builder |
| **4. Generate SVG** | `generator/svg_renderer.py` | Executor phase (LLM) | Same approach — LLM constrained by spec_lock |
| **5. SVG → PPTX** | `converter/svg_to_pptx/` (bundled) | `svg_to_pptx/` package | **Exact copy** — self-contained, no import needed |

---

## Architecture

```
User uploads reference.pptx
        │
        ▼
┌─────────────────────────────────────────────────┐
│  main.py — SlideRegenerator pipeline            │
│                                                  │
│  [1] ParsedSlide  ←── extractor/pptx_parser.py  │
│       │                                          │
│  [2] DesignTokens ←── extractor/design_extractor │
│       │                                          │
│  [3] spec_lock.md  ←── designer/spec_builder.py  │
│       design_spec.md                             │
│       │                                          │
│  [4] slide_01.svg  ←── generator/svg_renderer.py│
│       │            ←── AI backend (OpenAI/Claude)│
│       │                                          │
│  [5] output.pptx   ←── converter/svg_to_pptx/   │
│       │            ←── bundled (self-contained)  │
└─────────────────────────────────────────────────┘
        │
        ▼
    output.pptx  (native PowerPoint shapes, editable)
```

---

## Quick Start

### 1. Install dependencies

```bash
cd gen_svg_slide
pip install -r requirements.txt
```

### 2. Set your API keys in `.env`

```bash
# Copy the example and edit with your keys
cp .env.example .env
# Then edit .env — fill in at least one provider's key
```

### 3. Run the pipeline

```bash
# CLI — uses the backend and key from .env
python -m gen_svg_slide.main reference.pptx -o output.pptx

# Choose a specific backend
python -m gen_svg_slide.main reference.pptx --backend deepseek
python -m gen_svg_slide.main reference.pptx --backend qwen

# With custom instruction
python -m gen_svg_slide.main reference.pptx -o output.pptx \
    --instruction "Change the title to 'Q4 Results' and add a blue accent"

### 4. Programmatic API

```python
from gen_svg_slide.main import SlideRegenerator
from gen_svg_slide.generator.svg_renderer import DeepSeekSvgBackend

# API keys are auto-loaded from .env — no need to pass them explicitly
backend = DeepSeekSvgBackend()  # reads DEEPSEEK_API_KEY from .env
regen = SlideRegenerator(backend=backend)

result = regen.run(
    input_pptx="reference.pptx",
    output_pptx="output.pptx",
    user_instruction="Keep the same design, update content to Q4 numbers",
)

print(f"Success: {result.success}")
print(f"Output: {result.output_pptx}")
```

---

## Directory Structure

```
gen_svg_slide/
├── main.py                      # Pipeline orchestrator + CLI
├── config.py                    # Global constants, paths, defaults
├── requirements.txt             # Python dependencies
├── README.md                    # ← you are here
│
├── extractor/                   # Step 1-2: Parse PPTX → extract content + design
│   ├── __init__.py
│   ├── pptx_parser.py           #   Parse PPTX → ParsedSlide (shapes, text, images, tables)
│   └── design_extractor.py      #   ParsedSlide → DesignTokens (fonts, colors, layout)
│
├── designer/                    # Step 3: Build design spec
│   ├── __init__.py
│   └── spec_builder.py          #   DesignTokens → design_spec.md + spec_lock.md
│
├── generator/                   # Step 4: AI generates SVG
│   ├── __init__.py
│   └── svg_renderer.py          #   Prompt builder, OpenAI/Anthropic backends, orchestrator
│
├── converter/                   # Step 5: SVG → PPTX (self-contained engine)
│   ├── __init__.py              #   Re-exports from wrapper.py
│   ├── wrapper.py               #   Public API: svg_string_to_pptx, etc.
│   ├── config.py                #   Canvas format constants
│   ├── project_utils.py         #   Minimal stubs for the engine
│   ├── svg_to_pptx/             #   Bundled ppt-master SVG→DrawingML→PPTX engine (17 files)
│   └── svg_finalize/            #   Bundled ppt-master SVG post-processing (8 files)
│
├── utils/                       # Shared utilities
│   ├── __init__.py
│   └── image_utils.py           #   Image extraction, SVG validation, Base64 embedding
│
├── tests/                       # Test files (to be added)
│
└── output/                      # Generated output (per-run directories)
    └── <run_id>/
        ├── images/              #   Extracted images from reference slide
        ├── design_spec.md       #   Human-readable design narrative
        ├── spec_lock.md         #   Machine-readable execution contract
        ├── svg_output/          #   AI-generated SVG
        └── output.pptx          #   Final PPTX
```

---

## Key Design Decisions

### 1. SVG as the intermediate format

Same as ppt-master. SVG shares the same conceptual model as DrawingML (absolute-coordinate 2D vector graphics). The conversion is translation between dialects — not format mismatch.

### 2. Self-contained — the svg_to_pptx engine is bundled

The entire `svg_to_pptx/` and `svg_finalize/` packages from ppt-master live under `converter/`. No runtime dependency on the parent repo. The `converter/config.py` and `converter/project_utils.py` provide the minimal stubs needed by the engine.

### 3. Design extraction is rules-based, not LLM-based

The design extractor (`design_extractor.py`) uses heuristics (counters, position math) to derive fonts/colors/margins. This is deterministic, fast, and doesn't burn AI tokens. The LLM is reserved for the creative SVG generation step where it's needed.

### 4. spec_lock.md is the anti-drift contract

Following ppt-master's pattern, `design_spec.md` is human-readable narrative, and `spec_lock.md` is the machine-readable execution contract. The AI generator receives both — spec_lock ensures exact color/font values don't drift.

### 5. Two SVG output modes (from ppt-master)

- **Native shapes** (`use_native_shapes=True`): SVG → DrawingML → real PowerPoint objects. Editable, recolorable.
- **Compatibility mode** (`use_native_shapes=False`): SVG rendered as PNG + embedded SVG fallback. Works in older Office.

---

## What Each Module Does

### `extractor/pptx_parser.py`

Parses a `.pptx` file using `python-pptx` and produces a `ParsedSlide` dataclass containing:
- Every shape with its type, position, size
- All text runs with formatting (font, size, color, bold, italic)
- Embedded images as raw bytes
- Table data as 2D string arrays
- Raw XML for deep inspection

### `extractor/design_extractor.py`

Analyzes a `ParsedSlide` and produces `DesignTokens`:
- Color palette (background, primary, accent, text colors) via Counter analysis
- Font families and size ramp via frequency analysis
- Margins and title position via position math
- Follows the `spec_lock.md` format conventions exactly

### `designer/spec_builder.py`

Builds two markdown files from `DesignTokens`:
- `design_spec.md` — human-readable, follows ppt-master's §I–XI template
- `spec_lock.md` — machine-readable, exact format matching ppt-master

### `generator/svg_renderer.py`

The AI interface:
- `build_svg_generation_prompt()` — composes system + user prompts with all SVG constraints
- `OpenAiSvgBackend` / `AnthropicSvgBackend` — pluggable AI backends
- `generate_svg_slide()` — orchestrates the LLM call
- SVG extraction from LLM responses (handles ```svg fences)

### `converter/svg_to_pptx.py`

Thin wrapper over ppt-master:
- `svg_file_to_pptx()` — convert SVG file → PPTX
- `svg_string_to_pptx()` — convert SVG string → PPTX (writes temp file)
- `convert_svg_to_drawingml()` — low-level debug: see the DrawingML XML

### `utils/image_utils.py`

- `extract_images_from_parsed_slide()` — save embedded images to disk
- `validate_svg_quick()` — lightweight pre-check before full quality checker
- `embed_images_base64()` — inline images for self-contained SVGs
- `make_run_dir()` — timestamped output directories

---

## Extension Points

### Adding a new AI backend

Implement the `SvgGeneratorBackend` protocol:

```python
class MyCustomBackend:
    def generate(self, system_prompt: str, user_prompt: str) -> SvgGenerationResult:
        # Call your LLM here
        return SvgGenerationResult(svg_content="<svg>...</svg>")
```

### Template-based SVG generation (non-AI)

Create Jinja2 templates under `generator/templates/` and add a `TemplateSvgBackend` that renders them with the design tokens instead of calling an LLM.

### Multi-slide support

The pipeline is designed for single-slide regeneration. To extend:
1. Loop `parse_pptx_slide()` over all slides
2. Extract design once from slide 0, apply to all
3. Generate SVG per slide
4. Pass all SVGs to `create_pptx_with_native_svg(svg_files=[...])`

### Post-generation quality check

After SVG generation, optionally run ppt-master's full checker:

```python
# In your pipeline:
import subprocess
subprocess.run([
    "python3",
    "skills/ppt-master/scripts/svg_quality_checker.py",
    str(run_dir / "svg_output"),
])
```

---

## Dependencies on ppt-master

**None at runtime.** The `svg_to_pptx/` and `svg_finalize/` packages are bundled directly under `converter/`. The project is fully self-contained and can be deployed independently.

| Bundled component | Original source | Files |
|-------------------|----------------|-------|
| SVG → DrawingML engine | `skills/ppt-master/scripts/svg_to_pptx/` | 17 files |
| SVG post-processing | `skills/ppt-master/scripts/svg_finalize/` | 8 files |
| Canvas formats | `skills/ppt-master/scripts/config.py` | Adapted into `converter/config.py` |
| Project utilities | `skills/ppt-master/scripts/project_utils.py` | Minimal stubs in `converter/project_utils.py` |

## Supported AI Backends

| Backend | Class | .env Key | Default Model |
|---------|-------|----------|---------------|
| OpenAI | `OpenAiSvgBackend` | `OPENAI_API_KEY` | `gpt-4o` |
| Anthropic | `AnthropicSvgBackend` | `ANTHROPIC_API_KEY` | `claude-sonnet-4-20250514` |
| DeepSeek | `DeepSeekSvgBackend` | `DEEPSEEK_API_KEY` | `deepseek-chat` |
| Qwen | `QwenSvgBackend` | `QWEN_API_KEY` | `qwen-plus` |

All backends auto-load API keys from the `.env` file. Override with `api_key=` parameter or `--api-key` CLI flag.
