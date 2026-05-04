"""
SVG Generator — produces SVG slides guided by a design spec.

This module defines the interface for AI-driven SVG generation (Step 4).
The actual LLM call is abstracted behind ``SvgGeneratorBackend`` so you can
swap between OpenAI, Anthropic, DeepSeek, Qwen, or any OpenAI-compatible
provider without changing the pipeline.

API keys are loaded from a local ``.env`` file (via python-dotenv) so no
secrets need to live in code.

Supported backends and their .env keys:
  - OpenAI:      OPENAI_API_KEY
  - Anthropic:   ANTHROPIC_API_KEY
  - DeepSeek:    DEEPSEEK_API_KEY
  - Qwen:        QWEN_API_KEY (DASHSCOPE_API_KEY also accepted)
"""

from __future__ import annotations

import os
from pathlib import Path
from dataclasses import dataclass, field
from typing import Protocol

from ..extractor.pptx_parser import ParsedSlide
from ..designer.spec_builder import DesignTokens


# ---------------------------------------------------------------------------
# .env loading
# ---------------------------------------------------------------------------

def _load_dotenv() -> None:
    """Load environment variables from the local .env file (if present)."""
    try:
        from dotenv import load_dotenv as _load
    except ImportError:
        return  # python-dotenv not installed — rely on system env vars

    # Look for .env next to this file's package root (gen_svg_slide/)
    env_path = Path(__file__).resolve().parent.parent / ".env"
    if env_path.exists():
        _load(env_path)


# Load on module import so all backends see the vars
_load_dotenv()


# ---------------------------------------------------------------------------
# Backend protocol — swap AI providers here
# ---------------------------------------------------------------------------

@dataclass
class SvgGenerationResult:
    """Result of a single SVG generation call."""
    svg_content: str
    page_index: int = 1
    model: str = ""
    usage: dict = field(default_factory=dict)  # token counts, etc.


class SvgGeneratorBackend(Protocol):
    """Protocol for AI backends that generate SVG from a prompt."""

    def generate(self, system_prompt: str, user_prompt: str) -> SvgGenerationResult:
        """Generate SVG given system and user prompts.

        Args:
            system_prompt: System-level instructions (design constraints, SVG rules).
            user_prompt: Content + design spec to render.

        Returns:
            SvgGenerationResult with the generated SVG markup.
        """
        ...


# ---------------------------------------------------------------------------
# Prompt builder — composes the AI prompt from parsed slide + design tokens
# ---------------------------------------------------------------------------

def build_svg_generation_prompt(
    parsed: ParsedSlide,
    tokens: DesignTokens,
    design_spec: str,
    spec_lock: str,
    user_instruction: str = "",
) -> tuple[str, str]:
    """Build the (system_prompt, user_prompt) for SVG generation.

    The system prompt encodes all the SVG constraints from ppt-master's
    shared-standards.md. The user prompt provides the design spec and
    content to render.

    Args:
        parsed: The parsed reference slide.
        tokens: Extracted design tokens.
        design_spec: Human-readable design spec markdown.
        spec_lock: Machine-readable spec lock markdown.
        user_instruction: Optional additional instructions from the user.

    Returns:
        (system_prompt, user_prompt) tuple ready for the LLM.
    """
    system_prompt = f"""You are an expert SVG designer. Generate valid, clean SVG code for a presentation slide.

## CRITICAL SVG CONSTRAINTS (DO NOT VIOLATE):
1. Use viewBox="{tokens.viewbox}" — NO other viewBox values.
2. NO <mask>, NO <style>, NO class attributes, NO <foreignObject>.
3. NO <symbol> + <use>, NO textPath, NO @font-face, NO <animate*>, NO <script>.
4. NO rgba() — use fill="#RRGGBB" + fill-opacity="X.XX" instead.
5. NO <g opacity="..."> — set opacity on each child element individually.
6. ALL text must use raw Unicode characters (write — not &mdash;, → not &rarr;).
7. Escape XML reserved chars: & -> &amp;, < -> &lt;, > -> &gt;.
8. Text elements must use font-family, font-size, fill as inline attributes.
9. Do NOT use HTML named entities like &nbsp; &mdash; &copy; — use raw Unicode.
10. Embed images via <image href="data:image/png;base64,..."/> if needed.

## DESIGN SPEC TO FOLLOW:
{design_spec}

## EXECUTION LOCK (MUST use these exact values):
{spec_lock}

Output ONLY the SVG code wrapped in ```svg ... ``` — no explanations."""

    # Build content summary from parsed slide
    content_lines = ["## SLIDE CONTENT TO RENDER:"]
    for shape in parsed.shapes:
        if shape.paragraphs:
            text = " ".join(p.plain_text for p in shape.paragraphs if p.plain_text.strip())
            if text.strip():
                pos = f"({shape.left_emu/9525:.0f}, {shape.top_emu/9525:.0f})"
                content_lines.append(f"- [{shape.shape_type}] {pos}: {text.strip()}")
        elif shape.shape_type == "image":
            content_lines.append(f"- [image] ({shape.left_emu/9525:.0f}, {shape.top_emu/9525:.0f}) {shape.width_emu/9525:.0f}x{shape.height_emu/9525:.0f}px")
        elif shape.shape_type == "table":
            content_lines.append(f"- [table] ({shape.left_emu/9525:.0f}, {shape.top_emu/9525:.0f}) {len(shape.table_data)}r x {len(shape.table_data[0]) if shape.table_data else 0}c")
            for row in shape.table_data[:3]:  # first 3 rows only
                content_lines.append(f"  | {' | '.join(row)} |")

    if user_instruction:
        content_lines.append(f"\n## USER INSTRUCTION:\n{user_instruction}")

    user_prompt = "\n".join(content_lines)

    return system_prompt, user_prompt


# ---------------------------------------------------------------------------
# OpenAI-compatible base (shared by OpenAI, DeepSeek, Qwen, and others)
# ---------------------------------------------------------------------------

class _OpenAiCompatibleBackend:
    """Base for any OpenAI-compatible chat completions API.

    Subclasses set ``_default_model``, ``_default_base_url``, and
    ``_env_key`` to specialise for a particular provider.
    """

    _default_model: str = "gpt-4o"
    _default_base_url: str | None = None
    _env_key: str = "OPENAI_API_KEY"

    def __init__(
        self,
        api_key: str | None = None,
        model: str | None = None,
        base_url: str | None = None,
    ):
        self.api_key = api_key or os.environ.get(self._env_key, "")
        self.model = model or self._default_model
        self.base_url = base_url or self._default_base_url

        if not self.api_key:
            raise ValueError(
                f"No API key found for {self.__class__.__name__}. "
                f"Set {self._env_key} in .env or pass api_key= explicitly."
            )

    def generate(self, system_prompt: str, user_prompt: str) -> SvgGenerationResult:
        try:
            from openai import OpenAI
        except ImportError:
            raise ImportError(
                "openai package required for this backend: pip install openai"
            )

        client = OpenAI(api_key=self.api_key, base_url=self.base_url)
        response = client.chat.completions.create(
            model=self.model,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            temperature=0.3,
            max_tokens=8000,
        )

        content = response.choices[0].message.content or ""
        svg = _extract_svg_from_response(content)

        return SvgGenerationResult(
            svg_content=svg,
            model=self.model,
            usage={
                "prompt_tokens": response.usage.prompt_tokens if response.usage else 0,
                "completion_tokens": response.usage.completion_tokens if response.usage else 0,
            },
        )


# ---------------------------------------------------------------------------
# Concrete backends
# ---------------------------------------------------------------------------

class OpenAiSvgBackend(_OpenAiCompatibleBackend):
    """SVG generation via OpenAI API.

    Requires: ``OPENAI_API_KEY`` in .env or environment.
    Default model: ``gpt-4o``.
    """

    _default_model = "gpt-4o"
    _default_base_url = None  # uses OpenAI default
    _env_key = "OPENAI_API_KEY"


class DeepSeekSvgBackend(_OpenAiCompatibleBackend):
    """SVG generation via DeepSeek API (OpenAI-compatible).

    Requires: ``DEEPSEEK_API_KEY`` in .env or environment.
    Default model: ``deepseek-chat``.
    Base URL: ``https://api.deepseek.com``.
    """

    _default_model = "deepseek-chat"
    _default_base_url = "https://api.deepseek.com"
    _env_key = "DEEPSEEK_API_KEY"


class QwenSvgBackend(_OpenAiCompatibleBackend):
    """SVG generation via Qwen (Tongyi Qianwen) API — OpenAI-compatible.

    Requires: ``QWEN_API_KEY`` (or ``DASHSCOPE_API_KEY``) in .env or environment.
    Default model: ``qwen-plus``.
    Base URL: ``https://dashscope.aliyuncs.com/compatible-mode/v1``.
    """

    _default_model = "qwen-plus"
    _default_base_url = "https://dashscope.aliyuncs.com/compatible-mode/v1"
    _env_key = "QWEN_API_KEY"

    def __init__(
        self,
        api_key: str | None = None,
        model: str | None = None,
        base_url: str | None = None,
    ):
        # Qwen also accepts DASHSCOPE_API_KEY as fallback
        resolved = api_key or os.environ.get("QWEN_API_KEY") or os.environ.get("DASHSCOPE_API_KEY", "")
        super().__init__(api_key=resolved, model=model, base_url=base_url)


# ---------------------------------------------------------------------------
# Anthropic backend
# ---------------------------------------------------------------------------

class AnthropicSvgBackend:
    """SVG generation via Anthropic Claude API.

    Requires: ``ANTHROPIC_API_KEY`` in .env or environment.
    Default model: ``claude-sonnet-4-20250514``.
    """

    _default_model = "claude-sonnet-4-20250514"
    _env_key = "ANTHROPIC_API_KEY"

    def __init__(
        self,
        api_key: str | None = None,
        model: str | None = None,
    ):
        self.api_key = api_key or os.environ.get(self._env_key, "")
        self.model = model or self._default_model

        if not self.api_key:
            raise ValueError(
                f"No API key found for AnthropicSvgBackend. "
                f"Set {self._env_key} in .env or pass api_key= explicitly."
            )

    def generate(self, system_prompt: str, user_prompt: str) -> SvgGenerationResult:
        """Call Anthropic API to generate SVG."""
        try:
            import anthropic
        except ImportError:
            raise ImportError("anthropic package required: pip install anthropic")

        client = anthropic.Anthropic(api_key=self.api_key)
        response = client.messages.create(
            model=self.model,
            max_tokens=8000,
            system=system_prompt,
            messages=[{"role": "user", "content": user_prompt}],
        )

        content = response.content[0].text if response.content else ""
        svg = _extract_svg_from_response(content)

        return SvgGenerationResult(
            svg_content=svg,
            model=self.model,
            usage={
                "input_tokens": response.usage.input_tokens if response.usage else 0,
                "output_tokens": response.usage.output_tokens if response.usage else 0,
            },
        )


# ---------------------------------------------------------------------------
# SVG extraction from LLM response
# ---------------------------------------------------------------------------

def _extract_svg_from_response(text: str) -> str:
    """Extract SVG code from an LLM response (may be wrapped in ```svg ... ```)."""
    import re

    # Try fenced code block
    m = re.search(r"```(?:svg|xml)\s*\n(.*?)\n```", text, re.DOTALL)
    if m:
        return m.group(1).strip()

    # Try raw <svg>...</svg>
    m = re.search(r"(<svg\b.*?</svg>)", text, re.DOTALL | re.IGNORECASE)
    if m:
        return m.group(1).strip()

    # Return as-is (best effort)
    return text.strip()


# ---------------------------------------------------------------------------
# Orchestrator: generate SVG from parsed slide + design tokens
# ---------------------------------------------------------------------------

def generate_svg_slide(
    parsed: ParsedSlide,
    tokens: DesignTokens,
    design_spec: str,
    spec_lock: str,
    backend: SvgGeneratorBackend,
    user_instruction: str = "",
) -> SvgGenerationResult:
    """Generate a single SVG slide from parsed content and design tokens.

    Args:
        parsed: Parsed reference slide.
        tokens: Extracted design tokens.
        design_spec: Human-readable design spec markdown.
        spec_lock: Machine-readable spec lock markdown.
        backend: AI backend to use for generation.
        user_instruction: Optional user-provided content changes.

    Returns:
        SvgGenerationResult with the generated SVG.
    """
    system_prompt, user_prompt = build_svg_generation_prompt(
        parsed, tokens, design_spec, spec_lock, user_instruction,
    )
    return backend.generate(system_prompt, user_prompt)
