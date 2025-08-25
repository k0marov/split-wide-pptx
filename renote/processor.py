from __future__ import annotations

from dataclasses import dataclass
from typing import Optional

from .transforms import split_into_thirds_direct
from .vlm import classify_slides


@dataclass
class ProcessingOptions:
    title_min_font_pt: float = 36.0
    title_min_width_ratio: float = 1.2
    vlm_mode: str = "heuristic"  # heuristic | ollama
    soffice_path: Optional[str] = None
    ollama_model: str = "llava:latest"
    ollama_url: str = "http://localhost:11434"


def process_pptx(
    input_pptx: str,
    output_pptx: str,
    *,
    options: Optional[ProcessingOptions] = None,
    direct: bool = True,
) -> None:
    """High-level scenario: direct split into thirds.

    Args:
        input_pptx: source presentation path
        output_pptx: destination presentation path
        options: detection/tuning parameters
        triplicate_first: if True, create an intermediate triplicated file
    """
    opts = options or ProcessingOptions()

    # Decide scenario per slide using VLM/heuristics
    scenarios = classify_slides(
        input_pptx,
        mode=opts.vlm_mode,
        title_min_font_pt=opts.title_min_font_pt,
        title_min_width_ratio=opts.title_min_width_ratio,
        soffice_path=opts.soffice_path,
        ollama_model=opts.ollama_model,
        ollama_url=opts.ollama_url,
    )

    split_into_thirds_direct(
        input_pptx,
        output_pptx,
        title_min_font_pt=opts.title_min_font_pt,
        title_min_width_ratio=opts.title_min_width_ratio,
        scenarios=scenarios,
    )


