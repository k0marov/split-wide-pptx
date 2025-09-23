import base64
import os
import shutil
import subprocess
import tempfile
from typing import List, Optional

import requests
from pptx import Presentation

from .detectors import is_title_slide

def classify_slides(
    pptx_path: str,
    *,
    title_min_font_pt: float = 36.0,
    title_min_width_ratio: float = 1.2,
) -> List[str]:
    prs = Presentation(pptx_path)
    third_width = int(prs.slide_width // 3)
    return [
        "title" if is_title_slide(slide, third_width, min_font_pt=title_min_font_pt, min_width_ratio=title_min_width_ratio) else "split"
        for slide in prs.slides
    ]


