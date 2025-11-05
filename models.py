from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional

from pdf_utils import PDFEntry


@dataclass
class ScrapeJob:
    entry: PDFEntry
    category: str
    pages: List[int]
    prompt_text: str
    model_name: str
    upload_mode: str
    target_dir: Path
    temp_pdf: Optional[Path] = None
    text_payload: Optional[str] = None
