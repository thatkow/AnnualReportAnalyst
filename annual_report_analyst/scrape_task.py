"""Task description for scraping AI responses for PDF matches."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class ScrapeTask:
    entry_path: Path
    entry_name: str
    entry_year: str
    category: str
    page_index: int
    prompt_text: str
    page_text: str
    doc_dir: Path
