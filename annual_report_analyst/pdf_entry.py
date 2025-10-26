"""Representation of a PDF document that has been scanned for matches."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional

import fitz  # PyMuPDF

from .config import COLUMNS
from .match import Match


@dataclass
class PDFEntry:
    """Container for a PDF document and its associated match data."""

    path: Path
    doc: fitz.Document
    matches: Dict[str, List[Match]] = field(default_factory=dict)
    current_index: Dict[str, Optional[int]] = field(default_factory=dict)
    year: str = ""

    def __post_init__(self) -> None:
        for column in COLUMNS:
            self.matches.setdefault(column, [])
            self.current_index.setdefault(column, 0 if self.matches[column] else None)

    @property
    def stem(self) -> str:
        return self.path.stem
