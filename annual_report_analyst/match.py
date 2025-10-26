"""Data structures related to PDF match information."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Optional


@dataclass
class Match:
    """Represents a single pattern match within a PDF document."""

    page_index: int
    source: str
    pattern: Optional[str] = None
