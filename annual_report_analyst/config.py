"""Application-wide configuration constants."""

from __future__ import annotations

import re
from typing import Dict, List

COLUMNS: List[str] = ["Financial", "Income", "Shares"]

DEFAULT_PATTERNS: Dict[str, List[str]] = {
    "Financial": ["statement of financial position"],
    "Income": ["statement of profit or loss"],
    "Shares": ["Movements in issued capital"],
}

YEAR_DEFAULT_PATTERNS: List[str] = [r"(\d{4})\s+Annual\s+Report"]

DEFAULT_NOTE_OPTIONS: List[str] = ["", "excluded", "negated", "share_count"]

MAX_COMBINED_DATE_COLUMNS = 2

DEFAULT_NOTE_BACKGROUND_COLORS: Dict[str, str] = {
    "": "",
    "excluded": "#ff4d4f",
    "negated": "#4da6ff",
    "share_count": "#ffb6c1",
}

HEX_COLOR_RE = re.compile(r"^#(?:[0-9a-fA-F]{6})$")

DEFAULT_NOTE_KEY_BINDINGS: Dict[str, str] = {
    "": "`",
    "excluded": "1",
    "negated": "2",
    "share_count": "3",
}

SPECIAL_KEYSYM_ALIASES: Dict[str, str] = {
    "`": "grave",
    " ": "space",
}

SHIFT_MASK = 0x0001
