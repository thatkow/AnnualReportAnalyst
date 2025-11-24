"""Shared constants for the Annual Report Analyst application."""

from __future__ import annotations

from typing import Dict, List

COLUMNS = ["Financial", "Income", "Shares"]
DEFAULT_PATTERNS = {
    "Financial": ["statement of financial position"],
    "Income": ["statement of profit or loss"],
    "Shares": ["Movements in issued capital"],
}
YEAR_DEFAULT_PATTERNS = [r"(\d{4})\s+Annual\s+Report"]

DEFAULT_OPENAI_MODEL = "gpt-4o-mini"

SCRAPE_EXPECTED_COLUMNS = [
    "CATEGORY",
    "SUBCATEGORY",
    "ITEM",
    "NOTE",
    "30.06.2023",
    "30.06.2022",
]
SCRAPE_PLACEHOLDER_ROWS = 10

SHIFT_MASK = 0x0001
CONTROL_MASK = 0x0004

# Default note color scheme and fallback palette
DEFAULT_NOTE_COLOR_SCHEME: Dict[str, str] = {
    "1": "#FFF2CC",  # light yellow
    "2": "#E2EFDA",  # light green
    "3": "#FCE4D6",  # light orange
    "4": "#E7E6FF",  # light purple
    "5": "#DDEBF7",  # light blue
    "6": "#F4CCCC",  # light red/pink
    "7": "#D9EAD3",
    "8": "#CFE2F3",
    "9": "#FFD966",
    "10": "#D9D2E9",
    "goodwill": "#FFEFD5",  # peach puff
}
FALLBACK_NOTE_PALETTE: List[str] = [
    "#FFF2CC", "#E2EFDA", "#FCE4D6", "#E7E6FF", "#DDEBF7",
    "#F4CCCC", "#D9EAD3", "#CFE2F3", "#FFD966", "#D9D2E9",
    "#F8CBAD", "#C9DAF8", "#EAD1DC", "#D0E0E3", "#E6B8AF",
]
