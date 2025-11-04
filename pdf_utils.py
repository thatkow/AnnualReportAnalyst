"""PDF-related utilities and models for the Annual Report Analyst."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, TYPE_CHECKING

import tkinter as tk
from tkinter import ttk

try:
    import fitz  # type: ignore[import-untyped]
except ImportError:  # pragma: no cover - handled at runtime
    fitz = None  # type: ignore[assignment]

from PIL import ImageTk

from constants import COLUMNS, CONTROL_MASK, SCRAPE_EXPECTED_COLUMNS, SHIFT_MASK

if TYPE_CHECKING:  # pragma: no cover
    from ui_widgets import CategoryRow
    from report_app import ReportAppV2


@dataclass
class Match:
    page_index: int
    source: str
    pattern: Optional[str] = None
    matched_text: Optional[str] = None


@dataclass
class PDFEntry:
    path: Path
    doc: "fitz.Document"
    matches: Dict[str, List[Match]] = field(default_factory=dict)
    current_index: Dict[str, Optional[int]] = field(default_factory=dict)
    selected_pages: Dict[str, List[int]] = field(default_factory=dict)
    year: str = ""

    def __post_init__(self) -> None:
        for column in COLUMNS:
            self.matches.setdefault(column, [])
            self.current_index.setdefault(column, 0 if self.matches[column] else None)
            self.selected_pages.setdefault(column, [])
            index = self.current_index.get(column)
            if index is not None and 0 <= index < len(self.matches[column]):
                page_index = self.matches[column][index].page_index
                self.selected_pages[column] = [page_index]


class MatchThumbnail:
    SELECTED_COLOR = "#1E90FF"
    UNSELECTED_COLOR = "#c3c3c3"
    MULTI_COLOR = "#FFD666"

    def __init__(self, row: "CategoryRow", match_index: int, match: Match) -> None:
        self.row = row
        self.app: "ReportAppV2" = row.app
        self.entry = row.entry
        self.match = match
        self.match_index = match_index
        self.photo: Optional[ImageTk.PhotoImage] = None

        self.container = tk.Frame(
            row.inner,
            highlightthickness=1,
            highlightbackground=self.UNSELECTED_COLOR,
        )
        self.container.pack(side=tk.LEFT, padx=4, pady=4)
        self.container.columnconfigure(0, weight=1)

        self.image_label = ttk.Label(self.container)
        self.image_label.grid(row=0, column=0, sticky="nsew")
        self.info_label = ttk.Label(self.container, anchor="center", justify=tk.CENTER)
        self.info_label.grid(row=1, column=0, sticky="ew", pady=(4, 0))

        for widget in (self.container, self.image_label, self.info_label):
            widget.bind("<Button-1>", self._on_click)
            widget.bind("<Double-Button-1>", self._open_pdf)

        self.refresh()

    def refresh(self) -> None:
        photo = self.app.render_page(
            self.entry.doc,
            self.match.page_index,
            target_width=self.row.target_width,
        )
        self.photo = photo
        if photo is not None:
            self.image_label.configure(image=photo, text="")
        else:
            self.image_label.configure(image="", text="Preview unavailable")

        info_parts = [f"Page {self.match.page_index + 1}"]
        if self.match.source == "manual":
            info_parts.append("manual")
        elif self.match.pattern:
            info_parts.append(self.match.pattern)
        self.info_label.configure(text=" | ".join(info_parts))
        self.update_state()

    def destroy(self) -> None:
        self.container.destroy()

    def update_state(self) -> None:
        current_index = self.entry.current_index.get(self.row.category)
        selected = current_index == self.match_index
        multi_pages = self.match.page_index in self.entry.selected_pages.get(
            self.row.category, []
        )
        if selected:
            color = self.SELECTED_COLOR
            thickness = 3
        elif multi_pages:
            color = self.MULTI_COLOR
            thickness = 2
        else:
            color = self.UNSELECTED_COLOR
            thickness = 1
        self.container.configure(
            highlightbackground=color,
            highlightcolor=color,
            highlightthickness=thickness,
        )

    def _on_click(self, event: tk.Event) -> None:  # type: ignore[override]
        try:
            state = int(event.state)
        except Exception:
            state = 0
        if state & CONTROL_MASK:
            self.app.toggle_fullscreen_preview(self.entry, self.match.page_index)
            return
        extend = bool(state & SHIFT_MASK)
        self.app.select_match(
            self.entry,
            self.row.category,
            self.match_index,
            extend_selection=extend,
        )

    def _open_pdf(self, _: tk.Event) -> None:  # type: ignore[override]
        self.app.open_pdf(self.entry.path)


def normalize_header_row(cells: List[str]) -> Optional[List[str]]:
    normalized = [cell.strip() for cell in cells]
    if not any(normalized):
        return None

    lower_values = [cell.lower() for cell in normalized]
    expected_lower = [column.lower() for column in SCRAPE_EXPECTED_COLUMNS]

    prefix_length = min(len(lower_values), len(expected_lower))
    prefix_matches = all(
        lower_values[idx] == expected_lower[idx] for idx in range(prefix_length)
    )
    if prefix_matches:
        return normalized[: len(lower_values)]

    if lower_values and lower_values[0] == "category":
        return normalized

    return None
