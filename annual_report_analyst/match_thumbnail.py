"""UI representation of a match thumbnail in the review grid."""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk
from typing import Optional, TYPE_CHECKING

from PIL import ImageTk

from .config import SHIFT_MASK
from .match import Match

if TYPE_CHECKING:  # pragma: no cover - imported for type checking only
    from .category_row import CategoryRow


class MatchThumbnail:
    SELECTED_COLOR = "#1E90FF"
    UNSELECTED_COLOR = "#c3c3c3"

    def __init__(self, row: "CategoryRow", match_index: int, match: Match) -> None:
        self.row = row
        self.app = row.app
        self.entry = row.entry
        self.match = match
        self.match_index = match_index
        self.photo: Optional[ImageTk.PhotoImage] = None
        self.container = tk.Frame(row.inner, highlightthickness=1, highlightbackground=self.UNSELECTED_COLOR)
        self.container.pack(side=tk.LEFT, padx=4, pady=4)
        self.container.columnconfigure(0, weight=1)
        self.image_label = ttk.Label(self.container)
        self.image_label.grid(row=0, column=0, sticky="nsew")
        self.info_label = ttk.Label(self.container, anchor="center", justify=tk.CENTER)
        self.info_label.grid(row=1, column=0, sticky="ew", pady=(4, 0))
        for widget in (self.container, self.image_label, self.info_label):
            widget.bind("<Button-1>", self._on_click)
            widget.bind("<Double-Button-1>", self._open_pdf)
            widget.bind("<Button-3>", self._open_context_menu)
        self._context_menu: Optional[tk.Menu] = None
        self.refresh()

    def refresh(self) -> None:
        target_width = self.row.target_width
        photo = self.app._render_page(self.entry.doc, self.match.page_index, target_width)
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
        self.set_selected(self.match_index == self.entry.current_index.get(self.row.category))

    def destroy(self) -> None:
        self.container.destroy()

    def set_selected(self, selected: bool) -> None:
        color = self.SELECTED_COLOR if selected else self.UNSELECTED_COLOR
        thickness = 3 if selected else 1
        self.container.configure(highlightbackground=color, highlightcolor=color, highlightthickness=thickness)

    def _ensure_context_menu(self) -> tk.Menu:
        if self._context_menu is None:
            menu = tk.Menu(self.container, tearoff=False)
            menu.add_command(label="Open PDF", command=lambda: self.app._open_pdf(self.entry.path, self.match.page_index))
            menu.add_command(
                label="Manual Entry",
                command=lambda: self.app.manual_select(self.entry, self.row.category),
            )
            self._context_menu = menu
        return self._context_menu

    def _on_click(self, event: tk.Event) -> Optional[str]:  # type: ignore[override]
        state = getattr(event, "state", 0)
        if state & SHIFT_MASK:
            self.app.open_thumbnail_zoom(self.entry, self.match.page_index)
            return "break"
        self.app.select_match_index(self.entry, self.row.category, self.match_index)
        return None

    def _open_pdf(self, _: tk.Event) -> None:  # type: ignore[override]
        self.app._open_pdf(self.entry.path, self.match.page_index)

    def _open_context_menu(self, event: tk.Event) -> None:  # type: ignore[override]
        menu = self._ensure_context_menu()
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()
