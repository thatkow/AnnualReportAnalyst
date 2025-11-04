"""Tkinter widgets used throughout the Annual Report Analyst UI."""

from __future__ import annotations

from typing import List, Optional, TYPE_CHECKING

import tkinter as tk
from tkinter import ttk

from pdf_utils import MatchThumbnail

if TYPE_CHECKING:  # pragma: no cover
    from pdf_utils import PDFEntry
    from report_app import ReportAppV2


class CollapsibleFrame(ttk.Frame):
    def __init__(self, master: tk.Widget, title: str, initially_open: bool = False) -> None:
        super().__init__(master)
        self._title = title
        self._open = initially_open
        self._header = ttk.Button(
            self,
            text=self._formatted_title(),
            command=self._toggle,
            style="Toolbutton",
        )
        self._header.pack(fill=tk.X)
        self._content = ttk.Frame(self)
        if self._open:
            self._content.pack(fill=tk.BOTH, expand=True)

    @property
    def content(self) -> ttk.Frame:
        return self._content

    def _formatted_title(self) -> str:
        return ("▼ " if self._open else "► ") + self._title

    def _toggle(self) -> None:
        self._open = not self._open
        if self._open:
            self._content.pack(fill=tk.BOTH, expand=True)
        else:
            self._content.pack_forget()
        self._header.configure(text=self._formatted_title())


class CategoryRow:
    def __init__(
        self,
        parent: tk.Widget,
        app: "ReportAppV2",
        entry: "PDFEntry",
        category: str,
    ) -> None:
        self.app = app
        self.entry = entry
        self.category = category

        self.frame = ttk.Frame(parent, padding=(0, 4, 0, 4))
        self.frame.columnconfigure(0, weight=1)

        header = ttk.Frame(self.frame)
        header.grid(row=0, column=0, sticky="ew")
        ttk.Label(header, text=category, font=("TkDefaultFont", 10, "bold")).pack(side=tk.LEFT)
        ttk.Button(header, text="Manual", width=7, command=self._manual_select).pack(side=tk.RIGHT)

        self.target_width = self.app.thumbnail_width_var.get()
        self.canvas = tk.Canvas(self.frame, height=self._compute_canvas_height())
        self.canvas.grid(row=1, column=0, sticky="ew")
        self.scrollbar = ttk.Scrollbar(self.frame, orient=tk.HORIZONTAL, command=self.canvas.xview)
        self.scrollbar.grid(row=2, column=0, sticky="ew")
        self.canvas.configure(xscrollcommand=self.scrollbar.set)

        self.inner = ttk.Frame(self.canvas)
        self.window = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")

        self.inner.bind("<Configure>", self._on_inner_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.canvas.bind("<Shift-MouseWheel>", self._on_mousewheel)
        self.canvas.bind("<Shift-Button-4>", self._on_mousewheel)
        self.canvas.bind("<Shift-Button-5>", self._on_mousewheel)

        self.thumbnails: List[MatchThumbnail] = []
        self.empty_label: Optional[ttk.Label] = None

    def refresh(self) -> None:
        for thumb in self.thumbnails:
            thumb.destroy()
        self.thumbnails.clear()
        if self.empty_label is not None:
            self.empty_label.destroy()
            self.empty_label = None

        matches = self.entry.matches.get(self.category, [])
        if not matches:
            self.empty_label = ttk.Label(self.inner, text="No matches found", foreground="#666666")
            self.empty_label.pack(side=tk.LEFT, padx=8, pady=16)
        else:
            for idx, match in enumerate(matches):
                thumbnail = MatchThumbnail(self, idx, match)
                self.thumbnails.append(thumbnail)

        self.update_selection()
        self.frame.after_idle(self._update_scrollbar_visibility)

    def update_selection(self) -> None:
        for thumb in self.thumbnails:
            thumb.update_state()

    def set_thumbnail_width(self, width: int) -> None:
        if width == self.target_width:
            return
        self.target_width = max(80, width)
        self.canvas.configure(height=self._compute_canvas_height())
        for thumb in self.thumbnails:
            thumb.refresh()
        self.frame.after_idle(self._update_scrollbar_visibility)

    def _manual_select(self) -> None:
        self.app.manual_select(self.entry, self.category)

    def _compute_canvas_height(self) -> int:
        return max(160, int(self.target_width * 1.2))

    def _on_inner_configure(self, _: tk.Event) -> None:  # type: ignore[override]
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event: tk.Event) -> None:  # type: ignore[override]
        self.canvas.itemconfigure(self.window, height=event.height)

    def _on_mousewheel(self, event: tk.Event) -> None:  # type: ignore[override]
        if getattr(event, "delta", 0):
            step = -1 if event.delta > 0 else 1
        else:
            step = -1 if event.num == 4 else 1
        self.canvas.xview_scroll(step, "units")

    def _update_scrollbar_visibility(self) -> None:
        bbox = self.canvas.bbox("all")
        if bbox is None:
            self.scrollbar.grid_remove()
            return
        content_width = bbox[2] - bbox[0]
        canvas_width = self.canvas.winfo_width()
        if content_width <= canvas_width:
            self.scrollbar.grid_remove()
        else:
            self.scrollbar.grid()
