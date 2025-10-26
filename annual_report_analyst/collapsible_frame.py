"""Reusable collapsible frame widget for the Tkinter UI."""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk


class CollapsibleFrame(ttk.Frame):
    """A frame with a header that can toggle the visibility of its content."""

    def __init__(self, parent: tk.Widget, title: str, *, initially_open: bool = True) -> None:
        super().__init__(parent)
        self._title = title
        self._open = initially_open
        self._header = ttk.Button(self, text=self._formatted_title(), command=self._toggle, style="Toolbutton")
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
