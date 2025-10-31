from __future__ import annotations

import csv
import io
import json
import logging
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

try:
    import fitz  # type: ignore[import-untyped]
except ImportError:  # pragma: no cover - handled at runtime
    fitz = None  # type: ignore[assignment]

from PIL import Image, ImageTk

try:
    from openai import OpenAI
except ImportError:  # pragma: no cover - handled at runtime
    OpenAI = None  # type: ignore[assignment]


logger = logging.getLogger(__name__)


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


@dataclass
class Match:
    page_index: int
    source: str
    pattern: Optional[str] = None
    matched_text: Optional[str] = None


@dataclass
class PDFEntry:
    path: Path
    doc: fitz.Document
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


class CollapsibleFrame(ttk.Frame):
    def __init__(self, master: tk.Widget, title: str, initially_open: bool = False) -> None:
        super().__init__(master)
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


class MatchThumbnail:
    SELECTED_COLOR = "#1E90FF"
    UNSELECTED_COLOR = "#c3c3c3"
    MULTI_COLOR = "#FFD666"

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
        multi_pages = self.match.page_index in self.entry.selected_pages.get(self.row.category, [])
        if selected:
            color = self.SELECTED_COLOR
            thickness = 3
        elif multi_pages:
            color = self.MULTI_COLOR
            thickness = 2
        else:
            color = self.UNSELECTED_COLOR
            thickness = 1
        self.container.configure(highlightbackground=color, highlightcolor=color, highlightthickness=thickness)

    def _on_click(self, event: tk.Event) -> None:  # type: ignore[override]
        try:
            state = int(event.state)
        except Exception:
            state = 0
        if state & CONTROL_MASK:
            self.app.toggle_fullscreen_preview(self.entry, self.match.page_index)
            return
        extend = bool(state & SHIFT_MASK)
        self.app.select_match(self.entry, self.row.category, self.match_index, extend_selection=extend)

    def _open_pdf(self, _: tk.Event) -> None:  # type: ignore[override]
        self.app.open_pdf(self.entry.path)


class CategoryRow:
    def __init__(self, parent: tk.Widget, app: "ReportAppV2", entry: PDFEntry, category: str) -> None:
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


class ScrapeResultPanel:
    def __init__(
        self,
        parent: tk.Widget,
        app: "ReportAppV2",
        entry: PDFEntry,
        category: str,
        target_dir: Path,
    ) -> None:
        self.app = app
        self.entry = entry
        self.category = category
        self.target_dir = target_dir
        self.csv_path = target_dir / f"{category}.csv"
        self.multiplier_path = target_dir / f"{category}_multiplier.txt"
        self.raw_path = target_dir / f"{category}_raw.txt"
        self.has_csv_data = False
        self._updating_multiplier = False

        self.container = tk.Frame(
            parent,
            highlightbackground="#c3c3c3",
            highlightcolor="#c3c3c3",
            highlightthickness=1,
            bd=1,
            relief=tk.FLAT,
        )
        self.container.pack(fill=tk.X, pady=(0, 8))

        self.frame = ttk.Frame(self.container, padding=8)
        self.frame.pack(fill=tk.BOTH, expand=True)

        header = ttk.Frame(self.frame)
        header.pack(fill=tk.X)
        title_text = f"{entry.path.name} – {category}"
        self.title_label = ttk.Label(header, text=title_text, font=("TkDefaultFont", 10, "bold"))
        self.title_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        multiplier_box = ttk.Frame(header)
        multiplier_box.pack(side=tk.RIGHT)
        ttk.Label(multiplier_box, text="Multiplier:").pack(side=tk.LEFT)
        self.multiplier_var = tk.StringVar(master=self.frame)
        self.multiplier_entry = ttk.Entry(multiplier_box, textvariable=self.multiplier_var, width=16)
        self.multiplier_entry.pack(side=tk.LEFT, padx=(4, 0))
        self.multiplier_entry.bind("<FocusOut>", self._on_multiplier_changed)
        self.multiplier_entry.bind("<Return>", self._on_multiplier_submit)
        self.multiplier_entry.bind("<KP_Enter>", self._on_multiplier_submit)

        actions_row = ttk.Frame(self.frame)
        actions_row.pack(fill=tk.X, pady=(6, 0))
        self.open_csv_button = ttk.Button(actions_row, text="Open CSV", command=self.open_csv)
        self.open_csv_button.pack(side=tk.LEFT)
        self.view_raw_button = ttk.Button(actions_row, text="View Raw", command=self.view_raw_text)
        self.view_raw_button.pack(side=tk.LEFT, padx=(6, 0))

        table_container = ttk.Frame(self.frame)
        table_container.pack(fill=tk.BOTH, expand=True, pady=(6, 0))
        self.table = ttk.Treeview(
            table_container,
            columns=SCRAPE_EXPECTED_COLUMNS,
            show="headings",
            height=SCRAPE_PLACEHOLDER_ROWS,
        )
        for column in SCRAPE_EXPECTED_COLUMNS:
            self.table.heading(column, text=column)
            anchor = tk.W if column in {"CATEGORY", "SUBCATEGORY", "ITEM", "NOTE"} else tk.E
            self.table.column(column, anchor=anchor, width=140, stretch=True)
        self.table.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar = ttk.Scrollbar(table_container, orient=tk.VERTICAL, command=self.table.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.table.configure(yscrollcommand=scrollbar.set)

        for widget in (
            self.container,
            self.frame,
            header,
            self.title_label,
            table_container,
            self.table,
        ):
            widget.bind("<Button-1>", self._handle_activate, add="+")

        self.set_placeholder("-")
        self._update_action_states()

    def destroy(self) -> None:
        self.container.destroy()

    def set_placeholder(self, fill: str) -> None:
        rows = [[fill for _ in SCRAPE_EXPECTED_COLUMNS] for _ in range(SCRAPE_PLACEHOLDER_ROWS)]
        self._populate(rows)
        self.has_csv_data = False
        self._update_action_states()

    def mark_loading(self) -> None:
        if self.has_csv_data:
            return
        rows = [["?" for _ in SCRAPE_EXPECTED_COLUMNS] for _ in range(SCRAPE_PLACEHOLDER_ROWS)]
        self._populate(rows)
        self._update_action_states()

    def load_from_files(self) -> None:
        rows: List[List[str]] = []
        if self.csv_path.exists():
            try:
                with self.csv_path.open("r", encoding="utf-8", newline="") as fh:
                    reader = csv.reader(fh)
                    for raw_row in reader:
                        if any(cell.strip() for cell in raw_row):
                            rows.append([cell.strip() for cell in raw_row])
            except OSError:
                rows = []

        if rows:
            self._populate(rows)
            self.has_csv_data = True
        else:
            self.set_placeholder("-")

        if self.multiplier_path.exists():
            try:
                text = self.multiplier_path.read_text(encoding="utf-8").strip()
            except OSError:
                text = ""
            self._set_multiplier(text)
        else:
            self._set_multiplier("")
        self._update_action_states()

    def set_multiplier(self, value: str) -> None:
        self._set_multiplier(value)

    def _set_multiplier(self, value: str) -> None:
        self._updating_multiplier = True
        self.multiplier_var.set(value)
        self._updating_multiplier = False

    def save_multiplier(self) -> None:
        if self._updating_multiplier:
            return
        value = self.multiplier_var.get().strip()
        try:
            self.target_dir.mkdir(parents=True, exist_ok=True)
        except OSError:
            logger.exception("Unable to ensure scrape directory exists: %s", self.target_dir)
            return
        try:
            if value:
                self.multiplier_path.write_text(value, encoding="utf-8")
            elif self.multiplier_path.exists():
                self.multiplier_path.unlink()
        except OSError:
            logger.exception("Unable to persist multiplier for %s - %s", self.entry.path.name, self.category)

    def set_active(self, active: bool) -> None:
        color = "#1E90FF" if active else "#c3c3c3"
        thickness = 2 if active else 1
        self.container.configure(highlightbackground=color, highlightcolor=color, highlightthickness=thickness)

    def open_csv(self) -> None:
        if not self.csv_path.exists():
            messagebox.showinfo("Open CSV", "CSV file not available yet.")
            return
        self.app.open_file_path(self.csv_path)

    def view_raw_text(self) -> None:
        if not self.raw_path.exists():
            messagebox.showinfo("Raw Response", "Raw response not available yet.")
            return
        self.app.show_raw_text_dialog(
            self.raw_path,
            f"{self.entry.path.name} – {self.category} raw response",
        )

    def _populate(self, rows: List[List[str]]) -> None:
        for item in self.table.get_children(""):
            self.table.delete(item)
        column_count = len(SCRAPE_EXPECTED_COLUMNS)
        for row in rows:
            values = list(row[:column_count])
            if len(values) < column_count:
                values.extend([""] * (column_count - len(values)))
            self.table.insert("", "end", values=values)

    def _handle_activate(self, _: tk.Event) -> None:  # type: ignore[override]
        self.app._on_scrape_panel_clicked(self)

    def _on_multiplier_changed(self, _: tk.Event) -> None:  # type: ignore[override]
        self.save_multiplier()

    def _on_multiplier_submit(self, _: tk.Event) -> str:  # type: ignore[override]
        self.save_multiplier()
        return "break"

    def _update_action_states(self) -> None:
        has_csv = self.csv_path.exists()
        self.open_csv_button.configure(state="normal" if has_csv else "disabled")
        has_raw = self.raw_path.exists()
        self.view_raw_button.configure(state="normal" if has_raw else "disabled")


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


class ReportAppV2:
    def __init__(self, root: tk.Misc) -> None:
        if fitz is None:
            messagebox.showerror("PyMuPDF Required", "Install PyMuPDF (import name 'fitz') to use this app.")
            raise RuntimeError("PyMuPDF (fitz) is not installed")

        self.root = root
        if hasattr(self.root, "title"):
            try:
                self.root.title("Annual Report Analyst (Preview)")
            except tk.TclError:
                pass
        self.root.after(0, self._maximize_window)

        self.app_root = Path(__file__).resolve().parent
        self.companies_dir = self.app_root / "companies"
        self.config_path = self.app_root / "data2_config.json"
        self.pattern_config_path = self.app_root / "pattern_config.json"
        self.local_config_path = self.app_root / "local_config.json"
        self.prompts_dir = self.app_root / "prompts"

        self.company_var = tk.StringVar(master=self.root)
        self.folder_path = tk.StringVar(master=self.root)
        self.thumbnail_width_var = tk.IntVar(master=self.root, value=220)
        self.api_key_var = tk.StringVar(master=self.root)
        self.local_config_data: Dict[str, Any] = {}
        self._api_key_save_after: Optional[str] = None
        self._suspend_api_key_save = False
        self.api_key_var.trace_add("write", self._on_api_key_var_changed)

        self.pattern_texts: Dict[str, tk.Text] = {}
        self.case_insensitive_vars: Dict[str, tk.BooleanVar] = {}
        self.whitespace_as_space_vars: Dict[str, tk.BooleanVar] = {}
        self.year_pattern_text: Optional[tk.Text] = None
        self.year_case_insensitive_var = tk.BooleanVar(master=self.root, value=True)
        self.year_whitespace_as_space_var = tk.BooleanVar(master=self.root, value=True)
        self.openai_model_vars: Dict[str, tk.StringVar] = {}
        self.scrape_upload_mode_vars: Dict[str, tk.StringVar] = {}

        self.pdf_entries: List[PDFEntry] = []
        self.category_rows: Dict[Tuple[Path, str], CategoryRow] = {}
        self.assigned_pages: Dict[str, Dict[str, Any]] = {}
        self.assigned_pages_path: Optional[Path] = None
        self.fullscreen_preview_window: Optional[tk.Toplevel] = None
        self.fullscreen_preview_image: Optional[ImageTk.PhotoImage] = None
        self.fullscreen_preview_entry: Optional[Path] = None
        self.fullscreen_preview_page: Optional[int] = None

        self.scrape_panels: Dict[Tuple[Path, str], ScrapeResultPanel] = {}
        self.active_scrape_key: Optional[Tuple[Path, str]] = None
        self.scrape_preview_photo: Optional[ImageTk.PhotoImage] = None
        self.scrape_preview_pages: List[int] = []
        self.scrape_preview_entry: Optional[PDFEntry] = None
        self.scrape_preview_category: Optional[str] = None
        self.scrape_preview_cycle_index: int = 0
        self.scrape_preview_last_width: int = 0
        self.scrape_preview_render_width: int = 0
        self.scrape_preview_render_page: Optional[int] = None
        self._scrape_thread: Optional[threading.Thread] = None

        self.downloads_dir = tk.StringVar(master=self.root)
        self.recent_download_minutes = tk.IntVar(master=self.root, value=5)

        self._suspend_api_key_save = True
        self._load_local_config()
        self._suspend_api_key_save = False
        self._build_ui()
        self._load_pattern_config()
        self._load_config()
        self._refresh_company_options()

    # ------------------------------------------------------------------ UI
    def _build_ui(self) -> None:
        menu_bar = tk.Menu(self.root)
        file_menu = tk.Menu(menu_bar, tearoff=False)
        file_menu.add_command(label="New Company", command=self.create_company)
        file_menu.add_command(label="Set Downloads Dir", command=self._set_downloads_dir)
        menu_bar.add_cascade(label="File", menu=file_menu)
        try:
            self.root.config(menu=menu_bar)
        except tk.TclError:
            pass

        top = ttk.Frame(self.root, padding=8)
        top.pack(fill=tk.X)

        ttk.Label(top, text="Company:").pack(side=tk.LEFT)
        self.company_combo = ttk.Combobox(top, textvariable=self.company_var, state="readonly", width=30)
        self.company_combo.pack(side=tk.LEFT, padx=(4, 8))
        self.company_combo.bind("<<ComboboxSelected>>", self._on_company_selected)

        ttk.Button(top, text="Load PDFs", command=self.load_pdfs).pack(side=tk.LEFT)

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        review_tab = ttk.Frame(self.notebook)
        self.notebook.add(review_tab, text="Review")

        options_section = CollapsibleFrame(review_tab, "Patterns & Review Options", initially_open=False)
        options_section.pack(fill=tk.X, padx=8, pady=(4, 0))

        options_inner = ttk.Frame(options_section.content, padding=8)
        options_inner.pack(fill=tk.BOTH, expand=True)

        patterns_frame = ttk.LabelFrame(options_inner, text="Regex patterns (one per line)", padding=8)
        patterns_frame.pack(fill=tk.BOTH, expand=True)

        columns_frame = ttk.Frame(patterns_frame)
        columns_frame.pack(fill=tk.X)

        for idx, column in enumerate(COLUMNS):
            column_frame = ttk.Frame(columns_frame)
            column_frame.grid(row=0, column=idx, padx=4, sticky="nsew")
            columns_frame.columnconfigure(idx, weight=1)

            ttk.Label(column_frame, text=column).pack(anchor="w")
            text_widget = tk.Text(column_frame, height=4, width=30)
            text_widget.pack(fill=tk.BOTH, expand=True)
            defaults = DEFAULT_PATTERNS.get(column, [])
            text_widget.insert("1.0", "\n".join(defaults))
            self.pattern_texts[column] = text_widget

            model_var = tk.StringVar(master=self.root, value=DEFAULT_OPENAI_MODEL)
            self.openai_model_vars[column] = model_var

            case_var = tk.BooleanVar(master=self.root, value=True)
            self.case_insensitive_vars[column] = case_var
            ttk.Checkbutton(column_frame, text="Case-insensitive", variable=case_var).pack(anchor="w", pady=(4, 0))

            whitespace_var = tk.BooleanVar(master=self.root, value=True)
            self.whitespace_as_space_vars[column] = whitespace_var
            ttk.Checkbutton(
                column_frame,
                text="Treat spaces as any whitespace",
                variable=whitespace_var,
            ).pack(anchor="w")

        apply_button = ttk.Button(patterns_frame, text="Apply Patterns", command=self.load_pdfs)
        apply_button.pack(anchor="e", pady=(8, 0))

        year_frame = ttk.LabelFrame(options_inner, text="Year pattern", padding=8)
        year_frame.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
        year_text = tk.Text(year_frame, height=2, width=30)
        year_text.pack(fill=tk.BOTH, expand=True)
        year_text.insert("1.0", "\n".join(YEAR_DEFAULT_PATTERNS))
        self.year_pattern_text = year_text
        ttk.Checkbutton(year_frame, text="Case-insensitive", variable=self.year_case_insensitive_var).pack(anchor="w", pady=(4, 0))
        ttk.Checkbutton(
            year_frame,
            text="Treat spaces as any whitespace",
            variable=self.year_whitespace_as_space_var,
        ).pack(anchor="w")

        review_controls = ttk.Frame(options_inner, padding=(0, 8, 0, 0))
        review_controls.pack(fill=tk.X, pady=(8, 0))
        ttk.Label(review_controls, text="Thumbnail width:").pack(side=tk.LEFT)
        self.thumbnail_scale = ttk.Scale(
            review_controls,
            from_=160,
            to=420,
            orient=tk.HORIZONTAL,
            command=self._on_thumbnail_scale,
        )
        self.thumbnail_scale.set(self.thumbnail_width_var.get())
        self.thumbnail_scale.pack(side=tk.LEFT, padx=8, fill=tk.X, expand=True)
        ttk.Label(review_controls, textvariable=self.thumbnail_width_var).pack(side=tk.LEFT)

        review_container = ttk.Frame(review_tab)
        review_container.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

        self.review_canvas = tk.Canvas(review_container)
        self.review_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        review_scrollbar = ttk.Scrollbar(review_container, orient=tk.VERTICAL, command=self.review_canvas.yview)
        review_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.review_canvas.configure(yscrollcommand=review_scrollbar.set)
        self.review_canvas.bind("<Enter>", self._bind_review_mousewheel)
        self.review_canvas.bind("<Leave>", self._unbind_review_mousewheel)

        self.inner_frame = ttk.Frame(self.review_canvas)
        self.canvas_window = self.review_canvas.create_window((0, 0), window=self.inner_frame, anchor="nw")
        self.inner_frame.bind("<Configure>", lambda _e: self.review_canvas.configure(scrollregion=self.review_canvas.bbox("all")))
        self.review_canvas.bind(
            "<Configure>",
            lambda event: self.review_canvas.itemconfigure(self.canvas_window, width=event.width),
        )

        actions_frame = ttk.Frame(review_tab, padding=8)
        actions_frame.pack(fill=tk.X, padx=8, pady=(0, 8))
        self.commit_button = ttk.Button(actions_frame, text="Commit", command=self.commit_assignments)
        self.commit_button.pack(side=tk.RIGHT)

        scrape_tab = ttk.Frame(self.notebook)
        self.notebook.add(scrape_tab, text="Scrape")

        scrape_controls = ttk.Frame(scrape_tab, padding=8)
        scrape_controls.pack(fill=tk.X)
        ttk.Label(scrape_controls, text="API key:").pack(side=tk.LEFT)
        self.api_key_entry = ttk.Entry(scrape_controls, textvariable=self.api_key_var, width=40, show="*")
        self.api_key_entry.pack(side=tk.LEFT, padx=(4, 8))
        self.scrape_button = ttk.Button(scrape_controls, text="AIScrape", command=self.scrape_selected_pages)
        self.scrape_button.pack(side=tk.LEFT)
        self.scrape_progress = ttk.Progressbar(scrape_controls, orient=tk.HORIZONTAL, mode="determinate", length=200)
        self.scrape_progress.pack(side=tk.LEFT, padx=(8, 0), fill=tk.X, expand=True)

        model_frame = ttk.LabelFrame(scrape_tab, text="OpenAI models", padding=8)
        model_frame.pack(fill=tk.X, padx=8)
        for column in COLUMNS:
            row = ttk.Frame(model_frame)
            row.pack(fill=tk.X, pady=2)
            ttk.Label(row, text=f"{column}:").pack(side=tk.LEFT)
            entry = ttk.Entry(row, textvariable=self.openai_model_vars[column])
            entry.pack(side=tk.LEFT, padx=(4, 0), fill=tk.X, expand=True)
            mode_var = tk.StringVar(master=self.root, value="pdf")
            self.scrape_upload_mode_vars[column] = mode_var
            ttk.Radiobutton(row, text="Upload PDF", variable=mode_var, value="pdf").pack(
                side=tk.LEFT, padx=(8, 0)
            )
            ttk.Radiobutton(row, text="Extract text", variable=mode_var, value="text").pack(
                side=tk.LEFT, padx=(4, 0)
            )

        scrape_body = ttk.Frame(scrape_tab, padding=(8, 0, 8, 8))
        scrape_body.pack(fill=tk.BOTH, expand=True)

        scrape_split = ttk.Panedwindow(scrape_body, orient=tk.HORIZONTAL)
        scrape_split.pack(fill=tk.BOTH, expand=True)

        preview_frame = ttk.Frame(scrape_split, padding=(0, 0, 8, 0))
        scrape_split.add(preview_frame, weight=1)

        preview_header = ttk.Frame(preview_frame)
        preview_header.pack(fill=tk.X)
        self.scrape_preview_title_var = tk.StringVar(value="Select a section to preview.")
        ttk.Label(preview_header, textvariable=self.scrape_preview_title_var, font=("TkDefaultFont", 10, "bold")).pack(
            side=tk.LEFT,
            fill=tk.X,
            expand=True,
        )
        self.scrape_preview_page_var = tk.StringVar(value="")
        ttk.Label(preview_header, textvariable=self.scrape_preview_page_var).pack(side=tk.RIGHT)

        self.scrape_preview_label = tk.Label(
            preview_frame,
            text="Select a section to preview.",
            justify=tk.CENTER,
            anchor="center",
            background="#f0f0f0",
        )
        self.scrape_preview_label.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
        self.scrape_preview_label.bind("<Button-1>", self._on_scrape_preview_click)
        self.scrape_preview_label.bind("<Configure>", self._on_scrape_preview_resize)

        tables_frame = ttk.Frame(scrape_split)
        scrape_split.add(tables_frame, weight=1)

        self.scrape_pdf_notebook = ttk.Notebook(tables_frame)
        self.scrape_pdf_notebook.pack(fill=tk.BOTH, expand=True)
        self.scrape_pdf_tabs: Dict[Path, ttk.Frame] = {}
        self.scrape_pdf_category_notebooks: Dict[Path, ttk.Notebook] = {}
        self.scrape_category_tabs: Dict[Tuple[Path, str], ttk.Frame] = {}
        self.scrape_category_canvases: Dict[Tuple[Path, str], tk.Canvas] = {}
        self.scrape_category_inners: Dict[Tuple[Path, str], ttk.Frame] = {}
        self.scrape_category_windows: Dict[Tuple[Path, str], int] = {}
        self.scrape_category_placeholders: Dict[Tuple[Path, str], Optional[tk.Widget]] = {}

    def _maximize_window(self) -> None:
        try:
            self.root.state("zoomed")
        except tk.TclError:
            pass
        try:
            self.root.attributes("-zoomed", True)
        except tk.TclError:
            screen_width = self.root.winfo_screenwidth()
            screen_height = self.root.winfo_screenheight()
            try:
                self.root.geometry(f"{screen_width}x{screen_height}")
            except tk.TclError:
                pass

    # ------------------------------------------------------------------ Config
    def _collect_pattern_config_payload(self) -> Dict[str, Any]:
        patterns = {
            column: [line for line in self._read_text_lines(widget)]
            for column, widget in self.pattern_texts.items()
        }
        case_flags = {column: bool(var.get()) for column, var in self.case_insensitive_vars.items()}
        whitespace_flags = {
            column: bool(var.get()) for column, var in self.whitespace_as_space_vars.items()
        }
        year_patterns = []
        if self.year_pattern_text is not None:
            year_patterns = [line for line in self._read_text_lines(self.year_pattern_text)]
        openai_models: Dict[str, str] = {}
        for column in COLUMNS:
            var = self.openai_model_vars.get(column)
            if var is not None:
                value = var.get().strip()
            else:
                value = ""
            openai_models[column] = value or DEFAULT_OPENAI_MODEL
        upload_modes: Dict[str, str] = {}
        for column in COLUMNS:
            mode_var = self.scrape_upload_mode_vars.get(column)
            if mode_var is not None:
                upload_modes[column] = mode_var.get() or "pdf"
            else:
                upload_modes[column] = "pdf"
        payload = {
            "patterns": patterns,
            "case_insensitive": case_flags,
            "space_as_whitespace": whitespace_flags,
            "year_patterns": year_patterns,
            "year_case_insensitive": bool(self.year_case_insensitive_var.get()),
            "year_space_as_whitespace": bool(self.year_whitespace_as_space_var.get()),
            "downloads_minutes": int(self.recent_download_minutes.get()),
            "openai_models": openai_models,
            "upload_modes": upload_modes,
        }
        return payload

    def _save_pattern_config(self) -> None:
        payload = self._collect_pattern_config_payload()
        try:
            with self.pattern_config_path.open("w", encoding="utf-8") as fh:
                json.dump(payload, fh, indent=2)
        except OSError:
            messagebox.showwarning("Save Patterns", "Unable to save pattern configuration to disk.")

    def _load_pattern_config(self) -> None:
        if not self.pattern_config_path.exists():
            return
        try:
            with self.pattern_config_path.open("r", encoding="utf-8") as fh:
                data = json.load(fh)
        except (OSError, json.JSONDecodeError):
            messagebox.showwarning("Load Patterns", "Unable to read pattern configuration; using defaults.")
            return

        if not isinstance(data, dict):
            messagebox.showwarning("Load Patterns", "Pattern configuration format is invalid; using defaults.")
            return

        patterns = data.get("patterns", {})
        if isinstance(patterns, dict):
            for column, widget in self.pattern_texts.items():
                values = patterns.get(column)
                if isinstance(values, list):
                    widget.delete("1.0", tk.END)
                    widget.insert(
                        "1.0",
                        "\n".join(str(item) for item in values if isinstance(item, str)),
                    )

        case_flags = data.get("case_insensitive", {})
        if isinstance(case_flags, dict):
            for column, var in self.case_insensitive_vars.items():
                if column in case_flags:
                    var.set(bool(case_flags[column]))

        whitespace_flags = data.get("space_as_whitespace", {})
        if isinstance(whitespace_flags, dict):
            for column, var in self.whitespace_as_space_vars.items():
                if column in whitespace_flags:
                    var.set(bool(whitespace_flags[column]))

        year_patterns = data.get("year_patterns")
        if isinstance(year_patterns, list) and self.year_pattern_text is not None:
            self.year_pattern_text.delete("1.0", tk.END)
            self.year_pattern_text.insert(
                "1.0", "\n".join(str(item) for item in year_patterns if isinstance(item, str))
            )

        if "year_case_insensitive" in data:
            self.year_case_insensitive_var.set(bool(data["year_case_insensitive"]))
        if "year_space_as_whitespace" in data:
            self.year_whitespace_as_space_var.set(bool(data["year_space_as_whitespace"]))

        downloads_minutes = data.get("downloads_minutes")
        if isinstance(downloads_minutes, int) and downloads_minutes > 0:
            self.recent_download_minutes.set(downloads_minutes)

        models = data.get("openai_models")
        if isinstance(models, dict):
            for column, var in self.openai_model_vars.items():
                model_name = models.get(column)
                if isinstance(model_name, str) and model_name.strip():
                    var.set(model_name.strip())
        modes = data.get("upload_modes")
        if isinstance(modes, dict):
            for column, var in self.scrape_upload_mode_vars.items():
                mode_value = modes.get(column)
                if isinstance(mode_value, str) and mode_value in {"pdf", "text"}:
                    var.set(mode_value)

    def _load_local_config(self) -> None:
        self.local_config_data = {}
        if not self.local_config_path.exists():
            return
        try:
            with self.local_config_path.open("r", encoding="utf-8") as fh:
                data = json.load(fh)
        except (OSError, json.JSONDecodeError):
            return

        if not isinstance(data, dict):
            return

        self.local_config_data = data

        api_key_value = data.get("api_key")
        if isinstance(api_key_value, str):
            trimmed = api_key_value.strip()
            if trimmed:
                self.api_key_var.set(trimmed)

        downloads_dir_value = data.get("downloads_dir")
        if isinstance(downloads_dir_value, str) and not self.downloads_dir.get().strip():
            trimmed_downloads = downloads_dir_value.strip()
            if trimmed_downloads:
                self.downloads_dir.set(trimmed_downloads)

    def _write_local_config(self) -> None:
        data = dict(self.local_config_data)
        api_key_value = self.api_key_var.get().strip()
        if api_key_value:
            data["api_key"] = api_key_value
        else:
            data.pop("api_key", None)

        if not data:
            if self.local_config_path.exists():
                try:
                    self.local_config_path.unlink()
                except OSError:
                    messagebox.showwarning(
                        "Local Config", "Unable to remove local configuration file."
                    )
                    return
        else:
            try:
                with self.local_config_path.open("w", encoding="utf-8") as fh:
                    json.dump(data, fh, indent=2)
            except OSError:
                messagebox.showwarning(
                    "Local Config", "Unable to save local configuration file."
                )
                return

        self.local_config_data = data

    def _persist_api_key(self, value: str) -> None:
        trimmed = value.strip()
        if trimmed:
            if self.local_config_data.get("api_key") == trimmed:
                return
            self.local_config_data["api_key"] = trimmed
        elif "api_key" in self.local_config_data:
            self.local_config_data.pop("api_key", None)
        else:
            return
        self._write_local_config()

    def _on_api_key_var_changed(self, *_: Any) -> None:
        if self._suspend_api_key_save:
            return
        if self._api_key_save_after is not None:
            try:
                self.root.after_cancel(self._api_key_save_after)
            except Exception:
                pass
        self._api_key_save_after = self.root.after(600, self._flush_api_key_save)

    def _flush_api_key_save(self) -> None:
        self._api_key_save_after = None
        self._persist_api_key(self.api_key_var.get())

    def _load_config(self) -> None:
        if not self.config_path.exists():
            return
        try:
            with self.config_path.open("r", encoding="utf-8") as fh:
                data = json.load(fh)
        except (OSError, json.JSONDecodeError):
            return

        downloads = data.get("downloads_dir")
        if isinstance(downloads, str):
            self.downloads_dir.set(downloads)

        last_company = data.get("last_company")
        if isinstance(last_company, str):
            self.company_var.set(last_company)

    def _save_config(self) -> None:
        data = {
            "downloads_dir": self.downloads_dir.get().strip(),
            "last_company": self.company_var.get().strip(),
        }
        try:
            with self.config_path.open("w", encoding="utf-8") as fh:
                json.dump(data, fh, indent=2)
        except OSError:
            messagebox.showwarning("Save Config", "Unable to save configuration to disk.")

    # ------------------------------------------------------------------ Companies
    def _refresh_company_options(self) -> None:
        if not self.companies_dir.exists():
            self.companies_dir.mkdir(parents=True, exist_ok=True)

        companies = sorted([p.name for p in self.companies_dir.iterdir() if p.is_dir()])
        self.company_combo.configure(values=companies)

        current = self.company_var.get()
        if current and current in companies:
            self.company_combo.set(current)
            self._set_folder_for_company(current)
        elif companies:
            first = companies[0]
            self.company_combo.set(first)
            self.company_var.set(first)
            self._set_folder_for_company(first)

    def _on_company_selected(self, _: tk.Event) -> None:  # type: ignore[override]
        name = self.company_var.get()
        if name:
            self._set_folder_for_company(name)
            self._save_config()

    def _set_folder_for_company(self, company: str) -> None:
        folder = self.companies_dir / company / "raw"
        self.folder_path.set(str(folder))
        self._load_assigned_pages(company)
        self.clear_entries()

    def _load_assigned_pages(self, company: str) -> None:
        self.assigned_pages = {}
        self.assigned_pages_path = self.companies_dir / company / "assigned.json"
        if not self.assigned_pages_path.exists():
            return
        try:
            with self.assigned_pages_path.open("r", encoding="utf-8") as fh:
                data = json.load(fh)
        except Exception:
            return

        if not isinstance(data, dict):
            return

        parsed: Dict[str, Dict[str, Any]] = {}
        for pdf_name, value in data.items():
            if not isinstance(pdf_name, str) or not isinstance(value, dict):
                continue

            record: Dict[str, Any] = {}
            selections_obj = value.get("selections") if "selections" in value else value
            if isinstance(selections_obj, dict):
                selections: Dict[str, int] = {}
                for category, raw_page in selections_obj.items():
                    try:
                        selections[category] = int(raw_page)
                    except (TypeError, ValueError):
                        continue
                if selections:
                    record["selections"] = selections

            multi_obj = value.get("multi_selections")
            if isinstance(multi_obj, dict):
                multi: Dict[str, List[int]] = {}
                for category, raw_list in multi_obj.items():
                    if not isinstance(raw_list, list):
                        continue
                    pages: List[int] = []
                    for raw_page in raw_list:
                        try:
                            pages.append(int(raw_page))
                        except (TypeError, ValueError):
                            continue
                    if pages:
                        multi[category] = pages
                if multi:
                    record["multi_selections"] = multi

            year_value = value.get("year")
            if isinstance(year_value, str):
                record["year"] = year_value
            elif isinstance(year_value, (int, float)):
                record["year"] = str(int(year_value))

            if record:
                parsed[pdf_name] = record

        self.assigned_pages = parsed

    # ------------------------------------------------------------------ Pattern helpers
    def _read_text_lines(self, widget: tk.Text) -> List[str]:
        text = widget.get("1.0", tk.END)
        lines: List[str] = []
        for raw in text.splitlines():
            cleaned = raw.strip()
            if cleaned:
                lines.append(cleaned)
        return lines

    def _compile_patterns(self) -> Tuple[Dict[str, List[re.Pattern[str]]], List[re.Pattern[str]]]:
        pattern_map: Dict[str, List[re.Pattern[str]]] = {}
        for column, widget in self.pattern_texts.items():
            lines = self._read_text_lines(widget)
            compiled: List[re.Pattern[str]] = []
            flags = re.IGNORECASE if self.case_insensitive_vars[column].get() else 0
            whitespace = self.whitespace_as_space_vars[column].get()
            for line in lines:
                pattern_text = line.replace(" ", r"\s+") if whitespace else line
                try:
                    compiled.append(re.compile(pattern_text, flags))
                except re.error as exc:
                    messagebox.showerror(
                        "Invalid Pattern",
                        f"Could not compile pattern '{line}' for {column}: {exc}",
                    )
                    compiled.clear()
                    break
            pattern_map[column] = compiled

        year_patterns: List[re.Pattern[str]] = []
        if self.year_pattern_text is not None:
            lines = self._read_text_lines(self.year_pattern_text)
            flags = re.IGNORECASE if self.year_case_insensitive_var.get() else 0
            whitespace = self.year_whitespace_as_space_var.get()
            for line in lines:
                pattern_text = line.replace(" ", r"\s+") if whitespace else line
                try:
                    year_patterns.append(re.compile(pattern_text, flags))
                except re.error as exc:
                    messagebox.showerror("Invalid Year Pattern", f"Could not compile '{line}': {exc}")
                    year_patterns.clear()
                    break
        self._save_pattern_config()
        return pattern_map, year_patterns

    # ------------------------------------------------------------------ PDF loading
    def clear_entries(self) -> None:
        for entry in self.pdf_entries:
            try:
                entry.doc.close()
            except Exception:
                pass
        self.pdf_entries.clear()
        self.category_rows.clear()
        for child in self.inner_frame.winfo_children():
            child.destroy()
        self.active_scrape_key = None
        self._refresh_scrape_results()
        self._clear_scrape_preview()

    def load_pdfs(self) -> None:
        folder = self.folder_path.get()
        if not folder:
            messagebox.showinfo("Select Folder", "Choose a company before loading PDFs.")
            return

        folder_path = Path(folder)
        if not folder_path.exists():
            messagebox.showerror("Folder Not Found", f"The folder '{folder}' does not exist.")
            return

        pattern_map, year_patterns = self._compile_patterns()
        if any(not patterns for patterns in pattern_map.values()):
            return

        self.clear_entries()

        pdf_paths = sorted(folder_path.rglob("*.pdf"))
        if not pdf_paths:
            messagebox.showinfo("No PDFs", "No PDF files were found in the selected folder.")
            return

        for pdf_path in pdf_paths:
            try:
                doc = fitz.open(pdf_path)
            except Exception as exc:
                messagebox.showwarning("PDF Error", f"Could not open '{pdf_path}': {exc}")
                continue

            matches: Dict[str, List[Match]] = {column: [] for column in COLUMNS}
            year_value = ""

            for page_index in range(len(doc)):
                page = doc.load_page(page_index)
                page_text = page.get_text("text")

                for column, patterns in pattern_map.items():
                    for pattern in patterns:
                        match_obj = pattern.search(page_text)
                        if match_obj:
                            matches[column].append(
                                Match(
                                    page_index=page_index,
                                    source="regex",
                                    pattern=pattern.pattern,
                                    matched_text=match_obj.group(0).strip(),
                                )
                            )
                            break

                if not year_value:
                    for pattern in year_patterns:
                        year_match = pattern.search(page_text)
                        if year_match:
                            year_value = year_match.group(1) if year_match.groups() else year_match.group(0)
                            break

            entry = PDFEntry(path=pdf_path, doc=doc, matches=matches, year=year_value)
            self._apply_existing_assignments(entry)
            self.pdf_entries.append(entry)

        self._rebuild_review_grid()
        self._save_config()

    def _apply_existing_assignments(self, entry: PDFEntry) -> None:
        record = self.assigned_pages.get(entry.path.name)
        if not isinstance(record, dict):
            return

        stored_year = record.get("year")
        if isinstance(stored_year, str) and stored_year:
            entry.year = stored_year
        elif isinstance(stored_year, (int, float)):
            entry.year = str(int(stored_year))

        selections = record.get("selections")
        if not isinstance(selections, dict):
            return

        total_pages = len(entry.doc)
        for category, raw_page in selections.items():
            try:
                page_index = int(raw_page)
            except (TypeError, ValueError):
                continue
            if page_index < 0 or page_index >= total_pages:
                continue

            matches = entry.matches.setdefault(category, [])
            selected_index: Optional[int] = None
            for idx, match in enumerate(matches):
                if match.page_index == page_index:
                    selected_index = idx
                    break
            if selected_index is None:
                manual_match = Match(page_index=page_index, source="manual")
                matches.append(manual_match)
                matches.sort(key=lambda m: m.page_index)
                try:
                    selected_index = matches.index(manual_match)
                except ValueError:
                    selected_index = None
            if selected_index is not None:
                entry.current_index[category] = selected_index
                entry.selected_pages[category] = [matches[selected_index].page_index]

        multi_map = record.get("multi_selections")
        if isinstance(multi_map, dict):
            for category, values in multi_map.items():
                if not isinstance(values, list):
                    continue
                valid_pages: List[int] = []
                matches = entry.matches.setdefault(category, [])
                for value in values:
                    try:
                        page_index = int(value)
                    except (TypeError, ValueError):
                        continue
                    if page_index < 0 or page_index >= total_pages:
                        continue
                    if all(match.page_index != page_index for match in matches):
                        matches.append(Match(page_index=page_index, source="manual"))
                    valid_pages.append(page_index)
                if matches:
                    matches.sort(key=lambda m: m.page_index)
                if valid_pages:
                    unique_sorted = sorted(dict.fromkeys(valid_pages))
                    entry.selected_pages[category] = unique_sorted
                    if unique_sorted:
                        first_page = unique_sorted[0]
                        try:
                            first_index = next(
                                idx for idx, match in enumerate(matches) if match.page_index == first_page
                            )
                        except StopIteration:
                            first_index = None
                        if first_index is not None:
                            entry.current_index[category] = first_index

    def _rebuild_review_grid(self) -> None:
        for child in self.inner_frame.winfo_children():
            child.destroy()
        self.category_rows.clear()

        if not self.pdf_entries:
            ttk.Label(self.inner_frame, text="Load PDFs to begin reviewing.").grid(
                row=0, column=0, padx=16, pady=16, sticky="nw"
            )
            return

        for row_index, entry in enumerate(self.pdf_entries):
            container = ttk.Frame(self.inner_frame, padding=8)
            container.grid(row=row_index, column=0, sticky="ew", padx=4, pady=4)
            container.columnconfigure(1, weight=1)

            info_frame = ttk.Frame(container)
            info_frame.grid(row=0, column=0, sticky="nw", padx=(0, 12))
            ttk.Label(info_frame, text=str(entry.path.name), anchor="w", width=30, wraplength=200).pack(anchor="w")
            if entry.year:
                ttk.Label(info_frame, text=f"Year: {entry.year}", foreground="#555555").pack(anchor="w", pady=(4, 0))

            types_frame = ttk.Frame(container)
            types_frame.grid(row=0, column=1, sticky="ew")
            types_frame.columnconfigure(0, weight=1)

            for idx, column in enumerate(COLUMNS):
                row = CategoryRow(types_frame, self, entry, column)
                row.frame.grid(row=idx, column=0, sticky="ew")
                if idx:
                    row.frame.grid_configure(pady=(8, 0))
                self.category_rows[(entry.path, column)] = row
                row.refresh()

        self.inner_frame.columnconfigure(0, weight=1)
        self._refresh_scrape_results()

    def _refresh_scrape_results(self) -> None:
        if not hasattr(self, "scrape_pdf_notebook"):
            return

        for panel in self.scrape_panels.values():
            panel.destroy()
        self.scrape_panels.clear()

        for child in list(self.scrape_pdf_notebook.winfo_children()):
            child.destroy()

        self.scrape_pdf_tabs.clear()
        self.scrape_pdf_category_notebooks.clear()
        self.scrape_category_tabs.clear()
        self.scrape_category_canvases.clear()
        self.scrape_category_inners.clear()
        self.scrape_category_windows.clear()
        self.scrape_category_placeholders.clear()

        if not self.pdf_entries:
            placeholder_tab = ttk.Frame(self.scrape_pdf_notebook)
            ttk.Label(
                placeholder_tab,
                text="Load PDFs and choose pages to prepare scraping results.",
                foreground="#666666",
                wraplength=360,
                justify=tk.LEFT,
            ).pack(anchor="w", padx=12, pady=12)
            self.scrape_pdf_notebook.add(placeholder_tab, text="No PDFs")
            self.scrape_pdf_notebook.tab(placeholder_tab, state="disabled")
            self._clear_scrape_preview()
            self.active_scrape_key = None
            return

        company = self.company_var.get().strip()
        entry_lookup: Dict[Path, PDFEntry] = {entry.path: entry for entry in self.pdf_entries}
        default_entry: Optional[PDFEntry] = None
        default_category: Optional[str] = None

        for entry in self.pdf_entries:
            pdf_tab = ttk.Frame(self.scrape_pdf_notebook)
            self.scrape_pdf_notebook.add(pdf_tab, text=entry.path.name)
            self.scrape_pdf_tabs[entry.path] = pdf_tab

            category_notebook = ttk.Notebook(pdf_tab)
            category_notebook.pack(fill=tk.BOTH, expand=True)
            self.scrape_pdf_category_notebooks[entry.path] = category_notebook

            for category in COLUMNS:
                tab = ttk.Frame(category_notebook)
                category_notebook.add(tab, text=category)
                canvas = tk.Canvas(tab)
                canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
                scrollbar = ttk.Scrollbar(tab, orient=tk.VERTICAL, command=canvas.yview)
                scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
                canvas.configure(yscrollcommand=scrollbar.set)
                inner = ttk.Frame(canvas)
                window = canvas.create_window((0, 0), window=inner, anchor="nw")
                inner.bind(
                    "<Configure>",
                    lambda _e, c=canvas: c.configure(scrollregion=c.bbox("all")),
                )
                canvas.bind(
                    "<Configure>",
                    lambda event, cv=canvas, win=window: cv.itemconfigure(win, width=event.width),
                )
                placeholder = ttk.Label(
                    inner,
                    text="Select pages in the Review tab to stage AIScrape jobs.",
                    foreground="#666666",
                )
                placeholder.pack(anchor="w", padx=12, pady=12)
                key = (entry.path, category)
                self.scrape_category_tabs[key] = tab
                self.scrape_category_canvases[key] = canvas
                self.scrape_category_inners[key] = inner
                self.scrape_category_windows[key] = window
                self.scrape_category_placeholders[key] = placeholder

        for entry in self.pdf_entries:
            if company:
                target_base = self.companies_dir / company / "openapiscrape" / entry.path.stem
            else:
                target_base = entry.path.parent / "openapiscrape" / entry.path.stem
            for category in COLUMNS:
                key = (entry.path, category)
                parent_inner = self.scrape_category_inners.get(key)
                if parent_inner is None:
                    continue
                pages = self._get_selected_pages(entry, category)
                csv_path = target_base / f"{category}.csv"
                multiplier_path = target_base / f"{category}_multiplier.txt"
                if not pages and not csv_path.exists() and not multiplier_path.exists():
                    continue
                placeholder = self.scrape_category_placeholders.get(key)
                if placeholder is not None:
                    placeholder.destroy()
                    self.scrape_category_placeholders[key] = None
                panel = ScrapeResultPanel(parent_inner, self, entry, category, target_base)
                panel.load_from_files()
                self.scrape_panels[key] = panel
                if panel.has_csv_data or pages:
                    if default_entry is None:
                        default_entry = entry
                        default_category = category

        if not self.scrape_panels:
            self._clear_scrape_preview()
            self.active_scrape_key = None
            return

        if default_entry is None or default_category is None:
            first_path, first_category = next(iter(self.scrape_panels.keys()))
            default_entry = entry_lookup.get(first_path)
            default_category = first_category

        if default_entry is None or default_category is None:
            return

        self.set_active_scrape_panel(default_entry, default_category)

    def _clear_scrape_preview(self) -> None:
        if not hasattr(self, "scrape_preview_label"):
            return
        self.scrape_preview_photo = None
        self.scrape_preview_label.configure(image="", text="Select a section to preview.", background="#f0f0f0")
        if hasattr(self, "scrape_preview_title_var"):
            self.scrape_preview_title_var.set("Select a section to preview.")
        if hasattr(self, "scrape_preview_page_var"):
            self.scrape_preview_page_var.set("")
        self.scrape_preview_pages = []
        self.scrape_preview_entry = None
        self.scrape_preview_category = None
        self.scrape_preview_cycle_index = 0
        self.scrape_preview_render_page = None
        self.scrape_preview_render_width = 0

    def _on_scrape_panel_clicked(self, panel: ScrapeResultPanel) -> None:
        self.set_active_scrape_panel(panel.entry, panel.category)

    def set_active_scrape_panel(self, entry: PDFEntry, category: str) -> None:
        key = (entry.path, category)
        if key not in self.scrape_panels:
            return
        self.active_scrape_key = key
        pdf_tab = self.scrape_pdf_tabs.get(entry.path)
        if pdf_tab is not None:
            try:
                self.scrape_pdf_notebook.select(pdf_tab)
            except Exception:
                pass
        category_tab = self.scrape_category_tabs.get(key)
        category_notebook = self.scrape_pdf_category_notebooks.get(entry.path)
        if category_tab is not None and category_notebook is not None:
            try:
                category_notebook.select(category_tab)
            except Exception:
                pass
        for panel_key, panel in self.scrape_panels.items():
            panel.set_active(panel_key == key)
        self._show_scrape_preview(entry, category)

    def _show_scrape_preview(self, entry: PDFEntry, category: str) -> None:
        self.scrape_preview_entry = entry
        self.scrape_preview_category = category
        self.scrape_preview_pages = self._get_selected_pages(entry, category)
        self.scrape_preview_cycle_index = 0
        title = f"{entry.path.name} – {category}"
        self.scrape_preview_title_var.set(title)
        if not self.scrape_preview_pages:
            self.scrape_preview_label.configure(
                image="",
                text="No pages selected for this category.",
                background="#f0f0f0",
            )
            self.scrape_preview_page_var.set("")
            self.scrape_preview_photo = None
            return
        self.scrape_preview_render_page = None
        self.scrape_preview_render_width = 0
        self._display_scrape_preview_page(force=True)

    def _display_scrape_preview_page(self, force: bool = False) -> None:
        if not self.scrape_preview_entry or not self.scrape_preview_pages:
            self._clear_scrape_preview()
            return
        page_count = len(self.scrape_preview_pages)
        self.scrape_preview_cycle_index %= max(page_count, 1)
        page_index = self.scrape_preview_pages[self.scrape_preview_cycle_index]
        available_width = self.scrape_preview_last_width
        if available_width <= 1:
            available_width = self.scrape_preview_label.winfo_width()
        if available_width <= 1:
            available_width = self.scrape_preview_label.winfo_reqwidth()
        if available_width <= 1:
            available_width = max(self.thumbnail_width_var.get(), 360)
        display_width = max(int(available_width) - 16, 200)
        if (
            not force
            and self.scrape_preview_render_page == page_index
            and self.scrape_preview_render_width == display_width
        ):
            photo = self.scrape_preview_photo
        else:
            photo = self.render_page(
                self.scrape_preview_entry.doc,
                page_index,
                target_width=display_width,
            )
        if photo is None:
            self.scrape_preview_label.configure(image="", text="Preview unavailable", background="#f0f0f0")
            self.scrape_preview_photo = None
            self.scrape_preview_render_page = None
        else:
            self.scrape_preview_photo = photo
            self.scrape_preview_label.configure(image=photo, text="", background="#000000")
            self.scrape_preview_render_page = page_index
            self.scrape_preview_render_width = display_width
        self.scrape_preview_page_var.set(
            f"Page {page_index + 1} ({self.scrape_preview_cycle_index + 1}/{page_count})"
        )

    def _on_scrape_preview_resize(self, event: tk.Event) -> None:  # type: ignore[override]
        try:
            width = int(getattr(event, "width", 0))
        except Exception:
            return
        if width <= 1:
            return
        if abs(width - self.scrape_preview_last_width) <= 2:
            return
        self.scrape_preview_last_width = width
        if self.scrape_preview_pages:
            self._display_scrape_preview_page(force=True)

    def _on_scrape_preview_click(self, event: tk.Event) -> None:  # type: ignore[override]
        try:
            state = int(event.state)
        except Exception:
            state = 0
        if state & CONTROL_MASK:
            self._cycle_scrape_preview()

    def _cycle_scrape_preview(self) -> None:
        if len(self.scrape_preview_pages) < 2:
            return
        self.scrape_preview_cycle_index = (self.scrape_preview_cycle_index + 1) % len(self.scrape_preview_pages)
        self._display_scrape_preview_page(force=True)

    # ------------------------------------------------------------------ Interactions
    def _bind_review_mousewheel(self, _: tk.Event) -> None:  # type: ignore[override]
        self.review_canvas.bind_all("<MouseWheel>", self._on_review_mousewheel)
        self.review_canvas.bind_all("<Button-4>", self._on_review_mousewheel)
        self.review_canvas.bind_all("<Button-5>", self._on_review_mousewheel)

    def _unbind_review_mousewheel(self, _: tk.Event) -> None:  # type: ignore[override]
        self.review_canvas.unbind_all("<MouseWheel>")
        self.review_canvas.unbind_all("<Button-4>")
        self.review_canvas.unbind_all("<Button-5>")

    def _on_review_mousewheel(self, event: tk.Event) -> None:  # type: ignore[override]
        if getattr(event, "delta", 0):
            step = -1 if event.delta > 0 else 1
        else:
            step = -1 if getattr(event, "num", 0) == 4 else 1
        self.review_canvas.yview_scroll(step, "units")

    def _on_thumbnail_scale(self, value: str) -> None:
        try:
            width = int(float(value))
        except (TypeError, ValueError):
            return
        self.thumbnail_width_var.set(width)
        for row in self.category_rows.values():
            row.set_thumbnail_width(width)

    def select_match(
        self,
        entry: PDFEntry,
        category: str,
        index: int,
        *,
        extend_selection: bool = False,
    ) -> None:
        matches = entry.matches.get(category, [])
        if not matches:
            return
        index = max(0, min(index, len(matches) - 1))
        entry.current_index[category] = index
        match = matches[index]
        page_index = match.page_index
        if extend_selection:
            pages = entry.selected_pages.setdefault(category, [])
            if page_index not in pages:
                pages.append(page_index)
                pages.sort()
        else:
            entry.selected_pages[category] = [page_index]
        row = self.category_rows.get((entry.path, category))
        if row is not None:
            row.update_selection()
        self._refresh_scrape_results()

    def manual_select(self, entry: PDFEntry, category: str) -> None:
        self.open_pdf(entry.path)
        max_pages = len(entry.doc)
        value = simpledialog.askinteger(
            "Manual Selection",
            f"Enter the page number (1-{max_pages}) for {category}:",
            parent=self.root,
            minvalue=1,
            maxvalue=max_pages,
        )
        if value is None:
            return
        page_index = value - 1
        match = Match(page_index=page_index, source="manual")
        entry.matches.setdefault(category, []).append(match)
        entry.current_index[category] = len(entry.matches[category]) - 1
        entry.matches[category].sort(key=lambda m: m.page_index)
        entry.current_index[category] = next(
            (idx for idx, m in enumerate(entry.matches[category]) if m.page_index == page_index),
            entry.current_index[category],
        )
        entry.selected_pages[category] = [page_index]
        row = self.category_rows.get((entry.path, category))
        if row is not None:
            row.refresh()
        self._refresh_scrape_results()

    def open_pdf(self, path: Path) -> None:
        self._open_with_default_app(path, "Open PDF")

    def open_file_path(self, path: Path) -> None:
        self._open_with_default_app(path, "Open File")

    def show_raw_text_dialog(self, path: Path, title: str) -> None:
        try:
            content = path.read_text(encoding="utf-8")
        except OSError:
            messagebox.showwarning("Raw Response", "Unable to read the raw response file.")
            return
        window = tk.Toplevel(self.root)
        window.title(title)
        window.transient(self.root)

        container = ttk.Frame(window, padding=8)
        container.pack(fill=tk.BOTH, expand=True)

        text_widget = tk.Text(container, wrap=tk.WORD)
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        text_widget.insert("1.0", content)
        text_widget.configure(state=tk.DISABLED)

        scrollbar = ttk.Scrollbar(container, orient=tk.VERTICAL, command=text_widget.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        text_widget.configure(yscrollcommand=scrollbar.set)

        ttk.Button(window, text="Close", command=window.destroy).pack(pady=(0, 8))

    def _open_with_default_app(self, path: Path, failure_title: str) -> None:
        try:
            if sys.platform.startswith("win"):
                os.startfile(path)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(path)])
            else:
                subprocess.Popen(["xdg-open", str(path)])
        except Exception:
            messagebox.showwarning(
                failure_title,
                f"Unable to open '{path.name}' with the default application.",
            )

    def toggle_fullscreen_preview(self, entry: PDFEntry, page_index: int) -> None:
        if self.fullscreen_preview_window is not None and self.fullscreen_preview_window.winfo_exists():
            if (
                self.fullscreen_preview_entry == entry.path
                and self.fullscreen_preview_page == page_index
            ):
                self._close_fullscreen_preview()
                return
            self._close_fullscreen_preview()
        self._open_fullscreen_preview(entry, page_index)

    def _open_fullscreen_preview(self, entry: PDFEntry, page_index: int) -> None:
        window = tk.Toplevel(self.root)
        window.title(f"{entry.path.name} - Page {page_index + 1}")
        try:
            window.attributes("-fullscreen", True)
        except tk.TclError:
            try:
                window.state("zoomed")
            except tk.TclError:
                screen_width = window.winfo_screenwidth()
                screen_height = window.winfo_screenheight()
                try:
                    window.geometry(f"{screen_width}x{screen_height}")
                except tk.TclError:
                    pass

        window.bind("<Escape>", lambda _e: self._close_fullscreen_preview())

        window.rowconfigure(0, weight=1)
        window.columnconfigure(0, weight=1)

        container = ttk.Frame(window)
        container.grid(row=0, column=0, sticky="nsew")
        container.rowconfigure(0, weight=1)
        container.columnconfigure(0, weight=1)
        container.columnconfigure(1, weight=0)

        canvas = tk.Canvas(container, background="#111111", highlightthickness=0)
        canvas.grid(row=0, column=0, sticky="nsew")
        v_scroll = ttk.Scrollbar(container, orient=tk.VERTICAL, command=canvas.yview)
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll = ttk.Scrollbar(container, orient=tk.HORIZONTAL, command=canvas.xview)
        h_scroll.grid(row=1, column=0, sticky="ew")
        container.grid_rowconfigure(1, weight=0)
        canvas.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

        inner = ttk.Frame(canvas)
        window_item = canvas.create_window((0, 0), window=inner, anchor="nw")

        def _update_scroll_region(_: tk.Event) -> None:
            bbox = canvas.bbox(window_item)
            if bbox:
                canvas.configure(scrollregion=bbox)

        inner.bind("<Configure>", _update_scroll_region)

        def _on_mousewheel(event: tk.Event) -> None:
            if getattr(event, "delta", 0):
                delta = -event.delta
                step = int(delta / 120) or (1 if delta > 0 else -1)
                canvas.yview_scroll(step * 4, "units")
            else:
                step = -1 if getattr(event, "num", 0) == 4 else 1
                canvas.yview_scroll(step * 4, "units")

        def _on_shift_mousewheel(event: tk.Event) -> None:
            if getattr(event, "delta", 0):
                delta = -event.delta
                step = int(delta / 120) or (1 if delta > 0 else -1)
                canvas.xview_scroll(step * 4, "units")
            else:
                step = -1 if getattr(event, "num", 0) == 4 else 1
                canvas.xview_scroll(step * 4, "units")

        canvas.bind("<MouseWheel>", _on_mousewheel)
        canvas.bind("<Button-4>", _on_mousewheel)
        canvas.bind("<Button-5>", _on_mousewheel)
        canvas.bind("<Shift-MouseWheel>", _on_shift_mousewheel)
        canvas.bind("<Shift-Button-4>", _on_shift_mousewheel)
        canvas.bind("<Shift-Button-5>", _on_shift_mousewheel)

        screen_height = window.winfo_screenheight()
        target_height = max(screen_height - 160, 400)
        photo = self.render_page(entry.doc, page_index, target_height=target_height)
        if photo is None:
            label = ttk.Label(inner, text="Preview unavailable", padding=24)
            label.pack(expand=True, fill=tk.BOTH)
            self.fullscreen_preview_image = None
        else:
            label = tk.Label(inner, image=photo, background="#111111")
            label.pack()
            label.bind("<Button-1>", lambda _e: self._close_fullscreen_preview())
            self.fullscreen_preview_image = photo

        self.fullscreen_preview_window = window
        self.fullscreen_preview_entry = entry.path
        self.fullscreen_preview_page = page_index

    def _close_fullscreen_preview(self) -> None:
        if self.fullscreen_preview_window is not None and self.fullscreen_preview_window.winfo_exists():
            self.fullscreen_preview_window.destroy()
        self.fullscreen_preview_window = None
        self.fullscreen_preview_image = None
        self.fullscreen_preview_entry = None
        self.fullscreen_preview_page = None

    def render_page(
        self,
        doc: fitz.Document,
        page_index: int,
        *,
        target_width: Optional[int] = None,
        target_height: Optional[int] = None,
    ) -> Optional[ImageTk.PhotoImage]:
        try:
            page = doc.load_page(page_index)
            zoom_matrix = fitz.Matrix(1.5, 1.5)
            pix = page.get_pixmap(matrix=zoom_matrix)
            mode = "RGBA" if pix.alpha else "RGB"
            image = Image.frombytes(mode, [pix.width, pix.height], pix.samples)
            if image.mode == "RGBA":
                image = image.convert("RGB")
            scale: Optional[float] = None
            if target_width and target_width > 0:
                scale = target_width / image.width
            if target_height and target_height > 0:
                height_scale = target_height / image.height
                scale = min(scale, height_scale) if scale else height_scale
            if scale and abs(scale - 1.0) > 0.01:
                new_size = (
                    max(1, int(image.width * scale)),
                    max(1, int(image.height * scale)),
                )
                image = image.resize(new_size, Image.LANCZOS)
            return ImageTk.PhotoImage(image, master=self.root)
        except Exception:
            return None

    def _export_pages_to_pdf(
        self, doc: fitz.Document, pages: List[int]
    ) -> Optional[Path]:
        if not pages:
            return None
        temp_path: Optional[Path] = None
        try:
            unique_pages = sorted(dict.fromkeys(int(page) for page in pages))
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                temp_path = Path(tmp.name)
            new_doc = fitz.open()
            try:
                for page_index in unique_pages:
                    new_doc.insert_pdf(doc, from_page=page_index, to_page=page_index)
                new_doc.save(temp_path)
            finally:
                new_doc.close()
            return temp_path
        except Exception:
            try:
                if temp_path is not None:
                    temp_path.unlink()
            except Exception:
                pass
            return None

    def _extract_pages_text(self, doc: fitz.Document, pages: List[int]) -> Optional[str]:
        if not pages:
            return None
        snippets: List[str] = []
        seen: List[int] = sorted(dict.fromkeys(int(page) for page in pages))
        for page_index in seen:
            try:
                page = doc.load_page(page_index)
                text = page.get_text("text")
            except Exception:
                logger.exception(
                    "Failed to extract text for page %s in %s", page_index + 1, getattr(doc, "name", "document")
                )
                continue
            cleaned = text.strip()
            if not cleaned:
                continue
            snippets.append(f"--- Page {page_index + 1} ---\n{cleaned}")
        combined = "\n\n".join(snippets).strip()
        return combined or None

    def _get_selected_page_index(self, entry: PDFEntry, category: str) -> Optional[int]:
        matches = entry.matches.get(category, [])
        index = entry.current_index.get(category)
        if index is None or index < 0 or index >= len(matches):
            return None
        return matches[index].page_index

    def _get_selected_pages(self, entry: PDFEntry, category: str) -> List[int]:
        pages = entry.selected_pages.get(category, [])
        if pages:
            return sorted(dict.fromkeys(int(page) for page in pages))
        page_index = self._get_selected_page_index(entry, category)
        if page_index is None:
            return []
        return [int(page_index)]

    def _get_multi_page_indexes(self, entry: PDFEntry, category: str) -> List[int]:
        return self._get_selected_pages(entry, category)

    def _write_assigned_pages(self) -> bool:
        if self.assigned_pages_path is None:
            company = self.company_var.get().strip()
            if not company:
                return False
            self.assigned_pages_path = self.companies_dir / company / "assigned.json"
        try:
            self.assigned_pages_path.parent.mkdir(parents=True, exist_ok=True)
        except OSError:
            messagebox.showwarning(
                "Commit Assignments",
                "Could not create the folder for saving assignments.",
            )
            return False
        try:
            with self.assigned_pages_path.open("w", encoding="utf-8") as fh:
                json.dump(self.assigned_pages, fh, indent=2)
        except OSError as exc:
            messagebox.showwarning("Commit Assignments", f"Could not save assignments: {exc}")
            return False
        return True

    def commit_assignments(self) -> None:
        company = self.company_var.get().strip()
        if not company:
            messagebox.showinfo("Select Company", "Please choose a company before committing assignments.")
            return
        if self.pdf_entries:
            for entry in self.pdf_entries:
                record: Dict[str, Any] = self.assigned_pages.get(entry.path.name, {})
                if not isinstance(record, dict):
                    record = {}
                if entry.year:
                    record["year"] = entry.year
                else:
                    record.pop("year", None)

                selections = record.get("selections")
                if not isinstance(selections, dict):
                    selections = {}

                multi = record.get("multi_selections")
                if not isinstance(multi, dict):
                    multi = {}

                for category in COLUMNS:
                    page_index = self._get_selected_page_index(entry, category)
                    if page_index is None:
                        selections.pop(category, None)
                    else:
                        selections[category] = int(page_index)

                    multi_pages = self._get_multi_page_indexes(entry, category)
                    if multi_pages:
                        multi[category] = [int(idx) for idx in multi_pages]
                    else:
                        multi.pop(category, None)

                if selections:
                    record["selections"] = selections
                else:
                    record.pop("selections", None)

                if multi:
                    record["multi_selections"] = multi
                else:
                    record.pop("multi_selections", None)

                if record:
                    self.assigned_pages[entry.path.name] = record
                else:
                    self.assigned_pages.pop(entry.path.name, None)

        if self._write_assigned_pages():
            messagebox.showinfo("Commit Assignments", "Assignments saved.")

    # ------------------------------------------------------------------ Scrape
    def _get_prompt_text(self, company: str, category: str) -> Optional[str]:
        candidate_paths: List[Path] = []
        if company:
            company_dir = self.companies_dir / company
            candidate_paths.extend(
                [
                    company_dir / "prompts" / f"{category}.txt",
                    company_dir / "prompt" / f"{category}.txt",
                ]
            )
        candidate_paths.append(self.prompts_dir / f"{category}.txt")
        for path in candidate_paths:
            if path.exists():
                try:
                    return path.read_text(encoding="utf-8")
                except Exception:
                    continue
        return None

    def _strip_code_fence(self, text: str) -> str:
        fence = re.search(r"```(?:[^`\n]*)\n([\s\S]*?)```", text)
        if fence:
            return fence.group(1)
        return text

    def _parse_multiplier_response(self, response: str) -> Tuple[Optional[str], List[List[str]]]:
        cleaned = self._strip_code_fence(response)
        raw_lines = [line for line in cleaned.splitlines() if line.strip()]
        multiplier: Optional[str] = None
        data_lines: List[str] = []
        for line in raw_lines:
            stripped = line.strip()
            if multiplier is None and stripped.lower().startswith("multiplier"):
                match_obj = re.search(r"([-+]?\d[\d,]*\.?\d*)", stripped)
                if match_obj:
                    multiplier = match_obj.group(1)
                continue
            data_lines.append(stripped)

        rows: List[List[str]] = []
        if data_lines:
            reader = csv.reader(io.StringIO("\n".join(data_lines)))
            try:
                for parsed in reader:
                    rows.append([cell.strip() for cell in parsed])
            except csv.Error:
                rows.extend([line.split(",") for line in data_lines])
                rows = [[cell.strip() for cell in row] for row in rows]
        return multiplier, rows

    def _call_openai_with_pdfs(
        self, api_key: str, prompt: str, pdf_paths: List[Path], model_name: str
    ) -> str:
        sanitized_key = api_key.strip()
        if not sanitized_key:
            raise ValueError("API key is required")
        if not pdf_paths:
            raise ValueError("No PDF pages available for OpenAI request")

        selected_model = model_name.strip() or DEFAULT_OPENAI_MODEL

        client = OpenAI(api_key=sanitized_key)
        file_ids: List[str] = []
        for pdf_path in pdf_paths:
            logger.info("AIScrape uploading %s", pdf_path)
            with pdf_path.open("rb") as pdf_file:
                uploaded = client.files.create(file=pdf_file, purpose="assistants")
                file_id = getattr(uploaded, "id", None)
                if not file_id:
                    raise ValueError(f"Failed to upload {pdf_path.name} to OpenAI")
                file_ids.append(str(file_id))
                logger.info("AIScrape uploaded %s as file id %s", pdf_path.name, file_id)

        user_entries: List[Dict[str, Any]] = [
            {"type": "input_text", "text": prompt},
            {
                "type": "input_text",
                "text": "Parse the attached PDFs and return the multiplier value and CSV rows.",
            },
        ]
        user_entries.extend({"type": "input_file", "file_id": fid} for fid in file_ids)

        logger.info(
            "AIScrape submitting request (model=%s, files=%s)",
            selected_model,
            file_ids,
        )
        response = client.responses.create(
            model=selected_model,
            input=
            [
                {
                    "role": "system",
                    "content": "You are a financial statement parser.",
                },
                {"role": "user", "content": user_entries},
            ],
        )
        logger.info("AIScrape response received (model=%s)", selected_model)
        return self._extract_openai_response_text(response)

    def _call_openai_with_text(
        self, api_key: str, prompt: str, text_payload: str, model_name: str
    ) -> str:
        sanitized_key = api_key.strip()
        if not sanitized_key:
            raise ValueError("API key is required")
        cleaned_text = text_payload.strip()
        if not cleaned_text:
            raise ValueError("Extracted text is empty")

        selected_model = model_name.strip() or DEFAULT_OPENAI_MODEL
        client = OpenAI(api_key=sanitized_key)

        user_entries: List[Dict[str, Any]] = [
            {"type": "input_text", "text": prompt},
            {
                "type": "input_text",
                "text": "Parse the provided text excerpt and return the multiplier value and CSV rows.",
            },
            {"type": "input_text", "text": cleaned_text},
        ]

        logger.info(
            "AIScrape submitting text request (model=%s, characters=%s)",
            selected_model,
            len(cleaned_text),
        )
        response = client.responses.create(
            model=selected_model,
            input=[
                {
                    "role": "system",
                    "content": "You are a financial statement parser.",
                },
                {"role": "user", "content": user_entries},
            ],
        )
        logger.info("AIScrape text response received (model=%s)", selected_model)
        return self._extract_openai_response_text(response)

    def _call_openai_for_job(self, job: ScrapeJob, api_key: str) -> str:
        if job.upload_mode == "text":
            if not job.text_payload:
                raise ValueError("No extracted text available for OpenAI request")
            return self._call_openai_with_text(
                api_key,
                job.prompt_text,
                job.text_payload,
                job.model_name,
            )
        if job.temp_pdf is None:
            raise ValueError("No PDF prepared for OpenAI request")
        return self._call_openai_with_pdfs(
            api_key,
            job.prompt_text,
            [job.temp_pdf],
            job.model_name,
        )

    def _extract_openai_response_text(self, response: Any) -> str:
        text_output = getattr(response, "output_text", None)
        if text_output:
            combined = str(text_output).strip()
            if combined:
                return combined

        output_items = getattr(response, "output", None)
        if output_items:
            collected: List[str] = []
            for item in output_items:
                contents = getattr(item, "content", None)
                if not contents:
                    continue
                for content in contents:
                    if getattr(content, "type", None) == "output_text":
                        collected.append(str(getattr(content, "text", "")))
            combined = "\n".join(part.strip() for part in collected if part).strip()
            if combined:
                return combined

        if hasattr(response, "choices"):
            for choice in getattr(response, "choices", []):
                message = getattr(choice, "message", None)
                content = getattr(message, "content", None)
                if isinstance(content, str) and content.strip():
                    return content.strip()

        raise ValueError("OpenAI response did not contain any text output")

    def scrape_selected_pages(self) -> None:
        if OpenAI is None:
            messagebox.showwarning(
                "OpenAI Required",
                "Install the 'openai' package to use AIScrape.",
            )
            return

        if not self.pdf_entries:
            messagebox.showinfo("AIScrape", "Load PDFs before running AIScrape.")
            return

        company = self.company_var.get().strip()
        if not company:
            messagebox.showinfo("AIScrape", "Select a company before running AIScrape.")
            return

        api_key = self.api_key_var.get().strip()
        if not api_key:
            messagebox.showwarning("AIScrape", "Enter an OpenAI API key before running AIScrape.")
            if hasattr(self, "api_key_entry"):
                self.api_key_entry.focus_set()
            return
        self._persist_api_key(api_key)

        prompts: Dict[str, str] = {}
        missing: List[str] = []
        for category in COLUMNS:
            prompt_text = self._get_prompt_text(company, category)
            if prompt_text is None:
                missing.append(category)
            else:
                prompts[category] = prompt_text

        if missing:
            messagebox.showerror(
                "AIScrape", f"Prompt files not found for: {', '.join(missing)}."
            )
            return

        self._save_pattern_config()

        scrape_root = self.companies_dir / company / "openapiscrape"
        scrape_root.mkdir(parents=True, exist_ok=True)

        jobs: List[ScrapeJob] = []
        prep_errors: List[str] = []

        for entry in self.pdf_entries:
            for category in COLUMNS:
                pages = self._get_selected_pages(entry, category)
                if not pages:
                    continue
                prompt_text = prompts.get(category)
                if not prompt_text:
                    continue
                model_var = self.openai_model_vars.get(category)
                model_name = model_var.get() if model_var is not None else DEFAULT_OPENAI_MODEL
                mode_var = self.scrape_upload_mode_vars.get(category)
                upload_mode = mode_var.get() if mode_var is not None else "pdf"
                temp_pdf: Optional[Path] = None
                text_payload: Optional[str] = None
                if upload_mode == "text":
                    text_payload = self._extract_pages_text(entry.doc, pages)
                    if not text_payload:
                        prep_errors.append(
                            f"{entry.path.name} - {category}: Unable to extract text from selected pages"
                        )
                        continue
                else:
                    temp_pdf = self._export_pages_to_pdf(entry.doc, pages)
                    if temp_pdf is None:
                        prep_errors.append(
                            f"{entry.path.name} - {category}: Unable to prepare selected pages"
                        )
                        continue
                target_dir = scrape_root / entry.path.stem
                jobs.append(
                    ScrapeJob(
                        entry=entry,
                        category=category,
                        pages=pages,
                        prompt_text=prompt_text,
                        model_name=model_name,
                        upload_mode=upload_mode,
                        target_dir=target_dir,
                        temp_pdf=temp_pdf,
                        text_payload=text_payload,
                    )
                )
                panel = self.scrape_panels.get((entry.path, category))
                if panel is not None and not panel.has_csv_data:
                    panel.mark_loading()

        if prep_errors and not jobs:
            messagebox.showerror("AIScrape", "\n".join(prep_errors))
            return
        if not jobs:
            messagebox.showinfo("AIScrape", "Select pages before running AIScrape.")
            return

        self.scrape_button.configure(state="disabled")
        self.scrape_progress.configure(value=0, maximum=len(jobs))

        thread = threading.Thread(
            target=self._run_scrape_jobs,
            args=(jobs, api_key, prep_errors),
            daemon=True,
        )
        self._scrape_thread = thread
        thread.start()

    def _run_scrape_jobs(
        self,
        jobs: List[ScrapeJob],
        api_key: str,
        prep_errors: List[str],
    ) -> None:
        errors: List[str] = list(prep_errors)
        total = len(jobs)
        for index, job in enumerate(jobs, start=1):
            multiplier: Optional[str] = None
            success = False
            try:
                logger.info(
                    "AIScrape starting for %s | %s | pages=%s | model=%s",
                    job.entry.path.name,
                    job.category,
                    job.pages,
                    job.model_name,
                )
                logger.info("AIScrape mode=%s", job.upload_mode)
                response_text = self._call_openai_for_job(job, api_key)
                multiplier, rows = self._parse_multiplier_response(response_text)
                logger.info(
                    "AIScrape completed for %s | %s | multiplier=%s | rows=%s",
                    job.entry.path.name,
                    job.category,
                    multiplier,
                    len(rows),
                )

                job.target_dir.mkdir(parents=True, exist_ok=True)
                raw_path = job.target_dir / f"{job.category}_raw.txt"
                raw_path.write_text(response_text, encoding="utf-8")
                if multiplier is not None:
                    multiplier_path = job.target_dir / f"{job.category}_multiplier.txt"
                    multiplier_path.write_text(str(multiplier).strip(), encoding="utf-8")

                csv_path = job.target_dir / f"{job.category}.csv"
                if rows:
                    with csv_path.open("w", encoding="utf-8", newline="") as fh:
                        writer = csv.writer(fh, quoting=csv.QUOTE_MINIMAL)
                        writer.writerows(rows)
                else:
                    csv_path.write_text("", encoding="utf-8")
                success = True
            except Exception as exc:
                logger.exception(
                    "AIScrape failed for %s | %s", job.entry.path.name, job.category
                )
                errors.append(f"{job.entry.path.name} - {job.category}: {exc}")
            finally:
                try:
                    if job.temp_pdf is not None and job.temp_pdf.exists():
                        job.temp_pdf.unlink()
                except Exception:
                    pass
                self.root.after(
                    0,
                    self._on_scrape_job_progress,
                    job,
                    index,
                    success,
                    multiplier,
                )
        self.root.after(0, self._on_scrape_jobs_finished, total, errors)

    def _on_scrape_job_progress(
        self,
        job: ScrapeJob,
        completed: int,
        _success: bool,
        _multiplier: Optional[str],
    ) -> None:
        self.scrape_progress.configure(value=completed)
        panel = self.scrape_panels.get((job.entry.path, job.category))
        if panel is not None:
            panel.load_from_files()
            panel.set_active(self.active_scrape_key == (job.entry.path, job.category))
        if self.active_scrape_key == (job.entry.path, job.category):
            if self.scrape_preview_pages:
                self._display_scrape_preview_page(force=True)
            else:
                self._show_scrape_preview(job.entry, job.category)

    def _on_scrape_jobs_finished(self, total: int, errors: List[str]) -> None:
        self.scrape_button.configure(state="normal")
        self.scrape_progress.configure(value=0)
        self._scrape_thread = None
        if errors:
            messagebox.showerror("AIScrape", "\n".join(errors))
        else:
            messagebox.showinfo(
                "AIScrape",
                f"Saved {total} OpenAI response(s) to 'openapiscrape'.",
            )

    # ------------------------------------------------------------------ Company creation
    def _set_downloads_dir(self) -> None:
        initial = self.downloads_dir.get() or str(Path.home())
        selected = filedialog.askdirectory(parent=self.root, initialdir=initial, title="Select Downloads Directory")
        if not selected:
            return
        self.downloads_dir.set(selected)
        self._save_config()

    def create_company(self) -> None:
        recent_pdfs = self._collect_recent_downloads()
        preview_window: Optional[tk.Toplevel] = None
        if recent_pdfs:
            preview_window = self._show_recent_download_previews(recent_pdfs)
        else:
            downloads_dir = self.downloads_dir.get().strip()
            if downloads_dir:
                messagebox.showinfo(
                    "Recent Downloads",
                    "No recently downloaded PDFs were found in the configured downloads folder.",
                )

        name = simpledialog.askstring("New Company", "Enter a name for the new company:", parent=self.root)
        if name is None:
            if preview_window is not None and preview_window.winfo_exists():
                preview_window.destroy()
            return

        normalized = name.strip()
        if not normalized:
            messagebox.showwarning("Invalid Name", "Company name cannot be empty.")
            if preview_window is not None and preview_window.winfo_exists():
                preview_window.destroy()
            return

        safe_name = re.sub(r"[\\/:*?\"<>|]", "_", normalized).strip().strip(".")
        if not safe_name:
            messagebox.showwarning(
                "Invalid Name",
                "Company name contains only unsupported characters. Please choose a different name.",
            )
            if preview_window is not None and preview_window.winfo_exists():
                preview_window.destroy()
            return

        if safe_name != normalized:
            messagebox.showinfo(
                "Company Name Adjusted",
                f"Using '{safe_name}' as the folder name due to unsupported characters.",
            )

        company_dir = self.companies_dir / safe_name
        raw_dir = company_dir / "raw"
        try:
            raw_dir.mkdir(parents=True, exist_ok=True)
        except Exception as exc:
            messagebox.showerror("Create Company", f"Could not create folders for '{safe_name}': {exc}")
            if preview_window is not None and preview_window.winfo_exists():
                preview_window.destroy()
            return

        moved_files = 0
        for pdf_path in recent_pdfs:
            try:
                destination = self._ensure_unique_path(raw_dir / pdf_path.name)
                shutil.move(str(pdf_path), str(destination))
                moved_files += 1
            except Exception as exc:
                messagebox.showwarning("Move PDF", f"Could not move '{pdf_path.name}': {exc}")

        self._refresh_company_options()
        self.company_combo.set(safe_name)
        self.company_var.set(safe_name)
        self._set_folder_for_company(safe_name)
        self._save_config()
        self._open_in_file_manager(raw_dir)

        if preview_window is not None and preview_window.winfo_exists():
            preview_window.destroy()
        if moved_files:
            messagebox.showinfo("Create Company", f"Moved {moved_files} PDF(s) into '{safe_name}/raw'.")

    def _collect_recent_downloads(self) -> List[Path]:
        downloads_dir = self.downloads_dir.get().strip()
        if not downloads_dir:
            messagebox.showinfo(
                "Downloads Directory",
                "Configure the downloads directory before creating a new company.",
            )
            return []

        directory = Path(downloads_dir)
        if not directory.exists():
            messagebox.showerror("Downloads Directory", f"The folder '{downloads_dir}' does not exist.")
            return []

        minutes = self.recent_download_minutes.get()
        if minutes <= 0:
            minutes = 5
            self.recent_download_minutes.set(minutes)
            self._save_pattern_config()
        cutoff_ts = (datetime.now() - timedelta(minutes=minutes)).timestamp()

        recent: List[Tuple[float, Path]] = []
        for path in directory.iterdir():
            if not path.is_file() or path.suffix.lower() != ".pdf":
                continue
            try:
                stat_info = path.stat()
            except OSError:
                continue
            if stat_info.st_mtime >= cutoff_ts:
                recent.append((stat_info.st_mtime, path))

        recent.sort(key=lambda item: item[0], reverse=True)
        return [path for _mtime, path in recent]

    def _show_recent_download_previews(self, pdf_paths: List[Path]) -> tk.Toplevel:
        window = tk.Toplevel(self.root)
        window.title("Recent Downloads Preview")
        window.transient(self.root)

        container = ttk.Frame(window, padding=8)
        container.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(container, borderwidth=0)
        scrollbar = ttk.Scrollbar(container, orient=tk.VERTICAL, command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        inner = ttk.Frame(canvas)
        window_id = canvas.create_window((0, 0), window=inner, anchor="nw")

        inner.bind("<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda event: canvas.itemconfigure(window_id, width=event.width))

        previews: List[ImageTk.PhotoImage] = []
        for pdf_path in pdf_paths:
            item_frame = ttk.Frame(inner, padding=8)
            item_frame.pack(fill=tk.X, expand=True, pady=4)
            modified_text = ""
            try:
                modified = datetime.fromtimestamp(pdf_path.stat().st_mtime)
                modified_text = modified.strftime("%Y-%m-%d %H:%M")
            except OSError:
                pass
            header = pdf_path.name if not modified_text else f"{pdf_path.name} (modified {modified_text})"
            ttk.Label(item_frame, text=header, font=("TkDefaultFont", 10, "bold")).pack(anchor="w")
            try:
                doc = fitz.open(pdf_path)
            except Exception as exc:
                ttk.Label(item_frame, text=f"Unable to open PDF: {exc}").pack(anchor="w", pady=(4, 0))
                continue
            try:
                photo = self.render_page(doc, 0, target_width=220)
            finally:
                doc.close()
            if photo is None:
                ttk.Label(item_frame, text="Preview unavailable").pack(anchor="w", pady=(4, 0))
            else:
                previews.append(photo)
                ttk.Label(item_frame, image=photo).pack(anchor="w", pady=(4, 0))

        window.preview_images = previews  # type: ignore[attr-defined]
        return window

    def _ensure_unique_path(self, target: Path) -> Path:
        if not target.exists():
            return target
        stem = target.stem
        suffix = target.suffix
        counter = 1
        while True:
            candidate = target.with_name(f"{stem}_{counter}{suffix}")
            if not candidate.exists():
                return candidate
            counter += 1

    def _open_in_file_manager(self, path: Path) -> None:
        try:
            if sys.platform.startswith("win"):
                os.startfile(path)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(path)])
            else:
                subprocess.Popen(["xdg-open", str(path)])
        except Exception:
            messagebox.showwarning("Open Folder", "Could not open the folder in the file manager.")


def main() -> None:
    logging.basicConfig(level=logging.INFO)
    root = tk.Tk()
    app = ReportAppV2(root)
    root.protocol("WM_DELETE_WINDOW", root.destroy)
    root.mainloop()


if __name__ == "__main__":
    main()
