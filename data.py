from __future__ import annotations

import csv
import threading
import io
import json
import logging
import os
import re
import shutil
import subprocess
import sys
import tempfile
import tkinter as tk
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime, timedelta
from tkinter import colorchooser, filedialog, messagebox, simpledialog, ttk
from tkinter import font as tkfont
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Set, Tuple

try:
    import fitz  # type: ignore[import-untyped]
except ImportError:  # pragma: no cover - handled via runtime warning
    fitz = None  # type: ignore[assignment]


PYMUPDF_REQUIRED_MESSAGE = (
    "PyMuPDF (import name 'fitz') is required to preview, assign, and export PDF pages.\n"
    "Install it with 'pip install PyMuPDF' and then restart thatkowfinance_data."
)

from PIL import Image, ImageTk
import webbrowser
from openai import (
    APIConnectionError,
    APIError,
    APIStatusError,
    OpenAI,
    RateLimitError,
)

from analyst import FinanceDataset, FinancePlotFrame


logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)


COLUMNS = ["Financial", "Income", "Shares"]
DEFAULT_PATTERNS = {
    "Financial": ["statement of financial position"],
    "Income": ["statement of profit or loss"],
    "Shares": ["Movements in issued capital"],
}
YEAR_DEFAULT_PATTERNS = [r"(\d{4})\s+Annual\s+Report"]
DEFAULT_NOTE_OPTIONS = ["", "asis", "excluded", "negated", "share_count"]
DEFAULT_NOTE_BACKGROUND_COLORS = {
    "": "",
    "asis": "",
    "excluded": "#ff4d4f",
    "negated": "#4da6ff",
    "share_count": "#ffb6c1",
}
DEFAULT_NOTE_LABELS = {
    "": "Clear note",
    "asis": "As Is",
    "excluded": "Excluded",
    "negated": "Negated",
    "share_count": "Share Count",
}
HEX_COLOR_RE = re.compile(r"^#(?:[0-9a-fA-F]{6})$")
DEFAULT_NOTE_KEY_BINDINGS = {
    "": "`",
    "asis": "1",
    "excluded": "2",
    "negated": "3",
    "share_count": "4",
}
SPECIAL_KEYSYM_ALIASES = {
    "`": "grave",
    " ": "space",
}

DEFAULT_OPENAI_MODEL = "gpt-4o-mini"


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
    all_matches: Dict[str, List[Match]] = field(default_factory=dict)
    current_index: Dict[str, Optional[int]] = field(default_factory=dict)
    year: str = ""

    def __post_init__(self) -> None:
        for column in COLUMNS:
            self.matches.setdefault(column, [])
            self.all_matches.setdefault(column, list(self.matches[column]))
            self.current_index.setdefault(column, 0 if self.matches[column] else None)

    @property
    def stem(self) -> str:
        return self.path.stem


@dataclass(frozen=True)
class ScrapeTask:
    entry_path: Path
    entry_name: str
    entry_year: str
    category: str
    page_indexes: List[int]
    prompt_text: str
    page_text: str
    scrape_root: Path


class CollapsibleFrame(ttk.Frame):
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


SHIFT_MASK = 0x0001
CONTROL_MASK = 0x0004


class MatchThumbnail:
    SELECTED_COLOR = "#1E90FF"
    UNSELECTED_COLOR = "#c3c3c3"
    HIGHLIGHTED_COLOR = "#FFD666"

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
        self.update_state()

    def destroy(self) -> None:
        self.container.destroy()

    def update_state(self) -> None:
        current_index = self.entry.current_index.get(self.row.category)
        highlighted = self.app.is_match_highlighted(
            self.entry, self.row.category, self.match_index
        )
        selected = current_index == self.match_index
        self._apply_state(selected, highlighted)

    def _apply_state(self, selected: bool, highlighted: bool) -> None:
        if selected:
            color = self.SELECTED_COLOR
            thickness = 3
        elif highlighted:
            color = self.HIGHLIGHTED_COLOR
            thickness = 3
        else:
            color = self.UNSELECTED_COLOR
            thickness = 1
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
        control_pressed = bool(state & CONTROL_MASK)
        shift_pressed = bool(state & SHIFT_MASK)
        if control_pressed:
            was_highlighted = self.app.is_match_highlighted(
                self.entry, self.row.category, self.match_index
            )
            if was_highlighted:
                self.app._remove_review_highlight(self.entry, self.row.category, self.match_index)
            self.app.open_thumbnail_zoom(self.entry, self.match.page_index)
            if was_highlighted:
                self.app._refresh_category_row(self.entry, self.row.category, rebuild=False)
            return "break"
        if shift_pressed:
            self.app._add_review_highlight(self.entry, self.row.category, self.match_index)
            self.app.select_match_index(self.entry, self.row.category, self.match_index)
            return "break"
        self.app._set_review_highlights(self.entry, self.row.category, [self.match_index])
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


class CategoryRow:
    def __init__(self, parent: tk.Widget, app: "ReportApp", entry: PDFEntry, category: str) -> None:
        self.app = app
        self.entry = entry
        self.category = category
        self.frame = ttk.Frame(parent, padding=(0, 4, 0, 4))
        self.frame.columnconfigure(0, weight=1)
        header = ttk.Frame(self.frame)
        header.grid(row=0, column=0, sticky="ew")
        ttk.Label(header, text=category, font=("TkDefaultFont", 10, "bold")).pack(side=tk.LEFT)
        controls = ttk.Frame(header)
        controls.pack(side=tk.RIGHT)
        ttk.Button(
            controls,
            text="Manual",
            command=lambda: self.app.manual_select(self.entry, self.category),
            width=7,
        ).pack(side=tk.RIGHT, padx=(4, 0))
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
        self.canvas.bind("<MouseWheel>", self._on_generic_mousewheel)
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
        self.app._prune_review_highlights(self.entry, self.category, len(matches))
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

    def _compute_canvas_height(self) -> int:
        return max(160, int(self.target_width * 1.2))

    def _on_inner_configure(self, _: tk.Event) -> None:  # type: ignore[override]
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event: tk.Event) -> None:  # type: ignore[override]
        self.canvas.itemconfigure(self.window, height=event.height)

    def _on_generic_mousewheel(self, event: tk.Event) -> Optional[str]:  # type: ignore[override]
        state = getattr(event, "state", 0)
        # Only translate generic mouse wheel events into horizontal scrolling when Shift is held.
        if state & SHIFT_MASK:
            self._on_mousewheel(event)
            return "break"
        return None

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


class ReportApp:
    def __init__(
        self,
        root: tk.Misc,
        *,
        embedded: bool = False,
        company_name: Optional[str] = None,
        folder_override: Optional[Path] = None,
    ) -> None:
        self.root = root
        self.embedded = embedded
        if hasattr(self.root, "title") and not embedded:
            try:
                self.root.title("thatkowfinance_data")
            except tk.TclError:
                pass
        if fitz is None:
            messagebox.showerror("PyMuPDF Required", PYMUPDF_REQUIRED_MESSAGE)
            raise RuntimeError("PyMuPDF (fitz) is not installed")
        self.folder_path = tk.StringVar(master=self.root)
        if folder_override is not None:
            self.folder_path.set(str(folder_override))
        self.company_var = tk.StringVar(master=self.root)
        if company_name:
            self.company_var.set(company_name)
        self.api_key_var = tk.StringVar(master=self.root)
        self.openai_model_var = tk.StringVar(master=self.root, value=DEFAULT_OPENAI_MODEL)
        self.thumbnail_width_var = tk.IntVar(master=self.root, value=220)
        self.review_primary_match_filter_var = tk.BooleanVar(master=self.root, value=False)
        self.pattern_texts: Dict[str, tk.Text] = {}
        self.case_insensitive_vars: Dict[str, tk.BooleanVar] = {}
        self.whitespace_as_space_vars: Dict[str, tk.BooleanVar] = {}
        self.pdf_entries: List[PDFEntry] = []
        self.pdf_entry_by_path: Dict[Path, PDFEntry] = {}
        self.review_highlighted_matches: Dict[Tuple[Path, str], Set[int]] = {}
        self.category_rows: Dict[tuple[Path, str], CategoryRow] = {}
        self.year_vars: Dict[Path, tk.StringVar] = {}
        self.year_pattern_text: Optional[tk.Text] = None
        self.year_case_insensitive_var = tk.BooleanVar(master=self.root, value=True)
        self.year_whitespace_as_space_var = tk.BooleanVar(master=self.root, value=True)
        self.app_root = Path(__file__).resolve().parent
        self.companies_dir = self.app_root / "companies"
        self.prompts_dir = self.app_root / "prompts"
        self.pattern_config_path = self.app_root / "pattern_config.json"
        self.local_config_path = self.app_root / "local_config.json"
        self.type_item_category_path = self.app_root / "type_item_category.csv"
        self.combined_order_path = self.app_root / "combined_order.csv"
        self.type_category_sort_order_path = self.combined_order_path
        self.global_note_assignments_path = self.app_root / "type_category_item_assignments.csv"
        self.config_data: Dict[str, Any] = {}
        self.local_config_data: Dict[str, Any] = {}
        self.last_company_preference: str = ""
        self._config_loaded = False
        self.assigned_pages: Dict[str, Dict[str, Any]] = {}
        self.assigned_pages_path: Optional[Path] = None
        self.scraped_images: List[ImageTk.PhotoImage] = []
        self.scraped_preview_states: Dict[tk.Widget, Dict[str, Any]] = {}
        self._thumbnail_resize_job: Optional[str] = None
        self.company_tabs: Dict[str, "ReportApp"] = {}
        self.company_frames: Dict[str, ttk.Frame] = {}
        self.downloads_dir_var = tk.StringVar(master=self.root)
        self.recent_download_minutes_var = tk.IntVar(master=self.root, value=5)
        self.combined_column_label_vars: Dict[Tuple[Path, int], tk.StringVar] = {}
        self.combined_csv_sources: Dict[Tuple[Path, str], Dict[str, Any]] = {}
        self.combined_pdf_order: List[Path] = []
        self.combined_result_tree: Optional[ttk.Treeview] = None
        self.combined_max_data_columns: int = 0
        self.combined_note_record_keys: Dict[str, Tuple[str, str, str]] = {}
        self.combined_note_editor: Optional[ttk.Combobox] = None
        self.combined_note_column_id: Optional[str] = None
        self.combined_note_editor_item: Optional[str] = None
        self.combined_ordered_columns: List[str] = []
        self.combined_all_records: List[Dict[str, str]] = []
        self.combined_record_lookup: Dict[Tuple[str, str, str], Dict[str, str]] = {}
        self.combined_column_defaults: Dict[Tuple[Path, int], str] = {}
        self.combined_labels_by_pdf: Dict[Path, List[str]] = {}
        self.combined_column_name_map: Dict[str, Tuple[Path, str]] = {}
        self.combined_column_ids: Dict[str, str] = {}
        self.combined_note_cell_tags: Dict[str, str] = {}
        self.combined_type_cell_tags: Dict[str, str] = {}
        self.combined_category_cell_tags: Dict[str, str] = {}
        self.combined_row_sources: Dict[Tuple[str, str, str], List[Path]] = {}
        self.combined_base_column_widths: Dict[str, int] = {}
        self.combined_other_column_width: Optional[int] = None
        self.combined_header_label_widgets: Dict[
            Tuple[Path, int], List[Tuple[tk.Widget, str]]
        ] = {}
        self.combined_column_label_traces: Dict[Tuple[Path, int], str] = {}
        try:
            default_label_font = tkfont.nametofont("TkDefaultFont")
        except tk.TclError:
            default_label_font = tkfont.Font(root=self.root)
        self.combined_header_label_font = default_label_font
        self.combined_header_label_bold_font = default_label_font.copy()
        self.combined_header_label_bold_font.configure(weight="bold")
        # Default to iterating blank notes so reviewers immediately focus on
        # records that still need attention when the combined table loads.
        self.combined_show_blank_notes_var = tk.BooleanVar(master=self.root, value=True)
        self.combined_save_button: Optional[ttk.Button] = None
        self.combined_preview_frame: Optional[ttk.Frame] = None
        self.combined_preview_canvas: Optional[tk.Canvas] = None
        self.combined_preview_canvas_image: Optional[int] = None
        self.combined_preview_image: Optional[ImageTk.PhotoImage] = None
        self.combined_preview_zoom_var = tk.DoubleVar(value=1.0)
        self.combined_preview_zoom_display_var = tk.StringVar(value="100%")
        self.combined_preview_target: Optional[
            Tuple[PDFEntry, int, Tuple[str, str, str]]
        ] = None
        self.scraped_table_sources: Dict[ttk.Treeview, Dict[str, Any]] = {}
        self.combined_split_pane: Optional[ttk.Panedwindow] = None
        self.combined_preview_detail_var = tk.StringVar(
            master=self.root, value="Select a row to view the PDF page."
        )
        self._combined_zoom_save_job: Optional[str] = None
        self._combined_blank_notification_shown = False
        self.type_color_map: Dict[str, str] = {}
        self.type_color_labels: Dict[str, str] = {}
        self.category_color_map: Dict[str, str] = {}
        self.category_color_labels: Dict[str, str] = {}
        self.note_assignments: Dict[Tuple[str, str, str], str] = {}
        self.note_assignments_path: Optional[Path] = None
        self.note_options: List[str] = list(DEFAULT_NOTE_OPTIONS)
        self.note_background_colors: Dict[str, str] = DEFAULT_NOTE_BACKGROUND_COLORS.copy()
        self.note_key_bindings: Dict[str, str] = DEFAULT_NOTE_KEY_BINDINGS.copy()
        self.note_display_labels: Dict[str, str] = DEFAULT_NOTE_LABELS.copy()

        self.chart_plot_frame: Optional[FinancePlotFrame] = None

        self._load_local_config()
        self._build_ui()
        self._load_pattern_config()
        self._apply_configured_note_key_bindings()
        self._load_note_assignments(self.company_var.get().strip())
        if not self.embedded:
            self._maximize_window()
            self.root.after(0, self._load_pdfs_on_start)

    def _build_ui(self) -> None:
        if not self.embedded:
            self._create_menus()
            top_frame = ttk.Frame(self.root, padding=8)
            top_frame.pack(fill=tk.X)
            self.company_combo = ttk.Combobox(top_frame, textvariable=self.company_var, state="readonly", width=30)
            self.company_combo.pack(side=tk.LEFT, padx=(0, 4))
            self.company_combo.bind("<<ComboboxSelected>>", self._on_company_selected)
            load_button = ttk.Button(top_frame, text="Load PDFs", command=self._open_company_tab)
            load_button.pack(side=tk.LEFT, padx=4)
            self._refresh_company_options()
            self.company_notebook = ttk.Notebook(self.root)
            self.company_notebook.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)
            self.company_notebook.bind("<<NotebookTabChanged>>", self._on_company_tab_changed)
            return
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)
        self.notebook = notebook
        review_container = ttk.Frame(notebook)
        notebook.add(review_container, text="Review")
        pattern_section = CollapsibleFrame(
            review_container, "Regex patterns (one per line)", initially_open=False
        )
        pattern_section.pack(fill=tk.X, padx=8, pady=4)
        pattern_frame = ttk.Frame(pattern_section.content, padding=8)
        pattern_frame.pack(fill=tk.X)
        for idx, column in enumerate(COLUMNS):
            column_frame = ttk.Frame(pattern_frame)
            column_frame.grid(row=0, column=idx, padx=4, sticky="nsew")
            pattern_frame.columnconfigure(idx, weight=1)
            ttk.Label(column_frame, text=column).pack(anchor="w")
            text_widget = tk.Text(column_frame, height=4, width=30)
            text_widget.pack(fill=tk.BOTH, expand=True)
            text_widget.insert("1.0", "\n".join(DEFAULT_PATTERNS[column]))
            self.pattern_texts[column] = text_widget
            var = tk.BooleanVar(master=self.root, value=True)
            self.case_insensitive_vars[column] = var
            ttk.Checkbutton(column_frame, text="Case-insensitive", variable=var).pack(anchor="w", pady=(4, 0))
            whitespace_var = tk.BooleanVar(master=self.root, value=True)
            self.whitespace_as_space_vars[column] = whitespace_var
            ttk.Checkbutton(
                column_frame,
                text="Treat spaces as any whitespace",
                variable=whitespace_var,
            ).pack(anchor="w")
        update_button = ttk.Button(pattern_section.content, text="Apply Patterns", command=self.apply_patterns)
        update_button.pack(pady=4)
        year_frame = ttk.Frame(pattern_section.content)
        year_frame.pack(fill=tk.X, pady=(8, 0))
        ttk.Label(year_frame, text="Year pattern").pack(anchor="w")
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
        size_frame = ttk.Frame(review_container, padding=(8, 0))
        size_frame.pack(fill=tk.X, padx=0, pady=(0, 4))
        ttk.Label(size_frame, text="Review").pack(side=tk.LEFT, padx=(0, 8))
        ttk.Label(size_frame, text="Thumbnail width:").pack(side=tk.LEFT)
        self.thumbnail_scale = ttk.Scale(
            size_frame,
            from_=160,
            to=420,
            orient=tk.HORIZONTAL,
            command=self._on_thumbnail_scale,
        )
        self.thumbnail_scale.set(self.thumbnail_width_var.get())
        self.thumbnail_scale.pack(side=tk.LEFT, padx=8, fill=tk.X, expand=True)
        ttk.Label(size_frame, textvariable=self.thumbnail_width_var).pack(side=tk.LEFT)
        ttk.Checkbutton(
            size_frame,
            text="Filter others by first match",
            variable=self.review_primary_match_filter_var,
            command=self._on_review_primary_match_filter_toggle,
        ).pack(side=tk.LEFT, padx=(12, 0))
        grid_container = ttk.Frame(review_container)
        grid_container.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 4))
        self.canvas = tk.Canvas(grid_container)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar = ttk.Scrollbar(grid_container, orient=tk.VERTICAL, command=self.canvas.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.bind("<Enter>", self._bind_mousewheel)
        self.canvas.bind("<Leave>", self._unbind_mousewheel)
        self.canvas.bind("<Button-4>", self._on_mousewheel)
        self.canvas.bind("<Button-5>", self._on_mousewheel)
        self.inner_frame = ttk.Frame(self.canvas)
        self.canvas_window = self.canvas.create_window((0, 0), window=self.inner_frame, anchor="nw")
        self.inner_frame.bind("<Configure>", self._on_frame_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        action_frame = ttk.Frame(review_container, padding=8)
        action_frame.pack(fill=tk.X, pady=(0, 8))
        self.commit_button = ttk.Button(action_frame, text="Commit", command=self.commit_assignments)
        self.commit_button.pack(side=tk.RIGHT)
        scraped_container = ttk.Frame(notebook)
        self.scraped_frame = scraped_container
        notebook.add(scraped_container, text="Scraped")
        scraped_controls = ttk.Frame(scraped_container, padding=8)
        scraped_controls.pack(fill=tk.X)
        ttk.Label(scraped_controls, text="API key:").pack(side=tk.LEFT)
        self.api_key_entry = ttk.Entry(scraped_controls, textvariable=self.api_key_var, width=40)
        self.api_key_entry.pack(side=tk.LEFT, padx=4)
        self.api_key_entry.bind("<Return>", self._on_api_key_enter)
        self.api_key_entry.bind("<KP_Enter>", self._on_api_key_enter)
        self.scrape_button = ttk.Button(scraped_controls, text="AIScrape", command=self.scrape_selections)
        self.scrape_button.pack(side=tk.LEFT, padx=(8, 0))
        self.scrape_progress = ttk.Progressbar(scraped_controls, orient=tk.HORIZONTAL, mode="determinate", length=200)
        self.scrape_progress.pack(side=tk.LEFT, padx=(8, 0), fill=tk.X, expand=True)
        self.delete_all_button = ttk.Button(
            scraped_controls,
            text="Delete All",
            command=self._delete_all_scraped,
        )
        self.delete_all_button.pack(side=tk.RIGHT, padx=(8, 0))
        self.scraped_canvas = tk.Canvas(scraped_container)
        self.scraped_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scraped_scrollbar = ttk.Scrollbar(scraped_container, orient=tk.VERTICAL, command=self.scraped_canvas.yview)
        scraped_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.scraped_canvas.configure(yscrollcommand=scraped_scrollbar.set)
        self.scraped_inner = ttk.Frame(self.scraped_canvas)
        self.scraped_window = self.scraped_canvas.create_window((0, 0), window=self.scraped_inner, anchor="nw")
        self.scraped_inner.bind("<Configure>", lambda _e: self.scraped_canvas.configure(scrollregion=self.scraped_canvas.bbox("all")))
        self.scraped_canvas.bind("<Configure>", lambda e: self.scraped_canvas.itemconfigure(self.scraped_window, width=e.width))

        labels_tab = ttk.Frame(notebook)
        self.combined_labels_tab = labels_tab
        notebook.add(labels_tab, text="Labels")

        labels_controls = ttk.Frame(labels_tab, padding=8)
        labels_controls.pack(fill=tk.X)
        ttk.Label(
            labels_controls,
            text="Review scraped headers, adjust the final column labels, and click Parse to build the combined table.",
        ).pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.combine_reload_button = ttk.Button(
            labels_controls,
            text="Reload",
            command=self._on_reload_combined_clicked,
        )
        self.combine_reload_button.pack(side=tk.RIGHT, padx=(0, 8))
        self.combine_confirm_button = ttk.Button(
            labels_controls,
            text="Parse",
            command=self._on_confirm_combined_clicked,
            state="disabled",
        )
        self.combine_confirm_button.pack(side=tk.RIGHT)

        labels_canvas = tk.Canvas(labels_tab)
        labels_canvas.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

        def _labels_canvas_yview(*args: Any) -> None:
            self._destroy_note_editor()
            labels_canvas.yview(*args)

        def _labels_canvas_xview(*args: Any) -> None:
            self._destroy_note_editor()
            labels_canvas.xview(*args)

        labels_vscroll = ttk.Scrollbar(labels_tab, orient=tk.VERTICAL, command=_labels_canvas_yview)
        labels_vscroll.pack(side=tk.RIGHT, fill=tk.Y)
        labels_hscroll = ttk.Scrollbar(labels_tab, orient=tk.HORIZONTAL, command=_labels_canvas_xview)
        labels_hscroll.pack(side=tk.BOTTOM, fill=tk.X)
        labels_canvas.configure(yscrollcommand=labels_vscroll.set, xscrollcommand=labels_hscroll.set)
        self.combined_labels_canvas = labels_canvas
        self.combined_labels_inner = ttk.Frame(labels_canvas)
        self.combined_labels_window = labels_canvas.create_window((0, 0), window=self.combined_labels_inner, anchor="nw")
        self.combined_labels_inner.bind(
            "<Configure>",
            lambda _e: labels_canvas.configure(scrollregion=labels_canvas.bbox("all")),
        )
        self.combined_header_frame = ttk.Frame(self.combined_labels_inner, padding=8)
        self.combined_header_frame.grid(row=0, column=0, sticky="nsew")
        self.combined_labels_inner.columnconfigure(0, weight=1)

        combined_container = ttk.Frame(notebook)
        self.combined_frame = combined_container
        notebook.add(combined_container, text="Combined")

        self.combined_result_frame = ttk.Frame(combined_container, padding=8)
        self.combined_result_frame.pack(fill=tk.BOTH, expand=True)

        chart_container = ttk.Frame(notebook)
        notebook.add(chart_container, text="Chart")
        chart_content = ttk.Frame(chart_container, padding=8)
        chart_content.pack(fill=tk.BOTH, expand=True)
        chart_frame = FinancePlotFrame(chart_content)
        chart_frame.pack(fill=tk.BOTH, expand=True)
        chart_frame.set_display_mode(FinancePlotFrame.MODE_STACKED)
        chart_frame.set_normalization_mode(FinanceDataset.NORMALIZATION_REPORTED)
        self.chart_plot_frame = chart_frame

        notebook.bind("<<NotebookTabChanged>>", self._on_primary_tab_changed)

    def _on_frame_configure(self, _: tk.Event) -> None:  # type: ignore[override]
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event: tk.Event) -> None:  # type: ignore[override]
        self.canvas.itemconfigure(self.canvas_window, width=event.width)

    def _mousewheel_sequences(self) -> Tuple[str, ...]:
        return (
            "<MouseWheel>",
            "<Button-4>",
            "<Button-5>",
            "<Shift-Button-4>",
            "<Shift-Button-5>",
        )

    def _bind_mousewheel(self, _: tk.Event) -> None:  # type: ignore[override]
        for sequence in self._mousewheel_sequences():
            self.canvas.bind_all(sequence, self._on_mousewheel)

    def _unbind_mousewheel(self, _: tk.Event) -> None:  # type: ignore[override]
        for sequence in self._mousewheel_sequences():
            self.canvas.unbind_all(sequence)

    def _on_mousewheel(self, event: tk.Event) -> None:  # type: ignore[override]
        if event.num == 4 or event.delta > 0:
            self.canvas.yview_scroll(-1, "units")
        elif event.num == 5 or event.delta < 0:
            self.canvas.yview_scroll(1, "units")

    def _event_has_control(self, event: tk.Event) -> bool:
        try:
            return bool(int(event.state) & 0x0004)
        except Exception:
            return False

    def _on_scraped_image_click(self, event: tk.Event, widget: tk.Widget) -> Optional[str]:
        state = self.scraped_preview_states.get(widget)
        if not state:
            return None
        if self._event_has_control(event):
            self._cycle_scraped_preview(widget)
            return "break"
        entry = state.get("entry")
        if not isinstance(entry, PDFEntry):
            return "break"
        page_indexes: List[int] = state.get("page_indexes", [])
        if not page_indexes:
            return "break"
        position = state.get("position", 0)
        if not isinstance(position, int) or position < 0:
            position = 0
        if position >= len(page_indexes):
            position = len(page_indexes) - 1
        page_index = int(page_indexes[position])
        state["current_page"] = page_index
        self.open_thumbnail_zoom(entry, page_index)
        info = state.get("table_info") if isinstance(state, dict) else None
        if isinstance(info, dict):
            self._update_scraped_open_button(info)
        return "break"

    def _cycle_scraped_preview(self, widget: tk.Widget, direction: int = 1) -> None:
        state = self.scraped_preview_states.get(widget)
        if not state:
            return
        page_indexes: List[int] = state.get("page_indexes", [])
        if len(page_indexes) <= 1:
            return
        position = state.get("position", 0)
        try:
            position_int = int(position)
        except (TypeError, ValueError):
            position_int = 0
        new_position = (position_int + direction) % len(page_indexes)
        entry = state.get("entry")
        if not isinstance(entry, PDFEntry):
            return
        target_width = state.get("target_width", self.thumbnail_width_var.get())
        try:
            width_int = int(target_width)
        except (TypeError, ValueError):
            width_int = self.thumbnail_width_var.get()
        page_index = int(page_indexes[new_position])
        photo = self._render_page(entry.doc, page_index, width_int)
        if photo is None:
            return
        self.scraped_images.append(photo)
        try:
            widget.configure(image=photo)
        except Exception:
            return
        setattr(widget, "image", photo)
        state["position"] = new_position
        state["current_page"] = page_index
        title_label = state.get("title_label")
        base_text = state.get("title_base_text")
        if isinstance(title_label, ttk.Label) and isinstance(base_text, str):
            title_label.configure(text=f"{base_text} (Page {page_index + 1})")
        info = state.get("table_info") if isinstance(state, dict) else None
        if isinstance(info, dict):
            self._update_scraped_open_button(info)

    def _on_scraped_image_frame_resize(
        self, event: tk.Event, widget: tk.Widget
    ) -> None:
        state = self.scraped_preview_states.get(widget)
        if not state:
            return
        width = getattr(event, "width", 0)
        if width <= 0:
            return
        horizontal_padding = 0
        try:
            padding_value = event.widget.cget("padding")
        except (tk.TclError, AttributeError):
            padding_value = 0
        if isinstance(padding_value, str):
            parts = padding_value.split()
            parsed: List[float] = []
            for part in parts:
                try:
                    parsed.append(float(part))
                except ValueError:
                    continue
            if len(parsed) == 1:
                horizontal_padding = int(parsed[0] * 2)
            elif len(parsed) == 2:
                horizontal_padding = int(parsed[0] + parsed[1])
            elif len(parsed) >= 4:
                horizontal_padding = int(parsed[0] + parsed[2])
        elif isinstance(padding_value, (tuple, list)):
            parsed_pad: List[float] = []
            for value in padding_value:
                try:
                    parsed_pad.append(float(value))
                except (TypeError, ValueError):
                    continue
            if len(parsed_pad) == 1:
                horizontal_padding = int(parsed_pad[0] * 2)
            elif len(parsed_pad) >= 2:
                horizontal_padding = int(parsed_pad[0] + parsed_pad[1])
        available_width = max(0, width - horizontal_padding)
        if available_width <= 0:
            return
        resize_job = state.get("resize_job")
        if resize_job is not None:
            try:
                self.root.after_cancel(resize_job)
            except tk.TclError:
                pass
        state["pending_width"] = available_width
        state["resize_job"] = self.root.after(
            120,
            lambda lbl=widget: self._apply_scraped_frame_resize(lbl),
        )

    def _apply_scraped_frame_resize(self, widget: tk.Widget) -> None:
        state = self.scraped_preview_states.get(widget)
        if not state:
            return
        state.pop("resize_job", None)
        pending_width = state.pop("pending_width", None)
        if pending_width is None:
            return
        try:
            available_width = int(pending_width)
        except (TypeError, ValueError):
            return
        if available_width <= 0:
            return
        current_target = state.get("target_width")
        try:
            current_target_int = int(current_target)
        except (TypeError, ValueError):
            current_target_int = 0
        if abs(current_target_int - available_width) < 3:
            return
        entry = state.get("entry")
        if not isinstance(entry, PDFEntry):
            return
        current_page = state.get("current_page")
        try:
            page_index = int(current_page)
        except (TypeError, ValueError):
            page_index = 0
        photo = self._render_page(entry.doc, page_index, available_width)
        if photo is None:
            return
        self.scraped_images.append(photo)
        try:
            widget.configure(image=photo)
        except Exception:
            return
        setattr(widget, "image", photo)
        state["target_width"] = available_width

    def _on_thumbnail_scale(self, value: str) -> None:
        try:
            width = int(float(value))
        except (TypeError, ValueError):
            return
        self.thumbnail_width_var.set(width)
        if self._thumbnail_resize_job is not None:
            self.root.after_cancel(self._thumbnail_resize_job)
        self._thumbnail_resize_job = self.root.after(120, self._apply_thumbnail_width)

    def _apply_thumbnail_width(self) -> None:
        self._thumbnail_resize_job = None
        width = self.thumbnail_width_var.get()
        for row in self.category_rows.values():
            row.set_thumbnail_width(width)

    def _on_review_primary_match_filter_toggle(self) -> None:
        self._apply_review_primary_match_filter()

    def _maybe_reapply_primary_match_filter(self, entry: PDFEntry) -> None:
        if not self.review_primary_match_filter_var.get():
            return
        if not self.pdf_entries:
            return
        if self.pdf_entries[0] is not entry:
            return
        self._apply_review_primary_match_filter()

    def _normalize_match_text(self, text: str) -> str:
        normalized = re.sub(r"\s+", " ", text).strip().casefold()
        return normalized

    def _find_all_match_by_page(self, entry: PDFEntry, category: str, page_index: int) -> Optional[Match]:
        matches = entry.all_matches.get(category, [])
        for match in matches:
            if match.page_index == page_index:
                return match
        return None

    def _add_manual_match(self, entry: PDFEntry, category: str, page_index: int) -> Match:
        matches = entry.matches.setdefault(category, [])
        all_matches = entry.all_matches.setdefault(category, [])
        for match in all_matches:
            if match.page_index == page_index and match.source == "manual":
                if match not in matches:
                    matches.append(match)
                return match
        match = Match(page_index=page_index, source="manual")
        all_matches.append(match)
        matches.append(match)
        return match

    def _apply_review_primary_match_filter(self) -> None:
        if not self.pdf_entries:
            return

        selected_pages: Dict[Tuple[Path, str], Optional[int]] = {}
        for entry in self.pdf_entries:
            for column in COLUMNS:
                selected_pages[(entry.path, column)] = self._get_selected_page_index(entry, column)

        filter_enabled = self.review_primary_match_filter_var.get()
        normalized_primary: Dict[str, Optional[str]] = {column: None for column in COLUMNS}
        if filter_enabled:
            primary_entry = self.pdf_entries[0]
            for column in COLUMNS:
                page_index = selected_pages.get((primary_entry.path, column))
                if page_index is None:
                    continue
                match = self._find_all_match_by_page(primary_entry, column, page_index)
                if match and match.matched_text:
                    normalized = self._normalize_match_text(match.matched_text)
                    if normalized:
                        normalized_primary[column] = normalized

        for idx, entry in enumerate(self.pdf_entries):
            for column in COLUMNS:
                base_matches = list(entry.all_matches.get(column, []))
                if filter_enabled and idx != 0:
                    target_text = normalized_primary.get(column)
                    if target_text:
                        filtered = [
                            match
                            for match in base_matches
                            if match.source == "manual"
                            or (
                                match.matched_text
                                and self._normalize_match_text(match.matched_text) == target_text
                            )
                        ]
                    else:
                        filtered = base_matches
                else:
                    filtered = base_matches
                entry.matches[column] = filtered

        for entry in self.pdf_entries:
            for column in COLUMNS:
                matches = entry.matches.get(column, [])
                selected_page = selected_pages.get((entry.path, column))
                if matches:
                    if selected_page is not None:
                        for idx, match in enumerate(matches):
                            if match.page_index == selected_page:
                                entry.current_index[column] = idx
                                break
                        else:
                            entry.current_index[column] = 0
                    else:
                        entry.current_index[column] = 0
                else:
                    entry.current_index[column] = None
                self._refresh_category_row(entry, column, rebuild=True)

    def _refresh_company_options(self) -> None:
        if self.embedded or not hasattr(self, "company_combo"):
            return
        if not self.companies_dir.exists():
            self.company_combo.configure(values=[])
            return
        companies = sorted([d.name for d in self.companies_dir.iterdir() if d.is_dir()])
        self.company_combo.configure(values=companies)
        preferred = self.last_company_preference or self.company_var.get()
        if preferred and preferred in companies:
            self.company_combo.set(preferred)
            self.company_var.set(preferred)
            self._set_folder_from_company(preferred)
        elif companies and not self.company_var.get():
            self.company_combo.current(0)
            self._set_folder_from_company(companies[0])

    def _open_company_tab(self) -> None:
        if self.embedded:
            return
        company = self.company_var.get()
        if not company:
            messagebox.showinfo("Select Company", "Please choose a company before loading PDFs.")
            return
        folder = self.companies_dir / company / "raw"
        self.folder_path.set(str(folder))
        frame = self.company_frames.get(company)
        if frame is None:
            frame = ttk.Frame(self.company_notebook)
            frame.pack(fill=tk.BOTH, expand=True)
            child = ReportApp(
                frame,
                embedded=True,
                company_name=company,
                folder_override=folder,
            )
            self.company_frames[company] = frame
            self.company_tabs[company] = child
            self.company_notebook.add(frame, text=company)
            self.company_notebook.select(frame)
            self.root.after(0, child.load_pdfs)
        else:
            self.company_notebook.select(frame)

    def _on_company_tab_changed(self, event: tk.Event) -> None:  # type: ignore[override]
        if self.embedded:
            return
        widget = event.widget
        if not isinstance(widget, ttk.Notebook):
            return
        tab_id = widget.select()
        for name, frame in self.company_frames.items():
            if str(frame) == tab_id:
                self.company_var.set(name)
                self.folder_path.set(str(self.companies_dir / name / "raw"))
                break

    def _reload_current_company_tab(self) -> None:
        if self.embedded:
            self.load_pdfs()
            return
        if not hasattr(self, "company_notebook"):
            return
        try:
            tab_id = self.company_notebook.select()
        except tk.TclError:
            tab_id = ""
        if not tab_id:
            messagebox.showinfo("Reload PDFs", "Open a company tab before reloading PDFs.")
            return
        for name, frame in self.company_frames.items():
            if str(frame) == tab_id:
                child = self.company_tabs.get(name)
                if child is None:
                    break
                child.load_pdfs()
                return
        messagebox.showinfo("Reload PDFs", "Open a company tab before reloading PDFs.")

    def _on_company_selected(self, _: tk.Event) -> None:  # type: ignore[override]
        company = self.company_var.get()
        self._set_folder_from_company(company)

    def _set_folder_from_company(self, company: str) -> None:
        if not company:
            self.folder_path.set("")
            self.assigned_pages = {}
            self.assigned_pages_path = None
            self.note_assignments = {}
            self.note_assignments_path = None
            if self.embedded:
                self._reset_review_scroll()
            return
        folder = self.companies_dir / company / "raw"
        self.folder_path.set(str(folder))
        if self.embedded:
            self._load_assigned_pages(company)
            self._load_note_assignments(company)
            self._refresh_scraped_tab()
            self._reset_review_scroll()
        else:
            self._load_note_assignments(company)
        if self._config_loaded:
            self._update_last_company(company)

    def _on_api_key_enter(self, _: tk.Event) -> str:  # type: ignore[override]
        self._save_api_key()
        return "break"

    def _save_api_key(self) -> None:
        api_key = self.api_key_var.get().strip()
        if api_key:
            self.local_config_data["api_key"] = api_key
        else:
            self.local_config_data.pop("api_key", None)
        self._write_local_config()

    def _load_pdfs_on_start(self) -> None:
        if self.embedded:
            if self.folder_path.get():
                self.load_pdfs()
        else:
            # Do not automatically load a company when the application launches.
            # Users should explicitly choose a company and click the load button.
            return

    def apply_patterns(self) -> None:
        if not self.folder_path.get():
            messagebox.showinfo("Select Folder", "Please select a folder before applying patterns.")
            return
        if not self.pdf_entries:
            self.load_pdfs()
            return

        prev_patterns = {key: list(value) for key, value in self.config_data.get("patterns", {}).items()}
        prev_case = {key: bool(value) for key, value in self.config_data.get("case_insensitive", {}).items()}
        prev_whitespace = {key: bool(value) for key, value in self.config_data.get("space_as_whitespace", {}).items()}
        prev_year_patterns = list(self.config_data.get("year_patterns", YEAR_DEFAULT_PATTERNS))
        prev_year_case = bool(self.config_data.get("year_case_insensitive", True))
        prev_year_whitespace = bool(self.config_data.get("year_space_as_whitespace", True))

        pattern_map, year_patterns = self._gather_patterns()

        new_patterns = self.config_data.get("patterns", {})
        new_case = self.config_data.get("case_insensitive", {})
        new_whitespace = self.config_data.get("space_as_whitespace", {})

        changed_columns = set()
        for column in COLUMNS:
            old_patterns = prev_patterns.get(column, [])
            new_column_patterns = new_patterns.get(column, [])
            if old_patterns != new_column_patterns:
                changed_columns.add(column)
                continue
            old_case = prev_case.get(column, True)
            new_case_flag = bool(new_case.get(column, True))
            old_whitespace = prev_whitespace.get(column, True)
            new_whitespace_flag = bool(new_whitespace.get(column, True))
            if old_case != new_case_flag or old_whitespace != new_whitespace_flag:
                changed_columns.add(column)

        new_year_patterns = list(self.config_data.get("year_patterns", YEAR_DEFAULT_PATTERNS))
        new_year_case = bool(self.config_data.get("year_case_insensitive", True))
        new_year_whitespace = bool(self.config_data.get("year_space_as_whitespace", True))

        year_changed = (
            prev_year_patterns != new_year_patterns
            or prev_year_case != new_year_case
            or prev_year_whitespace != new_year_whitespace
        )

        if not changed_columns and not year_changed:
            return

        self._rescan_entries(pattern_map, year_patterns, changed_columns, year_changed)

    def _create_menus(self) -> None:
        if self.embedded:
            return
        menubar = tk.Menu(self.root)
        file_menu = tk.Menu(menubar, tearoff=False)
        file_menu.add_command(label="New Company", command=self.create_company)
        file_menu.add_command(label="Reload PDFs", command=self._reload_current_company_tab)
        file_menu.add_command(
            label="Append Type/Item/Category CSV",
            command=self._append_type_item_category_csv,
        )
        menubar.add_cascade(label="File", menu=file_menu)

        configuration_menu = tk.Menu(menubar, tearoff=False)
        configuration_menu.add_command(label="Set Downloads Dir", command=self._set_downloads_dir)
        configuration_menu.add_command(
            label="Set Download Window (minutes)",
            command=self._set_download_window,
        )
        configuration_menu.add_command(
            label="Set OpenAI Model",
            command=self._set_openai_model,
        )
        configuration_menu.add_command(
            label="Configure Combined Note Keys",
            command=self._configure_note_key_bindings,
        )
        configuration_menu.add_command(
            label="Configure Combined Column Widths",
            command=self._configure_combined_column_widths,
        )
        configuration_menu.add_command(
            label="Configure Type Colors",
            command=lambda: self._configure_value_colors("Type"),
        )
        configuration_menu.add_command(
            label="Configure Category Colors",
            command=lambda: self._configure_value_colors("Category"),
        )
        configuration_menu.add_command(
            label="Generate Type/Category Sort Order CSV",
            command=self._generate_type_category_sort_order_csv,
        )
        menubar.add_cascade(label="Configuration", menu=configuration_menu)
        try:
            self.root.configure(menu=menubar)
        except tk.TclError:
            pass

    def _set_downloads_dir(self) -> None:
        initial = self.downloads_dir_var.get() or str(Path.home())
        try:
            selected = filedialog.askdirectory(
                parent=self.root,
                initialdir=initial,
                title="Select Downloads Directory",
            )
        except tk.TclError:
            return
        if not selected:
            return
        self.downloads_dir_var.set(selected)
        self.local_config_data["downloads_dir"] = selected
        self._write_local_config()

    def _set_download_window(self) -> None:
        current_value = max(1, self.recent_download_minutes_var.get())
        value = simpledialog.askinteger(
            "Download Window",
            "Enter the number of minutes to consider downloads recent:",
            parent=self.root,
            minvalue=1,
            initialvalue=current_value,
        )
        if value is None:
            return
        self.recent_download_minutes_var.set(value)
        self.config_data["downloads_minutes"] = int(value)
        self._write_config()

    def _set_openai_model(self) -> None:
        current_value = self.openai_model_var.get().strip() or DEFAULT_OPENAI_MODEL
        try:
            value = simpledialog.askstring(
                "OpenAI Model",
                "Enter the OpenAI model name to use:",
                parent=self.root,
                initialvalue=current_value,
            )
        except tk.TclError:
            return
        if value is None:
            return
        cleaned = value.strip()
        if not cleaned:
            self.openai_model_var.set(DEFAULT_OPENAI_MODEL)
            self.config_data.pop("openai_model", None)
        else:
            self.openai_model_var.set(cleaned)
            self.config_data["openai_model"] = cleaned
        self._write_config()

    def _normalize_note_binding_value(self, value: str) -> str:
        if not value:
            return ""
        text = value.strip()
        if not text:
            return ""
        first = text[0]
        if not first.isprintable() or first.isspace():
            return ""
        return first.lower()

    def _configure_note_key_bindings(self) -> None:
        window = tk.Toplevel(self.root)
        window.title("Configure Combined Note Keys")
        window.transient(self.root)
        window.grab_set()
        container = ttk.Frame(window, padding=12)
        container.pack(fill=tk.BOTH, expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(1, weight=1)

        ttk.Label(
            container,
            text=(
                "Manage the Combined tab note shortcuts and colors."
                " Shortcuts must be a single visible key; leave blank to disable."
            ),
            wraplength=460,
            justify=tk.LEFT,
        ).grid(row=0, column=0, sticky="w")

        rows_container = ttk.Frame(container)
        rows_container.grid(row=1, column=0, sticky="nsew", pady=(12, 0))
        rows_container.columnconfigure(0, weight=1)

        header = ttk.Frame(rows_container)
        header.grid(row=0, column=0, sticky="ew")
        ttk.Label(header, text="Note value", width=24).grid(row=0, column=0, sticky="w")
        ttk.Label(header, text="Shortcut", width=10).grid(row=0, column=1, sticky="w", padx=(8, 0))
        ttk.Label(header, text="Color", width=12).grid(row=0, column=2, sticky="w", padx=(12, 0))

        label_map = {
            "": "Clear note",
            "excluded": "Set note to 'excluded'",
            "negated": "Set note to 'negated'",
            "share_count": "Set note to 'share_count'",
        }

        rows: List[Dict[str, Any]] = []

        def _update_color_preview(row: Dict[str, Any]) -> None:
            color_value = row["color_var"].get().strip()
            normalized = self._normalize_hex_color(color_value) if color_value else None
            if normalized:
                row["color_var"].set(normalized)
                row["color_text_var"].set(normalized)
                row["color_label"].configure(
                    background=normalized, foreground=self._foreground_for_color(normalized)
                )
            else:
                row["color_var"].set("")
                row["color_text_var"].set("None")
                row["color_label"].configure(background=row["default_bg"], foreground="#000000")

        def _regrid_rows() -> None:
            for index, row in enumerate(rows, start=1):
                row["frame"].grid_configure(row=index, column=0)

        def _choose_color(row: Dict[str, Any]) -> None:
            initial = row["color_var"].get() or "#ffffff"
            try:
                result = colorchooser.askcolor(initialcolor=initial, parent=window)
            except tk.TclError:
                return
            if not result or not result[1]:
                return
            row["color_var"].set(result[1])
            _update_color_preview(row)

        def _clear_color(row: Dict[str, Any]) -> None:
            row["color_var"].set("")
            _update_color_preview(row)

        def _remove_row(row: Dict[str, Any]) -> None:
            if not row["value"]:
                return
            rows.remove(row)
            row["frame"].destroy()
            _regrid_rows()

        def _create_row(
            note_value: str,
            *,
            display_name: Optional[str] = None,
            allow_remove: bool = True,
        ) -> None:
            normalized = note_value.strip().lower() if note_value else ""
            stored_label = self.note_display_labels.get(normalized)
            if stored_label is None:
                if normalized in DEFAULT_NOTE_LABELS:
                    stored_label = DEFAULT_NOTE_LABELS[normalized]
                elif normalized:
                    stored_label = normalized
                else:
                    stored_label = "Clear note"
            display = display_name or label_map.get(normalized, stored_label or (normalized or "Clear note"))
            frame = ttk.Frame(rows_container)
            frame.grid(row=len(rows) + 1, column=0, sticky="ew", pady=4)
            frame.columnconfigure(2, weight=1)
            ttk.Label(frame, text=display, width=24).grid(row=0, column=0, sticky="w")
            entry = ttk.Entry(frame, width=6)
            entry.grid(row=0, column=1, sticky="w", padx=(8, 0))
            current_shortcut = self.note_key_bindings.get(normalized, "")
            if current_shortcut:
                entry.insert(0, current_shortcut)
            color_value = self.note_background_colors.get(normalized, "")
            color_var = tk.StringVar(value=color_value)
            color_text_var = tk.StringVar(value=color_value or "None")
            color_label = tk.Label(frame, textvariable=color_text_var, width=12, relief=tk.SOLID, bd=1)
            color_label.grid(row=0, column=2, sticky="w", padx=(12, 0))
            default_bg = color_label.cget("background")
            row_data: Dict[str, Any] = {
                "value": normalized,
                "display": display,
                "label_value": display_name if display_name is not None else stored_label,
                "frame": frame,
                "entry": entry,
                "color_var": color_var,
                "color_text_var": color_text_var,
                "color_label": color_label,
                "default_bg": default_bg,
            }
            rows.append(row_data)
            _update_color_preview(row_data)
            ttk.Button(frame, text="Color…", command=lambda r=row_data: _choose_color(r)).grid(
                row=0, column=3, sticky="w", padx=(8, 0)
            )
            ttk.Button(frame, text="Clear", command=lambda r=row_data: _clear_color(r)).grid(
                row=0, column=4, sticky="w", padx=(4, 0)
            )
            if allow_remove:
                ttk.Button(frame, text="Remove", command=lambda r=row_data: _remove_row(r)).grid(
                    row=0, column=5, sticky="w", padx=(8, 0)
                )

        _create_row("", allow_remove=False)
        for note_value in self.note_options:
            if not note_value:
                continue
            _create_row(note_value)

        controls = ttk.Frame(container)
        controls.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        controls.columnconfigure(0, weight=1)

        def _on_add_entry() -> None:
            try:
                raw_value = simpledialog.askstring(
                    "Add Note Entry",
                    "Enter a new note value:",
                    parent=window,
                )
            except tk.TclError:
                return
            if raw_value is None:
                return
            candidate = raw_value.strip()
            if not candidate:
                messagebox.showerror("Add Note", "Note value cannot be blank.", parent=window)
                return
            normalized = candidate.lower()
            existing = {row["value"] for row in rows}
            if normalized in existing:
                messagebox.showerror("Add Note", "That note value already exists.", parent=window)
                return
            _create_row(normalized, display_name=candidate, allow_remove=True)
            _regrid_rows()

        ttk.Button(controls, text="Add Entry", command=_on_add_entry).grid(row=0, column=0, sticky="w")

        button_row = ttk.Frame(controls)
        button_row.grid(row=0, column=1, sticky="e")

        def _on_cancel() -> None:
            window.grab_release()
            window.destroy()

        def _on_save() -> None:
            new_order: List[str] = []
            new_bindings: Dict[str, str] = {}
            new_colors: Dict[str, str] = {}
            used_shortcuts: Set[str] = set()
            label_values: Dict[str, str] = {}
            for row in rows:
                value = row["value"]
                entry_widget: tk.Entry = row["entry"]
                raw_shortcut = entry_widget.get().strip()
                normalized_shortcut = self._normalize_note_binding_value(raw_shortcut)
                if raw_shortcut and not normalized_shortcut:
                    messagebox.showerror(
                        "Invalid Shortcut",
                        f"Shortcut for '{row['display']}' must be a single visible character.",
                        parent=window,
                    )
                    entry_widget.focus_set()
                    return
                if normalized_shortcut:
                    if normalized_shortcut in used_shortcuts:
                        messagebox.showerror(
                            "Duplicate Shortcut",
                            "Each shortcut must be unique.",
                            parent=window,
                        )
                        entry_widget.focus_set()
                        return
                    used_shortcuts.add(normalized_shortcut)
                color_value = row["color_var"].get().strip()
                normalized_color = self._normalize_hex_color(color_value) if color_value else ""
                new_order.append(value)
                new_bindings[value] = normalized_shortcut or ""
                new_colors[value] = normalized_color or ""
                label_text = row.get("label_value")
                if not label_text:
                    if value in DEFAULT_NOTE_LABELS:
                        label_text = DEFAULT_NOTE_LABELS[value]
                    elif value:
                        label_text = value
                    else:
                        label_text = "Clear note"
                label_values[value] = label_text
            if "" not in new_order:
                new_order.insert(0, "")
                new_bindings.setdefault("", "")
                new_colors.setdefault("", "")
                label_values.setdefault("", "Clear note")
            else:
                new_order = [""] + [value for value in new_order if value]
                label_values.setdefault("", "Clear note")
            sanitized_bindings = {value: new_bindings.get(value, "") for value in new_order}
            sanitized_colors = {value: new_colors.get(value, "") for value in new_order}
            sanitized_labels = {}
            for value in new_order:
                label_text = label_values.get(value)
                if not label_text:
                    if value in DEFAULT_NOTE_LABELS:
                        label_text = DEFAULT_NOTE_LABELS[value]
                    elif value:
                        label_text = value
                    else:
                        label_text = "Clear note"
                sanitized_labels[value] = label_text
            self.note_options = new_order
            self.note_key_bindings = sanitized_bindings
            self.note_background_colors = sanitized_colors
            self.note_display_labels = sanitized_labels
            valid_values = set(new_order)
            removed_assignment = False
            for key, note_value in list(self.note_assignments.items()):
                if note_value and note_value not in valid_values:
                    self.note_assignments.pop(key, None)
                    removed_assignment = True
            if removed_assignment:
                self._write_note_assignments()
            self._write_config()
            self._refresh_note_tags()
            self._update_combined_notes()
            window.grab_release()
            window.destroy()

        ttk.Button(button_row, text="Cancel", command=_on_cancel).pack(side=tk.RIGHT)
        ttk.Button(button_row, text="Save", command=_on_save).pack(side=tk.RIGHT, padx=(0, 8))

        window.bind("<Escape>", lambda _e: _on_cancel())
        window.bind("<Return>", lambda _e: _on_save())
        window.protocol("WM_DELETE_WINDOW", _on_cancel)

    def _configure_combined_column_widths(self) -> None:
        window = tk.Toplevel(self.root)
        window.title("Configure Combined Column Widths")
        window.transient(self.root)
        window.grab_set()

        container = ttk.Frame(window, padding=12)
        container.pack(fill=tk.BOTH, expand=True)
        container.columnconfigure(1, weight=1)

        ttk.Label(
            container,
            text=(
                "Set default widths for the Combined tab columns. "
                "Leave a value blank to allow the application to size it automatically."
            ),
            wraplength=420,
            justify=tk.LEFT,
        ).grid(row=0, column=0, columnspan=2, sticky="w")

        entries: Dict[str, ttk.Entry] = {}
        base_columns = ["Type", "Category", "Item", "Note"]
        tree = self.combined_result_tree

        def _current_width(column_name: str) -> Optional[int]:
            stored_width = self.combined_base_column_widths.get(column_name)
            if isinstance(stored_width, int) and stored_width > 0:
                return stored_width
            if tree is not None and column_name in self.combined_ordered_columns:
                try:
                    column_info = tree.column(column_name)
                except tk.TclError:
                    return None
                if isinstance(column_info, dict):
                    width_value = int(column_info.get("width", 0))
                    return width_value if width_value > 0 else None
            return None

        for index, column_name in enumerate(base_columns, start=1):
            ttk.Label(container, text=column_name).grid(row=index, column=0, sticky="w", pady=4)
            entry = ttk.Entry(container, width=10)
            entry.grid(row=index, column=1, sticky="ew", pady=4, padx=(8, 0))
            current_width = _current_width(column_name)
            if current_width is not None:
                entry.insert(0, str(current_width))
            entries[column_name] = entry

        other_label = "<Others>"
        other_row = len(base_columns) + 1
        ttk.Label(container, text=other_label).grid(row=other_row, column=0, sticky="w", pady=4)
        other_entry = ttk.Entry(container, width=10)
        other_entry.grid(row=other_row, column=1, sticky="ew", pady=4, padx=(8, 0))
        entries[other_label] = other_entry

        other_width = self.combined_other_column_width
        if not isinstance(other_width, int) or other_width <= 0:
            if tree is not None:
                for column_name in self.combined_ordered_columns:
                    if column_name in base_columns:
                        continue
                    try:
                        column_info = tree.column(column_name)
                    except tk.TclError:
                        continue
                    if isinstance(column_info, dict):
                        width_value = int(column_info.get("width", 0))
                        if width_value > 0:
                            other_width = width_value
                            break
        if isinstance(other_width, int) and other_width > 0:
            other_entry.insert(0, str(other_width))

        button_row = ttk.Frame(container)
        button_row.grid(row=other_row + 1, column=0, columnspan=2, pady=(12, 0), sticky="e")

        def _close() -> None:
            window.grab_release()
            window.destroy()

        def _parse_width(value: str, label: str) -> Optional[int]:
            text = value.strip()
            if not text:
                return None
            try:
                parsed = int(text)
            except ValueError:
                messagebox.showerror(
                    "Invalid Width",
                    f"{label} width must be a positive integer.",
                    parent=window,
                )
                return None
            if parsed <= 0:
                messagebox.showerror(
                    "Invalid Width",
                    f"{label} width must be greater than zero.",
                    parent=window,
                )
                return None
            return parsed

        def _on_save() -> None:
            new_widths: Dict[str, int] = {}
            for column_name in base_columns:
                value = entries[column_name].get()
                parsed = _parse_width(value, column_name)
                if parsed is None:
                    if value.strip():
                        return
                else:
                    new_widths[column_name] = parsed
            other_value = entries[other_label].get()
            parsed_other = _parse_width(other_value, "Other")
            if parsed_other is None and other_value.strip():
                return
            self.combined_base_column_widths = {
                column: width for column, width in new_widths.items()
            }
            self.combined_other_column_width = parsed_other
            self._persist_combined_base_column_widths()
            self._refresh_combined_column_widths()
            window.grab_release()
            window.destroy()

        ttk.Button(button_row, text="Cancel", command=_close).pack(side=tk.RIGHT)
        ttk.Button(button_row, text="Save", command=_on_save).pack(side=tk.RIGHT, padx=(0, 8))

        window.bind("<Escape>", lambda _e: _close())
        window.bind("<Return>", lambda _e: _on_save())

        window.protocol("WM_DELETE_WINDOW", _close)

    def _configure_value_colors(self, target: str) -> None:
        if target not in {"Type", "Category"}:
            return

        window = tk.Toplevel(self.root)
        window.title(f"Configure {target} Colors")
        window.transient(self.root)
        window.grab_set()

        container = ttk.Frame(window, padding=12)
        container.pack(fill=tk.BOTH, expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(1, weight=1)

        ttk.Label(
            container,
            text=(
                f"Assign background colors to {target} values. "
                "Cells matching these values will use the configured color when no note color overrides it."
            ),
            wraplength=460,
            justify=tk.LEFT,
        ).grid(row=0, column=0, columnspan=2, sticky="w")

        tree = ttk.Treeview(
            container,
            columns=("value", "color"),
            show="headings",
            height=8,
            selectmode="browse",
        )
        tree.heading("value", text="Value")
        tree.heading("color", text="Color")
        tree.column("value", anchor="w", width=220)
        tree.column("color", anchor="center", width=120)
        tree.grid(row=1, column=0, sticky="nsew", pady=(12, 0))
        y_scroll = ttk.Scrollbar(container, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=y_scroll.set)
        y_scroll.grid(row=1, column=1, sticky="ns", pady=(12, 0))

        if target == "Type":
            working_map = dict(self.type_color_map)
            working_labels = dict(self.type_color_labels)
        else:
            working_map = dict(self.category_color_map)
            working_labels = dict(self.category_color_labels)

        def _refresh_tree() -> None:
            for child in tree.get_children(""):
                tree.delete(child)
            entries: List[Tuple[str, str, str]] = []
            for norm_key, color in working_map.items():
                label = working_labels.get(norm_key, norm_key)
                entries.append((label, color, norm_key))
            for label, color, norm_key in sorted(entries, key=lambda item: item[0].lower()):
                tree.insert("", tk.END, iid=norm_key, values=(label, color))

        def _prompt_value(initial: str = "") -> Optional[Tuple[str, str]]:
            value = simpledialog.askstring(
                f"{target} Value",
                f"Enter the {target.lower()} value to color:",
                parent=window,
                initialvalue=initial,
            )
            if value is None:
                return None
            normalized = self._normalize_type_category_value(value)
            if not normalized:
                messagebox.showerror(
                    "Invalid Value",
                    "A non-empty value is required.",
                    parent=window,
                )
                return None
            trimmed = value.strip()
            return normalized, trimmed

        def _select_key() -> Optional[str]:
            selection = tree.selection()
            if not selection:
                return None
            return selection[0]

        def _on_add() -> None:
            result = _prompt_value()
            if result is None:
                return
            normalized, label = result
            if normalized in working_map:
                messagebox.showerror(
                    "Duplicate Value",
                    f"A color is already configured for that {target.lower()} value.",
                    parent=window,
                )
                return
            _, color_hex = colorchooser.askcolor(parent=window, title="Choose Color")
            if not color_hex:
                return
            normalized_color = self._normalize_hex_color(color_hex)
            if not normalized_color:
                messagebox.showerror(
                    "Invalid Color",
                    "Please choose a valid RGB color.",
                    parent=window,
                )
                return
            working_map[normalized] = normalized_color
            working_labels[normalized] = label
            _refresh_tree()

        def _on_edit() -> None:
            key = _select_key()
            if key is None:
                messagebox.showinfo(
                    "Edit Color",
                    f"Select a {target.lower()} value to edit first.",
                    parent=window,
                )
                return
            current_label = working_labels.get(key, key)
            result = _prompt_value(current_label)
            if result is None:
                return
            new_key, new_label = result
            if new_key != key and new_key in working_map:
                messagebox.showerror(
                    "Duplicate Value",
                    f"Another entry already uses that {target.lower()} value.",
                    parent=window,
                )
                return
            initial_color = working_map.get(key, "#ffffff")
            _, color_hex = colorchooser.askcolor(
                parent=window,
                title="Choose Color",
                initialcolor=initial_color,
            )
            if color_hex:
                normalized_color = self._normalize_hex_color(color_hex)
                if not normalized_color:
                    messagebox.showerror(
                        "Invalid Color",
                        "Please choose a valid RGB color.",
                        parent=window,
                    )
                    return
            else:
                normalized_color = working_map.get(key)
                if not normalized_color:
                    return
            if new_key != key:
                working_map.pop(key, None)
                working_labels.pop(key, None)
            working_map[new_key] = normalized_color
            working_labels[new_key] = new_label
            _refresh_tree()
            tree.selection_set(new_key)

        def _on_remove() -> None:
            key = _select_key()
            if key is None:
                messagebox.showinfo(
                    "Remove Color",
                    f"Select a {target.lower()} value to remove first.",
                    parent=window,
                )
                return
            if not messagebox.askyesno(
                "Remove Color",
                "Remove the selected color mapping?",
                parent=window,
            ):
                return
            working_map.pop(key, None)
            working_labels.pop(key, None)
            _refresh_tree()

        _refresh_tree()

        button_row = ttk.Frame(container)
        button_row.grid(row=2, column=0, columnspan=2, sticky="e", pady=(12, 0))

        ttk.Button(button_row, text="Remove", command=_on_remove).pack(side=tk.RIGHT)
        ttk.Button(button_row, text="Edit", command=_on_edit).pack(side=tk.RIGHT, padx=(0, 8))
        ttk.Button(button_row, text="Add", command=_on_add).pack(side=tk.RIGHT, padx=(0, 8))

        def _close() -> None:
            window.grab_release()
            window.destroy()

        def _on_save() -> None:
            filtered_map = {key: value for key, value in working_map.items() if value}
            filtered_labels = {key: working_labels.get(key, key) for key in filtered_map}
            if target == "Type":
                self.type_color_map = filtered_map
                self.type_color_labels = filtered_labels
            else:
                self.category_color_map = filtered_map
                self.category_color_labels = filtered_labels
            self._persist_type_category_colors()
            self._refresh_type_category_colors()
            _close()

        ttk.Button(button_row, text="Cancel", command=_close).pack(side=tk.RIGHT, padx=(0, 8))
        ttk.Button(button_row, text="Save", command=_on_save).pack(side=tk.RIGHT, padx=(0, 8))

        window.bind("<Escape>", lambda _e: _close())
        window.bind("<Return>", lambda _e: _on_save())
        window.protocol("WM_DELETE_WINDOW", _close)

    def _collect_recent_downloads(self) -> List[Path]:
        downloads_dir = self.downloads_dir_var.get().strip()
        if not downloads_dir:
            messagebox.showinfo(
                "Downloads Directory",
                "Please configure the downloads directory from Configuration → Set Downloads Dir.",
            )
            return []
        directory = Path(downloads_dir)
        if not directory.exists():
            messagebox.showerror("Downloads Directory", f"The folder '{downloads_dir}' does not exist.")
            return []

        minutes = self.recent_download_minutes_var.get()
        if minutes <= 0:
            minutes = 5
            self.recent_download_minutes_var.set(minutes)
            self.config_data["downloads_minutes"] = minutes
            self._write_config()
        cutoff_ts = (datetime.now() - timedelta(minutes=minutes)).timestamp()

        recent: List[tuple[float, Path]] = []
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

        def _on_inner_configure(_: tk.Event) -> None:
            canvas.configure(scrollregion=canvas.bbox("all"))

        inner.bind("<Configure>", _on_inner_configure)

        def _on_canvas_configure(event: tk.Event) -> None:  # type: ignore[override]
            canvas.itemconfigure(window_id, width=event.width)

        canvas.bind("<Configure>", _on_canvas_configure)

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
            except Exception as exc:  # pragma: no cover - guard for invalid PDFs
                ttk.Label(item_frame, text=f"Unable to open PDF: {exc}").pack(anchor="w", pady=(4, 0))
                continue
            try:
                photo = self._render_page(doc, 0, self.thumbnail_width_var.get())
            finally:
                doc.close()
            if photo is None:
                ttk.Label(item_frame, text="Preview unavailable").pack(anchor="w", pady=(4, 0))
            else:
                previews.append(photo)
                ttk.Label(item_frame, image=photo).pack(anchor="w", pady=(4, 0))

        window.preview_images = previews  # type: ignore[attr-defined]
        window.update_idletasks()
        try:
            window.lift()
        except tk.TclError:
            pass
        return window

    def load_pdfs(self) -> None:
        folder = self.folder_path.get()
        if not folder:
            messagebox.showinfo("Select Folder", "Please select a folder containing PDFs.")
            return

        folder_path = Path(folder)
        if not folder_path.exists():
            messagebox.showerror("Folder Not Found", f"The folder '{folder}' does not exist.")
            return

        company_name = self.company_var.get().strip()
        if company_name:
            self._load_assigned_pages(company_name)

        self._clear_entries()
        pattern_map, year_patterns = self._gather_patterns()
        pdf_paths = sorted(folder_path.rglob("*.pdf"))
        if not pdf_paths:
            messagebox.showinfo("No PDFs", "No PDF files were found in the selected folder.")
            self._rebuild_grid()
            return

        for pdf_path in pdf_paths:
            try:
                doc = fitz.open(pdf_path)
            except Exception as exc:  # pragma: no cover - guard for invalid PDFs
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
                            matched_text = match_obj.group(0).strip()
                            matches[column].append(
                                Match(
                                    page_index=page_index,
                                    source="regex",
                                    pattern=pattern.pattern,
                                    matched_text=matched_text,
                                )
                            )
                            break
                if not year_value:
                    for pattern in year_patterns:
                        year_match = pattern.search(page_text)
                        if year_match:
                            if year_match.groups():
                                year_value = year_match.group(1)
                            else:
                                year_value = year_match.group(0)
                            break

            entry = PDFEntry(path=pdf_path, doc=doc, matches=matches, year=year_value)
            # Reset current indices based on available matches
            for column in COLUMNS:
                entry.current_index[column] = 0 if entry.matches[column] else None
            self._apply_saved_selection(entry)
            self.pdf_entries.append(entry)
            self.pdf_entry_by_path[entry.path] = entry

        self._rebuild_grid()
        self._apply_review_primary_match_filter()
        self._refresh_scraped_tab()
        self._reset_review_scroll()

    def _reset_review_scroll(self) -> None:
        if not hasattr(self, "canvas"):
            return

        def _scroll() -> None:
            try:
                self.canvas.yview_moveto(0)
                self.canvas.xview_moveto(0)
            except tk.TclError:
                pass

        self.root.after_idle(_scroll)

    def create_company(self) -> None:
        if self.embedded:
            return
        preview_window: Optional[tk.Toplevel] = None
        recent_pdfs = self._collect_recent_downloads()
        if recent_pdfs:
            preview_window = self._show_recent_download_previews(recent_pdfs)
        else:
            downloads_dir = self.downloads_dir_var.get().strip()
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
        if hasattr(self, "company_combo"):
            self.company_combo.set(safe_name)
        self.company_var.set(safe_name)
        self._set_folder_from_company(safe_name)
        if self._config_loaded:
            self._update_last_company(safe_name)

        self._open_in_file_manager(raw_dir)
        self._open_company_tab()
        if preview_window is not None and preview_window.winfo_exists():
            preview_window.destroy()
        if moved_files:
            messagebox.showinfo("Create Company", f"Moved {moved_files} PDF(s) into '{safe_name}/raw'.")

    def _open_in_file_manager(self, path: Path) -> None:
        try:
            if sys.platform.startswith("win"):
                os.startfile(path)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(path)])
            else:
                subprocess.Popen(["xdg-open", str(path)])
        except Exception as exc:
            messagebox.showwarning("Open Folder", f"Could not open file browser: {exc}")

    def _load_assigned_pages(self, company: str) -> None:
        company_dir = self.companies_dir / company
        self.assigned_pages_path = company_dir / "assigned.json"
        if not self.assigned_pages_path.exists():
            self.assigned_pages = {}
            return
        try:
            with self.assigned_pages_path.open("r", encoding="utf-8") as fh:
                data = json.load(fh)
        except Exception:
            self.assigned_pages = {}
            return

        parsed: Dict[str, Dict[str, Any]] = {}
        for pdf_name, value in data.items():
            if not isinstance(value, dict):
                continue

            selections_obj = value.get("selections") if "selections" in value else value
            selections: Dict[str, int] = {}
            if isinstance(selections_obj, dict):
                for category_key, raw_page in selections_obj.items():
                    if isinstance(raw_page, (int, float)):
                        selections[category_key] = int(raw_page)

            highlights_map: Dict[str, List[int]] = {}
            raw_highlights = value.get("highlights") if isinstance(value, dict) else None
            if isinstance(raw_highlights, dict):
                for category_key, stored_pages in raw_highlights.items():
                    pages: List[int] = []
                    if isinstance(stored_pages, (list, tuple, set)):
                        iterable = stored_pages
                    else:
                        iterable = [stored_pages]
                    for page_value in iterable:
                        if isinstance(page_value, (int, float)):
                            pages.append(int(page_value))
                        else:
                            try:
                                pages.append(int(page_value))
                            except (TypeError, ValueError):
                                continue
                    if pages:
                        highlights_map[category_key] = pages

            parsed[pdf_name] = {
                "selections": selections,
                "year": value.get("year", ""),
            }
            if highlights_map:
                parsed[pdf_name]["highlights"] = highlights_map
        self.assigned_pages = parsed

    def _register_note_option(
        self,
        value: str,
        *,
        color: Optional[str] = None,
        shortcut: Optional[str] = None,
    ) -> None:
        if not isinstance(value, str):
            return
        normalized = value.strip().lower()
        if not normalized:
            return
        if normalized not in self.note_options:
            self.note_options.append(normalized)
        if normalized not in self.note_background_colors:
            if color:
                normalized_color = self._normalize_hex_color(color)
            else:
                normalized_color = DEFAULT_NOTE_BACKGROUND_COLORS.get(normalized, "")
            self.note_background_colors[normalized] = normalized_color or ""
        if normalized not in self.note_key_bindings:
            normalized_shortcut = self._normalize_note_binding_value(shortcut) if shortcut else ""
            if not normalized_shortcut:
                normalized_shortcut = DEFAULT_NOTE_KEY_BINDINGS.get(normalized, "")
            self.note_key_bindings[normalized] = normalized_shortcut or ""
        self._update_note_settings_config_entries()
        self._refresh_note_tags()

    def _read_note_assignments_file(self, path: Path) -> Dict[Tuple[str, str, str], str]:
        assignments: Dict[Tuple[str, str, str], str] = {}
        if not path.exists():
            return assignments
        try:
            with path.open("r", encoding="utf-8-sig", newline="") as fh:
                reader = csv.DictReader(fh)
                for row in reader:
                    if not isinstance(row, dict):
                        continue
                    type_value = (row.get("Type") or row.get("type") or "").strip()
                    category_value = (row.get("Category") or row.get("category") or "").strip()
                    item_value = (row.get("Item") or row.get("item") or "").strip()
                    if not (type_value and category_value and item_value):
                        continue
                    note_raw = (row.get("Note") or row.get("note") or "").strip()
                    note_value = note_raw.lower()
                    if note_value:
                        self._register_note_option(note_value)
                        assignments[(type_value, category_value, item_value)] = note_value
        except Exception:
            return {}
        return assignments

    def _load_note_assignments(self, company: str) -> None:
        path = self.global_note_assignments_path
        self.note_assignments_path = path
        self.note_assignments = self._read_note_assignments_file(path)
        self._update_combined_notes()
        self._apply_reference_sort_to_combined()

    def _ensure_note_assignments_path(self) -> Optional[Path]:
        if self.note_assignments_path is not None:
            return self.note_assignments_path
        self.note_assignments_path = self.global_note_assignments_path
        return self.note_assignments_path

    def _write_note_assignments(self) -> None:
        path = self._ensure_note_assignments_path()
        if path is None:
            return
        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            with path.open("w", encoding="utf-8", newline="") as fh:
                writer = csv.writer(fh, quoting=csv.QUOTE_ALL)
                writer.writerow(["Type", "Category", "Item", "Note"])
                for type_value, category_value, item_value in sorted(self.note_assignments.keys()):
                    note_value = self.note_assignments.get((type_value, category_value, item_value), "")
                    writer.writerow([type_value, category_value, item_value, note_value])
        except Exception as exc:
            messagebox.showwarning("Assignments", f"Could not save note assignments: {exc}")

    def _import_note_assignments_from_path(self, path: Path) -> bool:
        assignments = self._read_note_assignments_file(path)
        if not assignments and not path.exists():
            messagebox.showwarning("Assignments", f"Could not read assignments from {path}")
            return False
        self.note_assignments = assignments
        self._write_note_assignments()
        self._update_combined_notes()
        self._apply_reference_sort_to_combined()
        return True

    def _prompt_import_note_assignments(self) -> None:
        company = self.company_var.get().strip()
        if company:
            initial_dir = self.companies_dir / company
            if not initial_dir.exists():
                initial_dir = self.companies_dir
        else:
            initial_dir = self.app_root
        filename = filedialog.askopenfilename(
            title="Load type/category/item assignments",
            initialdir=str(initial_dir),
            filetypes=(("CSV files", "*.csv"), ("All files", "*.*")),
        )
        if not filename:
            return
        path = Path(filename)
        if not path.exists():
            messagebox.showwarning("Assignments", f"File not found: {path}")
            return
        self._import_note_assignments_from_path(path)

    def _destroy_note_editor(self) -> None:
        if self.combined_note_editor is not None:
            try:
                self.combined_note_editor.destroy()
            except Exception:
                pass
            self.combined_note_editor = None
        self.combined_note_editor_item = None

    @staticmethod
    def _parse_numeric_value(value: str) -> Optional[float]:
        text = value.strip()
        if not text:
            return None
        if text in {"-", "--"}:
            return None
        lower_text = text.lower()
        if lower_text in {"na", "n/a", "nil", "none"}:
            return None
        negative = False
        if text.startswith("(") and text.endswith(")"):
            negative = True
            text = text[1:-1].strip()
            if not text:
                return None
        cleaned = text.replace(",", "").strip()
        cleaned = re.sub(r"[^0-9eE\.\+\-]", "", cleaned)
        if not cleaned:
            return None
        try:
            number = float(cleaned)
        except ValueError:
            return None
        if negative and number > 0:
            number = -number
        return number

    def _format_combined_value(self, value: Any) -> str:
        if value is None:
            return ""
        if isinstance(value, (int, float)):
            return f"{float(value):.2e}"
        text = str(value).strip()
        if not text:
            return ""
        suffix = ""
        normalized = text
        if normalized.endswith("%"):
            suffix = "%"
            normalized = normalized[:-1]
        prefix = ""
        currency_symbols = "$€£¥"
        for symbol in currency_symbols:
            index = normalized.find(symbol)
            if index != -1:
                prefix = symbol
                normalized = normalized[:index] + normalized[index + 1 :]
                break
        normalized = normalized.strip()
        numeric = self._parse_numeric_value(normalized)
        if numeric is None:
            return text
        formatted = f"{numeric:.2e}"
        return f"{prefix}{formatted}{suffix}"

    def _note_tag_for_value(self, value: str) -> str:
        normalized = value.strip().lower() if isinstance(value, str) else ""
        if normalized not in self.note_background_colors:
            normalized = ""
        return f"note_value_{normalized or 'blank'}"

    def _tree_cell_tag_add(
        self, tree: ttk.Treeview, tag_name: str, item_id: str, column_id: str
    ) -> None:
        try:
            tree.tk.call(tree, "tag", "add", tag_name, item_id, column_id)
        except tk.TclError:
            pass

    def _tree_cell_tag_remove(
        self, tree: ttk.Treeview, tag_name: str, item_id: str, column_id: str
    ) -> None:
        try:
            tree.tk.call(tree, "tag", "remove", tag_name, item_id, column_id)
        except tk.TclError:
            pass

    def _foreground_for_color(self, color: str) -> str:
        normalized = self._normalize_hex_color(color)
        if not normalized:
            return "#000000"
        red = int(normalized[1:3], 16)
        green = int(normalized[3:5], 16)
        blue = int(normalized[5:7], 16)
        brightness = (0.299 * red) + (0.587 * green) + (0.114 * blue)
        return "#000000" if brightness >= 160 else "#ffffff"

    def _configure_note_tags(self, tree: ttk.Treeview) -> None:
        for note_value, color in self.note_background_colors.items():
            tag_name = self._note_tag_for_value(note_value)
            kwargs: Dict[str, Any] = {"background": ""}
            if color:
                kwargs["foreground"] = color
            else:
                kwargs["foreground"] = ""
            tree.tag_configure(tag_name, **kwargs)

    def _apply_note_value_tag(self, item_id: str, value: str) -> None:
        tree = self.combined_result_tree
        if tree is None:
            return
        column_id = self.combined_column_ids.get("Note")
        if not column_id:
            return
        previous_tag = self.combined_note_cell_tags.pop(item_id, None)
        if previous_tag:
            self._tree_cell_tag_remove(tree, previous_tag, item_id, column_id)
        tag_name = self._note_tag_for_value(value)
        if tag_name == "note_value_blank":
            return
        self._tree_cell_tag_add(tree, tag_name, item_id, column_id)
        self.combined_note_cell_tags[item_id] = tag_name

    def _normalize_type_category_value(self, value: Any) -> str:
        if isinstance(value, str):
            normalized = value.strip().lower()
        elif value is None:
            normalized = ""
        else:
            normalized = str(value).strip().lower()
        return normalized

    def _normalize_hex_color(self, color: str) -> Optional[str]:
        if not isinstance(color, str):
            return None
        candidate = color.strip()
        if not candidate:
            return None
        if not candidate.startswith("#"):
            candidate = f"#{candidate}"
        if HEX_COLOR_RE.match(candidate):
            return candidate.lower()
        return None

    def _apply_type_color_tag(self, item_id: str, type_value: Any) -> None:
        tree = self.combined_result_tree
        if tree is None:
            return
        column_id = self.combined_column_ids.get("Type")
        if not column_id:
            return
        previous_tag = self.combined_type_cell_tags.pop(item_id, None)
        if previous_tag:
            self._tree_cell_tag_remove(tree, previous_tag, item_id, column_id)
        normalized = self._normalize_type_category_value(type_value)
        color_value = ""
        if normalized and normalized in self.type_color_map:
            candidate = self.type_color_map.get(normalized, "")
            color_value = self._normalize_hex_color(candidate) or ""
        if not color_value:
            return
        tag_key = re.sub(r"[^0-9a-zA-Z]+", "_", normalized) or "value"
        tag_name = f"type_color_{tag_key}"
        tree.tag_configure(tag_name, background="", foreground=color_value)
        self._tree_cell_tag_add(tree, tag_name, item_id, column_id)
        self.combined_type_cell_tags[item_id] = tag_name

    def _apply_category_color_tag(self, item_id: str, category_value: Any) -> None:
        tree = self.combined_result_tree
        if tree is None:
            return
        column_id = self.combined_column_ids.get("Category")
        if not column_id:
            return
        previous_tag = self.combined_category_cell_tags.pop(item_id, None)
        if previous_tag:
            self._tree_cell_tag_remove(tree, previous_tag, item_id, column_id)
        normalized = self._normalize_type_category_value(category_value)
        color_value = ""
        if normalized and normalized in self.category_color_map:
            candidate = self.category_color_map.get(normalized, "")
            color_value = self._normalize_hex_color(candidate) or ""
        if not color_value:
            return
        tag_key = re.sub(r"[^0-9a-zA-Z]+", "_", normalized) or "value"
        tag_name = f"category_color_{tag_key}"
        tree.tag_configure(tag_name, background="", foreground=color_value)
        self._tree_cell_tag_add(tree, tag_name, item_id, column_id)
        self.combined_category_cell_tags[item_id] = tag_name

    def _refresh_type_category_colors(self) -> None:
        tree = self.combined_result_tree
        if tree is None or not self.combined_ordered_columns:
            return
        try:
            type_index = self.combined_ordered_columns.index("Type")
        except ValueError:
            type_index = None
        try:
            category_index = self.combined_ordered_columns.index("Category")
        except ValueError:
            category_index = None
        for item_id in tree.get_children(""):
            values = list(tree.item(item_id, "values") or [])
            type_value = values[type_index] if type_index is not None and type_index < len(values) else ""
            category_value = (
                values[category_index]
                if category_index is not None and category_index < len(values)
                else ""
            )
            self._apply_type_color_tag(item_id, type_value)
            self._apply_category_color_tag(item_id, category_value)

    def _refresh_note_tags(self) -> None:
        tree = self.combined_result_tree
        if tree is None:
            return
        self._configure_note_tags(tree)
        if not self.combined_ordered_columns or "Note" not in self.combined_ordered_columns:
            return
        note_index = self.combined_ordered_columns.index("Note")
        for item_id in tree.get_children(""):
            values = list(tree.item(item_id, "values") or [])
            note_value = values[note_index] if note_index < len(values) else ""
            self._apply_note_value_tag(item_id, note_value)
        self._destroy_note_editor()

    def _update_combined_notes(self) -> None:
        for key, record in self.combined_record_lookup.items():
            raw_note = self.note_assignments.get(key, "")
            if isinstance(raw_note, str):
                note_value = raw_note.strip().lower()
            else:
                note_value = str(raw_note or "").strip().lower()
            if note_value and note_value not in self.note_options:
                self._register_note_option(note_value)
            if note_value and note_value not in self.note_options:
                note_value = ""
            record["Note"] = note_value
        tree = self.combined_result_tree
        if tree is None:
            return
        if not self.combined_ordered_columns or "Note" not in self.combined_ordered_columns:
            return
        note_index = self.combined_ordered_columns.index("Note")
        for item_id, key in self.combined_note_record_keys.items():
            values = list(tree.item(item_id, "values"))
            if note_index >= len(values):
                continue
            raw_note = self.note_assignments.get(key, "")
            if isinstance(raw_note, str):
                note_value = raw_note.strip().lower()
            else:
                note_value = str(raw_note or "").strip().lower()
            if note_value and note_value not in self.note_options:
                self._register_note_option(note_value)
            if note_value and note_value not in self.note_options:
                note_value = ""
            if values[note_index] != note_value:
                values[note_index] = note_value
                tree.item(item_id, values=values)
            self._apply_note_value_tag(item_id, note_value)
        self._destroy_note_editor()

    def _on_combined_tree_click(self, event: tk.Event) -> None:  # type: ignore[override]
        tree = self.combined_result_tree
        if tree is None or self.combined_note_column_id is None:
            return
        try:
            tree.focus_set()
        except tk.TclError:
            pass
        region = tree.identify("region", event.x, event.y)
        if region != "cell":
            self._destroy_note_editor()
            return
        column_id = tree.identify_column(event.x)
        item_id = tree.identify_row(event.y)
        if column_id != self.combined_note_column_id or not item_id:
            self._destroy_note_editor()
            return
        self._show_note_editor(item_id)

    def _show_note_editor(self, item_id: str) -> None:
        tree = self.combined_result_tree
        if tree is None or self.combined_note_column_id is None:
            return
        bbox = tree.bbox(item_id, self.combined_note_column_id)
        if not bbox:
            self._destroy_note_editor()
            return
        x, y, width, height = bbox
        current_value = tree.set(item_id, "Note")
        self._destroy_note_editor()
        editor = ttk.Combobox(tree, values=self.note_options, state="readonly")
        editor.place(x=x, y=y, width=width, height=height)
        normalized_current = current_value.strip().lower() if isinstance(current_value, str) else ""
        if normalized_current and normalized_current not in self.note_options:
            self._register_note_option(normalized_current)
        if normalized_current in self.note_options:
            editor.set(normalized_current)
        else:
            editor.set("")
        editor.focus_set()
        editor.bind("<<ComboboxSelected>>", lambda _e, itm=item_id: self._commit_note_value(itm))
        editor.bind("<FocusOut>", lambda _e, itm=item_id: self._commit_note_value(itm))
        editor.bind("<Return>", lambda _e, itm=item_id: self._commit_note_value(itm))
        self.combined_note_editor = editor
        self.combined_note_editor_item = item_id

    def _commit_note_value(self, item_id: str) -> None:
        if self.combined_note_editor is None:
            return
        raw_value = self.combined_note_editor.get().strip().lower()
        if raw_value and raw_value not in self.note_options:
            self._register_note_option(raw_value)
        value = raw_value if raw_value in self.note_options else ""
        self._set_note_value(item_id, value)
        self._destroy_note_editor()

    def _normalized_note_binding_map(self) -> Dict[str, str]:
        mapping: Dict[str, str] = {}
        for note_value, key_text in self.note_key_bindings.items():
            if not key_text:
                continue
            normalized_key = key_text.lower()
            mapping[normalized_key] = note_value
            alias = SPECIAL_KEYSYM_ALIASES.get(key_text)
            if alias:
                mapping[alias.lower()] = note_value
        return mapping

    def _note_value_for_event(self, event: tk.Event) -> Optional[str]:
        mapping = self._normalized_note_binding_map()
        char = getattr(event, "char", "")
        if char:
            note_value = mapping.get(char.lower())
            if note_value is not None:
                return note_value
        keysym = getattr(event, "keysym", "")
        if keysym:
            return mapping.get(keysym.lower())
        return None

    def _apply_note_binding_to_selection(self, value: str) -> None:
        tree = self.combined_result_tree
        if tree is None:
            return
        selection = tree.selection()
        if not selection:
            focus_item = tree.focus()
            if focus_item:
                selection = (focus_item,)
        if not selection:
            return
        self._destroy_note_editor()
        for item_id in selection:
            self._set_note_value(item_id, value)

    def _on_combined_tree_key(self, event: tk.Event) -> Optional[str]:  # type: ignore[override]
        if event.keysym == "Tab":
            return None
        note_value = self._note_value_for_event(event)
        if note_value is None:
            return None
        self._apply_note_binding_to_selection(note_value)
        return "break"

    def _on_combined_tree_release(self, _: tk.Event) -> None:  # type: ignore[override]
        self._store_combined_base_column_widths()

    def _on_combined_tree_tab(self, event: tk.Event) -> Optional[str]:  # type: ignore[override]
        return self._move_combined_selection(1)

    def _is_blank_note_tree_item(self, tree: ttk.Treeview, item_id: str) -> bool:
        if not self.combined_ordered_columns or "Note" not in self.combined_ordered_columns:
            return False
        value = tree.set(item_id, "Note")
        if isinstance(value, str):
            normalized = value.strip()
        else:
            normalized = str(value or "").strip()
        return normalized == ""

    def _move_combined_selection(self, offset: int) -> Optional[str]:
        tree = self.combined_result_tree
        if tree is None:
            return None
        children = tree.get_children("")
        if not children:
            return "break"
        selection = tree.selection()
        current_item = selection[0] if selection else tree.focus()
        children_list = list(children)
        if not children_list:
            return "break"
        if offset == 0:
            return "break"
        iterate_blank_only = bool(self.combined_show_blank_notes_var.get())
        if iterate_blank_only and self.combined_ordered_columns and "Note" in self.combined_ordered_columns:
            direction = 1 if offset >= 0 else -1
            if current_item in children_list:
                start_index = children_list.index(current_item)
            else:
                start_index = -1 if direction > 0 else len(children_list)
            target_item: Optional[str] = None
            index = start_index + direction
            while 0 <= index < len(children_list):
                candidate = children_list[index]
                if self._is_blank_note_tree_item(tree, candidate):
                    target_item = candidate
                    break
                index += direction
            if target_item is None:
                return "break"
            new_item = target_item
        else:
            try:
                index = children_list.index(current_item)
            except ValueError:
                index = 0 if offset >= 0 else len(children_list) - 1
            new_index = index + offset
            if new_index < 0:
                new_index = 0
            elif new_index >= len(children_list):
                new_index = len(children_list) - 1
            new_item = children_list[new_index]
        tree.selection_set(new_item)
        tree.focus(new_item)
        tree.see(new_item)
        self._destroy_note_editor()
        return "break"

    def _on_combined_tree_shift_tab(self, _: tk.Event) -> Optional[str]:  # type: ignore[override]
        return self._move_combined_selection(-1)

    def _on_combined_tree_page(self, offset: int) -> Optional[str]:
        step = offset if offset != 0 else 0
        if step == 0:
            return "break"
        return self._move_combined_selection(step)

    def _on_combined_tree_return(self, _: tk.Event) -> Optional[str]:  # type: ignore[override]
        tree = self.combined_result_tree
        if tree is None or not self.combined_ordered_columns:
            return "break"
        self._destroy_note_editor()
        selection = tree.selection()
        item_id = selection[0] if selection else tree.focus()
        if not item_id:
            return "break"
        row_type = tree.set(item_id, "Type")
        row_category = tree.set(item_id, "Category")
        row_item = tree.set(item_id, "Item")
        key = (str(row_type), str(row_category), str(row_item))

        pdf_value_map: Dict[Path, List[Tuple[str, str]]] = {}
        for column_name in self.combined_ordered_columns:
            if column_name in {"Type", "Category", "Item", "Note"}:
                continue
            mapping = self.combined_column_name_map.get(column_name)
            if not mapping:
                continue
            pdf_path, label_value = mapping
            value_text = tree.set(item_id, column_name) or ""
            if not str(value_text).strip():
                continue
            pdf_value_map.setdefault(pdf_path, []).append((label_value, str(value_text)))

        if not pdf_value_map:
            messagebox.showinfo(
                "Row Values",
                "No PDF columns contain a value for the selected row.",
            )
            return "break"

        ordered_paths: List[Path] = list(self.combined_row_sources.get(key, []))
        for path in pdf_value_map:
            if path not in ordered_paths:
                ordered_paths.append(path)

        preview_items: List[Tuple[PDFEntry, Optional[int], Path, List[Tuple[str, str]]]] = []
        category_key = str(row_type or "")
        for path in ordered_paths:
            values = pdf_value_map.get(path)
            if not values:
                continue
            entry = self.pdf_entry_by_path.get(path)
            if entry is None:
                continue
            page_index = None
            if category_key:
                page_index = self._get_selected_page_index(entry, category_key)
                if page_index is None:
                    matches = entry.matches.get(category_key, [])
                    if matches:
                        page_index = matches[0].page_index
            preview_items.append((entry, page_index, path, values))

        if not preview_items:
            messagebox.showinfo(
                "Row Values",
                "No PDF pages are available for the selected row.",
            )
            return "break"

        dialog = tk.Toplevel(self.root)
        dialog.title("Row PDF Pages")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.focus_set()

        def _close() -> None:
            try:
                dialog.grab_release()
            except tk.TclError:
                pass
            dialog.destroy()

        dialog.protocol("WM_DELETE_WINDOW", _close)
        dialog.bind("<Escape>", lambda _e: (_close(), "break"))

        header_frame = ttk.Frame(dialog, padding=(12, 12, 12, 4))
        header_frame.pack(fill=tk.X)
        ttk.Label(
            header_frame,
            text=f"Type: {row_type or ''} | Category: {row_category or ''} | Item: {row_item or ''}",
            font=("TkDefaultFont", 10, "bold"),
        ).pack(anchor="w")

        content_frame = ttk.Frame(dialog, padding=(12, 0, 12, 0))
        content_frame.pack(fill=tk.BOTH, expand=True)
        canvas = tk.Canvas(content_frame, borderwidth=0)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar = ttk.Scrollbar(content_frame, orient=tk.VERTICAL, command=canvas.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.configure(yscrollcommand=scrollbar.set)
        inner = ttk.Frame(canvas)
        canvas_window = canvas.create_window((0, 0), window=inner, anchor="nw")

        def _on_inner_configure(_: tk.Event) -> None:  # type: ignore[override]
            bbox = canvas.bbox("all")
            if bbox is not None:
                canvas.configure(scrollregion=bbox)

        inner.bind("<Configure>", _on_inner_configure)

        preview_images: List[ImageTk.PhotoImage] = []
        for entry, page_index, path, values in preview_items:
            panel = ttk.Frame(inner, padding=(0, 12))
            panel.pack(fill=tk.X, expand=True)
            panel.columnconfigure(0, weight=1)

            panel_header = ttk.Frame(panel)
            panel_header.grid(row=0, column=0, sticky="ew")
            if page_index is not None:
                title_text = f"{path.name} — Page {page_index + 1}"
            else:
                title_text = f"{path.name} — Page not selected"
            ttk.Label(panel_header, text=title_text, font=("TkDefaultFont", 10, "bold")).pack(side=tk.LEFT)
            ttk.Button(
                panel_header,
                text="Open PDF",
                command=lambda p=path, idx=page_index: self._open_pdf(p, idx),
            ).pack(side=tk.RIGHT)

            preview_label = ttk.Label(panel)
            preview_label.grid(row=1, column=0, sticky="nsew", pady=(8, 4))
            if page_index is not None:
                photo = self._render_page(entry.doc, page_index, 420)
            else:
                photo = None
            if photo is not None:
                preview_label.configure(image=photo)
                preview_label.image = photo  # type: ignore[attr-defined]
                preview_images.append(photo)
            else:
                preview_label.configure(text="Preview unavailable", anchor="center")

            value_lines = [f"{label}: {value}" for label, value in values]
            if value_lines:
                ttk.Label(
                    panel,
                    text="\n".join(value_lines),
                    justify=tk.LEFT,
                    wraplength=520,
                ).grid(row=2, column=0, sticky="ew")

            ttk.Separator(panel, orient=tk.HORIZONTAL).grid(row=3, column=0, sticky="ew", pady=(12, 0))

        canvas.update_idletasks()
        bbox = canvas.bbox(canvas_window)
        if bbox is not None:
            canvas.configure(scrollregion=bbox)

        # Preserve image references on the dialog to avoid garbage collection.
        dialog._preview_images = preview_images  # type: ignore[attr-defined]

        button_frame = ttk.Frame(dialog, padding=(12, 8, 12, 12))
        button_frame.pack(fill=tk.X)
        ttk.Button(button_frame, text="Close", command=_close).pack(side=tk.RIGHT)

        dialog.wait_visibility()
        dialog.focus_set()
        return "break"

    def _set_note_value(self, item_id: str, value: str) -> None:
        tree = self.combined_result_tree
        if tree is None or not self.combined_ordered_columns or "Note" not in self.combined_ordered_columns:
            return
        if isinstance(value, str):
            normalized_value = value.strip().lower()
        else:
            normalized_value = str(value or "").strip().lower()
        if normalized_value and normalized_value not in self.note_options:
            self._register_note_option(normalized_value)
        if normalized_value and normalized_value not in self.note_options:
            normalized_value = ""
        note_index = self.combined_ordered_columns.index("Note")
        values = list(tree.item(item_id, "values"))
        if note_index < len(values):
            if values[note_index] != normalized_value:
                values[note_index] = normalized_value
                tree.item(item_id, values=values)
        self._apply_note_value_tag(item_id, normalized_value)
        key = self.combined_note_record_keys.get(item_id)
        if not key:
            return
        if normalized_value:
            self.note_assignments[key] = normalized_value
        else:
            self.note_assignments.pop(key, None)
        self._update_combined_record_note_value(key, normalized_value)
        self._write_note_assignments()

    def _apply_saved_selection(self, entry: PDFEntry) -> None:
        key = entry.path.name
        saved = self.assigned_pages.get(key)
        if not saved:
            return
        selections = saved.get("selections") if isinstance(saved, dict) else None
        if selections is None:
            selections = saved
        highlight_sources: Dict[str, Iterable[Any]] = {}
        if isinstance(saved, dict):
            raw_highlights = saved.get("highlights")
            if isinstance(raw_highlights, dict):
                for category, stored_pages in raw_highlights.items():
                    if isinstance(stored_pages, (list, tuple, set)):
                        highlight_sources[category] = stored_pages
                    else:
                        highlight_sources[category] = [stored_pages]

        highlight_indexes: Dict[str, List[int]] = {}
        for category, values in highlight_sources.items():
            matches = entry.matches.get(category)
            if matches is None:
                continue
            indexes: List[int] = []
            for value in values:
                try:
                    page_int = int(value)
                except (TypeError, ValueError):
                    continue
                target_idx: Optional[int] = None
                for idx, match in enumerate(matches):
                    if match.page_index == page_int:
                        target_idx = idx
                        break
                if target_idx is None:
                    match = self._add_manual_match(entry, category, page_int)
                    matches = entry.matches[category]
                    target_idx = matches.index(match)
                if target_idx not in indexes:
                    indexes.append(target_idx)
            if indexes:
                self._set_review_highlights(entry, category, indexes)
                resolved_indexes = self._get_highlight_match_indexes(entry, category)
                if resolved_indexes:
                    highlight_indexes[category] = resolved_indexes
                    entry.current_index[category] = resolved_indexes[0]

        for category, page_index in selections.items():
            matches = entry.matches.get(category)
            if matches is None:
                continue
            try:
                page_int = int(page_index)
            except (TypeError, ValueError):
                continue
            target_idx: Optional[int] = None
            for idx, match in enumerate(matches):
                if match.page_index == page_int:
                    target_idx = idx
                    break
            if target_idx is None:
                match = self._add_manual_match(entry, category, page_int)
                matches = entry.matches[category]
                target_idx = matches.index(match)
            existing = highlight_indexes.get(category)
            if existing:
                if target_idx not in existing:
                    new_indexes = existing + [target_idx]
                    self._set_review_highlights(entry, category, new_indexes)
                    highlight_indexes[category] = self._get_highlight_match_indexes(entry, category)
            else:
                self._set_review_highlights(entry, category, [target_idx])
                highlight_indexes[category] = self._get_highlight_match_indexes(entry, category)
            entry.current_index[category] = target_idx
        if not entry.year and isinstance(saved, dict):
            stored_year = saved.get("year")
            if isinstance(stored_year, str):
                entry.year = stored_year

    def _clear_entries(self) -> None:
        for entry in self.pdf_entries:
            try:
                entry.doc.close()
            except Exception:
                pass
        self.pdf_entries.clear()
        self.pdf_entry_by_path.clear()
        self.review_highlighted_matches.clear()
        self.category_rows.clear()
        self.year_vars.clear()
        for child in self.inner_frame.winfo_children():
            child.destroy()

    def _rebuild_grid(self) -> None:
        for child in self.inner_frame.winfo_children():
            child.destroy()
        self.category_rows.clear()

        if not self.pdf_entries:
            ttk.Label(self.inner_frame, text="Load PDFs to begin reviewing assignments.").grid(
                row=0, column=0, padx=16, pady=16, sticky="nw"
            )
            return

        for row_index, entry in enumerate(self.pdf_entries):
            container = ttk.Frame(self.inner_frame, padding=8)
            container.grid(row=row_index, column=0, sticky="ew", padx=4, pady=4)
            container.columnconfigure(1, weight=1)

            relative_path = (
                entry.path.relative_to(Path(self.folder_path.get())) if self.folder_path.get() else entry.path.name
            )
            info_frame = ttk.Frame(container)
            info_frame.grid(row=0, column=0, sticky="nw", padx=(0, 12))
            name_label = ttk.Label(info_frame, text=relative_path, anchor="w", width=30, wraplength=200)
            name_label.pack(anchor="w")
            name_label.bind("<Alt-Button-1>", lambda _e, ent=entry: self.open_entry_overview(ent))
            year_var = tk.StringVar(value=entry.year)
            self.year_vars[entry.path] = year_var
            year_var.trace_add("write", lambda *_args, e=entry, v=year_var: setattr(e, "year", v.get()))
            ttk.Entry(info_frame, textvariable=year_var, width=30).pack(fill=tk.X, pady=(4, 0))

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

    def _refresh_category_row(self, entry: PDFEntry, category: str, *, rebuild: bool) -> None:
        row = self.category_rows.get((entry.path, category))
        if row is None:
            return
        if rebuild:
            row.refresh()
        else:
            row.update_selection()

    def _review_highlight_key(self, entry: PDFEntry, category: str) -> Tuple[Path, str]:
        return (entry.path, category)

    def _prune_review_highlights(self, entry: PDFEntry, category: str, match_count: int) -> None:
        key = self._review_highlight_key(entry, category)
        highlights = self.review_highlighted_matches.get(key)
        if not highlights:
            return
        filtered = {idx for idx in highlights if 0 <= idx < match_count}
        if filtered:
            self.review_highlighted_matches[key] = filtered
        else:
            self.review_highlighted_matches.pop(key, None)
        self._sync_entry_highlights(entry, category, persist=False)

    def _get_highlight_match_indexes(self, entry: PDFEntry, category: str) -> List[int]:
        key = self._review_highlight_key(entry, category)
        highlights = self.review_highlighted_matches.get(key)
        if not highlights:
            return []
        matches = entry.matches.get(category, [])
        if not matches:
            return []
        valid_indexes: List[int] = []
        for idx in sorted(highlights):
            if 0 <= idx < len(matches):
                valid_indexes.append(int(idx))
        return valid_indexes

    def _get_highlight_page_numbers(self, entry: PDFEntry, category: str) -> List[int]:
        matches = entry.matches.get(category, [])
        if not matches:
            return []
        indexes = self._get_highlight_match_indexes(entry, category)
        if not indexes:
            return []
        seen: Set[int] = set()
        pages: List[int] = []
        for idx in sorted(indexes, key=lambda i: matches[i].page_index if 0 <= i < len(matches) else i):
            if 0 <= idx < len(matches):
                page_value = int(matches[idx].page_index)
                if page_value not in seen:
                    seen.add(page_value)
                    pages.append(page_value)
        return pages

    def _ensure_assigned_record(self, entry: PDFEntry) -> Dict[str, Any]:
        key = entry.path.name
        record = self.assigned_pages.get(key)
        if not isinstance(record, dict):
            record = {"selections": {}, "year": entry.year}
            self.assigned_pages[key] = record
        if entry.year:
            record["year"] = entry.year
        selections = record.get("selections")
        if not isinstance(selections, dict):
            selections = {}
            record["selections"] = selections
        return record

    def _sync_entry_highlights(
        self, entry: PDFEntry, category: str, *, persist: bool = False
    ) -> None:
        pages = self._get_highlight_page_numbers(entry, category)
        existing_record = self.assigned_pages.get(entry.path.name)
        if not pages and not isinstance(existing_record, dict):
            return
        record = self._ensure_assigned_record(entry)
        highlights = record.get("highlights")
        if not isinstance(highlights, dict):
            highlights = {}
            record["highlights"] = highlights
        if pages:
            highlights[category] = list(pages)
        else:
            highlights.pop(category, None)
            if not highlights:
                record.pop("highlights", None)
                if not record.get("selections") and not record.get("year"):
                    self.assigned_pages.pop(entry.path.name, None)
        if persist:
            self._write_assigned_pages()

    def _set_review_highlights(
        self, entry: PDFEntry, category: str, indexes: Iterable[int]
    ) -> None:
        key = self._review_highlight_key(entry, category)
        normalized = {int(idx) for idx in indexes}
        if normalized:
            self.review_highlighted_matches[key] = normalized
        else:
            self.review_highlighted_matches.pop(key, None)
        self._sync_entry_highlights(entry, category, persist=False)

    def _add_review_highlight(self, entry: PDFEntry, category: str, index: int) -> None:
        key = self._review_highlight_key(entry, category)
        highlights = self.review_highlighted_matches.setdefault(key, set())
        highlights.add(int(index))
        self._sync_entry_highlights(entry, category, persist=False)

    def _remove_review_highlight(self, entry: PDFEntry, category: str, index: int) -> None:
        key = self._review_highlight_key(entry, category)
        highlights = self.review_highlighted_matches.get(key)
        if not highlights:
            return
        highlights.discard(int(index))
        if not highlights:
            self.review_highlighted_matches.pop(key, None)
        self._sync_entry_highlights(entry, category, persist=False)

    def is_match_highlighted(self, entry: PDFEntry, category: str, index: int) -> bool:
        key = self._review_highlight_key(entry, category)
        highlights = self.review_highlighted_matches.get(key)
        return bool(highlights and int(index) in highlights)

    def select_match_index(self, entry: PDFEntry, category: str, index: int) -> None:
        matches = entry.matches.get(category, [])
        if not matches:
            return
        index = max(0, min(index, len(matches) - 1))
        entry.current_index[category] = index
        self._refresh_category_row(entry, category, rebuild=False)
        page_index = matches[index].page_index
        self._update_assigned_entry(entry, category, page_index, persist=False)
        self._maybe_reapply_primary_match_filter(entry)

    def _rescan_entries(
        self,
        pattern_map: Dict[str, List[re.Pattern[str]]],
        year_patterns: List[re.Pattern[str]],
        columns: set[str],
        year_changed: bool,
    ) -> None:
        for entry in self.pdf_entries:
            previous_pages: Dict[str, Optional[int]] = {}
            manual_matches: Dict[str, List[Match]] = {}
            new_matches: Dict[str, List[Match]] = {column: [] for column in columns}

            for column in columns:
                existing_matches = entry.matches.get(column, [])
                manual_matches[column] = [m for m in existing_matches if m.source == "manual"]
                current_index = entry.current_index.get(column)
                if current_index is not None and 0 <= current_index < len(existing_matches):
                    previous_pages[column] = existing_matches[current_index].page_index
                else:
                    previous_pages[column] = None

            detected_year: Optional[str] = None

            for page_index in range(len(entry.doc)):
                page = entry.doc.load_page(page_index)
                page_text = page.get_text("text")

                if year_changed and detected_year is None:
                    for pattern in year_patterns:
                        year_match = pattern.search(page_text)
                        if year_match:
                            if year_match.groups():
                                detected_year = year_match.group(1)
                            else:
                                detected_year = year_match.group(0)
                            break

                for column in columns:
                    compiled_patterns = pattern_map.get(column, [])
                    for pattern in compiled_patterns:
                        match_obj = pattern.search(page_text)
                        if match_obj:
                            matched_text = match_obj.group(0).strip()
                            new_matches[column].append(
                                Match(
                                    page_index=page_index,
                                    source="regex",
                                    pattern=pattern.pattern,
                                    matched_text=matched_text,
                                )
                            )
                            break

            for column in columns:
                matches = new_matches[column]
                manual = manual_matches[column]
                manual_pages = {match.page_index for match in matches}
                for manual_match in manual:
                    if manual_match.page_index not in manual_pages:
                        matches.append(manual_match)
                entry.matches[column] = matches
                entry.all_matches[column] = list(matches)

                if matches:
                    target_page = previous_pages[column]
                    if target_page is not None:
                        for idx, match in enumerate(matches):
                            if match.page_index == target_page:
                                entry.current_index[column] = idx
                                break
                        else:
                            entry.current_index[column] = 0
                    else:
                        entry.current_index[column] = 0
                    assigned_index = entry.current_index[column]
                    if assigned_index is not None:
                        self._set_review_highlights(entry, column, [assigned_index])
                    else:
                        self._set_review_highlights(entry, column, [])
                else:
                    entry.current_index[column] = None
                    self._set_review_highlights(entry, column, [])

                self._refresh_category_row(entry, column, rebuild=True)

            if year_changed:
                if detected_year is not None:
                    entry.year = detected_year
                    year_var = self.year_vars.get(entry.path)
                    if year_var is not None and year_var.get() != detected_year:
                        year_var.set(detected_year)
                else:
                    year_var = self.year_vars.get(entry.path)
                    if year_var is not None and year_var.get() != entry.year:
                        year_var.set(entry.year)

        if self.review_primary_match_filter_var.get():
            self._apply_review_primary_match_filter()

    def _render_page(self, doc: fitz.Document, page_index: int, target_width: int) -> Optional[ImageTk.PhotoImage]:
        try:
            page = doc.load_page(page_index)
            zoom_matrix = fitz.Matrix(1.5, 1.5)
            pix = page.get_pixmap(matrix=zoom_matrix)
            mode = "RGBA" if pix.alpha else "RGB"
            image = Image.frombytes(mode, [pix.width, pix.height], pix.samples)
            if image.mode == "RGBA":
                image = image.convert("RGB")
            if target_width > 0 and image.width != target_width:
                ratio = target_width / image.width
                new_size = (int(image.width * ratio), int(image.height * ratio))
                image = image.resize(new_size, Image.LANCZOS)
            return ImageTk.PhotoImage(image, master=self.root)
        except Exception as exc:  # pragma: no cover - guard for rendering issues
            messagebox.showwarning("Render Error", f"Could not render page {page_index + 1}: {exc}")
            return None

    def cycle_match(self, entry: PDFEntry, category: str, *, forward: bool) -> None:
        matches = entry.matches.get(category, [])
        if not matches:
            messagebox.showinfo("No Matches", f"No matches available for {category} in {entry.path.name}.")
            return

        current_index = entry.current_index.get(category) or 0
        if forward:
            if current_index + 1 >= len(matches):
                messagebox.showinfo("End of Matches", "Reached the last matched page.")
                return
            entry.current_index[category] = current_index + 1
        else:
            if current_index - 1 < 0:
                messagebox.showinfo("Start of Matches", "Reached the first matched page.")
                return
            entry.current_index[category] = current_index - 1
        new_index = entry.current_index.get(category)
        if new_index is not None:
            self._set_review_highlights(entry, category, [new_index])
        else:
            self._set_review_highlights(entry, category, [])
        self._refresh_category_row(entry, category, rebuild=False)
        page_index = self._get_selected_page_index(entry, category)
        if page_index is not None:
            self._update_assigned_entry(entry, category, page_index, persist=False)
        self._maybe_reapply_primary_match_filter(entry)

    def manual_select(self, entry: PDFEntry, category: str) -> None:
        pdf_path = entry.path
        self._open_pdf(pdf_path)
        page_number = simpledialog.askinteger(
            "Manual Page Selection",
            f"Enter page number for {category} in {pdf_path.name}:",
            parent=self.root,
            minvalue=1,
            maxvalue=len(entry.doc),
        )
        if page_number is None:
            return

        page_index = page_number - 1
        self._set_selected_page(entry, category, page_index)

    def open_current_match(self, entry: PDFEntry, category: str) -> None:
        matches = entry.matches.get(category, [])
        if not matches:
            messagebox.showinfo("No Matches", f"No matches available for {category} in {entry.path.name}.")
            return
        index = entry.current_index.get(category) or 0
        index = max(0, min(index, len(matches) - 1))
        page_index = matches[index].page_index
        self._open_pdf(entry.path, page_index)

    def _open_pdf(self, pdf_path: Path, page_index: Optional[int] = None) -> None:
        try:
            if page_index is not None:
                url = pdf_path.resolve().as_uri() + f"#page={page_index + 1}"
                opened = webbrowser.open(url)
                if opened:
                    return
            if sys.platform.startswith("win"):
                os.startfile(pdf_path)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                os.system(f"open '{pdf_path}'")
            else:
                os.system(f"xdg-open '{pdf_path}' >/dev/null 2>&1 &")
        except Exception as exc:
            messagebox.showwarning("Open PDF", f"Could not open PDF: {exc}")

    def open_entry_overview(self, entry: PDFEntry) -> None:
        dialog = tk.Toplevel(self.root)
        dialog.title(f"Overview - {entry.path.name}")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.focus_set()

        def _close() -> None:
            try:
                dialog.grab_release()
            except tk.TclError:
                pass
            dialog.destroy()

        dialog.protocol("WM_DELETE_WINDOW", _close)

        def _on_escape(_: tk.Event) -> str:  # type: ignore[override]
            _close()
            return "break"

        dialog.bind("<Escape>", _on_escape)

        container = ttk.Frame(dialog)
        container.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(container)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar = ttk.Scrollbar(container, orient=tk.VERTICAL, command=canvas.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.configure(yscrollcommand=scrollbar.set)

        inner = ttk.Frame(canvas)
        window = canvas.create_window((0, 0), window=inner, anchor="nw")

        def _update_scroll(_: tk.Event) -> None:  # type: ignore[override]
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _resize(event: tk.Event) -> None:  # type: ignore[override]
            canvas.itemconfigure(window, width=event.width)

        inner.bind("<Configure>", _update_scroll)
        canvas.bind("<Configure>", _resize)

        thumbnails: List[ImageTk.PhotoImage] = []
        for idx, column in enumerate(COLUMNS):
            frame = ttk.LabelFrame(inner, text=column, padding=8)
            row = idx // 2
            col = idx % 2
            frame.grid(row=row, column=col, padx=8, pady=8, sticky="nsew")
            inner.columnconfigure(col, weight=1)
            inner.rowconfigure(row, weight=1)
            page_index = self._get_selected_page_index(entry, column)
            if page_index is None:
                ttk.Label(frame, text="No page selected").pack(expand=True, fill=tk.BOTH)
                continue
            photo = self._render_page(entry.doc, page_index, 420)
            if photo is not None:
                thumbnails.append(photo)
                ttk.Label(frame, image=photo).pack(expand=True, fill=tk.BOTH)
            else:
                ttk.Label(frame, text="Preview unavailable").pack(expand=True, fill=tk.BOTH)
            ttk.Label(frame, text=f"Page {page_index + 1}").pack(anchor="w", pady=(4, 0))

        if len(COLUMNS) % 2:
            inner.columnconfigure(1, weight=1)

        if not thumbnails and not any(self._get_selected_page_index(entry, column) is not None for column in COLUMNS):
            ttk.Label(inner, text="No selections available for this PDF.").grid(row=0, column=0, padx=16, pady=16, sticky="nsew")

        dialog._thumbnails = thumbnails  # type: ignore[attr-defined]

    def open_thumbnail_zoom(self, entry: PDFEntry, page_index: int) -> None:
        dialog = tk.Toplevel(self.root)
        dialog.title(f"{entry.path.name} - Page {page_index + 1}")
        dialog.transient(self.root)

        def _attempt_grab() -> None:
            if not dialog.winfo_exists():
                return
            try:
                dialog.grab_set()
            except tk.TclError:
                dialog.after(50, _attempt_grab)

        dialog.after(0, _attempt_grab)
        dialog.focus_set()

        # Ensure the zoom dialog opens maximized so reviewers can inspect the page comfortably.
        self._maximize_window(dialog)
        dialog.update_idletasks()

        def _close() -> None:
            try:
                dialog.grab_release()
            except tk.TclError:
                pass
            dialog.destroy()

        dialog.protocol("WM_DELETE_WINDOW", _close)

        def _on_escape(_: tk.Event) -> str:  # type: ignore[override]
            _close()
            return "break"

        dialog.bind("<Escape>", _on_escape)

        container = ttk.Frame(dialog, padding=8)
        container.pack(fill=tk.BOTH, expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)

        canvas = tk.Canvas(container, highlightthickness=0)
        canvas.grid(row=0, column=0, sticky="nsew")

        y_scroll = ttk.Scrollbar(container, orient=tk.VERTICAL, command=canvas.yview)
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll = ttk.Scrollbar(container, orient=tk.HORIZONTAL, command=canvas.xview)
        x_scroll.grid(row=1, column=0, sticky="ew")
        canvas.configure(xscrollcommand=x_scroll.set, yscrollcommand=y_scroll.set)

        available_width = dialog.winfo_width()
        if available_width <= 1:
            available_width = dialog.winfo_screenwidth()
        target_width = max(480, available_width - 64, self.thumbnail_width_var.get() * 2)
        photo = self._render_page(entry.doc, page_index, target_width)

        if photo is not None:
            canvas.create_image(0, 0, anchor="nw", image=photo)
            dialog._photo = photo  # type: ignore[attr-defined]
        else:
            canvas.create_text(
                canvas.winfo_reqwidth() // 2,
                canvas.winfo_reqheight() // 2,
                text="Preview unavailable",
                anchor="center",
            )

        bbox = canvas.bbox("all")
        if bbox:
            canvas.configure(scrollregion=bbox)

        def _on_mousewheel(event: tk.Event) -> str:  # type: ignore[override]
            delta = getattr(event, "delta", 0)
            if delta == 0:
                button = getattr(event, "num", 0)
                if button in (4, 6):
                    delta = 120
                elif button in (5, 7):
                    delta = -120
                else:
                    delta = -120
            if getattr(event, "state", 0) & SHIFT_MASK:
                canvas.xview_scroll(-1 if delta > 0 else 1, "units")
            else:
                canvas.yview_scroll(-1 if delta > 0 else 1, "units")
            return "break"

        canvas.bind("<MouseWheel>", _on_mousewheel)
        canvas.bind("<Shift-MouseWheel>", _on_mousewheel)
        canvas.bind("<Button-4>", _on_mousewheel)
        canvas.bind("<Button-5>", _on_mousewheel)
        canvas.bind("<Shift-Button-4>", _on_mousewheel)
        canvas.bind("<Shift-Button-5>", _on_mousewheel)

        windowing_system = str(self.root.tk.call("tk", "windowingsystem")).lower()
        if windowing_system == "x11":
            for sequence in (
                "<Button-6>",
                "<Button-7>",
                "<Shift-Button-6>",
                "<Shift-Button-7>",
            ):
                try:
                    canvas.bind(sequence, _on_mousewheel)
                except tk.TclError:
                    # Some Tk builds may omit extended button bindings; ignore failures.
                    pass

        canvas.bind("<Button-1>", lambda _e: _close())

    def _set_selected_page(self, entry: PDFEntry, category: str, page_index: int) -> None:
        matches = entry.matches[category]
        for idx, match in enumerate(matches):
            if match.page_index == page_index:
                entry.current_index[category] = idx
                self._set_review_highlights(entry, category, [idx])
                self._refresh_category_row(entry, category, rebuild=False)
                self._update_assigned_entry(entry, category, page_index, persist=False)
                self._maybe_reapply_primary_match_filter(entry)
                return

        match = self._add_manual_match(entry, category, page_index)
        assigned_index = entry.matches[category].index(match)
        entry.current_index[category] = assigned_index
        self._set_review_highlights(entry, category, [assigned_index])
        self._refresh_category_row(entry, category, rebuild=True)
        self._update_assigned_entry(entry, category, page_index, persist=False)
        self._maybe_reapply_primary_match_filter(entry)

    def _get_selected_page_index(self, entry: PDFEntry, category: str) -> Optional[int]:
        matches = entry.matches.get(category, [])
        if not matches:
            return None
        index = entry.current_index.get(category)
        if index is None or not (0 <= index < len(matches)):
            return None
        return matches[index].page_index

    def _update_assigned_entry(
        self, entry: PDFEntry, category: str, page_index: int, *, persist: bool = False
    ) -> None:
        record = self._ensure_assigned_record(entry)
        selections = record.setdefault("selections", {})
        selections[category] = int(page_index)
        self._sync_entry_highlights(entry, category, persist=False)
        if persist:
            self._write_assigned_pages()

    def _write_assigned_pages(self) -> None:
        if self.assigned_pages_path is None:
            return
        self.assigned_pages_path.parent.mkdir(parents=True, exist_ok=True)
        try:
            with self.assigned_pages_path.open("w", encoding="utf-8") as fh:
                json.dump(self.assigned_pages, fh, indent=2)
        except Exception as exc:
            messagebox.showwarning("Save Assignments", f"Could not save assigned pages: {exc}")

    def commit_assignments(self) -> None:
        company = self.company_var.get()
        if not company:
            messagebox.showinfo("Select Company", "Please choose a company before committing assignments.")
            return
        if not messagebox.askyesno(
            "Confirm Commit",
            f"Commit the current assignments for {company}?",
            parent=self.root,
        ):
            return
        if self.assigned_pages_path is None:
            self.assigned_pages_path = self.companies_dir / company / "assigned.json"

        if self.pdf_entries:
            for entry in self.pdf_entries:
                record = self._ensure_assigned_record(entry)
                record["year"] = entry.year
                selections = record.setdefault("selections", {})
                highlights = record.setdefault("highlights", {})
                for category in COLUMNS:
                    page_index = self._get_selected_page_index(entry, category)
                    if page_index is None:
                        selections.pop(category, None)
                    else:
                        selections[category] = int(page_index)
                    highlight_pages = self._get_highlight_page_numbers(entry, category)
                    if highlight_pages:
                        highlights[category] = list(highlight_pages)
                    else:
                        highlights.pop(category, None)
                if not highlights:
                    record.pop("highlights", None)

        self._write_assigned_pages()
        self._refresh_scraped_tab()

    def _refresh_scraped_tab(self) -> None:
        if not hasattr(self, "scraped_inner"):
            return
        for child in self.scraped_inner.winfo_children():
            child.destroy()
        self.scraped_images.clear()
        self.scraped_table_sources = {}
        self.scraped_preview_states = {}
        self._clear_combined_tab()
        if hasattr(self, "scrape_progress"):
            self.scrape_progress["value"] = 0
        company = self.company_var.get()
        if not company:
            return
        scrape_root = self.companies_dir / company / "openapiscrape"
        if not scrape_root.exists():
            return

        self.scraped_inner.columnconfigure(0, weight=1, uniform="scraped_split")
        self.scraped_inner.columnconfigure(1, weight=2, uniform="scraped_split")
        if hasattr(self, "scraped_canvas"):
            self.scraped_canvas.update_idletasks()
            available_width = self.scraped_canvas.winfo_width()
        else:
            available_width = 0
        if available_width <= 0:
            available_width = self.root.winfo_width()
        if available_width <= 0:
            available_width = 900
        pdf_target_width = max(200, available_width // 3)
        row_index = 0
        header_added = False
        for entry in self.pdf_entries:
            entry_stem = entry.stem
            metadata_path = self._metadata_file_for_entry(scrape_root, entry_stem)
            metadata: Dict[str, Any] = {}
            if metadata_path.exists():
                metadata = self._load_doc_metadata(metadata_path)
            else:
                legacy_dir = self._legacy_doc_dir(scrape_root, entry_stem)
                if legacy_dir.exists():
                    metadata = self._load_doc_metadata(legacy_dir)
            if not metadata:
                continue
            for category in COLUMNS:
                meta = metadata.get(category)
                if not isinstance(meta, dict):
                    continue
                page_index_value = meta.get("page_index")
                try:
                    initial_page_index = (
                        int(page_index_value)
                        if page_index_value is not None
                        else None
                    )
                except (TypeError, ValueError):
                    initial_page_index = None
                meta_page_indexes = meta.get("page_indexes")
                parsed_page_indexes: List[int] = []
                if isinstance(meta_page_indexes, list):
                    for idx in meta_page_indexes:
                        try:
                            parsed_page_indexes.append(int(idx))
                        except (TypeError, ValueError):
                            continue
                if initial_page_index is not None:
                    try:
                        current_page_position = parsed_page_indexes.index(initial_page_index)
                    except ValueError:
                        parsed_page_indexes.insert(0, initial_page_index)
                        current_page_position = 0
                else:
                    current_page_position = 0
                if not parsed_page_indexes:
                    if initial_page_index is None:
                        continue
                    parsed_page_indexes = [initial_page_index]
                    current_page_position = 0
                display_page_index = parsed_page_indexes[current_page_position]
                preview_mode: Optional[str] = None
                preview_path: Optional[Path] = None
                preview_text: Optional[str] = None
                rows: List[List[str]] = []
                csv_target_path: Optional[Path] = None
                csv_delimiter = ","

                txt_name = meta.get("txt")
                if isinstance(txt_name, str) and txt_name:
                    candidate = self._resolve_scrape_path(scrape_root, entry_stem, txt_name)
                    if candidate.exists():
                        try:
                            preview_text = candidate.read_text(encoding="utf-8")
                        except Exception:
                            preview_text = None
                        else:
                            rows = self._convert_response_to_rows(preview_text)
                            preview_mode = "txt"
                            preview_path = candidate

                if not rows:
                    csv_name = meta.get("csv")
                    if isinstance(csv_name, str) and csv_name:
                        candidate_csv = self._resolve_scrape_path(
                            scrape_root, entry_stem, csv_name
                        )
                        csv_target_path = candidate_csv
                        if candidate_csv.exists():
                            csv_delimiter = self._detect_csv_delimiter(candidate_csv)
                            rows = self._read_csv_rows(candidate_csv)
                            if rows:
                                preview_mode = "csv"
                                preview_path = candidate_csv

                if not rows:
                    continue
                if not header_added:
                    header_frame = ttk.Frame(self.scraped_inner)
                    header_frame.grid(row=row_index, column=0, columnspan=2, sticky="ew", padx=8, pady=(8, 0))
                    ttk.Label(
                        header_frame,
                        text="Original PDF Page",
                        font=("TkDefaultFont", 10, "bold"),
                    ).grid(row=0, column=0, sticky="w")
                    ttk.Label(
                        header_frame,
                        text="Parsed Output",
                        font=("TkDefaultFont", 10, "bold"),
                    ).grid(row=0, column=1, sticky="w")
                    header_frame.columnconfigure(0, weight=1)
                    header_frame.columnconfigure(1, weight=1)
                    row_index += 1
                    header_added = True
                header_frame = ttk.Frame(self.scraped_inner)
                header_frame.grid(row=row_index, column=0, columnspan=2, sticky="ew", padx=8, pady=(8, 0))
                header_frame.columnconfigure(0, weight=1)
                header_frame.columnconfigure(1, weight=0)
                header_frame.columnconfigure(2, weight=0)
                header_frame.columnconfigure(3, weight=0)
                header_frame.columnconfigure(4, weight=0)
                header_frame.columnconfigure(5, weight=0)
                header_frame.columnconfigure(6, weight=0)
                base_header_text = f"{entry.path.name} - {category}"
                title_label = ttk.Label(
                    header_frame,
                    text=f"{base_header_text} (Page {display_page_index + 1})",
                    font=("TkDefaultFont", 10, "bold"),
                )
                title_label.grid(row=0, column=0, sticky="w")
                ttk.Button(
                    header_frame,
                    text="View Prompt",
                    command=lambda e=entry, c=category, p=int(display_page_index): self._show_prompt_preview(
                        e, c, p
                    ),
                    width=12,
                ).grid(row=0, column=1, sticky="e", padx=(0, 4))
                ttk.Button(
                    header_frame,
                    text="View Raw Text",
                    command=lambda e=entry, p=int(display_page_index): self._show_raw_text_dialog(e, p),
                    width=14,
                ).grid(row=0, column=2, sticky="e", padx=(0, 4))
                ttk.Button(
                    header_frame,
                    text="View TXT" if preview_mode == "txt" else "View CSV",
                    command=(
                        (lambda path=preview_path, text=preview_text: self._show_text_preview(path, text))
                        if preview_mode == "txt"
                        else (lambda path=preview_path: self._show_csv_preview(path) if path else None)
                    ),
                    width=10,
                ).grid(row=0, column=3, sticky="e", padx=(0, 4))
                csv_name = meta.get("csv") if isinstance(meta, dict) else None
                csv_path: Optional[Path] = None
                if isinstance(csv_name, str) and csv_name:
                    candidate_csv = self._resolve_scrape_path(
                        scrape_root, entry_stem, csv_name
                    )
                    if candidate_csv.exists():
                        csv_path = candidate_csv
                if csv_path is not None:
                    ttk.Button(
                        header_frame,
                        text="Reload CSV",
                        command=lambda root=scrape_root, meta_path=metadata_path, stem=entry_stem, c=category: self._reload_scraped_csv(
                            root, meta_path, stem, c
                        ),
                        width=12,
                    ).grid(row=0, column=4, sticky="e", padx=(0, 4))
                    ttk.Button(
                        header_frame,
                        text="Open CSV",
                        command=lambda path=csv_path: self._open_csv_file(path) if path else None,
                        width=10,
                    ).grid(row=0, column=5, sticky="e", padx=(0, 4))
                    delete_column = 6
                else:
                    delete_column = 4
                ttk.Button(
                    header_frame,
                    text="Delete",
                    command=lambda root=scrape_root, meta_path=metadata_path, stem=entry_stem, c=category: self._delete_scrape_output(
                        root, meta_path, stem, c
                    ),
                    width=10,
                ).grid(row=0, column=delete_column, sticky="e")
                row_index += 1

                image_frame = ttk.Frame(self.scraped_inner, padding=8)
                image_frame.grid(row=row_index, column=0, sticky="nsew", padx=4)
                image_frame.columnconfigure(0, weight=1)
                image_frame.rowconfigure(0, weight=1)
                table_frame = ttk.Frame(self.scraped_inner, padding=8)
                table_frame.grid(row=row_index, column=1, sticky="nsew", padx=4)
                self.scraped_inner.rowconfigure(row_index, weight=1)

                preview_widget: Optional[tk.Widget] = None
                photo = self._render_page(entry.doc, int(display_page_index), pdf_target_width)
                if photo is not None:
                    self.scraped_images.append(photo)
                    image_label = ttk.Label(image_frame, image=photo, cursor="hand2")
                    image_label.grid(row=0, column=0, sticky="nsew")
                    image_label.bind(
                        "<Button-1>",
                        lambda event, lbl=image_label: self._on_scraped_image_click(event, lbl),
                    )
                    image_label.bind(
                        "<Control-Button-1>",
                        lambda event, lbl=image_label: self._on_scraped_image_click(event, lbl),
                    )
                    self.scraped_preview_states[image_label] = {
                        "entry": entry,
                        "category": category,
                        "page_indexes": list(parsed_page_indexes),
                        "position": current_page_position,
                        "current_page": display_page_index,
                        "target_width": pdf_target_width,
                        "title_label": title_label,
                        "title_base_text": base_header_text,
                    }
                    preview_widget = image_label
                    image_frame.bind(
                        "<Configure>",
                        lambda event, lbl=image_label: self._on_scraped_image_frame_resize(
                            event, lbl
                        ),
                    )
                else:
                    ttk.Label(image_frame, text="Preview unavailable").grid(row=0, column=0, sticky="nsew")

                table_frame.columnconfigure(0, weight=1)
                table_frame.columnconfigure(1, weight=0)
                table_frame.rowconfigure(1, weight=1)

                if not rows:
                    ttk.Label(table_frame, text="No data available").grid(
                        row=0, column=0, sticky="nsew"
                    )
                else:
                    headings = rows[0]
                    data_rows = rows[1:] if len(rows) > 1 else []
                    if not any(heading.strip() for heading in headings):
                        headings = [f"Column {idx + 1}" for idx in range(len(headings))]
                    columns = [f"col_{idx}" for idx in range(len(headings))]
                    controls_frame = ttk.Frame(table_frame)
                    controls_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 4))
                    controls_frame.columnconfigure(0, weight=1)
                    tree = ttk.Treeview(
                        table_frame,
                        columns=columns,
                        show="headings",
                        height=max(3, len(data_rows)),
                        selectmode="extended",
                    )
                    y_scroll = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=tree.yview)
                    x_scroll = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=tree.xview)
                    tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
                    tree.grid(row=1, column=0, sticky="nsew")
                    y_scroll.grid(row=1, column=1, sticky="ns")
                    x_scroll.grid(row=2, column=0, sticky="ew")
                    for col, heading in zip(columns, headings):
                        tree.heading(col, text=heading)
                        tree.column(col, anchor="center", stretch=True, width=120)
                    table_rows: List[List[str]] = []
                    scalable_columns = self._scraped_scalable_column_indices(headings)
                    if data_rows:
                        for data in data_rows:
                            values = list(data)
                            if len(values) < len(columns):
                                values.extend([""] * (len(columns) - len(values)))
                            elif len(values) > len(columns):
                                values = values[: len(columns)]
                            table_rows.append(values.copy())
                            display_values = [self._format_table_value(value) for value in values]
                            tree.insert("", tk.END, values=display_values)

                    multiply_button = ttk.Button(
                        controls_frame,
                        text="×10",
                        command=lambda t=tree: self._scale_scraped_table(t, 10.0),
                        width=12,
                    )
                    multiply_button.grid(row=0, column=1, sticky="e", padx=(0, 4))

                    divide_button = ttk.Button(
                        controls_frame,
                        text="÷10",
                        command=lambda t=tree: self._scale_scraped_table(t, 0.1),
                        width=12,
                    )
                    divide_button.grid(row=0, column=2, sticky="e", padx=(0, 4))

                    open_page_button = ttk.Button(
                        controls_frame,
                        text="Open Page",
                        command=lambda t=tree: self._open_scraped_table_preview(t),
                        width=12,
                    )
                    open_page_button.grid(row=0, column=3, sticky="e", padx=(0, 4))

                    relabel_button = ttk.Button(
                        controls_frame,
                        text="Relabel Dates",
                        command=lambda t=tree: self._prompt_relabel_scraped_dates(t),
                        width=14,
                    )
                    relabel_button.grid(row=0, column=4, sticky="e", padx=(0, 4))

                    dedupe_button = ttk.Button(
                        controls_frame,
                        text="Make Dates Unique",
                        command=lambda t=tree: self._dedupe_scraped_headers(t),
                        width=18,
                    )
                    dedupe_button.grid(row=0, column=5, sticky="e", padx=(0, 4))

                    dedupe_items_button = ttk.Button(
                        controls_frame,
                        text="Make Items Unique",
                        command=lambda t=tree: self._dedupe_scraped_items(t),
                        width=18,
                    )
                    dedupe_items_button.grid(row=0, column=6, sticky="e", padx=(0, 4))

                    delete_row_button = ttk.Button(
                        controls_frame,
                        text="Delete Row",
                        command=lambda t=tree: self._delete_scraped_row(t),
                        width=12,
                    )
                    delete_row_button.grid(row=0, column=7, sticky="e", padx=(0, 4))

                    delete_column_button = tk.Menubutton(
                        controls_frame,
                        text="Delete Column",
                        width=14,
                        relief=tk.RAISED,
                        direction="below",
                    )
                    delete_column_button.grid(row=0, column=8, sticky="e")
                    delete_column_menu = tk.Menu(delete_column_button, tearoff=False)
                    delete_column_button.configure(menu=delete_column_menu)

                    info = {
                        "header": list(headings),
                        "rows": table_rows,
                        "csv_path": csv_target_path,
                        "delimiter": csv_delimiter,
                        "scrape_root": scrape_root,
                        "metadata_path": metadata_path,
                        "entry_stem": entry_stem,
                        "category": category,
                        "button": delete_row_button,
                        "scale_buttons": (multiply_button, divide_button),
                        "open_button": open_page_button,
                        "relabel_button": relabel_button,
                        "dedupe_button": dedupe_button,
                        "dedupe_items_button": dedupe_items_button,
                        "delete_column_button": delete_column_button,
                        "delete_column_menu": delete_column_menu,
                        "preview_widget": preview_widget,
                        "entry": entry,
                        "scalable_columns": scalable_columns,
                        "tree": tree,
                    }
                    self.scraped_table_sources[tree] = info
                    if isinstance(preview_widget, tk.Widget):
                        state = self.scraped_preview_states.get(preview_widget)
                        if isinstance(state, dict):
                            state["table_info"] = info
                    self._update_scraped_controls_state(info)

                row_index += 1

        self._refresh_combined_tab(auto_update=True)

    def _scraped_scalable_column_indices(self, headings: List[str]) -> List[int]:
        if not headings:
            return []
        return [index for index in range(len(headings)) if index >= 2]

    def _scraped_item_column_index(self, headings: List[str]) -> Optional[int]:
        for index, value in enumerate(headings):
            if isinstance(value, str) and value.strip().lower() == "item":
                return index
        return None

    def _update_scraped_controls_state(self, info: Dict[str, Any]) -> None:
        rows: List[List[str]] = info.get("rows", [])
        csv_path = info.get("csv_path")
        has_rows = bool(rows)
        has_csv = isinstance(csv_path, Path)
        has_scalable_columns = bool(info.get("scalable_columns"))

        delete_button = info.get("button")
        if isinstance(delete_button, ttk.Button):
            if has_rows and has_csv:
                delete_button.state(["!disabled"])
            else:
                delete_button.state(["disabled"])

        scale_buttons = info.get("scale_buttons")
        if isinstance(scale_buttons, tuple):
            for button in scale_buttons:
                if isinstance(button, ttk.Button):
                    if has_rows and has_csv and has_scalable_columns:
                        button.state(["!disabled"])
                    else:
                        button.state(["disabled"])

        relabel_button = info.get("relabel_button")
        header: List[str] = info.get("header", [])
        if isinstance(relabel_button, ttk.Button):
            if has_csv and len(header) > 2:
                relabel_button.state(["!disabled"])
            else:
                relabel_button.state(["disabled"])

        dedupe_button = info.get("dedupe_button")
        if isinstance(dedupe_button, ttk.Button):
            if has_csv and len(header) > 2:
                dedupe_button.state(["!disabled"])
            else:
                dedupe_button.state(["disabled"])

        dedupe_items_button = info.get("dedupe_items_button")
        if isinstance(dedupe_items_button, ttk.Button):
            if has_csv and has_rows and self._scraped_item_column_index(header) is not None:
                dedupe_items_button.state(["!disabled"])
            else:
                dedupe_items_button.state(["disabled"])

        delete_column_button = info.get("delete_column_button")
        if isinstance(delete_column_button, tk.Menubutton):
            if has_csv and len(header) > 2:
                delete_column_button.configure(state="normal")
            else:
                delete_column_button.configure(state="disabled")
        self._populate_delete_column_menu(info)

        self._update_scraped_open_button(info)

    def _update_scraped_open_button(self, info: Dict[str, Any]) -> None:
        open_button = info.get("open_button")
        if not isinstance(open_button, ttk.Button):
            return
        preview_widget = info.get("preview_widget")
        if not isinstance(preview_widget, tk.Widget):
            open_button.state(["disabled"])
            return
        state = self.scraped_preview_states.get(preview_widget)
        entry = state.get("entry") if isinstance(state, dict) else None
        page_index = state.get("current_page") if isinstance(state, dict) else None
        if isinstance(entry, PDFEntry) and isinstance(page_index, int) and page_index >= 0:
            open_button.state(["!disabled"])
        else:
            open_button.state(["disabled"])

    def _open_scraped_table_preview(self, tree: ttk.Treeview) -> None:
        info = self.scraped_table_sources.get(tree)
        if not isinstance(info, dict):
            return
        preview_widget = info.get("preview_widget")
        if not isinstance(preview_widget, tk.Widget):
            return
        state = self.scraped_preview_states.get(preview_widget)
        if not isinstance(state, dict):
            return
        entry = state.get("entry")
        page_index = state.get("current_page")
        if isinstance(entry, PDFEntry) and isinstance(page_index, int) and page_index >= 0:
            self._open_pdf_page_external(entry, page_index)

    def _prompt_relabel_scraped_dates(self, tree: ttk.Treeview) -> None:
        info = self.scraped_table_sources.get(tree)
        if not isinstance(info, dict):
            return

        header: List[str] = info.get("header", [])
        if len(header) <= 2:
            messagebox.showinfo(
                "Relabel Dates", "This table does not have any date columns to relabel."
            )
            return

        csv_path = info.get("csv_path")
        if not isinstance(csv_path, Path):
            messagebox.showinfo(
                "Relabel Dates",
                "This table is not associated with a CSV file that can be updated.",
            )
            return

        rows: List[List[str]] = info.get("rows", [])
        delimiter: str = info.get("delimiter", ",")

        try:
            dialog = tk.Toplevel(self.root)
        except tk.TclError:
            return

        dialog.title("Relabel Dates")
        dialog.transient(self.root)
        try:
            dialog.grab_set()
        except tk.TclError:
            pass

        content = ttk.Frame(dialog, padding=12)
        content.pack(fill=tk.BOTH, expand=True)
        content.columnconfigure(1, weight=1)

        ttk.Label(
            content,
            text="Update the column labels for the date values and click Save to apply the changes.",
            wraplength=400,
            justify=tk.LEFT,
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 12))

        entry_rows: List[Tuple[int, tk.StringVar, ttk.Entry]] = []
        current_row = 1
        for index in range(2, len(header)):
            label_text = header[index].strip() or f"Column {index + 1}"
            ttk.Label(content, text=f"{label_text}:").grid(
                row=current_row, column=0, sticky="w", padx=(0, 8), pady=2
            )
            var = tk.StringVar(value=header[index])
            entry = ttk.Entry(content, textvariable=var)
            entry.grid(row=current_row, column=1, sticky="ew", pady=2)
            entry_rows.append((index, var, entry))
            current_row += 1

        button_frame = ttk.Frame(dialog, padding=(12, 0, 12, 12))
        button_frame.pack(fill=tk.X)
        button_frame.columnconfigure(0, weight=1)

        def _close_dialog() -> None:
            try:
                dialog.grab_release()
            except tk.TclError:
                pass
            dialog.destroy()

        def _on_cancel() -> None:
            _close_dialog()

        def _on_save() -> None:
            new_header = list(header)
            seen: Set[str] = set()
            for idx, existing in enumerate(new_header):
                if idx < 2:
                    normalized = existing.strip().lower()
                    if normalized:
                        seen.add(normalized)
            for column_index, var, _entry in entry_rows:
                value = var.get().strip()
                if not value:
                    messagebox.showerror(
                        "Relabel Dates", "Column labels cannot be left blank.", parent=dialog
                    )
                    return
                normalized = value.lower()
                if normalized in seen:
                    messagebox.showerror(
                        "Relabel Dates",
                        "Each column label must be unique.",
                        parent=dialog,
                    )
                    return
                seen.add(normalized)
                new_header[column_index] = value

            if not self._write_scraped_csv_rows(
                csv_path, new_header, rows, delimiter=delimiter
            ):
                return

            info["header"] = new_header
            info["scalable_columns"] = self._scraped_scalable_column_indices(new_header)
            columns = list(tree["columns"])
            for column_id, heading_text in zip(columns, new_header):
                tree.heading(column_id, text=heading_text)

            self._update_scraped_controls_state(info)
            self._refresh_combined_tab(auto_update=True)
            _close_dialog()

        ttk.Button(button_frame, text="Cancel", command=_on_cancel).grid(
            row=0, column=1, sticky="e"
        )
        ttk.Button(button_frame, text="Save", command=_on_save).grid(
            row=0, column=2, sticky="e", padx=(8, 0)
        )

        dialog.protocol("WM_DELETE_WINDOW", _on_cancel)
        dialog.bind("<Escape>", lambda _e: _on_cancel())
        dialog.bind("<Return>", lambda _e: _on_save())

        dialog.wait_visibility()
        if entry_rows:
            entry_rows[0][2].focus_set()
        else:
            dialog.focus_set()

    def _populate_delete_column_menu(self, info: Dict[str, Any]) -> None:
        delete_column_button = info.get("delete_column_button")
        delete_column_menu = info.get("delete_column_menu")
        tree = info.get("tree")
        if not isinstance(delete_column_button, tk.Menubutton):
            return
        if not isinstance(delete_column_menu, tk.Menu):
            delete_column_button.configure(state="disabled")
            return
        if not isinstance(tree, ttk.Treeview):
            delete_column_button.configure(state="disabled")
            return

        delete_column_menu.delete(0, tk.END)
        header: List[str] = info.get("header", [])
        deletable_indices = [index for index in range(len(header)) if index >= 2]
        if not deletable_indices:
            delete_column_button.configure(state="disabled")
            return

        for column_index in deletable_indices:
            heading = header[column_index]
            label = heading.strip() or f"Column {column_index + 1}"
            delete_column_menu.add_command(
                label=label,
                command=lambda idx=column_index, t=tree: self._delete_scraped_column(t, idx),
            )

    def _dedupe_scraped_headers(self, tree: ttk.Treeview) -> None:
        info = self.scraped_table_sources.get(tree)
        if not isinstance(info, dict):
            return

        header: List[str] = info.get("header", [])
        if len(header) <= 2:
            messagebox.showinfo(
                "Make Dates Unique", "This table does not have any date columns to update.",
            )
            return

        csv_path = info.get("csv_path")
        if not isinstance(csv_path, Path):
            messagebox.showinfo(
                "Make Dates Unique",
                "This table is not associated with a CSV file that can be updated.",
            )
            return

        rows: List[List[str]] = info.get("rows", [])
        delimiter: str = info.get("delimiter", ",")

        new_header = list(header)
        seen: Dict[str, int] = {}
        duplicates_found = False
        for index, value in enumerate(new_header):
            normalized = value.strip().lower()
            if not normalized:
                continue
            count = seen.get(normalized, 0)
            if count > 0 and index >= 2:
                trimmed = value.strip()
                suffix = f".{count}"
                if not trimmed.endswith(suffix):
                    trimmed = f"{trimmed}{suffix}"
                new_header[index] = trimmed
                duplicates_found = True
            seen[normalized] = count + 1

        if not duplicates_found:
            messagebox.showinfo(
                "Make Dates Unique", "No duplicate column labels were found after the first two columns.",
            )
            return

        if not self._write_scraped_csv_rows(csv_path, new_header, rows, delimiter=delimiter):
            return

        info["header"] = new_header
        info["scalable_columns"] = self._scraped_scalable_column_indices(new_header)
        columns = list(tree["columns"])
        for column_id, heading_text in zip(columns, new_header):
            tree.heading(column_id, text=heading_text)

        self._update_scraped_controls_state(info)
        self._refresh_combined_tab(auto_update=True)
        messagebox.showinfo(
            "Make Dates Unique", "Duplicate column labels have been updated.",
        )

    def _dedupe_scraped_items(self, tree: ttk.Treeview) -> None:
        info = self.scraped_table_sources.get(tree)
        if not isinstance(info, dict):
            return

        header: List[str] = info.get("header", [])
        if not header:
            messagebox.showinfo(
                "Make Items Unique", "This table does not have any columns to update.",
            )
            return

        item_index = self._scraped_item_column_index(header)
        if item_index is None:
            messagebox.showinfo(
                "Make Items Unique", "No 'ITEM' column was found in this table.",
            )
            return

        csv_path = info.get("csv_path")
        if not isinstance(csv_path, Path):
            messagebox.showinfo(
                "Make Items Unique",
                "This table is not associated with a CSV file that can be updated.",
            )
            return

        rows: List[List[str]] = info.get("rows", [])
        if not rows:
            messagebox.showinfo(
                "Make Items Unique", "There are no rows available to update.",
            )
            return

        updated_rows: List[List[str]] = []
        seen: Dict[str, int] = {}
        duplicates_found = False
        for row in rows:
            values = list(row)
            if item_index >= len(values):
                values.extend([""] * (item_index - len(values) + 1))
            raw_value = values[item_index].strip()
            if raw_value:
                base_value = re.sub(r"\.\d+$", "", raw_value)
                normalized = base_value.lower()
                count = seen.get(normalized, 0)
                if count == 0:
                    if base_value != values[item_index]:
                        values[item_index] = base_value
                else:
                    values[item_index] = f"{base_value}.{count}"
                    duplicates_found = True
                seen[normalized] = count + 1
            updated_rows.append(values)

        if not duplicates_found:
            messagebox.showinfo(
                "Make Items Unique", "No duplicate ITEM entries were found to update.",
            )
            return

        delimiter: str = info.get("delimiter", ",")
        if not self._write_scraped_csv_rows(csv_path, header, updated_rows, delimiter=delimiter):
            return

        info["rows"] = updated_rows
        self._refresh_scraped_tree_display(tree, info)
        self._update_scraped_controls_state(info)
        self._refresh_combined_tab(auto_update=True)
        messagebox.showinfo(
            "Make Items Unique", "Duplicate ITEM entries have been updated.",
        )

    def _delete_scraped_column(self, tree: ttk.Treeview, column_index: int) -> None:
        info = self.scraped_table_sources.get(tree)
        if not isinstance(info, dict):
            return

        header: List[str] = info.get("header", [])
        if not header or not (0 <= column_index < len(header)):
            return

        if column_index < 2:
            messagebox.showinfo(
                "Delete Column",
                "The first two columns cannot be deleted from this table.",
            )
            return

        csv_path = info.get("csv_path")
        if not isinstance(csv_path, Path):
            messagebox.showinfo(
                "Delete Column",
                "This table is not associated with a CSV file that can be updated.",
            )
            return

        column_label = header[column_index].strip() or f"Column {column_index + 1}"
        confirm = messagebox.askyesno(
            "Delete Column", f"Delete column '{column_label}' from the table?", icon="warning"
        )
        if not confirm:
            return

        rows: List[List[str]] = info.get("rows", [])
        updated_rows: List[List[str]] = []
        for row in rows:
            values = list(row)
            if 0 <= column_index < len(values):
                values.pop(column_index)
            updated_rows.append(values)

        new_header = [value for idx, value in enumerate(header) if idx != column_index]
        delimiter: str = info.get("delimiter", ",")

        if not self._write_scraped_csv_rows(csv_path, new_header, updated_rows, delimiter=delimiter):
            return

        info["header"] = new_header
        info["rows"] = updated_rows
        info["scalable_columns"] = self._scraped_scalable_column_indices(new_header)

        columns = list(tree["columns"])
        if 0 <= column_index < len(columns):
            columns.pop(column_index)
        tree.configure(columns=columns)
        for column_id, heading_text in zip(columns, new_header):
            tree.heading(column_id, text=heading_text)
            tree.column(column_id, anchor="center", stretch=True, width=120)

        self._refresh_scraped_tree_display(tree, info)
        self._update_scraped_controls_state(info)
        self._refresh_combined_tab(auto_update=True)

    def _open_pdf_page_external(self, entry: PDFEntry, page_index: int) -> None:
        if fitz is None:
            messagebox.showwarning("Open Page", PYMUPDF_REQUIRED_MESSAGE)
            return

        if page_index < 0 or page_index >= len(entry.doc):
            messagebox.showinfo(
                "Open Page", "The requested page could not be found in the PDF document.",
            )
            return

        try:
            temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf", prefix="scraped_page_")
        except Exception as exc:
            messagebox.showwarning("Open Page", f"Could not create a temporary file: {exc}")
            return

        temp_path = Path(temp_file.name)
        temp_file.close()

        new_doc: Optional[fitz.Document] = None
        try:
            new_doc = fitz.open()
            new_doc.insert_pdf(entry.doc, from_page=page_index, to_page=page_index)
            new_doc.save(str(temp_path))
        except Exception as exc:
            messagebox.showwarning("Open Page", f"Could not extract PDF page: {exc}")
            temp_path.unlink(missing_ok=True)
            return
        finally:
            if new_doc is not None:
                new_doc.close()

        self._open_pdf(temp_path)

    def _refresh_scraped_tree_display(self, tree: ttk.Treeview, info: Dict[str, Any]) -> None:
        for item in tree.get_children():
            tree.delete(item)
        columns = list(tree["columns"])
        data_rows: List[List[str]] = info.get("rows", [])
        for row in data_rows:
            values = list(row)
            if len(values) < len(columns):
                values.extend([""] * (len(columns) - len(values)))
            elif len(values) > len(columns):
                values = values[: len(columns)]
            display_values = [self._format_table_value(value) for value in values]
            tree.insert("", tk.END, values=display_values)
        tree.configure(height=min(15, max(3, len(data_rows))))

    def _scale_scraped_table(self, tree: ttk.Treeview, factor: float) -> None:
        info = self.scraped_table_sources.get(tree)
        if not info:
            return

        csv_path = info.get("csv_path")
        if not isinstance(csv_path, Path):
            messagebox.showinfo(
                "Scale Values", "This table is not associated with a CSV file that can be updated."
            )
            return

        scalable_columns: List[int] = info.get("scalable_columns", [])
        if not scalable_columns:
            messagebox.showinfo(
                "Scale Values",
                "No columns beyond the first two are available for scaling in this table.",
            )
            return

        data_rows: List[List[str]] = info.get("rows", [])
        if not data_rows:
            messagebox.showinfo("Scale Values", "There are no rows available to scale.")
            return

        header: List[str] = info.get("header", [])
        header_length = len(header)
        delimiter: str = info.get("delimiter", ",")

        updated_rows: List[List[str]] = []
        changed = False
        for original_row in data_rows:
            values = list(original_row)
            if len(values) < header_length:
                values.extend([""] * (header_length - len(values)))
            elif len(values) > header_length:
                values = values[:header_length]
            for column_index in scalable_columns:
                if 0 <= column_index < len(values):
                    new_value = self._scale_scraped_cell_value(values[column_index], factor)
                    if new_value is not None and new_value != values[column_index]:
                        values[column_index] = new_value
                        changed = True
            updated_rows.append(values)

        if not changed:
            messagebox.showinfo(
                "Scale Values", "No numeric values were updated for this table."
            )
            return

        if not self._write_scraped_csv_rows(csv_path, header, updated_rows, delimiter=delimiter):
            scrape_root: Optional[Path] = info.get("scrape_root")
            metadata_path: Optional[Path] = info.get("metadata_path")
            entry_stem = info.get("entry_stem")
            category: Optional[str] = info.get("category")
            if (
                isinstance(scrape_root, Path)
                and isinstance(metadata_path, Path)
                and isinstance(entry_stem, str)
                and isinstance(category, str)
            ):
                self._reload_scraped_csv(scrape_root, metadata_path, entry_stem, category)
            return

        info["rows"] = updated_rows
        self._refresh_scraped_tree_display(tree, info)
        self._update_scraped_controls_state(info)
        self._refresh_combined_tab(auto_update=True)

    def _scale_scraped_cell_value(self, original: str, factor: float) -> Optional[str]:
        numeric_value = self._parse_numeric_value(original)
        if numeric_value is None:
            return None

        scaled_value = numeric_value * factor
        stripped = original.strip()
        if not stripped:
            return None

        leading_ws_match = re.match(r"^\s*", original)
        trailing_ws_match = re.search(r"\s*$", original)
        leading_ws = leading_ws_match.group(0) if leading_ws_match else ""
        trailing_ws = trailing_ws_match.group(0) if trailing_ws_match else ""

        has_percent = stripped.endswith("%")
        if has_percent:
            stripped = stripped[:-1].rstrip()

        used_parentheses = stripped.startswith("(") and stripped.endswith(")")
        if used_parentheses:
            stripped = stripped[1:-1].strip()

        if stripped.startswith("-") or stripped.startswith("+"):
            stripped = stripped[1:].lstrip()

        prefix_match = re.match(r"^[^\d]*", stripped)
        prefix = prefix_match.group(0) if prefix_match else ""
        remainder = stripped[len(prefix) :]
        if not remainder:
            return None

        number_match = re.match(
            r"[-+]?(?:(?:\d[\d,]*)(?:\.\d+)?|\.\d+)(?:[eE][-+]?\d+)?",
            remainder,
        )
        if not number_match:
            formatted_number = f"{abs(scaled_value):.6g}"
            formatted = f"{prefix}{formatted_number}"
            if scaled_value < 0:
                if used_parentheses:
                    formatted = f"({formatted})"
                else:
                    formatted = f"-{formatted}"
            if has_percent:
                formatted = f"{formatted}%"
            return f"{leading_ws}{formatted}{trailing_ws}"

        number_segment = number_match.group(0)
        suffix_extra = remainder[len(number_segment) :]
        normalized_segment = number_segment.replace(",", "")
        normalized_segment = re.sub(r"[eE][-+]?\d+", "", normalized_segment)
        decimal_match = re.search(r"\.(\d+)", normalized_segment)
        decimal_places = len(decimal_match.group(1)) if decimal_match else 0

        absolute_value = abs(scaled_value)
        if decimal_places > 0:
            formatted_number = f"{absolute_value:,.{decimal_places}f}"
        else:
            formatted_number = f"{absolute_value:,.0f}"

        formatted = f"{prefix}{formatted_number}{suffix_extra}"
        if scaled_value < 0:
            if used_parentheses:
                formatted = f"({formatted})"
            else:
                formatted = f"-{formatted}"

        if has_percent:
            formatted = f"{formatted}%"

        return f"{leading_ws}{formatted}{trailing_ws}"

    def _clear_combined_tab(self) -> None:
        if not hasattr(self, "combined_header_frame"):
            return
        for child in self.combined_header_frame.winfo_children():
            child.destroy()
        for child in self.combined_result_frame.winfo_children():
            child.destroy()
        self._destroy_note_editor()
        self.combined_csv_sources.clear()
        self.combined_pdf_order = []
        self.combined_result_tree = None
        if hasattr(self, "combine_confirm_button"):
            self.combine_confirm_button.state(["disabled"])
        self.combined_max_data_columns = 0
        self.combined_note_record_keys.clear()
        self.combined_note_column_id = None
        self.combined_ordered_columns = []
        self.combined_all_records = []
        self.combined_record_lookup = {}
        self.combined_column_defaults.clear()
        self.combined_labels_by_pdf.clear()
        self.combined_column_name_map.clear()
        self.combined_row_sources.clear()
        self.combined_preview_frame = None
        self.combined_preview_canvas = None
        self.combined_preview_canvas_image = None
        self.combined_preview_image = None
        self.combined_split_pane = None
        self.combined_save_button = None
        self.combined_preview_detail_var.set("Select a row to view the PDF page.")
        self.combined_header_label_widgets.clear()
        self._refresh_chart_tab()

    def _refresh_chart_tab(self) -> None:
        plot_frame = self.chart_plot_frame
        if plot_frame is None:
            return
        plot_frame.clear_companies()
        if not self.combined_all_records or not self.combined_ordered_columns:
            return
        company = self.company_var.get().strip()
        if not company:
            return
        self._export_final_combined_csv()
        combined_path = self.companies_dir / company / "combined.csv"
        try:
            dataset = FinanceDataset(combined_path)
        except (FileNotFoundError, ValueError) as exc:
            logger.info("Chart dataset unavailable for %s: %s", company, exc)
            return
        if not dataset.has_data():
            logger.info(
                "Combined CSV for %s does not contain Finance or Income data for chart.",
                company,
            )
            return
        try:
            _, warning = plot_frame.add_company(company, dataset)
        except ValueError as exc:
            logger.warning("Unable to render chart for %s: %s", company, exc)
            return
        except Exception:
            logger.exception("Unexpected error while rendering chart for %s", company)
            return
        if warning:
            logger.warning("Chart normalization warning for %s: %s", company, warning)
        plot_frame.set_display_mode(FinancePlotFrame.MODE_STACKED)
        plot_frame.set_normalization_mode(FinanceDataset.NORMALIZATION_REPORTED)

    def _refresh_combined_tab(self, *, auto_update: bool = False) -> None:
        if not hasattr(self, "combined_header_frame"):
            return
        self._clear_combined_tab()
        company = self.company_var.get()
        if not company:
            ttk.Label(
                self.combined_header_frame,
                text="Select a company and run AIScrape to review combined data.",
            ).grid(row=0, column=0, sticky="w", padx=8, pady=8)
            return
        if not self.pdf_entries:
            ttk.Label(
                self.combined_header_frame,
                text="Load PDFs to review scraped output headers.",
            ).grid(row=0, column=0, sticky="w", padx=8, pady=8)
            return
        scrape_root = self.companies_dir / company / "openapiscrape"
        if not scrape_root.exists():
            ttk.Label(
                self.combined_header_frame,
                text="Run AIScrape to populate scraped results for combination.",
            ).grid(row=0, column=0, sticky="w", padx=8, pady=8)
            return

        type_header_map: Dict[str, Dict[Path, Dict[str, Any]]] = {}
        type_heading_labels: Dict[str, Tuple[str, str]] = {}
        pdf_entries_with_data: List[Tuple[PDFEntry, int, int]] = []
        max_data_columns = 0

        for entry_position, entry in enumerate(self.pdf_entries):
            entry_stem = entry.stem
            metadata_path = self._metadata_file_for_entry(scrape_root, entry_stem)
            metadata: Dict[str, Any] = {}
            if metadata_path.exists():
                metadata = self._load_doc_metadata(metadata_path)
            else:
                legacy_dir = self._legacy_doc_dir(scrape_root, entry_stem)
                if legacy_dir.exists():
                    metadata = self._load_doc_metadata(legacy_dir)
            if not metadata:
                continue
            entry_has_data = False
            min_page_index: Optional[int] = None
            for category in COLUMNS:
                meta = metadata.get(category)
                if not isinstance(meta, dict):
                    continue
                page_value = meta.get("page_index")
                page_index: Optional[int]
                try:
                    page_index = int(page_value)
                except (TypeError, ValueError):
                    try:
                        page_index = int(float(page_value))
                    except (TypeError, ValueError):
                        page_index = None
                if page_index is not None:
                    if min_page_index is None or page_index < min_page_index:
                        min_page_index = page_index
                rows: List[List[str]] = []
                txt_name = meta.get("txt")
                if isinstance(txt_name, str) and txt_name:
                    txt_path = self._resolve_scrape_path(scrape_root, entry_stem, txt_name)
                    if txt_path.exists():
                        try:
                            text = txt_path.read_text(encoding="utf-8")
                        except Exception:
                            text = ""
                        rows = self._convert_response_to_rows(text)
                if not rows:
                    csv_name = meta.get("csv")
                    if isinstance(csv_name, str) and csv_name:
                        csv_path = self._resolve_scrape_path(scrape_root, entry_stem, csv_name)
                        if csv_path.exists():
                            rows = self._read_csv_rows(csv_path)
                if not rows:
                    continue
                headings_row = rows[0]
                if not headings_row:
                    continue
                category_index: Optional[int] = None
                item_index: Optional[int] = None
                data_indices: List[int] = []
                display_headers: List[str] = []
                for idx, heading in enumerate(headings_row):
                    normalized = heading.strip().lower()
                    if normalized == "category":
                        category_index = idx
                    elif normalized == "item":
                        item_index = idx
                    else:
                        data_indices.append(idx)
                        display_headers.append(heading.strip() or f"Column {idx + 1}")
                if category_index is None or item_index is None:
                    continue
                data_rows = rows[1:] if len(rows) > 1 else []
                category_heading_text = (
                    headings_row[category_index] if category_index < len(headings_row) else "Category"
                )
                item_heading_text = (
                    headings_row[item_index] if item_index < len(headings_row) else "Item"
                )
                logger.info(
                    "Combined header capture for %s | %s -> indices=%s headers=%s",
                    entry.path.name,
                    category,
                    data_indices,
                    display_headers,
                )
                self.combined_csv_sources[(entry.path, category)] = {
                    "rows": data_rows,
                    "category_index": category_index,
                    "item_index": item_index,
                    "data_indices": data_indices,
                    "headings": [headings_row[i] if i < len(headings_row) else "" for i in data_indices],
                    "category_heading": category_heading_text,
                    "item_heading": item_heading_text,
                }
                header_entry = type_header_map.setdefault(category, {}).setdefault(entry.path, {})
                header_entry.update(
                    {
                        "headers": display_headers[:],
                        "category_heading": category_heading_text,
                        "item_heading": item_heading_text,
                    }
                )
                if category not in type_heading_labels:
                    type_heading_labels[category] = (category_heading_text, item_heading_text)
                for idx, header_text in enumerate(display_headers):
                    key = (entry.path, idx)
                    if header_text.strip():
                        self.combined_column_defaults.setdefault(key, header_text)
                max_data_columns = max(max_data_columns, len(display_headers))
                entry_has_data = True
            if entry_has_data:
                normalized_page = min_page_index if min_page_index is not None else sys.maxsize
                pdf_entries_with_data.append((entry, normalized_page, entry_position))

        if not pdf_entries_with_data:
            ttk.Label(
                self.combined_header_frame,
                text="No scraped results available to combine.",
            ).grid(row=0, column=0, sticky="w", padx=8, pady=8)
            return

        pdf_entries_with_data.sort(key=lambda item: (item[1], item[2], item[0].path.name.lower()))
        self.combined_pdf_order = [entry.path for entry, _page, _pos in pdf_entries_with_data]
        self.combined_max_data_columns = max_data_columns
        logger.info(
            "Combined PDF order: %s with max data columns=%d",
            [path.name for path in self.combined_pdf_order],
            self.combined_max_data_columns,
        )

        desired_keys = set()
        for path in self.combined_pdf_order:
            for idx in range(max_data_columns):
                key = (path, idx)
                desired_keys.add(key)
                default_text = self.combined_column_defaults.get(key, "")
                if not default_text:
                    default_text = f"Value {idx + 1}"
                var = self.combined_column_label_vars.get(key)
                if var is None:
                    var = tk.StringVar(master=self.root, value=default_text)
                    self.combined_column_label_vars[key] = var
                elif not var.get().strip():
                    var.set(default_text)
                if var is not None and key not in self.combined_column_label_traces:
                    self._register_combined_label_trace(key, var)
        for stale_key in list(self.combined_column_label_vars.keys()):
            if stale_key not in desired_keys:
                self._remove_combined_label_trace(stale_key)
                self.combined_column_label_vars.pop(stale_key, None)
                self.combined_header_label_widgets.pop(stale_key, None)

        displayed_categories: List[str] = []
        for category in COLUMNS:
            header_by_pdf = type_header_map.get(category, {})
            if header_by_pdf or category in type_heading_labels:
                displayed_categories.append(category)

        if not displayed_categories:
            displayed_categories = COLUMNS[:]

        total_columns = 1 + len(displayed_categories)
        logger.info(
            "Combined label grid layout -> rows for %d PDFs and %d categories",
            len(self.combined_pdf_order),
            len(displayed_categories),
        )
        for column_index in range(total_columns):
            weight = 1 if column_index > 0 else 0
            self.combined_header_frame.columnconfigure(column_index, weight=weight)

        default_category_heading = "Category"
        default_item_heading = "Item"
        for category in COLUMNS:
            headings = type_heading_labels.get(category)
            if headings:
                if headings[0].strip():
                    default_category_heading = headings[0]
                if headings[1].strip():
                    default_item_heading = headings[1]
                break

        category_heading_display: Dict[str, str] = {}
        item_heading_display: Dict[str, str] = {}
        for category in displayed_categories:
            heading_values = type_heading_labels.get(category)
            category_heading_text = default_category_heading
            item_heading_text = default_item_heading
            if heading_values:
                if heading_values[0].strip():
                    category_heading_text = heading_values[0]
                if heading_values[1].strip():
                    item_heading_text = heading_values[1]
            else:
                header_by_pdf = type_header_map.get(category, {})
                for entry_meta in header_by_pdf.values():
                    category_heading_text = entry_meta.get("category_heading", category_heading_text)
                    item_heading_text = entry_meta.get("item_heading", item_heading_text)
                    break
            category_heading_display[category] = category_heading_text
            item_heading_display[category] = item_heading_text

        current_row = 0
        ttk.Label(
            self.combined_header_frame,
            text="",
        ).grid(row=current_row, column=0, padx=4, pady=(0, 4), sticky="w")
        for column_index, category in enumerate(displayed_categories, start=1):
            ttk.Label(
                self.combined_header_frame,
                text=category,
                font=("TkDefaultFont", 10, "bold"),
            ).grid(row=current_row, column=column_index, padx=4, pady=(0, 4), sticky="w")

        current_row += 1
        ttk.Label(
            self.combined_header_frame,
            text="Type",
            font=("TkDefaultFont", 10, "bold"),
        ).grid(row=current_row, column=0, padx=4, pady=4, sticky="w")
        for column_index, category in enumerate(displayed_categories, start=1):
            ttk.Label(
                self.combined_header_frame,
                text=category,
            ).grid(row=current_row, column=column_index, padx=4, pady=4, sticky="w")

        current_row += 1
        ttk.Label(
            self.combined_header_frame,
            text=default_category_heading,
            font=("TkDefaultFont", 10, "bold"),
        ).grid(row=current_row, column=0, padx=4, pady=4, sticky="w")
        for column_index, category in enumerate(displayed_categories, start=1):
            ttk.Label(
                self.combined_header_frame,
                text=category_heading_display.get(category, ""),
            ).grid(row=current_row, column=column_index, padx=4, pady=4, sticky="w")

        current_row += 1
        ttk.Label(
            self.combined_header_frame,
            text=default_item_heading,
            font=("TkDefaultFont", 10, "bold"),
        ).grid(row=current_row, column=0, padx=4, pady=4, sticky="w")
        for column_index, category in enumerate(displayed_categories, start=1):
            ttk.Label(
                self.combined_header_frame,
                text=item_heading_display.get(category, ""),
            ).grid(row=current_row, column=column_index, padx=4, pady=4, sticky="w")

        if max_data_columns:
            for pdf_path in self.combined_pdf_order:
                pdf_label = pdf_path.stem
                for idx in range(max_data_columns):
                    current_row += 1
                    key = (pdf_path, idx)
                    var = self.combined_column_label_vars.get(key)
                    self.combined_header_label_widgets[key] = []
                    row_container = ttk.Frame(self.combined_header_frame)
                    row_container.grid(row=current_row, column=0, padx=4, pady=4, sticky="nsew")
                    ttk.Label(
                        row_container,
                        text=pdf_label,
                        font=("TkDefaultFont", 10, "bold"),
                    ).pack(anchor="w")
                    if var is not None:
                        entry = ttk.Entry(row_container, textvariable=var, width=22)
                        entry.pack(anchor="w", fill=tk.X, pady=(2, 0))
                    else:
                        ttk.Label(row_container, text="").pack(anchor="w")
                    for column_index, category in enumerate(displayed_categories, start=1):
                        headers = type_header_map.get(category, {}).get(pdf_path, {}).get("headers", [])
                        header_text = headers[idx] if idx < len(headers) else ""
                        label_widget = ttk.Label(
                            self.combined_header_frame,
                            text=header_text,
                        )
                        label_widget.grid(
                            row=current_row, column=column_index, padx=4, pady=4, sticky="w"
                        )
                        self.combined_header_label_widgets[key].append(
                            (label_widget, header_text.strip())
                        )
                    self._update_combined_header_label_styles(key)

        if hasattr(self, "combine_confirm_button"):
            self.combine_confirm_button.state(["!disabled"])

        if auto_update and self.combined_pdf_order and self.combined_csv_sources and self.combined_result_tree is None:
            self._confirm_combined_table(auto=True)

    def _register_combined_label_trace(
        self, key: Tuple[Path, int], var: tk.StringVar
    ) -> None:
        self._remove_combined_label_trace(key)

        def _on_change(*_args: Any, target_key: Tuple[Path, int] = key) -> None:
            self._update_combined_header_label_styles(target_key)

        trace_id = var.trace_add("write", _on_change)
        self.combined_column_label_traces[key] = trace_id

    def _remove_combined_label_trace(self, key: Tuple[Path, int]) -> None:
        trace_id = self.combined_column_label_traces.pop(key, None)
        if not trace_id:
            return
        var = self.combined_column_label_vars.get(key)
        if var is None:
            return
        try:
            var.trace_remove("write", trace_id)
        except tk.TclError:
            pass

    def _update_combined_header_label_styles(self, key: Tuple[Path, int]) -> None:
        label_entries = self.combined_header_label_widgets.get(key)
        if not label_entries:
            return
        var = self.combined_column_label_vars.get(key)
        assigned_value = ""
        if var is not None:
            assigned_value = str(var.get() or "").strip()
        for widget, source_value in label_entries:
            normalized_source = source_value.strip()
            if assigned_value != normalized_source:
                try:
                    widget.configure(font=self.combined_header_label_bold_font)
                except tk.TclError:
                    continue
            else:
                try:
                    widget.configure(font=self.combined_header_label_font)
                except tk.TclError:
                    continue

    def _filter_combined_records(self, records: List[Dict[str, str]]) -> List[Dict[str, str]]:
        if not records:
            return []
        # The combined table currently displays all rows regardless of the
        # blank-note iteration toggle; navigation handles the blank-only
        # workflow.
        return records

    def _normalize_type_item_category_key(
        self, type_value: str, item_value: str, category_value: str
    ) -> Tuple[str, str, str]:
        return (
            str(type_value or "").strip().casefold(),
            str(item_value or "").strip().casefold(),
            str(category_value or "").strip().casefold(),
        )

    def _normalize_type_category_key(self, type_value: str, category_value: str) -> Tuple[str, str]:
        return (
            str(type_value or "").strip().casefold(),
            str(category_value or "").strip().casefold(),
        )

    def _load_type_category_entries(self) -> List[Tuple[str, str]]:
        path = getattr(self, "type_category_sort_order_path", None)
        if path is None:
            return []
        entries: List[Tuple[str, str]] = []
        try:
            with path.open("r", encoding="utf-8-sig", newline="") as fh:
                reader = csv.reader(fh)
                header_parsed = False
                type_idx: Optional[int] = None
                category_idx: Optional[int] = None
                for row in reader:
                    if not row:
                        continue
                    trimmed = [cell.strip() for cell in row]
                    if not any(trimmed):
                        continue
                    if not header_parsed:
                        lowered = [cell.lower() for cell in trimmed]
                        try:
                            type_idx = lowered.index("type")
                            category_idx = lowered.index("category")
                        except ValueError:
                            type_idx = 0 if len(trimmed) > 0 else None
                            category_idx = 1 if len(trimmed) > 1 else None
                            header_parsed = True
                        else:
                            header_parsed = True
                            continue
                    if type_idx is None or category_idx is None:
                        continue
                    if type_idx >= len(trimmed) or category_idx >= len(trimmed):
                        continue
                    type_value = trimmed[type_idx]
                    category_value = trimmed[category_idx]
                    if not (type_value or category_value):
                        continue
                    entries.append((type_value, category_value))
        except FileNotFoundError:
            return []
        except Exception as exc:
            logger.warning("Could not read %s: %s", path, exc)
        return entries

    def _ensure_combined_order_entries(self, records: List[Dict[str, str]]) -> None:
        if not records:
            return
        path = getattr(self, "combined_order_path", None)
        if path is None:
            return
        unique_records: List[Tuple[str, str]] = []
        seen_keys: Set[Tuple[str, str]] = set()
        for record in records:
            type_value = str(record.get("Type", "") or "").strip()
            category_value = str(record.get("Category", "") or "").strip()
            if not type_value or not category_value:
                continue
            normalized = self._normalize_type_category_key(type_value, category_value)
            if normalized in seen_keys:
                continue
            seen_keys.add(normalized)
            unique_records.append((type_value, category_value))
        if not unique_records:
            return
        existing_entries = self._load_type_category_entries()
        existing_keys = {
            self._normalize_type_category_key(type_value, category_value)
            for type_value, category_value in existing_entries
        }
        new_entries = [
            entry
            for entry in unique_records
            if self._normalize_type_category_key(*entry) not in existing_keys
        ]
        try:
            should_initialize = not path.exists() or path.stat().st_size == 0
        except OSError:
            should_initialize = False
        if should_initialize:
            entries_to_write = unique_records
            mode = "w"
        else:
            entries_to_write = new_entries
            mode = "a"
        if not entries_to_write:
            return
        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            with path.open(mode, encoding="utf-8", newline="") as fh:
                writer = csv.writer(fh, quoting=csv.QUOTE_ALL)
                if should_initialize:
                    writer.writerow(["Type", "Category"])
                for type_value, category_value in entries_to_write:
                    writer.writerow([type_value, category_value])
            logger.info(
                "Recorded %d Type/Category combination(s) in %s",
                len(entries_to_write),
                path,
            )
        except Exception as exc:
            logger.warning("Could not update %s: %s", path, exc)

    def _load_type_category_order_map(self) -> Dict[Tuple[str, str], int]:
        entries = self._load_type_category_entries()
        order: Dict[Tuple[str, str], int] = {}
        for index, (type_value, category_value) in enumerate(entries):
            key = self._normalize_type_category_key(type_value, category_value)
            if key not in order:
                order[key] = index
        return order

    def _load_type_item_category_entries(self) -> List[Tuple[str, str, str]]:
        path = getattr(self, "type_item_category_path", None)
        if path is None:
            return []
        entries: List[Tuple[str, str, str]] = []
        try:
            with path.open("r", encoding="utf-8-sig", newline="") as fh:
                reader = csv.reader(fh)
                header_parsed = False
                type_idx: Optional[int] = None
                item_idx: Optional[int] = None
                category_idx: Optional[int] = None
                for row in reader:
                    if not row:
                        continue
                    trimmed = [cell.strip() for cell in row]
                    if not any(trimmed):
                        continue
                    if not header_parsed:
                        lowered = [cell.lower() for cell in trimmed]
                        try:
                            type_idx = lowered.index("type")
                            item_idx = lowered.index("item")
                            category_idx = lowered.index("category")
                        except ValueError:
                            type_idx = 0 if len(trimmed) > 0 else None
                            item_idx = 1 if len(trimmed) > 1 else None
                            category_idx = 2 if len(trimmed) > 2 else None
                            header_parsed = True
                        else:
                            header_parsed = True
                            continue
                    if type_idx is None or item_idx is None or category_idx is None:
                        continue
                    if (
                        type_idx >= len(trimmed)
                        or item_idx >= len(trimmed)
                        or category_idx >= len(trimmed)
                    ):
                        continue
                    type_value = trimmed[type_idx]
                    item_value = trimmed[item_idx]
                    category_value = trimmed[category_idx]
                    if not (type_value or item_value or category_value):
                        continue
                    entries.append((type_value, item_value, category_value))
        except FileNotFoundError:
            return []
        except Exception as exc:
            logger.warning("Could not read %s: %s", path, exc)
        return entries

    def _load_type_item_category_order_map(self) -> Dict[Tuple[str, str, str], int]:
        entries = self._load_type_item_category_entries()
        order: Dict[Tuple[str, str, str], int] = {}
        for index, (type_value, item_value, category_value) in enumerate(entries):
            key = self._normalize_type_item_category_key(type_value, item_value, category_value)
            if key not in order:
                order[key] = index
        return order

    def _sort_combined_records(self, records: List[Dict[str, str]]) -> None:
        if not records:
            return
        type_category_order_map = self._load_type_category_order_map()
        type_item_order_map = self._load_type_item_category_order_map()
        type_priority = {name: index for index, name in enumerate(COLUMNS)}
        max_type_priority = len(type_priority)

        def _sort_key(record: Dict[str, str]) -> Tuple[Any, ...]:
            type_value = str(record.get("Type", "") or "").strip()
            category_value = str(record.get("Category", "") or "").strip()
            item_value = str(record.get("Item", "") or "").strip()
            normalized_full = self._normalize_type_item_category_key(
                type_value, item_value, category_value
            )
            normalized_category = self._normalize_type_category_key(
                type_value, category_value
            )
            category_index = type_category_order_map.get(normalized_category)
            item_index = type_item_order_map.get(normalized_full)
            fallback_strings = (
                category_value.casefold(),
                item_value.casefold(),
                type_value.casefold(),
            )
            if category_index is not None:
                return (
                    0,
                    category_index,
                    0 if item_index is not None else 1,
                    item_index if item_index is not None else 0,
                    fallback_strings,
                )
            if item_index is not None:
                return (1, item_index, fallback_strings)
            return (
                2,
                type_priority.get(type_value, max_type_priority),
                category_value.casefold(),
                item_value.casefold(),
                type_value.casefold(),
            )

        records.sort(key=_sort_key)

    def _parse_combined_period_label(self, label: str) -> Optional[date]:
        text = str(label).strip()
        if not text:
            return None
        formats = [
            "%d.%m.%Y",
            "%Y-%m-%d",
            "%m/%d/%Y",
            "%d/%m/%Y",
            "%Y/%m/%d",
            "%Y.%m.%d",
            "%b %d %Y",
            "%B %d %Y",
            "%b %Y",
            "%B %Y",
            "%Y",
        ]
        for fmt in formats:
            try:
                parsed = datetime.strptime(text, fmt)
            except ValueError:
                continue
            if fmt in {"%b %Y", "%B %Y"}:
                return date(parsed.year, parsed.month, 1)
            if fmt == "%Y":
                return date(parsed.year, 1, 1)
            return parsed.date()
        forward_quarter = re.match(r"(?i)^Q([1-4])\s*(\d{4})$", text)
        if forward_quarter:
            quarter = int(forward_quarter.group(1))
            year = int(forward_quarter.group(2))
            month = (quarter - 1) * 3 + 1
            return date(year, month, 1)
        reverse_quarter = re.match(r"(?i)^(\d{4})\s*Q([1-4])$", text)
        if reverse_quarter:
            year = int(reverse_quarter.group(1))
            quarter = int(reverse_quarter.group(2))
            month = (quarter - 1) * 3 + 1
            return date(year, month, 1)
        fiscal_match = re.match(r"(?i)^FY\s*(\d{4})$", text)
        if fiscal_match:
            year = int(fiscal_match.group(1))
            return date(year, 1, 1)
        return None

    def _sort_combined_period_columns(
        self,
        columns: List[str],
        column_name_map: Dict[str, Tuple[Path, str]],
    ) -> List[str]:
        if not columns:
            return columns
        base_columns_order = [
            column for column in columns if column in {"Type", "Category", "Item", "Note"}
        ]
        data_columns = [
            column for column in columns if column not in {"Type", "Category", "Item", "Note"}
        ]
        if not data_columns:
            return base_columns_order
        pdf_priority = {path: index for index, path in enumerate(self.combined_pdf_order)}
        decorated: List[Tuple[int, Optional[date], int, int, str]] = []
        for index, column_name in enumerate(data_columns):
            mapping = column_name_map.get(column_name)
            label_value = mapping[1] if mapping else column_name
            parsed_date = self._parse_combined_period_label(label_value)
            pdf_index = pdf_priority.get(mapping[0], len(pdf_priority)) if mapping else len(pdf_priority)
            sort_group = 0 if parsed_date is not None else 1
            decorated.append((sort_group, parsed_date, pdf_index, index, column_name))
        decorated.sort(
            key=lambda item: (
                item[0],
                item[1] or date.max,
                item[2],
                item[3],
            )
        )
        sorted_data_columns = [item[4] for item in decorated]
        return base_columns_order + sorted_data_columns

    def _apply_reference_sort_to_combined(self) -> None:
        if not self.combined_all_records:
            return
        self._sort_combined_records(self.combined_all_records)
        if self.combined_ordered_columns:
            self._update_combined_tree_display()

    def _on_resort_combined_clicked(self) -> None:
        if not self.combined_all_records:
            messagebox.showinfo(
                "Combined Order",
                "Build the combined table before sorting by the combined order file.",
            )
            return
        self._ensure_combined_order_entries(self.combined_all_records)
        self._apply_reference_sort_to_combined()

    def _generate_type_category_sort_order_csv(self) -> None:
        if not self.combined_all_records:
            messagebox.showinfo(
                "Type/Category Sort Order",
                "Load a combined table before generating the Type/Category sort order.",
            )
            return

        seen_keys: Set[Tuple[str, str]] = set()
        unique_records: List[Tuple[str, str]] = []
        for record in self.combined_all_records:
            type_value = str(record.get("Type", "") or "").strip()
            category_value = str(record.get("Category", "") or "").strip()
            if not type_value or not category_value:
                continue
            normalized = self._normalize_type_category_key(type_value, category_value)
            if normalized in seen_keys:
                continue
            seen_keys.add(normalized)
            unique_records.append((type_value, category_value))

        if not unique_records:
            messagebox.showinfo(
                "Type/Category Sort Order",
                "No complete Type/Category values were found in the current combined table.",
            )
            return

        existing_entries = self._load_type_category_entries()
        existing_keys = {
            self._normalize_type_category_key(type_value, category_value)
            for type_value, category_value in existing_entries
        }
        new_entries = [
            entry
            for entry in unique_records
            if self._normalize_type_category_key(*entry) not in existing_keys
        ]

        path = self.type_category_sort_order_path
        appended = 0
        if new_entries:
            try:
                write_header = not path.exists() or path.stat().st_size == 0
            except OSError:
                write_header = True
            try:
                with path.open("a", encoding="utf-8", newline="") as fh:
                    writer = csv.writer(fh, quoting=csv.QUOTE_ALL)
                    if write_header:
                        writer.writerow(["Type", "Category"])
                    for type_value, category_value in new_entries:
                        writer.writerow([type_value, category_value])
                        appended += 1
            except Exception as exc:
                messagebox.showerror(
                    "Type/Category Sort Order",
                    f"Could not append new entries: {exc}",
                )
                return
            logger.info(
                "Appended %d Type/Category combination(s) to %s",
                appended,
                path,
            )
        else:
            if not path.exists():
                try:
                    with path.open("w", encoding="utf-8", newline="") as fh:
                        writer = csv.writer(fh, quoting=csv.QUOTE_ALL)
                        writer.writerow(["Type", "Category"])
                        for type_value, category_value in unique_records:
                            writer.writerow([type_value, category_value])
                    logger.info("Created %s with existing Type/Category combinations.", path)
                except Exception as exc:
                    messagebox.showerror(
                        "Type/Category Sort Order",
                        f"Could not create the CSV: {exc}",
                    )
                    return
                appended = len(unique_records)

        self._sort_combined_records(self.combined_all_records)
        if self.combined_ordered_columns:
            self._update_combined_tree_display()

        if appended:
            message = f"Appended {appended} new combination(s) to {path.name}."
        else:
            message = "No new Type/Category combinations were found."

        messagebox.showinfo(
            "Type/Category Sort Order",
            message + " Opening the file for review.",
        )
        self._open_in_file_manager(path)

    def _append_type_item_category_csv(self) -> None:
        if not self.combined_all_records:
            messagebox.showinfo(
                "Type/Item/Category CSV",
                "Build the combined table before appending Type/Item/Category entries.",
            )
            return

        current_records = self._filter_combined_records(self.combined_all_records)
        unique_records: List[Tuple[str, str, str]] = []
        seen_keys: Set[Tuple[str, str, str]] = set()
        for record in current_records:
            type_value = str(record.get("Type", "") or "").strip()
            category_value = str(record.get("Category", "") or "").strip()
            item_value = str(record.get("Item", "") or "").strip()
            if not (type_value and category_value and item_value):
                continue
            normalized = self._normalize_type_item_category_key(
                type_value, item_value, category_value
            )
            if normalized in seen_keys:
                continue
            seen_keys.add(normalized)
            unique_records.append((type_value, item_value, category_value))

        if not unique_records:
            messagebox.showinfo(
                "Type/Item/Category CSV",
                "No complete Type/Item/Category values were found in the current combined table.",
            )
            return

        existing_entries = self._load_type_item_category_entries()
        existing_keys = {
            self._normalize_type_item_category_key(type_value, item_value, category_value)
            for type_value, item_value, category_value in existing_entries
        }
        new_entries = [
            entry
            for entry in unique_records
            if self._normalize_type_item_category_key(*entry) not in existing_keys
        ]

        path = self.type_item_category_path
        appended = 0
        if new_entries:
            try:
                write_header = not path.exists() or path.stat().st_size == 0
            except OSError:
                write_header = True
            try:
                with path.open("a", encoding="utf-8", newline="") as fh:
                    writer = csv.writer(fh, quoting=csv.QUOTE_ALL)
                    if write_header:
                        writer.writerow(["Type", "Item", "Category"])
                    for type_value, item_value, category_value in new_entries:
                        writer.writerow([type_value, item_value, category_value])
                        appended += 1
            except Exception as exc:
                messagebox.showerror(
                    "Type/Item/Category CSV",
                    f"Could not append new entries: {exc}",
                )
                return
            logger.info(
                "Appended %d Type/Item/Category combination(s) to %s",
                appended,
                path,
            )
        else:
            logger.info("No new Type/Item/Category combinations to append.")

        if appended:
            self._sort_combined_records(self.combined_all_records)
            if self.combined_ordered_columns:
                self._update_combined_tree_display()
            message = f"Appended {appended} new combination(s) to {path.name}."
        else:
            message = "No new Type/Item/Category combinations were found."

        messagebox.showinfo("Type/Item/Category CSV", message + " Opening the file for review.")
        self._open_in_file_manager(path)

    def _update_combined_tree_display(self) -> None:
        if not self.combined_ordered_columns:
            return
        records = self._filter_combined_records(self.combined_all_records)
        self._render_combined_tree(records)

    def _on_combined_show_blank_notes_toggle(self) -> None:
        if not self.combined_ordered_columns:
            return
        tree = self.combined_result_tree
        if tree is None:
            return
        self._destroy_note_editor()
        if not self.combined_show_blank_notes_var.get():
            self._combined_blank_notification_shown = False
            return
        self._focus_first_blank_note(notify_if_missing=True)

    def _focus_first_blank_note(self, *, notify_if_missing: bool = False) -> None:
        if not self.combined_show_blank_notes_var.get():
            return
        tree = self.combined_result_tree
        if tree is None:
            return
        if not self.combined_ordered_columns or "Note" not in self.combined_ordered_columns:
            return
        selection = tree.selection()
        current_item = selection[0] if selection else tree.focus()
        if current_item and self._is_blank_note_tree_item(tree, current_item):
            self._combined_blank_notification_shown = False
            return
        children = tree.get_children("")
        first_blank: Optional[str] = None
        for item_id in children:
            if self._is_blank_note_tree_item(tree, item_id):
                first_blank = item_id
                break
        if first_blank:
            tree.selection_set(first_blank)
            tree.focus(first_blank)
            tree.see(first_blank)
            self._combined_blank_notification_shown = False
            return
        if notify_if_missing and not self._combined_blank_notification_shown and children:
            self._export_final_combined_csv()
            messagebox.showinfo(
                "Iterate blank notes only",
                "All notes assigned. Do you wish to finalize and save report?",
            )
            self._combined_blank_notification_shown = True

    def _render_combined_tree(
        self,
        records: List[Dict[str, str]],
        ordered_columns: Optional[List[str]] = None,
    ) -> None:
        if ordered_columns is not None:
            self.combined_ordered_columns = ordered_columns[:]
        if not self.combined_ordered_columns:
            return
        columns = self.combined_ordered_columns
        self._destroy_note_editor()
        for child in self.combined_result_frame.winfo_children():
            child.destroy()
        # Use the column names directly as identifiers so cell-level tagging and
        # geometry queries can reliably reference the correct Treeview column.
        self.combined_column_ids = {name: f"#{index + 1}" for index, name in enumerate(columns)}
        self.combined_note_cell_tags.clear()
        self.combined_type_cell_tags.clear()
        self.combined_category_cell_tags.clear()
        split_pane = ttk.Panedwindow(self.combined_result_frame, orient=tk.HORIZONTAL)
        split_pane.pack(fill=tk.BOTH, expand=True)
        self.combined_split_pane = split_pane

        table_container = ttk.Frame(split_pane)
        table_container.columnconfigure(0, weight=1)
        table_container.rowconfigure(1, weight=1)
        split_pane.add(table_container, weight=3)

        controls_frame = ttk.Frame(table_container)
        controls_frame.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        controls_frame.columnconfigure(0, weight=1)

        ttk.Checkbutton(
            controls_frame,
            text="Iterate blank notes only",
            variable=self.combined_show_blank_notes_var,
            command=self._on_combined_show_blank_notes_toggle,
        ).grid(row=0, column=0, sticky="w")

        ttk.Button(
            controls_frame,
            text="Load Assignments",
            command=self._prompt_import_note_assignments,
        ).grid(row=0, column=1, sticky="e", padx=(8, 0))

        ttk.Button(
            controls_frame,
            text="Sort by Combined Order",
            command=self._on_resort_combined_clicked,
        ).grid(row=0, column=2, sticky="e", padx=(8, 0))

        tree_container = ttk.Frame(table_container)
        tree_container.grid(row=1, column=0, sticky="nsew")
        tree_container.columnconfigure(0, weight=1)
        tree_container.rowconfigure(0, weight=1)

        tree = ttk.Treeview(tree_container, columns=columns, show="headings")
        tree.grid(row=0, column=0, sticky="nsew")

        def _combined_xscroll(*args: Any) -> None:
            tree.xview(*args)
            self._destroy_note_editor()

        y_scroll = ttk.Scrollbar(tree_container, orient=tk.VERTICAL, command=tree.yview)
        y_scroll.grid(row=0, column=1, sticky="ns")
        tree.configure(yscrollcommand=y_scroll.set)

        x_scroll = ttk.Scrollbar(tree_container, orient=tk.HORIZONTAL, command=_combined_xscroll)
        x_scroll.grid(row=1, column=0, columnspan=2, sticky="ew")
        tree.configure(xscrollcommand=x_scroll.set)

        preview_container = ttk.Frame(split_pane, padding=0)
        preview_container.columnconfigure(0, weight=1)
        preview_container.rowconfigure(0, weight=1)
        split_pane.add(preview_container, weight=2)

        def _lock_split_position(event: tk.Event) -> None:  # type: ignore[override]
            total_width = event.width
            if total_width <= 2:
                return
            target = int(total_width * 0.6)
            try:
                current = split_pane.sashpos(0)
            except tk.TclError:
                current = None
            if current is None or abs(current - target) > 2:
                try:
                    split_pane.sashpos(0, target)
                except tk.TclError:
                    pass

        split_pane.bind("<Configure>", _lock_split_position)

        def _initial_split_adjustment() -> None:
            total_width = split_pane.winfo_width()
            if total_width <= 2:
                return
            target = int(total_width * 0.6)
            try:
                split_pane.sashpos(0, target)
            except tk.TclError:
                pass

        self.root.after_idle(_initial_split_adjustment)
        self.combined_preview_frame = preview_container

        canvas = tk.Canvas(preview_container, highlightthickness=0, borderwidth=0)
        canvas.grid(row=0, column=0, sticky="nsew")
        y_scroll = ttk.Scrollbar(preview_container, orient=tk.VERTICAL, command=canvas.yview)
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll = ttk.Scrollbar(preview_container, orient=tk.HORIZONTAL, command=canvas.xview)
        x_scroll.grid(row=1, column=0, columnspan=2, sticky="ew")
        canvas.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        self.combined_preview_canvas = canvas
        self.combined_preview_canvas_image = None
        self.combined_preview_image = None
        self._update_combined_preview_zoom_display()
        self._clear_combined_preview()

        other_has_fixed_width = (
            isinstance(self.combined_other_column_width, int)
            and self.combined_other_column_width > 0
        )
        for column_name in columns:
            heading_text = column_name
            if column_name not in {"Type", "Category", "Item", "Note"}:
                mapping = self.combined_column_name_map.get(column_name)
                if mapping:
                    _, label_text = mapping
                    if label_text.strip():
                        heading_text = label_text
                    elif "." in column_name:
                        heading_text = column_name.split(".", 1)[1]
                elif "." in column_name:
                    heading_text = column_name.split(".", 1)[1]
            tree.heading(column_name, text=heading_text)
            if column_name in {"Type", "Category", "Item"}:
                tree.column(column_name, anchor="w", stretch=False)
            elif column_name == "Note":
                tree.column(column_name, anchor="center", stretch=False, minwidth=120)
            else:
                tree.column(
                    column_name,
                    anchor="center",
                    stretch=not other_has_fixed_width,
                )

        note_index = None
        if "Note" in columns:
            note_index = columns.index("Note")
            self.combined_note_column_id = self.combined_column_ids.get("Note")
        else:
            self.combined_note_column_id = None
        self.combined_note_record_keys.clear()
        self._configure_note_tags(tree)
        self.combined_result_tree = tree

        save_frame = ttk.Frame(table_container)
        save_frame.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        save_frame.columnconfigure(0, weight=1)

        zoom_controls = ttk.Frame(save_frame)
        zoom_controls.grid(row=0, column=0, sticky="ew")
        zoom_controls.columnconfigure(1, weight=1)

        ttk.Label(zoom_controls, text="Zoom").grid(row=0, column=0, padx=(0, 8))

        zoom_scale = ttk.Scale(
            zoom_controls,
            from_=0.5,
            to=3.0,
            orient=tk.HORIZONTAL,
            variable=self.combined_preview_zoom_var,
            command=lambda value: self._on_combined_preview_zoom_scale(float(value)),
        )
        zoom_scale.grid(row=0, column=1, sticky="ew")

        ttk.Label(
            zoom_controls,
            textvariable=self.combined_preview_zoom_display_var,
            width=6,
            anchor="e",
        ).grid(row=0, column=2, padx=8)

        ttk.Button(
            zoom_controls,
            text="-",
            width=3,
            command=lambda: self._adjust_combined_preview_zoom(-0.1),
        ).grid(row=0, column=3, padx=(0, 4))
        ttk.Button(
            zoom_controls,
            text="Reset",
            command=self._reset_combined_preview_zoom,
        ).grid(row=0, column=4, padx=(0, 4))
        ttk.Button(
            zoom_controls,
            text="+",
            width=3,
            command=lambda: self._adjust_combined_preview_zoom(0.1),
        ).grid(row=0, column=5)

        save_button = ttk.Button(
            save_frame,
            text="Save Combined CSV",
            command=self._save_combined_csv,
        )
        save_button.grid(row=0, column=1, sticky="e", padx=(8, 0))
        self.combined_save_button = save_button

        for record in records:
            values = [record.get(column_name, "") for column_name in columns]
            item_id = tree.insert("", tk.END, values=values)
            if note_index is not None and item_id:
                note_key = (
                    record.get("Type", ""),
                    record.get("Category", ""),
                    record.get("Item", ""),
                )
                if all(note_key):
                    self.combined_note_record_keys[item_id] = (
                        str(note_key[0]),
                        str(note_key[1]),
                        str(note_key[2]),
                    )
            self._apply_type_color_tag(item_id, record.get("Type", ""))
            self._apply_category_color_tag(item_id, record.get("Category", ""))
            note_value = record.get("Note", "")
            if not isinstance(note_value, str):
                note_value = str(note_value or "")
            self._apply_note_value_tag(item_id, note_value)

        tree.bind("<Button-1>", self._on_combined_tree_click)
        tree.bind("<ButtonRelease-1>", self._on_combined_tree_release, add="+")
        tree.bind("<Configure>", lambda _e: self._destroy_note_editor())
        tree.bind("<<TreeviewSelect>>", self._on_combined_tree_select)
        tree.bind("<Tab>", self._on_combined_tree_tab)
        tree.bind("<Shift-Tab>", self._on_combined_tree_shift_tab)
        tree.bind("<ISO_Left_Tab>", self._on_combined_tree_shift_tab)
        tree.bind("<Prior>", lambda _e: self._on_combined_tree_page(-10))
        tree.bind("<Next>", lambda _e: self._on_combined_tree_page(10))
        tree.bind("<Return>", self._on_combined_tree_return)
        tree.bind("<Key>", self._on_combined_tree_key)

        def _tree_mousewheel(event: tk.Event) -> str:  # type: ignore[override]
            self._destroy_note_editor()
            if event.delta:
                delta = -int(event.delta / 120)
                if delta:
                    tree.yview_scroll(delta, "units")
            elif getattr(event, "num", None) == 4:
                tree.yview_scroll(-1, "units")
            elif getattr(event, "num", None) == 5:
                tree.yview_scroll(1, "units")
            return "break"

        tree.bind("<MouseWheel>", _tree_mousewheel)
        tree.bind("<Button-4>", _tree_mousewheel)
        tree.bind("<Button-5>", _tree_mousewheel)

        self._refresh_combined_column_widths()
        self._refresh_type_category_colors()
        current_selection = tree.selection()
        if current_selection:
            self._update_combined_preview(current_selection[0])
        else:
            self._clear_combined_preview()
        self._update_combined_save_button_state()
        self._refresh_chart_tab()
        if self.combined_show_blank_notes_var.get():
            self._focus_first_blank_note(notify_if_missing=True)

    def _on_confirm_combined_clicked(self) -> None:
        self._confirm_combined_table()
        if self.combined_result_tree is not None:
            try:
                self.notebook.select(self.combined_frame)
            except tk.TclError:
                pass

    def _on_reload_combined_clicked(self) -> None:
        self._refresh_combined_tab(auto_update=True)
        try:
            self.notebook.select(self.combined_labels_tab)
        except tk.TclError:
            pass

    def _update_combined_save_button_state(self) -> None:
        button = self.combined_save_button
        if button is None or not button.winfo_exists():
            return
        if self.combined_all_records and self.combined_ordered_columns:
            button.state(["!disabled"])
        else:
            button.state(["disabled"])

    def _save_combined_csv(self) -> None:
        if not self.combined_all_records or not self.combined_ordered_columns:
            messagebox.showinfo(
                "Save Combined CSV",
                "No combined data is currently available to save.",
            )
            return
        company = self.company_var.get().strip()
        if not company:
            messagebox.showinfo(
                "Save Combined CSV",
                "Select a company before saving the combined CSV.",
            )
            return
        company_root = self.companies_dir / company
        combined_path = company_root / "combined.csv"
        try:
            csv_rows = self._build_combined_csv_rows()
            if not csv_rows:
                messagebox.showinfo(
                    "Save Combined CSV",
                    "No combined data is currently available to save.",
                )
                return
            combined_path.parent.mkdir(parents=True, exist_ok=True)
            self._write_csv_rows(csv_rows, combined_path)
            metadata = self._build_combined_metadata(company_root)
            if metadata:
                metadata_path = combined_path.with_name("combined_metadata.json")
                with metadata_path.open("w", encoding="utf-8") as fh:
                    json.dump(metadata, fh, indent=2)
        except Exception as exc:
            messagebox.showwarning("Save Combined CSV", f"Could not save the combined CSV: {exc}")
            return
        messagebox.showinfo("Save Combined CSV", f"Combined CSV saved to:\n{combined_path}")

    def _build_combined_csv_rows(self) -> Optional[List[List[str]]]:
        if not self.combined_all_records or not self.combined_ordered_columns:
            return None
        csv_header: List[str] = []
        for column_name in self.combined_ordered_columns:
            if column_name in {"Type", "Category", "Item", "Note"}:
                csv_header.append(column_name)
            elif "." in column_name:
                csv_header.append(column_name.split(".", 1)[1])
            else:
                csv_header.append(column_name)
        csv_rows: List[List[str]] = [csv_header]
        for record in self.combined_all_records:
            csv_rows.append(
                [record.get(column_name, "") for column_name in self.combined_ordered_columns]
            )
        return csv_rows

    def _build_combined_metadata(self, company_root: Path) -> Optional[Dict[str, Any]]:
        if not self.combined_all_records or not self.combined_ordered_columns:
            return None
        if not getattr(self, "combined_column_name_map", None):
            return None

        periods: Dict[str, Dict[str, str]] = {}

        def _relative_pdf(path: Path) -> str:
            try:
                return str(path.relative_to(company_root))
            except ValueError:
                return str(path)

        for column_name in self.combined_ordered_columns:
            if column_name in {"Type", "Category", "Item", "Note"}:
                continue
            mapping = self.combined_column_name_map.get(column_name)
            if not mapping:
                continue
            pdf_path, label_value = mapping
            periods.setdefault(
                label_value,
                {
                    "label": label_value,
                    "pdf": _relative_pdf(pdf_path),
                    "pdf_name": pdf_path.name,
                },
            )

        if not periods:
            return None

        rows: List[Dict[str, Any]] = []

        for record in self.combined_all_records:
            type_value = str(record.get("Type", ""))
            category_value = str(record.get("Category", ""))
            item_value = str(record.get("Item", ""))
            row_entry: Dict[str, Any] = {
                "type": type_value,
                "category": category_value,
                "item": item_value,
                "note": str(record.get("Note", "")),
                "periods": {},
            }

            for column_name in self.combined_ordered_columns:
                if column_name in {"Type", "Category", "Item", "Note"}:
                    continue
                mapping = self.combined_column_name_map.get(column_name)
                if not mapping:
                    continue
                pdf_path, label_value = mapping
                assigned_page = self._get_assigned_page_number(pdf_path, type_value)
                row_entry["periods"][label_value] = {
                    "pdf": _relative_pdf(pdf_path),
                    "page": assigned_page,
                }

            if row_entry["periods"]:
                rows.append(row_entry)

        if not rows:
            return None

        return {
            "periods": periods,
            "rows": rows,
        }

    def _get_assigned_page_number(self, pdf_path: Path, type_value: str) -> Optional[int]:
        if not self.assigned_pages:
            return None
        record = self.assigned_pages.get(pdf_path.name)
        if not isinstance(record, dict):
            return None
        selections = record.get("selections")
        if not isinstance(selections, dict):
            return None
        page_value = selections.get(type_value)
        if page_value is None:
            return None
        try:
            page_index = int(page_value)
        except (TypeError, ValueError):
            return None
        return page_index + 1 if page_index >= 0 else None

    def _export_final_combined_csv(self) -> None:
        if not self.combined_all_records or not self.combined_ordered_columns:
            return
        company = self.company_var.get().strip()
        if not company:
            return
        csv_rows = self._build_combined_csv_rows()
        if not csv_rows:
            return
        company_root = self.companies_dir / company
        final_path = company_root / "combined.csv"
        try:
            final_path.parent.mkdir(parents=True, exist_ok=True)
            self._write_csv_rows(csv_rows, final_path)
            logger.info("Exported final combined CSV to %s", final_path)
            metadata = self._build_combined_metadata(company_root)
            if metadata:
                metadata_path = final_path.with_name("combined_metadata.json")
                with metadata_path.open("w", encoding="utf-8") as fh:
                    json.dump(metadata, fh, indent=2)
                logger.info("Exported combined metadata to %s", metadata_path)
        except Exception:
            logger.exception("Failed to export final combined CSV to company root")

    def _clear_combined_preview(self, message: Optional[str] = None) -> None:
        canvas = self.combined_preview_canvas
        if canvas is None or not canvas.winfo_exists():
            return
        canvas.delete("all")
        canvas.configure(scrollregion=(0, 0, 0, 0))
        self.combined_preview_image = None
        self.combined_preview_canvas_image = None
        self.combined_preview_target = None
        display_text = message or "Select a row to view the PDF page."
        self.combined_preview_detail_var.set(display_text)

    def _update_combined_preview_zoom_display(self) -> None:
        zoom = self.combined_preview_zoom_var.get()
        percent = int(round(zoom * 100))
        self.combined_preview_zoom_display_var.set(f"{percent}%")

    def _schedule_combined_zoom_persist(self) -> None:
        if not self._config_loaded:
            return
        if self._combined_zoom_save_job is not None:
            try:
                self.root.after_cancel(self._combined_zoom_save_job)
            except Exception:
                pass

        def _persist() -> None:
            self._combined_zoom_save_job = None
            zoom_value = max(0.5, min(3.0, float(self.combined_preview_zoom_var.get())))
            self.combined_preview_zoom_var.set(zoom_value)
            self._write_config()

        self._combined_zoom_save_job = self.root.after(250, _persist)

    def _adjust_combined_preview_zoom(self, delta: float) -> None:
        zoom = self.combined_preview_zoom_var.get() + delta
        zoom = max(0.5, min(3.0, zoom))
        self.combined_preview_zoom_var.set(zoom)
        self._update_combined_preview_zoom_display()
        self._rerender_combined_preview()
        self._schedule_combined_zoom_persist()

    def _reset_combined_preview_zoom(self) -> None:
        self.combined_preview_zoom_var.set(1.0)
        self._update_combined_preview_zoom_display()
        self._rerender_combined_preview()
        self._schedule_combined_zoom_persist()

    def _on_combined_preview_zoom_scale(self, value: float) -> None:
        zoom = max(0.5, min(3.0, float(value)))
        if abs(self.combined_preview_zoom_var.get() - zoom) > 1e-6:
            self.combined_preview_zoom_var.set(zoom)
        self._update_combined_preview_zoom_display()
        self._rerender_combined_preview()
        self._schedule_combined_zoom_persist()

    def _rerender_combined_preview(self) -> None:
        target = self.combined_preview_target
        if not target:
            return
        entry, page_index, key = target
        self._display_combined_preview(entry, page_index, key)

    def _display_combined_preview(
        self, entry: PDFEntry, page_index: int, key: Tuple[str, str, str]
    ) -> None:
        canvas = self.combined_preview_canvas
        if canvas is None or not canvas.winfo_exists():
            return
        detail_text = f"{entry.path.name} — {key[0]} (Page {page_index + 1})"
        zoom = self.combined_preview_zoom_var.get()
        self._update_combined_preview_zoom_display()
        canvas.update_idletasks()
        canvas_width = canvas.winfo_width()
        if canvas_width <= 2:
            try:
                parent_width = canvas.master.winfo_width() if canvas.master else 0
            except Exception:
                parent_width = 0
            canvas_width = max(canvas_width, parent_width)
        if canvas_width <= 2:
            canvas_width = 480
        target_width = max(120, int(canvas_width * zoom))
        photo = self._render_page(entry.doc, page_index, target_width)
        if photo is None:
            canvas.delete("all")
            canvas.configure(scrollregion=(0, 0, 0, 0))
            self.combined_preview_image = None
            self.combined_preview_canvas_image = None
            self.combined_preview_detail_var.set(detail_text)
            return
        self.combined_preview_image = photo
        canvas.delete("all")
        image_id = canvas.create_image(0, 0, anchor="nw", image=photo)
        canvas.configure(scrollregion=canvas.bbox("all") or (0, 0, 0, 0))
        self.combined_preview_canvas_image = image_id
        self.combined_preview_detail_var.set(detail_text)

    def _resolve_combined_preview_target(
        self, key: Tuple[str, str, str]
    ) -> Optional[Tuple[PDFEntry, int]]:
        source_paths = self.combined_row_sources.get(key, [])
        if not source_paths:
            return None
        category = key[0]
        for path in source_paths:
            entry = self.pdf_entry_by_path.get(path)
            if entry is None:
                continue
            page_index = self._get_selected_page_index(entry, category)
            if page_index is None:
                matches = entry.matches.get(category, [])
                if matches:
                    page_index = matches[0].page_index
            if page_index is not None:
                return entry, page_index
        return None

    def _update_combined_preview(self, item_id: Optional[str]) -> None:
        canvas = self.combined_preview_canvas
        if canvas is None or not canvas.winfo_exists():
            return
        tree = self.combined_result_tree
        if tree is None:
            self._clear_combined_preview()
            return
        if not item_id:
            self._clear_combined_preview()
            return
        self.combined_preview_target = None
        key = self.combined_note_record_keys.get(item_id)
        if not key:
            if not self.combined_ordered_columns:
                self._clear_combined_preview()
                return
            try:
                type_index = self.combined_ordered_columns.index("Type")
                category_index = self.combined_ordered_columns.index("Category")
                item_index = self.combined_ordered_columns.index("Item")
            except ValueError:
                self._clear_combined_preview()
                return
            values = list(tree.item(item_id, "values"))
            if (
                type_index < len(values)
                and category_index < len(values)
                and item_index < len(values)
            ):
                key = (
                    str(values[type_index]),
                    str(values[category_index]),
                    str(values[item_index]),
                )
        if not key:
            self._clear_combined_preview("No PDF page available for this row.")
            return
        target = self._resolve_combined_preview_target(key)
        if target is None:
            self._clear_combined_preview("No PDF page available for this row.")
            return
        entry, page_index = target
        self.combined_preview_target = (entry, page_index, key)
        self._display_combined_preview(entry, page_index, key)

    def _on_combined_tree_select(self, _: tk.Event) -> None:  # type: ignore[override]
        self._destroy_note_editor()
        tree = self.combined_result_tree
        if tree is None:
            self._clear_combined_preview()
            return
        selection = tree.selection()
        item_id = selection[0] if selection else None
        self._update_combined_preview(item_id)

    def _update_combined_record_note_value(self, key: Tuple[str, str, str], value: str) -> None:
        record = self.combined_record_lookup.get(key)
        if record is None:
            return
        record["Note"] = value
        if isinstance(value, str):
            normalized = value.strip()
        else:
            normalized = str(value or "").strip()
        if not normalized:
            self._combined_blank_notification_shown = False

    def _on_primary_tab_changed(self, event: tk.Event) -> None:  # type: ignore[override]
        widget = event.widget
        if not isinstance(widget, ttk.Notebook):
            return
        try:
            selected = widget.select()
        except tk.TclError:
            return
        if not selected:
            return
        if self.combined_frame is not None and str(self.combined_frame) == selected:
            if self.combined_show_blank_notes_var.get():
                self._focus_first_blank_note(notify_if_missing=True)

    def _confirm_combined_table(self, auto: bool = False) -> None:
        if not self.combined_pdf_order or not self.combined_csv_sources:
            if not auto:
                messagebox.showinfo("Combine Results", "No scraped data is available to combine.")
            return

        ordered_columns = ["Type", "Category", "Item", "Note"]
        column_labels_by_pdf: Dict[Path, List[str]] = {}
        column_name_map: Dict[str, Tuple[Path, str]] = {}
        for path in self.combined_pdf_order:
            pdf_label = path.stem
            labels: List[str] = []
            for position in range(self.combined_max_data_columns):
                key = (path, position)
                var = self.combined_column_label_vars.get(key)
                label_value = var.get().strip() if var else ""
                if not label_value:
                    label_value = self.combined_column_defaults.get(key, f"Value {position + 1}")
                labels.append(label_value)
                column_name = f"{pdf_label}.{label_value}"
                ordered_columns.append(column_name)
                column_name_map[column_name] = (path, label_value)
            logger.info("Combined labels for %s -> %s", path.name, labels)
            column_labels_by_pdf[path] = labels

        ordered_columns = self._sort_combined_period_columns(ordered_columns, column_name_map)

        logger.info("Combined ordered columns: %s", ordered_columns)
        self.combined_labels_by_pdf = {path: labels[:] for path, labels in column_labels_by_pdf.items()}
        self.combined_column_name_map = column_name_map

        record_map: Dict[Tuple[str, str, str], Dict[Tuple[Path, int], str]] = {}

        for path in self.combined_pdf_order:
            for category in COLUMNS:
                source = self.combined_csv_sources.get((path, category))
                if not source:
                    continue
                rows: List[List[str]] = source.get("rows", [])
                if not rows:
                    continue
                data_indices: List[int] = source.get("data_indices", [])
                category_index = source.get("category_index")
                item_index = source.get("item_index")
                if not isinstance(category_index, int) or not isinstance(item_index, int):
                    continue
                for row in rows:
                    if category_index >= len(row) or item_index >= len(row):
                        continue
                    category_value = row[category_index]
                    item_value = row[item_index]
                    key = (category, category_value, item_value)
                    value_map = record_map.setdefault(key, {})
                    for position in range(self.combined_max_data_columns):
                        value = ""
                        if position < len(data_indices):
                            data_index = data_indices[position]
                            if data_index < len(row):
                                value = row[data_index]
                        value_map[(path, position)] = value

        combined_records: List[Dict[str, str]] = []
        self.combined_row_sources = {}
        for category in COLUMNS:
            category_keys = [key for key in record_map if key[0] == category]
            category_keys.sort(key=lambda item: (item[1], item[2]))
            logger.info("Combined category %s -> %d rows", category, len(category_keys))
            for key in category_keys:
                _type, category_value, item_value = key
                value_map = record_map.get(key, {})
                raw_note = self.note_assignments.get((category, category_value, item_value), "")
                if isinstance(raw_note, str):
                    note_value = raw_note.strip().lower()
                else:
                    note_value = str(raw_note or "").strip().lower()
                if note_value and note_value not in self.note_options:
                    self._register_note_option(note_value)
                if note_value and note_value not in self.note_options:
                    note_value = ""
                record: Dict[str, str] = {
                    "Type": category,
                    "Category": category_value,
                    "Item": item_value,
                    "Note": note_value,
                }
                paths_with_data = {
                    path_key for path_key, _ in value_map.keys()
                }
                ordered_sources: List[Path] = [
                    path for path in self.combined_pdf_order if path in paths_with_data
                ]
                for path in self.combined_pdf_order:
                    pdf_label = path.stem
                    labels = column_labels_by_pdf.get(path, [])
                    for position, base_label in enumerate(labels):
                        column_name = f"{pdf_label}.{base_label}"
                        raw_value = value_map.get((path, position), "")
                        record[column_name] = self._format_combined_value(raw_value)
                if ordered_sources:
                    self.combined_row_sources[
                        (str(category), str(category_value), str(item_value))
                    ] = ordered_sources
                combined_records.append(record)

        self._ensure_combined_order_entries(combined_records)
        self._sort_combined_records(combined_records)

        logger.info(
            "Combined result set -> %d records with %d columns",
            len(combined_records),
            len(ordered_columns),
        )
        if not combined_records:
            if not auto:
                messagebox.showinfo(
                    "Combine Results", "No rows were available after processing the scraped results."
                )
            return

        self.combined_all_records = combined_records
        record_lookup: Dict[Tuple[str, str, str], Dict[str, str]] = {}
        for record in combined_records:
            key = (
                str(record.get("Type", "")),
                str(record.get("Category", "")),
                str(record.get("Item", "")),
            )
            if all(key):
                record_lookup[key] = record
        self.combined_record_lookup = record_lookup

        display_records = self._filter_combined_records(combined_records)
        self._render_combined_tree(display_records, ordered_columns)

    def _get_tree_font(self, tree: ttk.Treeview) -> tkfont.Font:
        try:
            return tkfont.nametofont(tree.cget("font"))
        except tk.TclError:
            return tkfont.nametofont("TkDefaultFont")

    def _refresh_combined_column_widths(self) -> None:
        self._destroy_note_editor()
        self._apply_combined_base_column_widths()
        self._apply_combined_other_column_widths()

    def _apply_combined_base_column_widths(self) -> None:
        tree = self.combined_result_tree
        if tree is None or not self.combined_ordered_columns:
            return
        tree.update_idletasks()
        tree_font = self._get_tree_font(tree)
        for column_name in self.combined_ordered_columns:
            if column_name not in {"Type", "Category", "Item", "Note"}:
                continue
            stored_width = self.combined_base_column_widths.get(column_name)
            if isinstance(stored_width, int) and stored_width > 0:
                tree.column(
                    column_name,
                    width=stored_width,
                    minwidth=max(20, stored_width),
                    stretch=False,
                )
                continue
            max_width = tree_font.measure(column_name)
            for item_id in tree.get_children(""):
                value_text = tree.set(item_id, column_name) or ""
                max_width = max(max_width, tree_font.measure(str(value_text)))
            if column_name == "Note":
                reference_width = tree_font.measure("share_count")
                desired_width = max(max_width + 32, reference_width + 32, 120)
            else:
                desired_width = max(max_width + 24, 160)
            tree.column(
                column_name,
                width=desired_width,
                minwidth=max(20, desired_width),
                stretch=False,
            )
        other_width = self.combined_other_column_width
        if isinstance(other_width, int) and other_width > 0:
            for column_name in self.combined_ordered_columns:
                if column_name in {"Type", "Category", "Item", "Note"}:
                    continue
                tree.column(
                    column_name,
                    width=other_width,
                    minwidth=max(20, other_width),
                    stretch=False,
                )

    def _persist_combined_base_column_widths(self) -> None:
        if not self.combined_base_column_widths:
            self.config_data.pop("combined_base_column_widths", None)
        else:
            self.config_data["combined_base_column_widths"] = {
                column: int(width)
                for column, width in self.combined_base_column_widths.items()
                if column in {"Type", "Category", "Item", "Note"} and width > 0
            }
        if isinstance(self.combined_other_column_width, int) and self.combined_other_column_width > 0:
            self.config_data["combined_other_column_width"] = int(self.combined_other_column_width)
        else:
            self.config_data.pop("combined_other_column_width", None)
        self._write_config()

    def _persist_type_category_colors(self) -> None:
        def _build_store(
            color_map: Dict[str, str], labels: Dict[str, str]
        ) -> Dict[str, str]:
            stored: Dict[str, str] = {}
            for normalized, color in color_map.items():
                normalized_color = self._normalize_hex_color(color)
                if not normalized_color:
                    continue
                label = labels.get(normalized, normalized)
                stored[label] = normalized_color
            return stored

        type_store = _build_store(self.type_color_map, self.type_color_labels)
        category_store = _build_store(self.category_color_map, self.category_color_labels)

        if type_store:
            self.config_data["combined_type_colors"] = type_store
        else:
            self.config_data.pop("combined_type_colors", None)

        if category_store:
            self.config_data["combined_category_colors"] = category_store
        else:
            self.config_data.pop("combined_category_colors", None)

        # Remove legacy configuration key if present
        self.config_data.pop("combined_type_category_colors", None)
        self._write_config()

    def _store_combined_base_column_widths(self) -> None:
        tree = self.combined_result_tree
        if tree is None:
            return
        changed = False
        for column_name in ("Type", "Category", "Item", "Note"):
            if column_name not in self.combined_ordered_columns:
                continue
            try:
                column_info = tree.column(column_name)
            except tk.TclError:
                continue
            width = int(column_info.get("width", 0)) if isinstance(column_info, dict) else 0
            if width <= 0:
                continue
            if self.combined_base_column_widths.get(column_name) != width:
                self.combined_base_column_widths[column_name] = width
                changed = True
        if changed:
            self._persist_combined_base_column_widths()

    def _apply_combined_other_column_widths(self) -> None:
        tree = self.combined_result_tree
        if tree is None or not self.combined_ordered_columns:
            return
        date_columns = [
            column_name
            for column_name in self.combined_ordered_columns
            if column_name not in {"Type", "Category", "Item", "Note"}
        ]
        if not date_columns:
            return
        self._destroy_note_editor()
        tree.update_idletasks()
        if isinstance(self.combined_other_column_width, int) and self.combined_other_column_width > 0:
            for column_name in date_columns:
                tree.column(
                    column_name,
                    width=self.combined_other_column_width,
                    minwidth=max(20, self.combined_other_column_width),
                    stretch=False,
                )
            return
        tree_font = self._get_tree_font(tree)
        for column_name in date_columns:
            max_width = tree_font.measure(column_name)
            for item_id in tree.get_children(""):
                value_text = tree.set(item_id, column_name)
                if value_text is None:
                    value_text = ""
                max_width = max(max_width, tree_font.measure(str(value_text)))
            desired_width = max(max_width + 24, 120)
            tree.column(
                column_name,
                width=desired_width,
                minwidth=max(20, desired_width),
                stretch=False,
            )

    def _show_raw_text_dialog(self, entry: PDFEntry, page_index: int) -> None:
        try:
            page = entry.doc.load_page(page_index)
            page_text = page.get_text("text")
        except Exception as exc:
            page_text = f"Unable to load the parsed page text for page {page_index + 1}: {exc}"

        dialog = tk.Toplevel(self.root)
        dialog.title(f"Raw Text - {entry.path.name} (Page {page_index + 1})")
        dialog.transient(self.root)
        dialog.grab_set()

        content_frame = ttk.Frame(dialog, padding=(12, 12, 12, 0))
        content_frame.pack(fill=tk.BOTH, expand=True)

        text_widget = tk.Text(content_frame, wrap="word", width=80, height=25)
        text_widget.insert("1.0", page_text)
        text_widget.configure(font=("TkFixedFont", 9), state="disabled")
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(content_frame, orient=tk.VERTICAL, command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        button_frame = ttk.Frame(dialog, padding=(12, 0, 12, 12))
        button_frame.pack(fill=tk.X)
        ttk.Button(button_frame, text="Close", command=dialog.destroy).pack(side=tk.RIGHT)
        ttk.Button(
            button_frame,
            text="Copy",
            command=lambda: self._copy_to_clipboard(page_text),
        ).pack(side=tk.RIGHT, padx=(0, 8))

        dialog.bind("<Escape>", lambda _e: dialog.destroy())
        dialog.wait_visibility()
        dialog.focus_set()

    def _show_text_preview(self, txt_path: Optional[Path], preloaded_text: Optional[str] = None) -> None:
        if txt_path is None or not txt_path.exists():
            messagebox.showinfo(
                "View TXT", "The selected response file could not be found on disk."
            )
            return

        text = preloaded_text
        if text is None:
            try:
                text = txt_path.read_text(encoding="utf-8")
            except Exception as exc:
                messagebox.showwarning("View TXT", f"Could not read response file: {exc}")
                return

        dialog = tk.Toplevel(self.root)
        dialog.title(f"Response Preview - {txt_path.name}")
        dialog.transient(self.root)
        dialog.grab_set()

        content_frame = ttk.Frame(dialog, padding=(12, 12, 12, 0))
        content_frame.pack(fill=tk.BOTH, expand=True)
        content_frame.columnconfigure(0, weight=1)
        content_frame.rowconfigure(0, weight=1)

        text_widget = tk.Text(content_frame, wrap="word", width=100, height=30)
        text_widget.insert("1.0", text)
        text_widget.configure(font=("TkFixedFont", 9), state="disabled")
        text_widget.grid(row=0, column=0, sticky="nsew")

        y_scroll = ttk.Scrollbar(content_frame, orient=tk.VERTICAL, command=text_widget.yview)
        y_scroll.grid(row=0, column=1, sticky="ns")
        text_widget.configure(yscrollcommand=y_scroll.set)

        button_frame = ttk.Frame(dialog, padding=(12, 0, 12, 12))
        button_frame.pack(fill=tk.X)
        ttk.Button(button_frame, text="Close", command=dialog.destroy).pack(side=tk.RIGHT)
        ttk.Button(
            button_frame,
            text="Copy",
            command=lambda data=text: self._copy_to_clipboard(data),
        ).pack(side=tk.RIGHT, padx=(0, 8))

        dialog.bind("<Escape>", lambda _e: dialog.destroy())
        dialog.wait_visibility()
        dialog.focus_set()

    def _show_csv_preview(self, csv_path: Path) -> None:
        if not csv_path.exists():
            messagebox.showinfo(
                "View CSV", "The selected CSV file could not be found on disk."
            )
            return

        try:
            csv_text = csv_path.read_text(encoding="utf-8")
        except Exception as exc:
            messagebox.showwarning("View CSV", f"Could not read CSV file: {exc}")
            return

        dialog = tk.Toplevel(self.root)
        dialog.title(f"CSV Preview - {csv_path.name}")
        dialog.transient(self.root)
        dialog.grab_set()

        content_frame = ttk.Frame(dialog, padding=(12, 12, 12, 0))
        content_frame.pack(fill=tk.BOTH, expand=True)
        content_frame.columnconfigure(0, weight=1)
        content_frame.rowconfigure(0, weight=1)

        text_widget = tk.Text(content_frame, wrap="none", width=100, height=30)
        text_widget.insert("1.0", csv_text)
        text_widget.configure(font=("TkFixedFont", 9), state="disabled")
        text_widget.grid(row=0, column=0, sticky="nsew")

        y_scroll = ttk.Scrollbar(content_frame, orient=tk.VERTICAL, command=text_widget.yview)
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll = ttk.Scrollbar(content_frame, orient=tk.HORIZONTAL, command=text_widget.xview)
        x_scroll.grid(row=1, column=0, sticky="ew")
        text_widget.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

        button_frame = ttk.Frame(dialog, padding=(12, 0, 12, 12))
        button_frame.pack(fill=tk.X)
        ttk.Button(button_frame, text="Close", command=dialog.destroy).pack(side=tk.RIGHT)
        ttk.Button(
            button_frame,
            text="Copy",
            command=lambda: self._copy_to_clipboard(csv_text),
        ).pack(side=tk.RIGHT, padx=(0, 8))

        dialog.bind("<Escape>", lambda _e: dialog.destroy())
        dialog.wait_visibility()
        dialog.focus_set()

    def _show_prompt_preview(self, entry: PDFEntry, category: str, page_index: int) -> None:
        company = self.company_var.get()
        if not company:
            messagebox.showinfo("Prompt Preview", "Select a company to view its prompts.")
            return
        prompt_text = self._get_prompt_text(company, category)
        if prompt_text is None:
            messagebox.showerror(
                "Prompt Preview",
                f"No prompt file found for the '{category}' category.",
            )
            return

        try:
            page = entry.doc.load_page(page_index)
            page_text = page.get_text("text")
        except Exception as exc:
            messagebox.showerror(
                "Prompt Preview",
                f"Unable to load the parsed page text for page {page_index + 1}: {exc}",
            )
            return

        combined_text = (
            f"=== System Prompt ===\n{prompt_text}\n\n=== Parsed Page Text ===\n{page_text}"
        )

        dialog = tk.Toplevel(self.root)
        dialog.title(f"{category} Prompt Preview")
        dialog.transient(self.root)
        dialog.grab_set()

        content_frame = ttk.Frame(dialog, padding=(12, 12, 12, 0))
        content_frame.pack(fill=tk.BOTH, expand=True)

        text_widget = tk.Text(content_frame, wrap="word", width=80, height=25)
        text_widget.insert("1.0", combined_text)
        text_widget.configure(font=("TkFixedFont", 9), state="disabled")
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(content_frame, orient=tk.VERTICAL, command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        button_frame = ttk.Frame(dialog, padding=(12, 0, 12, 12))
        button_frame.pack(fill=tk.X)
        ttk.Button(button_frame, text="Close", command=dialog.destroy).pack(side=tk.RIGHT)
        ttk.Button(
            button_frame,
            text="Copy",
            command=lambda: self._copy_to_clipboard(combined_text),
        ).pack(side=tk.RIGHT, padx=(0, 8))

        dialog.bind("<Escape>", lambda _e: dialog.destroy())
        dialog.wait_visibility()
        dialog.focus_set()

    def _copy_to_clipboard(self, text: str) -> None:
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
        except tk.TclError:
            messagebox.showwarning(
                "Copy Failed", "Could not access the clipboard to copy the prompt preview."
            )

    def _format_table_value(self, value: str) -> str:
        stripped = value.strip()
        if not stripped:
            return ""

        suffix = ""
        normalized = stripped
        if normalized.endswith("%"):
            suffix = "%"
            normalized = normalized[:-1]

        prefix = ""
        if normalized and normalized[0] in "$€£¥":
            prefix = normalized[0]
            normalized = normalized[1:]

        normalized = normalized.strip()
        if normalized.startswith("(") and normalized.endswith(")") and len(normalized) > 2:
            normalized = f"-{normalized[1:-1]}"

        normalized = normalized.replace(",", "")

        try:
            numeric_value = float(normalized)
        except ValueError:
            return stripped

        formatted = f"{numeric_value:.6e}"
        return f"{prefix}{formatted}{suffix}"

    def _metadata_file_for_entry(self, scrape_root: Path, entry_stem: str) -> Path:
        return scrape_root / f"{entry_stem}_metadata.json"

    def _legacy_doc_dir(self, scrape_root: Path, entry_stem: str) -> Path:
        return scrape_root / entry_stem

    def _scrape_candidate_paths(self, scrape_root: Path, entry_stem: str, name: str) -> List[Path]:
        candidates: List[Path] = []
        modern = scrape_root / name
        candidates.append(modern)
        legacy_dir = self._legacy_doc_dir(scrape_root, entry_stem)
        legacy = legacy_dir / name
        if legacy != modern:
            candidates.append(legacy)
        return candidates

    def _resolve_scrape_path(
        self, scrape_root: Path, entry_stem: str, name: str
    ) -> Path:
        for candidate in self._scrape_candidate_paths(scrape_root, entry_stem, name):
            if candidate.exists():
                return candidate
        return scrape_root / name

    def _build_scrape_filename(self, pdf_stem: str, category: str) -> str:
        safe_category = re.sub(r"[^0-9A-Za-z]+", "_", category).strip("_")
        if not safe_category:
            safe_category = "data"
        return f"{pdf_stem}_{safe_category}.csv"

    def _load_doc_metadata(self, target: Path) -> Dict[str, Any]:
        if target.is_dir():
            metadata_path = target / "metadata.json"
        else:
            metadata_path = target
        if not metadata_path.exists():
            return {}
        try:
            with metadata_path.open("r", encoding="utf-8") as fh:
                data = json.load(fh)
        except Exception:
            return {}
        return data if isinstance(data, dict) else {}

    def _write_doc_metadata(self, target: Path, data: Dict[str, Any]) -> None:
        if target.is_dir():
            metadata_path = target / "metadata.json"
        else:
            metadata_path = target
        metadata_path.parent.mkdir(parents=True, exist_ok=True)
        try:
            with metadata_path.open("w", encoding="utf-8") as fh:
                json.dump(data, fh, indent=2)
        except Exception as exc:
            messagebox.showwarning("Save Metadata", f"Could not save scrape metadata: {exc}")

    def _delete_scrape_output(
        self, scrape_root: Path, metadata_path: Path, entry_stem: str, category: str
    ) -> None:
        metadata = self._load_doc_metadata(metadata_path)
        if not metadata:
            legacy_dir = self._legacy_doc_dir(scrape_root, entry_stem)
            if legacy_dir.exists():
                metadata = self._load_doc_metadata(legacy_dir)
        if not metadata:
            return
        meta = metadata.get(category)
        errors: List[str] = []
        seen_paths: Set[Path] = set()
        if isinstance(meta, dict):
            for key in ("csv", "txt"):
                name = meta.get(key)
                if not isinstance(name, str) or not name:
                    continue
                for path in self._scrape_candidate_paths(scrape_root, entry_stem, name):
                    if path in seen_paths:
                        continue
                    seen_paths.add(path)
                    try:
                        path.unlink(missing_ok=True)
                    except TypeError:
                        try:
                            if path.exists():
                                path.unlink()
                        except FileNotFoundError:
                            pass
                        except Exception as exc:
                            errors.append(f"Could not delete {name}: {exc}")
                    except FileNotFoundError:
                        pass
                    except Exception as exc:
                        errors.append(f"Could not delete {name}: {exc}")
        metadata.pop(category, None)
        self._write_doc_metadata(metadata_path, metadata)
        self._refresh_scraped_tab()
        if errors:
            messagebox.showwarning("Delete Output", "\n".join(errors))

    def _reload_scraped_csv(
        self, scrape_root: Path, metadata_path: Path, entry_stem: str, category: str
    ) -> None:
        metadata = self._load_doc_metadata(metadata_path)
        if not metadata:
            legacy_dir = self._legacy_doc_dir(scrape_root, entry_stem)
            if legacy_dir.exists():
                metadata = self._load_doc_metadata(legacy_dir)
        if not metadata:
            messagebox.showinfo("Reload CSV", "No scrape metadata found for this selection.")
            return
        meta = metadata.get(category)
        if not isinstance(meta, dict):
            messagebox.showinfo("Reload CSV", "No CSV file is associated with this selection.")
            return
        csv_name = meta.get("csv")
        if not isinstance(csv_name, str) or not csv_name:
            messagebox.showinfo("Reload CSV", "No CSV file is associated with this selection.")
            return
        csv_path = self._resolve_scrape_path(scrape_root, entry_stem, csv_name)
        if not csv_path.exists():
            messagebox.showinfo("Reload CSV", "The recorded CSV file could not be found on disk.")
            self._refresh_scraped_tab()
            return
        rows = self._read_csv_rows(csv_path)
        if not rows:
            messagebox.showwarning(
                "Reload CSV",
                "The CSV file could not be loaded or did not contain any data.",
            )
        self._refresh_scraped_tab()

    def _open_csv_file(self, csv_path: Path) -> None:
        if not csv_path.exists():
            messagebox.showinfo("Open CSV", "The selected CSV file could not be found on disk.")
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(csv_path)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(csv_path)])
            else:
                subprocess.Popen(["xdg-open", str(csv_path)])
        except Exception as exc:
            messagebox.showwarning("Open CSV", f"Could not open CSV file: {exc}")

    def _delete_scraped_row(self, tree: ttk.Treeview) -> None:
        info = self.scraped_table_sources.get(tree)
        if not info:
            return

        selection = tree.selection()
        if not selection:
            messagebox.showinfo("Delete Row", "Select a row to delete from the table.")
            return

        csv_path: Optional[Path] = info.get("csv_path")
        if csv_path is None:
            messagebox.showinfo(
                "Delete Row",
                "This table is not associated with a CSV file that can be updated.",
            )
            return

        data_rows: List[List[str]] = info.get("rows", [])
        if not data_rows:
            messagebox.showinfo("Delete Row", "There are no rows available to delete.")
            return

        # Map current tree items to the backing data row positions.
        index_map = {item: idx for idx, item in enumerate(tree.get_children())}
        delete_indices = sorted(
            {index_map[item] for item in selection if item in index_map}, reverse=True
        )
        if not delete_indices:
            messagebox.showinfo("Delete Row", "Select a row to delete from the table.")
            return

        for item in selection:
            tree.delete(item)

        for idx in delete_indices:
            if 0 <= idx < len(data_rows):
                data_rows.pop(idx)

        tree.configure(height=min(15, max(3, len(data_rows))))

        header: List[str] = info.get("header", [])
        delimiter: str = info.get("delimiter", ",")
        if not self._write_scraped_csv_rows(csv_path, header, data_rows, delimiter=delimiter):
            scrape_root: Optional[Path] = info.get("scrape_root")
            metadata_path: Optional[Path] = info.get("metadata_path")
            entry_stem = info.get("entry_stem")
            category: Optional[str] = info.get("category")
            if (
                isinstance(scrape_root, Path)
                and isinstance(metadata_path, Path)
                and isinstance(entry_stem, str)
                and isinstance(category, str)
            ):
                self._reload_scraped_csv(scrape_root, metadata_path, entry_stem, category)
            return

        info["rows"] = data_rows
        self._update_scraped_controls_state(info)

        self._refresh_combined_tab(auto_update=True)

    def _delete_all_scraped(self) -> None:
        company = self.company_var.get()
        if not company:
            messagebox.showinfo("Delete Scraped Data", "Select a company before deleting scraped results.")
            return

        scrape_root = self.companies_dir / company / "openapiscrape"
        if not scrape_root.exists():
            messagebox.showinfo("Delete Scraped Data", "No scraped results found for the selected company.")
            return

        confirm = messagebox.askyesno(
            "Delete Scraped Data",
            "Are you sure you want to delete all scraped files for this company?",
            icon="warning",
        )
        if not confirm:
            return

        try:
            shutil.rmtree(scrape_root)
        except FileNotFoundError:
            pass
        except Exception as exc:
            messagebox.showwarning("Delete Scraped Data", f"Could not delete scraped data: {exc}")

        self._refresh_scraped_tab()

    def _read_csv_rows(self, csv_path: Path) -> List[List[str]]:
        try:
            with csv_path.open("r", encoding="utf-8", newline="") as fh:
                text = fh.read()
        except Exception:
            return []

        if not text.strip():
            return []

        delimiter = self._infer_delimiter_from_text(text)
        try:
            reader = csv.reader(io.StringIO(text), delimiter=delimiter)
            return [
                [self._normalize_scraped_cell(cell) for cell in row]
                for row in reader
            ]
        except Exception:
            return []

    def _detect_csv_delimiter(self, csv_path: Path) -> str:
        try:
            with csv_path.open("r", encoding="utf-8", newline="") as fh:
                sample = fh.read(4096)
        except Exception:
            return ","
        return self._infer_delimiter_from_text(sample)

    def _infer_delimiter_from_text(self, text: str) -> str:
        if "\t" in text:
            return "\t"
        if ";" in text:
            return ";"
        return ","

    def _write_scraped_csv_rows(
        self,
        csv_path: Path,
        header: List[str],
        rows: List[List[str]],
        *,
        delimiter: str = ",",
    ) -> bool:
        if not header:
            messagebox.showwarning(
                "Delete Row", "The table could not be saved because it does not include headers."
            )
            return False

        normalized_rows: List[List[str]] = []
        header_length = len(header)
        for row in rows:
            values = list(row)
            if len(values) < header_length:
                values.extend([""] * (header_length - len(values)))
            elif len(values) > header_length:
                values = values[:header_length]
            normalized_rows.append(values)

        try:
            with csv_path.open("w", encoding="utf-8", newline="") as fh:
                writer = csv.writer(fh, delimiter=delimiter)
                writer.writerow(header)
                writer.writerows(normalized_rows)
        except Exception as exc:
            messagebox.showwarning("Delete Row", f"Could not save updated CSV: {exc}")
            return False
        return True

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

    def _call_openai(self, api_key: str, prompt: str, page_text: str) -> str:
        thread_name = threading.current_thread().name
        sanitized_key = api_key.strip()
        if not sanitized_key:
            logger.error(
                "OpenAI call aborted | thread=%s | reason=missing_api_key",
                thread_name,
            )
            raise ValueError("OpenAI API key is missing")

        model_name = self.openai_model_var.get().strip() or DEFAULT_OPENAI_MODEL
        prompt_length = len(prompt)
        page_text_length = len(page_text)
        logger.info(
            "OpenAI call starting | thread=%s | model=%s | prompt_chars=%d | page_chars=%d",
            thread_name,
            model_name,
            prompt_length,
            page_text_length,
        )

        client_start = time.perf_counter()
        client = OpenAI(api_key=sanitized_key)
        logger.info(
            "OpenAI client initialized | thread=%s | elapsed_ms=%.2f",
            thread_name,
            (time.perf_counter() - client_start) * 1000,
        )

        request_started = time.perf_counter()
        logger.info(
            "OpenAI request initiated | thread=%s | model=%s | prompt_chars=%d | page_chars=%d",
            thread_name,
            model_name,
            prompt_length,
            page_text_length,
        )
        try:
            response = client.chat.completions.create(
                model=model_name,
                messages=[
                    {"role": "system", "content": prompt},
                    {"role": "user", "content": page_text},
                ],
            )
        except (APIConnectionError, APIError, APIStatusError, RateLimitError) as exc:
            logger.exception(
                "OpenAI request failed | model=%s | status=error | detail=%s",
                model_name,
                exc,
            )
            raise

        usage = getattr(response, "usage", None)
        prompt_tokens = getattr(usage, "prompt_tokens", None) if usage else None
        completion_tokens = getattr(usage, "completion_tokens", None) if usage else None
        total_tokens = getattr(usage, "total_tokens", None) if usage else None
        logger.info(
            "OpenAI response received | thread=%s | model=%s | status=success | completion_id=%s | "
            "prompt_tokens=%s | completion_tokens=%s | total_tokens=%s | latency_ms=%.2f",
            thread_name,
            model_name,
            getattr(response, "id", ""),
            prompt_tokens,
            completion_tokens,
            total_tokens,
            (time.perf_counter() - request_started) * 1000,
        )

        choices = getattr(response, "choices", None)
        if not choices:
            logger.error("OpenAI response missing choices | model=%s", model_name)
            raise ValueError("No choices returned from OpenAI API")

        first_choice = choices[0]
        message = getattr(first_choice, "message", None) or {}
        content = message.get("content") if isinstance(message, dict) else getattr(message, "content", None)
        if not content:
            logger.error("OpenAI response missing content | model=%s", model_name)
            raise ValueError("Empty response from OpenAI API")
        return str(content)

    def _strip_code_fence(self, text: str) -> str:
        fence_match = re.search(r"```(?:[^`\n]*)\n([\s\S]*?)```", text)
        if fence_match:
            return fence_match.group(1)
        return text

    def _convert_response_to_rows(self, response: str) -> List[List[str]]:
        cleaned = self._strip_code_fence(response)
        lines = [line.strip() for line in cleaned.splitlines() if line.strip()]
        if not lines:
            return self._normalize_scraped_rows([["response"], [""]])

        if any("\t" in line for line in lines):
            rows: List[List[str]] = []
            max_columns = 0
            for line in lines:
                parts = [segment.strip() for segment in line.split("\t")]
                rows.append(parts)
                if len(parts) > max_columns:
                    max_columns = len(parts)
            if max_columns:
                for row in rows:
                    if len(row) < max_columns:
                        row.extend([""] * (max_columns - len(row)))
            if rows:
                return self._normalize_scraped_rows(rows)

        if all(
            line.startswith("|") and line.endswith("|") and "|" in line.strip("|") for line in lines
        ):
            rows: List[List[str]] = []
            for idx, line in enumerate(lines):
                segments = [segment.strip() for segment in line.strip("|").split("|")]
                if idx == 1 and all(set(seg) <= {"-", ":"} for seg in segments):
                    continue
                rows.append(segments)
            if rows:
                return self._normalize_scraped_rows(rows)

        if any(
            ";" in line and not line.startswith("#") for line in lines
        ):
            rows = [[segment.strip() for segment in line.split(";")] for line in lines]
            return self._normalize_scraped_rows(rows)

        if any(
            "," in line and not line.startswith("#") for line in lines
        ):
            rows = [[segment.strip() for segment in line.split(",")] for line in lines]
            return self._normalize_scraped_rows(rows)

        fallback = [["response"], *[[line] for line in lines]]
        return self._normalize_scraped_rows(fallback)

    def _write_csv_rows(self, rows: List[List[str]], csv_path: Path) -> None:
        csv_path.parent.mkdir(parents=True, exist_ok=True)
        with csv_path.open("w", encoding="utf-8", newline="") as fh:
            writer = csv.writer(fh, quoting=csv.QUOTE_ALL)
            for row in rows:
                writer.writerow(row)

    def _normalize_scraped_rows(self, rows: List[List[Any]]) -> List[List[str]]:
        normalized: List[List[str]] = []
        for row in rows:
            normalized.append([self._normalize_scraped_cell(cell) for cell in row])
        return normalized

    def _normalize_scraped_cell(self, value: Any) -> str:
        if value is None:
            return ""
        original_text = str(value)
        if not original_text:
            return ""
        trimmed = original_text.strip()
        if not trimmed:
            return ""

        unwrapped = trimmed
        changed = False
        # remove balanced quoting repeatedly
        while len(unwrapped) >= 2 and unwrapped[0] == unwrapped[-1] == '"':
            inner = unwrapped[1:-1]
            if inner == unwrapped:
                break
            unwrapped = inner
            changed = True

        collapsed = unwrapped.replace('""', '"')
        if collapsed != unwrapped:
            changed = True
        collapsed_escaped = collapsed.replace('\\"', '"')
        if collapsed_escaped != collapsed:
            changed = True

        if changed:
            return collapsed_escaped
        return trimmed

    def _load_local_config(self) -> None:
        self.local_config_data = {}
        if not self.local_config_path.exists():
            return
        try:
            with self.local_config_path.open("r", encoding="utf-8") as fh:
                data = json.load(fh)
        except Exception as exc:
            messagebox.showwarning("Load Local Config", f"Could not load local configuration: {exc}")
            return
        if not isinstance(data, dict):
            messagebox.showwarning(
                "Load Local Config",
                "Local configuration file is invalid; ignoring its contents.",
            )
            return

        api_key_value = str(data.get("api_key", "")).strip()
        downloads_dir_value = str(data.get("downloads_dir", "")).strip()

        if api_key_value:
            self.local_config_data["api_key"] = api_key_value
        self.api_key_var.set(api_key_value)

        if downloads_dir_value:
            self.local_config_data["downloads_dir"] = downloads_dir_value
            self.downloads_dir_var.set(downloads_dir_value)

    def _write_local_config(self) -> None:
        payload: Dict[str, Any] = {}
        api_key_value = str(self.local_config_data.get("api_key", "")).strip()
        downloads_dir_value = str(self.local_config_data.get("downloads_dir", "")).strip()

        if api_key_value:
            payload["api_key"] = api_key_value
        if downloads_dir_value:
            payload["downloads_dir"] = downloads_dir_value

        if not payload:
            if self.local_config_path.exists():
                try:
                    self.local_config_path.unlink()
                except Exception as exc:
                    messagebox.showwarning("Save Local Config", f"Could not remove local configuration: {exc}")
            return

        try:
            with self.local_config_path.open("w", encoding="utf-8") as fh:
                json.dump(payload, fh, indent=2)
        except Exception as exc:
            messagebox.showwarning("Save Local Config", f"Could not save local configuration: {exc}")

    def _ensure_unique_path(self, path: Path) -> Path:
        if not path.exists():
            return path
        counter = 2
        stem = path.stem
        suffix = path.suffix
        while True:
            candidate = path.with_name(f"{stem}_{counter}{suffix}")
            if not candidate.exists():
                return candidate
            counter += 1

    def _gather_patterns(self) -> Tuple[Dict[str, List[re.Pattern[str]]], List[re.Pattern[str]]]:
        pattern_map: Dict[str, List[re.Pattern[str]]] = {}
        raw_map: Dict[str, List[str]] = {}
        case_flags: Dict[str, bool] = {}
        whitespace_flags: Dict[str, bool] = {}
        for column in COLUMNS:
            text_widget = self.pattern_texts[column]
            raw_text = text_widget.get("1.0", tk.END)
            patterns = [line.strip() for line in raw_text.splitlines() if line.strip()]
            if not patterns:
                patterns = DEFAULT_PATTERNS[column]
                text_widget.delete("1.0", tk.END)
                text_widget.insert("1.0", "\n".join(patterns))
            compiled = []
            flags = re.IGNORECASE if self.case_insensitive_vars[column].get() else 0
            for pattern in patterns:
                try:
                    compiled_pattern = self._apply_whitespace_option(pattern, self.whitespace_as_space_vars[column].get())
                    compiled.append(re.compile(compiled_pattern, flags))
                except re.error as exc:
                    messagebox.showerror("Invalid Pattern", f"Invalid regex '{pattern}' for {column}: {exc}")
            pattern_map[column] = compiled
            raw_map[column] = patterns
            case_flags[column] = self.case_insensitive_vars[column].get()
            whitespace_flags[column] = self.whitespace_as_space_vars[column].get()

        year_patterns: List[str] = []
        year_compiled: List[re.Pattern[str]] = []
        if self.year_pattern_text is not None:
            raw_text = self.year_pattern_text.get("1.0", tk.END)
            year_patterns = [line.strip() for line in raw_text.splitlines() if line.strip()]
        if not year_patterns:
            year_patterns = YEAR_DEFAULT_PATTERNS
            if self.year_pattern_text is not None:
                self.year_pattern_text.delete("1.0", tk.END)
                self.year_pattern_text.insert("1.0", "\n".join(year_patterns))
        year_flags = re.IGNORECASE if self.year_case_insensitive_var.get() else 0
        for pattern in year_patterns:
            try:
                compiled_pattern = self._apply_whitespace_option(pattern, self.year_whitespace_as_space_var.get())
                year_compiled.append(re.compile(compiled_pattern, year_flags))
            except re.error as exc:
                messagebox.showerror("Invalid Pattern", f"Invalid year regex '{pattern}': {exc}")
        self._save_pattern_config(
            raw_map,
            year_patterns,
            case_flags,
            self.year_case_insensitive_var.get(),
            whitespace_flags,
            self.year_whitespace_as_space_var.get(),
        )
        return pattern_map, year_compiled

    def _save_pattern_config(
        self,
        patterns: Dict[str, List[str]],
        year_patterns: List[str],
        case_flags: Dict[str, bool],
        year_case_flag: bool,
        whitespace_flags: Dict[str, bool],
        year_whitespace_flag: bool,
    ) -> None:
        self.config_data.update(
            {
                "patterns": patterns,
                "case_insensitive": case_flags,
                "year_patterns": year_patterns,
                "year_case_insensitive": year_case_flag,
                "space_as_whitespace": whitespace_flags,
                "year_space_as_whitespace": year_whitespace_flag,
            }
        )
        current_company = self.company_var.get()
        if current_company:
            self.config_data["last_company"] = current_company
        self._write_config()

    def _load_pattern_config(self) -> None:
        if not self.pattern_config_path.exists():
            self._config_loaded = True
            self._ensure_download_settings()
            return
        try:
            with self.pattern_config_path.open("r", encoding="utf-8") as fh:
                raw_data = json.load(fh)
        except Exception as exc:  # pragma: no cover - guard for IO issues
            messagebox.showwarning("Load Patterns", f"Could not load pattern configuration: {exc}")
            self._config_loaded = True
            return

        if not isinstance(raw_data, dict):
            messagebox.showwarning(
                "Load Patterns",
                "Pattern configuration file is invalid; ignoring its contents.",
            )
            data: Dict[str, Any] = {}
        else:
            data = dict(raw_data)

        migrated_local_settings = False
        raw_api_key = data.pop("api_key", None)
        if isinstance(raw_api_key, str):
            trimmed_key = raw_api_key.strip()
            if not self.api_key_var.get().strip():
                self.api_key_var.set(trimmed_key)
            if trimmed_key:
                if self.local_config_data.get("api_key") != trimmed_key:
                    self.local_config_data["api_key"] = trimmed_key
                    migrated_local_settings = True
            elif "api_key" in self.local_config_data:
                self.local_config_data.pop("api_key", None)
                migrated_local_settings = True

        raw_downloads_dir = data.pop("downloads_dir", None)
        if isinstance(raw_downloads_dir, str):
            trimmed_dir = raw_downloads_dir.strip()
            if not self.downloads_dir_var.get().strip() and trimmed_dir:
                self.downloads_dir_var.set(trimmed_dir)
            if trimmed_dir:
                if self.local_config_data.get("downloads_dir") != trimmed_dir:
                    self.local_config_data["downloads_dir"] = trimmed_dir
                    migrated_local_settings = True
            elif "downloads_dir" in self.local_config_data:
                self.local_config_data.pop("downloads_dir", None)
                migrated_local_settings = True

        if migrated_local_settings:
            self._write_local_config()

        self.config_data = data
        try:
            stored_widths = data.get("combined_base_column_widths")
            if isinstance(stored_widths, dict):
                valid_widths: Dict[str, int] = {}
                for column_name in ("Type", "Category", "Item", "Note"):
                    width_value = stored_widths.get(column_name)
                    if isinstance(width_value, (int, float)) and width_value > 0:
                        valid_widths[column_name] = int(width_value)
                self.combined_base_column_widths = valid_widths
            else:
                self.combined_base_column_widths = {}
            other_width_value = data.get("combined_other_column_width")
            if isinstance(other_width_value, (int, float)) and other_width_value > 0:
                self.combined_other_column_width = int(other_width_value)
            else:
                self.combined_other_column_width = None
            zoom_value = data.get("combined_preview_zoom")
            if isinstance(zoom_value, (int, float)):
                zoom = max(0.5, min(3.0, float(zoom_value)))
                self.combined_preview_zoom_var.set(zoom)
            self._update_combined_preview_zoom_display()
            self._load_type_category_colors_from_config(data)
        except Exception as exc:  # pragma: no cover - guard for IO issues
            messagebox.showwarning("Load Patterns", f"Could not load pattern configuration: {exc}")
            self._config_loaded = True
            return

        patterns = data.get("patterns", {})
        case_flags = data.get("case_insensitive", {})
        whitespace_flags = data.get("space_as_whitespace", {})
        for column, text_widget in self.pattern_texts.items():
            column_patterns = patterns.get(column)
            if column_patterns:
                text_widget.delete("1.0", tk.END)
                text_widget.insert("1.0", "\n".join(column_patterns))
            if column in case_flags and column in self.case_insensitive_vars:
                self.case_insensitive_vars[column].set(bool(case_flags[column]))
            if column in whitespace_flags and column in self.whitespace_as_space_vars:
                self.whitespace_as_space_vars[column].set(bool(whitespace_flags[column]))

        year_patterns = data.get("year_patterns")
        if year_patterns and self.year_pattern_text is not None:
            self.year_pattern_text.delete("1.0", tk.END)
            self.year_pattern_text.insert("1.0", "\n".join(year_patterns))
        if "year_case_insensitive" in data:
            self.year_case_insensitive_var.set(bool(data["year_case_insensitive"]))
        if "year_space_as_whitespace" in data:
            self.year_whitespace_as_space_var.set(bool(data["year_space_as_whitespace"]))
        self.last_company_preference = data.get("last_company", "")
        raw_model = data.get("openai_model")
        if isinstance(raw_model, str) and raw_model.strip():
            self.openai_model_var.set(raw_model.strip())
        else:
            self.openai_model_var.set(DEFAULT_OPENAI_MODEL)
        self._ensure_download_settings()
        self._apply_last_company_selection()
        self._config_loaded = True

    def _load_type_category_colors_from_config(self, data: Any) -> None:
        self.type_color_map = {}
        self.type_color_labels = {}
        self.category_color_map = {}
        self.category_color_labels = {}

        if isinstance(data, dict):
            type_raw = data.get("combined_type_colors")
            category_raw = data.get("combined_category_colors")
            legacy_raw = data.get("combined_type_category_colors")
        else:
            type_raw = None
            category_raw = None
            legacy_raw = None

        if legacy_raw and not type_raw and not category_raw:
            type_raw = legacy_raw
            category_raw = legacy_raw

        def _iter_entries(raw: Any) -> List[Tuple[Any, Any]]:
            if isinstance(raw, dict):
                return list(raw.items())
            if isinstance(raw, list):
                entries: List[Tuple[Any, Any]] = []
                for entry in raw:
                    if isinstance(entry, dict):
                        entries.append((entry.get("value"), entry.get("color")))
                return entries
            return []

        def _populate(
            target_map: Dict[str, str], target_labels: Dict[str, str], raw: Any
        ) -> None:
            for key, color in _iter_entries(raw):
                normalized_key = self._normalize_type_category_value(key)
                normalized_color = (
                    self._normalize_hex_color(str(color)) if color is not None else None
                )
                if not normalized_key or not normalized_color:
                    continue
                label_text = str(key).strip() if isinstance(key, str) else str(key)
                target_map[normalized_key] = normalized_color
                target_labels[normalized_key] = label_text

        _populate(self.type_color_map, self.type_color_labels, type_raw)
        _populate(self.category_color_map, self.category_color_labels, category_raw)

    def _apply_configured_note_key_bindings(self) -> None:
        note_order: List[str] = []
        seen: Set[str] = set()
        display_labels: Dict[str, str] = {}

        def _add_option(value: Any, *, display: Optional[str] = None) -> str:
            if value is None:
                normalized_value = ""
                source_text = ""
            elif isinstance(value, str):
                normalized_value = value.strip().lower()
                source_text = value.strip()
            else:
                normalized_value = str(value).strip().lower()
                source_text = str(value).strip()
            if not normalized_value:
                normalized_value = ""
            if normalized_value in seen:
                if display:
                    display_labels.setdefault(normalized_value, display)
                elif source_text:
                    display_labels.setdefault(normalized_value, source_text)
                return normalized_value
            if normalized_value:
                note_order.append(normalized_value)
            else:
                note_order.insert(0, "")
            seen.add(normalized_value)
            label_text = display
            if not label_text:
                if normalized_value in DEFAULT_NOTE_LABELS:
                    label_text = DEFAULT_NOTE_LABELS[normalized_value]
                elif source_text:
                    label_text = source_text
                elif not normalized_value:
                    label_text = "Clear note"
                else:
                    label_text = normalized_value
            display_labels[normalized_value] = label_text
            return normalized_value

        configured_colors: Dict[str, str] = {}
        configured_bindings: Dict[str, str] = {}

        _add_option("", display=DEFAULT_NOTE_LABELS.get(""))

        raw_settings = self.config_data.get("combined_note_settings")
        if isinstance(raw_settings, list):
            for entry in raw_settings:
                if not isinstance(entry, dict):
                    continue
                raw_value = entry.get("normalized", entry.get("value", ""))
                display_value = (
                    entry.get("label")
                    or entry.get("display")
                    or entry.get("name")
                )
                normalized_value = _add_option(raw_value, display=display_value)
                color = entry.get("color")
                normalized_color = self._normalize_hex_color(str(color)) if color else None
                if normalized_color:
                    configured_colors[normalized_value] = normalized_color
                shortcut_value = entry.get("shortcut") or entry.get("key") or entry.get("binding")
                normalized_shortcut = (
                    self._normalize_note_binding_value(str(shortcut_value)) if shortcut_value else ""
                )
                if normalized_value not in configured_bindings or normalized_shortcut:
                    configured_bindings[normalized_value] = normalized_shortcut

        raw_bindings = self.config_data.get("combined_note_key_bindings")
        if isinstance(raw_bindings, dict):
            for value, shortcut in raw_bindings.items():
                normalized_value = _add_option(value)
                normalized_shortcut = (
                    self._normalize_note_binding_value(str(shortcut)) if shortcut else ""
                )
                if normalized_value not in configured_bindings or normalized_shortcut:
                    configured_bindings[normalized_value] = normalized_shortcut

        for default_value in DEFAULT_NOTE_OPTIONS:
            _add_option(default_value, display=DEFAULT_NOTE_LABELS.get(default_value))

        self.note_options = note_order
        defaults_colors = DEFAULT_NOTE_BACKGROUND_COLORS.copy()
        self.note_background_colors = {}
        for value in note_order:
            color = configured_colors.get(value, defaults_colors.get(value, "")) or ""
            self.note_background_colors[value] = color
        for value, color in configured_colors.items():
            if value not in self.note_background_colors:
                self.note_background_colors[value] = color

        defaults_keys = DEFAULT_NOTE_KEY_BINDINGS.copy()
        self.note_key_bindings = {}
        for value in note_order:
            if value in configured_bindings:
                self.note_key_bindings[value] = configured_bindings[value] or ""
            else:
                self.note_key_bindings[value] = defaults_keys.get(value, "") or ""
        for value, shortcut in configured_bindings.items():
            if value not in self.note_key_bindings:
                self.note_key_bindings[value] = shortcut or ""

        self.note_display_labels = {}
        for value in note_order:
            label_text = display_labels.get(value)
            if not label_text:
                if value in DEFAULT_NOTE_LABELS:
                    label_text = DEFAULT_NOTE_LABELS[value]
                elif value:
                    label_text = value
                else:
                    label_text = "Clear note"
            self.note_display_labels[value] = label_text
        for value, label_text in display_labels.items():
            if value not in self.note_display_labels:
                self.note_display_labels[value] = label_text

        self._update_note_settings_config_entries()
        self._refresh_note_tags()

    def _apply_whitespace_option(self, pattern: str, enabled: bool) -> str:
        if not enabled:
            return pattern
        return pattern.replace(" ", r"\s+")

    def _ensure_download_settings(self) -> None:
        configured_dir = str(self.local_config_data.get("downloads_dir", "")).strip()
        if configured_dir:
            self.downloads_dir_var.set(configured_dir)
        elif not self.downloads_dir_var.get():
            default_download_dir = Path.home() / "Downloads"
            if default_download_dir.exists():
                self.downloads_dir_var.set(str(default_download_dir))
        current_dir = self.downloads_dir_var.get().strip()
        if current_dir and self.local_config_data.get("downloads_dir") != current_dir:
            self.local_config_data["downloads_dir"] = current_dir
            self._write_local_config()

        configured_minutes = self.config_data.get("downloads_minutes")
        if isinstance(configured_minutes, int) and configured_minutes > 0:
            self.recent_download_minutes_var.set(configured_minutes)
        else:
            minutes_value = self.recent_download_minutes_var.get()
            if minutes_value <= 0:
                minutes_value = 5
                self.recent_download_minutes_var.set(minutes_value)
            self.config_data["downloads_minutes"] = minutes_value

    def _maximize_window(self, window: Optional[tk.Misc] = None) -> None:
        target = window or self.root
        try:
            target.state("zoomed")
            return
        except tk.TclError:
            pass
        try:
            target.attributes("-zoomed", True)
            return
        except tk.TclError:
            pass
        screen_width = target.winfo_screenwidth()
        screen_height = target.winfo_screenheight()
        target.geometry(f"{screen_width}x{screen_height}+0+0")

    def _apply_last_company_selection(self) -> None:
        if self.embedded or not hasattr(self, "company_combo"):
            return
        if not self.last_company_preference:
            return
        companies = list(self.company_combo["values"])
        if self.last_company_preference in companies:
            self.company_combo.set(self.last_company_preference)
            self.company_var.set(self.last_company_preference)
            self._set_folder_from_company(self.last_company_preference)

    def _update_last_company(self, company: str) -> None:
        if not company:
            return
        if self.config_data.get("last_company") == company:
            return
        self.config_data["last_company"] = company
        self.last_company_preference = company
        self._write_config()

    def _update_note_settings_config_entries(self) -> None:
        sanitized_order: List[str] = []
        seen: Set[str] = set()
        for value in self.note_options:
            if isinstance(value, str):
                normalized_value = value.strip().lower()
            else:
                normalized_value = str(value).strip().lower()
            if not normalized_value:
                normalized_value = ""
            if normalized_value in seen:
                continue
            if normalized_value:
                sanitized_order.append(normalized_value)
            else:
                sanitized_order.insert(0, "")
            seen.add(normalized_value)
        if "" not in seen:
            sanitized_order.insert(0, "")
            seen.add("")

        sanitized_colors: Dict[str, str] = {}
        sanitized_bindings: Dict[str, str] = {}
        sanitized_labels: Dict[str, str] = {}
        for value in sanitized_order:
            raw_color = self.note_background_colors.get(value, DEFAULT_NOTE_BACKGROUND_COLORS.get(value, ""))
            normalized_color = self._normalize_hex_color(raw_color) if raw_color else None
            sanitized_colors[value] = normalized_color or ""
            stored_shortcut = self.note_key_bindings.get(value)
            if stored_shortcut is None:
                stored_shortcut = DEFAULT_NOTE_KEY_BINDINGS.get(value, "")
            normalized_shortcut = (
                self._normalize_note_binding_value(str(stored_shortcut)) if stored_shortcut else ""
            )
            sanitized_bindings[value] = normalized_shortcut or ""
            label_text = self.note_display_labels.get(value)
            if not label_text:
                if value in DEFAULT_NOTE_LABELS:
                    label_text = DEFAULT_NOTE_LABELS[value]
                elif value:
                    label_text = value
                else:
                    label_text = "Clear note"
            sanitized_labels[value] = label_text

        self.note_options = sanitized_order
        self.note_background_colors = sanitized_colors
        self.note_key_bindings = sanitized_bindings
        self.note_display_labels = sanitized_labels

        settings_payload: List[Dict[str, str]] = []
        for value in sanitized_order:
            entry: Dict[str, str] = {"value": value, "normalized": value}
            shortcut = sanitized_bindings.get(value, "")
            if shortcut:
                entry["shortcut"] = shortcut
            color = sanitized_colors.get(value, "")
            if color:
                entry["color"] = color
            label_text = sanitized_labels.get(value, "")
            if label_text:
                entry["label"] = label_text
            settings_payload.append(entry)

        self.config_data["combined_note_settings"] = settings_payload
        self.config_data["combined_note_key_bindings"] = {
            value: sanitized_bindings.get(value, "") for value in sanitized_order
        }

    def _write_config(self) -> None:
        self._update_note_settings_config_entries()
        self.config_data.pop("api_key", None)
        self.config_data.pop("downloads_dir", None)
        model_name = self.openai_model_var.get().strip()
        if model_name and model_name != DEFAULT_OPENAI_MODEL:
            self.config_data["openai_model"] = model_name
        else:
            self.config_data.pop("openai_model", None)
        self.config_data["combined_base_column_widths"] = {
            column: int(width)
            for column, width in self.combined_base_column_widths.items()
            if column in {"Type", "Category", "Item", "Note"} and width > 0
        }
        if isinstance(self.combined_other_column_width, int) and self.combined_other_column_width > 0:
            self.config_data["combined_other_column_width"] = int(self.combined_other_column_width)
        else:
            self.config_data.pop("combined_other_column_width", None)
        zoom_value = float(self.combined_preview_zoom_var.get())
        if 0.5 <= zoom_value <= 3.0:
            self.config_data["combined_preview_zoom"] = round(zoom_value, 3)
        else:
            self.config_data.pop("combined_preview_zoom", None)
        def _build_store(
            color_map: Dict[str, str], labels: Dict[str, str]
        ) -> Dict[str, str]:
            stored: Dict[str, str] = {}
            for normalized, color in color_map.items():
                normalized_color = self._normalize_hex_color(color)
                if not normalized_color:
                    continue
                label = labels.get(normalized, normalized)
                stored[label] = normalized_color
            return stored

        type_store = _build_store(self.type_color_map, self.type_color_labels)
        category_store = _build_store(self.category_color_map, self.category_color_labels)

        if type_store:
            self.config_data["combined_type_colors"] = type_store
        else:
            self.config_data.pop("combined_type_colors", None)

        if category_store:
            self.config_data["combined_category_colors"] = category_store
        else:
            self.config_data.pop("combined_category_colors", None)

        self.config_data.pop("combined_type_category_colors", None)
        try:
            with self.pattern_config_path.open("w", encoding="utf-8") as fh:
                json.dump(self.config_data, fh, indent=2)
        except Exception as exc:  # pragma: no cover - guard for IO issues
            messagebox.showwarning("Save Patterns", f"Could not save pattern configuration: {exc}")

    def scrape_selections(self) -> None:
        logger.info("AIScrape invoked | pdf_entries=%d", len(self.pdf_entries))
        if not self.pdf_entries:
            messagebox.showinfo("No PDFs", "Load PDFs before running AIScrape.")
            logger.info("AIScrape aborted | reason=no_pdfs")
            return

        company = self.company_var.get()
        if not company:
            messagebox.showinfo("Select Company", "Please choose a company before running AIScrape.")
            logger.info("AIScrape aborted | reason=no_company_selected")
            return

        if hasattr(self, "scrape_progress"):
            self.scrape_progress["value"] = 0

        api_key = self.api_key_var.get().strip()
        if not api_key:
            messagebox.showwarning("API Key Required", "Enter an API key and press Enter before running AIScrape.")
            self.api_key_entry.focus_set()
            logger.info("AIScrape aborted | reason=no_api_key")
            return
        if api_key != self.api_key_var.get():
            self.api_key_var.set(api_key)
        prompts: Dict[str, str] = {}
        missing_prompts: List[str] = []
        for category in COLUMNS:
            prompt_text = self._get_prompt_text(company, category)
            if prompt_text is None:
                missing_prompts.append(category)
            else:
                prompts[category] = prompt_text
        if missing_prompts:
            messagebox.showerror(
                "Missing Prompts",
                "Prompt files not found for: " + ", ".join(missing_prompts),
            )
            logger.info(
                "AIScrape aborted | reason=missing_prompts | missing=%s",
                ",".join(sorted(missing_prompts)),
            )
            return

        self._save_api_key()

        scrape_root = self.companies_dir / company / "openapiscrape"
        scrape_root.mkdir(parents=True, exist_ok=True)
        logger.info(
            "AIScrape preparing jobs | company=%s | scrape_root=%s",
            company,
            scrape_root,
        )

        errors: List[str] = []
        pending: List[Tuple[PDFEntry, List[Tuple[str, List[int]]]]] = []
        metadata_cache: Dict[Path, Dict[str, Any]] = {}
        metadata_path_map: Dict[Path, Path] = {}

        for entry in self.pdf_entries:
            entry_tasks: List[Tuple[str, List[int]]] = []
            entry_stem = entry.stem
            metadata_path = self._metadata_file_for_entry(scrape_root, entry_stem)
            if metadata_path.exists():
                metadata_data = self._load_doc_metadata(metadata_path)
            else:
                legacy_dir = self._legacy_doc_dir(scrape_root, entry_stem)
                if legacy_dir.exists():
                    metadata_data = self._load_doc_metadata(legacy_dir)
                else:
                    metadata_data = {}
            if not isinstance(metadata_data, dict):
                metadata_data = {}
            metadata = dict(metadata_data)
            metadata_cache[entry.path] = metadata
            metadata_path_map[entry.path] = metadata_path
            for category in COLUMNS:
                page_indexes = self._get_highlight_page_numbers(entry, category)
                if not page_indexes:
                    page_index = self._get_selected_page_index(entry, category)
                    if page_index is None:
                        continue
                    page_indexes = [int(page_index)]
                normalized_indexes = sorted({int(idx) for idx in page_indexes})
                if not normalized_indexes:
                    continue
                if not prompts.get(category):
                    continue
                existing = metadata.get(category)
                file_exists = False
                if isinstance(existing, dict):
                    csv_name = existing.get("csv")
                    if isinstance(csv_name, str) and csv_name:
                        csv_path = self._resolve_scrape_path(
                            scrape_root, entry_stem, csv_name
                        )
                        if csv_path.exists():
                            file_exists = True
                    else:
                        txt_name = existing.get("txt")
                        if isinstance(txt_name, str) and txt_name:
                            txt_path = self._resolve_scrape_path(
                                scrape_root, entry_stem, txt_name
                            )
                            if txt_path.exists():
                                file_exists = True
                if file_exists:
                    continue
                entry_tasks.append((category, normalized_indexes))

            if entry_tasks:
                logger.info(
                    "AIScrape pending entry | entry=%s | year=%s | tasks=%d",
                    entry.path.name,
                    entry.year,
                    len(entry_tasks),
                )
                pending.append((entry, entry_tasks))

        total_tasks = sum(len(items) for _, items in pending)
        if not total_tasks:
            messagebox.showinfo("AIScrape", "All selected sections already have scraped files.")
            logger.info("AIScrape aborted | reason=no_pending_tasks")
            return

        if hasattr(self, "scrape_button"):
            self.scrape_button.state(["disabled"])
        if hasattr(self, "scrape_progress"):
            self.scrape_progress.configure(maximum=total_tasks, value=0)
        logger.info("AIScrape starting | jobs=%d", total_tasks)

        successful = 0
        attempted = 0

        metadata_changed: Dict[Path, bool] = {}
        jobs: List[ScrapeTask] = []

        def _run_job(job: ScrapeTask) -> Tuple[bool, Optional[Dict[str, Any]], Optional[str]]:
            logger.info(
                "AIScrape job started | entry=%s | category=%s | page_count=%d",
                job.entry_name,
                job.category,
                len(job.page_indexes),
            )
            try:
                response_text = self._call_openai(api_key, job.prompt_text, job.page_text)
            except (APIConnectionError, APIError, APIStatusError, RateLimitError) as exc:
                logger.error(
                    "OpenAI request failed | entry=%s | category=%s | detail=%s",
                    job.entry_name,
                    job.category,
                    exc,
                )
                return False, None, f"{job.entry_name} - {job.category}: API request failed ({exc})"
            except Exception as exc:
                logger.error(
                    "OpenAI processing error | entry=%s | category=%s | detail=%s",
                    job.entry_name,
                    job.category,
                    exc,
                )
                return False, None, f"{job.entry_name} - {job.category}: {exc}"

            pdf_stem = Path(job.entry_name).stem
            csv_name = self._build_scrape_filename(pdf_stem, job.category)
            csv_path = job.scrape_root / csv_name

            try:
                csv_path.write_text(response_text, encoding="utf-8")
            except Exception as exc:
                return False, None, f"{job.entry_name} - {job.category}: Could not save response ({exc})"

            metadata_entry = {
                "csv": csv_path.name,
                "page_index": job.page_indexes[0] if job.page_indexes else None,
                "page_indexes": list(job.page_indexes),
                "year": job.entry_year,
            }
            logger.info(
                "AIScrape job completed | entry=%s | category=%s | response_chars=%d",
                job.entry_name,
                job.category,
                len(response_text),
            )
            return True, metadata_entry, None

        try:
            for entry, tasks in pending:
                metadata = metadata_cache.get(entry.path, {})
                if not isinstance(metadata, dict):
                    metadata = {}
                metadata = dict(metadata)
                metadata_cache[entry.path] = metadata
                metadata_changed.setdefault(entry.path, False)
                for category, page_indexes in tasks:
                    removed = metadata.pop(category, None)
                    if removed is not None:
                        metadata_changed[entry.path] = True
                    prompt_text = prompts.get(category)
                    if not prompt_text:
                        attempted += 1
                        if hasattr(self, "scrape_progress"):
                            self.scrape_progress["value"] = attempted
                            self.root.update()
                        continue
                    page_text_parts: List[str] = []
                    successful_pages: List[int] = []
                    for page_index in page_indexes:
                        try:
                            page = entry.doc.load_page(page_index)
                            page_text_parts.append(page.get_text("text"))
                            successful_pages.append(int(page_index))
                        except Exception as exc:
                            errors.append(
                                f"{entry.path.name} - {category}: Could not read page {page_index + 1} ({exc})"
                            )
                    if not successful_pages:
                        attempted += 1
                        if hasattr(self, "scrape_progress"):
                            self.scrape_progress["value"] = attempted
                            self.root.update()
                        continue
                    combined_text = "\n\n".join(part for part in page_text_parts if part)
                    jobs.append(
                        ScrapeTask(
                            entry_path=entry.path,
                            entry_name=entry.path.name,
                            entry_year=entry.year,
                            category=category,
                            page_indexes=successful_pages,
                            prompt_text=prompt_text,
                            page_text=combined_text,
                            scrape_root=scrape_root,
                        )
                    )
                    logger.info(
                        "AIScrape job queued | entry=%s | category=%s | pages=%s",
                        entry.path.name,
                        category,
                        ",".join(str(idx) for idx in successful_pages),
                    )

            if jobs:
                max_workers = min(8, max(2, os.cpu_count() or 4))
                logger.info("AIScrape executing | queued_jobs=%d | max_workers=%d", len(jobs), max_workers)
                with ThreadPoolExecutor(max_workers=max_workers) as executor:
                    future_map = {executor.submit(_run_job, job): job for job in jobs}
                    for future in as_completed(future_map):
                        job = future_map[future]
                        success = False
                        metadata_entry: Optional[Dict[str, Any]] = None
                        error_message: Optional[str] = None
                        try:
                            success, metadata_entry, error_message = future.result()
                        except Exception as exc:
                            error_message = f"{job.entry_name} - {job.category}: {exc}"
                            logger.exception(
                                "AIScrape job crashed | entry=%s | category=%s",
                                job.entry_name,
                                job.category,
                            )

                        attempted += 1
                        logger.info(
                            "AIScrape job finished | entry=%s | category=%s | success=%s | attempted=%d/%d",
                            job.entry_name,
                            job.category,
                            success,
                            attempted,
                            total_tasks,
                        )

                        if success and metadata_entry is not None:
                            metadata = metadata_cache.get(job.entry_path, {})
                            if not isinstance(metadata, dict):
                                metadata = {}
                                metadata_cache[job.entry_path] = metadata
                            metadata[job.category] = metadata_entry
                            metadata_changed[job.entry_path] = True
                            successful += 1
                            logger.info(
                                "AIScrape job success | entry=%s | category=%s | successful=%d",
                                job.entry_name,
                                job.category,
                                successful,
                            )
                        elif error_message:
                            errors.append(error_message)
                            logger.warning(
                                "AIScrape job error | entry=%s | category=%s | message=%s",
                                job.entry_name,
                                job.category,
                                error_message,
                            )

                        if hasattr(self, "scrape_progress"):
                            self.scrape_progress["value"] = attempted
                            self.root.update()

            for entry_path, changed in metadata_changed.items():
                if not changed:
                    continue
                metadata = metadata_cache.get(entry_path, {})
                if not isinstance(metadata, dict):
                    continue
                metadata_path = metadata_path_map.get(entry_path)
                if metadata_path is None:
                    continue
                logger.info(
                    "AIScrape writing metadata | entry=%s | path=%s",
                    entry_path.name,
                    metadata_path,
                )
                self._write_doc_metadata(metadata_path, metadata)
        finally:
            if hasattr(self, "scrape_button"):
                self.scrape_button.state(["!disabled"])
            if hasattr(self, "scrape_progress"):
                self.scrape_progress["value"] = 0
            logger.info(
                "AIScrape completed | attempted=%d | successful=%d | errors=%d",
                attempted,
                successful,
                len(errors),
            )

        self._refresh_scraped_tab()
        if successful:
            if hasattr(self, "notebook") and hasattr(self, "scraped_frame"):
                self.notebook.select(self.scraped_frame)
            messagebox.showinfo("AIScrape Complete", f"Saved {successful} OpenAI responses to 'openapiscrape'.")
        if errors:
            messagebox.showerror("AIScrape Issues", "\n".join(errors))


def main() -> None:
    root = tk.Tk()
    try:
        ReportApp(root)
    except RuntimeError:
        root.destroy()
        return
    root.mainloop()


if __name__ == "__main__":
    main()
