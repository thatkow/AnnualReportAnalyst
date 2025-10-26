"""Primary application class for the Annual Report Analyst UI."""

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
import tkinter as tk
import webbrowser
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta
from pathlib import Path
from tkinter import colorchooser, filedialog, messagebox, simpledialog, ttk
from tkinter import font as tkfont
from typing import Any, Dict, List, Optional, Set, Tuple

import fitz  # PyMuPDF
import requests
from PIL import Image, ImageTk

from .category_row import CategoryRow
from .collapsible_frame import CollapsibleFrame
from .config import (
    COLUMNS,
    DEFAULT_NOTE_BACKGROUND_COLORS,
    DEFAULT_NOTE_KEY_BINDINGS,
    DEFAULT_NOTE_OPTIONS,
    DEFAULT_PATTERNS,
    HEX_COLOR_RE,
    MAX_COMBINED_DATE_COLUMNS,
    SHIFT_MASK,
    SPECIAL_KEYSYM_ALIASES,
    YEAR_DEFAULT_PATTERNS,
)
from .match import Match
from .pdf_entry import PDFEntry
from .scrape_task import ScrapeTask


logger = logging.getLogger(__name__)

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
                self.root.title("Annual Report Analyst")
            except tk.TclError:
                pass
        self.folder_path = tk.StringVar(master=self.root)
        if folder_override is not None:
            self.folder_path.set(str(folder_override))
        self.company_var = tk.StringVar(master=self.root)
        if company_name:
            self.company_var.set(company_name)
        self.api_key_var = tk.StringVar(master=self.root)
        self.thumbnail_width_var = tk.IntVar(master=self.root, value=220)
        self.pattern_texts: Dict[str, tk.Text] = {}
        self.case_insensitive_vars: Dict[str, tk.BooleanVar] = {}
        self.whitespace_as_space_vars: Dict[str, tk.BooleanVar] = {}
        self.pdf_entries: List[PDFEntry] = []
        self.category_rows: Dict[tuple[Path, str], CategoryRow] = {}
        self.year_vars: Dict[Path, tk.StringVar] = {}
        self.year_pattern_text: Optional[tk.Text] = None
        self.year_case_insensitive_var = tk.BooleanVar(master=self.root, value=True)
        self.year_whitespace_as_space_var = tk.BooleanVar(master=self.root, value=True)
        self.companies_dir = Path(__file__).resolve().parent / "companies"
        self.prompts_dir = Path(__file__).resolve().parent / "prompts"
        self.pattern_config_path = Path(__file__).resolve().parent / "pattern_config.json"
        self.config_data: Dict[str, Any] = {}
        self.last_company_preference: str = ""
        self._config_loaded = False
        self.assigned_pages: Dict[str, Dict[str, Any]] = {}
        self.assigned_pages_path: Optional[Path] = None
        self.scraped_images: List[ImageTk.PhotoImage] = []
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
        self.combined_base_column_widths: Dict[str, int] = {}
        self.combined_other_column_width: Optional[int] = None
        self.combined_show_blank_notes_var = tk.BooleanVar(master=self.root, value=False)
        self.combined_result_message: Optional[str] = None
        self.type_category_color_map: Dict[str, str] = {}
        self.type_category_color_labels: Dict[str, str] = {}
        self.note_assignments: Dict[Tuple[str, str, str], str] = {}
        self.note_assignments_path: Optional[Path] = None
        self.note_options: List[str] = list(DEFAULT_NOTE_OPTIONS)
        self.note_background_colors: Dict[str, str] = DEFAULT_NOTE_BACKGROUND_COLORS.copy()
        self.note_key_bindings: Dict[str, str] = DEFAULT_NOTE_KEY_BINDINGS.copy()

        self._build_ui()
        self._load_pattern_config()
        self._apply_configured_note_key_bindings()
        if not self.embedded:
            self._maximize_window()
            self.root.after(0, self._load_pdfs_on_start)

    def _build_ui(self) -> None:
        if not self.embedded:
            self._create_menus()
            top_frame = ttk.Frame(self.root, padding=8)
            top_frame.pack(fill=tk.X)
            company_label = ttk.Label(top_frame, text="Company:")
            company_label.pack(side=tk.LEFT)
            self.company_combo = ttk.Combobox(top_frame, textvariable=self.company_var, state="readonly", width=30)
            self.company_combo.pack(side=tk.LEFT, padx=4)
            self.company_combo.bind("<<ComboboxSelected>>", self._on_company_selected)
            folder_entry = ttk.Entry(top_frame, textvariable=self.folder_path, width=60, state="readonly")
            folder_entry.pack(side=tk.LEFT, padx=4)
            load_button = ttk.Button(top_frame, text="Load PDFs", command=self._open_company_tab)
            load_button.pack(side=tk.LEFT, padx=4)
            self._refresh_company_options()
            self.company_notebook = ttk.Notebook(self.root)
            self.company_notebook.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)
            self.company_notebook.bind("<<NotebookTabChanged>>", self._on_company_tab_changed)
            return
        top_frame = ttk.Frame(self.root, padding=8)
        top_frame.pack(fill=tk.X)
        ttk.Label(top_frame, text=f"Company: {self.company_var.get() or 'Unknown'}", font=("TkDefaultFont", 11, "bold")).pack(side=tk.LEFT)
        folder_entry = ttk.Entry(top_frame, textvariable=self.folder_path, width=50, state="readonly")
        folder_entry.pack(side=tk.LEFT, padx=(8, 0), fill=tk.X, expand=True)
        ttk.Button(top_frame, text="Reload PDFs", command=self.load_pdfs).pack(side=tk.LEFT, padx=(8, 0))
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

        combined_container = ttk.Frame(notebook)
        self.combined_frame = combined_container
        notebook.add(combined_container, text="Combined")
        combined_controls = ttk.Frame(combined_container, padding=8)
        combined_controls.pack(fill=tk.X)
        ttk.Label(
            combined_controls,
            text="Review scraped headers, adjust the final column labels, and click Parse to build the combined table.",
        ).pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.combine_confirm_button = ttk.Button(
            combined_controls,
            text="Parse",
            command=self._confirm_combined_table,
            state="disabled",
        )
        self.combine_confirm_button.pack(side=tk.RIGHT)
        ttk.Button(
            combined_controls,
            text="Load Assignments",
            command=self._prompt_import_note_assignments,
        ).pack(side=tk.RIGHT, padx=(0, 8))
        ttk.Checkbutton(
            combined_controls,
            text="Show only blank notes",
            variable=self.combined_show_blank_notes_var,
            command=self._on_combined_show_blank_notes_toggle,
        ).pack(side=tk.RIGHT, padx=(0, 8))
        combined_canvas = tk.Canvas(combined_container)
        combined_canvas.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

        def _combined_canvas_yview(*args: Any) -> None:
            self._destroy_note_editor()
            combined_canvas.yview(*args)

        def _combined_canvas_xview(*args: Any) -> None:
            self._destroy_note_editor()
            combined_canvas.xview(*args)

        combined_vscroll = ttk.Scrollbar(combined_container, orient=tk.VERTICAL, command=_combined_canvas_yview)
        combined_vscroll.pack(side=tk.RIGHT, fill=tk.Y)
        combined_hscroll = ttk.Scrollbar(combined_container, orient=tk.HORIZONTAL, command=_combined_canvas_xview)
        combined_hscroll.pack(side=tk.BOTTOM, fill=tk.X)
        combined_canvas.configure(yscrollcommand=combined_vscroll.set, xscrollcommand=combined_hscroll.set)
        self.combined_canvas = combined_canvas
        self.combined_inner = ttk.Frame(combined_canvas)
        self.combined_window = combined_canvas.create_window((0, 0), window=self.combined_inner, anchor="nw")
        self.combined_inner.bind(
            "<Configure>",
            lambda _e: combined_canvas.configure(scrollregion=combined_canvas.bbox("all")),
        )
        self.combined_header_frame = ttk.Frame(self.combined_inner, padding=8)
        self.combined_header_frame.grid(row=0, column=0, sticky="nsew")
        self.combined_result_frame = ttk.Frame(self.combined_inner, padding=8)
        self.combined_result_frame.grid(row=1, column=0, sticky="nsew")
        self.combined_inner.columnconfigure(0, weight=1)
        self.combined_inner.rowconfigure(1, weight=1)

    def _on_frame_configure(self, _: tk.Event) -> None:  # type: ignore[override]
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event: tk.Event) -> None:  # type: ignore[override]
        self.canvas.itemconfigure(self.canvas_window, width=event.width)

    def _bind_mousewheel(self, _: tk.Event) -> None:  # type: ignore[override]
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _unbind_mousewheel(self, _: tk.Event) -> None:  # type: ignore[override]
        self.canvas.unbind_all("<MouseWheel>")

    def _on_mousewheel(self, event: tk.Event) -> None:  # type: ignore[override]
        if event.num == 4 or event.delta > 0:
            self.canvas.yview_scroll(-1, "units")
        elif event.num == 5 or event.delta < 0:
            self.canvas.yview_scroll(1, "units")

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
        self.config_data["api_key"] = api_key
        self._write_config()

    def _load_pdfs_on_start(self) -> None:
        if self.embedded:
            if self.folder_path.get():
                self.load_pdfs()
        else:
            if self.company_var.get():
                self._open_company_tab()

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
        menubar.add_cascade(label="File", menu=file_menu)

        configuration_menu = tk.Menu(menubar, tearoff=False)
        configuration_menu.add_command(label="Set Downloads Dir", command=self._set_downloads_dir)
        configuration_menu.add_command(
            label="Set Download Window (minutes)",
            command=self._set_download_window,
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
            label="Configure Type/Category Colors",
            command=self._configure_type_category_colors,
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
        self.config_data["downloads_dir"] = selected
        self._write_config()

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

        def _create_row(note_value: str, *, display_name: Optional[str] = None, allow_remove: bool = True) -> None:
            normalized = note_value.strip().lower() if note_value else ""
            display = display_name or label_map.get(normalized, normalized or "Clear note")
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
            if "" not in new_order:
                new_order.insert(0, "")
                new_bindings.setdefault("", "")
                new_colors.setdefault("", "")
            else:
                new_order = [""] + [value for value in new_order if value]
            sanitized_bindings = {value: new_bindings.get(value, "") for value in new_order}
            sanitized_colors = {value: new_colors.get(value, "") for value in new_order}
            self.note_options = new_order
            self.note_key_bindings = sanitized_bindings
            self.note_background_colors = sanitized_colors
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

    def _configure_type_category_colors(self) -> None:
        window = tk.Toplevel(self.root)
        window.title("Configure Type/Category Colors")
        window.transient(self.root)
        window.grab_set()

        container = ttk.Frame(window, padding=12)
        container.pack(fill=tk.BOTH, expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(1, weight=1)

        ttk.Label(
            container,
            text=(
                "Assign background colors to Type or Category values. "
                "Rows matching these values will use the configured color when no note color overrides it."
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

        working_map = dict(self.type_category_color_map)
        working_labels = dict(self.type_category_color_labels)

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
                "Value",
                "Enter the Type or Category value to color:",
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
                    "A color is already configured for that value.",
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
                messagebox.showinfo("Edit Color", "Select a value to edit first.", parent=window)
                return
            current_label = working_labels.get(key, key)
            result = _prompt_value(current_label)
            if result is None:
                return
            new_key, new_label = result
            if new_key != key and new_key in working_map:
                messagebox.showerror(
                    "Duplicate Value",
                    "Another entry already uses that value.",
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
                messagebox.showinfo("Remove Color", "Select a value to remove first.", parent=window)
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
            self.type_category_color_map = {key: value for key, value in working_map.items() if value}
            self.type_category_color_labels = {
                key: working_labels.get(key, key) for key in self.type_category_color_map
            }
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
                        if pattern.search(page_text):
                            matches[column].append(Match(page_index=page_index, source="regex", pattern=pattern.pattern))
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

        self._rebuild_grid()
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

        copied_files = 0
        for pdf_path in recent_pdfs:
            try:
                destination = self._ensure_unique_path(raw_dir / pdf_path.name)
                shutil.copy2(pdf_path, destination)
                copied_files += 1
            except Exception as exc:
                messagebox.showwarning("Copy PDF", f"Could not copy '{pdf_path.name}': {exc}")

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
        if copied_files:
            messagebox.showinfo("Create Company", f"Copied {copied_files} PDF(s) into '{safe_name}/raw'.")

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
            if isinstance(value, dict) and "selections" in value:
                selections = value.get("selections")
                if isinstance(selections, dict):
                    parsed[pdf_name] = {
                        "selections": {k: int(v) for k, v in selections.items() if isinstance(v, (int, float))},
                        "year": value.get("year", ""),
                    }
            elif isinstance(value, dict):
                parsed[pdf_name] = {
                    "selections": {k: int(v) for k, v in value.items() if isinstance(v, (int, float))},
                    "year": value.get("year", ""),
                }
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
        company_dir = self.companies_dir / company
        path = company_dir / "type_category_item_assignments.csv"
        self.note_assignments_path = path
        self.note_assignments = self._read_note_assignments_file(path)
        self._update_combined_notes()

    def _ensure_note_assignments_path(self) -> Optional[Path]:
        if self.note_assignments_path is not None:
            return self.note_assignments_path
        company = self.company_var.get().strip()
        if not company:
            return None
        path = self.companies_dir / company / "type_category_item_assignments.csv"
        self.note_assignments_path = path
        return path

    def _write_note_assignments(self) -> None:
        path = self._ensure_note_assignments_path()
        if path is None:
            return
        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            with path.open("w", encoding="utf-8", newline="") as fh:
                writer = csv.writer(fh)
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
        return True

    def _prompt_import_note_assignments(self) -> None:
        company = self.company_var.get().strip()
        if not company:
            messagebox.showinfo("Assignments", "Select a company before loading assignments.")
            return
        initial_dir = self.companies_dir / company
        if not initial_dir.exists():
            initial_dir = self.companies_dir
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
            kwargs: Dict[str, Any] = {"background": color or ""}
            if color:
                kwargs["foreground"] = self._foreground_for_color(color)
            else:
                kwargs["foreground"] = ""
            tree.tag_configure(tag_name, **kwargs)

    def _apply_note_value_tag(self, item_id: str, value: str) -> None:
        tree = self.combined_result_tree
        if tree is None:
            return
        tag_name = self._note_tag_for_value(value)
        existing_tags = list(tree.item(item_id, "tags") or ())
        other_tags = [tag for tag in existing_tags if not tag.startswith("note_value_")]
        tree.item(item_id, tags=tuple(other_tags + [tag_name]))

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

    def _contrast_foreground_for_color(self, color: str) -> str:
        hex_color = color.lstrip("#")
        if len(hex_color) != 6:
            return "#000000"
        try:
            r = int(hex_color[0:2], 16)
            g = int(hex_color[2:4], 16)
            b = int(hex_color[4:6], 16)
        except ValueError:
            return "#000000"
        luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
        return "#000000" if luminance > 0.6 else "#ffffff"

    def _apply_type_category_color_tag(self, item_id: str, type_value: Any, category_value: Any) -> None:
        tree = self.combined_result_tree
        if tree is None or not item_id:
            return
        target_value = ""
        color_value = ""
        for candidate in (type_value, category_value):
            normalized = self._normalize_type_category_value(candidate)
            if normalized and normalized in self.type_category_color_map:
                candidate_color = self._normalize_hex_color(self.type_category_color_map[normalized])
                if candidate_color:
                    target_value = normalized
                    color_value = candidate_color
                    break
        existing_tags = list(tree.item(item_id, "tags") or ())
        filtered_tags = [tag for tag in existing_tags if not tag.startswith("type_category_color_")]
        if not target_value or not color_value:
            if len(filtered_tags) != len(existing_tags):
                tree.item(item_id, tags=tuple(filtered_tags))
            return
        tag_key = re.sub(r"[^0-9a-zA-Z]+", "_", target_value) or "value"
        tag_name = f"type_category_color_{tag_key}"
        foreground = self._contrast_foreground_for_color(color_value)
        tree.tag_configure(tag_name, background=color_value, foreground=foreground)
        tree.item(item_id, tags=tuple(filtered_tags + [tag_name]))

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
            self._apply_type_category_color_tag(item_id, type_value, category_value)

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
        if self.combined_show_blank_notes_var.get():
            try:
                self.root.after_idle(self._update_combined_tree_display)
            except tk.TclError:
                pass

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
        pdf_entries: List[Tuple[str, str, str]] = []
        row_type = tree.set(item_id, "Type")
        row_category = tree.set(item_id, "Category")
        row_item = tree.set(item_id, "Item")
        for column_name in self.combined_ordered_columns:
            if column_name in {"Type", "Category", "Item", "Note"}:
                continue
            value_text = tree.set(item_id, column_name) or ""
            if not str(value_text).strip():
                continue
            mapping = self.combined_column_name_map.get(column_name)
            if not mapping:
                continue
            pdf_path, label_value = mapping
            pdf_entries.append((pdf_path.stem, label_value, str(value_text)))

        if not pdf_entries:
            messagebox.showinfo(
                "Row Values",
                "No PDF columns contain a value for the selected row.",
            )
            return "break"

        dialog = tk.Toplevel(self.root)
        dialog.title("Row PDF Values")
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

        tree_frame = ttk.Frame(dialog, padding=(12, 0, 12, 0))
        tree_frame.pack(fill=tk.BOTH, expand=True)
        value_tree = ttk.Treeview(tree_frame, columns=("PDF", "Column", "Value"), show="headings", height=len(pdf_entries))
        value_tree.pack(fill=tk.BOTH, expand=True)
        for col_name, heading_text, anchor in (
            ("PDF", "PDF", "w"),
            ("Column", "Column", "w"),
            ("Value", "Value", "e"),
        ):
            value_tree.heading(col_name, text=heading_text)
            value_tree.column(col_name, anchor=anchor, stretch=True)
        for pdf_label, column_label, value_text in pdf_entries:
            value_tree.insert("", tk.END, values=(pdf_label, column_label, value_text))

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
        if self.combined_show_blank_notes_var.get():
            try:
                self.root.after_idle(self._update_combined_tree_display)
            except tk.TclError:
                pass

    def _apply_saved_selection(self, entry: PDFEntry) -> None:
        key = entry.path.name
        saved = self.assigned_pages.get(key)
        if not saved:
            return
        selections = saved.get("selections") if isinstance(saved, dict) else None
        if selections is None:
            selections = saved
        for category, page_index in selections.items():
            if category not in entry.matches:
                continue
            page_int = int(page_index)
            found = False
            for idx, match in enumerate(entry.matches[category]):
                if match.page_index == page_int:
                    entry.current_index[category] = idx
                    found = True
                    break
            if not found:
                entry.matches[category].append(Match(page_index=page_int, source="manual"))
                entry.current_index[category] = len(entry.matches[category]) - 1
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

    def select_match_index(self, entry: PDFEntry, category: str, index: int) -> None:
        matches = entry.matches.get(category, [])
        if not matches:
            return
        index = max(0, min(index, len(matches) - 1))
        entry.current_index[category] = index
        self._refresh_category_row(entry, category, rebuild=False)
        page_index = matches[index].page_index
        self._update_assigned_entry(entry, category, page_index, persist=False)

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
                        if pattern.search(page_text):
                            new_matches[column].append(
                                Match(page_index=page_index, source="regex", pattern=pattern.pattern)
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
                else:
                    entry.current_index[column] = None

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

    def _render_page(self, doc: fitz.Document, page_index: int, target_width: int) -> Optional[ImageTk.PhotoImage]:
        try:
            page = doc.load_page(page_index)
            zoom_matrix = fitz.Matrix(1.5, 1.5)
            pix = page.get_pixmap(matrix=zoom_matrix)
            mode = "RGBA" if pix.alpha else "RGB"
            image = Image.frombytes(mode, [pix.width, pix.height], pix.samples)
            if target_width > 0 and image.width != target_width:
                ratio = target_width / image.width
                new_size = (int(image.width * ratio), int(image.height * ratio))
                image = image.resize(new_size, Image.LANCZOS)
            return ImageTk.PhotoImage(image)
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
        self._refresh_category_row(entry, category, rebuild=False)
        page_index = self._get_selected_page_index(entry, category)
        if page_index is not None:
            self._update_assigned_entry(entry, category, page_index, persist=False)

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
        dialog.grab_set()
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
                self._refresh_category_row(entry, category, rebuild=False)
                self._update_assigned_entry(entry, category, page_index, persist=False)
                return

        matches.append(Match(page_index=page_index, source="manual"))
        entry.current_index[category] = len(matches) - 1
        self._refresh_category_row(entry, category, rebuild=True)
        self._update_assigned_entry(entry, category, page_index, persist=False)

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
        key = entry.path.name
        record = self.assigned_pages.setdefault(key, {"selections": {}, "year": entry.year})
        selections = record.setdefault("selections", {})
        selections[category] = int(page_index)
        if entry.year:
            record["year"] = entry.year
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
        if self.assigned_pages_path is None:
            self.assigned_pages_path = self.companies_dir / company / "assigned.json"

        if self.pdf_entries:
            for entry in self.pdf_entries:
                record = self.assigned_pages.setdefault(
                    entry.path.name, {"selections": {}, "year": entry.year}
                )
                record["year"] = entry.year
                selections = record.setdefault("selections", {})
                for category in COLUMNS:
                    page_index = self._get_selected_page_index(entry, category)
                    if page_index is None:
                        selections.pop(category, None)
                    else:
                        selections[category] = int(page_index)

        self._write_assigned_pages()
        self.scrape_selections()

    def _refresh_scraped_tab(self) -> None:
        if not hasattr(self, "scraped_inner"):
            return
        for child in self.scraped_inner.winfo_children():
            child.destroy()
        self.scraped_images.clear()
        self._clear_combined_tab()
        if hasattr(self, "scrape_progress"):
            self.scrape_progress["value"] = 0
        company = self.company_var.get()
        if not company:
            return
        scrape_root = self.companies_dir / company / "openapiscrape"
        if not scrape_root.exists():
            return

        self.scraped_inner.columnconfigure(0, weight=1)
        self.scraped_inner.columnconfigure(1, weight=2)
        row_index = 0
        header_added = False
        for entry in self.pdf_entries:
            doc_dir = scrape_root / entry.stem
            metadata = self._load_doc_metadata(doc_dir)
            if not metadata:
                continue
            for category in COLUMNS:
                meta = metadata.get(category)
                if not isinstance(meta, dict):
                    continue
                page_index = meta.get("page_index")
                if page_index is None:
                    continue
                preview_mode: Optional[str] = None
                preview_path: Optional[Path] = None
                preview_text: Optional[str] = None
                rows: List[List[str]] = []

                txt_name = meta.get("txt")
                if isinstance(txt_name, str) and txt_name:
                    candidate = doc_dir / txt_name
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
                        candidate_csv = doc_dir / csv_name
                        if candidate_csv.exists():
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
                ttk.Label(
                    header_frame,
                    text=f"{entry.path.name} - {category} (Page {int(page_index) + 1})",
                    font=("TkDefaultFont", 10, "bold"),
                ).grid(row=0, column=0, sticky="w")
                ttk.Button(
                    header_frame,
                    text="View Prompt",
                    command=lambda e=entry, c=category, p=int(page_index): self._show_prompt_preview(
                        e, c, p
                    ),
                    width=12,
                ).grid(row=0, column=1, sticky="e", padx=(0, 4))
                ttk.Button(
                    header_frame,
                    text="View Raw Text",
                    command=lambda e=entry, p=int(page_index): self._show_raw_text_dialog(e, p),
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
                ttk.Button(
                    header_frame,
                    text="Delete",
                    command=lambda d=doc_dir, c=category: self._delete_scrape_output(d, c),
                    width=10,
                ).grid(row=0, column=4, sticky="e")
                row_index += 1

                image_frame = ttk.Frame(self.scraped_inner, padding=8)
                image_frame.grid(row=row_index, column=0, sticky="nsew", padx=4)
                table_frame = ttk.Frame(self.scraped_inner, padding=8)
                table_frame.grid(row=row_index, column=1, sticky="nsew", padx=4)
                self.scraped_inner.rowconfigure(row_index, weight=1)

                photo = self._render_page(entry.doc, int(page_index), 350)
                if photo is not None:
                    self.scraped_images.append(photo)
                    image_label = ttk.Label(image_frame, image=photo, cursor="hand2")
                    image_label.pack(expand=True, fill=tk.BOTH)
                    image_label.bind(
                        "<Button-1>",
                        lambda _e, e=entry, p=int(page_index): self.open_thumbnail_zoom(e, p),
                    )
                else:
                    ttk.Label(image_frame, text="Preview unavailable").pack(expand=True, fill=tk.BOTH)

                table_frame.columnconfigure(0, weight=1)
                table_frame.rowconfigure(0, weight=1)

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
                    tree = ttk.Treeview(
                        table_frame,
                        columns=columns,
                        show="headings",
                        height=min(15, max(3, len(data_rows))),
                    )
                    y_scroll = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=tree.yview)
                    x_scroll = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=tree.xview)
                    tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
                    tree.grid(row=0, column=0, sticky="nsew")
                    y_scroll.grid(row=0, column=1, sticky="ns")
                    x_scroll.grid(row=1, column=0, sticky="ew")
                    for col, heading in zip(columns, headings):
                        tree.heading(col, text=heading)
                        tree.column(col, anchor="center", stretch=True, width=120)
                    if data_rows:
                        for data in data_rows:
                            values = list(data)
                            if len(values) < len(columns):
                                values.extend([""] * (len(columns) - len(values)))
                            display_values = [self._format_table_value(value) for value in values]
                            tree.insert("", tk.END, values=display_values)

                row_index += 1

        self._refresh_combined_tab(auto_update=True)

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
        self.combined_result_message = None

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
        pdf_entries_with_data: List[PDFEntry] = []
        max_data_columns = 0

        for entry in self.pdf_entries:
            doc_dir = scrape_root / entry.stem
            metadata = self._load_doc_metadata(doc_dir)
            if not metadata:
                continue
            entry_has_data = False
            for category in COLUMNS:
                meta = metadata.get(category)
                if not isinstance(meta, dict):
                    continue
                rows: List[List[str]] = []
                txt_name = meta.get("txt")
                if isinstance(txt_name, str) and txt_name:
                    txt_path = doc_dir / txt_name
                    if txt_path.exists():
                        try:
                            text = txt_path.read_text(encoding="utf-8")
                        except Exception:
                            text = ""
                        rows = self._convert_response_to_rows(text)
                if not rows:
                    csv_name = meta.get("csv")
                    if isinstance(csv_name, str) and csv_name:
                        csv_path = doc_dir / csv_name
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
                if len(display_headers) > MAX_COMBINED_DATE_COLUMNS:
                    display_headers = display_headers[:MAX_COMBINED_DATE_COLUMNS]
                    data_indices = data_indices[:MAX_COMBINED_DATE_COLUMNS]
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
                pdf_entries_with_data.append(entry)

        if not pdf_entries_with_data:
            ttk.Label(
                self.combined_header_frame,
                text="No scraped results available to combine.",
            ).grid(row=0, column=0, sticky="w", padx=8, pady=8)
            return

        self.combined_pdf_order = [entry.path for entry in pdf_entries_with_data]
        self.combined_max_data_columns = min(max_data_columns, MAX_COMBINED_DATE_COLUMNS)
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
                    self.combined_column_label_vars[key] = tk.StringVar(master=self.root, value=default_text)
                elif not var.get().strip():
                    var.set(default_text)
        for stale_key in list(self.combined_column_label_vars.keys()):
            if stale_key not in desired_keys:
                self.combined_column_label_vars.pop(stale_key, None)

        base_columns = ["Type", "Category", "Item"]
        pdf_count = len(self.combined_pdf_order)
        total_columns = len(base_columns) + pdf_count * max_data_columns
        logger.info(
            "Combined header grid layout -> base_columns=%d pdf_count=%d total_columns=%d",
            len(base_columns),
            pdf_count,
            total_columns,
        )
        for column_index in range(total_columns):
            weight = 1 if column_index >= len(base_columns) else 0
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

        current_row = 0

        for column_index, label_text in enumerate(["Type", default_category_heading, default_item_heading]):
            ttk.Label(
                self.combined_header_frame,
                text=label_text,
                font=("TkDefaultFont", 10, "bold"),
            ).grid(row=current_row, column=column_index, padx=4, pady=(0, 4), sticky="w")

        column_position = len(base_columns)
        if max_data_columns:
            for pdf_path in self.combined_pdf_order:
                pdf_label = pdf_path.stem
                for idx in range(max_data_columns):
                    key = (pdf_path, idx)
                    var = self.combined_column_label_vars.get(key)
                    if var is None:
                        continue
                    cell = ttk.Frame(self.combined_header_frame)
                    cell.grid(
                        row=current_row,
                        column=column_position,
                        padx=4,
                        pady=(0, 4),
                        sticky="w",
                    )
                    ttk.Label(
                        cell,
                        text=pdf_label,
                        font=("TkDefaultFont", 10, "bold"),
                    ).pack(anchor="w")
                    entry = ttk.Entry(cell, textvariable=var, width=18)
                    entry.pack(anchor="w", fill=tk.X, pady=(2, 0))
                    column_position += 1

        current_row += 1
        for category in COLUMNS:
            header_by_pdf = type_header_map.get(category, {})
            if not header_by_pdf and category not in type_heading_labels:
                continue
            category_heading_text = default_category_heading
            item_heading_text = default_item_heading
            for entry_meta in header_by_pdf.values():
                category_heading_text = entry_meta.get("category_heading", category_heading_text)
                item_heading_text = entry_meta.get("item_heading", item_heading_text)
                break
            ttk.Label(self.combined_header_frame, text=category).grid(
                row=current_row,
                column=0,
                padx=4,
                pady=4,
                sticky="w",
            )
            ttk.Label(self.combined_header_frame, text=category_heading_text).grid(
                row=current_row,
                column=1,
                padx=4,
                pady=4,
                sticky="w",
            )
            ttk.Label(self.combined_header_frame, text=item_heading_text).grid(
                row=current_row,
                column=2,
                padx=4,
                pady=4,
                sticky="w",
            )
            column_position = len(base_columns)
            for pdf_path in self.combined_pdf_order:
                headers = header_by_pdf.get(pdf_path, {}).get("headers", [])
                for idx in range(max_data_columns):
                    header_text = headers[idx] if idx < len(headers) else ""
                    ttk.Label(self.combined_header_frame, text=header_text).grid(
                        row=current_row,
                        column=column_position,
                        padx=4,
                        pady=4,
                        sticky="w",
                    )
                    column_position += 1
            current_row += 1

        if hasattr(self, "combine_confirm_button"):
            self.combine_confirm_button.state(["!disabled"])

        if auto_update and self.combined_pdf_order and self.combined_csv_sources and self.combined_result_tree is None:
            self._confirm_combined_table(auto=True)

    def _filter_combined_records(self, records: List[Dict[str, str]]) -> List[Dict[str, str]]:
        if not records:
            return []
        if not self.combined_show_blank_notes_var.get():
            return records
        filtered: List[Dict[str, str]] = []
        for record in records:
            note_value = record.get("Note", "")
            if isinstance(note_value, str):
                normalized = note_value.strip()
            else:
                normalized = str(note_value or "").strip()
            if not normalized:
                filtered.append(record)
        return filtered

    def _update_combined_tree_display(self) -> None:
        if not self.combined_ordered_columns:
            return
        records = self._filter_combined_records(self.combined_all_records)
        self._render_combined_tree(records)

    def _on_combined_show_blank_notes_toggle(self) -> None:
        if not self.combined_ordered_columns:
            return
        self._update_combined_tree_display()

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
        result_container = ttk.Frame(self.combined_result_frame)
        result_container.pack(fill=tk.BOTH, expand=True)
        result_container.columnconfigure(0, weight=1)
        result_container.rowconfigure(0, weight=1)
        tree = ttk.Treeview(result_container, columns=columns, show="headings")
        tree.grid(row=0, column=0, sticky="nsew")
        tree.configure(height=max(len(records), 1))

        def _combined_xscroll(*args: Any) -> None:
            tree.xview(*args)
            self._destroy_note_editor()

        x_scroll = ttk.Scrollbar(result_container, orient=tk.HORIZONTAL, command=_combined_xscroll)
        x_scroll.grid(row=1, column=0, sticky="ew")
        tree.configure(xscrollcommand=x_scroll.set)

        for column_name in columns:
            tree.heading(column_name, text=column_name)
            if column_name in {"Type", "Category", "Item"}:
                tree.column(column_name, anchor="w", stretch=False)
            elif column_name == "Note":
                tree.column(column_name, anchor="center", stretch=False, minwidth=120)
            else:
                tree.column(column_name, anchor="center", stretch=True)

        note_index = None
        if "Note" in columns:
            note_index = columns.index("Note")
            self.combined_note_column_id = f"#{note_index + 1}"
        else:
            self.combined_note_column_id = None
        self.combined_note_record_keys.clear()
        self._configure_note_tags(tree)
        self.combined_result_tree = tree

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
            self._apply_type_category_color_tag(
                item_id,
                record.get("Type", ""),
                record.get("Category", ""),
            )
            note_value = record.get("Note", "")
            if not isinstance(note_value, str):
                note_value = str(note_value or "")
            self._apply_note_value_tag(item_id, note_value)

        tree.bind("<Button-1>", self._on_combined_tree_click)
        tree.bind("<ButtonRelease-1>", self._on_combined_tree_release, add="+")
        tree.bind("<Configure>", lambda _e: self._destroy_note_editor())
        tree.bind("<<TreeviewSelect>>", lambda _e: self._destroy_note_editor())
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
                    self.combined_canvas.yview_scroll(delta, "units")
            elif getattr(event, "num", None) == 4:
                self.combined_canvas.yview_scroll(-1, "units")
            elif getattr(event, "num", None) == 5:
                self.combined_canvas.yview_scroll(1, "units")
            return "break"

        tree.bind("<MouseWheel>", _tree_mousewheel)
        tree.bind("<Button-4>", _tree_mousewheel)
        tree.bind("<Button-5>", _tree_mousewheel)

        self._refresh_combined_column_widths()
        self._refresh_type_category_colors()
        self._render_combined_result_message()

    def _render_combined_result_message(self) -> None:
        if not self.combined_result_message:
            return
        ttk.Label(
            self.combined_result_frame,
            text=self.combined_result_message,
        ).pack(anchor="w", padx=8, pady=(4, 0))

    def _update_combined_record_note_value(self, key: Tuple[str, str, str], value: str) -> None:
        record = self.combined_record_lookup.get(key)
        if record is None:
            return
        record["Note"] = value

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
                for path in self.combined_pdf_order:
                    pdf_label = path.stem
                    labels = column_labels_by_pdf.get(path, [])
                    for position, base_label in enumerate(labels):
                        column_name = f"{pdf_label}.{base_label}"
                        raw_value = value_map.get((path, position), "")
                        record[column_name] = self._format_combined_value(raw_value)
                combined_records.append(record)

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

        self.combined_result_message = None
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

        company = self.company_var.get()
        if company:
            scrape_root = self.companies_dir / company / "openapiscrape"
            combined_path = scrape_root / "combined.csv"
            try:
                csv_header: List[str] = []
                for column_name in ordered_columns:
                    if column_name in {"Type", "Category", "Item", "Note"}:
                        csv_header.append(column_name)
                    elif "." in column_name:
                        csv_header.append(column_name.split(".", 1)[1])
                    else:
                        csv_header.append(column_name)
                csv_rows = [csv_header]
                for record in combined_records:
                    csv_rows.append([record.get(column_name, "") for column_name in ordered_columns])
                self._write_csv_rows(csv_rows, combined_path)
                self.combined_result_message = f"Combined CSV saved to: {combined_path}"
                self._render_combined_result_message()
            except Exception as exc:
                messagebox.showwarning("Combine Results", f"Could not save combined CSV: {exc}")

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
            tree.column(column_name, width=desired_width, minwidth=max(20, desired_width))
        other_width = self.combined_other_column_width
        if isinstance(other_width, int) and other_width > 0:
            for column_name in self.combined_ordered_columns:
                if column_name in {"Type", "Category", "Item", "Note"}:
                    continue
                tree.column(
                    column_name,
                    width=other_width,
                    minwidth=max(20, other_width),
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
        if self.type_category_color_map:
            stored: Dict[str, str] = {}
            for normalized, color in self.type_category_color_map.items():
                if not color:
                    continue
                label = self.type_category_color_labels.get(normalized, normalized)
                stored[label] = color
            if stored:
                self.config_data["combined_type_category_colors"] = stored
            else:
                self.config_data.pop("combined_type_category_colors", None)
        else:
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
            tree.column(column_name, width=desired_width, minwidth=max(20, desired_width))

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

    def _load_doc_metadata(self, doc_dir: Path) -> Dict[str, Any]:
        metadata_path = doc_dir / "metadata.json"
        if not metadata_path.exists():
            return {}
        try:
            with metadata_path.open("r", encoding="utf-8") as fh:
                data = json.load(fh)
        except Exception:
            return {}
        return data if isinstance(data, dict) else {}

    def _write_doc_metadata(self, doc_dir: Path, data: Dict[str, Any]) -> None:
        metadata_path = doc_dir / "metadata.json"
        try:
            with metadata_path.open("w", encoding="utf-8") as fh:
                json.dump(data, fh, indent=2)
        except Exception as exc:
            messagebox.showwarning("Save Metadata", f"Could not save scrape metadata: {exc}")

    def _delete_scrape_output(self, doc_dir: Path, category: str) -> None:
        metadata = self._load_doc_metadata(doc_dir)
        if not metadata:
            return
        meta = metadata.get(category)
        errors: List[str] = []
        if isinstance(meta, dict):
            for key in ("csv", "txt"):
                name = meta.get(key)
                if not isinstance(name, str) or not name:
                    continue
                path = doc_dir / name
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
        self._write_doc_metadata(doc_dir, metadata)
        self._refresh_scraped_tab()
        if errors:
            messagebox.showwarning("Delete Output", "\n".join(errors))

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

        delimiter = "\t" if "\t" in text else ","
        try:
            reader = csv.reader(io.StringIO(text), delimiter=delimiter)
            return [list(row) for row in reader]
        except Exception:
            return []

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
        url = "https://api.openai.com/v1/chat/completions"
        headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
        payload = {
            "model": "gpt-4o-mini",
            "messages": [
                {"role": "system", "content": prompt},
                {"role": "user", "content": page_text},
            ],
        }
        response = requests.post(url, headers=headers, json=payload, timeout=120)
        response.raise_for_status()
        data = response.json()
        choices = data.get("choices")
        if not choices:
            raise ValueError("No choices returned from OpenAI API")
        message = choices[0].get("message", {})
        content = message.get("content")
        if not content:
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
            return [["response"], [""]]

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
                return rows

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
                return rows

        if any(
            "," in line and not line.startswith("#") for line in lines
        ):
            rows = [[segment.strip() for segment in line.split(",")] for line in lines]
            return rows

        return [["response"], *[[line] for line in lines]]

    def _write_csv_rows(self, rows: List[List[str]], csv_path: Path) -> None:
        csv_path.parent.mkdir(parents=True, exist_ok=True)
        with csv_path.open("w", encoding="utf-8", newline="") as fh:
            writer = csv.writer(fh)
            for row in rows:
                writer.writerow(row)

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
                data = json.load(fh)
                self.config_data = data
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
                self._load_type_category_colors_from_config(data.get("combined_type_category_colors"))
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
        if "api_key" in data:
            self.api_key_var.set(str(data.get("api_key", "")))
        self._ensure_download_settings()
        self._apply_last_company_selection()
        self._config_loaded = True

    def _load_type_category_colors_from_config(self, raw: Any) -> None:
        self.type_category_color_map = {}
        self.type_category_color_labels = {}
        if isinstance(raw, dict):
            items = raw.items()
        elif isinstance(raw, list):
            items = []
            for entry in raw:
                if isinstance(entry, dict):
                    value = entry.get("value")
                    color = entry.get("color")
                    items.append((value, color))
        else:
            return
        for key, color in items:  # type: ignore[misc]
            normalized_key = self._normalize_type_category_value(key)
            normalized_color = self._normalize_hex_color(str(color)) if color is not None else None
            if not normalized_key or not normalized_color:
                continue
            self.type_category_color_map[normalized_key] = normalized_color
            self.type_category_color_labels[normalized_key] = str(key).strip()

    def _apply_configured_note_key_bindings(self) -> None:
        note_order: List[str] = []
        seen: Set[str] = set()

        def _add_option(value: Any) -> None:
            if isinstance(value, str):
                normalized_value = value.strip().lower()
            else:
                normalized_value = str(value).strip().lower()
            if not normalized_value:
                normalized_value = ""
            if normalized_value in seen:
                return
            if normalized_value:
                note_order.append(normalized_value)
            else:
                note_order.insert(0, "")
            seen.add(normalized_value)

        configured_colors: Dict[str, str] = {}
        configured_bindings: Dict[str, str] = {}

        _add_option("")

        raw_settings = self.config_data.get("combined_note_settings")
        if isinstance(raw_settings, list):
            for entry in raw_settings:
                if not isinstance(entry, dict):
                    continue
                value = entry.get("value", "")
                _add_option(value)
                normalized_value = value.strip().lower() if isinstance(value, str) else str(value).strip().lower()
                if not normalized_value:
                    normalized_value = ""
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
                _add_option(value)
                normalized_value = value.strip().lower() if isinstance(value, str) else str(value).strip().lower()
                if not normalized_value:
                    normalized_value = ""
                normalized_shortcut = (
                    self._normalize_note_binding_value(str(shortcut)) if shortcut else ""
                )
                if normalized_value not in configured_bindings or normalized_shortcut:
                    configured_bindings[normalized_value] = normalized_shortcut

        for default_value in DEFAULT_NOTE_OPTIONS:
            _add_option(default_value)

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

        self._update_note_settings_config_entries()
        self._refresh_note_tags()

    def _apply_whitespace_option(self, pattern: str, enabled: bool) -> str:
        if not enabled:
            return pattern
        return pattern.replace(" ", r"\s+")

    def _ensure_download_settings(self) -> None:
        configured_dir = str(self.config_data.get("downloads_dir", "")).strip()
        if configured_dir:
            self.downloads_dir_var.set(configured_dir)
        elif not self.downloads_dir_var.get():
            default_download_dir = Path.home() / "Downloads"
            if default_download_dir.exists():
                self.downloads_dir_var.set(str(default_download_dir))
        if self.downloads_dir_var.get():
            self.config_data["downloads_dir"] = self.downloads_dir_var.get()

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

        self.note_options = sanitized_order
        self.note_background_colors = sanitized_colors
        self.note_key_bindings = sanitized_bindings

        settings_payload: List[Dict[str, str]] = []
        for value in sanitized_order:
            entry: Dict[str, str] = {"value": value}
            shortcut = sanitized_bindings.get(value, "")
            if shortcut:
                entry["shortcut"] = shortcut
            color = sanitized_colors.get(value, "")
            if color:
                entry["color"] = color
            settings_payload.append(entry)

        self.config_data["combined_note_settings"] = settings_payload
        self.config_data["combined_note_key_bindings"] = {
            value: sanitized_bindings.get(value, "") for value in sanitized_order
        }

    def _write_config(self) -> None:
        self._update_note_settings_config_entries()
        self.config_data["combined_base_column_widths"] = {
            column: int(width)
            for column, width in self.combined_base_column_widths.items()
            if column in {"Type", "Category", "Item", "Note"} and width > 0
        }
        if isinstance(self.combined_other_column_width, int) and self.combined_other_column_width > 0:
            self.config_data["combined_other_column_width"] = int(self.combined_other_column_width)
        else:
            self.config_data.pop("combined_other_column_width", None)
        if self.type_category_color_map:
            stored: Dict[str, str] = {}
            for normalized, color in self.type_category_color_map.items():
                if not color:
                    continue
                label = self.type_category_color_labels.get(normalized, normalized)
                stored[label] = color
            if stored:
                self.config_data["combined_type_category_colors"] = stored
            else:
                self.config_data.pop("combined_type_category_colors", None)
        else:
            self.config_data.pop("combined_type_category_colors", None)
        try:
            with self.pattern_config_path.open("w", encoding="utf-8") as fh:
                json.dump(self.config_data, fh, indent=2)
        except Exception as exc:  # pragma: no cover - guard for IO issues
            messagebox.showwarning("Save Patterns", f"Could not save pattern configuration: {exc}")

    def scrape_selections(self) -> None:
        if not self.pdf_entries:
            messagebox.showinfo("No PDFs", "Load PDFs before running AIScrape.")
            return

        company = self.company_var.get()
        if not company:
            messagebox.showinfo("Select Company", "Please choose a company before running AIScrape.")
            return

        if hasattr(self, "scrape_progress"):
            self.scrape_progress["value"] = 0

        api_key = self.api_key_var.get().strip()
        if not api_key:
            messagebox.showwarning("API Key Required", "Enter an API key and press Enter before running AIScrape.")
            self.api_key_entry.focus_set()
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
            return

        self._save_api_key()

        scrape_root = self.companies_dir / company / "openapiscrape"
        scrape_root.mkdir(parents=True, exist_ok=True)

        errors: List[str] = []
        pending: List[Tuple[PDFEntry, List[Tuple[str, int]]]] = []
        metadata_cache: Dict[Path, Dict[str, Any]] = {}
        doc_dir_map: Dict[Path, Path] = {}

        for entry in self.pdf_entries:
            entry_tasks: List[Tuple[str, int]] = []
            doc_dir = scrape_root / entry.stem
            doc_dir.mkdir(parents=True, exist_ok=True)
            metadata = self._load_doc_metadata(doc_dir)
            if not isinstance(metadata, dict):
                metadata = {}
            else:
                metadata = dict(metadata)
            metadata_cache[entry.path] = metadata
            doc_dir_map[entry.path] = doc_dir
            for category in COLUMNS:
                page_index = self._get_selected_page_index(entry, category)
                if page_index is None:
                    continue
                if not prompts.get(category):
                    continue
                existing = metadata.get(category)
                file_exists = False
                if isinstance(existing, dict):
                    for key in ("csv", "txt"):
                        name = existing.get(key)
                        if isinstance(name, str) and name:
                            if (doc_dir / name).exists():
                                file_exists = True
                                break
                if file_exists:
                    continue
                entry_tasks.append((category, page_index))

            if entry_tasks:
                pending.append((entry, entry_tasks))

        total_tasks = sum(len(items) for _, items in pending)
        if not total_tasks:
            messagebox.showinfo("AIScrape", "All selected sections already have scraped files.")
            return

        if hasattr(self, "scrape_button"):
            self.scrape_button.state(["disabled"])
        if hasattr(self, "scrape_progress"):
            self.scrape_progress.configure(maximum=total_tasks, value=0)

        successful = 0
        attempted = 0

        metadata_changed: Dict[Path, bool] = {}
        jobs: List[ScrapeTask] = []

        def _run_job(job: ScrapeTask) -> Tuple[bool, Optional[Dict[str, Any]], Optional[str]]:
            try:
                response_text = self._call_openai(api_key, job.prompt_text, job.page_text)
            except requests.RequestException as exc:
                return False, None, f"{job.entry_name} - {job.category}: API request failed ({exc})"
            except Exception as exc:
                return False, None, f"{job.entry_name} - {job.category}: {exc}"

            base_name = f"{job.entry_year or 'unknown'}_{job.category}"
            csv_path = self._ensure_unique_path(job.doc_dir / f"{base_name}.csv")

            rows = self._convert_response_to_rows(response_text)
            if not rows:
                rows = [["response"], [response_text]]
            sanitized_rows = [
                ["" if cell is None else str(cell) for cell in row]
                for row in rows
            ]

            try:
                self._write_csv_rows(sanitized_rows, csv_path)
            except Exception as exc:
                return False, None, f"{job.entry_name} - {job.category}: Could not save response ({exc})"

            metadata_entry = {
                "csv": csv_path.name,
                "page_index": job.page_index,
                "year": job.entry_year,
            }
            return True, metadata_entry, None

        try:
            for entry, tasks in pending:
                metadata = metadata_cache.get(entry.path, {})
                if not isinstance(metadata, dict):
                    metadata = {}
                metadata = dict(metadata)
                metadata_cache[entry.path] = metadata
                metadata_changed.setdefault(entry.path, False)
                for category, page_index in tasks:
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
                    try:
                        page = entry.doc.load_page(page_index)
                        page_text = page.get_text("text")
                    except Exception as exc:
                        errors.append(f"{entry.path.name} - {category}: Could not read page ({exc})")
                        attempted += 1
                        if hasattr(self, "scrape_progress"):
                            self.scrape_progress["value"] = attempted
                            self.root.update()
                        continue

                    jobs.append(
                        ScrapeTask(
                            entry_path=entry.path,
                            entry_name=entry.path.name,
                            entry_year=entry.year,
                            category=category,
                            page_index=page_index,
                            prompt_text=prompt_text,
                            page_text=page_text,
                            doc_dir=doc_dir_map.get(entry.path, scrape_root / entry.stem),
                        )
                    )

            if jobs:
                max_workers = min(8, max(2, os.cpu_count() or 4))
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

                        attempted += 1

                        if success and metadata_entry is not None:
                            metadata = metadata_cache.get(job.entry_path, {})
                            if not isinstance(metadata, dict):
                                metadata = {}
                                metadata_cache[job.entry_path] = metadata
                            metadata[job.category] = metadata_entry
                            metadata_changed[job.entry_path] = True
                            successful += 1
                        elif error_message:
                            errors.append(error_message)

                        if hasattr(self, "scrape_progress"):
                            self.scrape_progress["value"] = attempted
                            self.root.update()

            for entry_path, changed in metadata_changed.items():
                if not changed:
                    continue
                metadata = metadata_cache.get(entry_path, {})
                if not isinstance(metadata, dict):
                    continue
                doc_dir = doc_dir_map.get(entry_path)
                if doc_dir is None:
                    continue
                self._write_doc_metadata(doc_dir, metadata)
        finally:
            if hasattr(self, "scrape_button"):
                self.scrape_button.state(["!disabled"])
            if hasattr(self, "scrape_progress"):
                self.scrape_progress["value"] = 0

        self._refresh_scraped_tab()
        if successful:
            if hasattr(self, "notebook") and hasattr(self, "scraped_frame"):
                self.notebook.select(self.scraped_frame)
            messagebox.showinfo("AIScrape Complete", f"Saved {successful} OpenAI responses to 'openapiscrape'.")
        if errors:
            messagebox.showerror("AIScrape Issues", "\n".join(errors))

