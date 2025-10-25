import csv
import json
import os
import re
import subprocess
import sys
import tkinter as tk
from concurrent.futures import ThreadPoolExecutor, as_completed
from tkinter import messagebox, simpledialog, ttk
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import fitz  # PyMuPDF
from PIL import Image, ImageTk
import webbrowser
import requests


COLUMNS = ["Financial", "Income", "Shares"]
DEFAULT_PATTERNS = {
    "Financial": ["statement of financial position"],
    "Income": ["statement of profit or loss"],
    "Shares": ["Movements in issued capital"],
}
YEAR_DEFAULT_PATTERNS = [r"(\d{4})\s+Annual\s+Report"]


@dataclass
class Match:
    page_index: int
    source: str
    pattern: Optional[str] = None


@dataclass
class PDFEntry:
    path: Path
    doc: fitz.Document
    matches: Dict[str, List[Match]] = field(default_factory=dict)
    current_index: Dict[str, Optional[int]] = field(default_factory=dict)
    year: str = ""

    def __post_init__(self) -> None:
        for column in COLUMNS:
            self.matches.setdefault(column, [])
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
    page_index: int
    prompt_text: str
    page_text: str
    doc_dir: Path


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
        current_index = self.entry.current_index.get(self.category)
        for thumb in self.thumbnails:
            thumb.set_selected(current_index == thumb.match_index)

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

        self._build_ui()
        self._load_pattern_config()
        if not self.embedded:
            self._maximize_window()
            self.root.after(0, self._load_pdfs_on_start)

    def _build_ui(self) -> None:
        if not self.embedded:
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
            new_company_button = ttk.Button(top_frame, text="New Company", command=self.create_company)
            new_company_button.pack(side=tk.LEFT, padx=(0, 4))
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
        pattern_section = CollapsibleFrame(review_container, "Regex patterns (one per line)")
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
        self.scraped_canvas = tk.Canvas(scraped_container)
        self.scraped_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scraped_scrollbar = ttk.Scrollbar(scraped_container, orient=tk.VERTICAL, command=self.scraped_canvas.yview)
        scraped_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.scraped_canvas.configure(yscrollcommand=scraped_scrollbar.set)
        self.scraped_inner = ttk.Frame(self.scraped_canvas)
        self.scraped_window = self.scraped_canvas.create_window((0, 0), window=self.scraped_inner, anchor="nw")
        self.scraped_inner.bind("<Configure>", lambda _e: self.scraped_canvas.configure(scrollregion=self.scraped_canvas.bbox("all")))
        self.scraped_canvas.bind("<Configure>", lambda e: self.scraped_canvas.itemconfigure(self.scraped_window, width=e.width))

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
            if self.embedded:
                self._reset_review_scroll()
            return
        folder = self.companies_dir / company / "raw"
        self.folder_path.set(str(folder))
        if self.embedded:
            self._load_assigned_pages(company)
            self._refresh_scraped_tab()
            self._reset_review_scroll()
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

    def load_pdfs(self) -> None:
        folder = self.folder_path.get()
        if not folder:
            messagebox.showinfo("Select Folder", "Please select a folder containing PDFs.")
            return

        folder_path = Path(folder)
        if not folder_path.exists():
            messagebox.showerror("Folder Not Found", f"The folder '{folder}' does not exist.")
            return

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
        name = simpledialog.askstring("New Company", "Enter a name for the new company:", parent=self.root)
        if name is None:
            return

        normalized = name.strip()
        if not normalized:
            messagebox.showwarning("Invalid Name", "Company name cannot be empty.")
            return

        safe_name = re.sub(r"[\\/:*?\"<>|]", "_", normalized).strip().strip(".")
        if not safe_name:
            messagebox.showwarning(
                "Invalid Name",
                "Company name contains only unsupported characters. Please choose a different name.",
            )
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
            return

        self._refresh_company_options()
        if hasattr(self, "company_combo"):
            self.company_combo.set(safe_name)
        self.company_var.set(safe_name)
        self._set_folder_from_company(safe_name)
        if self._config_loaded:
            self._update_last_company(safe_name)

        self._open_in_file_manager(raw_dir)
        self._open_company_tab()

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
        messagebox.showinfo("Assignments Saved", "Assignments have been committed to assigned.json.")

    def _refresh_scraped_tab(self) -> None:
        if not hasattr(self, "scraped_inner"):
            return
        for child in self.scraped_inner.winfo_children():
            child.destroy()
        self.scraped_images.clear()
        if hasattr(self, "scrape_progress"):
            self.scrape_progress["value"] = 0
        company = self.company_var.get()
        if not company:
            return
        scrape_root = self.companies_dir / company / "openapiscrape"
        if not scrape_root.exists():
            return

        self.scraped_inner.columnconfigure(0, weight=1)
        self.scraped_inner.columnconfigure(1, weight=1)
        self.scraped_inner.columnconfigure(2, weight=1)
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
                csv_name = meta.get("csv")
                page_index = meta.get("page_index")
                if csv_name is None or page_index is None:
                    continue
                csv_path = doc_dir / csv_name
                if not csv_path.exists():
                    continue
                if not header_added:
                    header_frame = ttk.Frame(self.scraped_inner)
                    header_frame.grid(row=row_index, column=0, columnspan=3, sticky="ew", padx=8, pady=(8, 0))
                    ttk.Label(
                        header_frame,
                        text="Original PDF Page",
                        font=("TkDefaultFont", 10, "bold"),
                    ).grid(row=0, column=0, sticky="w")
                    ttk.Label(
                        header_frame,
                        text="Raw Text",
                        font=("TkDefaultFont", 10, "bold"),
                    ).grid(row=0, column=1, sticky="w")
                    ttk.Label(
                        header_frame,
                        text="Converted CSV",
                        font=("TkDefaultFont", 10, "bold"),
                    ).grid(row=0, column=2, sticky="w")
                    header_frame.columnconfigure(0, weight=1)
                    header_frame.columnconfigure(1, weight=1)
                    header_frame.columnconfigure(2, weight=1)
                    row_index += 1
                    header_added = True
                header_frame = ttk.Frame(self.scraped_inner)
                header_frame.grid(row=row_index, column=0, columnspan=3, sticky="ew", padx=8, pady=(8, 0))
                header_frame.columnconfigure(0, weight=1)
                ttk.Label(
                    header_frame,
                    text=f"{entry.path.name} - {category} (Page {int(page_index) + 1})",
                    font=("TkDefaultFont", 10, "bold"),
                ).grid(row=0, column=0, sticky="w")
                ttk.Button(
                    header_frame,
                    text="View Prompt",
                    command=lambda c=category: self._show_prompt_preview(c),
                    width=12,
                ).grid(row=0, column=1, sticky="e", padx=(0, 4))
                ttk.Button(
                    header_frame,
                    text="Delete",
                    command=lambda d=doc_dir, c=category: self._delete_scrape_output(d, c),
                    width=10,
                ).grid(row=0, column=2, sticky="e")
                row_index += 1

                image_frame = ttk.Frame(self.scraped_inner, padding=8)
                image_frame.grid(row=row_index, column=0, sticky="nsew", padx=4)
                text_frame = ttk.Frame(self.scraped_inner, padding=8)
                text_frame.grid(row=row_index, column=1, sticky="nsew", padx=4)
                table_frame = ttk.Frame(self.scraped_inner, padding=8)
                table_frame.grid(row=row_index, column=2, sticky="nsew", padx=4)
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

                raw_text = ""
                try:
                    page = entry.doc.load_page(int(page_index))
                    raw_text = page.get_text("text")
                except Exception:
                    raw_text = "Could not load page text."

                text_container = tk.Text(text_frame, wrap="word", height=25)
                text_container.insert("1.0", raw_text)
                text_container.configure(state="disabled")
                text_scroll = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=text_container.yview)
                text_container.configure(yscrollcommand=text_scroll.set)
                text_container.pack(side=tk.LEFT, expand=True, fill=tk.BOTH)
                text_scroll.pack(side=tk.RIGHT, fill=tk.Y)

                rows = self._read_csv_rows(csv_path)
                if not rows:
                    ttk.Label(table_frame, text="CSV file is empty").pack(expand=True, fill=tk.BOTH)
                else:
                    headings = rows[0]
                    data_rows = rows[1:] if len(rows) > 1 else []
                    columns = [f"col_{idx}" for idx in range(len(headings))]
                    tree = ttk.Treeview(
                        table_frame,
                        columns=columns,
                        show="headings",
                        height=min(15, max(3, len(data_rows))),
                    )
                    for col, heading in zip(columns, headings):
                        tree.heading(col, text=heading)
                        tree.column(col, anchor="center")
                    if data_rows:
                        for data in data_rows:
                            values = list(data)
                            if len(values) < len(columns):
                                values.extend([""] * (len(columns) - len(values)))
                            tree.insert("", tk.END, values=values)
                    tree.pack(expand=True, fill=tk.BOTH)

                row_index += 1

    def _show_prompt_preview(self, category: str) -> None:
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

        dialog = tk.Toplevel(self.root)
        dialog.title(f"{category} Prompt Preview")
        dialog.transient(self.root)
        dialog.grab_set()

        content_frame = ttk.Frame(dialog, padding=(12, 12, 12, 0))
        content_frame.pack(fill=tk.BOTH, expand=True)

        text_widget = tk.Text(content_frame, wrap="word", width=80, height=25)
        text_widget.insert("1.0", prompt_text)
        text_widget.configure(state="disabled")
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(content_frame, orient=tk.VERTICAL, command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        button_frame = ttk.Frame(dialog, padding=(12, 0, 12, 12))
        button_frame.pack(fill=tk.X)
        ttk.Button(button_frame, text="Close", command=dialog.destroy).pack(side=tk.RIGHT)

        dialog.bind("<Escape>", lambda _e: dialog.destroy())
        dialog.wait_visibility()
        dialog.focus_set()

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

    def _read_csv_rows(self, csv_path: Path) -> List[List[str]]:
        try:
            with csv_path.open("r", encoding="utf-8", newline="") as fh:
                reader = csv.reader(fh)
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
            return
        try:
            with self.pattern_config_path.open("r", encoding="utf-8") as fh:
                data = json.load(fh)
                self.config_data = data
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
        self._apply_last_company_selection()
        self._config_loaded = True

    def _apply_whitespace_option(self, pattern: str, enabled: bool) -> str:
        if not enabled:
            return pattern
        return pattern.replace(" ", r"\s+")

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

    def _write_config(self) -> None:
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
                csv_exists = False
                txt_exists = False
                if isinstance(existing, dict):
                    csv_name = existing.get("csv")
                    txt_name = existing.get("txt")
                    if isinstance(csv_name, str) and csv_name:
                        csv_exists = (doc_dir / csv_name).exists()
                    if isinstance(txt_name, str) and txt_name:
                        txt_exists = (doc_dir / txt_name).exists()
                if csv_exists and txt_exists:
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
            txt_path = self._ensure_unique_path(job.doc_dir / f"{base_name}.txt")

            try:
                txt_path.write_text(response_text, encoding="utf-8")
            except Exception as exc:
                return False, None, f"{job.entry_name} - {job.category}: Could not save response ({exc})"

            rows = self._convert_response_to_rows(response_text)
            csv_path = txt_path.with_suffix(".csv")

            try:
                self._write_csv_rows(rows, csv_path)
            except Exception as exc:
                try:
                    txt_path.unlink(missing_ok=True)
                except TypeError:
                    try:
                        if txt_path.exists():
                            txt_path.unlink()
                    except FileNotFoundError:
                        pass
                    except Exception:
                        pass
                return False, None, f"{job.entry_name} - {job.category}: Could not write CSV ({exc})"

            metadata_entry = {
                "csv": csv_path.name,
                "txt": txt_path.name,
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


def main() -> None:
    root = tk.Tk()
    app = ReportApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
