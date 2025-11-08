from __future__ import annotations

import threading
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import tkinter as tk
from tkinter import messagebox, ttk

try:
    import fitz  # type: ignore[import-untyped]
except ImportError:  # pragma: no cover - handled at runtime
    fitz = None  # type: ignore[assignment]

from PIL import ImageTk


from company_manager import CompanyManagerMixin
from config_manager import ConfigManagerMixin
from constants import COLUMNS, DEFAULT_OPENAI_MODEL, DEFAULT_NOTE_COLOR_SCHEME, FALLBACK_NOTE_PALETTE
from pdf_manager import PDFManagerMixin
from pdf_utils import PDFEntry
from scrape_manager import ScrapeManagerMixin
from scrape_panel import ScrapeResultPanel
from ui_combined import CombinedUIMixin
from ui_main import MainUIMixin
from ui_review import ReviewUIMixin
from ui_scrape import ScrapeUIMixin
from ui_widgets import CategoryRow


class ReportAppV2(
    ReviewUIMixin,
    ScrapeUIMixin,
    CombinedUIMixin,
    PDFManagerMixin,
    ScrapeManagerMixin,
    CompanyManagerMixin,
    MainUIMixin,
    ConfigManagerMixin,
):
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
        self.company_options: List[str] = []
        self.company_selector_window: Optional[tk.Toplevel] = None
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
        self.auto_scale_tables_var = tk.BooleanVar(master=self.root, value=True)
        self.scrape_column_widths: Dict[str, int] = {
            "category": 140,
            "subcategory": 140,
            "item": 140,
            "note": 140,
            "dates": 140,
        }
        # Note color scheme and fallback palette
        self.note_color_scheme: Dict[str, str] = dict(DEFAULT_NOTE_COLOR_SCHEME)
        self.fallback_note_palette: List[str] = list(FALLBACK_NOTE_PALETTE)

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
        self.scrape_row_registry: Dict[Tuple[str, str, str], List[Tuple[ScrapeResultPanel, str]]] = {}
        self.scrape_row_state_by_key: Dict[Tuple[str, str, str], str] = {}
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
        # Auto-load last company toggle
        self.auto_load_last_company_var = tk.BooleanVar(master=self.root, value=False)

        # Combined tab state
        self.combined_date_tree: Optional[ttk.Treeview] = None
        self.combined_table: Optional[ttk.Treeview] = None
        self.combined_create_button: Optional[ttk.Button] = None
        self.combined_save_button: Optional[ttk.Button] = None
        self.combined_columns: List[str] = []
        self.combined_rows: List[List[str]] = []
        self.combined_tab: Optional[ttk.Frame] = None
        # Removed rename text boxes; keep placeholders for compatibility
        self.combined_rename_canvas: Optional[tk.Canvas] = None
        self.combined_rename_inner: Optional[ttk.Frame] = None
        self.combined_rename_scroll: Optional[ttk.Scrollbar] = None
        self.combined_rename_vars: List[tk.StringVar] = []
        self.combined_dyn_columns: List[Dict[str, Any]] = []
        self.combined_rename_names: List[str] = []  # dynamic column names used as headers
        self.combined_date_all_col_ids: List[str] = []
        self.combined_table_col_ids: List[str] = []

        self._suspend_api_key_save = True
        self._load_local_config()
        self._suspend_api_key_save = False
        self._build_ui()
        self._load_pattern_config()
        self._load_config()
        self._refresh_company_options()
        self._maybe_auto_load_last_company()

        # ---------------------- Logger setup ----------------------
        import logging
        self.logger = logging.getLogger("annualreport")
        if not self.logger.handlers:
            handler = logging.StreamHandler()
            handler.setFormatter(
                logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s")
            )
            self.logger.addHandler(handler)
            self.logger.setLevel(logging.INFO)
        # ---------------------- Initialize ScrapeManagerMixin ----------------------
        try:
            # Find correct MRO base with a defined __init__ before calling
            if hasattr(ScrapeManagerMixin, "__init__"):
                super(ScrapeManagerMixin, self).__init__()
            # Explicitly assign the logger to this instance
            self.logger.info("Logger initialized and bound to ReportAppV2")
        except Exception as e:
            print(f"⚠️ Logger setup warning: {e}")
            import traceback; traceback.print_exc()

        # Bind logger to ScrapeManagerMixin explicitly
        self.logger.info("✅ Shared logger setup complete")

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
            "scrape_column_widths": dict(self.scrape_column_widths),
            "note_colors": dict(self.note_color_scheme),
        }
        return payload


    # ------------------------------------------------------------------ Pattern helpers
    def _read_text_lines(self, widget: tk.Text) -> List[str]:
        text = widget.get("1.0", tk.END)
        lines: List[str] = []
        for raw in text.splitlines():
            cleaned = raw.strip()
            if cleaned:
                lines.append(cleaned)
        return lines
