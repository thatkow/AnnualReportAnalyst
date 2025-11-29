from __future__ import annotations

import logging
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


from app_logging import get_logger
from company_manager import CompanyManagerMixin
from config_manager import ConfigManager
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
):
    def __init__(self, root: tk.Misc) -> None:
        if fitz is None:
            messagebox.showerror("PyMuPDF Required", "Install PyMuPDF (import name 'fitz') to use this app.")
            raise RuntimeError("PyMuPDF (fitz) is not installed")

        self.root = root
        try:
            self.root.title("Annual Report Analyst (Preview)")
        except tk.TclError:
            pass
        self.root.after(0, self._maximize_window)

        self.app_root = Path(__file__).resolve().parent
        self.companies_dir = self.app_root / "companies"
        self.prompts_dir = self.app_root / "prompts"

        self.company_var = tk.StringVar(master=self.root)
        self.company_options: List[str] = []
        self.company_selector_window: Optional[tk.Toplevel] = None
        self.folder_path = tk.StringVar(master=self.root)
        self.thumbnail_width_var = tk.IntVar(master=self.root, value=220)
        self.api_key_var = tk.StringVar(master=self.root)
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
        self._suspend_openai_model_save = False
        self.config = ConfigManager.load()
        self._apply_config_state()
        self._suspend_api_key_save = False
        self._build_ui()
        self._load_pattern_config()
        self._load_config()
        # ---------------------- Logger setup ----------------------
        self.logger = get_logger()
        if not self.logger.handlers:
            handler = logging.StreamHandler()
            handler.setFormatter(
                logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s")
            )
            self.logger.addHandler(handler)
            self.logger.setLevel(logging.INFO)

        self.logger.info("✅ Logger initialized for ReportAppV2")

        # ---------------------- AIScrape thread count setup ----------------------
        try:
            self.thread_count = self.get_thread_count()
            self.logger.info("🔁 Loaded AIScrape thread count = %d", self.thread_count)
        except Exception as e:
            self.logger.warning("⚠️ Could not load AIScrape thread count from config: %s", e)
            self.thread_count = 3
            self.set_thread_count(self.thread_count)
            self.logger.info("🆕 Initialized AIScrape thread count = %d (default)", self.thread_count)

        self._refresh_company_options()
        self._maybe_auto_load_last_company()

        self.logger.info("Logger initialized and bound to ReportAppV2")
        self.logger.info("✅ Shared logger setup complete")

    def _create_openai_model_var(self, column: str, initial: str) -> tk.StringVar:
        """Return an OpenAI model StringVar that auto-saves on change."""

        var = tk.StringVar(master=self.root, value=initial)
        var.trace_add("write", lambda *_: self._on_openai_model_changed(column))
        return var

    def _apply_config_state(self) -> None:
        """Populate runtime state from the shared configuration object."""

        self.scrape_row_height = int(getattr(self.config, "scrape_row_height", 22))
        self.scrape_column_widths = dict(getattr(self.config, "scrape_column_widths", {})) or {
            "category": 140,
            "subcategory": 140,
            "item": 140,
            "note": 140,
            "dates": 140,
        }
        note_colors = dict(getattr(self.config, "note_colors", {})) or dict(DEFAULT_NOTE_COLOR_SCHEME)
        if note_colors:
            self.note_color_scheme = note_colors

        downloads_dir = getattr(self.config, "downloads_dir", "") or ""
        if downloads_dir:
            self.downloads_dir.set(downloads_dir)
        else:
            self.downloads_dir.set("")

        minutes = getattr(self.config, "downloads_minutes", 5)
        try:
            self.recent_download_minutes.set(int(minutes))
        except Exception:
            self.recent_download_minutes.set(5)

        last_company = getattr(self.config, "last_company", "") or ""
        if last_company:
            self.company_var.set(last_company)

        auto_load = bool(getattr(self.config, "auto_load_last_company", False))
        self.auto_load_last_company_var.set(auto_load)

        api_key = getattr(self.config, "api_key", "") or ""
        self.api_key_var.set(api_key)

    # ------------------------------------------------------------------ Pattern helpers
    def _read_text_lines(self, widget: tk.Text) -> List[str]:
        text = widget.get("1.0", tk.END)
        lines: List[str] = []
        for raw in text.splitlines():
            cleaned = raw.strip()
            if cleaned:
                lines.append(cleaned)
        return lines

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
            value = var.get().strip() if var is not None else ""
            openai_models[column] = value or DEFAULT_OPENAI_MODEL
        upload_modes: Dict[str, str] = {}
        for column in COLUMNS:
            mode_var = self.scrape_upload_mode_vars.get(column)
            upload_modes[column] = (mode_var.get() if mode_var is not None else "pdf") or "pdf"
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
            "scrape_row_height": int(getattr(self, "scrape_row_height", 22)),
        }
        return payload

    def _save_pattern_config(self) -> None:
        payload = self._collect_pattern_config_payload()
        self.config.patterns = payload["patterns"]
        self.config.case_insensitive = payload["case_insensitive"]
        self.config.space_as_whitespace = payload["space_as_whitespace"]
        self.config.year_patterns = payload["year_patterns"]
        self.config.year_case_insensitive = payload["year_case_insensitive"]
        self.config.year_space_as_whitespace = payload["year_space_as_whitespace"]
        self.config.downloads_minutes = payload["downloads_minutes"]
        self.config.openai_models = payload["openai_models"]
        self.config.upload_modes = payload["upload_modes"]
        self.config.scrape_column_widths = payload["scrape_column_widths"]
        self.config.note_colors = payload["note_colors"]
        self.config.scrape_row_height = payload["scrape_row_height"]
        self.scrape_row_height = self.config.scrape_row_height
        self.note_color_scheme = dict(self.config.note_colors)
        try:
            self.config.save()
        except OSError:
            messagebox.showwarning(
                "Save Patterns", "Unable to save pattern configuration to disk."
            )

    def _load_pattern_config(self) -> None:
        config = getattr(self, "config", None)
        if config is None:
            return

        self._suspend_openai_model_save = True
        try:
            for column, widget in self.pattern_texts.items():
                values = config.patterns.get(column, [])
                widget.delete("1.0", tk.END)
                widget.insert("1.0", "\n".join(values))

            for column, var in self.case_insensitive_vars.items():
                var.set(config.case_insensitive.get(column, True))

            for column, var in self.whitespace_as_space_vars.items():
                var.set(config.space_as_whitespace.get(column, True))

            if self.year_pattern_text is not None:
                self.year_pattern_text.delete("1.0", tk.END)
                self.year_pattern_text.insert("1.0", "\n".join(config.year_patterns))

            self.year_case_insensitive_var.set(config.year_case_insensitive)
            self.year_whitespace_as_space_var.set(config.year_space_as_whitespace)

            try:
                self.recent_download_minutes.set(int(config.downloads_minutes))
            except Exception:
                self.recent_download_minutes.set(5)

            for column, var in self.openai_model_vars.items():
                var.set(config.openai_models.get(column, DEFAULT_OPENAI_MODEL))

            for column, var in self.scrape_upload_mode_vars.items():
                var.set(config.upload_modes.get(column, "pdf"))

            self.scrape_column_widths = dict(config.scrape_column_widths)
            self.note_color_scheme = dict(config.note_colors)
            self.scrape_row_height = int(config.scrape_row_height)
        finally:
            self._suspend_openai_model_save = False

    def get_scrape_row_height(self) -> int:
        """Return the configured scrape table row height."""

        return int(getattr(self.config, "scrape_row_height", getattr(self, "scrape_row_height", 22)))

    def set_scrape_row_height(self, value: int) -> None:
        """Update and persist the scrape table row height."""

        try:
            height = int(value)
        except (TypeError, ValueError):
            height = 22
        height = max(10, min(height, 60))
        self.scrape_row_height = height
        self.config.scrape_row_height = height
        try:
            self.config.save()
        except OSError:
            messagebox.showwarning(
                "Save Patterns", "Unable to save pattern configuration to disk."
            )

    def _load_config(self) -> None:
        self._suspend_api_key_save = True
        try:
            self._apply_config_state()
        finally:
            self._suspend_api_key_save = False

    def _save_config(self) -> None:
        self.config.downloads_dir = self.downloads_dir.get().strip()
        self.config.last_company = self.company_var.get().strip()
        self.config.auto_load_last_company = bool(self.auto_load_last_company_var.get())
        try:
            self.config.save()
        except OSError:
            messagebox.showwarning(
                "Save Config", "Unable to save configuration to disk."
            )

    def get_thread_count(self, default: int = 3) -> int:
        """Retrieve stored thread count from config, or use default."""

        value = getattr(self.config, "thread_count", default)
        if isinstance(value, int) and value > 0:
            self.logger.info("🔁 Restored thread count = %d from config", value)
            return value
        return default

    def set_thread_count(self, value: int) -> None:
        """Save new thread count to config file."""

        try:
            count = int(value)
            if count <= 0:
                raise ValueError
        except (TypeError, ValueError):
            count = max(1, int(getattr(self.config, "thread_count", 3)))
        self.config.thread_count = count
        try:
            self.config.save()
            self.logger.info("💾 Saved thread count = %d to config", count)
        except OSError as exc:
            self.logger.warning("⚠️ Could not save thread count: %s", exc)

    def _persist_api_key(self, value: str) -> None:
        trimmed = value.strip()
        if trimmed == getattr(self.config, "api_key", ""):
            return
        self.config.api_key = trimmed
        try:
            self.config.save()
        except OSError:
            messagebox.showwarning(
                "Local Config", "Unable to save API key to configuration file."
            )

    def _on_openai_model_changed(self, column: str, *_: Any) -> None:
        if self._suspend_openai_model_save:
            return

        var = self.openai_model_vars.get(column)
        if var is None:
            return

        value = var.get().strip() or DEFAULT_OPENAI_MODEL
        if self.config.openai_models.get(column) == value:
            return

        self.config.openai_models[column] = value
        try:
            self.config.save()
        except OSError:
            messagebox.showwarning(
                "Save Config", "Unable to save OpenAI model configuration to disk."
            )

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
