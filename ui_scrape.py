from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import tkinter as tk
from tkinter import ttk

from constants import COLUMNS
from pdf_utils import PDFEntry
from scrape_panel import ScrapeResultPanel
from ui_widgets import CollapsibleFrame


class ScrapeUIMixin:
    root: tk.Misc
    scrape_tab: Optional[ttk.Frame]
    scrape_type_notebook: ttk.Notebook
    scrape_panels: Dict[Tuple[Path, str], ScrapeResultPanel]
    scrape_type_tabs: Dict[str, ttk.Frame]
    scrape_type_pdf_notebooks: Dict[str, ttk.Notebook]
    scrape_pdf_tabs_by_type: Dict[Tuple[str, Path], ttk.Frame]
    scrape_category_canvases: Dict[Tuple[Path, str], tk.Canvas]
    scrape_category_inners: Dict[Tuple[Path, str], ttk.Frame]
    scrape_category_windows: Dict[Tuple[Path, str], int]
    scrape_category_placeholders: Dict[Tuple[Path, str], Optional[tk.Widget]]
    pdf_entries: List[PDFEntry]
    company_var: tk.StringVar
    companies_dir: Path
    auto_scale_tables_var: tk.BooleanVar
    active_scrape_key: Optional[Tuple[Path, str]]
    scrape_preview_title_var: tk.StringVar
    scrape_preview_canvas: tk.Canvas
    scrape_preview_label: tk.Label
    scrape_preview_page_var: tk.StringVar
    scrape_preview_pages: List[int]
    scrape_preview_entry: Optional[PDFEntry]
    scrape_preview_category: Optional[str]
    scrape_preview_cycle_index: int
    scrape_preview_last_width: int
    scrape_preview_render_page: Optional[int]
    scrape_preview_render_width: int
    scrape_preview_photo: Any
    note_color_scheme: Dict[str, str]
    scrape_row_registry: Dict[Tuple[str, str, str], List[Tuple[ScrapeResultPanel, str]]]
    scrape_row_state_by_key: Dict[Tuple[str, str, str], str]
    scrape_column_widths: Dict[str, int]

    def build_scrape_tab(self, notebook: ttk.Notebook) -> None:
        scrape_tab = ttk.Frame(notebook)
        self.scrape_tab = scrape_tab
        notebook.add(scrape_tab, text="Scrape")

        scrape_controls = ttk.Frame(scrape_tab, padding=8)
        scrape_controls.pack(fill=tk.X)
        self.scrape_button = ttk.Button(scrape_controls, text="AIScrape", command=self.scrape_selected_pages)
        self.scrape_button.pack(side=tk.LEFT)
        self.open_scrape_dir_button = ttk.Button(
            scrape_controls, text="Open Folder", command=self.open_scrape_folder
        )
        self.open_scrape_dir_button.pack(side=tk.LEFT, padx=(6, 0))
        self.scrape_progress = ttk.Progressbar(scrape_controls, orient=tk.HORIZONTAL, mode="determinate", length=200)
        self.scrape_progress.pack(side=tk.LEFT, padx=(8, 0), fill=tk.X, expand=True)

        openai_section = CollapsibleFrame(scrape_tab, "OpenAI Settings", initially_open=False)
        openai_section.pack(fill=tk.X, padx=8)
        openai_content = openai_section.content
        settings_frame = ttk.Frame(openai_content, padding=8)
        settings_frame.pack(fill=tk.X)

        api_frame = ttk.Frame(settings_frame)
        api_frame.pack(fill=tk.X)
        ttk.Label(api_frame, text="API key:").pack(side=tk.LEFT)
        self.api_key_entry = ttk.Entry(api_frame, textvariable=self.api_key_var, width=40, show="*")
        self.api_key_entry.pack(side=tk.LEFT, padx=(4, 0), fill=tk.X, expand=True)

        model_frame = ttk.LabelFrame(settings_frame, text="OpenAI models", padding=8)
        model_frame.pack(fill=tk.X, pady=(8, 0))
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

        preview_body = ttk.Frame(preview_frame)
        preview_body.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
        preview_body.rowconfigure(0, weight=1)
        preview_body.columnconfigure(0, weight=1)

        self.scrape_preview_canvas = tk.Canvas(
            preview_body,
            background="#f0f0f0",
            highlightthickness=0,
        )
        self.scrape_preview_canvas.grid(row=0, column=0, sticky="nsew")
        preview_scrollbar = ttk.Scrollbar(
            preview_body, orient=tk.VERTICAL, command=self.scrape_preview_canvas.yview
        )
        preview_scrollbar.grid(row=0, column=1, sticky="ns")
        self.scrape_preview_canvas.configure(yscrollcommand=preview_scrollbar.set)

        self.scrape_preview_label = tk.Label(
            preview_body,
            text="Select a section to preview.",
            justify=tk.CENTER,
            anchor="center",
            background="#f0f0f0",
        )
        self.scrape_preview_window = self.scrape_preview_canvas.create_window(
            (0, 0), window=self.scrape_preview_label, anchor="n", width=0
        )
        self.scrape_preview_canvas.bind("<Configure>", self._on_scrape_preview_resize)
        self.scrape_preview_label.bind("<Configure>", self._on_scrape_preview_label_configure)
        self.scrape_preview_canvas.bind("<MouseWheel>", self._on_scrape_preview_mousewheel)
        self.scrape_preview_canvas.bind("<Button-4>", self._on_scrape_preview_mousewheel)
        self.scrape_preview_label.bind("<Button-5>", self._on_scrape_preview_mousewheel)
        self.scrape_preview_label.bind("<Button-1>", self._on_scrape_preview_click)
        self.scrape_preview_label.bind("<MouseWheel>", self._on_scrape_preview_mousewheel)

        # === Keyboard navigation for Scrape tab ===
        self.root.bind("a", lambda e: self._navigate_scrape_pdf_tab(-1))
        self.root.bind("d", lambda e: self._navigate_scrape_pdf_tab(1))
        self.root.bind("q", lambda e: self._navigate_scrape_type_tab(-1))
        self.root.bind("e", lambda e: self._navigate_scrape_type_tab(1))
        self.scrape_preview_label.bind("<Button-4>", self._on_scrape_preview_mousewheel)
        self.scrape_preview_label.bind("<Button-5>", self._on_scrape_preview_mousewheel)

        tables_frame = ttk.Frame(scrape_split)
        scrape_split.add(tables_frame, weight=1)

        self.scrape_type_notebook = ttk.Notebook(tables_frame)
        self.scrape_type_notebook.pack(fill=tk.BOTH, expand=True)

        self.scrape_type_tabs = {}
        self.scrape_type_pdf_notebooks = {}
        self.scrape_pdf_tabs_by_type = {}
        self.scrape_category_canvases = {}
        self.scrape_category_inners = {}
        self.scrape_category_windows = {}
        self.scrape_category_placeholders = {}

    def _navigate_scrape_pdf_tab(self, delta: int) -> None:
        """Switch to previous/next PDF filename tab within the active TYPE tab."""
        type_nb = self.scrape_type_notebook
        type_tab = type_nb.select()
        if not type_tab:
            raise RuntimeError("No active TYPE tab in scrape_type_notebook.")

        active_category = type_nb.tab(type_tab, "text")
        inner_nb = self.scrape_type_pdf_notebooks.get(active_category)
        if inner_nb is None:
            raise RuntimeError("Active TYPE tab missing inner notebook of PDF filenames.")

        total_tabs = inner_nb.index("end")
        if total_tabs <= 0:
            raise RuntimeError("No PDF tabs available.")

        current_idx = inner_nb.index(inner_nb.select())
        new_idx = (current_idx + delta) % total_tabs
        inner_nb.select(new_idx)

    def _navigate_scrape_type_tab(self, delta: int) -> None:
        """Switch to previous/next TYPE tab (Financial, Income, Shares) in Scrape view."""
        notebook = self.scrape_type_notebook
        total_tabs = notebook.index("end")
        if total_tabs <= 0:
            raise RuntimeError("No TYPE tabs found in Scrape view.")

        current_idx = notebook.index(notebook.select())
        new_idx = (current_idx + delta) % total_tabs
        notebook.select(new_idx)

    def _refresh_scrape_results(self) -> None:
        for panel in self.scrape_panels.values():
            panel.destroy()
        self.scrape_panels.clear()

        for child in list(self.scrape_type_notebook.winfo_children()):
            child.destroy()

        self.scrape_type_tabs.clear()
        self.scrape_type_pdf_notebooks.clear()
        self.scrape_pdf_tabs_by_type.clear()
        self.scrape_category_canvases.clear()
        self.scrape_category_inners.clear()
        self.scrape_category_windows.clear()
        self.scrape_category_placeholders.clear()

        if not self.pdf_entries:
            placeholder_tab = ttk.Frame(self.scrape_type_notebook)
            ttk.Label(
                placeholder_tab,
                text="Load PDFs and choose pages to prepare scraping results.",
                foreground="#666666",
                wraplength=360,
                justify=tk.LEFT,
            ).pack(anchor="w", padx=12, pady=12)
            self.scrape_type_notebook.add(placeholder_tab, text="No PDFs")
            self.scrape_type_notebook.tab(placeholder_tab, state="disabled")
            self._clear_scrape_preview()
            self.active_scrape_key = None
            self.refresh_combined_tab()
            return

        for category in COLUMNS:
            type_tab = ttk.Frame(self.scrape_type_notebook)
            self.scrape_type_notebook.add(type_tab, text=category)
            self.scrape_type_tabs[category] = type_tab

            inner_notebook = ttk.Notebook(type_tab)
            inner_notebook.pack(fill=tk.BOTH, expand=True)
            self.scrape_type_pdf_notebooks[category] = inner_notebook
            inner_notebook.bind(
                "<<NotebookTabChanged>>",
                lambda _e, cat=category: self._on_scrape_inner_pdf_tab_changed(cat),
            )

        self.scrape_type_notebook.bind("<<NotebookTabChanged>>", self._on_scrape_type_tab_changed)

        for category in COLUMNS:
            inner_nb = self.scrape_type_pdf_notebooks.get(category)
            if inner_nb is None:
                continue
            for entry in self.pdf_entries:
                tab = ttk.Frame(inner_nb)

                # === NEW: Use Year (from Review tab) for nicer tab text ===
                base_name = entry.path.stem
                if hasattr(entry, "year") and entry.year:
                    try:
                        year_str = str(entry.year).strip()
                        display = f"{base_name} ({year_str})"
                    except Exception:
                        display = base_name
                else:
                    display = base_name

                inner_nb.add(tab, text=display)
                self.scrape_pdf_tabs_by_type[(category, entry.path)] = tab

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
                self.scrape_category_canvases[key] = canvas
                self.scrape_category_inners[key] = inner
                self.scrape_category_windows[key] = window
                self.scrape_category_placeholders[key] = placeholder

        company = self.company_var.get().strip()
        entry_lookup: Dict[Path, PDFEntry] = {entry.path: entry for entry in self.pdf_entries}
        default_entry: Optional[PDFEntry] = None
        default_category: Optional[str] = None

        for entry in self.pdf_entries:
            if company:
                target_base_base = self.companies_dir / company / "openapiscrape" / entry.path.stem
            else:
                target_base_base = entry.path.parent / "openapiscrape" / entry.path.stem
            for category in COLUMNS:
                key = (entry.path, category)
                parent_inner = self.scrape_category_inners.get(key)
                if parent_inner is None:
                    continue
                pages = self.get_selected_pages(entry, category)
                target_base = target_base_base
                csv_path = target_base / f"{category}.csv"
                multiplier_path = target_base / f"{category}_multiplier.txt"
                if not pages and not csv_path.exists() and not multiplier_path.exists():
                    continue
                placeholder = self.scrape_category_placeholders.get(key)
                if placeholder is not None:
                    placeholder.destroy()
                    self.scrape_category_placeholders[key] = None
                panel = ScrapeResultPanel(
                    parent_inner,
                    self,
                    entry,
                    category,
                    target_base,
                    self.auto_scale_tables_var.get(),
                )
                panel.load_from_files()
                panel.update_note_coloring()
                self.scrape_panels[key] = panel
                if panel.has_csv_data or pages:
                    if default_entry is None:
                        default_entry = entry
                        default_category = category

        if not self.scrape_panels:
            self._clear_scrape_preview()
            self.active_scrape_key = None
            self.refresh_combined_tab()
            return

        if default_entry is None or default_category is None:
            first_path, first_category = next(iter(self.scrape_panels.keys()))
            default_entry = entry_lookup.get(first_path)
            default_category = first_category

        if default_entry is None or default_category is None:
            self.refresh_combined_tab()
            return

        self.set_active_scrape_panel(default_entry, default_category)
        self.refresh_combined_tab()

    def _clear_scrape_preview(self) -> None:
        self.scrape_preview_photo = None
        self.scrape_preview_canvas.configure(background="#f0f0f0")
        self.scrape_preview_label.configure(image="", text="Select a section to preview.", background="#f0f0f0")
        self.scrape_preview_title_var.set("Select a section to preview.")
        self.scrape_preview_page_var.set("")
        self.scrape_preview_pages = []
        self.scrape_preview_entry = None
        self.scrape_preview_category = None
        self.scrape_preview_cycle_index = 0
        self.scrape_preview_render_page = None
        self.scrape_preview_render_width = 0
        self._reset_scrape_preview_scroll()

    def _on_scrape_panel_clicked(self, panel: ScrapeResultPanel) -> None:
        self.set_active_scrape_panel(panel.entry, panel.category)

    def register_scrape_row(
        self, panel: ScrapeResultPanel, item_id: str, key: Tuple[str, str, str]
    ) -> None:
        entries = self.scrape_row_registry.setdefault(key, [])
        entries = [entry for entry in entries if entry[0] is not panel or entry[1] != item_id]
        entries.append((panel, item_id))
        self.scrape_row_registry[key] = entries

    def unregister_panel_rows(self, panel: ScrapeResultPanel) -> None:
        to_remove: List[Tuple[str, str, str]] = []
        for key, entries in self.scrape_row_registry.items():
            filtered = [entry for entry in entries if entry[0] is not panel]
            if filtered:
                self.scrape_row_registry[key] = filtered
            else:
                to_remove.append(key)
        for key in to_remove:
            self.scrape_row_registry.pop(key, None)

    def apply_row_state_to_all(self, key: Tuple[str, str, str], state: Optional[str]) -> None:
        if state in (None, ""):
            self.scrape_row_state_by_key.pop(key, None)
        else:
            self.scrape_row_state_by_key[key] = state

        for panel, item_id in list(self.scrape_row_registry.get(key, [])):
            panel.update_row_state(item_id, state)
            note_idx = panel._note_column_index()
            if note_idx is not None:
                vals = list(panel.table.item(item_id, "values"))
                if len(vals) <= note_idx:
                    vals.extend([""] * (note_idx + 1 - len(vals)))
                vals[note_idx] = state if state else "asis"
                panel.table.item(item_id, values=vals)
                panel._apply_note_color_to_item(item_id)
            panel.table.update_idletasks()
            panel.save_table_to_csv()

    def set_active_scrape_panel(self, entry: PDFEntry, category: str) -> None:
        key = (entry.path, category)
        self.active_scrape_key = key
        type_tab = self.scrape_type_tabs.get(category)
        if type_tab is not None:
            try:
                self.scrape_type_notebook.select(type_tab)
            except Exception:
                pass
        inner_nb = self.scrape_type_pdf_notebooks.get(category)
        if inner_nb is not None:
            pdf_tab = self.scrape_pdf_tabs_by_type.get((category, entry.path))
            if pdf_tab is not None:
                try:
                    inner_nb.select(pdf_tab)
                except Exception:
                    pass
        for panel_key, panel in self.scrape_panels.items():
            panel.set_active(panel_key == key)
        self._show_scrape_preview(entry, category)

    def reload_scrape_panels(self) -> None:
        for key, panel in self.scrape_panels.items():
            panel.load_from_files()
            panel.update_note_coloring()
            panel.set_active(self.active_scrape_key == (panel.entry, panel.category) if isinstance(panel.entry, Path) else self.active_scrape_key == (panel.entry.path, panel.category))
        if self.active_scrape_key is not None:
            path, category = self.active_scrape_key
            entry = self._get_entry_by_path(path)
            if entry is not None:
                self._show_scrape_preview(entry, category)
        self.refresh_combined_tab()

    def apply_note_colors_to_all_panels(self) -> None:
        for panel in self.scrape_panels.values():
            panel.update_note_coloring()
        self.refresh_combined_tab()

    def get_scrape_column_width(self, heading_text: str) -> int:
        key = heading_text.strip().lower()
        if key in ("category", "subcategory", "item", "note"):
            return int(self.scrape_column_widths.get(key, 140))
        return int(self.scrape_column_widths.get("dates", 140))

    def _on_scrape_type_tab_changed(self, event: tk.Event) -> None:  # type: ignore[override]
        try:
            widget = event.widget
        except Exception:
            return
        if widget is not self.scrape_type_notebook:
            return
        selected_id = self.scrape_type_notebook.select()
        if not selected_id:
            return
        selected_widget = self.root.nametowidget(selected_id)
        selected_category: Optional[str] = None
        for cat, frame in self.scrape_type_tabs.items():
            if frame is selected_widget:
                selected_category = cat
                break
        if selected_category is None:
            return
        inner_nb = self.scrape_type_pdf_notebooks.get(selected_category)
        if inner_nb is None:
            return
        pdf_sel_id = inner_nb.select()
        if not pdf_sel_id:
            children = inner_nb.winfo_children()
            if not children:
                return
            try:
                inner_nb.select(children[0])
                pdf_widget = children[0]
            except Exception:
                pdf_widget = children[0]
        else:
            pdf_widget = self.root.nametowidget(pdf_sel_id)
        selected_path: Optional[Path] = None
        for (cat, path), frame in self.scrape_pdf_tabs_by_type.items():
            if cat == selected_category and frame is pdf_widget:
                selected_path = path
                break
        if selected_path is None:
            return
        entry = self._get_entry_by_path(selected_path)
        if entry is None:
            return
        self.set_active_scrape_panel(entry, selected_category)

    def _on_scrape_inner_pdf_tab_changed(self, category: str) -> None:
        inner_nb = self.scrape_type_pdf_notebooks.get(category)
        if inner_nb is None:
            return
        sel_id = inner_nb.select()
        if not sel_id:
            return
        sel_widget = self.root.nametowidget(sel_id)
        selected_path: Optional[Path] = None
        for (cat, path), frame in self.scrape_pdf_tabs_by_type.items():
            if cat == category and frame is sel_widget:
                selected_path = path
                break
        if selected_path is None:
            return
        entry = self._get_entry_by_path(selected_path)
        if entry is None:
            return
        self.set_active_scrape_panel(entry, category)

    def _get_entry_by_path(self, path: Path) -> Optional[PDFEntry]:
        for entry in self.pdf_entries:
            if entry.path == path:
                return entry
        return None

    def get_note_color(self, note_value: str) -> Optional[str]:
        value = note_value.strip().lower()
        if not value:
            return None
        lower_map = {k.lower(): v for k, v in self.note_color_scheme.items()}
        return lower_map.get(value)
