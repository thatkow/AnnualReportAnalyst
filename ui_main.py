from __future__ import annotations

from pathlib import Path
from typing import Any, Dict

import tkinter as tk
from tkinter import colorchooser, messagebox, simpledialog, ttk

from constants import DEFAULT_NOTE_COLOR_SCHEME


class MainUIMixin:
    root: tk.Misc
    notebook: ttk.Notebook
    scrape_type_notebook: ttk.Notebook
    auto_scale_tables_var: tk.BooleanVar
    auto_load_last_company_var: tk.BooleanVar
    note_color_scheme: Dict[str, str]
    scrape_column_widths: Dict[str, int]
    scrape_panels: Dict[Any, Any]
    combined_tab: Any
    company_var: tk.StringVar
    folder_path: tk.StringVar

    def _build_ui(self) -> None:
        menu_bar = tk.Menu(self.root)

        file_menu = tk.Menu(menu_bar, tearoff=False)
        file_menu.add_command(label="New Company", command=self.create_company)
        file_menu.add_command(label="Set Downloads Dir", command=self._set_downloads_dir)
        menu_bar.add_cascade(label="File", menu=file_menu)

        company_menu = tk.Menu(menu_bar, tearoff=False)
        def _select_and_load():
            self._open_company_selector()
            self.load_pdfs()

        company_menu.add_command(label="Select Company & Load PDFs…", command=_select_and_load)
        company_menu.add_checkbutton(
            label="Auto Load Last Company on Startup",
            variable=self.auto_load_last_company_var,
            command=self._on_toggle_auto_load_last_company,
        )
        menu_bar.add_cascade(label="Company", menu=company_menu)

        view_menu = tk.Menu(menu_bar, tearoff=False)
        view_menu.add_checkbutton(
            label="Auto Scale Scrape Tables",
            variable=self.auto_scale_tables_var,
            command=self._on_toggle_auto_scale_tables,
        )
        view_menu.add_command(
            label="Configure Scrape Column Widths…",
            command=self._configure_scrape_column_widths,
        )
        view_menu.add_command(
            label="Configure Note Colors…",
            command=self._configure_note_colors,
        )
        menu_bar.add_cascade(label="View", menu=view_menu)

        try:
            self.root.config(menu=menu_bar)
        except tk.TclError:
            pass

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        self.notebook.bind("<<NotebookTabChanged>>", self._on_main_tab_changed)

        if hasattr(self, "build_review_tab"):
            self.build_review_tab(self.notebook)
        if hasattr(self, "build_scrape_tab"):
            self.build_scrape_tab(self.notebook)
        if hasattr(self, "build_combined_tab"):
            self.build_combined_tab(self.notebook)

    def _on_main_tab_changed(self, event: tk.Event) -> None:  # type: ignore[override]
        try:
            sel = self.notebook.select()
            if not sel:
                return
            widget = self.root.nametowidget(sel)
        except Exception:
            return
        if getattr(self, "combined_tab", None) is not None and widget is self.combined_tab:
            self.refresh_combined_tab()

    def _on_toggle_auto_scale_tables(self) -> None:
        enabled = self.auto_scale_tables_var.get()
        for panel in self.scrape_panels.values():
            panel.set_auto_scale(enabled)

    def _on_toggle_auto_load_last_company(self) -> None:
        self._save_config()

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

    def _configure_scrape_column_widths(self) -> None:
        window = tk.Toplevel(self.root)
        window.title("Configure Scrape Column Widths")
        window.transient(self.root)
        window.resizable(False, False)

        frame = ttk.Frame(window, padding=12)
        frame.pack(fill=tk.BOTH, expand=True)

        entries: Dict[str, tk.Entry] = {}
        for idx, heading in enumerate(["category", "subcategory", "item", "note", "dates"]):
            row = ttk.Frame(frame)
            row.pack(fill=tk.X, pady=4)
            ttk.Label(row, text=heading.capitalize(), width=15).pack(side=tk.LEFT)
            var = tk.StringVar(value=str(self.scrape_column_widths.get(heading, 140)))
            entry = ttk.Entry(row, textvariable=var, width=10)
            entry.pack(side=tk.LEFT)
            entries[heading] = entry

        button_row = ttk.Frame(frame)
        button_row.pack(fill=tk.X, pady=(12, 0))

        def on_save() -> None:
            for key, entry in entries.items():
                value = entry.get().strip()
                try:
                    width = int(value)
                except ValueError:
                    messagebox.showerror("Invalid Width", f"Width for '{key}' must be an integer.")
                    return
                self.scrape_column_widths[key] = max(40, min(width, 800))
            self._save_pattern_config()
            window.destroy()

        ttk.Button(button_row, text="Cancel", command=window.destroy).pack(side=tk.RIGHT)
        ttk.Button(button_row, text="Save", command=on_save).pack(side=tk.RIGHT, padx=(0, 8))

        window.grab_set()
        window.focus_force()

    def _configure_note_colors(self) -> None:
        window = tk.Toplevel(self.root)
        window.title("Configure Note Colors")
        window.transient(self.root)
        window.resizable(False, False)

        frame = ttk.Frame(window, padding=12)
        frame.pack(fill=tk.BOTH, expand=True)

        columns = ("value", "color")
        tree = ttk.Treeview(frame, columns=columns, show="headings", height=8)
        tree.heading("value", text="Note Value")
        tree.heading("color", text="Color (hex)")
        tree.column("value", width=140, anchor=tk.W)
        tree.column("color", width=120, anchor=tk.W)
        tree.pack(fill=tk.BOTH, expand=True)

        local_map: Dict[str, str] = dict(self.note_color_scheme)

        def refresh_view() -> None:
            for iid in tree.get_children(""):
                tree.delete(iid)
            for key in sorted(local_map.keys(), key=lambda s: (s.isdigit(), int(s) if s.isdigit() else s)):
                tree.insert("", "end", values=(key, local_map[key]))

        btns = ttk.Frame(frame)
        btns.pack(fill=tk.X, pady=(8, 0))

        def add_mapping() -> None:
            key = simpledialog.askstring("Add Note Mapping", "Enter note value:", parent=window)
            if key is None:
                return
            key = key.strip()
            if not key:
                return
            initial = local_map.get(key, "#FFF2CC")
            color = colorchooser.askcolor(color=initial, title=f"Choose color for note '{key}'")[1]
            if not color:
                return
            local_map[key] = color
            refresh_view()

        def edit_color() -> None:
            sel = tree.selection()
            if not sel:
                return
            values = tree.item(sel[0], "values")
            if not values:
                return
            key = str(values[0])
            current = local_map.get(key, "#FFF2CC")
            color = colorchooser.askcolor(color=current, title=f"Choose color for note '{key}'")[1]
            if not color:
                return
            local_map[key] = color
            refresh_view()

        def remove_mapping() -> None:
            sel = tree.selection()
            if not sel:
                return
            values = tree.item(sel[0], "values")
            if not values:
                return
            key = str(values[0])
            if key in local_map:
                del local_map[key]
            refresh_view()

        def reset_defaults() -> None:
            local_map.clear()
            local_map.update(DEFAULT_NOTE_COLOR_SCHEME)
            refresh_view()

        def on_save() -> None:
            self.note_color_scheme = dict(local_map)
            self._save_pattern_config()
            self.apply_note_colors_to_all_panels()
            window.destroy()

        def on_cancel() -> None:
            window.destroy()

        ttk.Button(btns, text="Add", command=add_mapping).pack(side=tk.LEFT)
        ttk.Button(btns, text="Set Color", command=edit_color).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(btns, text="Remove", command=remove_mapping).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(btns, text="Reset Defaults", command=reset_defaults).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(btns, text="Cancel", command=on_cancel).pack(side=tk.RIGHT)
        ttk.Button(btns, text="Save", command=on_save).pack(side=tk.RIGHT, padx=(0, 8))

        refresh_view()
        window.grab_set()
        window.focus_force()

    def _maybe_auto_load_last_company(self) -> None:
        try:
            if not self.auto_load_last_company_var.get():
                return
        except Exception:
            return
        company = self.company_var.get().strip()
        if not company:
            return
        folder_str = self.folder_path.get().strip()
        if not folder_str:
            return
        folder = Path(folder_str)
        if not folder.exists():
            return
        has_pdf = any(folder.rglob("*.pdf"))
        if not has_pdf:
            return
        self.root.after(0, self.load_pdfs)
