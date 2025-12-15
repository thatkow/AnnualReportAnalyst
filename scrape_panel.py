"""Scrape result panel widget coordinating model, view, and context menu."""

from __future__ import annotations

import os
from pathlib import Path
from typing import List, Optional, TYPE_CHECKING

import tkinter as tk
from tkinter import messagebox, simpledialog, ttk

from pdf_utils import PDFEntry
from scrape_context_menu import ScrapeContextMenu
from scrape_table_model import ScrapeTableModel
from scrape_table_view import ScrapeTableView

if TYPE_CHECKING:  # pragma: no cover
    from report_app import ReportAppV2


class ScrapeResultPanel:
    def __init__(
        self,
        parent: tk.Widget,
        app: "ReportAppV2",
        entry: PDFEntry,
        category: str,
        target_dir: Path,
        auto_scale_tables: bool,
    ) -> None:
        self.app = app
        self.entry = entry
        self.category = category
        self.target_dir = target_dir

        self.model = ScrapeTableModel(app, entry, category, target_dir)
        self.row_states = self.model.row_states
        self.row_keys = self.model.row_keys
        self._flip_column_index: Optional[int] = None

        self.container = tk.Frame(
            parent,
            highlightbackground="#c3c3c3",
            highlightcolor="#c3c3c3",
            highlightthickness=1,
            bd=1,
            relief=tk.FLAT,
        )
        self.container.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        self.frame = ttk.Frame(self.container, padding=8)
        self.frame.pack(fill=tk.BOTH, expand=True)

        header = ttk.Frame(self.frame)
        header.pack(fill=tk.X)
        title_text = f"{entry.path.name} – {category}"
        self.title_label = ttk.Label(
            header, text=title_text, font=("TkDefaultFont", 10, "bold")
        )
        self.title_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        multiplier_box = ttk.Frame(header)
        multiplier_box.pack(side=tk.RIGHT)
        ttk.Label(multiplier_box, text="Multiplier:").pack(side=tk.LEFT)
        self.multiplier_var = tk.StringVar(master=self.frame)
        self.multiplier_entry = ttk.Entry(
            multiplier_box,
            textvariable=self.multiplier_var,
            width=16,
            state="disabled",
        )
        self.multiplier_entry.pack(side=tk.LEFT, padx=(4, 0))
        self.multiplier_entry.bind("<FocusOut>", self._on_multiplier_changed)
        self.multiplier_entry.bind("<Return>", self._on_multiplier_submit)
        self.multiplier_entry.bind("<KP_Enter>", self._on_multiplier_submit)

        def _reload_multiplier_from_file(event: Optional[tk.Event] = None) -> None:  # type: ignore[override]
            try:
                if not self.model.csv_path.exists():
                    messagebox.showwarning(
                        "Reload Multiplier", "No base CSV file found."
                    )
                    return

                multiplier_txt = self.model.csv_path.with_name(
                    self.model.csv_path.stem + "_multiplier.txt"
                )
                if not multiplier_txt.exists():
                    messagebox.showwarning(
                        "Reload Multiplier",
                        f"{multiplier_txt.name} does not exist.",
                    )
                    return

                content = multiplier_txt.read_text(encoding="utf-8").strip()
                if content:
                    self._set_multiplier(content)
                    print(f"ℹ️ Reloaded multiplier from {multiplier_txt.name}")
                else:
                    messagebox.showinfo(
                        "Reload Multiplier", f"{multiplier_txt.name} is empty."
                    )
            except Exception as exc:
                import traceback

                traceback.print_exc()
                messagebox.showerror(
                    "Reload Multiplier", f"Failed to reload multiplier:\n{exc}"
                )

        self.multiplier_entry.bind("<Button-1>", _reload_multiplier_from_file, add="+")

        def _cycle_multiplier() -> None:
            cycle_values = ["1", "1000", "1000000", "1000000000"]
            current = self.multiplier_var.get().strip()
            try:
                idx = cycle_values.index(current)
                new_value = cycle_values[(idx + 1) % len(cycle_values)]
            except ValueError:
                new_value = cycle_values[0]

            self._set_multiplier(new_value)
            try:
                self.save_multiplier()
            except Exception as exc:
                import traceback

                traceback.print_exc()
                messagebox.showerror(
                    "Multiplier", f"Failed to save multiplier:\n{exc}"
                )

        cycle_btn = ttk.Button(multiplier_box, text="Cycle", command=_cycle_multiplier)
        cycle_btn.pack(side=tk.LEFT, padx=(8, 0))

        def _open_multiplier_txt() -> None:
            try:
                base_csv = self.model.csv_path
                if not base_csv.exists():
                    messagebox.showwarning(
                        "Open Multiplier", "No base CSV file found."
                    )
                    return

                multiplier_txt = base_csv.with_name(base_csv.stem + "_multiplier.txt")

                if not multiplier_txt.exists():
                    try:
                        multiplier_txt.write_text("1", encoding="utf-8")
                        messagebox.showinfo(
                            "Open Multiplier",
                            f"{multiplier_txt.name} did not exist and was created with default value 1.",
                        )
                    except Exception as exc:
                        messagebox.showerror(
                            "Open Multiplier",
                            f"Failed to create {multiplier_txt.name}:\n{exc}",
                        )
                        return

                self.app.open_file_path(multiplier_txt)
            except Exception as exc:
                messagebox.showerror(
                    "Open Multiplier", f"Failed to open multiplier.txt:\n{exc}"
                )

        open_multiplier_btn = ttk.Button(
            multiplier_box, text="Open _multiplier.txt", command=_open_multiplier_txt
        )
        open_multiplier_btn.pack(side=tk.LEFT, padx=(8, 0))

        actions_row = ttk.Frame(self.frame)
        actions_row.pack(fill=tk.X, pady=(6, 0))
        self.open_csv_button = ttk.Button(
            actions_row, text="Open CSV", command=self.open_csv
        )
        self.open_csv_button.pack(side=tk.LEFT)

        self.reload_csv_button = ttk.Button(
            actions_row, text="Reload CSV", command=self.load_from_files
        )
        self.reload_csv_button.pack(side=tk.LEFT, padx=(6, 0))
        self.delete_csv_button = ttk.Button(
            actions_row, text="Delete CSV", command=self.delete_csv
        )
        self.delete_csv_button.pack(side=tk.LEFT, padx=(6, 0))

        def _browse_pdf_folder() -> None:
            try:
                pdf_folder = self.target_dir / "PDF_FOLDER"
                if not pdf_folder.exists():
                    messagebox.showwarning(
                        "Browse PDF Folder",
                        f"PDF_FOLDER not found:\n{pdf_folder}",
                    )
                    return
                os.startfile(pdf_folder)
            except Exception as exc:
                messagebox.showerror(
                    "Browse PDF Folder", f"Failed to open directory:\n{exc}"
                )

        self.browse_pdf_button = ttk.Button(
            actions_row, text="Browse PDF", command=_browse_pdf_folder
        )
        self.browse_pdf_button.pack(side=tk.LEFT, padx=(6, 0))

        self.view = ScrapeTableView(
            self, self.frame, self.app, self.model, auto_scale_tables
        )
        self.view.container.pack(fill=tk.BOTH, expand=True, pady=(6, 0))
        self.context_menu = ScrapeContextMenu(self.view, self.model, self.app)
        self.table = self.view.table
        self.set_row_height = self.view.set_row_height  # type: ignore[assignment]

        self.set_placeholder("-")
        self._update_action_states()

    def destroy(self) -> None:
        self.app.unregister_panel_rows(self)
        self.container.destroy()

    def set_auto_scale(self, enabled: bool) -> None:
        self.view.set_auto_scale(enabled)

    def set_placeholder(self, fill: str) -> None:
        self.model.has_csv_data = False
        self.view.set_placeholder(fill)
        self._update_action_states()

    def mark_loading(self) -> None:
        if self.model.has_csv_data:
            return
        self.view.mark_loading()
        self._update_action_states()

    def load_from_files(self) -> None:
        header, data_rows = self.model.load_csv_rows()
        if data_rows or header is not None:
            self.view.populate(data_rows, register=True, header=header)
        else:
            self.set_placeholder("-")

        multiplier_value = self.model.load_multiplier_value()
        self._set_multiplier(multiplier_value)
        self._update_action_states()
        self.app.refresh_combined_tab()

    def set_multiplier(self, value: str) -> None:
        self._set_multiplier(value)

    def _set_multiplier(self, value: str) -> None:
        self.model.begin_multiplier_update()
        self.multiplier_var.set(value)
        self.model.end_multiplier_update()

    def save_multiplier(self) -> None:
        if self.model.is_updating_multiplier():
            return
        value = self.multiplier_var.get().strip()
        self.model.save_multiplier_value(value)

    def set_active(self, active: bool) -> None:
        color = "#1E90FF" if active else "#c3c3c3"
        thickness = 2 if active else 1
        self.container.configure(
            highlightbackground=color,
            highlightcolor=color,
            highlightthickness=thickness,
        )

    def open_csv(self) -> None:
        if not self.model.csv_path.exists():
            messagebox.showinfo("Open CSV", "CSV file not available yet.")
            return
        self.app.open_file_path(self.model.csv_path)

    def delete_column(self, index: Optional[int] = None) -> None:
        column_count = len(self.view.current_columns)
        if column_count <= 1:
            messagebox.showinfo(
                "Delete Column", "Cannot delete the last remaining column."
            )
            return

        if index is None:
            options = "\n".join(
                f"{idx + 1}. {name}"
                for idx, name in enumerate(self.view.current_columns)
            )
            selection = simpledialog.askinteger(
                "Delete Column",
                f"Select the column number to delete:\n{options}",
                parent=self.frame,
                minvalue=1,
                maxvalue=column_count,
                initialvalue=column_count,
            )
            if selection is None:
                return
            index = selection - 1
        elif index < 0 or index >= column_count:
            return

        new_columns = (
            self.view.current_columns[:index]
            + self.view.current_columns[index + 1 :]
        )
        if not new_columns:
            messagebox.showinfo(
                "Delete Column", "Cannot delete the last remaining column."
            )
            return

        rows: List[List[str]] = []
        for item_id in self.view.table.get_children(""):
            values = list(self.view.table.item(item_id, "values"))
            if index < len(values):
                del values[index]
            rows.append(values)

        normalized_rows: List[List[str]] = []
        expected_len = len(new_columns)
        for values in rows:
            trimmed = list(values[:expected_len])
            if len(trimmed) < expected_len:
                trimmed.extend([""] * (expected_len - len(trimmed)))
            normalized_rows.append(trimmed)

        self.view.populate(
            normalized_rows, register=self.model.has_csv_data, header=new_columns
        )
        self.model.has_csv_data = bool(normalized_rows)
        self.save_table_to_csv()
        self._update_action_states()

    def _delete_selected_rows(self) -> None:
        try:
            selected = list(self.view.table.selection())
            if not selected:
                messagebox.showinfo("Delete Row", "No rows selected.")
                return

            confirm = messagebox.askyesno(
                "Delete Row",
                f"Delete {len(selected)} selected row(s)?\nThis cannot be undone.",
                parent=self.frame,
            )
            if not confirm:
                return

            for item_id in selected:
                self.view.table.delete(item_id)

            self.save_table_to_csv()
            self.load_from_files()
        except Exception as exc:
            messagebox.showerror("Delete Row", f"Error deleting rows:\n{exc}")

    def delete_csv(self) -> None:
        confirm = messagebox.askyesno(
            "Delete CSV",
            f"Delete CSV for {self.entry.path.name} – {self.category}?\nThis cannot be undone.",
            parent=self.frame,
        )
        if not confirm:
            return
        try:
            base = self.model.csv_path
            prefix = base.stem
            parent = base.parent
            related_files = [
                parent / f"{prefix}.csv",
                parent / f"{prefix}_raw.csv",
                parent / f"{prefix}_multiplier.txt",
            ]

            deleted: List[str] = []
            for file_path in related_files:
                if file_path.exists():
                    try:
                        file_path.unlink()
                        deleted.append(file_path.name)
                    except Exception as exc:
                        messagebox.showwarning(
                            "Delete CSV", f"Failed to delete {file_path.name}: {exc}"
                        )

            if deleted:
                messagebox.showinfo(
                    "Delete CSV", "Deleted files:\n" + "\n".join(deleted)
                )
            else:
                messagebox.showinfo(
                    "Delete CSV", "No related files found to delete."
                )
        except Exception as exc:
            messagebox.showerror(
                "Delete CSV",
                f"Unexpected error while deleting CSV files:\n{exc}",
            )

        self.set_placeholder("-")
        self.model.has_csv_data = False
        self._update_action_states()
        self.app.refresh_combined_tab()

    def save_table_to_csv(self) -> None:
        rows = self.view.get_table_rows()
        self.model.save_table(self.view.current_columns, rows)

    def update_row_state(self, item_id: str, state: Optional[str]) -> None:
        self.view.update_row_state(item_id, state)
        self.save_table_to_csv()

    def update_note_coloring(self) -> None:
        self.view.update_note_coloring()

    def flash_row(self, category: str, subcategory: str, item: str) -> None:
        self.view.flash_row(category, subcategory, item)

    def _on_multiplier_changed(self, _: tk.Event) -> None:  # type: ignore[override]
        self.save_multiplier()

    def _on_multiplier_submit(self, _: tk.Event) -> str:  # type: ignore[override]
        self.save_multiplier()
        return "break"

    def _update_action_states(self) -> None:
        has_csv = self.model.csv_path.exists()
        self.open_csv_button.configure(state="normal" if has_csv else "disabled")
        self.delete_csv_button.configure(
            state="normal"
            if has_csv or self.model.has_csv_data
            else "disabled"
        )

    def _note_column_index(self) -> Optional[int]:
        return self.view._note_column_index()

    def _apply_note_color_to_item(self, item_id: str) -> None:
        self.view._apply_note_color_to_item(item_id)
