"""Scrape result panel widget and helpers."""

from __future__ import annotations

import csv
from pathlib import Path
from typing import Dict, List, Optional, Tuple, TYPE_CHECKING
import sys

import tkinter as tk
from tkinter import messagebox, simpledialog, ttk

from app_logging import get_logger
from constants import SCRAPE_EXPECTED_COLUMNS, SCRAPE_PLACEHOLDER_ROWS
from pdf_utils import PDFEntry, normalize_header_row

if TYPE_CHECKING:  # pragma: no cover
    from report_app import ReportAppV2

logger = get_logger()

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
        self.csv_path = target_dir / f"{category}.csv"
        self.multiplier_path = target_dir / f"{category}_multiplier.txt"
        self.raw_path = target_dir / f"{category}_raw.txt"
        self.has_csv_data = False
        self._updating_multiplier = False

        self.row_states: Dict[str, str] = {}
        self.row_keys: Dict[str, Tuple[str, str, str]] = {}
        self._context_item: Optional[str] = None

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
        self.title_label = ttk.Label(header, text=title_text, font=("TkDefaultFont", 10, "bold"))
        self.title_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        multiplier_box = ttk.Frame(header)
        multiplier_box.pack(side=tk.RIGHT)
        ttk.Label(multiplier_box, text="Multiplier:").pack(side=tk.LEFT)
        self.multiplier_var = tk.StringVar(master=self.frame)
        # Disable manual typing — controlled via cycle button
        self.multiplier_entry = ttk.Entry(multiplier_box, textvariable=self.multiplier_var,
                                          width=16, state="disabled")
        self.multiplier_entry.pack(side=tk.LEFT, padx=(4, 0))
        self.multiplier_entry.bind("<FocusOut>", self._on_multiplier_changed)
        self.multiplier_entry.bind("<Return>", self._on_multiplier_submit)
        self.multiplier_entry.bind("<KP_Enter>", self._on_multiplier_submit)

        # === Reload multiplier.txt when input box clicked ===
        def _reload_multiplier_from_file(event=None):
            try:
                if not self.csv_path.exists():
                    messagebox.showwarning("Reload Multiplier", "No base CSV file found.")
                    return

                multiplier_txt = self.csv_path.with_name(self.csv_path.stem + "_multiplier.txt")
                if not multiplier_txt.exists():
                    messagebox.showwarning("Reload Multiplier", f"{multiplier_txt.name} does not exist.")
                    return

                content = multiplier_txt.read_text(encoding="utf-8").strip()
                if content:
                    self._updating_multiplier = True
                    self.multiplier_var.set(content)
                    self._updating_multiplier = False
                    # Removed popup confirmation when clicking multiplier input
                    print(f"ℹ️ Reloaded multiplier from {multiplier_txt.name}")
                else:
                    messagebox.showinfo("Reload Multiplier", f"{multiplier_txt.name} is empty.")
            except Exception as e:
                import traceback
                traceback.print_exc()
                messagebox.showerror("Reload Multiplier", f"Failed to reload multiplier:\n{e}")
        self.multiplier_entry.bind("<Button-1>", _reload_multiplier_from_file, add="+")

        # === Multiplier cycle button ===
        def _cycle_multiplier():
            cycle_values = ["1", "1000", "1000000", "1000000000"]
            current = self.multiplier_var.get().strip()
            try:
                idx = cycle_values.index(current)
                new_value = cycle_values[(idx + 1) % len(cycle_values)]
            except ValueError:
                # If invalid, reset to 1
                new_value = cycle_values[0]

            # Update display (entry is disabled so use .set)
            self._updating_multiplier = True
            self.multiplier_var.set(new_value)
            self._updating_multiplier = False

            # Persist to file
            try:
                self.save_multiplier()
            except Exception as e:
                import traceback
                traceback.print_exc()
                messagebox.showerror(
                    "Multiplier",
                    f"Failed to save multiplier:\n{e}"
                )

        cycle_btn = ttk.Button(multiplier_box, text="Cycle", command=_cycle_multiplier)
        cycle_btn.pack(side=tk.LEFT, padx=(8, 0))

        # === Button to open the related _multiplier.txt file ===
        def _open_multiplier_txt():
            try:
                base_csv = self.csv_path
                if not base_csv.exists():
                    messagebox.showwarning("Open Multiplier", "No base CSV file found.")
                    return

                multiplier_txt = base_csv.with_name(base_csv.stem + "_multiplier.txt")

                # ✅ Create file with default value 1 if missing
                if not multiplier_txt.exists():
                    try:
                        multiplier_txt.write_text("1", encoding="utf-8")
                        messagebox.showinfo(
                            "Open Multiplier",
                            f"{multiplier_txt.name} did not exist and was created with default value 1."
                        )
                    except Exception as e:
                        messagebox.showerror("Open Multiplier", f"Failed to create {multiplier_txt.name}:\n{e}")
                        return

                # Open the file (existing or newly created)
                self.app.open_file_path(multiplier_txt)
            except Exception as e:
                messagebox.showerror("Open Multiplier", f"Failed to open multiplier.txt:\n{e}")

        open_multiplier_btn = ttk.Button(multiplier_box, text="Open _multiplier.txt", command=_open_multiplier_txt)
        open_multiplier_btn.pack(side=tk.LEFT, padx=(8, 0))

        actions_row = ttk.Frame(self.frame)
        actions_row.pack(fill=tk.X, pady=(6, 0))
        self.open_csv_button = ttk.Button(actions_row, text="Open CSV", command=self.open_csv)
        self.open_csv_button.pack(side=tk.LEFT)
        self.view_raw_button = ttk.Button(actions_row, text="View Raw", command=self.view_raw_text)
        self.view_raw_button.pack(side=tk.LEFT, padx=(6, 0))
        self.delete_column_button = ttk.Button(
            actions_row,
            text="Delete Column",
            command=self.delete_column,
        )
        self.delete_column_button.pack(side=tk.LEFT, padx=(6, 0))

        # === Reload CSV button (similar creation to Delete Column) ===
        self.reload_csv_button = ttk.Button(
            actions_row,
            text="Reload CSV",
            command=self.load_from_files,
        )
        self.reload_csv_button.pack(side=tk.LEFT, padx=(6, 0))
        self.delete_csv_button = ttk.Button(
            actions_row,
            text="Delete CSV",
            command=self.delete_csv,
        )
        self.delete_csv_button.pack(side=tk.LEFT, padx=(6, 0))

        table_container = ttk.Frame(self.frame)
        
        table_container.columnconfigure(0, weight=1)
        table_container.rowconfigure(0, weight=1)

        self.auto_scale_tables = auto_scale_tables
        self._current_row_count = SCRAPE_PLACEHOLDER_ROWS
        self.current_columns: List[str] = list(SCRAPE_EXPECTED_COLUMNS)
        self._column_ids: List[str] = [f"col{idx}" for idx in range(len(self.current_columns))]
        table_container.pack(fill=tk.BOTH, expand=True, pady=(6, 0))
        # Apply configured row height using a dedicated style
        style = ttk.Style()
        style_name = "Scrape.Treeview"
        row_h = self.app.get_scrape_row_height()
        style.configure(style_name, rowheight=row_h)

        self.table = ttk.Treeview(
            table_container,
            columns=self._column_ids,
            show="headings",
            height=SCRAPE_PLACEHOLDER_ROWS,
            style=style_name,
        )

        # Method to update height dynamically when changed in the menu
        def set_row_height(val: int) -> None:
            r = int(val)
            if r < 10 or r > 60:
                raise ValueError(f"Invalid row height {r}")
            style.configure(style_name, rowheight=r)
            self.table.update_idletasks()

        # Expose so UI can call self.scrape_panels[key].set_row_height()
        self.set_row_height = set_row_height

        self._apply_table_columns(self.current_columns)
        self.table.grid(row=0, column=0, sticky="nsew")

        y_scrollbar = ttk.Scrollbar(table_container, orient=tk.VERTICAL, command=self.table.yview)
        y_scrollbar.grid(row=0, column=1, sticky="ns")
        x_scrollbar = ttk.Scrollbar(table_container, orient=tk.HORIZONTAL, command=self.table.xview)
        x_scrollbar.grid(row=1, column=0, sticky="ew")
        self.table.configure(yscrollcommand=y_scrollbar.set, xscrollcommand=x_scrollbar.set)

        self.table.bind("<MouseWheel>", self._on_table_mousewheel, add="+")
        self.table.bind("<Shift-MouseWheel>", self._on_table_shift_mousewheel, add="+")
        self.table.bind("<Button-4>", self._on_table_linux_scroll, add="+")
        self.table.bind("<Button-5>", self._on_table_linux_scroll, add="+")
        self.table.tag_configure("state-negated", background="#FFE6E6")
        self.table.tag_configure("state-excluded", background="#E8E8E8", foreground="#666666")
        self.table.tag_configure("state-share_count", background="#E0F3FF")

        self._row_state_var = tk.StringVar(master=self.frame, value="asis")
        self._context_menu = tk.Menu(self.table, tearoff=False)
        self._context_menu.add_radiobutton(
            label="As is",
            variable=self._row_state_var,
            value="asis",
            command=lambda: self._set_row_state("asis", apply_all=False),
        )
        self._context_menu.add_radiobutton(
            label="Negated",
            variable=self._row_state_var,
            value="negated",
            command=lambda: self._set_row_state("negated", apply_all=False),
        )
        self._context_menu.add_radiobutton(
            label="Excluded",
            variable=self._row_state_var,
            value="excluded",
            command=lambda: self._set_row_state("excluded", apply_all=False),
        )
        self._context_menu.add_radiobutton(
            label="Share count",
            variable=self._row_state_var,
            value="share_count",
            command=lambda: self._set_row_state("share_count", apply_all=False),
        )
        self._context_menu.add_separator()
        self._context_menu.add_command(
            label="As is (all)",
            command=lambda: self._set_row_state("asis", apply_all=True),
        )
        self._context_menu.add_command(
            label="Negated (all)",
            command=lambda: self._set_row_state("negated", apply_all=True),
        )
        self._context_menu.add_command(
            label="Excluded (all)",
            command=lambda: self._set_row_state("excluded", apply_all=True),
        )
        self._context_menu.add_command(
            label="Share count (all)",
            command=lambda: self._set_row_state("share_count", apply_all=True),
        )

        # === New Section Separator ===
        self._context_menu.add_separator()

        # === New: Flip Sign for selected row ===
        def _flip_sign():
            row_id = self._context_item
            if not row_id:
                return

            values = list(self.table.item(row_id, "values"))
            new_vals = []
            for v in values:
                try:
                    new_vals.append(str(-1 * float(v)))
                except Exception:
                    new_vals.append(v)

            self.table.item(row_id, values=new_vals)
            self.save_table_to_csv()
            self.load_from_files()

        self._context_menu.add_command(
            label="Flip Sign (row)",
            command=_flip_sign,
        )

        # === New Section Separator ===
        self._context_menu.add_separator()

        # === New: Flip Sign option ===
        def _flip_sign():
            # Multiply all numeric date-column values by -1.0
            rows = []
            for item_id in self.table.get_children(""):
                values = list(self.table.item(item_id, "values"))
                new_vals = []
                for v in values:
                    try:
                        new_vals.append(str(-1 * float(v)))
                    except Exception:
                        new_vals.append(v)
                self.table.item(item_id, values=new_vals)
            self.save_table_to_csv()
            self.app.reload_scrape_panels()

        self._context_menu.add_command(
            label="Flip Sign",
            command=_flip_sign,
        )
        self.table.bind("<Button-3>", self._on_table_right_click)
        if sys.platform == "darwin":
            # macOS sends Control-Button-1 for context menus
            self.table.bind("<Control-Button-1>", self._on_table_right_click)

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
        self.app.unregister_panel_rows(self)
        self.container.destroy()

    def _apply_table_columns(self, columns: List[str]) -> None:
        headings = []
        for idx, column in enumerate(columns):
            text = column.strip()
            headings.append(text if text else f"Column {idx + 1}")

        if not headings:
            headings = list(SCRAPE_EXPECTED_COLUMNS)

        self.current_columns = headings
        self._column_ids = [f"col{idx}" for idx in range(len(headings))]
        self.table.configure(columns=self._column_ids, displaycolumns=self._column_ids)

        text_columns = {"category", "subcategory", "item", "note"}
        for idx, column_id in enumerate(self._column_ids):
            heading_text = headings[idx]
            anchor = tk.W if heading_text.lower() in text_columns else tk.E
            self.table.heading(column_id, text=heading_text)
            width = self.app.get_scrape_column_width(heading_text)
            self.table.column(column_id, anchor=anchor, width=width, stretch=True)

    def set_auto_scale(self, enabled: bool) -> None:
        if self.auto_scale_tables == enabled:
            return
        self.auto_scale_tables = enabled
        self._update_table_height()

    def _update_table_height(self) -> None:
        if self.auto_scale_tables:
            height = max(self._current_row_count, 1)
        else:
            height = SCRAPE_PLACEHOLDER_ROWS
        self.table.configure(height=height)

    def set_placeholder(self, fill: str) -> None:
        column_count = len(self._column_ids)
        rows = [[fill for _ in range(column_count)] for _ in range(SCRAPE_PLACEHOLDER_ROWS)]
        self._populate(rows, register=False)
        self.has_csv_data = False
        self._update_action_states()

    def mark_loading(self) -> None:
        if self.has_csv_data:
            return
        column_count = len(self._column_ids)
        rows = [["?" for _ in range(column_count)] for _ in range(SCRAPE_PLACEHOLDER_ROWS)]
        self._populate(rows, register=False)
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

        header: Optional[List[str]] = None
        data_rows: List[List[str]] = []
        if rows:
            candidate = normalize_header_row(rows[0])
            if candidate is not None:
                header = candidate
                data_rows = rows[1:]
            else:
                data_rows = rows

        if data_rows or header is not None:
            self._populate(data_rows, register=True, header=header)
            self.has_csv_data = bool(data_rows)
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

    def delete_column(self) -> None:
        column_count = len(self.current_columns)
        if column_count <= 1:
            messagebox.showinfo("Delete Column", "Cannot delete the last remaining column.")
            return

        options = "\n".join(
            f"{idx + 1}. {name}" for idx, name in enumerate(self.current_columns)
        )
        default_index = column_count  # default to last column
        selection = simpledialog.askinteger(
            "Delete Column",
            f"Select the column number to delete:\n{options}",
            parent=self.frame,
            minvalue=1,
            maxvalue=column_count,
            initialvalue=default_index,
        )
        if selection is None:
            return

        index = selection - 1
        new_columns = self.current_columns[:index] + self.current_columns[index + 1 :]
        if not new_columns:
            messagebox.showinfo("Delete Column", "Cannot delete the last remaining column.")
            return

        rows: List[List[str]] = []
        for item_id in self.table.get_children(""):
            values = list(self.table.item(item_id, "values"))
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
        rows = normalized_rows

        # Re-populate and immediately write out the updated table
        self._populate(rows, register=self.has_csv_data, header=new_columns)
        self.has_csv_data = bool(rows)
        self.save_table_to_csv()
        self._update_action_states()

    def delete_csv(self) -> None:
        # Remove CSV file and reset table to placeholder "missing" state
        confirm = messagebox.askyesno(
            "Delete CSV",
            f"Delete CSV for {self.entry.path.name} – {self.category}?\nThis cannot be undone.",
            parent=self.frame,
        )
        if not confirm:
            return
        try:
            base = self.csv_path
            prefix = base.stem
            parent = base.parent

            # Delete all related files for this category
            related_files = [
                parent / f"{prefix}.csv",
                parent / f"{prefix}_raw.csv",
                parent / f"{prefix}_multiplier.txt",
            ]

            deleted = []
            for f in related_files:
                if f.exists():
                    try:
                        f.unlink()
                        deleted.append(f.name)
                    except Exception as e:
                        messagebox.showwarning("Delete CSV", f"Failed to delete {f.name}: {e}")

            if deleted:
                messagebox.showinfo("Delete CSV", "Deleted files:\n" + "\n".join(deleted))
            else:
                messagebox.showinfo("Delete CSV", "No related files found to delete.")

        except Exception as e:
            messagebox.showerror("Delete CSV", f"Unexpected error while deleting CSV files:\n{e}")

        # Reset table to empty/missing status
        self.set_placeholder("-")
        self.has_csv_data = False
        self._update_action_states()
        # Update Combined tab (date columns)
        self.app.refresh_combined_tab()

    def save_table_to_csv(self) -> None:
        try:
            self.target_dir.mkdir(parents=True, exist_ok=True)
        except OSError:
            logger.exception(
                "Unable to ensure scrape directory exists: %s", self.target_dir
            )
            return
        try:
            with self.csv_path.open("w", encoding="utf-8", newline="") as fh:
                writer = csv.writer(fh, quoting=csv.QUOTE_MINIMAL)
                writer.writerow(self.current_columns)
                for item_id in self.table.get_children(""):
                    values = list(self.table.item(item_id, "values"))
                    expected_len = len(self.current_columns)
                    if len(values) < expected_len:
                        values.extend([""] * (expected_len - len(values)))
                    else:
                        values = values[:expected_len]
                    writer.writerow(values)
        except OSError:
            logger.exception(
                "Unable to persist CSV after table modification for %s - %s",
                self.entry.path.name,
                self.category,
            )

    def _populate(
        self,
        rows: List[List[str]],
        register: bool,
        header: Optional[List[str]] = None,
    ) -> None:
        self.app.unregister_panel_rows(self)
        self.row_states.clear()
        self.row_keys.clear()
        for item in self.table.get_children(""):
            self.table.delete(item)
        if header is not None:
            self._apply_table_columns(header)

        column_count = len(self._column_ids)
        for row in rows:
            values = list(row[:column_count])
            if len(values) < column_count:
                values.extend([""] * (column_count - len(values)))
            item_id = self.table.insert("", "end", values=values)
            if register:
                key = (
                    values[0] if column_count > 0 else "",
                    values[1] if column_count > 1 else "",
                    values[2] if column_count > 2 else "",
                )
                self.row_keys[item_id] = key
                self.app.register_scrape_row(self, item_id, key)
                initial_state = self.app.scrape_row_state_by_key.get(key)
                if initial_state:
                    self.update_row_state(item_id, initial_state)
            # Apply note-based coloring per row
            self._apply_note_color_to_item(item_id)
        self._current_row_count = max(len(rows), 1)
        self._update_table_height()

    def _handle_activate(self, _: tk.Event) -> None:  # type: ignore[override]
        self.app._on_scrape_panel_clicked(self)

    def _on_table_right_click(self, event: tk.Event) -> None:  # type: ignore[override]
        row_id = self.table.identify_row(event.y)
        if not row_id:
            return
        # Ensure both selection and focus are set to the clicked row
        try:
            self.table.selection_set(row_id)
            self.table.focus(row_id)
        except Exception:
            self.table.selection_set(row_id)
        self._context_item = row_id
        state_value = self.row_states.get(row_id, "asis")
        if not state_value:
            state_value = "asis"
        self._row_state_var.set(state_value)
        try:
            self._context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self._context_menu.grab_release()
        return "break"

    def _set_row_state(self, state_label: str, apply_all: bool) -> None:
        item_id = self._context_item
        if not item_id:
            return

        normalized = self._normalize_state_label(state_label)

        # --- Update NOTE column immediately on current table ---
        note_index = self._note_column_index()
        if note_index is not None:
            values = list(self.table.item(item_id, "values"))
            if len(values) <= note_index:
                values.extend([""] * (note_index + 1 - len(values)))
            values[note_index] = state_label
            self.table.item(item_id, values=values)
            self._apply_note_color_to_item(item_id)
            self.table.update_idletasks()

        key = self.row_keys.get(item_id)

        # --- Apply to all matching rows across panels (non-recursive) ---
        if apply_all and key and not getattr(self, "_is_propagating", False):
            self._is_propagating = True
            try:
                # Just update logical state, no config reload
                self.app.scrape_row_state_by_key[key] = normalized

                for panel in self.app.scrape_panels.values():
                    # Skip this panel (already updated)
                    if panel is self:
                        continue

                    # Update all matching rows in that panel
                    for other_id, other_key in panel.row_keys.items():
                        if other_key == key:
                            # --- update NOTE field ---
                            vals = list(panel.table.item(other_id, "values"))
                            if len(vals) <= note_index:
                                vals.extend([""] * (note_index + 1 - len(vals)))
                            vals[note_index] = state_label
                            panel.table.item(other_id, values=vals)

                            # --- apply color from user-defined mapping ---
                            panel._apply_note_color_to_item(other_id)
                            panel.table.update_idletasks()
                            break

                    # persist that panel’s table
                    panel.save_table_to_csv()
            finally:
                self._is_propagating = False

        # --- Local single-row update ---
        elif not apply_all:
            self.update_row_state(item_id, normalized)
            self.table.update_idletasks()

        # --- Save, deselect, and clear context ---
        self.save_table_to_csv()
        self.table.selection_remove(self.table.selection())  # ✅ clear highlight
        self._context_item = None


    def update_row_state(self, item_id: str, state: Optional[str]) -> None:
        if state in (None, ""):
            if item_id in self.row_states:
                self.row_states.pop(item_id, None)
            self._apply_state_to_item(item_id, None)
            if self._context_item == item_id:
                self._row_state_var.set("asis")
        else:
            self.row_states[item_id] = state
            self._apply_state_to_item(item_id, state)
            if self._context_item == item_id:
                self._row_state_var.set(state)
        # Ensure note coloring remains in sync with any changes
        self._apply_note_color_to_item(item_id)

        # Force redraw — fixes “last row not recoloring” issue
        self.table.update_idletasks()

        # Immediately write out on modification
        self.save_table_to_csv()


    def _apply_state_to_item(self, item_id: str, state: Optional[str]) -> None:
        existing_tags = list(self.table.item(item_id, "tags"))
        # Preserve note-color tag(s), replace state-* tag
        note_tags = [tag for tag in existing_tags if tag.startswith("note-color-")]
        filtered_tags = [tag for tag in existing_tags if not tag.startswith("state-") and not tag.startswith("note-color-")]
        if state:
            filtered_tags.append(f"state-{state}")
        # Ensure note color tag is last to override background
        note_tag = None
        if not note_tags:
            # recompute and add note tag if any
            self._apply_note_color_to_item(item_id)
            note_tags = [tag for tag in self.table.item(item_id, "tags") if tag.startswith("note-color-")]
        if note_tags:
            note_tag = note_tags[-1]
        if note_tag:
            filtered_tags.append(note_tag)
        self.table.item(item_id, tags=filtered_tags)

    @staticmethod
    def _normalize_state_label(state_label: str) -> Optional[str]:
        return None if state_label == "asis" else state_label

    def _on_multiplier_changed(self, _: tk.Event) -> None:  # type: ignore[override]
        self.save_multiplier()

    def _on_multiplier_submit(self, _: tk.Event) -> str:  # type: ignore[override]
        self.save_multiplier()
        return "break"

    def _update_action_states(self) -> None:
        has_csv = self.csv_path.exists()
        self.open_csv_button.configure(state="normal" if has_csv else "disabled")
        self.delete_csv_button.configure(state="normal" if has_csv or self.has_csv_data else "disabled")
        has_raw = self.raw_path.exists()
        self.view_raw_button.configure(state="normal" if has_raw else "disabled")

    def _on_table_mousewheel(self, event: tk.Event) -> str:  # type: ignore[override]
        if event.delta == 0:
            return "break"
        direction = -1 if event.delta > 0 else 1
        self.table.yview_scroll(direction, "units")
        return "break"

    def _on_table_shift_mousewheel(self, event: tk.Event) -> str:  # type: ignore[override]
        if event.delta == 0:
            return "break"
        direction = -1 if event.delta > 0 else 1
        self.table.xview_scroll(direction, "units")
        return "break"

    def _on_table_linux_scroll(self, event: tk.Event) -> str:  # type: ignore[override]
        if getattr(event, "num", None) == 4:
            self.table.yview_scroll(-1, "units")
        elif getattr(event, "num", None) == 5:
            self.table.yview_scroll(1, "units")
        return "break"

    # Note coloring helpers
    def _note_column_index(self) -> Optional[int]:
        for idx, name in enumerate(self.current_columns):
            if name.strip().lower() == "note":
                return idx
        return None

    def _get_note_value_for_item(self, item_id: str) -> str:
        idx = self._note_column_index()
        if idx is None:
            return ""
        values = list(self.table.item(item_id, "values"))
        if idx < len(values):
            return str(values[idx]).strip()
        return ""

    def _apply_note_color_to_item(self, item_id: str) -> None:
        note_val = self._get_note_value_for_item(item_id)
        color = self.app.get_note_color(note_val)
        existing_tags = list(self.table.item(item_id, "tags"))
        # remove any previous note-color-*
        filtered = [t for t in existing_tags if not t.startswith("note-color-")]
        if color:
            tag = f"note-color-{color.replace('#','')}"
            try:
                self.table.tag_configure(tag, background=color)
            except Exception:
                pass
            filtered.append(tag)
        self.table.item(item_id, tags=filtered)

    def update_note_coloring(self) -> None:
        for item_id in self.table.get_children(""):
            self._apply_note_color_to_item(item_id)

    # ============================================================
    # NEW: Visually flash a row (CATEGORY, SUBCATEGORY, ITEM)
    # ============================================================
    def flash_row(self, category: str, subcategory: str, item: str) -> None:
        """
        Flash-highlight the row matching (category, subcategory, item).
        """
        try:
            target_key = (category, subcategory, item)

            # Locate matching item_id
            target_id = None
            for row_id, key in self.row_keys.items():
                if key == target_key:
                    target_id = row_id
                    break

            if not target_id:
                print(f"⚠️ flash_row: No row found for {target_key}")
                return

            tv = self.table

            # Apply selection so scrolling and focus work properly
            try:
                tv.selection_set(target_id)
                tv.focus(target_id)
                tv.see(target_id)
            except Exception:
                pass

            # Flash color tag
            flash_tag = "flash-highlight"
            try:
                tv.tag_configure(flash_tag, background="#fff3a3")  # pale yellow
            except Exception:
                pass

            # Add tag
            tv.item(target_id, tags=(flash_tag,))

            # Remove after delay
            def _remove_flash():
                try:
                    tv.item(target_id, tags=())
                except Exception:
                    pass

            tv.after(600, _remove_flash)
        except Exception as exc:
            print(f"⚠️ flash_row failed: {exc}")

