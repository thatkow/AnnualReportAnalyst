"""Context menu controller for scrape results."""

from __future__ import annotations

import sys
from typing import List, Optional, TYPE_CHECKING

import tkinter as tk
from tkinter import messagebox, simpledialog

from scrape_table_model import ScrapeTableModel
from scrape_table_view import ScrapeTableView
import re

# --- helper: detect DD.MM.YYYY date columns ---
_DATE_RE = re.compile(r"^\d{2}\.\d{2}\.\d{4}$")
def _is_date_col(name: str) -> bool:
    return bool(_DATE_RE.match(name.strip()))


if TYPE_CHECKING:  # pragma: no cover
    from report_app import ReportAppV2


class ScrapeContextMenu:
    """Handle right-click menu construction and actions."""

    def __init__(
        self,
        view: ScrapeTableView,
        model: ScrapeTableModel,
        app: "ReportAppV2",
    ) -> None:
        self.view = view
        self.model = model
        self.app = app

        self._row_state_var = tk.StringVar(master=self.view.container, value="asis")
        self._context_item: Optional[str] = None
        self._is_propagating = False
        self._header_context_column: Optional[int] = None

        self.menu = tk.Menu(self.view.table, tearoff=False)
        self._build_menu()

        self.header_menu = tk.Menu(self.view.table, tearoff=False)
        self._build_header_menu()

        self.view.table.bind("<Button-3>", self._on_table_right_click)
        if sys.platform == "darwin":
            self.view.table.bind("<Control-Button-1>", self._on_table_right_click)

    def _build_menu(self) -> None:
        self.menu.add_radiobutton(
            label="As is",
            variable=self._row_state_var,
            value="asis",
            command=lambda: self._set_row_state("asis", apply_all=False),
        )
        self.menu.add_radiobutton(
            label="Negated",
            variable=self._row_state_var,
            value="negated",
            command=lambda: self._set_row_state("negated", apply_all=False),
        )
        self.menu.add_radiobutton(
            label="Excluded",
            variable=self._row_state_var,
            value="excluded",
            command=lambda: self._set_row_state("excluded", apply_all=False),
        )
        self.menu.add_radiobutton(
            label="Share count",
            variable=self._row_state_var,
            value="share_count",
            command=lambda: self._set_row_state("share_count", apply_all=False),
        )
        self.menu.add_radiobutton(
            label="Intangibles",
            variable=self._row_state_var,
            value="intangibles",
            command=lambda: self._set_row_state("intangibles", apply_all=False),
        )
        self.menu.add_separator()
        self.menu.add_command(
            label="As is (all)",
            command=lambda: self._set_row_state("asis", apply_all=True),
        )
        self.menu.add_command(
            label="Negated (all)",
            command=lambda: self._set_row_state("negated", apply_all=True),
        )
        self.menu.add_command(
            label="Excluded (all)",
            command=lambda: self._set_row_state("excluded", apply_all=True),
        )
        self.menu.add_command(
            label="Share count (all)",
            command=lambda: self._set_row_state("share_count", apply_all=True),
        )
        self.menu.add_command(
            label="Intangibles (all)",
            command=lambda: self._set_row_state("intangibles", apply_all=True),
        )
        self.menu.add_separator()
        self.menu.add_command(
            label="Flip Sign (row)",
            command=self._flip_sign_row,
        )
        self.menu.add_command(
            label="Flip Sign (odd date cols)",
            command=self._flip_sign_odd_date_columns,
        )
        self.menu.add_separator()
        self.menu.add_command(
            label="Flip Sign",
            command=self._flip_sign,
        )
        self.menu.add_separator()
        self.menu.add_command(
            label="Delete selected row(s)",
            command=self._delete_rows,
        )
        self.menu.add_separator()
        self._build_subcategory_menu()

    def _build_subcategory_menu(self) -> None:
        sub_menu = tk.Menu(self.menu, tearoff=False)
        sub_menu.add_command(
            label="Custom...",
            command=self._prompt_custom_subcategory,
        )
        for label in ("CURRENT", "NON-CURRENT", "SECONDARYGRP"):
            sub_menu.add_command(
                label=label,
                command=lambda val=label: self._set_subcategory_value(val),
            )
        self.menu.add_cascade(label="SUBCATEGORY", menu=sub_menu)

    def _build_header_menu(self) -> None:
        self.header_menu.add_command(
            label="Delete column",
            command=self._delete_current_column,
        )
        self.header_menu.add_separator()
        self.header_menu.add_command(
            label="Sum other column into...",
            command=self._sum_other_column_into_current,
        )

    def _on_table_right_click(self, event: tk.Event) -> Optional[str]:  # type: ignore[override]
        region = self.view.table.identify_region(event.x, event.y)
        if region == "heading":
            return self._on_header_right_click(event)

        row_id = self.view.table.identify_row(event.y)
        if not row_id:
            return None

        current_selection = list(self.view.table.selection())
        if len(current_selection) > 1 and row_id in current_selection:
            pass
        else:
            self.view.table.selection_set(row_id)
            current_selection = [row_id]

        try:
            self.view.table.focus(row_id)
        except Exception:
            pass

        self._context_item = row_id
        state_value = self.model.get_row_state(row_id) or "asis"
        self._row_state_var.set(state_value)
        try:
            self.menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.menu.grab_release()
        return "break"

    def _get_target_rows(self) -> List[str]:
        selected = list(self.view.table.selection())
        if not selected and self._context_item:
            selected = [self._context_item]
        return selected

    def _prompt_custom_subcategory(self) -> None:
        value = simpledialog.askstring(
            "Set SUBCATEGORY",
            "Enter the SUBCATEGORY value to apply:",
            parent=self.view.panel.frame,
        )
        if value is None:
            return
        self._set_subcategory_value(value.strip())

    def _set_subcategory_value(self, value: str) -> None:
        target_rows = self._get_target_rows()
        if not target_rows:
            return

        sub_idx = self.view._subcategory_column_index()
        if sub_idx is None:
            messagebox.showinfo(
                "Set SUBCATEGORY",
                "No SUBCATEGORY column was found in this table.",
                parent=self.view.panel.frame,
            )
            return

        for item_id in target_rows:
            values = list(self.view.table.item(item_id, "values"))
            if len(values) <= sub_idx:
                values.extend([""] * (sub_idx + 1 - len(values)))
            values[sub_idx] = value
            self.view.table.item(item_id, values=values)

        self.view.panel.save_table_to_csv()
        self.view.panel.load_from_files()
        self.view.table.selection_remove(self.view.table.selection())
        self._context_item = None

    def _set_row_state(self, state_label: str, apply_all: bool) -> None:
        item_id = self._context_item
        if not item_id:
            return

        selected = list(self.view.table.selection())
        target_rows = selected if selected else [item_id]

        normalized = self.view.normalize_state_label(state_label)

        if apply_all:
            for rid in target_rows:
                key = self.model.row_keys.get(rid)
                if key:
                    self.app.apply_row_state_to_all(key, normalized)
            self.view.table.selection_remove(self.view.table.selection())
            self._context_item = None
            return

        note_index = self.view._note_column_index()
        if note_index is not None:
            for rid in target_rows:
                values = list(self.view.table.item(rid, "values"))
                if len(values) <= note_index:
                    values.extend([""] * (note_index + 1 - len(values)))
                values[note_index] = state_label
                self.view.table.item(rid, values=values)
                self.view._apply_note_color_to_item(rid)
            self.view.table.update_idletasks()

        for rid in target_rows:
            self.view.update_row_state(rid, normalized)

        self.view.panel.save_table_to_csv()
        self.view.table.selection_remove(self.view.table.selection())
        self._context_item = None

    def _flip_sign_row(self) -> None:
        selected = self.view.table.selection()
        if not selected:
            selected = [self._context_item] if self._context_item else []

        for item_id in selected:
            values = list(self.view.table.item(item_id, "values"))
            new_vals = values[:]

            for idx, col_name in enumerate(self.view.current_columns):
                if not _is_date_col(col_name):
                    continue
                s = str(new_vals[idx]).strip()
                if s.startswith("-"):
                    new_vals[idx] = s[1:]
                elif s == "":
                    new_vals[idx] = ""
                else:
                    new_vals[idx] = "-" + s

            self.view.table.item(item_id, values=new_vals)

        self.view.panel.save_table_to_csv()
        self.view.panel.load_from_files()

    def _flip_sign_odd_date_columns(self) -> None:
        selected = self.view.table.selection()
        if not selected:
            selected = [self._context_item] if self._context_item else []
        if not selected:
            return

        date_indices = [
            idx for idx, name in enumerate(self.view.current_columns)
            if _is_date_col(name)
        ]
        target_indices = [
            idx for pos, idx in enumerate(date_indices, start=1)
            if pos % 2 == 1
        ]
        if not target_indices:
            return

        for item_id in selected:
            values = list(self.view.table.item(item_id, "values"))
            new_vals = values[:]

            for idx in target_indices:
                if idx >= len(new_vals):
                    continue
                s = str(new_vals[idx]).strip()
                if s.startswith("-"):
                    new_vals[idx] = s[1:]
                elif s == "":
                    new_vals[idx] = ""
                else:
                    new_vals[idx] = "-" + s

            self.view.table.item(item_id, values=new_vals)

        self.view.panel.save_table_to_csv()
        self.view.panel.load_from_files()


    def _flip_sign(self) -> None:
        selected = self.view.table.selection()
        if not selected:
            return

        col = getattr(self.view.panel, "_flip_column_index", None)

        for item_id in selected:
            values = list(self.view.table.item(item_id, "values"))
            new_vals = values[:]

            # Determine which columns to flip (only date columns)
            if col is None:
                indices = [
                    i for i, name in enumerate(self.view.current_columns)
                    if _is_date_col(name)
                ]
            else:
                indices = [col] if (
                    col < len(values) and _is_date_col(self.view.current_columns[col])
                ) else []

            for idx in indices:
                s = str(new_vals[idx]).strip()
                if s.startswith("-"):
                    new_vals[idx] = s[1:]
                elif s == "":
                    new_vals[idx] = ""
                else:
                    new_vals[idx] = "-" + s

            self.view.table.item(item_id, values=new_vals)

        self.view.panel.save_table_to_csv()
        self.app.reload_scrape_panels()

    def _delete_rows(self) -> None:
        if not self.view.table.selection() and self._context_item:
            self.view.table.selection_set(self._context_item)
        self.view.panel._delete_selected_rows()

    # ------------------------------------------------------------------
    # Header menu helpers
    # ------------------------------------------------------------------
    def _on_header_right_click(self, event: tk.Event) -> str:
        column_id = self.view.table.identify_column(event.x)
        try:
            index = int(column_id.replace("#", "")) - 1
        except ValueError:
            return "break"

        if index < 0 or index >= len(self.view.current_columns):
            return "break"

        self._header_context_column = index
        try:
            self.header_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.header_menu.grab_release()
        return "break"

    @staticmethod
    def _parse_numeric(value: str) -> Optional[float]:
        cleaned = value.strip()
        if not cleaned:
            return None
        cleaned = cleaned.replace(",", "").replace("$", "")
        if cleaned.startswith("(") and cleaned.endswith(")"):
            cleaned = f"-{cleaned[1:-1]}"
        try:
            return float(cleaned)
        except ValueError:
            return None

    @staticmethod
    def _format_numeric(value: float) -> str:
        if value.is_integer():
            return f"{int(value):,}"
        return f"{value:,.2f}"

    def _sum_other_column_into_current(self) -> None:
        target_idx = self._header_context_column
        self._header_context_column = None
        if target_idx is None or target_idx >= len(self.view.current_columns):
            return

        target_name = self.view.current_columns[target_idx]
        if not _is_date_col(target_name):
            messagebox.showinfo(
                "Sum Columns",
                "The selected column must be a date column.",
                parent=self.view.panel.frame,
            )
            return

        date_columns = [
            (idx, name)
            for idx, name in enumerate(self.view.current_columns)
            if idx != target_idx and _is_date_col(name)
        ]

        if not date_columns:
            messagebox.showinfo(
                "Sum Columns",
                "No other date columns are available to sum.",
                parent=self.view.panel.frame,
            )
            return

        options = "\n".join(
            f"{pos + 1}. {name}" for pos, (_, name) in enumerate(date_columns)
        )
        selection = simpledialog.askinteger(
            "Sum Columns",
            f"Select the column to sum into {target_name}:\n{options}",
            parent=self.view.panel.frame,
            minvalue=1,
            maxvalue=len(date_columns),
            initialvalue=len(date_columns),
        )
        if selection is None:
            return

        source_idx = date_columns[selection - 1][0]

        rows = self.view.get_table_rows()
        updated_rows: List[List[str]] = []
        for values in rows:
            row = list(values)
            while len(row) < len(self.view.current_columns):
                row.append("")
            left = self._parse_numeric(row[target_idx]) or 0.0
            right = self._parse_numeric(row[source_idx]) or 0.0
            total = left + right
            if total == 0 and not row[target_idx] and not row[source_idx]:
                row[target_idx] = ""
            else:
                row[target_idx] = self._format_numeric(total)
            updated_rows.append(row)

        new_columns = (
            self.view.current_columns[:source_idx]
            + self.view.current_columns[source_idx + 1 :]
        )

        trimmed_rows: List[List[str]] = []
        expected_len = len(new_columns)
        for row in updated_rows:
            trimmed = list(row)
            if source_idx < len(trimmed):
                del trimmed[source_idx]
            if len(trimmed) < expected_len:
                trimmed.extend([""] * (expected_len - len(trimmed)))
            else:
                trimmed = trimmed[:expected_len]
            trimmed_rows.append(trimmed)

        self.view.populate(
            trimmed_rows,
            register=self.view.panel.model.has_csv_data,
            header=new_columns,
        )
        self.view.panel.model.has_csv_data = bool(trimmed_rows)
        self.view.panel.save_table_to_csv()

    def _delete_current_column(self) -> None:
        target_idx = self._header_context_column
        self._header_context_column = None
        if target_idx is None:
            return

        self.view.panel.delete_column(index=target_idx)
