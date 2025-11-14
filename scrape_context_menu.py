"""Context menu controller for scrape results."""

from __future__ import annotations

import sys
from typing import List, Optional, TYPE_CHECKING

import tkinter as tk

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

        self.menu = tk.Menu(self.view.table, tearoff=False)
        self._build_menu()

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
        self.menu.add_separator()
        self.menu.add_command(
            label="Flip Sign (row)",
            command=self._flip_sign_row,
        )
        self.menu.add_separator()
        self.menu.add_command(
            label="Flip Sign",
            command=self._flip_sign,
        )

    def _on_table_right_click(self, event: tk.Event) -> Optional[str]:  # type: ignore[override]
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

    def _set_row_state(self, state_label: str, apply_all: bool) -> None:
        item_id = self._context_item
        if not item_id:
            return

        if apply_all:
            target_rows = [item_id]
        else:
            sel = list(self.view.table.selection())
            target_rows = sel if sel else [item_id]

        normalized = self.view.normalize_state_label(state_label)
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

        if not apply_all:
            self.view.panel.save_table_to_csv()
            return

        key = self.model.row_keys.get(item_id)
        if apply_all and key and not getattr(self, "_is_propagating", False):
            self._is_propagating = True
            try:
                if normalized is not None:
                    self.app.scrape_row_state_by_key[key] = normalized
                else:
                    self.app.scrape_row_state_by_key.pop(key, None)

                note_index = self.view._note_column_index()
                if note_index is not None:
                    for panel in self.app.scrape_panels.values():
                        if panel is self.view.panel:
                            continue
                        view = panel.view
                        for other_id, other_key in panel.model.row_keys.items():
                            if other_key == key:
                                vals = list(view.table.item(other_id, "values"))
                                if len(vals) <= note_index:
                                    vals.extend([""] * (note_index + 1 - len(vals)))
                                vals[note_index] = state_label
                                view.table.item(other_id, values=vals)
                                view._apply_note_color_to_item(other_id)
                                view.table.update_idletasks()
                                break
                        panel.save_table_to_csv()
            finally:
                self._is_propagating = False

        elif not apply_all:
            self.view.update_row_state(item_id, normalized)
            self.view.table.update_idletasks()

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
