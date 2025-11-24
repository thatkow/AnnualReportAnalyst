"""Treeview presentation logic for scrape results."""

from __future__ import annotations

from typing import List, Optional, Tuple, TYPE_CHECKING

import tkinter as tk
from tkinter import ttk

from constants import SCRAPE_EXPECTED_COLUMNS, SCRAPE_PLACEHOLDER_ROWS
from scrape_table_model import ScrapeTableModel

if TYPE_CHECKING:  # pragma: no cover
    from report_app import ReportAppV2
    from scrape_panel import ScrapeResultPanel


class ScrapeTableView:
    """Manage the Treeview UI for scrape results."""

    def __init__(
        self,
        panel: "ScrapeResultPanel",
        parent: tk.Widget,
        app: "ReportAppV2",
        model: ScrapeTableModel,
        auto_scale_tables: bool,
    ) -> None:
        self.panel = panel
        self.app = app
        self.model = model

        self.auto_scale_tables = auto_scale_tables
        self._current_row_count = SCRAPE_PLACEHOLDER_ROWS
        self.current_columns: List[str] = list(SCRAPE_EXPECTED_COLUMNS)
        self._column_ids: List[str] = [f"col{idx}" for idx in range(len(self.current_columns))]

        self.container = ttk.Frame(parent)
        self.container.columnconfigure(0, weight=1)
        self.container.rowconfigure(0, weight=1)

        style = ttk.Style()
        style_name = "Scrape.Treeview"
        row_h = self.app.get_scrape_row_height()
        style.configure(style_name, rowheight=row_h)

        self.table = ttk.Treeview(
            self.container,
            columns=self._column_ids,
            show="headings",
            height=SCRAPE_PLACEHOLDER_ROWS,
            style=style_name,
        )
        self.table.grid(row=0, column=0, sticky="nsew")

        y_scrollbar = ttk.Scrollbar(self.container, orient=tk.VERTICAL, command=self.table.yview)
        y_scrollbar.grid(row=0, column=1, sticky="ns")
        x_scrollbar = ttk.Scrollbar(self.container, orient=tk.HORIZONTAL, command=self.table.xview)
        x_scrollbar.grid(row=1, column=0, sticky="ew")
        self.table.configure(yscrollcommand=y_scrollbar.set, xscrollcommand=x_scrollbar.set)

        self.table.bind("<MouseWheel>", self._on_table_mousewheel, add="+")
        self.table.bind("<Shift-MouseWheel>", self._on_table_shift_mousewheel, add="+")
        self.table.bind("<Button-4>", self._on_table_linux_scroll, add="+")
        self.table.bind("<Button-5>", self._on_table_linux_scroll, add="+")

        self.table.tag_configure("state-negated", background="#FFE6E6")
        self.table.tag_configure("state-excluded", background="#E8E8E8", foreground="#666666")
        self.table.tag_configure("state-share_count", background="#E0F3FF")
        self.table.tag_configure("state-goodwill", background="#FFF7E6")

        for widget in (
            self.panel.container,
            self.panel.frame,
            self.panel.title_label,
            self.container,
            self.table,
        ):
            widget.bind("<Button-1>", self._handle_activate, add="+")

        def set_row_height(val: int) -> None:
            r = int(val)
            if r < 10 or r > 60:
                raise ValueError(f"Invalid row height {r}")
            style.configure(style_name, rowheight=r)
            self.table.update_idletasks()

        self.set_row_height = set_row_height  # type: ignore[assignment]

        self._apply_table_columns(self.current_columns)

    # ------------------------------------------------------------------
    # Layout helpers
    # ------------------------------------------------------------------
    def _handle_activate(self, _: tk.Event) -> None:  # type: ignore[override]
        self.app._on_scrape_panel_clicked(self.panel)

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

    def set_placeholder(self, fill: str) -> None:
        column_count = len(self._column_ids)
        rows = [[fill for _ in range(column_count)] for _ in range(SCRAPE_PLACEHOLDER_ROWS)]
        self.populate(rows, register=False)
        self.model.has_csv_data = False

    def mark_loading(self) -> None:
        if self.model.has_csv_data:
            return
        column_count = len(self._column_ids)
        rows = [["?" for _ in range(column_count)] for _ in range(SCRAPE_PLACEHOLDER_ROWS)]
        self.populate(rows, register=False)

    def populate(
        self,
        rows: List[List[str]],
        register: bool,
        header: Optional[List[str]] = None,
    ) -> None:
        self.app.unregister_panel_rows(self.panel)
        self.model.clear_row_tracking()
        for item in self.table.get_children(""):
            self.table.delete(item)
        if header is not None:
            self._apply_table_columns(header)

        column_count = len(self._column_ids)
        for row in rows:
            raw_values = list(row[:column_count])
            if len(raw_values) < column_count:
                raw_values.extend([""] * (column_count - len(raw_values)))

            formatted: List[str] = []
            for col_name, value in zip(self.current_columns, raw_values):
                s = str(value).strip()
                if col_name.lower() in ("category", "subcategory", "item", "note"):
                    formatted.append(s)
                    continue
                if s == "":
                    formatted.append("")
                    continue
                try:
                    num = float(s)
                    if num.is_integer():
                        formatted.append(f"{int(num):,}")
                    else:
                        formatted.append(f"{num:,.2f}")
                except Exception:
                    formatted.append(s)

            item_id = self.table.insert("", "end", values=formatted)

            if register:
                key = (
                    raw_values[0] if column_count > 0 else "",
                    raw_values[1] if column_count > 1 else "",
                    raw_values[2] if column_count > 2 else "",
                    raw_values[3] if column_count > 3 else "",
                    raw_values[4] if column_count > 4 else "",
                )
                self.model.register_row_key(item_id, key)
                self.app.register_scrape_row(self.panel, item_id, key)
                initial_state = self.app.scrape_row_state_by_key.get(key)
                if initial_state:
                    self.update_row_state(item_id, initial_state)
            self._apply_note_color_to_item(item_id)

        self._current_row_count = max(len(rows), 1)
        self._update_table_height()

    # ------------------------------------------------------------------
    # Row state helpers
    # ------------------------------------------------------------------
    def update_row_state(self, item_id: str, state: Optional[str]) -> None:
        self.model.set_row_state(item_id, state)
        if state in (None, ""):
            self._apply_state_to_item(item_id, None)
        else:
            self._apply_state_to_item(item_id, state)
        self._apply_note_color_to_item(item_id)
        self.table.update_idletasks()

    @staticmethod
    def normalize_state_label(state_label: str) -> Optional[str]:
        return None if state_label == "asis" else state_label

    def _apply_state_to_item(self, item_id: str, state: Optional[str]) -> None:
        existing_tags = list(self.table.item(item_id, "tags"))
        note_tags = [tag for tag in existing_tags if tag.startswith("note-color-")]
        filtered_tags = [
            tag
            for tag in existing_tags
            if not tag.startswith("state-") and not tag.startswith("note-color-")
        ]
        if state:
            filtered_tags.append(f"state-{state}")
        note_tag = None
        if not note_tags:
            self._apply_note_color_to_item(item_id)
            note_tags = [
                tag for tag in self.table.item(item_id, "tags") if tag.startswith("note-color-")
            ]
        if note_tags:
            note_tag = note_tags[-1]
        if note_tag:
            filtered_tags.append(note_tag)
        self.table.item(item_id, tags=filtered_tags)

    # ------------------------------------------------------------------
    # Mouse/scroll helpers
    # ------------------------------------------------------------------
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

    # ------------------------------------------------------------------
    # Data extraction helpers
    # ------------------------------------------------------------------
    def get_table_rows(self) -> List[List[str]]:
        rows: List[List[str]] = []
        for item_id in self.table.get_children(""):
            rows.append(list(self.table.item(item_id, "values")))
        return rows

    # ------------------------------------------------------------------
    # Note coloring helpers
    # ------------------------------------------------------------------
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

    # ------------------------------------------------------------------
    # Misc helpers
    # ------------------------------------------------------------------
    def flash_row(self, category: str, subcategory: str, item: str) -> None:
        try:
            target_key = (category, subcategory, item)
            target_id = None
            for row_id, key in self.model.row_keys.items():
                if key[:3] == target_key:
                    target_id = row_id
                    break
            if not target_id:
                print(f"⚠️ flash_row: No row found for {target_key}")
                return
            tv = self.table
            try:
                tv.selection_set(target_id)
                tv.focus(target_id)
                tv.see(target_id)
            except Exception:
                pass

            flash_tag = "flash-highlight"
            try:
                tv.tag_configure(flash_tag, background="#fff3a3")
            except Exception:
                pass

            tv.item(target_id, tags=(flash_tag,))

            def _remove_flash() -> None:
                try:
                    tv.item(target_id, tags=())
                except Exception:
                    pass

            tv.after(600, _remove_flash)
        except Exception as exc:
            print(f"⚠️ flash_row failed: {exc}")
