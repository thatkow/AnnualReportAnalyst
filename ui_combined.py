from __future__ import annotations

import csv
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import tkinter as tk
from tkinter import messagebox, simpledialog, ttk

from constants import COLUMNS, SCRAPE_EXPECTED_COLUMNS
from ui_widgets import CollapsibleFrame
from pdf_utils import PDFEntry, normalize_header_row


class CombinedUIMixin:
    root: tk.Misc
    combined_tab: Optional[ttk.Frame]
    combined_date_tree: Optional[ttk.Treeview]
    combined_table: Optional[ttk.Treeview]
    combined_create_button: Optional[ttk.Button]
    combined_save_button: Optional[ttk.Button]
    combined_columns: List[str]
    combined_rows: List[List[str]]
    combined_dyn_columns: List[Dict[str, Any]]
    combined_rename_names: List[str]
    combined_date_all_col_ids: List[str]
    combined_table_col_ids: List[str]
    note_color_scheme: Dict[str, str]
    pdf_entries: List[PDFEntry]
    company_var: tk.StringVar
    companies_dir: Path
    assigned_pages: Dict[str, Dict[str, Any]]
    canvas_window: int

    def build_combined_tab(self, notebook: ttk.Notebook) -> None:
        combined_tab = ttk.Frame(notebook)
        notebook.add(combined_tab, text="Combined")
        self.combined_tab = combined_tab

        combined_top = CollapsibleFrame(combined_tab, "Date Columns by PDF", initially_open=True)
        combined_top.pack(fill=tk.BOTH, expand=False, padx=8, pady=(8, 0))

        date_frame = ttk.Frame(combined_top.content, padding=(8, 0, 8, 8))
        date_frame.pack(fill=tk.BOTH, expand=True)
        cols = ("Type",)
        self.combined_date_tree = ttk.Treeview(date_frame, columns=cols, show="headings", height=6)
        self.combined_date_tree.heading("Type", text="Type")
        self.combined_date_tree.column("Type", width=140, anchor=tk.W, stretch=False)
        ysc = ttk.Scrollbar(date_frame, orient=tk.VERTICAL, command=self.combined_date_tree.yview)
        xsc = ttk.Scrollbar(date_frame, orient=tk.HORIZONTAL, command=self.combined_date_tree.xview)
        self.combined_date_tree.configure(yscrollcommand=ysc.set, xscrollcommand=xsc.set)
        self.combined_date_tree.grid(row=0, column=0, sticky="nsew")
        ysc.grid(row=0, column=1, sticky="ns")
        xsc.grid(row=1, column=0, sticky="ew")
        date_frame.rowconfigure(0, weight=1)
        date_frame.columnconfigure(0, weight=1)
        self.combined_date_tree.bind("<Button-3>", self._on_combined_header_right_click)
        if sys.platform == "darwin":
            self.combined_date_tree.bind("<Control-Button-1>", self._on_combined_header_right_click)

        controls = ttk.Frame(combined_tab, padding=8)
        controls.pack(fill=tk.X)
        self.combined_create_button = ttk.Button(controls, text="Create", command=self.create_combined_dataset)
        self.combined_create_button.pack(side=tk.LEFT)
        self.combined_save_button = ttk.Button(controls, text="Save CSV", command=self.save_combined_to_csv, state="disabled")
        self.combined_save_button.pack(side=tk.LEFT, padx=(6, 0))

        table_container = ttk.Frame(combined_tab, padding=(8, 0, 8, 8))
        table_container.pack(fill=tk.BOTH, expand=True)
        table_container.rowconfigure(0, weight=1)
        table_container.columnconfigure(0, weight=1)
        self.combined_table = ttk.Treeview(table_container, columns=("init",), show="headings")
        self.combined_table.heading("init", text="Combined table not yet created. Click 'Create'.")
        self.combined_table.column("init", width=600, anchor=tk.W, stretch=True)
        self.combined_table.grid(row=0, column=0, sticky="nsew")
        ysb = ttk.Scrollbar(table_container, orient=tk.VERTICAL, command=self.combined_table.yview)
        xsb = ttk.Scrollbar(table_container, orient=tk.HORIZONTAL, command=self.combined_table.xview)
        self.combined_table.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
        ysb.grid(row=0, column=1, sticky="ns")
        xsb.grid(row=1, column=0, sticky="ew")
        self.combined_table.bind("<Button-3>", self._on_combined_header_right_click)
        if sys.platform == "darwin":
            self.combined_table.bind("<Control-Button-1>", self._on_combined_header_right_click)

    def _combined_scrape_dir_for_entry(self, entry: PDFEntry) -> Path:
        company = self.company_var.get().strip()
        if company:
            return self.companies_dir / company / "openapiscrape" / entry.path.stem
        return entry.path.parent / "openapiscrape" / entry.path.stem

    def _read_csv_path(self, path: Path) -> Tuple[List[str], List[List[str]]]:
        header: List[str] = []
        rows: List[List[str]] = []
        if not path.exists():
            return header, rows
        try:
            with path.open("r", encoding="utf-8", newline="") as fh:
                reader = csv.reader(fh)
                for raw in reader:
                    if not raw:
                        continue
                    if not header:
                        cand = normalize_header_row([c.strip() for c in raw])
                        header = cand or [c.strip() for c in raw]
                    else:
                        rows.append([c.strip() for c in raw])
        except Exception:
            return [], []
        return header, rows

    @staticmethod
    def _date_columns_from_header(header: List[str]) -> List[str]:
        ignore = {"category", "subcategory", "item", "note"}
        out: List[str] = []
        for name in header:
            if name.strip().lower() in ignore:
                continue
            out.append(name.strip())
        return out

    @staticmethod
    def _parse_date_key(val: str) -> Tuple[int, int, int]:
        s = val.strip()
        for sep in (".", "/"):
            parts = s.split(sep)
            if len(parts) == 3 and all(part.isdigit() for part in parts):
                d, m, y = parts
                try:
                    return int(y), int(m), int(d)
                except Exception:
                    continue
        try:
            dt = datetime.strptime(s, "%Y-%m-%d")
            return (dt.year, dt.month, dt.day)
        except Exception:
            pass
        return (0, 0, 0)

    def _build_date_matrix_data(self) -> Tuple[List[Dict[str, Any]], Dict[str, List[str]], List[str]]:
        dyn_columns: List[Dict[str, Any]] = []
        warnings: List[str] = []
        for entry in self.pdf_entries:
            base = self._combined_scrape_dir_for_entry(entry)
            per_type_dates: Dict[str, List[str]] = {}
            for typ in COLUMNS:
                header, _ = self._read_csv_path(base / f"{typ}.csv")
                per_type_dates[typ] = self._date_columns_from_header(header) if header else []
            lengths = {typ: len(per_type_dates.get(typ, [])) for typ in COLUMNS}
            if len(set(lengths.values())) > 1:
                warnings.append(
                    f"{entry.path.name}: Financial={lengths['Financial']}, "
                    f"Income={lengths['Income']}, Shares={lengths['Shares']}"
                )
            max_len = max(lengths.values()) if lengths else 0
            for idx in range(max_len):
                fin = per_type_dates["Financial"][idx] if idx < len(per_type_dates["Financial"]) else ""
                inc = per_type_dates["Income"][idx] if idx < len(per_type_dates["Income"]) else ""
                sh = per_type_dates["Shares"][idx] if idx < len(per_type_dates["Shares"]) else ""
                primary = fin or inc or sh or f"date{idx+1}"
                display = f"{entry.path.name}:{primary}"
                default_name = fin or primary
                dyn_columns.append(
                    {
                        "pdf": entry.path.name,
                        "entry": entry,
                        "index": idx,
                        "labels": {"Financial": fin, "Income": inc, "Shares": sh},
                        "display_label": display,
                        "default_name": default_name,
                    }
                )
        rows_by_type: Dict[str, List[str]] = {t: [] for t in COLUMNS}
        for col in dyn_columns:
            for t in COLUMNS:
                rows_by_type[t].append(col["labels"].get(t, ""))
        return dyn_columns, rows_by_type, warnings

    def _rebuild_rename_inputs(self, dyn_columns: List[Dict[str, Any]]) -> None:
        return

    def _union_pages_for_entry(self, entry: PDFEntry) -> List[int]:
        pages_set: set[int] = set()
        for cat in COLUMNS:
            for p in entry.selected_pages.get(cat, []):
                try:
                    pages_set.add(int(p))
                except Exception:
                    continue
        if not pages_set:
            rec = self.assigned_pages.get(entry.path.name, {})
            if isinstance(rec, dict):
                multi = rec.get("multi_selections", {})
                if isinstance(multi, dict):
                    for lst in multi.values():
                        if isinstance(lst, list):
                            for p in lst:
                                try:
                                    pages_set.add(int(p))
                                except Exception:
                                    continue
                sel = rec.get("selections", {})
                if isinstance(sel, dict):
                    for p in sel.values():
                        try:
                            pages_set.add(int(p))
                        except Exception:
                            continue
        return sorted(pages_set)

    def _pages_list_string_for_entry(self, entry: PDFEntry) -> str:
        pages = self._union_pages_for_entry(entry)
        return ";".join(str(p + 1) for p in pages)

    def _pages_list_string_for_entry_and_type(self, entry: PDFEntry, category: str) -> str:
        pages_set: set[int] = set()
        for p in entry.selected_pages.get(category, []):
            try:
                pages_set.add(int(p))
            except Exception:
                continue
        if not pages_set:
            rec = self.assigned_pages.get(entry.path.name, {})
            if isinstance(rec, dict):
                multi = rec.get("multi_selections", {})
                if isinstance(multi, dict):
                    lst = multi.get(category, [])
                    if isinstance(lst, list):
                        for p in lst:
                            try:
                                pages_set.add(int(p))
                            except Exception:
                                continue
                sel = rec.get("selections", {})
                if isinstance(sel, dict):
                    p = sel.get(category)
                    try:
                        if p is not None:
                            pages_set.add(int(p))
                    except Exception:
                        pass
        pages = sorted(pages_set)
        return ";".join(str(p + 1) for p in pages)

    def refresh_combined_tab(self) -> None:
        if self.combined_date_tree is None:
            return
        dyn_columns, rows_by_type, warnings = self._build_date_matrix_data()
        self.combined_dyn_columns = dyn_columns

        if len(self.combined_rename_names) != len(dyn_columns):
            self.combined_rename_names = []
            for idx, col in enumerate(dyn_columns):
                name = col.get("default_name") or f"date{int(col.get('index', 0))+1}"
                self.combined_rename_names.append(name)

        tv = self.combined_date_tree
        # Safely clear Treeview without removing all columns
        try:
            # Clear all items (rows) first
            for iid in tv.get_children(""):
                tv.delete(iid)

            # If columns exist, keep a placeholder until reconfigured
            if not tv["columns"]:
                tv["columns"] = ("placeholder",)
                tv.heading("placeholder", text="")
                tv.column("placeholder", width=1)
        except tk.TclError:
            # Fallback safety reset
            tv["columns"] = ("placeholder",)

        # Now reconfigure columns properly below

        columns = ["Type"] + [col.get("display_label", "") for col in dyn_columns]
        col_ids = [f"c{idx}" for idx in range(len(columns))]
        self.combined_date_all_col_ids = col_ids
        tv.configure(columns=col_ids, displaycolumns=col_ids, show="headings")
        for idx, cid in enumerate(col_ids):
            tv.heading(cid, text=columns[idx])
            tv.column(cid, width=160 if idx else 140, anchor=tk.W if idx == 0 else tk.E)
        for typ in COLUMNS:
            values = [typ] + rows_by_type.get(typ, [])
            tv.insert("", "end", values=values)

        if warnings:
            messagebox.showwarning("Combined", "\n".join(warnings))

    def _apply_note_color_to_combined_item(self, item_id: str, columns: List[str], tv: ttk.Treeview) -> None:
        try:
            note_idx = next(i for i, n in enumerate(columns) if n.strip().lower() == "note")
        except StopIteration:
            return
        values = list(tv.item(item_id, "values"))
        if note_idx >= len(values):
            return
        note_val = str(values[note_idx]).strip()
        color = self.get_note_color(note_val)
        existing_tags = list(tv.item(item_id, "tags"))
        filtered = [t for t in existing_tags if not t.startswith("note-color-")]
        if color:
            tag = f"note-color-{color.replace('#','')}"
            try:
                tv.tag_configure(tag, background=color)
            except Exception:
                pass
            filtered.append(tag)
        tv.item(item_id, tags=filtered)

    def _populate_combined_table(self, columns: List[str], rows: List[List[str]]) -> None:
        if self.combined_table is None:
            return
        tv = self.combined_table
        for iid in tv.get_children(""):
            tv.delete(iid)
        col_ids = [f"c{idx}" for idx in range(len(columns))]
        self.combined_table_col_ids = col_ids
        tv.configure(columns=col_ids, displaycolumns=col_ids, show="headings")
        for idx, cid in enumerate(col_ids):
            name = columns[idx]
            anchor = tk.W if name.lower() in {"category", "subcategory", "item", "note"} else tk.E
            tv.heading(cid, text=name)
            width = self.get_scrape_column_width(name)
            tv.column(cid, width=width, anchor=anchor, stretch=True)
        for row in rows:
            values = list(row[:len(col_ids)])
            if len(values) < len(col_ids):
                values += [""] * (len(col_ids) - len(values))
            iid = tv.insert("", "end", values=values)
            self._apply_note_color_to_combined_item(iid, columns, tv)

    def save_combined_to_csv(self) -> None:
        if not self.combined_columns or not self.combined_rows:
            messagebox.showinfo("Save Combined", "No combined data to save. Click 'Create' first.")
            return
        company = self.company_var.get().strip()
        target_dir: Optional[Path] = None
        if company:
            target_dir = self.companies_dir / company
        elif self.pdf_entries:
            target_dir = self.pdf_entries[0].path.parent
        else:
            messagebox.showinfo("Save Combined", "Select a company or load PDFs first.")
            return
        try:
            target_dir.mkdir(parents=True, exist_ok=True)
            out_path = target_dir / "Combined.csv"
            with out_path.open("w", encoding="utf-8", newline="") as fh:
                writer = csv.writer(fh, quoting=csv.QUOTE_MINIMAL)
                writer.writerow(self.combined_columns)
                writer.writerows(self.combined_rows)
            messagebox.showinfo("Save Combined", f"Saved combined CSV to:\n{out_path}")
        except Exception as exc:
            messagebox.showerror("Save Combined", f"Failed to save combined CSV: {exc}")

    def _on_combined_header_right_click(self, event: tk.Event) -> None:  # type: ignore[override]
        try:
            tv: ttk.Treeview = event.widget  # type: ignore[assignment]
        except Exception:
            return
        region = tv.identify_region(event.x, event.y)
        if region != "heading":
            return
        col = tv.identify_column(event.x)
        if not col or not col.startswith("#"):
            return
        try:
            idx = int(col[1:]) - 1
        except Exception:
            return
        if tv is self.combined_date_tree:
            if idx < 1:
                return
            dyn_idx = idx - 1
            if dyn_idx < 0 or dyn_idx >= len(self.combined_rename_names):
                return
            current = self.combined_rename_names[dyn_idx]
            new_name = simpledialog.askstring("Rename Column", "Enter new column name:", initialvalue=current, parent=self.root)
            if new_name is None:
                return
            new_name = new_name.strip()
            if not new_name:
                return
            self.combined_rename_names[dyn_idx] = new_name
            all_cols_ids = self.combined_date_all_col_ids or []
            if idx < len(all_cols_ids):
                tv.heading(all_cols_ids[idx], text=new_name)
            if self.combined_table is not None and self.combined_columns:
                c_idx = 4 + dyn_idx
                if 0 <= c_idx < len(self.combined_columns):
                    self.combined_columns[c_idx] = new_name
                    if self.combined_table_col_ids and c_idx < len(self.combined_table_col_ids):
                        try:
                            self.combined_table.heading(self.combined_table_col_ids[c_idx], text=new_name)
                        except Exception:
                            pass
        elif tv is self.combined_table:
            if not self.combined_columns or idx < 0 or idx >= len(self.combined_columns):
                return
            current = self.combined_columns[idx]
            new_name = simpledialog.askstring("Rename Column", "Enter new column name:", initialvalue=current, parent=self.root)
            if new_name is None:
                return
            new_name = new_name.strip()
            if not new_name:
                return
            self.combined_columns[idx] = new_name
            if self.combined_table_col_ids and idx < len(self.combined_table_col_ids):
                try:
                    tv.heading(self.combined_table_col_ids[idx], text=new_name)
                except Exception:
                    pass
            dyn_idx = idx - 4
            if 0 <= dyn_idx < len(self.combined_rename_names):
                self.combined_rename_names[dyn_idx] = new_name

    def create_combined_dataset(self) -> None:
        dyn_cols, rows_by_type, warnings = self._build_date_matrix_data()
        self.combined_dyn_columns = dyn_cols

        conflicts = []
        key_data: Dict[Tuple[str, str, str], Dict[str, Any]] = {}
        for entry in self.pdf_entries:
            base = self._combined_scrape_dir_for_entry(entry)
            for typ in COLUMNS:
                csv_path = base / f"{typ}.csv"
                header, rows = self._read_csv_path(csv_path)
                header = header or SCRAPE_EXPECTED_COLUMNS
                normalized_header = [col.strip() for col in header]
                for row in rows or [[]]:
                    padded = list(row[: len(normalized_header)])
                    if len(padded) < len(normalized_header):
                        padded.extend([""] * (len(normalized_header) - len(padded)))
                    mapping = dict(zip(normalized_header, padded))
                    category = mapping.get("CATEGORY", "")
                    subcategory = mapping.get("SUBCATEGORY", "")
                    item = mapping.get("ITEM", "")
                    note_val = mapping.get("NOTE", "")
                    key = (category, subcategory, item)
                    record = key_data.setdefault(key, {"NOTE": "", "values_by_dyn": {}})
                    if note_val:
                        existing_note = record.get("NOTE")
                        if existing_note and existing_note.strip().lower() != note_val.strip().lower():
                            conflicts.append((key, [existing_note, note_val]))
                        else:
                            record["NOTE"] = note_val
                    col_list = record.setdefault("values_by_dyn", {})
                    for dc in dyn_cols:
                        pdf_name = dc.get("pdf")
                        idx = dc.get("index")
                        if pdf_name != entry.path.name or idx is None:
                            continue
                        label = dc.get("labels", {}).get(typ, "")
                        value = mapping.get(label, "") if label else ""
                        if label:
                            col_list[(pdf_name, idx)] = value
        if conflicts:
            msg_lines = ["Conflicting NOTE values detected:"]
            for key, vals in conflicts[:20]:
                cat, sub, item = key
                msg_lines.append(f"- {cat} | {sub} | {item}: {', '.join(vals)}")
            if len(conflicts) > 20:
                msg_lines.append(f"... and {len(conflicts) - 20} more")
            messagebox.showerror("Combined – NOTE conflict", "\n".join(msg_lines))
            return

        columns = ["CATEGORY", "SUBCATEGORY", "ITEM", "NOTE"] + self.combined_rename_names
        pdf_summary = ["PDF source", "", "", ""] + [dc.get("pdf", "") for dc in dyn_cols]
        pages_summary = ["Pages", "", "", ""] + [
            self._pages_list_string_for_entry(dc.get("entry")) if isinstance(dc.get("entry"), PDFEntry) else ""
            for dc in dyn_cols
        ]

        def key_sort(k: Tuple[str, str, str]) -> Tuple[str, str, str]:
            return (k[0] or "", k[1] or "", k[2] or "")

        rows_out: List[List[str]] = []
        for key in sorted(key_data.keys(), key=key_sort):
            rec = key_data[key]
            cat, sub, item = key
            note_val = rec.get("NOTE", "")
            values_map_dyn: Dict[Tuple[str, int], str] = rec.get("values_by_dyn", {})
            values_for_row: List[str] = []
            for dc in dyn_cols:
                pdf = str(dc.get("pdf", ""))
                pos = int(dc.get("index", 0))
                values_for_row.append(values_map_dyn.get((pdf, pos), ""))
            rows_out.append([cat, sub, item, note_val] + values_for_row)

        final_rows = [pdf_summary, pages_summary] + rows_out
        self._populate_combined_table(columns, final_rows)
        self.combined_columns = columns
        self.combined_rows = final_rows
        if self.combined_save_button is not None:
            self.combined_save_button.configure(state="normal")
        messagebox.showinfo("Combined", f"Combined dataset created with {len(final_rows)} rows and {len(columns)} columns.")
