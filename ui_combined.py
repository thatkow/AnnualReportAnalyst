from __future__ import annotations

import csv
import io
import json
import os
import subprocess
import sys
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Set, Tuple

import pandas as pd

import tkinter as tk
from tkinter import messagebox, simpledialog, ttk

from constants import COLUMNS, DEFAULT_OPENAI_MODEL, SCRAPE_EXPECTED_COLUMNS
# === New buttons: Copy ReleaseDate / Stock Multiplier Prompts ===
from combined_utils import (
    _sort_dates,
    build_release_date_prompt,
    build_stock_multiplier_prompt,
    generate_and_open_stock_multipliers,
)

COMBINED_BASE_COLUMNS = [
    "TYPE",
    "CATEGORY",
    "SUBCATEGORY",
    "ITEM",
    "NOTE",
    "Key4Coloring",
]
from ui_widgets import CollapsibleFrame
from pdf_utils import PDFEntry, normalize_header_row


class CombinedUIMixin:
    root: tk.Misc
    combined_tab: Optional[ttk.Frame]
    combined_date_tree: Optional[ttk.Treeview]
    combined_table: Optional[ttk.Treeview]
    combined_create_button: Optional[ttk.Button]
    mapping_create_button: Optional[ttk.Button]
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

    def _on_note_conflict_double_click(self, event: tk.Event, tree: ttk.Treeview) -> None:
        """Focus the Scrape TYPE/PDF tabs when a conflict row is double-clicked."""

        try:
            row_id = tree.identify_row(event.y)
            if not row_id:
                return
            values = tree.item(row_id, "values")
        except Exception:
            return

        if not values or len(values) < 6:
            return

        type_value = str(values[0]).strip()
        pdf_name = str(values[5]).strip()
        # Row identity for flashing
        category   = str(values[1]).strip()
        subcat     = str(values[2]).strip()
        item       = str(values[3]).strip()
        # TYPE is in type_value

        if not type_value or not pdf_name:
            return

        entry: Optional[PDFEntry] = None
        for candidate in getattr(self, "pdf_entries", []):
            if candidate.path.name == pdf_name:
                entry = candidate
                break
        if entry is None:
            return

        def _focus_scrape_panel() -> None:
            notebook = getattr(self, "notebook", None)
            scrape_tab = getattr(self, "scrape_tab", None)
            if notebook is not None and scrape_tab is not None:
                try:
                    notebook.select(scrape_tab)
                except Exception:
                    pass
            try:
                self.set_active_scrape_panel(entry, type_value)
            except Exception as exc:
                print(f"⚠️ Unable to focus Scrape panel for {type_value}/{pdf_name}: {exc}")

            # === NEW: flash the specific row in the active scrape panel ===
            try:
                active_key = getattr(self, "active_scrape_key", None)
                if active_key:
                    # Find the corresponding panel object
                    target_panel = None
                    for panel in self.scrape_panels.values():
                        p_entry = getattr(panel, "entry", None)
                        if p_entry and (p_entry.path, panel.category) == active_key:
                            target_panel = panel
                            break

                    # Flash-highlight the matching row
                    if target_panel and hasattr(target_panel, "flash_row"):
                        target_panel.flash_row(category, subcat, item)
                    else:
                        print(f"⚠️ flash_row: No panel found for {active_key}")
                else:
                    print("⚠️ flash_row: No active scrape key")
            except Exception as exc:
                print(f"⚠️ flash_row failed for row ({category}, {subcat}, {item}): {exc}")

        try:
            self.root.after(0, _focus_scrape_panel)
        except Exception:
            _focus_scrape_panel()


   
    def _get_pdf_table_dates(self) -> List[str]:
        """Return sorted date columns sourced from the Date Columns by PDF table."""

        dyn_cols = list(getattr(self, "combined_dyn_columns", []) or [])
        if not dyn_cols:
            dyn_cols, _, _ = self._build_date_matrix_data()
            self.combined_dyn_columns = dyn_cols

        rename_names = list(getattr(self, "combined_rename_names", []) or [])
        if len(rename_names) != len(dyn_cols):
            rename_names = [
                dc.get("default_name") or f"date{idx + 1}"
                for idx, dc in enumerate(dyn_cols)
            ]
            self.combined_rename_names = rename_names

        candidates = [name for name in rename_names if isinstance(name, str) and name.strip()]
        if not candidates:
            candidates = [
                dc.get("default_name", "")
                for dc in dyn_cols
                if isinstance(dc, dict)
            ]

        return _sort_dates([c for c in candidates if isinstance(c, str) and c.strip()])

    def _on_copy_releasedate_prompt(self):
        try:
            company = self.company_var.get().strip()
            if not company:
                messagebox.showwarning("Company Missing", "Select a company before generating the prompt.")
                return

            # Need the combined dataset first
            if not self.combined_columns or not self.combined_rows:
                messagebox.showwarning("No Combined Data", "Create the combined dataset first.")
                return

            date_cols = self._get_pdf_table_dates()

            # Build prompt text
            text = build_release_date_prompt(company, date_cols)

            # Copy to clipboard
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            self.root.update()

            # ------------------------------------------------------------
            # NEW BEHAVIOUR:
            # If ReleaseDates.csv exists, append missing dates and open it.
            # If not, create a new one.
            # ------------------------------------------------------------
            import csv, os, sys
            company_folder = self.companies_dir / company
            company_folder.mkdir(parents=True, exist_ok=True)
            csv_path = company_folder / "ReleaseDates.csv"

            # Case 1: File exists → merge
            if csv_path.exists():
                existing = {}
                with csv_path.open("r", encoding="utf-8", newline="") as fh:
                    reader = csv.DictReader(fh)
                    for row in reader:
                        date_key = row.get("Date", "").strip()
                        if date_key:
                            existing[date_key] = row.get("ReleaseDate", "").strip()

                # Append missing dates
                changed = False
                for d in date_cols:
                    if d not in existing:
                        existing[d] = ""
                        changed = True

                # Rewrite ONLY if needed (otherwise leave intact)
                if changed:
                    with csv_path.open("w", encoding="utf-8", newline="") as fh:
                        writer = csv.writer(fh)
                        writer.writerow(["Date", "ReleaseDate"])
                        for dt, rel in existing.items():
                            writer.writerow([dt, rel])

            else:
                # Case 2: Create fresh file
                with csv_path.open("w", encoding="utf-8", newline="") as fh:
                    writer = csv.writer(fh)
                    writer.writerow(["Date", "ReleaseDate"])
                    for d in date_cols:
                        writer.writerow([d, ""])

            # Open with system default
            try:
                if sys.platform.startswith("win"):
                    os.startfile(str(csv_path))  # type: ignore[attr-defined]
                elif sys.platform == "darwin":
                    subprocess.run(["open", str(csv_path)], check=False)
                else:
                    subprocess.run(["xdg-open", str(csv_path)], check=False)
            except Exception:
                pass

        except Exception as exc:
            messagebox.showerror("Error", f"Failed to generate ReleaseDate prompt:\n{exc}")

    def _on_copy_stock_multiplier_prompt(self) -> None:
        try:
            self._generate_multipliers_with_prompt()

        except FileNotFoundError as exc:
            messagebox.showerror("Stock Multipliers", str(exc))
        except Exception as exc:
            messagebox.showerror("Stock Multipliers", f"Failed to generate stock multiplier prompt:\n{exc}")

    def _on_get_stock_prices(self) -> None:
        try:
            company = self.company_var.get().strip()
            if not company:
                messagebox.showwarning("Company Missing", "Select a company before fetching stock prices.")
                return

            # Load release dates
            release_csv = self.companies_dir / company / "ReleaseDates.csv"
            if not release_csv.exists():
                messagebox.showerror(
                    "ReleaseDates.csv missing",
                    "Generate the release dates first using the Copy ReleaseDate Prompt button.",
                )
                return

            release_dates: List[str] = []
            with release_csv.open("r", encoding="utf-8") as fh:
                reader = csv.DictReader(fh)
                for row in reader:
                    rel = str(row.get("ReleaseDate", "")).strip()
                    if rel:
                        release_dates.append(rel)

            release_dates = _sort_dates(list(set(release_dates)))
            if not release_dates:
                messagebox.showwarning("No Release Dates", "Fill in ReleaseDates.csv before fetching stock prices.")
                return

            shift_values = [-30, -7, -1, 0, 1, 7, 30]

            from analyst.yahoo import get_stock_data_for_dates

            cache_path = self.companies_dir / company / "stock_cache.json"
            stock_df = get_stock_data_for_dates(
                ticker=company,
                dates=release_dates,
                days=shift_values,
                cache_filepath=str(cache_path),
            )

            if stock_df is None or stock_df.empty:
                messagebox.showwarning("No Prices", "No stock prices were retrieved for the selected company.")
                return

            offsets = [str(v) for v in shift_values]
            price_lookup: Dict[str, Dict[str, float]] = {}
            for _, row in stock_df.iterrows():
                release = str(row.get("BaseDate", "")).strip()
                offset_key = str(int(row.get("OffsetDays", 0)))
                price = row.get("Price")
                if not release:
                    continue
                if release not in price_lookup:
                    price_lookup[release] = {k: "" for k in offsets}  # type: ignore[assignment]
                if price_lookup[release].get(offset_key) in ("", None, float("nan")):
                    price_lookup[release][offset_key] = "" if pd.isna(price) else round(float(price), 4)

            # Merge with existing CSV if present
            csv_path = self.companies_dir / company / "StockPrices.csv"
            if csv_path.exists():
                with csv_path.open("r", encoding="utf-8", newline="") as fh:
                    reader = csv.DictReader(fh)
                    existing_cols = [c for c in (reader.fieldnames or []) if c != "ReleaseDate"]
                    for col in existing_cols:
                        if col not in offsets:
                            offsets.append(col)
                    for row in reader:
                        release = str(row.get("ReleaseDate", "")).strip()
                        if not release:
                            continue
                        if release not in price_lookup:
                            price_lookup[release] = {k: "" for k in offsets}  # type: ignore[assignment]
                        for off in existing_cols:
                            if off == "ReleaseDate":
                                continue
                            val = row.get(off, "")
                            if val not in ("", None):
                                price_lookup[release][off] = val

            # Ensure all release dates exist
            for rel in release_dates:
                price_lookup.setdefault(rel, {k: "" for k in offsets})

            csv_path.parent.mkdir(parents=True, exist_ok=True)
            with csv_path.open("w", encoding="utf-8", newline="") as fh:
                writer = csv.writer(fh)
                header = ["ReleaseDate"] + offsets
                writer.writerow(header)
                for rel in _sort_dates(list(price_lookup.keys())):
                    row_vals = price_lookup.get(rel, {})
                    writer.writerow([rel] + [row_vals.get(off, "") for off in offsets])

            try:
                if sys.platform.startswith("win"):
                    os.startfile(str(csv_path))  # type: ignore[attr-defined]
                elif sys.platform == "darwin":
                    subprocess.run(["open", str(csv_path)], check=False)
                else:
                    subprocess.run(["xdg-open", str(csv_path)], check=False)
            except Exception:
                pass

        except Exception as exc:
            messagebox.showerror("Stock Prices", f"Failed to fetch stock prices:\n{exc}")

    def _generate_multipliers_with_prompt(self) -> None:
        company = self.company_var.get().strip()
        if not company:
            messagebox.showwarning("Company Missing", "Select a company before generating stock multipliers.")
            return

        date_cols = self._get_pdf_table_dates()
        if not date_cols:
            messagebox.showwarning(
                "No Dates",
                "No date columns were found in the Date Columns by PDF table.",
            )
            return

        try:
            generate_and_open_stock_multipliers(
                logger=getattr(self, "logger", None),
                company_dir=self.companies_dir,
                date_columns=date_cols,
                current_company_name=company,
            )

            app_root = Path(getattr(self, "app_root", Path(__file__).resolve().parent))
            template_path = app_root / "prompts" / "Stock_Multipliers.txt"
            text = build_stock_multiplier_prompt(company, date_cols, template_path)

            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            self.root.update()

        except FileNotFoundError as exc:
            messagebox.showerror("Stock Multipliers", str(exc))
        except Exception as exc:
            messagebox.showerror(
                "Stock Multipliers",
                f"Failed to generate stock multipliers and prompt:\n{exc}",
            )

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


        # === New stock multiplier control buttons ===
        # 1) Generate Multipliers
        self.generate_multipliers_button = ttk.Button(
            controls, text="Generate Multipliers",
            command=self._generate_multipliers_with_prompt,
        )
        self.generate_multipliers_button.pack(side=tk.LEFT)

        # 2) Create Mapping
        self.mapping_create_button = ttk.Button(
            controls, text="Create Mapping", command=self.create_mapping_csv
        )
        self.mapping_create_button.pack(side=tk.LEFT)

        # 3) Copy ReleaseDate Prompt
        copy_prompt_btn = ttk.Button(
            controls, text="Copy ReleaseDate Prompt",
            command=self._on_copy_releasedate_prompt
        )
        copy_prompt_btn.pack(side=tk.LEFT, padx=(0, 20))

        # 4) Get stock prices
        self.get_stock_prices_button = ttk.Button(
            controls, text="Get stock prices", command=self._on_get_stock_prices
        )
        self.get_stock_prices_button.pack(side=tk.LEFT)

        # 5) Generate Table (rename Create)
        self.combined_create_button = ttk.Button(
            controls, text="Generate Table", command=self.create_combined_dataset
        )
        self.combined_create_button.pack(side=tk.LEFT)

        # === New button: Plot Stacked Visuals ===
        def _on_plot_stacked_visuals():
            try:
                if not self.combined_rows:
                    messagebox.showwarning("No Data", "No combined data loaded or generated.")
                    return

                from analyst.data import Company
                from analyst.plots import plot_stacked_financials

                company_name = self.company_var.get().strip()
                df_all = pd.DataFrame(
                    self.combined_rows, columns=self.combined_columns
                ).fillna("")

                company = Company.from_combined(
                    company_name, df_all, companies_dir=self.companies_dir
                )

                plot_stacked_financials(company)

            except Exception as e:
                messagebox.showerror("Plot Error", f"Failed to plot stacked visuals:\n{e}")
        self.plot_visuals_button = ttk.Button(
            controls, text="Plot Stacked Visuals", command=_on_plot_stacked_visuals
        )
        self.plot_visuals_button.pack(side=tk.LEFT)

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

    def clear_combined_table(self) -> None:
        """Remove any Combined.csv data from the UI and disable save actions."""

        # Reset in-memory data
        self.combined_columns = []
        self.combined_rows = []

        tv = getattr(self, "combined_table", None)
        if tv is not None:
            try:
                for iid in tv.get_children(""):
                    tv.delete(iid)
            except Exception:
                pass

            placeholder = ("placeholder",)
            try:
                tv.configure(columns=placeholder, displaycolumns=placeholder, show="headings")
                tv.heading("placeholder", text="No Combined data loaded")
                tv.column("placeholder", width=220, anchor=tk.CENTER, stretch=True)
            except Exception:
                pass


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
            text_columns = {
                "category",
                "subcategory",
                "item",
                "note",
                "key4coloring",
            }
            anchor = tk.W if name.lower() in text_columns else tk.E
            tv.heading(cid, text=name)
            width = self.get_scrape_column_width(name)
            tv.column(cid, width=width, anchor=anchor, stretch=True)
        for row in rows:
            raw_values = list(row[:len(col_ids)])
            if len(raw_values) < len(col_ids):
                raw_values += [""] * (len(col_ids) - len(raw_values))

            # === NEW: Visual formatting only (commas) ===
            formatted = []
            for c_name, val in zip(columns, raw_values):
                s = str(val).strip()

                # Text columns render unchanged
                if c_name.lower() in ("type", "category", "subcategory", "item", "note", "key4coloring"):
                    formatted.append(s)
                    continue

                # Empty → show empty
                if s == "":
                    formatted.append("")
                    continue

                # Try formatting numbers
                try:
                    num = float(s)
                    if num.is_integer():
                        formatted.append(f"{int(num):,}")
                    else:
                        formatted.append(f"{num:,.2f}")
                except Exception:
                    # Non-numeric → render raw
                    formatted.append(s)

            # Insert *formatted* display values
            iid = tv.insert("", "end", values=formatted)

            # Apply NOTE coloring based on raw data (unchanged)
            self._apply_note_color_to_combined_item(iid, columns, tv)

    def save_combined_to_csv(self, quiet: bool = False) -> None:
        if not self.combined_columns or not self.combined_rows:
            if not quiet:
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
            if not quiet:
                messagebox.showinfo("Save Combined", f"Saved combined CSV to:\n{out_path}")
        except Exception as exc:
            messagebox.showerror("Save Combined", f"Failed to save combined CSV: {exc}")

    def _combined_dynamic_column_offset(self) -> int:
        return len(COMBINED_BASE_COLUMNS)

    def _get_mapping_json_paths(self, company_name: str) -> Dict[str, Path]:
        base = (self.companies_dir / company_name).resolve()
        return {
            "Financial": base / "mapping_financial.json",
            "Income": base / "mapping_income.json",
        }

    def _load_key4color_lookup(self, company_name: str) -> Dict[str, Dict[str, str]]:
        lookup: Dict[str, Dict[str, str]] = {}
        if not company_name:
            return lookup
        paths = self._get_mapping_json_paths(company_name)
        for typ, path in paths.items():
            if not path.exists():
                continue
            try:
                raw_text = path.read_text(encoding="utf-8")
                data = json.loads(raw_text)
            except Exception as exc:  # pragma: no cover - defensive logging
                logger = getattr(self, "logger", None)
                if logger is not None:
                    logger.error("Failed to load %s: %s", path, exc)
                continue
            if not isinstance(data, dict):
                continue
            type_map = lookup.setdefault(typ, {})
            for canonical, originals in data.items():
                if not isinstance(canonical, str):
                    continue
                canonical_clean = canonical.strip()
                if canonical_clean:
                    type_map.setdefault(canonical_clean.casefold(), canonical_clean)
                if not isinstance(originals, list):
                    continue
                for item in originals:
                    if not isinstance(item, str):
                        continue
                    raw_item = item.strip()
                    if not raw_item:
                        continue
                    type_map.setdefault(raw_item.casefold(), canonical_clean or raw_item)
        return lookup

    def _update_mapping_buttons(self) -> None:
        return

    def open_mapping_csv(self) -> None:
        company_name = self.company_var.get().strip()
        if not company_name:
            messagebox.showinfo("Mapping", "Select a company before opening mapping files.")
            return
        paths = self._get_mapping_json_paths(company_name)
        existing = [p for p in paths.values() if p.exists()]

        # === NEW: open both files; show error only if none exist ===
        if not existing:
            expected = "\n".join(str(p) for p in paths.values()) or "<no files>"
            messagebox.showinfo(
                "Mapping",
                f"No mapping JSON files found for {company_name}.\n\nExpected at:\n{expected}",
            )
            self._update_mapping_buttons()
            return

        # === NEW: open ALL existing mapping files ===
        for path in existing:
            try:
                if sys.platform.startswith("win"):
                    os.startfile(str(path))  # type: ignore[attr-defined]
                elif sys.platform == "darwin":
                    subprocess.run(["open", str(path)], check=False)
                else:
                    subprocess.run(["xdg-open", str(path)], check=False)
            except Exception as exc:
                # Continue trying to open the others; fail only if all fail
                print(f"⚠️ Failed to open {path}: {exc}")

        try:
            # Optional: success message (only if both opened silently)
            pass
        except Exception as exc:
            messagebox.showerror("Mapping", f"Failed to open mapping file: {exc}")

    def create_mapping_csv(self) -> None:
        company_name = self.company_var.get().strip()
        if not company_name:
            messagebox.showinfo("Mapping", "Select a company before creating mapping files.")
            return
        if not self.combined_columns or not self.combined_rows:
            messagebox.showinfo(
                "Mapping",
                "Create the combined dataset first before generating mapping files.",
            )
            return

        required_cols = ["TYPE", "ITEM"]
        missing_cols = [col for col in required_cols if col not in self.combined_columns]
        if missing_cols:
            messagebox.showerror(
                "Mapping",
                f"Combined table is missing required columns: {', '.join(missing_cols)}",
            )
            return

        col_index = {col: self.combined_columns.index(col) for col in required_cols}
        items_by_type: Dict[str, Set[str]] = {}
        for row in self.combined_rows:
            try:
                typ = str(row[col_index["TYPE"]]).strip() if col_index["TYPE"] < len(row) else ""
                item = str(row[col_index["ITEM"]]).strip() if col_index["ITEM"] < len(row) else ""
            except Exception:
                continue
            if typ not in COLUMNS:
                continue
            if not item:
                continue
            items_by_type.setdefault(typ, set()).add(item)

        if not items_by_type:
            messagebox.showinfo(
                "Mapping",
                "No eligible rows were found to generate mapping files.",
            )
            return

        prompt_root_obj = getattr(self, "app_root", Path(__file__).resolve().parent)
        prompt_root = Path(prompt_root_obj)
        prompt_dir = prompt_root / "prompts"
        prompt_files: Dict[str, Path] = {
            "Financial": prompt_dir / "Mapping_Financial.txt",
            "Income": prompt_dir / "Mapping_Income.txt",
        }

        unsupported_types = [typ for typ in items_by_type if typ not in prompt_files and items_by_type.get(typ)]
        items_for_prompt: Dict[str, List[str]] = {}
        for typ, items in items_by_type.items():
            if typ not in prompt_files:
                continue
            if not items:
                continue
            items_for_prompt[typ] = sorted(items)

        if not items_for_prompt:
            note = ""
            if unsupported_types:
                note = (
                    "\n\nThese TYPE values are currently unsupported for auto-mapping: "
                    + ", ".join(sorted(set(unsupported_types)))
                )
            messagebox.showinfo(
                "Mapping",
                "No supported Financial or Income items were found for mapping." + note,
            )
            return

        missing_prompts = [typ for typ, path in prompt_files.items() if typ in items_for_prompt and not path.exists()]
        if missing_prompts:
            detail = "\n".join(str(prompt_files[typ]) for typ in missing_prompts)
            messagebox.showerror(
                "Mapping",
                f"Prompt file not found for TYPE(s): {', '.join(missing_prompts)}.\n\nChecked paths:\n{detail}",
            )
            return

        prompt_payloads: Dict[str, str] = {}
        empty_prompt_types: List[str] = []
        for typ, items in items_for_prompt.items():
            prompt_path = prompt_files[typ]
            prompt_text = prompt_path.read_text(encoding="utf-8").strip()
            if not prompt_text:
                empty_prompt_types.append(typ)
                continue
            item_block = "\n".join(items)
            combined_prompt = f"{prompt_text}\n\n{item_block}".strip()
            prompt_payloads[typ] = combined_prompt

        if empty_prompt_types:
            detail = "\n".join(str(prompt_files[typ]) for typ in empty_prompt_types)
            messagebox.showerror(
                "Mapping",
                f"Prompt file is empty for TYPE(s): {', '.join(empty_prompt_types)}.\n\nChecked paths:\n{detail}",
            )
            return

        if not prompt_payloads:
            messagebox.showerror(
                "Mapping",
                "No prompts were prepared for OpenAI processing.",
            )
            return

        api_key = self.api_key_var.get().strip()
        if not api_key:
            messagebox.showinfo("Mapping", "Enter an OpenAI API key before creating mapping files.")
            return

        try:
            from openai import OpenAI
        except ImportError:
            messagebox.showerror(
                "Mapping",
                "The 'openai' package is not installed. Install it to generate mapping files.",
            )
            return

        def _open_paths(paths: Iterable[Path]) -> None:
            for path in paths:
                try:
                    if sys.platform.startswith("win"):
                        os.startfile(str(path))  # type: ignore[attr-defined]
                    elif sys.platform == "darwin":
                        subprocess.run(["open", str(path)], check=False)
                    else:
                        subprocess.run(["xdg-open", str(path)], check=False)
                except Exception:
                    continue

        def _prompt_mapping_action(existing_paths: List[Path]) -> str:
            dialog = tk.Toplevel(self.root)
            dialog.title("Mapping")
            dialog.resizable(False, False)

            ttk.Label(
                dialog,
                text=(
                    "Mapping files already exist. Choose an action:\n"
                    "\n• Rerun generation\n• Open existing files\n• Cancel"
                ),
                justify=tk.LEFT,
                padding=(12, 10),
                wraplength=360,
            ).pack(fill=tk.X)

            files_frame = ttk.Frame(dialog, padding=(12, 0, 12, 8))
            files_frame.pack(fill=tk.BOTH, expand=True)
            ttk.Label(files_frame, text="Existing files:", justify=tk.LEFT).pack(anchor=tk.W)
            files_list = tk.Text(files_frame, height=4, width=50)
            files_list.pack(fill=tk.BOTH, expand=True)
            files_list.insert("1.0", "\n".join(str(p) for p in existing_paths))
            files_list.configure(state="disabled")

            choice: Dict[str, str] = {"value": "cancel"}

            def set_choice(value: str) -> None:
                choice["value"] = value
                dialog.destroy()

            btn_frame = ttk.Frame(dialog, padding=(12, 4, 12, 12))
            btn_frame.pack(fill=tk.X)
            ttk.Button(btn_frame, text="Rerun Generation", command=lambda: set_choice("rerun")).pack(
                side=tk.LEFT, padx=(0, 6)
            )
            ttk.Button(btn_frame, text="Open Existing", command=lambda: set_choice("open")).pack(
                side=tk.LEFT, padx=(0, 6)
            )
            ttk.Button(btn_frame, text="Cancel", command=lambda: set_choice("cancel")).pack(side=tk.LEFT)

            dialog.transient(self.root)
            dialog.grab_set()
            dialog.wait_window()
            return choice["value"]

        paths = self._get_mapping_json_paths(company_name)
        existing = [p for p in paths.values() if p.exists()]
        if existing:
            action = _prompt_mapping_action(existing)
            if action == "open":
                _open_paths(existing)
                self._update_mapping_buttons()
                return
            if action != "rerun":
                return

        logger = getattr(self, "logger", None)
        if logger is not None:
            summary = {typ: len(items) for typ, items in items_for_prompt.items()}
            logger.info("Starting Key4Coloring mapping for %s: %s", company_name, summary)

        progress_win = tk.Toplevel(self.root)
        progress_win.title("Mapping")
        progress_win.resizable(False, False)
        ttk.Label(
            progress_win,
            text="Processing mapping with OpenAI. This may take a moment...",
            padding=(12, 10),
            wraplength=320,
            justify=tk.LEFT,
        ).pack(fill=tk.X)
        progress_bar = ttk.Progressbar(progress_win, mode="indeterminate", length=260)
        progress_bar.pack(padx=16, pady=(0, 16))
        progress_bar.start(10)

        def close_progress() -> None:
            try:
                progress_bar.stop()
            except Exception:
                pass
            try:
                progress_win.destroy()
            except Exception:
                pass

        try:
            progress_win.transient(self.root)
            progress_win.lift()
        except Exception:
            pass
        progress_win.protocol("WM_DELETE_WINDOW", close_progress)

        model_name = DEFAULT_OPENAI_MODEL or "gpt-5"

        def run_prompt(typ: str, prompt_text: str) -> Dict[str, List[str]]:
            client = OpenAI(api_key=api_key)
            response = client.responses.create(
                model=model_name,
                input=[
                    {
                        "role": "system",
                        "content": "You are a financial statement mapping assistant.",
                    },
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "input_text",
                                "text": prompt_text,
                            }
                        ],
                    },
                ],
            )
            raw_response = self._extract_openai_response_text(response)
            cleaned = self._strip_code_fence(raw_response).strip()
            if not cleaned:
                raise RuntimeError(f"OpenAI returned an empty response for {typ} mapping")
            try:
                parsed = json.loads(cleaned)
            except json.JSONDecodeError as exc:
                raise RuntimeError(f"OpenAI response for {typ} is not valid JSON: {exc}") from exc
            if not isinstance(parsed, dict):
                raise RuntimeError(f"OpenAI response for {typ} is not a JSON object")
            return parsed

        def on_success(saved_paths: List[Path]) -> None:
            close_progress()
            if saved_paths:
                files_list = "\n".join(str(p) for p in saved_paths)
                note = ""
                if unsupported_types:
                    note = (
                        "\n\nSkipped TYPE values without prompts: "
                        + ", ".join(sorted(set(unsupported_types)))
                    )
                messagebox.showinfo(
                    "Mapping",
                    "Mapping JSON files created successfully:\n"
                    + files_list
                    + note,
                )
            else:
                messagebox.showwarning(
                    "Mapping",
                    "OpenAI responses were received but no files were written.",
                )
            self._update_mapping_buttons()
            try:
                self.create_combined_dataset()
            except Exception as exc:  # pragma: no cover - refresh best effort
                if logger is not None:
                    logger.warning(
                        "Failed to refresh combined dataset after Key4Coloring mapping: %s",
                        exc,
                    )

        def on_error(exc: Exception) -> None:
            close_progress()
            messagebox.showerror("Mapping", f"Failed to create mapping files: {exc}")

        def worker() -> None:
            try:
                results: Dict[str, Dict[str, List[str]]] = {}
                with ThreadPoolExecutor(max_workers=max(1, len(prompt_payloads))) as executor:
                    future_map = {
                        executor.submit(run_prompt, typ, prompt): typ
                        for typ, prompt in prompt_payloads.items()
                    }
                    for future in as_completed(future_map):
                        typ = future_map[future]
                        results[typ] = future.result()

                mapping_paths = self._get_mapping_json_paths(company_name)
                saved_paths: List[Path] = []
                for typ, data in results.items():
                    path = mapping_paths.get(typ)
                    if path is None:
                        continue
                    path.parent.mkdir(parents=True, exist_ok=True)
                    with path.open("w", encoding="utf-8") as fh:
                        json.dump(data, fh, ensure_ascii=False, indent=2)
                    saved_paths.append(path)

                self.root.after(0, lambda: on_success(saved_paths))
            except Exception as exc:  # pragma: no cover - best effort reporting
                if logger is not None:
                    logger.error("Failed to create Key4Coloring mapping: %s", exc, exc_info=True)
                self.root.after(0, lambda: on_error(exc))

        threading.Thread(target=worker, daemon=True).start()

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
                base_offset = self._combined_dynamic_column_offset()
                c_idx = base_offset + dyn_idx
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
            dyn_idx = idx - self._combined_dynamic_column_offset()
            if 0 <= dyn_idx < len(self.combined_rename_names):
                self.combined_rename_names[dyn_idx] = new_name

    def create_combined_dataset(self) -> None:
        dyn_cols, rows_by_type, warnings = self._build_date_matrix_data()
        self.combined_dyn_columns = dyn_cols

        company_name = self.company_var.get().strip()
        mapping_lookup: Dict[str, Dict[str, str]] = {}
        if company_name:
            mapping_lookup = self._load_key4color_lookup(company_name)

        conflicts = []
        duplicate_tracker: Dict[
            Tuple[str, str], Dict[Tuple[str, str, str], List[Dict[str, str]]]
        ] = {}
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
                    category = str(mapping.get("CATEGORY", "")).strip()
                    subcategory = str(mapping.get("SUBCATEGORY", "")).strip()
                    item = str(mapping.get("ITEM", "")).strip()
                    note_val = mapping.get("NOTE", "")
                    table_key = (entry.path.name, typ)
                    row_info = {
                        "type": str(mapping.get("TYPE", typ) or typ),
                        "category": category,
                        "subcategory": subcategory,
                        "item": item,
                        "note": str(note_val),
                        "pdf": entry.path.name,
                    }
                    duplicate_tracker.setdefault(table_key, {}).setdefault(
                        (category, subcategory, item),
                        [],
                    ).append(row_info)
                    # Include typ in key for downstream TYPE assignment
                    key = (category, subcategory, item, typ)
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
        duplicate_rows: Dict[Tuple[str, str], List[Dict[str, str]]] = {}
        for table_key, key_map in duplicate_tracker.items():
            for rows in key_map.values():
                if len(rows) > 1:
                    duplicate_rows.setdefault(table_key, []).extend(rows)

        if duplicate_rows:
            viewer = tk.Toplevel(self.root)
            viewer.title("Duplicate Entry Viewer")
            viewer.geometry("900x520")

            ttk.Label(
                viewer,
                text=(
                    "Duplicate CATEGORY/SUBCATEGORY/ITEM combinations were found "
                    "in one or more tables.\nEach table must contain unique "
                    "CATEGORY, SUBCATEGORY, ITEM entries. Please resolve the "
                    "duplicates in the source CSV files."
                ),
                font=("Segoe UI", 11, "bold"),
                wraplength=850,
                justify=tk.LEFT,
            ).pack(pady=6)

            nb = ttk.Notebook(viewer)
            nb.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

            columns = ("TYPE", "Category", "Subcategory", "Item", "Note", "PDF File")
            for (pdf_name, typ), rows in sorted(duplicate_rows.items()):
                tab = ttk.Frame(nb)
                nb.add(tab, text=f"{typ} | {pdf_name[:14]}")

                tree = ttk.Treeview(tab, columns=columns, show="headings", height=10)
                for col in columns:
                    tree.heading(col, text=col)
                    tree.column(
                        col,
                        width=150 if col in ("TYPE", "PDF File") else 140,
                        anchor="center",
                    )

                for row in rows:
                    tree.insert(
                        "",
                        tk.END,
                        values=(
                            row.get("type", typ),
                            row.get("category", ""),
                            row.get("subcategory", ""),
                            row.get("item", ""),
                            row.get("note", ""),
                            row.get("pdf", pdf_name),
                        ),
                    )

                vsb = ttk.Scrollbar(tab, orient="vertical", command=tree.yview)
                tree.configure(yscroll=vsb.set)
                vsb.pack(side=tk.RIGHT, fill=tk.Y)
                tree.pack(fill=tk.BOTH, expand=True)

            ttk.Button(viewer, text="Close", command=viewer.destroy).pack(pady=8)

            try:
                viewer.transient(self.root)
                viewer.focus_set()
                viewer.lift()
                viewer.grab_release()
            except Exception:
                pass

            viewer.protocol("WM_DELETE_WINDOW", viewer.destroy)
            return

        if conflicts:
            # === Build interactive NOTE Conflict Viewer ===
            viewer = tk.Toplevel(self.root)
            viewer.title("NOTE Conflict Viewer")
            viewer.geometry("950x550")

            ttk.Label(
                viewer,
                text=(
                    "Conflicting NOTE values detected while combining datasets.\n"
                    "TYPE column is included to show which statement the conflict belongs to."
                ),
                font=("Segoe UI", 11, "bold")
            ).pack(pady=6)

            nb = ttk.Notebook(viewer)
            nb.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

            for key, vals in conflicts:
                if len(key) != 4:
                    continue

                cat, sub, item, typ = key
                tab = ttk.Frame(nb)
                nb.add(tab, text=f"{typ} | {cat[:12]} | {sub[:12]}")

                tree = ttk.Treeview(
                    tab,
                    columns=("TYPE", "Category", "Subcategory", "Item", "Note", "PDF File"),
                    show="headings",
                    height=10
                )

                # Columns reordered to put TYPE first
                for col in ("TYPE", "Category", "Subcategory", "Item", "Note", "PDF File"):
                    tree.heading(col, text=col)
                    tree.column(col, width=150 if col == "TYPE" else 140, anchor="center")

                # Collect matching rows directly from source CSVs
                for entry in self.pdf_entries:
                    base = self._combined_scrape_dir_for_entry(entry)
                    csv_path = base / f"{typ}.csv"
                    if not csv_path.exists():
                        continue

                    try:
                        import csv as csvmod
                        with csv_path.open("r", encoding="utf-8") as fh:
                            reader = csvmod.DictReader(fh)
                            for row in reader:
                                if (
                                    row.get("CATEGORY", "") == cat
                                    and row.get("SUBCATEGORY", "") == sub
                                    and row.get("ITEM", "") == item
                                    and row.get("TYPE", typ) == typ
                                ):
                                    tree.insert(
                                        "",
                                        tk.END,
                                        values=(
                                            row.get("TYPE", typ),          # TYPE
                                            row.get("CATEGORY", ""),       # Category
                                            row.get("SUBCATEGORY", ""),    # Subcategory
                                            row.get("ITEM", ""),           # Item
                                            row.get("NOTE", ""),           # Note
                                            entry.path.name,               # PDF
                                        ),
                                    )
                    except Exception as e:
                        print(f"⚠️ Failed to read {csv_path}: {e}")

                # Add scroll bar
                vsb = ttk.Scrollbar(tab, orient="vertical", command=tree.yview)
                tree.configure(yscroll=vsb.set)
                vsb.pack(side=tk.RIGHT, fill=tk.Y)
                tree.pack(fill=tk.BOTH, expand=True)

                tree.bind(
                    "<Double-1>",
                    lambda e, tr=tree: self._on_note_conflict_double_click(e, tr),
                )

            ttk.Button(viewer, text="Close", command=viewer.destroy).pack(pady=8)

            # === Non-blocking configuration ===
            try:
                viewer.transient(self.root)
                viewer.focus_set()
                viewer.lift()
                viewer.grab_release()
            except Exception:
                pass

            viewer.protocol("WM_DELETE_WINDOW", viewer.destroy)
            return

        # Add TYPE as first column
        columns = COMBINED_BASE_COLUMNS + self.combined_rename_names

        # PDF source and multiplier rows are considered Meta
        pdf_summary_values = [dc.get("pdf", "") for dc in dyn_cols]
        pdf_summary = ["Meta", "PDF source", "", "", "excluded", ""] + pdf_summary_values

        # === Collect multiplier values for each type (Financial, Income, Shares) ===
        multipliers: Dict[str, Dict[str, str]] = {}
        for entry in self.pdf_entries:
            base = self._combined_scrape_dir_for_entry(entry)
            pdf_name = entry.path.name
            mults = {}
            for typ in ("Financial", "Income", "Shares"):
                mult_path = base / f"{typ}_multiplier.txt"
                if mult_path.exists():
                    try:
                        with mult_path.open("r", encoding="utf-8") as fh:
                            mults[typ] = fh.read().strip()
                    except Exception:
                        mults[typ] = ""
                else:
                    mults[typ] = ""
            multipliers[pdf_name] = mults

        # Create a summary line for multipliers for each type
        multiplier_rows = []
        for typ in ("Financial", "Income", "Shares"):
            # Mark NOTE column as 'meta' for metadata rows
            row = [f"{typ}", f"{typ} Multiplier", "", "", "excluded", ""]
            for dc in dyn_cols:
                pdf = dc.get("pdf", "")
                val = multipliers.get(pdf, {}).get(typ, "")
                row.append(val)
            multiplier_rows.append(row)

        # Append Share Multiplier row to pdf_summary
        share_row: List[str] = ["Meta", "Stock Multiplier", "", "", "excluded", ""] + ["" for _ in dyn_cols]
        try:
            import csv
            from pathlib import Path

            stock_path = self.companies_dir / company_name / "stock_multipliers.csv"

            # Ensure stock_multipliers.csv exists
            if not stock_path.exists():
                msg = f"The required file 'stock_multipliers.csv' is missing for {company_name}.\n\nExpected at:\n{stock_path}"
                self.logger.error(f"❌ {msg}")
                try:
                    messagebox.showerror(
                        "Missing Stock Multipliers",
                        f"{msg}\n\nPlease generate the file first using the 'Generate Multipliers' button."
                    )
                except Exception:
                    print(f"❌ {msg}")
                return

            # Load multipliers
            stock_data = {}
            with stock_path.open("r", encoding="utf-8") as fh:
                reader = csv.reader(fh)
                header = next(reader, [])
                for row in reader:
                    if len(row) >= 2:
                        stock_data[row[0].strip()] = row[1].strip()

            # Build Stock Multiplier row aligned by date columns
            share_row = ["Meta", "Stock Multiplier", "", "", "excluded", ""]
            for dc in dyn_cols:
                date_label = dc.get("default_name", "").strip()
                val = stock_data.get(date_label, "1")
                share_row.append(val)

            self.logger.info(f"🧮 Added 'Stock Multiplier' row from {stock_path} ({len(stock_data)} entries)")

        except FileNotFoundError:
            raise
        except Exception as e:
            self.logger.error(f"❌ Failed to append 'Stock Multiplier' row to PDF summary: {e}")
            raise

        # === Append Release Dates row and Stock Prices rows (one per offset) ===
        stock_price_rows: List[List[str]] = []
        release_map: Dict[str, str] = {}
        release_row: List[str] = ["Meta", "ReleaseDate", "", "", "excluded", ""] + ["" for _ in dyn_cols]

        release_csv = self.companies_dir / company_name / "ReleaseDates.csv"
        if release_csv.exists():
            with release_csv.open("r", encoding="utf-8") as fh:
                reader = csv.DictReader(fh)
                for row in reader:
                    date_key = str(row.get("Date", "")).strip()
                    rel_val = str(row.get("ReleaseDate", "")).strip()
                    if date_key:
                        release_map[date_key] = rel_val

        # Populate the release row using the current date columns
        for dc_idx, dc in enumerate(dyn_cols):
            date_label = dc.get("default_name", "").strip()
            release_row[len(COMBINED_BASE_COLUMNS) + dc_idx] = release_map.get(date_label, "")

        stock_prices_path = self.companies_dir / company_name / "StockPrices.csv"
        if stock_prices_path.exists():
            with stock_prices_path.open("r", encoding="utf-8", newline="") as fh:
                reader = csv.DictReader(fh)
                offsets = [c for c in (reader.fieldnames or []) if c and c != "ReleaseDate"]
                default_offsets = ["-30", "-7", "-1", "0", "1", "7", "30"]
                # Preserve requested order
                for off in default_offsets:
                    if off not in offsets:
                        offsets.append(off)

                price_table: Dict[str, Dict[str, str]] = {}
                for row in reader:
                    rel = str(row.get("ReleaseDate", "")).strip()
                    if not rel:
                        continue
                    price_table[rel] = {k: row.get(k, "") for k in offsets}

                for offset in offsets:
                    price_row = ["Stock", "Prices", offset, "", "excluded", ""]
                    for dc in dyn_cols:
                        fin_date = dc.get("default_name", "").strip()
                        rel_date = release_map.get(fin_date, "")
                        val = price_table.get(rel_date, {}).get(offset, "")
                        price_row.append(val)
                    stock_price_rows.append(price_row)

        if stock_price_rows:
            self.logger.info(f"🧾 Added {len(stock_price_rows)} Stock Prices rows from {stock_prices_path}")

        def key_sort(k: Tuple[str, str, str]) -> Tuple[str, str, str]:
            return (k[0] or "", k[1] or "", k[2] or "")

        rows_out: List[List[str]] = []
        for key in sorted(key_data.keys(), key=key_sort):
            rec = key_data[key]
            # Unpack typ as well
            cat, sub, item, typ = key
            note_val = rec.get("NOTE", "")
            values_map_dyn: Dict[Tuple[str, int], str] = rec.get("values_by_dyn", {})
            values_for_row: List[str] = []
            for dc in dyn_cols:
                pdf = str(dc.get("pdf", ""))
                pos = int(dc.get("index", 0))
                values_for_row.append(values_map_dyn.get((pdf, pos), ""))

            # Determine TYPE by current CSV being processed (for typ in COLUMNS)
            assigned_type = None
            for csv_type in COLUMNS:
                if csv_type.lower() in cat.lower():
                    assigned_type = csv_type
                    break
            if assigned_type is None:
                assigned_type = "Meta"

            # Use the typ from the key directly as TYPE
            assigned_type = typ if typ in COLUMNS else "Meta"
            key4_value = ""
            item_clean = item.strip()
            type_map = mapping_lookup.get(assigned_type, {})
            if item_clean:
                lookup_key = item_clean.casefold()
                key4_value = type_map.get(lookup_key, "") or type_map.get(item_clean, "")
            if not key4_value and item_clean:
                key4_value = item_clean
            rows_out.append([assigned_type, cat, sub, item, note_val, key4_value] + values_for_row)

        final_rows = [pdf_summary] + multiplier_rows + [share_row, release_row] + stock_price_rows + rows_out
        self._populate_combined_table(columns, final_rows)
        self.combined_columns = columns
        self.combined_rows = final_rows
        self._update_mapping_buttons()
        # === Sort date columns (and reorder rows accordingly) ===
        import re
        from datetime import datetime

        def _is_date_column(col: str) -> bool:
            return bool(
                re.match(r"\d{2}\.\d{2}\.\d{4}", str(col))
                or re.match(r"\d{4}-\d{2}-\d{2}", str(col))
                or re.match(r"\d{2}/\d{2}/\d{4}", str(col))
            )

        def _parse_date_str(s: str):
            for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y", "%d.%m.%y"):
                try:
                    return datetime.strptime(s, fmt)
                except Exception:
                    continue
            return datetime.min

        # Identify and sort date columns ascending
        date_cols = [c for c in columns if _is_date_column(c)]
        non_date_cols = [c for c in columns if c not in date_cols]
        sorted_dates = sorted(date_cols, key=_parse_date_str)
        sorted_columns = non_date_cols + sorted_dates

        # Rebuild rows to match sorted columns
        reordered_rows = []
        for row in self.combined_rows:
            mapping = {columns[i]: row[i] for i in range(min(len(columns), len(row)))}
            reordered_rows.append([mapping.get(c, "") for c in sorted_columns])

        # Update table and memory
        self.combined_columns = sorted_columns
        self.combined_rows = reordered_rows
        self._populate_combined_table(sorted_columns, reordered_rows)

        # Auto-save Combined.csv when the table is generated
        self.save_combined_to_csv()


    # === New: Load Combined.csv for a specific company ===
    def load_company_combined_csv(self, company_name: str) -> None:
        import pandas as pd, os
        try:
            company_name = company_name.strip()
            if not company_name:
                print("⚠️ No company name provided to load Combined.csv.")
                return

            company_folder = os.path.join(os.getcwd(), "companies", company_name)
            csv_path = os.path.join(company_folder, "Combined.csv")

            if not os.path.exists(csv_path):
                print(f"ℹ️ No Combined.csv found for {company_name}")
                return

            df = pd.read_csv(csv_path)
            # Replace NaN with empty string before populating
            df = df.fillna("")

            columns = list(df.columns)
            rows = df.values.tolist()

            # Sanitize to ensure all None or nan-like values are blank
            rows = [[("" if str(v).lower() in ("nan", "none", "null") else v) for v in row] for row in rows]

            self.combined_columns = columns
            self.combined_rows = rows

            self._populate_combined_table(columns, rows)
            self._update_mapping_buttons()

            print(f"✅ Loaded Combined.csv for {company_name} ({len(rows)} rows, {len(columns)} columns)")

        except Exception as e:
            print(f"❌ Error loading Combined.csv for {company_name}: {e}")
