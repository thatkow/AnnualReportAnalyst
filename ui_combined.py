from __future__ import annotations

import csv
import io
import os
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

import tkinter as tk
from tkinter import messagebox, simpledialog, ttk

from constants import COLUMNS, DEFAULT_OPENAI_MODEL, SCRAPE_EXPECTED_COLUMNS


COMBINED_BASE_COLUMNS = [
    "TYPE",
    "CATEGORY",
    "SUBCATEGORY",
    "ITEM",
    "NOTE",
    "CATEGORY_M",
    "SUBCATEGORY_M",
    "ITEM_M",
]
from ui_widgets import CollapsibleFrame
from pdf_utils import PDFEntry, normalize_header_row


class CombinedUIMixin:
    root: tk.Misc
    combined_tab: Optional[ttk.Frame]
    combined_date_tree: Optional[ttk.Treeview]
    combined_table: Optional[ttk.Treeview]
    combined_create_button: Optional[ttk.Button]
    combined_save_button: Optional[ttk.Button]
    mapping_create_button: Optional[ttk.Button]
    mapping_open_button: Optional[ttk.Button]
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

        # === New stock multiplier control buttons ===
        from combined_utils import open_stock_multipliers_file, reload_stock_multipliers, generate_and_open_stock_multipliers

        self.stock_multiplier_button = ttk.Button(
            controls, text="Stock Multiplier",
            command=lambda: open_stock_multipliers_file(logger=getattr(self, "logger", None),
                                                       company_dir=self.companies_dir,
                                                       current_company_name=self.company_var.get())
        )
        self.stock_multiplier_button.pack(side=tk.LEFT)

        self.reload_multipliers_button = ttk.Button(
            controls,
            text="Reload Multipliers",
            command=lambda: reload_stock_multipliers(self),
        )
        self.reload_multipliers_button.pack(side=tk.LEFT, padx=(6, 0))

        self.generate_multipliers_button = ttk.Button(
            controls, text="Generate Multipliers",
            command=lambda: generate_and_open_stock_multipliers(logger=getattr(self, "logger", None),
                                                                company_dir=self.companies_dir,
                                                                date_columns=self.combined_rename_names,
                                                                current_company_name=self.company_var.get())
        )
        self.generate_multipliers_button.pack(side=tk.LEFT, padx=(6, 0))
        self.combined_create_button = ttk.Button(controls, text="Create", command=self.create_combined_dataset)
        self.combined_create_button.pack(side=tk.LEFT)
        self.mapping_create_button = ttk.Button(controls, text="Create Mapping", command=self.create_mapping_csv)
        self.mapping_create_button.pack(side=tk.LEFT, padx=(6, 0))
        self.mapping_open_button = ttk.Button(controls, text="Open Mapping", command=self.open_mapping_csv, state="disabled")
        self.mapping_open_button.pack(side=tk.LEFT)
        self._update_mapping_buttons()
        self.combined_save_button = ttk.Button(controls, text="Save CSV", command=self.save_combined_to_csv, state="disabled")
        self.combined_save_button.pack(side=tk.LEFT, padx=(6, 0))

        # === New button: Plot Stacked Visuals ===
        from analyst_stackedvisuals import render_stacked_annual_report
        import pandas as pd

        def _on_plot_stacked_visuals():
            try:
                if not self.combined_rows:
                    messagebox.showwarning("No Data", "No combined data loaded or generated.")
                    return

                # Construct DataFrame from current table
                df = pd.DataFrame(self.combined_rows, columns=self.combined_columns).fillna("")

                # Extract rows containing multipliers
                share_mult_row = df[df["CATEGORY"].str.lower() == "shares multiplier"]
                stock_mult_row = df[df["CATEGORY"].str.lower() == "stock multiplier"]
                fin_mult_row = df[(df["CATEGORY"].str.lower() == "financial multiplier")]
                inc_mult_row = df[(df["CATEGORY"].str.lower() == "income multiplier")]
               
                excluded_cols = set(COMBINED_BASE_COLUMNS + ["Ticker"])
                num_cols = [c for c in df.columns if c not in excluded_cols]

                def extract_mult(row_df):
                    if row_df.empty:
                        return {}
                    row = row_df.iloc[0]
                    return {col: float(row[col]) if pd.notna(row[col]) and str(row[col]).strip() != "" else 1.0 for col in num_cols}

                share_mult = extract_mult(share_mult_row)
                stock_mult = extract_mult(stock_mult_row)
                fin_mult = extract_mult(fin_mult_row)
                inc_mult = extract_mult(inc_mult_row)

                # Determine company/ticker name
                company_name = self.company_var.get().strip()
                df["Ticker"] = company_name

                # Remove NOTE=exclude
                df = df[df["NOTE"].str.lower() != "excluded"].copy()

                # Negate NOTE=negated                
                for c in num_cols:
                    df[c] = pd.to_numeric(df[c], errors="coerce")
                neg_idx = df["NOTE"].str.lower() == "negated"
                df.loc[neg_idx, num_cols] = df.loc[neg_idx, num_cols].applymap(
                    lambda x: -1.0 * x if pd.notna(x) else x
                )

                # === Load multipliers ===
                from pathlib import Path
                import csv

                company_dir = self.companies_dir / company_name       

                # Apply Stock Multiplier and Share Multiplier to all "Number of shares" rows
                share_rows = df[df["ITEM"].str.lower() == "number of shares"]
                for c in num_cols:
                    df.loc[share_rows.index, c] = df.loc[share_rows.index, c] * share_mult.get(c, 1.0) * stock_mult.get(c, 1.0)

                # Apply Financial / Income multipliers to corresponding TYPE rows (excluding the multiplier rows)
                for c in num_cols:
                    df.loc[(df["TYPE"].str.lower() == "financial") & (df["ITEM"].str.lower() != "financial multiplier"), c] *= fin_mult.get(c, 1.0)
                    df.loc[(df["TYPE"].str.lower() == "income") & (df["ITEM"].str.lower() != "income multiplier"), c] *= inc_mult.get(c, 1.0)

                self.logger.info("✅ Applied share, stock, and type multipliers before plotting.")

                # === Inject Ticker column using company_name ===
                df["Ticker"] = company_name

                # === Identify year columns ===
                year_cols = [c for c in df.columns if c[:2].isdigit() or c.startswith("31.")]

                # === Replace factor_lookup with stock price-based lookup (inverted prices) ===
                from analyst_yahoo import get_stock_data_for_dates

                # Provide optional parameter for lookup label
                lookup_label = getattr(self, "factor_label", "Stock-adjusted factors")

                # Define date shifts for keys
                shift_keys = {
                    "": 0,  # blank entry maps to 1
                    "Prior month (-30)": -30,
                    "Prior week (-7)": -7,
                    "Prior day (-1)": -1,
                    "On release (0)": 0,
                    "Next day (+1)": 1,
                    "Next week (+7)": 7,
                    "Next month (+30)": 30,
                }

                ticker = company_name or "UNKNOWN"
                print(f"📈 Fetching stock prices for {ticker} (label: {lookup_label})...")

                stock_df = get_stock_data_for_dates(
                    ticker=ticker,
                    dates=year_cols,
                    days=[d for d in shift_keys.values()],
                    cache_filepath="stock_cache.json"
                )

                factor_lookup = {}


                # Blank entry maps every year to factor=1.0
                factor_lookup[""] = {y: 1.0 for y in year_cols}

                # Informational log
                print(f"🟦 Added blank factor_lookup entry for all years: {len(year_cols)} columns mapped to 1.0")

                for key_label, day_offset in shift_keys.items():
                    if key_label == "":
                        continue

                    subset = stock_df[stock_df["OffsetDays"] == day_offset]
                    price_map = {}
                    for _, row in subset.iterrows():
                        base = row["BaseDate"]
                        price = row["Price"]
                        if pd.notna(price) and price > 0:
                            price_map[base] = 1.0 / float(price)
                        else:
                            price_map[base] = float("nan")

                    factor_lookup[key_label] = price_map

                if factor_lookup:
                    print(f"✅ {lookup_label} generated:")
                    for k, v in factor_lookup.items():
                        ex = list(v.items())[:2]
                        print(f"  {k}: sample {ex}")
                else:
                    print("⚠️ Stock price lookup returned no data; continuing without adjustment.")

                # === Build factor_tooltip map (date → ["offset: price"]) ===
                factor_tooltip = {}
                for y in year_cols:
                    entries = []
                    for key_label, day_offset in shift_keys.items():
                        if key_label == "":
                            continue
                        # Retrieve 1/price from factor_lookup, invert to display original price
                        lookup_map = factor_lookup.get(key_label, {})
                        inv_val = lookup_map.get(y, None)
                        if inv_val is None or pd.isna(inv_val) or inv_val == 0:
                            entries.append(f"{day_offset:+d}: NaN")
                        else:
                            price_val = 1.0 / inv_val
                            entries.append(f"{day_offset:+d}: {price_val:.3f}")

                    factor_tooltip[y] = entries

                factor_tooltip_label = "Stock Prices"
                print(f"🟩 Built factor_tooltip (showing stock prices) with {len(factor_tooltip)} entries.")

                factor_tooltip_label = "Prices"
                print(f"🟩 Built factor_tooltip with {len(factor_tooltip)} entries.")

                # === Extract share counts from 'Number of shares' rows ===
                share_counts = {}
                share_rows = df[df["ITEM"].str.lower().str.contains("number of shares", na=False)]
                share_counts[company_name] = {}
                if not share_rows.empty:
                    row = share_rows.iloc[0]
                    for y in year_cols:
                        val = row.get(y)
                        share_counts[company_name][y] = float(val) if pd.notna(val) else 1.0
                else:
                    share_counts[company_name] = {y: 1.0 for y in year_cols}

                # === Remove 'Number of shares' rows for plotting ===
                df_plot = df[~df["ITEM"].str.lower().str.contains("number of shares", na=False)].copy()

                # === Output directory ===
                from pathlib import Path
                visuals_dir = Path("companies") / "visuals"
                visuals_dir.mkdir(parents=True, exist_ok=True)

                out_path = visuals_dir / f"ARVisuals_{company_name}.html"

                # === Call the stacked visuals plotter ===
                render_stacked_annual_report(
                    df_plot,
                    title=f"Financial/Income for {company_name}",
                    factor_lookup=factor_lookup,
                    factor_label="Value Per Stock Price Dollar",
                    factor_tooltip=factor_tooltip,
                    factor_tooltip_label=factor_tooltip_label,
                    share_counts=share_counts,
                    out_path=out_path,
                )

            except Exception as e:
                messagebox.showerror("Plot Error", f"Failed to plot stacked visuals:\n{e}")

        self.plot_visuals_button = ttk.Button(controls, text="Plot Stacked Visuals", command=_on_plot_stacked_visuals)
        self.plot_visuals_button.pack(side=tk.LEFT, padx=(6, 0))

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
                "category_m",
                "subcategory_m",
                "item_m",
            }
            anchor = tk.W if name.lower() in text_columns else tk.E
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

    def _combined_dynamic_column_offset(self) -> int:
        return len(COMBINED_BASE_COLUMNS)

    def _get_mapping_path(self, company_name: str) -> Path:
        return (self.companies_dir / company_name / "mapping.csv").resolve()

    def _load_mapping_lookup(
        self, company_name: str
    ) -> Dict[Tuple[str, str, str, str], Dict[str, str]]:
        result: Dict[Tuple[str, str, str, str], Dict[str, str]] = {}
        if not company_name:
            return result
        path = self._get_mapping_path(company_name)
        if not path.exists():
            return result
        try:
            with path.open("r", encoding="utf-8", newline="") as fh:
                reader = csv.DictReader(fh)
                for row in reader:
                    typ = str(row.get("TYPE", "")).strip()
                    category = str(row.get("CATEGORY", "")).strip()
                    subcategory = str(row.get("SUBCATEGORY", "")).strip()
                    item = str(row.get("ITEM", "")).strip()
                    key = (typ, category, subcategory, item)
                    result[key] = {
                        "CATEGORY_M": str(row.get("CATEGORY_M", "")).strip(),
                        "SUBCATEGORY_M": str(row.get("SUBCATEGORY_M", "")).strip(),
                        "ITEM_M": str(row.get("ITEM_M", "")).strip(),
                    }
        except Exception as exc:  # pragma: no cover - defensive logging
            logger = getattr(self, "logger", None)
            if logger is not None:
                logger.error("Failed to load mapping.csv for %s: %s", company_name, exc)
        return result

    def _update_mapping_buttons(self) -> None:
        button = getattr(self, "mapping_open_button", None)
        if button is None:
            return
        company_name = self.company_var.get().strip()
        path: Optional[Path] = None
        if company_name:
            try:
                path = self._get_mapping_path(company_name)
            except Exception:
                path = None
        exists = bool(path and path.exists())
        try:
            button.configure(state="normal" if exists else "disabled")
        except Exception:
            pass

    def open_mapping_csv(self) -> None:
        company_name = self.company_var.get().strip()
        if not company_name:
            messagebox.showinfo("Mapping CSV", "Select a company before opening mapping.csv.")
            return
        path = self._get_mapping_path(company_name)
        if not path.exists():
            messagebox.showinfo(
                "Mapping CSV",
                f"No mapping.csv found for {company_name}.\n\nExpected at:\n{path}",
            )
            self._update_mapping_buttons()
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(path))  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.run(["open", str(path)], check=False)
            else:
                subprocess.run(["xdg-open", str(path)], check=False)
        except Exception as exc:
            messagebox.showerror("Mapping CSV", f"Failed to open mapping.csv: {exc}")

    def create_mapping_csv(self) -> None:
        company_name = self.company_var.get().strip()
        if not company_name:
            messagebox.showinfo("Mapping CSV", "Select a company before creating mapping.csv.")
            return
        if not self.combined_columns or not self.combined_rows:
            messagebox.showinfo(
                "Mapping CSV",
                "Create the combined dataset first before generating mapping.csv.",
            )
            return

        required_cols = ["TYPE", "CATEGORY", "SUBCATEGORY", "ITEM"]
        missing_cols = [col for col in required_cols if col not in self.combined_columns]
        if missing_cols:
            messagebox.showerror(
                "Mapping CSV",
                f"Combined table is missing required columns: {', '.join(missing_cols)}",
            )
            return

        col_index = {col: self.combined_columns.index(col) for col in required_cols}
        seen: Set[Tuple[str, str, str, str]] = set()
        export_rows: List[List[str]] = []
        for row in self.combined_rows:
            try:
                values = {
                    col: str(row[col_index[col]]).strip() if col_index[col] < len(row) else ""
                    for col in required_cols
                }
            except Exception:
                continue
            typ = values["TYPE"]
            if typ not in COLUMNS:
                continue
            category = values["CATEGORY"]
            subcategory = values["SUBCATEGORY"]
            item = values["ITEM"]
            key = (typ, category, subcategory, item)
            if key in seen:
                continue
            seen.add(key)
            export_rows.append([typ, category, subcategory, item])

        if not export_rows:
            messagebox.showinfo(
                "Mapping CSV",
                "No eligible rows were found to generate mapping.csv.",
            )
            return

        export_rows.sort()

        prompt_root = getattr(self, "app_root", Path(__file__).resolve().parent)
        prompt_path = prompt_root / "mapping_prompt.txt"
        if not prompt_path.exists():
            messagebox.showerror(
                "Mapping CSV",
                f"Prompt file not found at {prompt_path}.",
            )
            return
        prompt_text = prompt_path.read_text(encoding="utf-8").strip()
        if not prompt_text:
            messagebox.showerror("Mapping CSV", "mapping_prompt.txt is empty.")
            return

        api_key = self.api_key_var.get().strip()
        if not api_key:
            messagebox.showinfo("Mapping CSV", "Enter an OpenAI API key before creating mapping.csv.")
            return

        try:
            from openai import OpenAI
        except ImportError as exc:
            messagebox.showerror(
                "Mapping CSV",
                "The 'openai' package is not installed. Install it to generate mapping.csv.",
            )
            return

        csv_buffer = io.StringIO()
        writer = csv.writer(csv_buffer, quoting=csv.QUOTE_ALL)
        writer.writerow(required_cols)
        writer.writerows(export_rows)
        payload = csv_buffer.getvalue().strip()

        if not payload:
            messagebox.showerror(
                "Mapping CSV",
                "No data was generated for mapping.csv.",
            )
            return

        combined_prompt = f"{prompt_text}\n\n{payload}" if prompt_text else payload

        try:
            client = OpenAI(api_key=api_key)
            response = client.responses.create(
                model=DEFAULT_OPENAI_MODEL,
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
                                "text": combined_prompt,
                            }
                        ],
                    },
                ],
            )
            raw_response = self._extract_openai_response_text(response)
            cleaned = self._strip_code_fence(raw_response).strip()
            if not cleaned:
                raise RuntimeError("OpenAI returned an empty response")

            reader = csv.DictReader(io.StringIO(cleaned))
            expected = required_cols + ["CATEGORY_M", "SUBCATEGORY_M", "ITEM_M"]
            fieldnames = reader.fieldnames or []
            missing = [col for col in expected if col not in fieldnames]
            if missing:
                raise RuntimeError(
                    f"OpenAI response is missing expected columns: {', '.join(missing)}"
                )

            mapping_path = self._get_mapping_path(company_name)
            mapping_path.parent.mkdir(parents=True, exist_ok=True)
            mapping_path.write_text(cleaned, encoding="utf-8")

            messagebox.showinfo(
                "Mapping CSV",
                f"mapping.csv created successfully at:\n{mapping_path}",
            )
            self._update_mapping_buttons()
            try:
                self.create_combined_dataset()
            except Exception as exc:  # pragma: no cover - refresh best effort
                logger = getattr(self, "logger", None)
                if logger is not None:
                    logger.warning("Failed to refresh combined dataset after mapping.csv: %s", exc)
        except Exception as exc:
            logger = getattr(self, "logger", None)
            if logger is not None:
                logger.error("Failed to create mapping.csv: %s", exc)
            messagebox.showerror("Mapping CSV", f"Failed to create mapping.csv: {exc}")
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
        mapping_lookup: Dict[Tuple[str, str, str, str], Dict[str, str]] = {}
        if company_name:
            mapping_lookup = self._load_mapping_lookup(company_name)

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
                    category = str(mapping.get("CATEGORY", "")).strip()
                    subcategory = str(mapping.get("SUBCATEGORY", "")).strip()
                    item = str(mapping.get("ITEM", "")).strip()
                    note_val = mapping.get("NOTE", "")
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
        if conflicts:
            # === Build interactive NOTE Conflict Viewer ===
            viewer = tk.Toplevel(self.root)
            viewer.title("NOTE Conflict Viewer")
            viewer.geometry("950x550")

            ttk.Label(
                viewer,
                text="Conflicting NOTE values detected while combining datasets.",
                font=("Segoe UI", 11, "bold")
            ).pack(pady=6)

            nb = ttk.Notebook(viewer)
            nb.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

            for key, vals in conflicts:
                if len(key) != 4:
                    continue

                cat, sub, item, typ = key
                tab = ttk.Frame(nb)
                nb.add(tab, text=f"{cat[:12]} | {sub[:12]} | {typ}")

                tree = ttk.Treeview(
                    tab,
                    columns=("Category", "Subcategory", "Item", "Type", "Note", "PDF File"),
                    show="headings",
                    height=10
                )

                for col in ("Category", "Subcategory", "Item", "Type", "Note", "PDF File"):
                    tree.heading(col, text=col)
                    tree.column(col, width=140, anchor="center")

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
                                            row.get("CATEGORY", ""),
                                            row.get("SUBCATEGORY", ""),
                                            row.get("ITEM", ""),
                                            row.get("TYPE", ""),
                                            row.get("NOTE", ""),
                                            entry.path.name,
                                        ),
                                    )
                    except Exception as e:
                        print(f"⚠️ Failed to read {csv_path}: {e}")

                # Add scroll bar
                vsb = ttk.Scrollbar(tab, orient="vertical", command=tree.yview)
                tree.configure(yscroll=vsb.set)
                vsb.pack(side=tk.RIGHT, fill=tk.Y)
                tree.pack(fill=tk.BOTH, expand=True)

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
        pdf_summary = ["Meta", "PDF source", "", "", "excluded", "", "", ""] + pdf_summary_values

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
            row = [f"{typ}", f"{typ} Multiplier", "", "", "excluded", "", "", ""]
            for dc in dyn_cols:
                pdf = dc.get("pdf", "")
                val = multipliers.get(pdf, {}).get(typ, "")
                row.append(val)
            multiplier_rows.append(row)

        # Append Share Multiplier row to pdf_summary
        share_row: List[str] = ["Meta", "Stock Multiplier", "", "", "excluded", "", "", ""] + ["" for _ in dyn_cols]
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
            share_row = ["Meta", "Stock Multiplier", "", "", "excluded", "", "", ""]
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
            mapping_key = (assigned_type, cat, sub, item)
            mapped_values = mapping_lookup.get(mapping_key, {})
            cat_m = mapped_values.get("CATEGORY_M", "")
            sub_m = mapped_values.get("SUBCATEGORY_M", "")
            item_m = mapped_values.get("ITEM_M", "")
            rows_out.append([assigned_type, cat, sub, item, note_val, cat_m, sub_m, item_m] + values_for_row)

        final_rows = [pdf_summary] + multiplier_rows + [share_row] + rows_out
        self._populate_combined_table(columns, final_rows)
        self.combined_columns = columns
        self.combined_rows = final_rows
        self._update_mapping_buttons()
        if self.combined_save_button is not None:
            self.combined_save_button.configure(state="normal")
        messagebox.showinfo("Combined", f"Combined dataset created with {len(final_rows)} rows and {len(columns)} columns.")

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

            if self.combined_save_button is not None:
                self.combined_save_button.configure(state="normal")

            print(f"✅ Loaded Combined.csv for {company_name} ({len(rows)} rows, {len(columns)} columns)")

        except Exception as e:
            print(f"❌ Error loading Combined.csv for {company_name}: {e}")
