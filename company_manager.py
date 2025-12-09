from __future__ import annotations

import json
import os
import sys
import re
import shutil
import subprocess
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

try:
    import fitz  # type: ignore[import-untyped]
except ImportError:  # pragma: no cover - handled at runtime
    fitz = None  # type: ignore[assignment]

from PIL import ImageTk

import pandas as pd

from pdf_utils import PDFEntry
from analyst import yahoo


class CompanyManagerMixin:
    root: tk.Misc
    companies_dir: Path
    company_var: tk.StringVar
    company_options: List[str]
    company_selector_window: Optional[tk.Toplevel]
    folder_path: tk.StringVar
    assigned_pages: Dict[str, Dict[str, Any]]
    assigned_pages_path: Optional[Path]
    pdf_entries: List[PDFEntry]
    downloads_dir: tk.StringVar
    recent_download_minutes: tk.IntVar

    def _open_company_selector(self) -> None:
        if not self.company_options:
            messagebox.showinfo("Select Company", "No companies available. Create one first.")
            return

        if self.company_selector_window is not None and self.company_selector_window.winfo_exists():
            try:
                self.company_selector_window.lift()
                self.company_selector_window.focus_force()
            except tk.TclError:
                pass
            return

        window = tk.Toplevel(self.root)
        window.title("Select Company")
        window.transient(self.root)
        window.resizable(False, False)
        self.company_selector_window = window

        # Build UI elements before making modal
        frame = ttk.Frame(window, padding=12)
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="Company:").pack(anchor="w")
        initial = self.company_var.get().strip()
        if initial not in self.company_options and self.company_options:
            initial = self.company_options[0]
        combo_var = tk.StringVar(value=initial)
        combo = ttk.Combobox(frame, state="readonly", values=self.company_options, textvariable=combo_var)
        combo.pack(fill=tk.X, pady=(4, 12))

        button_row = ttk.Frame(frame)
        button_row.pack(fill=tk.X)

        def close_dialog() -> None:
            if self.company_selector_window is None:
                return
            try:
                window.grab_release()
            except tk.TclError:
                pass
            try:
                window.destroy()
            except tk.TclError:
                pass
            self.company_selector_window = None

        def on_confirm() -> None:
            selection = combo.get().strip()
            if selection:
                self._set_active_company(selection)
                window.result = True
            else:
                window.result = False
            close_dialog()

        def on_cancel() -> None:
            window.result = False
            close_dialog()

        ttk.Button(button_row, text="Cancel", command=on_cancel).pack(side=tk.RIGHT)
        ttk.Button(button_row, text="Select", command=on_confirm).pack(side=tk.RIGHT, padx=(0, 8))

        window.bind("<Return>", lambda _e: on_confirm())
        window.bind("<Escape>", lambda _e: on_cancel())
        window.protocol("WM_DELETE_WINDOW", on_cancel)
        combo.focus_set()

        # Now make the dialog modal (after UI exists)
        window.grab_set()
        self.root.wait_window(window)

        # Return result to caller
        return getattr(window, "result", False)

        frame = ttk.Frame(window, padding=12)
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="Company:").pack(anchor="w")
        initial = self.company_var.get().strip()
        if initial not in self.company_options and self.company_options:
            initial = self.company_options[0]
        combo_var = tk.StringVar(value=initial)
        combo = ttk.Combobox(frame, state="readonly", values=self.company_options, textvariable=combo_var)
        combo.pack(fill=tk.X, pady=(4, 12))

        button_row = ttk.Frame(frame)
        button_row.pack(fill=tk.X)

        def close_dialog() -> None:
            if self.company_selector_window is None:
                return
            try:
                window.grab_release()
            except tk.TclError:
                pass
            try:
                window.destroy()
            except tk.TclError:
                pass
            self.company_selector_window = None

        def on_confirm() -> None:
            selection = combo.get().strip()
            if selection:
                self._set_active_company(selection)
            close_dialog()

        def on_cancel() -> None:
            close_dialog()

        ttk.Button(button_row, text="Cancel", command=on_cancel).pack(side=tk.RIGHT)
        ttk.Button(button_row, text="Select", command=on_confirm).pack(side=tk.RIGHT, padx=(0, 8))

        window.bind("<Return>", lambda _e: on_confirm())
        window.bind("<Escape>", lambda _e: on_cancel())
        window.protocol("WM_DELETE_WINDOW", on_cancel)
        combo.focus_set()

    def _set_active_company(self, name: str, *, save: bool = True) -> None:
        normalized = name.strip()
        if not normalized:
            return
        if normalized not in self.company_options:
            self.company_options.append(normalized)
            self.company_options.sort()
        self.company_var.set(normalized)
        self._set_folder_for_company(normalized)
        if save:
            self._save_config()

    def _refresh_company_options(self) -> None:
        if not self.companies_dir.exists():
            self.companies_dir.mkdir(parents=True, exist_ok=True)

        companies = sorted([p.name for p in self.companies_dir.iterdir() if p.is_dir()])
        self.company_options = companies

        current = self.company_var.get().strip()
        if current and current in companies:
            self._set_active_company(current, save=False)
        elif companies:
            self._set_active_company(companies[0], save=False)
        else:
            self.company_var.set("")
            self.folder_path.set("")
            self.clear_entries()

    def _set_folder_for_company(self, company: str) -> None:
        folder = self.companies_dir / company / "raw"
        self.folder_path.set(str(folder))
        self._load_assigned_pages(company)
        self.clear_entries()

    def _load_assigned_pages(self, company: str) -> None:
        self.assigned_pages = {}
        self.assigned_pages_path = self.companies_dir / company / "assigned.json"
        if not self.assigned_pages_path.exists():
            return
        try:
            with self.assigned_pages_path.open("r", encoding="utf-8") as fh:
                data = json.load(fh)
        except Exception:
            return

        if not isinstance(data, dict):
            return

        parsed: Dict[str, Dict[str, Any]] = {}
        for pdf_name, value in data.items():
            if not isinstance(pdf_name, str) or not isinstance(value, dict):
                continue

            record: Dict[str, Any] = {}
            selections_obj = value.get("selections") if "selections" in value else value
            if isinstance(selections_obj, dict):
                selections: Dict[str, int] = {}
                for category, raw_page in selections_obj.items():
                    try:
                        selections[category] = int(raw_page)
                    except (TypeError, ValueError):
                        continue
                if selections:
                    record["selections"] = selections

            multi_obj = value.get("multi_selections")
            if isinstance(multi_obj, dict):
                multi: Dict[str, List[int]] = {}
                for category, raw_list in multi_obj.items():
                    if not isinstance(raw_list, list):
                        continue
                    pages: List[int] = []
                    for raw_page in raw_list:
                        try:
                            pages.append(int(raw_page))
                        except (TypeError, ValueError):
                            continue
                    if pages:
                        multi[category] = pages
                if multi:
                    record["multi_selections"] = multi

            year_value = value.get("year")
            if isinstance(year_value, str):
                record["year"] = year_value
            elif isinstance(year_value, (int, float)):
                record["year"] = str(int(year_value))

            if record:
                parsed[pdf_name] = record

        self.assigned_pages = parsed

    def _set_downloads_dir(self) -> None:
        initial = self.downloads_dir.get() or str(Path.home())
        selected = filedialog.askdirectory(parent=self.root, initialdir=initial, title="Select Downloads Directory")
        if not selected:
            return
        self.downloads_dir.set(selected)
        self._save_config()

    def _lookup_recent_stock_price(self, ticker: str) -> Optional[float]:
        """Return the most recent stock price within the last month for ``ticker``.

        The lookup relies on ``analyst.yahoo.get_stock_prices`` and walks backward day
        by day (up to 30 days) from today to find the latest available closing price.
        """

        try:
            prices = yahoo.get_stock_prices(ticker, years=1)
        except Exception as exc:
            print(f"⚠️ Unable to fetch stock prices for {ticker}: {exc}")
            return None

        if prices.empty:
            return None

        normalized = prices.copy()

        # Flatten multi-index columns that may include ticker names (e.g., [('Price', 'NXT.AX')])
        if isinstance(normalized.columns, pd.MultiIndex):
            normalized.columns = [
                " ".join(str(part) for part in col if str(part)).strip()
                for col in normalized.columns
            ]

        def _pick_column(columns: pd.Index, target: str) -> Optional[str]:
            """Return the best column match for ``target`` allowing suffixes like ticker codes."""

            if target in columns:
                return target

            target_lower = target.lower()
            candidates = [
                col
                for col in columns
                if str(col).lower() == target_lower
                or str(col).split(" ")[0].lower() == target_lower
                or target_lower in str(col).lower()
            ]

            return candidates[0] if candidates else None

        date_col = _pick_column(normalized.columns, "Date")
        price_col = _pick_column(normalized.columns, "Price")

        if not date_col or not price_col:
            print(
                f"⚠️ Stock price data for {ticker} missing expected columns: {normalized.columns.tolist()}"
            )
            return None

        normalized = normalized.rename(columns={date_col: "Date", price_col: "Price"})

        normalized["Date"] = pd.to_datetime(normalized["Date"], errors="coerce").dt.date

        price_values = normalized["Price"]
        if isinstance(price_values, pd.DataFrame):
            price_values = price_values.squeeze(axis=1)

        normalized["Price"] = pd.to_numeric(price_values, errors="coerce")
        normalized = normalized.dropna(subset=["Date", "Price"])

        if normalized.empty:
            return None

        price_lookup = {d: float(p) for d, p in zip(normalized["Date"], normalized["Price"])}
        today = datetime.now().date()

        for offset in range(0, 31):  # today + past 30 days
            candidate_date = today - timedelta(days=offset)
            price = price_lookup.get(candidate_date)
            if price is not None:
                return price

        return None

    def create_company(self) -> None:
        recent_pdfs = self._collect_recent_downloads()
        preview_window: Optional[tk.Toplevel] = None
        if recent_pdfs:
            preview_window = self._show_recent_download_previews(recent_pdfs)
        else:
            downloads_dir = self.downloads_dir.get().strip()
            if downloads_dir:
                messagebox.showinfo(
                    "Recent Downloads",
                    "No recently downloaded PDFs were found in the configured downloads folder.",
                )

        name = simpledialog.askstring("New Company", "Enter a name for the new company:", parent=self.root)
        if name is None:
            if preview_window is not None and preview_window.winfo_exists():
                preview_window.destroy()
            return

        normalized = name.strip()
        if not normalized:
            messagebox.showwarning("Invalid Name", "Company name cannot be empty.")
            if preview_window is not None and preview_window.winfo_exists():
                preview_window.destroy()
            return

        safe_name = re.sub(r"[\\/:*?\"<>|]", "_", normalized).strip().strip(".")
        if not safe_name:
            messagebox.showwarning(
                "Invalid Name",
                "Company name contains only unsupported characters. Please choose a different name.",
            )
            if preview_window is not None and preview_window.winfo_exists():
                preview_window.destroy()
            return

        if safe_name != normalized:
            messagebox.showinfo(
                "Company Name Adjusted",
                f"Using '{safe_name}' as the folder name due to unsupported characters.",
            )

        company_dir = self.companies_dir / safe_name
        raw_dir = company_dir / "raw"
        try:
            raw_dir.mkdir(parents=True, exist_ok=True)
        except Exception as exc:
            messagebox.showerror("Create Company", f"Could not create folders for '{safe_name}': {exc}")
            if preview_window is not None and preview_window.winfo_exists():
                preview_window.destroy()
            return

        moved_files = 0
        for pdf_path in recent_pdfs:
            try:
                destination = self._ensure_unique_path(raw_dir / pdf_path.name)
                shutil.move(str(pdf_path), str(destination))
                moved_files += 1
            except Exception as exc:
                messagebox.showwarning("Move PDF", f"Could not move '{pdf_path.name}': {exc}")

        self.company_var.set(safe_name)
        self._refresh_company_options()
        self._set_active_company(safe_name)
        self._open_in_file_manager(raw_dir)

        # Automatically load the new company's PDFs and Combined.csv (if available)
        folder_exists = raw_dir.exists()
        if folder_exists:
            load_pdfs = getattr(self, "load_pdfs", None)
            if callable(load_pdfs):
                try:
                    load_pdfs()
                except Exception as exc:
                    messagebox.showwarning(
                        "Load Company",
                        f"Created '{safe_name}', but failed to load its PDFs automatically:\n{exc}",
                    )

            load_combined = getattr(self, "load_company_combined_csv", None)
            if callable(load_combined):
                try:
                    load_combined(safe_name)
                except Exception as exc:
                    print(f"⚠️ Unable to load Combined.csv for {safe_name}: {exc}")

        price = self._lookup_recent_stock_price(safe_name)
        if price is not None:
            messagebox.showinfo(
                "Stock Price Available",
                f"Latest available stock price for {safe_name}: ${price:,.2f}",
                parent=self.root,
            )
        else:
            messagebox.showwarning(
                "Stock Price Unavailable",
                f"No stock price found for {safe_name} within the last month.",
                parent=self.root,
            )

        if preview_window is not None and preview_window.winfo_exists():
            preview_window.destroy()
        if moved_files:
            messagebox.showinfo("Create Company", f"Moved {moved_files} PDF(s) into '{safe_name}/raw'.")

    def _collect_recent_downloads(self) -> List[Path]:
        downloads_dir = self.downloads_dir.get().strip()
        if not downloads_dir:
            messagebox.showinfo(
                "Downloads Directory",
                "Configure the downloads directory before creating a new company.",
            )
            return []

        directory = Path(downloads_dir)
        if not directory.exists():
            messagebox.showerror("Downloads Directory", f"The folder '{downloads_dir}' does not exist.")
            return []

        minutes = self.recent_download_minutes.get()
        if minutes <= 0:
            minutes = 5
            self.recent_download_minutes.set(minutes)
            self._save_pattern_config()
        cutoff_ts = (datetime.now() - timedelta(minutes=minutes)).timestamp()

        recent: List[Tuple[float, Path]] = []
        for path in directory.iterdir():
            if not path.is_file() or path.suffix.lower() != ".pdf":
                continue
            try:
                stat_info = path.stat()
            except OSError:
                continue
            if stat_info.st_mtime >= cutoff_ts:
                recent.append((stat_info.st_mtime, path))

        recent.sort(key=lambda item: item[0], reverse=True)
        return [path for _mtime, path in recent]

    def _show_recent_download_previews(self, pdf_paths: List[Path]) -> tk.Toplevel:
        window = tk.Toplevel(self.root)
        window.title("Recent Downloads Preview")
        window.transient(self.root)

        container = ttk.Frame(window, padding=8)
        container.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(container, borderwidth=0)
        scrollbar = ttk.Scrollbar(container, orient=tk.VERTICAL, command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        inner = ttk.Frame(canvas)
        window_id = canvas.create_window((0, 0), window=inner, anchor="nw")

        inner.bind("<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda event: canvas.itemconfigure(window_id, width=event.width))

        previews: List[ImageTk.PhotoImage] = []
        for pdf_path in pdf_paths:
            item_frame = ttk.Frame(inner, padding=8)
            item_frame.pack(fill=tk.X, expand=True, pady=4)
            modified_text = ""
            try:
                modified = datetime.fromtimestamp(pdf_path.stat().st_mtime)
                modified_text = modified.strftime("%Y-%m-%d %H:%M")
            except OSError:
                pass
            header = pdf_path.name if not modified_text else f"{pdf_path.name} (modified {modified_text})"
            ttk.Label(item_frame, text=header, font=("TkDefaultFont", 10, "bold")).pack(anchor="w")
            try:
                doc = fitz.open(pdf_path)
            except Exception as exc:
                ttk.Label(item_frame, text=f"Unable to open PDF: {exc}").pack(anchor="w", pady=(4, 0))
                continue
            try:
                photo = self.render_page(doc, 0, target_width=220)
            finally:
                doc.close()
            if photo is None:
                ttk.Label(item_frame, text="Preview unavailable").pack(anchor="w", pady=(4, 0))
            else:
                previews.append(photo)
                ttk.Label(item_frame, image=photo).pack(anchor="w", pady=(4, 0))

        window.preview_images = previews  # type: ignore[attr-defined]
        return window

    def _ensure_unique_path(self, target: Path) -> Path:
        if not target.exists():
            return target
        stem = target.stem
        suffix = target.suffix
        counter = 1
        while True:
            candidate = target.with_name(f"{stem}_{counter}{suffix}")
            if not candidate.exists():
                return candidate
            counter += 1

    def _open_in_file_manager(self, path: Path) -> None:
        try:
            if os.name == "nt":
                os.startfile(path)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(path)])
            else:
                subprocess.Popen(["xdg-open", str(path)])
        except Exception:
            messagebox.showwarning("Open Folder", "Could not open the folder in the file manager.")
