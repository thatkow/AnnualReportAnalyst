"""Standalone analyst application for combined CSV summaries.

This lightweight Tkinter application lets reviewers choose companies that have
``combined.csv`` files in ``companies/<company>/``. When a company is opened, it
is added to a shared set of side-by-side Finance and Income Statement charts so
reviewers can compare multiple companies without switching tabs. Each bar is
stacked by the underlying rows from the ``combined.csv`` file so reviewers can
see how individual entries contribute to each reporting period.
"""

from __future__ import annotations

import csv
import math
import re
from collections import OrderedDict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Sequence, Tuple

import json
import os
import sys
import webbrowser
from datetime import datetime

import matplotlib

try:
    import fitz  # type: ignore[import-untyped]
except ImportError:  # pragma: no cover - handled via runtime warning
    fitz = None  # type: ignore[assignment]

from PIL import Image, ImageTk

# Ensure TkAgg is used when embedding plots inside Tkinter widgets.
matplotlib.use("TkAgg")

from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from matplotlib import colors, colormaps
from matplotlib.patches import Patch, Rectangle
from matplotlib.ticker import ScalarFormatter

import tkinter as tk
from tkinter import messagebox, ttk


BASE_FIELDS = {"Type", "Category", "Item", "Note"}
NOTE_ASIS = "asis"
NOTE_NEGATED = "negated"
NOTE_EXCLUDED = "excluded"
NOTE_SHARE_COUNT = "share_count"
NOTE_SHARE_PRICE = "share_price"
VALID_NOTES = {
    NOTE_ASIS,
    NOTE_NEGATED,
    NOTE_EXCLUDED,
    NOTE_SHARE_COUNT,
    NOTE_SHARE_PRICE,
}


@dataclass
class RowSegment:
    """Represent a single CSV row broken into period values for plotting."""

    key: str
    type_label: str
    type_value: str
    category: str
    item: str
    periods: List[str]
    values: List[float]
    sources: Dict[str, "SegmentSource"] = field(default_factory=dict)
    raw_values: List[float] = field(default_factory=list)


@dataclass
class SegmentSource:
    """Describe the PDF resource tied to a row segment period."""

    pdf_path: Path
    page: Optional[int]


def open_pdf(pdf_path: Path, page: Optional[int] = None) -> None:
    """Open a PDF, attempting to jump to the requested page when available."""

    try:
        if page is not None and page >= 1:
            url = pdf_path.resolve().as_uri() + f"#page={page}"
            if webbrowser.open(url):
                return
        if sys.platform.startswith("win"):
            os.startfile(pdf_path)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            os.system(f"open '{pdf_path}'")
        else:
            os.system(f"xdg-open '{pdf_path}' >/dev/null 2>&1 &")
    except Exception as exc:
        messagebox.showwarning("Open PDF", f"Could not open PDF: {exc}")


class PdfPageViewer(tk.Toplevel):
    """Display a single rendered PDF page inside a scrollable window."""

    def __init__(self, master: tk.Widget, pdf_path: Path, page_number: int) -> None:
        super().__init__(master)
        if fitz is None:
            messagebox.showerror(
                "PyMuPDF Required",
                (
                    "Viewing a single PDF page requires the PyMuPDF package (import name 'fitz').\n"
                    "Install it with 'pip install PyMuPDF' to enable in-app previews."
                ),
            )
            self.after(0, self.destroy)
            raise RuntimeError("PyMuPDF (fitz) is not installed")
        self.title(f"{pdf_path.name} — Page {page_number}")
        self._photo: Optional[ImageTk.PhotoImage] = None
        self._pdf_path = pdf_path
        self._page_number = page_number
        self._build_widgets()
        self._render_page()
        self.bind("<Escape>", lambda _event: self.destroy())

    def _build_widgets(self) -> None:
        self.configure(bg="#2b2b2b")
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self._canvas = tk.Canvas(self, bg="#2b2b2b", highlightthickness=0)
        self._canvas.grid(row=0, column=0, sticky="nsew")

        self._vbar = tk.Scrollbar(self, orient=tk.VERTICAL, command=self._canvas.yview)
        self._vbar.grid(row=0, column=1, sticky="ns")
        self._hbar = tk.Scrollbar(self, orient=tk.HORIZONTAL, command=self._canvas.xview)
        self._hbar.grid(row=1, column=0, sticky="ew")

        self._canvas.configure(yscrollcommand=self._vbar.set, xscrollcommand=self._hbar.set)

    def _render_page(self) -> None:
        if fitz is None:
            raise RuntimeError("PyMuPDF (fitz) is not installed")
        try:
            with fitz.open(self._pdf_path) as doc:
                page = doc.load_page(self._page_number - 1)
                zoom_matrix = fitz.Matrix(1.5, 1.5)
                pix = page.get_pixmap(matrix=zoom_matrix)
        except Exception as exc:
            messagebox.showwarning(
                "Open PDF",
                f"Could not render page {self._page_number} from {self._pdf_path.name}: {exc}",
            )
            self.after(0, self.destroy)
            return

        mode = "RGBA" if pix.alpha else "RGB"
        image = Image.frombytes(mode, [pix.width, pix.height], pix.samples)
        self._photo = ImageTk.PhotoImage(image)
        self._canvas.delete("all")
        self._canvas.create_image(0, 0, anchor="nw", image=self._photo)
        self._canvas.configure(scrollregion=(0, 0, pix.width, pix.height))


class FinanceDataset:
    """Load and aggregate Finance and Income Statement totals from a CSV."""

    FINANCE_LABEL = "Finance"
    INCOME_LABEL = "Income Statement"

    NORMALIZATION_SHARES = "shares"
    NORMALIZATION_REPORTED = "reported"
    NORMALIZATION_SHARE_PRICE = "share_price"

    _TYPE_TO_LABEL = {
        "financial": FINANCE_LABEL,
        "finance": FINANCE_LABEL,
        "income": INCOME_LABEL,
    }

    def __init__(self, path: Path) -> None:
        self.path = path
        self.company_root = path.parent
        self.headers: List[str] = []
        self.periods: List[str] = []
        self.finance_segments: List[RowSegment] = []
        self.income_segments: List[RowSegment] = []
        self._color_cache: Dict[str, str] = {}
        self.share_counts: List[float] = []
        self.share_prices: List[float] = []
        self._normalization_mode = self.NORMALIZATION_REPORTED
        self._load()
        self._load_metadata()

    @staticmethod
    def _clean_numeric(value: str) -> float:
        text = value.strip()
        if not text or text in {"-", "--"}:
            return 0.0
        negative = False
        if text.startswith("(") and text.endswith(")"):
            negative = True
            text = text[1:-1].strip()
        cleaned = text.replace(",", "")
        cleaned = re.sub(r"[^0-9eE\.\+\-]", "", cleaned)
        if not cleaned:
            return 0.0
        try:
            number = float(cleaned)
        except ValueError:
            return 0.0
        if negative and number > 0:
            number = -number
        return number

    def _load(self) -> None:
        if not self.path.exists():
            raise FileNotFoundError(f"Combined CSV not found: {self.path}")

        with self.path.open(newline="", encoding="utf-8") as csv_file:
            reader = csv.reader(csv_file)
            try:
                header = next(reader)
            except StopIteration as exc:
                raise ValueError("Combined CSV is empty") from exc

            self.headers = header
            type_index = header.index("Type") if "Type" in header else None
            if type_index is None:
                raise ValueError("Combined CSV missing 'Type' column")

            category_index = header.index("Category") if "Category" in header else None
            item_index = header.index("Item") if "Item" in header else None
            note_index = header.index("Note") if "Note" in header else None
            if note_index is None:
                raise ValueError("Combined CSV missing 'Note' column")

            data_indices: List[int] = []
            self.periods = []
            for idx, column_name in enumerate(header):
                if column_name in BASE_FIELDS:
                    continue
                data_indices.append(idx)
                self.periods.append(column_name)

            if not data_indices:
                raise ValueError("Combined CSV does not contain numeric data columns")
            self.finance_segments = []
            self.income_segments = []
            share_counts: Optional[List[float]] = None
            share_prices: Optional[List[float]] = None

            for row_number, row in enumerate(reader, start=1):
                if note_index >= len(row):
                    raise ValueError(f"Row {row_number} missing Note value")
                note_value = row[note_index].strip()
                if not note_value:
                    raise ValueError(f"Row {row_number} missing Note value")
                note_key = note_value.lower()
                if note_key not in VALID_NOTES:
                    raise ValueError(
                        f"Row {row_number} has unsupported Note value '{note_value}'"
                    )

                values: List[float] = []
                for column_index in data_indices:
                    if column_index >= len(row):
                        raise ValueError(
                            f"Row {row_number} is missing data for period column index {column_index}"
                        )
                    values.append(self._clean_numeric(row[column_index]))

                if note_key == NOTE_SHARE_COUNT:
                    if share_counts is not None:
                        raise ValueError("Multiple share_count rows found in combined.csv")
                    share_counts = values
                    continue

                if note_key == NOTE_SHARE_PRICE:
                    if share_prices is not None:
                        raise ValueError("Multiple share_price rows found in combined.csv")
                    share_prices = values
                    continue

                if not any(value != 0 for value in values):
                    continue

                if note_key == NOTE_EXCLUDED:
                    continue

                if type_index >= len(row):
                    raise ValueError(f"Row {row_number} missing Type value")

                raw_type = row[type_index].strip()
                type_value = raw_type.lower()
                series_label = self._TYPE_TO_LABEL.get(type_value)
                if not series_label:
                    # Skip other sections such as Shares.
                    continue

                if note_key == NOTE_NEGATED:
                    values = [-value for value in values]

                category = (
                    row[category_index].strip()
                    if category_index is not None and category_index < len(row)
                    else ""
                )
                item = (
                    row[item_index].strip()
                    if item_index is not None and item_index < len(row)
                    else ""
                )
                key_parts: List[str] = []
                if raw_type:
                    key_parts.append(raw_type)
                if category:
                    key_parts.append(category)
                if item:
                    key_parts.append(item)
                key = "_".join(key_parts)
                if not key:
                    key = f"{series_label} Row {row_number}"

                segment = RowSegment(
                    key=key,
                    type_label=series_label,
                    type_value=raw_type,
                    category=category,
                    item=item,
                    periods=list(self.periods),
                    values=values,
                )
                target = (
                    self.finance_segments
                    if series_label == self.FINANCE_LABEL
                    else self.income_segments
                )
                target.append(segment)

            if share_counts is None:
                raise ValueError("combined.csv is missing a share_count row")

            if len(share_counts) != len(self.periods):
                raise ValueError("share_count row does not match the number of periods")

            for index, count in enumerate(share_counts):
                if count == 0:
                    raise ValueError(
                        f"share_count value for period '{self.periods[index]}' is zero"
                    )

            if share_prices is not None and len(share_prices) != len(self.periods):
                raise ValueError("share_price row does not match the number of periods")

            sort_indices = list(range(len(self.periods)))
            try:
                parsed_periods = [
                    datetime.strptime(self.periods[idx].strip(), "%d.%m.%Y")
                    for idx in sort_indices
                ]
            except ValueError:
                parsed_periods = []

            if parsed_periods:
                sort_indices = [
                    index
                    for _, index in sorted(
                        zip(parsed_periods, sort_indices), key=lambda pair: pair[0]
                    )
                ]
                self.periods = [self.periods[idx] for idx in sort_indices]
                share_counts = [share_counts[idx] for idx in sort_indices]
                if share_prices is not None:
                    share_prices = [share_prices[idx] for idx in sort_indices]
            for segment in self.finance_segments + self.income_segments:
                segment.values = [segment.values[idx] for idx in sort_indices]
                segment.periods = [self.periods[idx] for idx in sort_indices]
                segment.raw_values = list(segment.values)

            self.share_counts = list(share_counts)
            self.share_prices = list(share_prices) if share_prices is not None else []
            self.set_normalization_mode(self.NORMALIZATION_SHARES)

    def set_normalization_mode(self, mode: str) -> None:
        valid_modes = {
            self.NORMALIZATION_SHARES,
            self.NORMALIZATION_REPORTED,
            self.NORMALIZATION_SHARE_PRICE,
        }
        if mode not in valid_modes:
            raise ValueError(f"Unsupported normalization mode: {mode}")

        if mode == self._normalization_mode:
            return

        if mode == self.NORMALIZATION_SHARES:
            if not self.share_counts:
                raise ValueError("combined.csv is missing a share_count row")
            for segment in self.finance_segments + self.income_segments:
                per_share: List[float] = []
                for idx, value in enumerate(segment.raw_values):
                    count = self.share_counts[idx] if idx < len(self.share_counts) else 0.0
                    if count == 0:
                        per_share.append(0.0)
                    else:
                        per_share.append(value / count)
                segment.values = per_share
        elif mode == self.NORMALIZATION_SHARE_PRICE:
            if not self.share_counts:
                raise ValueError("combined.csv is missing a share_count row")
            latest_price = self.latest_share_price()
            if latest_price is None:
                raise ValueError(
                    f"Share price data is not available for {self.company_root.name or 'the selected company'}."
                )
            for segment in self.finance_segments + self.income_segments:
                multiples: List[float] = []
                for idx, value in enumerate(segment.raw_values):
                    count = self.share_counts[idx] if idx < len(self.share_counts) else 0.0
                    if count == 0:
                        multiples.append(0.0)
                        continue
                    per_share = value / count
                    if per_share == 0:
                        multiples.append(0.0)
                    else:
                        multiples.append(latest_price / per_share)
                segment.values = multiples
        else:
            for segment in self.finance_segments + self.income_segments:
                segment.values = list(segment.raw_values)
        self._normalization_mode = mode

    def _load_metadata(self) -> None:
        metadata_path = self.path.with_name("combined_metadata.json")
        if not metadata_path.exists():
            return
        try:
            with metadata_path.open("r", encoding="utf-8") as fh:
                payload = json.load(fh)
        except (OSError, json.JSONDecodeError):
            return

        rows = payload.get("rows")
        if not isinstance(rows, list):
            return

        metadata_map: Dict[Tuple[str, str, str], Dict[str, SegmentSource]] = {}

        for entry in rows:
            if not isinstance(entry, dict):
                continue
            type_value = str(entry.get("type", ""))
            category_value = str(entry.get("category", ""))
            item_value = str(entry.get("item", ""))
            key = (type_value, category_value, item_value)
            periods = entry.get("periods", {})
            if not isinstance(periods, dict):
                continue
            period_sources: Dict[str, SegmentSource] = {}
            for label, info in periods.items():
                if not isinstance(info, dict):
                    continue
                pdf_value = info.get("pdf")
                if not pdf_value:
                    continue
                pdf_path = Path(pdf_value)
                if not pdf_path.is_absolute():
                    pdf_path = (metadata_path.parent / pdf_path).resolve()
                page_value = info.get("page")
                page: Optional[int]
                if isinstance(page_value, int):
                    page = page_value if page_value > 0 else None
                else:
                    try:
                        page = int(str(page_value))
                        if page <= 0:
                            page = None
                    except (TypeError, ValueError):
                        page = None
                period_sources[str(label)] = SegmentSource(pdf_path=pdf_path, page=page)
            if period_sources:
                metadata_map[key] = period_sources

        if not metadata_map:
            return

        for segment in self.finance_segments + self.income_segments:
            key = (segment.type_value, segment.category, segment.item)
            sources = metadata_map.get(key)
            if sources:
                segment.sources = sources

    def aggregate_totals(
        self, periods: Sequence[str], *, series: str
    ) -> Tuple[List[float], List[bool]]:
        if series == self.FINANCE_LABEL:
            segments = self.finance_segments
        elif series == self.INCOME_LABEL:
            segments = self.income_segments
        else:
            return [0.0] * len(periods), [False] * len(periods)

        period_index = {label: idx for idx, label in enumerate(self.periods)}
        totals: List[float] = []
        has_data: List[bool] = []

        for label in periods:
            idx = period_index.get(label)
            if idx is None:
                totals.append(0.0)
                has_data.append(False)
                continue

            total = 0.0
            present = False
            for segment in segments:
                if idx >= len(segment.values):
                    continue
                present = True
                total += segment.values[idx]

            totals.append(total)
            has_data.append(present)

        return totals, has_data

    def share_counts_for(self, periods: Sequence[str]) -> Tuple[List[float], List[bool]]:
        if not self.share_counts:
            return [math.nan] * len(periods), [False] * len(periods)

        period_index = {label: idx for idx, label in enumerate(self.periods)}
        values: List[float] = []
        present: List[bool] = []

        for label in periods:
            idx = period_index.get(label)
            if idx is None or idx >= len(self.share_counts):
                values.append(math.nan)
                present.append(False)
                continue
            values.append(self.share_counts[idx])
            present.append(True)

        return values, present

    def latest_share_price(self) -> Optional[float]:
        for value in reversed(self.share_prices):
            if math.isfinite(value) and value > 0:
                return value
        return None

    def has_data(self) -> bool:
        return bool(self.finance_segments or self.income_segments)

    def color_for_key(self, key: str) -> str:
        if key not in self._color_cache:
            palette = colormaps["tab20"].colors
            index = len(self._color_cache) % len(palette)
            self._color_cache[key] = colors.to_hex(palette[index])
        return self._color_cache[key]


class FinancePlotFrame(ttk.Frame):
    """Container that renders stacked bar plots for one or more companies."""

    _PLOT_WIDTH = 0.82
    MODE_STACKED = "stacked"
    MODE_LINE = "line"
    VALUE_MODE_SHARE_COUNT = "share_count"

    def __init__(self, parent: tk.Widget) -> None:
        super().__init__(parent)
        self.figure = Figure(figsize=(8, 5), dpi=100)
        self.axis = self.figure.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.figure, master=self)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        self.hover_helper = BarHoverHelper(self.axis)
        self.hover_helper.attach(self.canvas)
        self.display_mode = self.MODE_STACKED
        self.datasets: "OrderedDict[str, FinanceDataset]" = OrderedDict()
        self.periods: Optional[List[str]] = None
        self._periods_by_key: "OrderedDict[Tuple[str, str, str], List[str]]" = OrderedDict()
        self._visible_periods: Optional[List[str]] = None
        self._company_colors: Dict[str, Tuple[str, str]] = {}
        self._palette = list(colormaps["tab10"].colors)
        self._context_menu: Optional[tk.Menu] = None
        self._context_metadata: Optional[Dict[str, Any]] = None
        self.normalization_mode = FinanceDataset.NORMALIZATION_SHARES
        self._dataset_normalization_mode = FinanceDataset.NORMALIZATION_SHARES
        self._per_share_formatter = ScalarFormatter()
        self._reported_formatter = ScalarFormatter(useOffset=False)
        self._reported_formatter.set_scientific(True)
        self._reported_formatter.set_powerlimits((0, 0))
        self._share_count_formatter = ScalarFormatter(useOffset=False)
        self._share_count_formatter.set_scientific(True)
        self._share_count_formatter.set_powerlimits((0, 0))
        self._share_price_formatter = ScalarFormatter(useOffset=False)
        self._share_price_formatter.set_scientific(True)
        self._share_price_formatter.set_powerlimits((0, 0))
        self._render_empty()

    @staticmethod
    def _compute_bottoms(values: Sequence[float], pos_totals: List[float], neg_totals: List[float]) -> List[float]:
        bottoms: List[float] = []
        for index, value in enumerate(values):
            if value >= 0:
                bottom = pos_totals[index]
                pos_totals[index] += value
            else:
                bottom = neg_totals[index]
                neg_totals[index] += value
            bottoms.append(bottom)
        return bottoms

    def _render_empty(self) -> None:
        self.axis.clear()
        self.axis.set_title("Select a company to begin analysis")
        self.axis.set_xticks([])
        self.axis.set_yticks([])
        self.axis.axhline(0, color="#333333", linewidth=0.8)
        context_callback = (
            self._show_context_menu if self.display_mode == self.MODE_STACKED else None
        )
        self.hover_helper.begin_update(context_callback)
        self.figure.tight_layout()
        self.canvas.draw()

    def set_display_mode(self, mode: str) -> None:
        if mode not in {self.MODE_STACKED, self.MODE_LINE}:
            return
        if self.display_mode == mode:
            return
        self.display_mode = mode
        if self.datasets:
            self._render()
        else:
            self._render_empty()

    def set_normalization_mode(self, mode: str) -> bool:
        valid_modes = {
            FinanceDataset.NORMALIZATION_SHARES,
            FinanceDataset.NORMALIZATION_REPORTED,
            FinanceDataset.NORMALIZATION_SHARE_PRICE,
            self.VALUE_MODE_SHARE_COUNT,
        }
        if mode not in valid_modes:
            return False

        if mode == self.VALUE_MODE_SHARE_COUNT:
            if self.normalization_mode == mode:
                return True
            self.normalization_mode = mode
            if self.datasets:
                self._render()
            else:
                self._render_empty()
            return True

        if self._dataset_normalization_mode != mode:
            previous_mode = self._dataset_normalization_mode
            updated: List[FinanceDataset] = []
            for dataset in self.datasets.values():
                try:
                    dataset.set_normalization_mode(mode)
                except ValueError as exc:
                    messagebox.showwarning("Normalize Values", str(exc))
                    for updated_dataset in updated:
                        try:
                            updated_dataset.set_normalization_mode(previous_mode)
                        except ValueError:
                            pass
                    return False
                else:
                    updated.append(dataset)
            self._dataset_normalization_mode = mode

        if self.normalization_mode == mode:
            return True

        self.normalization_mode = mode
        if self.datasets:
            self._render()
        else:
            self._render_empty()
        return True

    def _next_color_pair(self) -> Tuple[str, str]:
        index = len(self._company_colors)
        if not self._palette:
            return ("#1f77b4", "#ff7f0e")
        finance_index = (2 * index) % len(self._palette)
        income_index = (finance_index + 1) % len(self._palette)
        finance_color = colors.to_hex(self._palette[finance_index])
        income_color = colors.to_hex(self._palette[income_index])
        return finance_color, income_color

    @staticmethod
    def _sort_periods(periods: Sequence[str]) -> List[str]:
        def sort_key(label: str) -> Tuple[int, Any]:
            try:
                parsed = datetime.strptime(label.strip(), "%d.%m.%Y")
                return (0, parsed)
            except ValueError:
                return (1, label)

        return sorted(periods, key=sort_key)

    @staticmethod
    def _merge_periods(existing: Sequence[str], new_periods: Sequence[str]) -> List[str]:
        ordered_unique: List[str] = []
        seen = set()
        for label in list(existing) + list(new_periods):
            if label in seen:
                continue
            seen.add(label)
            ordered_unique.append(label)
        return FinancePlotFrame._sort_periods(ordered_unique)

    def _register_periods(self, dataset: FinanceDataset) -> None:
        for segment in dataset.finance_segments + dataset.income_segments:
            key_tuple = (segment.type_value, segment.category, segment.item)
            existing = self._periods_by_key.get(key_tuple, [])
            merged = self._merge_periods(existing, segment.periods)
            self._periods_by_key[key_tuple] = merged

    def _rebuild_period_sequence(self) -> List[str]:
        if not self._periods_by_key:
            return []
        ordered: List[str] = []
        seen = set()
        for period_list in self._periods_by_key.values():
            for label in period_list:
                if label in seen:
                    continue
                seen.add(label)
                ordered.append(label)
        return self._sort_periods(ordered)

    def clear_companies(self) -> None:
        self.datasets.clear()
        self.periods = None
        self._periods_by_key.clear()
        self._visible_periods = None
        self._render_empty()

    def add_company(self, company: str, dataset: FinanceDataset) -> Tuple[bool, Optional[str]]:
        previous_periods = set(self.periods or [])
        self._register_periods(dataset)
        previous_visible = (
            list(self._visible_periods) if self._visible_periods is not None else None
        )
        self.periods = self._rebuild_period_sequence()
        new_periods = [label for label in self.periods if label not in previous_periods]
        if previous_visible is None:
            self._visible_periods = list(self.periods)
        else:
            visible_set = set(previous_visible)
            updated_visible = [label for label in self.periods if label in visible_set]
            updated_visible.extend(
                [label for label in new_periods if label not in visible_set]
            )
            self._visible_periods = updated_visible

        is_new_company = company not in self.datasets
        if is_new_company and company not in self._company_colors:
            self._company_colors[company] = self._next_color_pair()

        normalization_warning: Optional[str] = None
        try:
            dataset.set_normalization_mode(self._dataset_normalization_mode)
        except ValueError as exc:
            if self._dataset_normalization_mode == FinanceDataset.NORMALIZATION_SHARE_PRICE:
                normalization_warning = str(exc)
                fallback_mode = FinanceDataset.NORMALIZATION_SHARES
                dataset.set_normalization_mode(fallback_mode)
                for existing in self.datasets.values():
                    existing.set_normalization_mode(fallback_mode)
                self._dataset_normalization_mode = fallback_mode
                if self.normalization_mode == FinanceDataset.NORMALIZATION_SHARE_PRICE:
                    self.normalization_mode = fallback_mode
            else:
                raise

        self.datasets[company] = dataset
        # Preserve insertion order by moving refreshed companies to the end.
        self.datasets.move_to_end(company)
        self._render()
        return is_new_company, normalization_warning

    def all_periods(self) -> List[str]:
        return list(self.periods or [])

    def visible_periods(self) -> List[str]:
        if self._visible_periods is not None:
            return list(self._visible_periods)
        return list(self.periods or [])

    def set_visible_periods(self, periods: Sequence[str]) -> None:
        if not self.periods:
            self._visible_periods = []
            self._render_empty()
            return

        allowed = {label for label in periods}
        new_visible = [label for label in self.periods if label in allowed]
        if not allowed:
            new_visible = []

        if self._visible_periods == new_visible:
            return

        self._visible_periods = new_visible
        if self.datasets:
            self._render()
        else:
            self._render_empty()

    def _render(self) -> None:
        if not self.datasets:
            self._render_empty()
            return

        if self._visible_periods is not None:
            periods = list(self._visible_periods)
        else:
            periods = self.periods or []
        period_indices = list(range(len(periods)))
        num_companies = len(self.datasets)
        if num_companies == 0:
            self._render_empty()
            return

        self.axis.clear()
        hover_helper = self.hover_helper
        mode = self.display_mode

        legend_handles: List[Any] = []

        if self.normalization_mode == self.VALUE_MODE_SHARE_COUNT:
            hover_helper.begin_update(None)
            x_positions = period_indices
            for company, dataset in self.datasets.items():
                finance_color, _ = self._company_colors.setdefault(
                    company, self._next_color_pair()
                )
                share_values, share_presence = dataset.share_counts_for(periods)
                if not any(share_presence):
                    continue
                share_points = [
                    share_values[idx] if share_presence[idx] else math.nan
                    for idx in range(len(share_values))
                ]
                line = self.axis.plot(
                    x_positions,
                    share_points,
                    linestyle=":",
                    marker="o",
                    color=finance_color,
                    markerfacecolor=finance_color,
                    markeredgecolor=finance_color,
                    label=f"{company} Number of shares",
                )
                legend_handles.append(line[0])

            if period_indices:
                self.axis.set_xticks(period_indices)
                self.axis.set_xticklabels(periods, rotation=45, ha="right")
                self.axis.set_xlim(-0.5, len(period_indices) - 0.5)
            else:
                self.axis.set_xlim(-0.6, 0.6)
                self.axis.set_xticks([])

        elif mode == self.MODE_STACKED:
            hover_helper.begin_update(self._show_context_menu)
            group_width = self._PLOT_WIDTH
            bar_width = group_width / max(2 * num_companies, 1)
            start_offset = -group_width / 2

            for company_index, (company, dataset) in enumerate(self.datasets.items()):
                finance_color, income_color = self._company_colors.setdefault(
                    company, self._next_color_pair()
                )
                company_color = finance_color
                base_offset = start_offset + (2 * company_index) * bar_width
                finance_positions = [
                    index + base_offset + 0.5 * bar_width for index in period_indices
                ]
                income_positions = [
                    index + base_offset + 1.5 * bar_width for index in period_indices
                ]

                finance_segment_values: List[Tuple[RowSegment, List[float]]] = []
                finance_nonzero = [False] * len(periods)
                for segment in dataset.finance_segments:
                    segment_period_index = {
                        label: idx for idx, label in enumerate(segment.periods)
                    }
                    segment_values = [
                        segment.values[segment_period_index[label]]
                        if label in segment_period_index
                        and segment_period_index[label] < len(segment.values)
                        else 0.0
                        for label in periods
                    ]
                    for idx, value in enumerate(segment_values):
                        if value != 0:
                            finance_nonzero[idx] = True
                    finance_segment_values.append((segment, segment_values))

                finance_active_indices = [
                    idx for idx, flag in enumerate(finance_nonzero) if flag
                ]
                finance_positions_active = [
                    finance_positions[idx] for idx in finance_active_indices
                ]
                finance_periods_active = [
                    periods[idx] for idx in finance_active_indices
                ]
                finance_pos_totals = [0.0] * len(finance_active_indices)
                finance_neg_totals = [0.0] * len(finance_active_indices)

                if finance_active_indices:
                    for segment, segment_values in finance_segment_values:
                        filtered_values = [
                            segment_values[idx] for idx in finance_active_indices
                        ]
                        if not any(value != 0 for value in filtered_values):
                            continue
                        bottoms = self._compute_bottoms(
                            filtered_values, finance_pos_totals, finance_neg_totals
                        )
                        rectangles = self.axis.bar(
                            finance_positions_active,
                            filtered_values,
                            width=bar_width,
                            bottom=bottoms,
                            color=dataset.color_for_key(segment.key),
                            edgecolor=finance_color,
                            linewidth=2.4,
                        )
                        hover_helper.add_segment(
                            rectangles,
                            segment,
                            finance_periods_active,
                            company,
                            values=filtered_values,
                        )

                income_segment_values: List[Tuple[RowSegment, List[float]]] = []
                income_nonzero = [False] * len(periods)
                for segment in dataset.income_segments:
                    segment_period_index = {
                        label: idx for idx, label in enumerate(segment.periods)
                    }
                    segment_values = [
                        segment.values[segment_period_index[label]]
                        if label in segment_period_index
                        and segment_period_index[label] < len(segment.values)
                        else 0.0
                        for label in periods
                    ]
                    for idx, value in enumerate(segment_values):
                        if value != 0:
                            income_nonzero[idx] = True
                    income_segment_values.append((segment, segment_values))

                income_active_indices = [
                    idx for idx, flag in enumerate(income_nonzero) if flag
                ]
                income_positions_active = [
                    income_positions[idx] for idx in income_active_indices
                ]
                income_periods_active = [
                    periods[idx] for idx in income_active_indices
                ]
                income_pos_totals = [0.0] * len(income_active_indices)
                income_neg_totals = [0.0] * len(income_active_indices)

                if income_active_indices:
                    for segment, segment_values in income_segment_values:
                        filtered_values = [
                            segment_values[idx] for idx in income_active_indices
                        ]
                        if not any(value != 0 for value in filtered_values):
                            continue
                        bottoms = self._compute_bottoms(
                            filtered_values, income_pos_totals, income_neg_totals
                        )
                        rectangles = self.axis.bar(
                            income_positions_active,
                            filtered_values,
                            width=bar_width,
                            bottom=bottoms,
                            color=dataset.color_for_key(segment.key),
                            edgecolor=income_color,
                            linewidth=2.4,
                        )
                        hover_helper.add_segment(
                            rectangles,
                            segment,
                            income_periods_active,
                            company,
                            values=filtered_values,
                        )

                if finance_active_indices:
                    finance_totals = [
                        pos + neg for pos, neg in zip(finance_pos_totals, finance_neg_totals)
                    ]
                    dot_color = finance_color
                    self.axis.scatter(
                        finance_positions_active,
                        finance_totals,
                        color=dot_color,
                        marker="o",
                        s=36,
                        zorder=5,
                    )
                    legend_handles.append(
                        Patch(
                            facecolor="white",
                            edgecolor=finance_color,
                            linewidth=1.0,
                            label=f"{company} {FinanceDataset.FINANCE_LABEL}",
                        )
                    )

                if income_active_indices:
                    income_totals = [
                        pos + neg for pos, neg in zip(income_pos_totals, income_neg_totals)
                    ]
                    dot_color = finance_color
                    self.axis.scatter(
                        income_positions_active,
                        income_totals,
                        color=dot_color,
                        marker="o",
                        s=36,
                        zorder=5,
                    )
                    legend_handles.append(
                        Patch(
                            facecolor="white",
                            edgecolor=income_color,
                            linewidth=1.0,
                            label=f"{company} {FinanceDataset.INCOME_LABEL}",
                        )
                    )

            if period_indices:
                left_limit = period_indices[0] + start_offset
                right_limit = (
                    period_indices[-1]
                    + start_offset
                    + 2 * num_companies * bar_width
                )
                padding = bar_width
                self.axis.set_xlim(left_limit - padding, right_limit + padding)
                self.axis.set_xticks(period_indices)
                self.axis.set_xticklabels(periods, rotation=45, ha="right")
            else:
                self.axis.set_xlim(-0.6, 0.6)
                self.axis.set_xticks([])

        else:
            hover_helper.begin_update(None)
            x_positions = period_indices
            for company, dataset in self.datasets.items():
                finance_color, income_color = self._company_colors.setdefault(
                    company, self._next_color_pair()
                )
                company_color = finance_color

                finance_totals, finance_presence = dataset.aggregate_totals(
                    periods, series=FinanceDataset.FINANCE_LABEL
                )
                finance_included = [
                    finance_presence[idx]
                    and not math.isclose(value, 0.0, abs_tol=1e-12)
                    for idx, value in enumerate(finance_totals)
                ]
                finance_points = [
                    value if finance_included[idx] else math.nan
                    for idx, value in enumerate(finance_totals)
                ]
                if any(finance_included):
                    line = self.axis.plot(
                        x_positions,
                        finance_points,
                        linestyle="-",
                        marker="o",
                        color=company_color,
                        markerfacecolor=company_color,
                        markeredgecolor=company_color,
                        label=f"{company} {FinanceDataset.FINANCE_LABEL}",
                    )
                    legend_handles.append(line[0])

                income_totals, income_presence = dataset.aggregate_totals(
                    periods, series=FinanceDataset.INCOME_LABEL
                )
                income_included = [
                    income_presence[idx]
                    and not math.isclose(value, 0.0, abs_tol=1e-12)
                    for idx, value in enumerate(income_totals)
                ]
                income_points = [
                    value if income_included[idx] else math.nan
                    for idx, value in enumerate(income_totals)
                ]
                if any(income_included):
                    line = self.axis.plot(
                        x_positions,
                        income_points,
                        linestyle="--",
                        marker="o",
                        color=company_color,
                        markerfacecolor=company_color,
                        markeredgecolor=company_color,
                        label=f"{company} {FinanceDataset.INCOME_LABEL}",
                    )
                    legend_handles.append(line[0])

            if period_indices:
                self.axis.set_xticks(period_indices)
                self.axis.set_xticklabels(periods, rotation=45, ha="right")
                if period_indices:
                    self.axis.set_xlim(-0.5, len(period_indices) - 0.5)
            else:
                self.axis.set_xlim(-0.6, 0.6)
                self.axis.set_xticks([])

        self.axis.axhline(0, color="#333333", linewidth=0.8)
        if self.normalization_mode == self.VALUE_MODE_SHARE_COUNT:
            self.axis.yaxis.set_major_formatter(self._share_count_formatter)
            self.axis.set_ylabel("Number of Shares")
        elif self.normalization_mode == FinanceDataset.NORMALIZATION_SHARES:
            self.axis.yaxis.set_major_formatter(self._per_share_formatter)
            self.axis.set_ylabel("Value per Share")
        elif self.normalization_mode == FinanceDataset.NORMALIZATION_SHARE_PRICE:
            self.axis.yaxis.set_major_formatter(self._share_price_formatter)
            self.axis.set_ylabel("Share Price Multiple")
        else:
            self.axis.yaxis.set_major_formatter(self._reported_formatter)
            self.axis.set_ylabel("Reported Value")
        if self.datasets:
            companies_list = ", ".join(self.datasets.keys())
            if self.normalization_mode == self.VALUE_MODE_SHARE_COUNT:
                self.axis.set_title(f"Number of Shares — {companies_list}")
            else:
                self.axis.set_title(f"Finance vs Income Statement — {companies_list}")
        else:
            if self.normalization_mode == self.VALUE_MODE_SHARE_COUNT:
                self.axis.set_title("Number of Shares")
            else:
                self.axis.set_title("Finance vs Income Statement")

        if legend_handles:
            legend_kwargs = {"loc": "upper left"}
            if mode == self.MODE_STACKED and self.normalization_mode != self.VALUE_MODE_SHARE_COUNT:
                legend_kwargs["ncol"] = 2
            self.axis.legend(handles=legend_handles, **legend_kwargs)

        self.figure.tight_layout()
        self.canvas.draw()

    def _ensure_context_menu(self) -> tk.Menu:
        if self._context_menu is None:
            self._context_menu = tk.Menu(self, tearoff=False)
        else:
            self._context_menu.delete(0, tk.END)
        return self._context_menu

    def _show_context_menu(self, metadata: Dict[str, Any], gui_event: Any) -> None:
        menu = self._ensure_context_menu()
        self._context_metadata = None
        period = metadata.get("period") or "Period"
        pdf_path = metadata.get("pdf_path")
        page = metadata.get("pdf_page")
        label = f"View PDF for {period}"
        if pdf_path:
            if page is None:
                label += " (page unknown)"
            self._context_metadata = metadata
            menu.add_command(label=label, command=self._open_pdf_from_context)
        else:
            menu.add_command(label=label, state=tk.DISABLED)
        try:
            if gui_event is not None and hasattr(gui_event, "x_root") and hasattr(gui_event, "y_root"):
                menu.tk_popup(gui_event.x_root, gui_event.y_root)
            else:
                menu.tk_popup(self.winfo_pointerx(), self.winfo_pointery())
        finally:
            menu.grab_release()

    def _open_pdf_from_context(self) -> None:
        if not self._context_metadata:
            return
        pdf_path_value = self._context_metadata.get("pdf_path")
        if not pdf_path_value:
            messagebox.showinfo("Open PDF", "No PDF is associated with this bar segment.")
            return
        pdf_path = Path(str(pdf_path_value))
        page_value = self._context_metadata.get("pdf_page")
        page: Optional[int]
        try:
            page = int(page_value) if page_value is not None else None
        except (TypeError, ValueError):
            page = None
        if page is not None and page <= 0:
            page = None
        if not pdf_path.exists():
            messagebox.showwarning(
                "Open PDF",
                f"The PDF for this bar segment could not be found at {pdf_path}.",
            )
            return
        if page is None:
            messagebox.showinfo(
                "Open PDF",
                "No page number is available for this entry. The PDF will open to its first page.",
            )
            open_pdf(pdf_path, None)
            self._context_metadata = None
            return
        if fitz is None:
            messagebox.showwarning(
                "Open PDF",
                (
                    "Viewing individual PDF pages requires the PyMuPDF package (import name 'fitz').\n"
                    "The full PDF will be opened instead."
                ),
            )
            open_pdf(pdf_path, page)
            self._context_metadata = None
            return
        try:
            viewer = PdfPageViewer(self, pdf_path, page)
            viewer.focus_set()
        except Exception as exc:
            messagebox.showwarning(
                "Open PDF",
                f"Could not display page {page} from {pdf_path.name}: {exc}\n"
                "The full PDF will be opened instead.",
            )
            open_pdf(pdf_path, page)
        self._context_metadata = None


class BarHoverHelper:
    """Manage hover annotations for bar segments."""

    def __init__(self, axis) -> None:
        self.axis = axis
        self._rectangles: List[Tuple[Rectangle, Dict[str, Any]]] = []
        self._annotation = axis.annotate(
            "",
            xy=(0, 0),
            xytext=(12, 12),
            textcoords="offset points",
            bbox={"boxstyle": "round", "fc": "#f9f9f9", "ec": "#555555", "alpha": 0.95},
            arrowprops={"arrowstyle": "->", "color": "#555555"},
        )
        # Ensure the tooltip annotation renders above all bar segments so it is always visible.
        self._annotation.set_zorder(1000)
        bbox_patch = self._annotation.get_bbox_patch()
        if bbox_patch is not None:
            bbox_patch.set_zorder(1000)
        if hasattr(self._annotation, "arrow_patch") and self._annotation.arrow_patch is not None:
            self._annotation.arrow_patch.set_zorder(1000)
        self._annotation.set_visible(False)
        self._canvas: Optional[FigureCanvasTkAgg] = None
        self._tk_widget: Optional[tk.Widget] = None
        self._connection_id: Optional[int] = None
        self._button_connection_id: Optional[int] = None
        self._motion_binding_id: Optional[str] = None
        self._leave_binding_id: Optional[str] = None
        self._active_rectangle: Optional[Rectangle] = None
        self._context_callback: Optional[Callable[[Dict[str, Any], Any], None]] = None

    def attach(self, canvas: FigureCanvasTkAgg) -> None:
        if self._connection_id is not None and self._canvas is not None:
            self._canvas.mpl_disconnect(self._connection_id)
        if self._button_connection_id is not None and self._canvas is not None:
            self._canvas.mpl_disconnect(self._button_connection_id)
        if self._tk_widget is not None:
            if self._motion_binding_id is not None:
                self._tk_widget.unbind("<Motion>", self._motion_binding_id)
            if self._leave_binding_id is not None:
                self._tk_widget.unbind("<Leave>", self._leave_binding_id)
        self._motion_binding_id = None
        self._leave_binding_id = None
        self._canvas = canvas
        self._tk_widget = canvas.get_tk_widget()
        self.reset()
        self._connection_id = canvas.mpl_connect("motion_notify_event", self._on_motion)
        self._button_connection_id = None
        self._context_callback = None
        if self._tk_widget is not None:
            self._motion_binding_id = self._tk_widget.bind("<Motion>", self._on_tk_motion, add="+")
            self._leave_binding_id = self._tk_widget.bind("<Leave>", self._on_tk_leave, add="+")

    def _set_context_callback(
        self, callback: Optional[Callable[[Dict[str, Any], Any], None]]
    ) -> None:
        if self._canvas is None:
            self._context_callback = callback
            return
        if callback is None:
            if self._button_connection_id is not None:
                self._canvas.mpl_disconnect(self._button_connection_id)
                self._button_connection_id = None
        else:
            if self._button_connection_id is None:
                self._button_connection_id = self._canvas.mpl_connect(
                    "button_press_event", self._on_button_press
                )
        self._context_callback = callback

    def begin_update(
        self, context_callback: Optional[Callable[[Dict[str, Any], Any], None]]
    ) -> None:
        self._set_context_callback(context_callback)
        self.reset()

    def reset(self) -> None:
        self._rectangles.clear()
        self._active_rectangle = None
        self._annotation.set_visible(False)

    def _on_tk_motion(self, tk_event: tk.Event) -> None:  # type: ignore[name-defined]
        if self._canvas is None:
            return
        widget = self._tk_widget if self._tk_widget is not None else getattr(tk_event, "widget", None)
        if widget is None:
            return
        width = widget.winfo_width()
        height = widget.winfo_height()
        if height <= 0 or width <= 0:
            return
        canvas_backend = getattr(self._canvas.figure, "canvas", None)
        if canvas_backend is None:
            return
        backend_width, backend_height = canvas_backend.get_width_height()
        if backend_width <= 0 or backend_height <= 0:
            backend_width, backend_height = width, height
        scale_x = backend_width / width if width else 1.0
        scale_y = backend_height / height if height else 1.0
        tk_canvas = getattr(canvas_backend, "_tkcanvas", None)
        if tk_canvas is not None:
            canvas_x = tk_canvas.canvasx(tk_event.x)
            canvas_y = tk_canvas.canvasy(tk_event.y)
        else:
            canvas_x = tk_event.x
            canvas_y = tk_event.y
        display_x = canvas_x * scale_x
        display_y = (backend_height - canvas_y) * scale_y
        self._handle_motion_at(display_x, display_y)

    def _on_tk_leave(self, _event: tk.Event) -> None:  # type: ignore[name-defined]
        self._hide_annotation()

    def add_segment(
        self,
        bar_container,
        segment: RowSegment,
        periods: Sequence[str],
        company: str,
        *,
        values: Optional[Sequence[float]] = None,
    ) -> None:
        for index, rect in enumerate(bar_container.patches):
            value_list = values if values is not None else segment.values
            metadata: Dict[str, Any] = {
                "key": segment.key,
                "type_label": segment.type_label,
                "type_value": segment.type_value,
                "category": segment.category,
                "item": segment.item,
                "period": periods[index] if index < len(periods) else str(index),
                "value": value_list[index] if index < len(value_list) else 0.0,
                "company": company,
            }
            period_label = metadata["period"]
            if period_label is not None and segment.sources:
                source = segment.sources.get(str(period_label))
                if source:
                    metadata["pdf_path"] = str(source.pdf_path)
                    metadata["pdf_page"] = source.page
            self._rectangles.append((rect, metadata))

    def clear(self) -> None:
        self.reset()
        if self._canvas is not None:
            self._canvas.draw_idle()

    def _format_text(self, metadata: Dict[str, Any]) -> str:
        value = metadata.get("value")
        if isinstance(value, (int, float)) and math.isfinite(value):
            formatted_value = f"{value:,.2f}"
        else:
            formatted_value = "—"
        type_value = metadata.get("type_value") or metadata.get("type_label") or "—"
        category = metadata.get("category") or "—"
        item = metadata.get("item") or "—"
        return (
            f"TYPE: {type_value}\n"
            f"CATEGORY: {category}\n"
            f"ITEM: {item}\n"
            f"AMOUNT: {formatted_value}"
        )

    def _position_annotation(self, x: float, y: float, value: float) -> None:
        xlim = self.axis.get_xlim()
        ylim = self.axis.get_ylim()
        x_range = xlim[1] - xlim[0] if xlim[1] != xlim[0] else 1.0
        y_range = ylim[1] - ylim[0] if ylim[1] != ylim[0] else 1.0

        offset_x = 12
        ha = "left"
        right_threshold = xlim[1] - 0.08 * x_range
        left_threshold = xlim[0] + 0.08 * x_range
        if x >= right_threshold:
            offset_x = -12
            ha = "right"
        elif x <= left_threshold:
            offset_x = 12
            ha = "left"

        if value >= 0:
            offset_y = 12
            va = "bottom"
            top_threshold = ylim[1] - 0.08 * y_range
            if y >= top_threshold:
                offset_y = -12
                va = "top"
        else:
            offset_y = -12
            va = "top"
            bottom_threshold = ylim[0] + 0.08 * y_range
            if y <= bottom_threshold:
                offset_y = 12
                va = "bottom"

        self._annotation.xy = (x, y)
        self._annotation.xytext = (offset_x, offset_y)
        self._annotation.set_ha(ha)
        self._annotation.set_va(va)

    def _hide_annotation(self) -> None:
        if not self._annotation.get_visible():
            return
        self._annotation.set_visible(False)
        self._active_rectangle = None
        if self._canvas is not None:
            self._canvas.draw_idle()

    def _handle_motion_at(self, display_x: float, display_y: float) -> None:
        if self._canvas is None:
            return
        if math.isnan(display_x) or math.isnan(display_y):
            self._hide_annotation()
            return
        if not self.axis.bbox.contains(display_x, display_y):
            self._hide_annotation()
            return

        for rect, metadata in self._rectangles:
            if not rect.get_visible():
                continue
            if not rect.contains_point((display_x, display_y)):
                continue
            if self._active_rectangle is not rect or not self._annotation.get_visible():
                self._annotation.set_text(self._format_text(metadata))
                self._annotation.set_visible(True)
            self._active_rectangle = rect
            center_x = rect.get_x() + rect.get_width() / 2
            value = float(metadata.get("value", 0.0))
            if value >= 0:
                anchor_y = rect.get_y() + rect.get_height()
            else:
                anchor_y = rect.get_y()
            self._position_annotation(center_x, anchor_y, value)
            self._canvas.draw_idle()
            return

        self._hide_annotation()

    def _on_motion(self, event) -> None:
        if self._canvas is None:
            return
        if event.inaxes not in (None, self.axis):
            self._hide_annotation()
            return
        if event.x is None or event.y is None:
            self._hide_annotation()
            return
        self._handle_motion_at(event.x, event.y)

    def _metadata_for_event(self, event) -> Optional[Dict[str, Any]]:
        if event.inaxes != self.axis:
            return None
        for rect, metadata in self._rectangles:
            contains, _ = rect.contains(event)
            if contains:
                return metadata
        return None

    def _on_button_press(self, event) -> None:
        if event.button != 3:
            return
        if self._context_callback is None:
            return
        metadata = self._metadata_for_event(event)
        if metadata is None:
            return
        self._context_callback(metadata, getattr(event, "guiEvent", None))


class FinanceAnalystApp:
    """Main window controller for the analyst application."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("analyst")
        self._maximize_window()
        self.base_dir = Path(__file__).resolve().parent
        self.companies_dir = self.base_dir / "companies"

        self.company_var = tk.StringVar()
        self.mode_var = tk.StringVar(value=FinancePlotFrame.MODE_STACKED)
        self.normalization_var = tk.StringVar(
            value=FinanceDataset.NORMALIZATION_SHARES
        )
        self.period_vars: Dict[str, tk.BooleanVar] = {}

        self._build_ui()
        self._refresh_company_list()

    def _build_ui(self) -> None:
        outer = ttk.Frame(self.root, padding=12)
        outer.pack(fill=tk.BOTH, expand=True)

        controls = ttk.Frame(outer)
        controls.pack(fill=tk.X, pady=(0, 12))

        ttk.Label(controls, text="Company:").pack(side=tk.LEFT)
        self.company_combo = ttk.Combobox(
            controls,
            textvariable=self.company_var,
            state="readonly",
            width=40,
        )
        self.company_combo.pack(side=tk.LEFT, padx=(6, 6))

        self.open_button = ttk.Button(controls, text="Add", command=self._open_selected_company)
        self.open_button.pack(side=tk.LEFT)

        mode_frame = ttk.Frame(outer)
        mode_frame.pack(fill=tk.X, pady=(0, 12))

        ttk.Label(mode_frame, text="View:").pack(side=tk.LEFT)
        ttk.Radiobutton(
            mode_frame,
            text="Stacked bar components",
            variable=self.mode_var,
            value=FinancePlotFrame.MODE_STACKED,
            command=self._on_mode_change,
        ).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Radiobutton(
            mode_frame,
            text="Dotted Line",
            variable=self.mode_var,
            value=FinancePlotFrame.MODE_LINE,
            command=self._on_mode_change,
        ).pack(side=tk.LEFT, padx=(6, 0))

        values_frame = ttk.Frame(outer)
        values_frame.pack(fill=tk.X, pady=(0, 12))
        ttk.Label(values_frame, text="Values:").pack(side=tk.LEFT)
        ttk.Radiobutton(
            values_frame,
            text="Normalize over number of shares",
            variable=self.normalization_var,
            value=FinanceDataset.NORMALIZATION_SHARES,
            command=self._on_normalization_change,
        ).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Radiobutton(
            values_frame,
            text="Normalize over latest share price",
            variable=self.normalization_var,
            value=FinanceDataset.NORMALIZATION_SHARE_PRICE,
            command=self._on_normalization_change,
        ).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Radiobutton(
            values_frame,
            text="As reported",
            variable=self.normalization_var,
            value=FinanceDataset.NORMALIZATION_REPORTED,
            command=self._on_normalization_change,
        ).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Radiobutton(
            values_frame,
            text="Number of shares",
            variable=self.normalization_var,
            value=FinancePlotFrame.VALUE_MODE_SHARE_COUNT,
            command=self._on_normalization_change,
        ).pack(side=tk.LEFT, padx=(6, 0))

        self.periods_frame = ttk.Frame(outer)
        self.periods_frame.pack(fill=tk.X, pady=(0, 12))
        ttk.Label(self.periods_frame, text="Dates:").pack(side=tk.LEFT)
        self.period_checks_frame = ttk.Frame(self.periods_frame)
        self.period_checks_frame.pack(side=tk.LEFT, padx=(6, 0))

        self.plot_frame = FinancePlotFrame(outer)
        self.plot_frame.pack(fill=tk.BOTH, expand=True)

    def _on_mode_change(self) -> None:
        mode = self.mode_var.get()
        self.plot_frame.set_display_mode(mode)
        success = self.plot_frame.set_normalization_mode(self.normalization_var.get())
        if not success:
            self.normalization_var.set(self.plot_frame.normalization_mode)

    def _on_normalization_change(self) -> None:
        mode = self.normalization_var.get()
        if not self.plot_frame.set_normalization_mode(mode):
            self.normalization_var.set(self.plot_frame.normalization_mode)

    def _on_period_toggle(self) -> None:
        self._apply_period_filters()

    def _apply_period_filters(self) -> None:
        active_periods = [
            label for label, var in self.period_vars.items() if var.get()
        ]
        if not active_periods and self.plot_frame.all_periods():
            self.plot_frame.set_visible_periods([])
            return
        self.plot_frame.set_visible_periods(active_periods)

    def _refresh_period_controls(self) -> None:
        periods = self.plot_frame.all_periods()
        existing_state = {label: var.get() for label, var in self.period_vars.items()}

        for child in self.period_checks_frame.winfo_children():
            child.destroy()
        self.period_vars.clear()

        for label in periods:
            var = tk.BooleanVar(value=existing_state.get(label, True))
            check = ttk.Checkbutton(
                self.period_checks_frame,
                text=label,
                variable=var,
                command=self._on_period_toggle,
            )
            check.pack(side=tk.LEFT, padx=(0, 6))
            self.period_vars[label] = var

        self._apply_period_filters()


    def _maximize_window(self) -> None:
        try:
            self.root.state("zoomed")
        except tk.TclError:
            try:
                self.root.attributes("-zoomed", True)
            except tk.TclError:
                self.root.attributes("-fullscreen", True)
                self.root.after(100, lambda: self.root.attributes("-fullscreen", False))

    def _refresh_company_list(self) -> None:
        companies = self._discover_companies()
        self.company_combo.configure(values=companies)
        if companies:
            self.company_combo.current(0)
            self.open_button.state(["!disabled"])
        else:
            self.company_var.set("")
            self.open_button.state(["disabled"])

    def _discover_companies(self) -> Sequence[str]:
        if not self.companies_dir.exists():
            return []
        candidates: List[str] = []
        for path in sorted(self.companies_dir.iterdir()):
            if not path.is_dir():
                continue
            combined_path = path / "combined.csv"
            if combined_path.exists():
                candidates.append(path.name)
        return candidates

    def _open_selected_company(self) -> None:
        company = self.company_var.get().strip()
        if not company:
            messagebox.showinfo("Select Company", "Choose a company to analyse.")
            return
        combined_path = self.companies_dir / company / "combined.csv"
        try:
            dataset = FinanceDataset(combined_path)
        except (FileNotFoundError, ValueError) as exc:
            messagebox.showerror("Load Company", str(exc))
            return
        if not dataset.has_data():
            messagebox.showinfo(
                "Load Company",
                "The selected combined.csv does not contain Finance or Income data.",
            )
            return
        try:
            added, normalization_warning = self.plot_frame.add_company(company, dataset)
        except ValueError as exc:
            messagebox.showerror("Load Company", str(exc))
            return
        if not added:
            messagebox.showinfo("Load Company", f"{company} data refreshed on the chart.")
        if normalization_warning:
            messagebox.showwarning("Normalize Values", normalization_warning)
            self.normalization_var.set(self.plot_frame.normalization_mode)
        success = self.plot_frame.set_normalization_mode(self.normalization_var.get())
        if not success:
            self.normalization_var.set(self.plot_frame.normalization_mode)
        self._refresh_period_controls()


def main() -> None:
    root = tk.Tk()
    app = FinanceAnalystApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
