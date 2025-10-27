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

import requests

import matplotlib

import fitz  # PyMuPDF
from PIL import Image, ImageTk

# Ensure TkAgg is used when embedding plots inside Tkinter widgets.
matplotlib.use("TkAgg")

from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from matplotlib import colors, colormaps
from matplotlib.patches import Patch, Rectangle

import tkinter as tk
from tkinter import messagebox, ttk


BASE_FIELDS = {"Type", "Category", "Item", "Note"}
NOTE_ASIS = "asis"
NOTE_NEGATED = "negated"
NOTE_EXCLUDED = "excluded"
NOTE_SHARE_COUNT = "share_count"
VALID_NOTES = {NOTE_ASIS, NOTE_NEGATED, NOTE_EXCLUDED, NOTE_SHARE_COUNT}


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
            for segment in self.finance_segments + self.income_segments:
                segment.values = [segment.values[idx] for idx in sort_indices]
                segment.periods = [self.periods[idx] for idx in sort_indices]

            for segment in self.finance_segments + self.income_segments:
                segment.values = [
                    value / share_counts[idx]
                    for idx, value in enumerate(segment.values)
                ]
                segment.periods = list(self.periods)

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
        self._company_colors: Dict[str, Tuple[str, str]] = {}
        self._palette = list(colormaps["tab10"].colors)
        self._context_menu: Optional[tk.Menu] = None
        self._context_metadata: Optional[Dict[str, Any]] = None
        self._normalize_with_price = False
        self._price_map: Dict[str, float] = {}
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

    def set_price_normalization(
        self, enabled: bool, price_map: Optional[Dict[str, float]] = None
    ) -> None:
        self._normalize_with_price = enabled
        self._price_map = dict(price_map or {})
        if self.datasets:
            self._render()
        else:
            self._render_empty()

    def _price_factor(self, company: str) -> float:
        if not self._normalize_with_price:
            return 1.0
        price = self._price_map.get(company)
        if price is None:
            return 1.0
        if price == 0:
            return 1.0
        return 1.0 / price

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
        self._price_map = {}
        self._render_empty()

    def add_company(self, company: str, dataset: FinanceDataset) -> bool:
        self._register_periods(dataset)
        self.periods = self._rebuild_period_sequence()

        is_new_company = company not in self.datasets
        if is_new_company and company not in self._company_colors:
            self._company_colors[company] = self._next_color_pair()

        self.datasets[company] = dataset
        # Preserve insertion order by moving refreshed companies to the end.
        self.datasets.move_to_end(company)
        self._render()
        return is_new_company

    def _render(self) -> None:
        if not self.datasets:
            self._render_empty()
            return

        periods = self.periods or []
        period_indices = list(range(len(periods)))
        num_companies = len(self.datasets)
        if num_companies == 0:
            self._render_empty()
            return

        self.axis.clear()
        hover_helper = self.hover_helper
        mode = self.display_mode
        context_callback = self._show_context_menu if mode == self.MODE_STACKED else None
        hover_helper.begin_update(context_callback)

        legend_handles: List[Any] = []

        if mode == self.MODE_STACKED:
            group_width = self._PLOT_WIDTH
            bar_width = group_width / max(2 * num_companies, 1)
            start_offset = -group_width / 2

            for company_index, (company, dataset) in enumerate(self.datasets.items()):
                finance_color, income_color = self._company_colors.setdefault(
                    company, self._next_color_pair()
                )
                company_color = finance_color
                normalization_factor = self._price_factor(company)
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
                    if normalization_factor != 1.0:
                        segment_values = [value * normalization_factor for value in segment_values]
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
                            linewidth=0.8,
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
                    if normalization_factor != 1.0:
                        segment_values = [value * normalization_factor for value in segment_values]
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
                            linewidth=0.8,
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
            x_positions = period_indices
            for company, dataset in self.datasets.items():
                finance_color, income_color = self._company_colors.setdefault(
                    company, self._next_color_pair()
                )
                company_color = finance_color

                finance_totals, finance_presence = dataset.aggregate_totals(
                    periods, series=FinanceDataset.FINANCE_LABEL
                )
                normalization_factor = self._price_factor(company)
                if normalization_factor != 1.0:
                    finance_totals = [value * normalization_factor for value in finance_totals]
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
                if normalization_factor != 1.0:
                    income_totals = [value * normalization_factor for value in income_totals]
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
        if self._normalize_with_price:
            self.axis.set_ylabel("Value per Share (Price-Normalized)")
        else:
            self.axis.set_ylabel("Value per Share")
        if self.datasets:
            companies_list = ", ".join(self.datasets.keys())
            self.axis.set_title(f"Finance vs Income Statement — {companies_list}")
        else:
            self.axis.set_title("Finance vs Income Statement")

        if legend_handles:
            legend_kwargs = {"loc": "upper left"}
            if mode == self.MODE_STACKED:
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
        self._connection_id: Optional[int] = None
        self._button_connection_id: Optional[int] = None
        self._active_rectangle: Optional[Rectangle] = None
        self._context_callback: Optional[Callable[[Dict[str, Any], Any], None]] = None

    def attach(self, canvas: FigureCanvasTkAgg) -> None:
        if self._connection_id is not None and self._canvas is not None:
            self._canvas.mpl_disconnect(self._connection_id)
        if self._button_connection_id is not None and self._canvas is not None:
            self._canvas.mpl_disconnect(self._button_connection_id)
        self._canvas = canvas
        self.reset()
        self._connection_id = canvas.mpl_connect("motion_notify_event", self._on_motion)
        self._button_connection_id = None
        self._context_callback = None

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
        value = metadata.get("value", 0.0)
        formatted_value = f"{value:,.2f}"
        type_value = metadata.get("type_value") or metadata.get("type_label") or "—"
        category = metadata.get("category") or "—"
        item = metadata.get("item") or "—"
        key = metadata.get("key") or "—"
        period = metadata.get("period") or "Period"
        company = metadata.get("company") or "—"
        return (
            f"Company: {company}\n"
            f"Key: {key}\n"
            f"Type: {type_value}\n"
            f"Category: {category}\n"
            f"Item: {item}\n"
            f"{period}: {formatted_value}"
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

    def _on_motion(self, event) -> None:
        if self._canvas is None:
            return
        if event.inaxes != self.axis:
            if self._annotation.get_visible():
                self._annotation.set_visible(False)
                self._canvas.draw_idle()
            self._active_rectangle = None
            return

        for rect, metadata in self._rectangles:
            contains, _ = rect.contains(event)
            if not contains:
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

        if self._annotation.get_visible():
            self._annotation.set_visible(False)
            self._canvas.draw_idle()
        self._active_rectangle = None

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
        self.normalize_var = tk.BooleanVar(value=False)
        self._price_cache: Dict[str, float] = {}
        self._active_prices: Dict[str, float] = {}

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
        ttk.Checkbutton(
            mode_frame,
            text="Normalize against stock price",
            variable=self.normalize_var,
            command=self._on_normalize_toggle,
        ).pack(side=tk.LEFT, padx=(18, 0))

        self.plot_frame = FinancePlotFrame(outer)
        self.plot_frame.pack(fill=tk.BOTH, expand=True)

        self.price_frame = ttk.Frame(outer)
        self.price_label = ttk.Label(self.price_frame, text="", anchor="w")
        self.price_label.pack(fill=tk.X)

    def _on_mode_change(self) -> None:
        mode = self.mode_var.get()
        self.plot_frame.set_display_mode(mode)
        if self.normalize_var.get():
            # Changing modes may redraw the plot; ensure price state is applied.
            self._update_prices()

    def _on_normalize_toggle(self) -> None:
        if self.normalize_var.get():
            self._update_prices()
        else:
            self._active_prices = {}
            self.plot_frame.set_price_normalization(False, {})
            self._update_price_display()

    def _ensure_price_frame_visible(self) -> None:
        if not self.price_frame.winfo_ismapped():
            self.price_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=(8, 0))

    def _hide_price_frame(self) -> None:
        if self.price_frame.winfo_ismapped():
            self.price_frame.pack_forget()

    def _update_price_display(self) -> None:
        if self.normalize_var.get() and self._active_prices:
            self._ensure_price_frame_visible()
            parts = [
                f"{company}: ${price:,.2f}" for company, price in self._active_prices.items()
            ]
            display_text = "Stock prices: " + " | ".join(parts)
            self.price_label.configure(text=display_text)
        elif self.normalize_var.get():
            self._ensure_price_frame_visible()
            self.price_label.configure(text="Stock prices: awaiting data...")
        else:
            self.price_label.configure(text="")
            self._hide_price_frame()

    def _fetch_stock_price(self, symbol: str) -> float:
        if symbol in self._price_cache:
            return self._price_cache[symbol]

        url = f"https://query1.finance.yahoo.com/v8/finance/chart/{symbol}"
        params = {"range": "1d", "interval": "1d"}
        try:
            response = requests.get(url, params=params, timeout=10)
            response.raise_for_status()
            payload = response.json()
        except requests.RequestException as exc:
            raise ValueError(f"Could not download price for {symbol}: {exc}") from exc
        except ValueError as exc:
            raise ValueError(
                f"Received an invalid response while downloading price for {symbol}"
            ) from exc

        chart_data = payload.get("chart")
        if not isinstance(chart_data, dict):
            raise ValueError(f"No quote data returned for {symbol}")
        if chart_data.get("error"):
            raise ValueError(f"Quote service returned an error for {symbol}")
        results = chart_data.get("result")
        if not isinstance(results, list) or not results:
            raise ValueError(f"No quote data returned for {symbol}")
        first_entry = results[0]
        if not isinstance(first_entry, dict):
            raise ValueError(f"Quote for {symbol} was malformed")
        meta = first_entry.get("meta")
        if not isinstance(meta, dict):
            raise ValueError(f"Quote for {symbol} was missing metadata")
        price_value = meta.get("regularMarketPrice")
        if price_value in (None, 0):
            price_value = meta.get("previousClose")
        try:
            numeric_price = float(price_value)
        except (TypeError, ValueError) as exc:
            raise ValueError(f"Quote for {symbol} returned an invalid price value") from exc
        if numeric_price <= 0:
            raise ValueError(f"Quote for {symbol} returned a non-positive price")

        self._price_cache[symbol] = numeric_price
        return numeric_price

    def _update_prices(self) -> bool:
        if not self.normalize_var.get():
            self._active_prices = {}
            self.plot_frame.set_price_normalization(False, {})
            self._update_price_display()
            return True

        companies = list(self.plot_frame.datasets.keys())
        if not companies:
            self._active_prices = {}
            self.plot_frame.set_price_normalization(False, {})
            self._update_price_display()
            return True

        price_map: Dict[str, float] = {}
        for company in companies:
            try:
                price = self._price_cache[company]
            except KeyError:
                try:
                    price = self._fetch_stock_price(company)
                except ValueError as exc:
                    messagebox.showerror("Stock Price", str(exc))
                    self.normalize_var.set(False)
                    self._active_prices = {}
                    self.plot_frame.set_price_normalization(False, {})
                    self._update_price_display()
                    return False
            price_map[company] = price

        self._active_prices = price_map
        self.plot_frame.set_price_normalization(True, price_map)
        self._update_price_display()
        return True

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
            added = self.plot_frame.add_company(company, dataset)
        except ValueError as exc:
            messagebox.showerror("Load Company", str(exc))
            return
        if not added:
            messagebox.showinfo("Load Company", f"{company} data refreshed on the chart.")
        if self.normalize_var.get():
            self._update_prices()


def main() -> None:
    root = tk.Tk()
    app = FinanceAnalystApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
