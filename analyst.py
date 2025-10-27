"""Standalone analyst application for combined CSV summaries.

This lightweight Tkinter application lets reviewers choose companies that have
``combined.csv`` files in ``companies/<company>/``. When a company is opened, it
is added to a shared set of side-by-side Finance and Income Statement charts so
reviewers can compare multiple companies without switching tabs. Each bar is
stacked by the underlying rows from the ``combined.csv`` file so reviewers can
see how individual entries contribute to each reporting period.
"""

from __future__ import annotations

import calendar
import csv
import math
import re
from bisect import bisect_left
from collections import OrderedDict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Sequence, Set, Tuple

import json
import os
import sys
import webbrowser
from datetime import date, datetime, timedelta, timezone

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

import urllib.error
import urllib.request
from urllib.parse import quote


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


_PERIOD_DATE_FORMATS = (
    "%d.%m.%Y",
    "%d/%m/%Y",
    "%d-%m-%Y",
    "%Y-%m-%d",
    "%Y.%m.%d",
    "%Y/%m/%d",
    "%d %b %Y",
    "%d %B %Y",
    "%b %Y",
    "%B %Y",
    "%Y",
)


def _normalize_two_digit_year(year_text: str) -> Optional[int]:
    try:
        value = int(year_text)
    except ValueError:
        return None
    if 0 <= value <= 99:
        return 2000 + value if value < 50 else 1900 + value
    return value if value >= 1900 else None


def _parse_period_label(label: str) -> Optional[datetime]:
    """Best-effort parsing for period labels into comparable datetimes."""

    text = label.strip()
    if not text:
        return None

    normalized = (
        text.replace("\u2013", "-")
        .replace("\u2014", "-")
        .replace("\u2212", "-")
    )
    candidates = {normalized, normalized.replace("  ", " ")}
    simplified = normalized.replace(".", "/")
    candidates.update({simplified, simplified.replace("/", "-"), normalized.replace("/", "-")})

    upper_text = normalized.upper()
    if upper_text.startswith("FY"):
        candidates.add(normalized[2:].strip())
    candidates.add(re.sub(r"\bFY\b", "", normalized, flags=re.IGNORECASE).strip())

    for candidate in list(candidates):
        collapsed = re.sub(r"\s+", " ", candidate).strip()
        if collapsed:
            candidates.add(collapsed)

    for candidate in candidates:
        for date_format in _PERIOD_DATE_FORMATS:
            try:
                parsed = datetime.strptime(candidate, date_format)
            except ValueError:
                continue
            else:
                return parsed

    year_match = re.search(r"(?<!\d)(\d{4})(?!\d)", normalized)
    year_value: Optional[int] = None
    if year_match:
        year_value = int(year_match.group(1))
    else:
        short_match = re.search(r"(?<!\d)(\d{2})(?!\d)", normalized)
        if short_match:
            year_value = _normalize_two_digit_year(short_match.group(1))

    if year_value is None:
        return None

    month = 1
    quarter_match = re.search(r"Q([1-4])", normalized, flags=re.IGNORECASE)
    if quarter_match:
        quarter = int(quarter_match.group(1))
        month = min(quarter * 3, 12)
    else:
        half_match = re.search(r"H([12])", normalized, flags=re.IGNORECASE)
        if half_match:
            half = int(half_match.group(1))
            month = 6 if half == 1 else 12
        elif re.search(r"FY", normalized, flags=re.IGNORECASE):
            month = 12

    day = min(28, calendar.monthrange(year_value, month)[1])
    return datetime(year_value, month, day)


def _period_sort_key(label: str) -> Tuple[int, Any]:
    parsed = _parse_period_label(label)
    if parsed is not None:
        return (0, parsed)
    return (1, label.strip().lower())


def _normalized_period_token(label: str) -> Tuple[str, int, int, str]:
    """Return a hashable token so equivalent period labels can be aligned."""

    normalized = re.sub(r"\s+", " ", label.strip().lower())
    if not normalized:
        return ("", 0, 0, "")

    parsed = _parse_period_label(label)
    if parsed is None:
        return ("text", 0, 0, normalized)

    year = parsed.year
    quarter_match = re.search(r"q([1-4])", normalized)
    if quarter_match:
        return ("quarter", year, int(quarter_match.group(1)), "")
    half_match = re.search(r"h([12])", normalized)
    if half_match:
        return ("half", year, int(half_match.group(1)), "")
    if re.search(r"fy", normalized) or re.fullmatch(r"\d{4}", normalized):
        return ("year", year, 0, "")

    month = parsed.month
    day = parsed.day
    return ("date", year, month * 100 + day, "")


def _lookup_period_value(collection: Dict[str, Any], label: str) -> Any:
    """Attempt to find a period entry by direct label or normalized alias."""

    if label in collection:
        return collection[label]
    token = _normalized_period_token(label)
    for candidate_label, value in collection.items():
        if _normalized_period_token(candidate_label) == token:
            return value
    raise KeyError(label)


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


class _PdfMonitorPlacementMixin:
    """Utilities for positioning PDF viewers on the preferred monitor."""

    _placement: str = "full"

    def _preferred_monitor_info(self) -> Dict[str, int]:
        monitor_info: Optional[Dict[str, int]] = None
        monitor_count = 1
        try:
            monitor_count = int(self.tk.call("winfo", "monitorcount", self))  # type: ignore[attr-defined]
        except tk.TclError:
            monitor_count = 1

        if monitor_count > 1:
            monitors: List[Dict[str, Any]] = []
            try:
                monitors_raw = self.tk.splitlist(self.tk.call("winfo", "monitors", self))  # type: ignore[attr-defined]
            except tk.TclError:
                monitors_raw = []

            for raw_monitor in monitors_raw:
                parts = list(self.tk.splitlist(raw_monitor))  # type: ignore[attr-defined]
                if len(parts) < 4:
                    continue
                try:
                    x = int(parts[0])
                    y = int(parts[1])
                    width = int(parts[2])
                    height = int(parts[3])
                except (TypeError, ValueError):
                    continue

                primary = False
                for extra in parts[4:]:
                    text = str(extra).strip().lower()
                    if text in {"1", "0", "true", "false"}:
                        primary = text in {"1", "true"}
                        break

                monitors.append(
                    {
                        "x": x,
                        "y": y,
                        "width": width,
                        "height": height,
                        "primary": primary,
                    }
                )

            non_primary = [monitor for monitor in monitors if not monitor.get("primary")]
            if non_primary:
                monitor_info = {
                    "x": int(non_primary[0]["x"]),
                    "y": int(non_primary[0]["y"]),
                    "width": int(non_primary[0]["width"]),
                    "height": int(non_primary[0]["height"]),
                }
            elif monitors:
                monitor_info = {
                    "x": int(monitors[0]["x"]),
                    "y": int(monitors[0]["y"]),
                    "width": int(monitors[0]["width"]),
                    "height": int(monitors[0]["height"]),
                }

        if monitor_info is None:
            monitor_info = {
                "x": 0,
                "y": 0,
                "width": int(self.winfo_screenwidth()),
                "height": int(self.winfo_screenheight()),
            }

        return monitor_info

    def _apply_monitor_geometry(self, placement: str = "full") -> None:
        monitor_info = self._preferred_monitor_info()
        width = int(monitor_info["width"])
        height = int(monitor_info["height"])
        x = int(monitor_info["x"])
        y = int(monitor_info["y"])

        normalized = placement.lower()
        if normalized not in {"full", "left", "right"}:
            normalized = "full"
        self._placement = normalized

        if normalized in {"left", "right"} and width > 1:
            left_width = max(1, width // 2)
            right_width = max(1, width - left_width)
            if normalized == "left":
                width = left_width
            else:
                width = right_width
                x += left_width

        geometry = f"{width}x{height}+{x}+{y}"
        self.geometry(geometry)  # type: ignore[misc]
        self.update_idletasks()

        if normalized == "full":
            try:
                self.attributes("-fullscreen", True)  # type: ignore[attr-defined]
                return
            except tk.TclError:
                try:
                    self.state("zoomed")  # type: ignore[attr-defined]
                except tk.TclError:
                    pass
        else:
            try:
                self.attributes("-fullscreen", False)  # type: ignore[attr-defined]
            except tk.TclError:
                pass
            try:
                self.state("normal")  # type: ignore[attr-defined]
            except tk.TclError:
                pass


def _render_pdf_page(pdf_path: Path, page_number: int) -> ImageTk.PhotoImage:
    if fitz is None:
        raise RuntimeError("PyMuPDF (fitz) is not installed")
    if page_number <= 0:
        raise ValueError("page_number must be 1 or greater")
    with fitz.open(pdf_path) as doc:  # type: ignore[union-attr]
        page = doc.load_page(page_number - 1)
        zoom_matrix = fitz.Matrix(1.5, 1.5)
        pix = page.get_pixmap(matrix=zoom_matrix)
    mode = "RGBA" if pix.alpha else "RGB"
    image = Image.frombytes(mode, [pix.width, pix.height], pix.samples)
    return ImageTk.PhotoImage(image)


class PdfPageViewer(_PdfMonitorPlacementMixin, tk.Toplevel):
    """Display a single rendered PDF page inside a scrollable window."""

    def __init__(
        self,
        master: tk.Widget,
        pdf_path: Path,
        page_number: int,
        *,
        placement: str = "full",
    ) -> None:
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
        self._apply_monitor_geometry(placement)
        self.bind("<Escape>", self._close_from_escape)

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
            self._photo = _render_pdf_page(self._pdf_path, self._page_number)
        except Exception as exc:
            messagebox.showwarning(
                "Open PDF",
                f"Could not render page {self._page_number} from {self._pdf_path.name}: {exc}",
            )
            self.after(0, self.destroy)
            return

        if self._photo is None:
            return

        width = self._photo.width()
        height = self._photo.height()
        self._canvas.delete("all")
        self._canvas.create_image(0, 0, anchor="nw", image=self._photo)
        self._canvas.configure(scrollregion=(0, 0, width, height))

    def _close_from_escape(self, _event: Any) -> None:
        try:
            self.attributes("-fullscreen", False)
        except tk.TclError:
            pass
        self.destroy()


class PdfSplitViewer(_PdfMonitorPlacementMixin, tk.Toplevel):
    """Display two PDF pages side-by-side on the preferred monitor."""

    def __init__(
        self,
        master: tk.Widget,
        left: Tuple[str, SegmentSource],
        right: Tuple[str, SegmentSource],
        title: str,
    ) -> None:
        super().__init__(master)
        if fitz is None:
            messagebox.showerror(
                "PyMuPDF Required",
                (
                    "Viewing individual PDF pages requires the PyMuPDF package (import name 'fitz').\n"
                    "Install PyMuPDF to enable in-app previews."
                ),
            )
            self.after(0, self.destroy)
            raise RuntimeError("PyMuPDF (fitz) is not installed")

        self.title(title)
        self._left_label, self._left_source = left
        self._right_label, self._right_source = right
        self._left_photo: Optional[ImageTk.PhotoImage] = None
        self._right_photo: Optional[ImageTk.PhotoImage] = None
        self._build_widgets()
        self._render_pages()
        self._apply_monitor_geometry("full")
        self.bind("<Escape>", self._close_from_escape)

    def _build_widgets(self) -> None:
        self.configure(bg="#1f1f1f")
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        paned = ttk.Panedwindow(self, orient=tk.HORIZONTAL)
        paned.grid(row=0, column=0, sticky="nsew")

        self._left_frame = ttk.Frame(paned)
        self._left_frame.grid_rowconfigure(1, weight=1)
        self._left_frame.grid_columnconfigure(0, weight=1)
        self._right_frame = ttk.Frame(paned)
        self._right_frame.grid_rowconfigure(1, weight=1)
        self._right_frame.grid_columnconfigure(0, weight=1)
        paned.add(self._left_frame, weight=1)
        paned.add(self._right_frame, weight=1)

        ttk.Label(
            self._left_frame,
            text=self._left_label,
            anchor="center",
            padding=(0, 6),
        ).grid(row=0, column=0, columnspan=2, sticky="ew")

        self._left_canvas = tk.Canvas(
            self._left_frame, bg="#2b2b2b", highlightthickness=0
        )
        self._left_canvas.grid(row=1, column=0, sticky="nsew")
        self._left_vbar = tk.Scrollbar(
            self._left_frame, orient=tk.VERTICAL, command=self._left_canvas.yview
        )
        self._left_vbar.grid(row=1, column=1, sticky="ns")
        self._left_canvas.configure(yscrollcommand=self._left_vbar.set)
        self._left_hbar = tk.Scrollbar(
            self._left_frame, orient=tk.HORIZONTAL, command=self._left_canvas.xview
        )
        self._left_hbar.grid(row=2, column=0, columnspan=2, sticky="ew")
        self._left_canvas.configure(xscrollcommand=self._left_hbar.set)

        ttk.Label(
            self._right_frame,
            text=self._right_label,
            anchor="center",
            padding=(0, 6),
        ).grid(row=0, column=0, columnspan=2, sticky="ew")

        self._right_canvas = tk.Canvas(
            self._right_frame, bg="#2b2b2b", highlightthickness=0
        )
        self._right_canvas.grid(row=1, column=0, sticky="nsew")
        self._right_vbar = tk.Scrollbar(
            self._right_frame, orient=tk.VERTICAL, command=self._right_canvas.yview
        )
        self._right_vbar.grid(row=1, column=1, sticky="ns")
        self._right_canvas.configure(yscrollcommand=self._right_vbar.set)
        self._right_hbar = tk.Scrollbar(
            self._right_frame, orient=tk.HORIZONTAL, command=self._right_canvas.xview
        )
        self._right_hbar.grid(row=2, column=0, columnspan=2, sticky="ew")
        self._right_canvas.configure(xscrollcommand=self._right_hbar.set)

    def _render_pages(self) -> None:
        if fitz is None:
            raise RuntimeError("PyMuPDF (fitz) is not installed")

        try:
            left_page = int(self._left_source.page)
            self._left_photo = _render_pdf_page(self._left_source.pdf_path, left_page)
        except Exception as exc:
            messagebox.showwarning(
                "Open PDF",
                f"Could not render {self._left_label} page {self._left_source.page}: {exc}",
            )
            self.after(0, self.destroy)
            return

        try:
            right_page = int(self._right_source.page)
            self._right_photo = _render_pdf_page(self._right_source.pdf_path, right_page)
        except Exception as exc:
            messagebox.showwarning(
                "Open PDF",
                f"Could not render {self._right_label} page {self._right_source.page}: {exc}",
            )
            self.after(0, self.destroy)
            return

        if self._left_photo is not None:
            width = self._left_photo.width()
            height = self._left_photo.height()
            self._left_canvas.delete("all")
            self._left_canvas.create_image(0, 0, anchor="nw", image=self._left_photo)
            self._left_canvas.configure(scrollregion=(0, 0, width, height))

        if self._right_photo is not None:
            width = self._right_photo.width()
            height = self._right_photo.height()
            self._right_canvas.delete("all")
            self._right_canvas.create_image(0, 0, anchor="nw", image=self._right_photo)
            self._right_canvas.configure(scrollregion=(0, 0, width, height))

    def _close_from_escape(self, _event: Any) -> None:
        try:
            self.attributes("-fullscreen", False)
        except tk.TclError:
            pass
        self.destroy()


class FinanceDataset:
    """Load and aggregate Finance and Income Statement totals from a CSV."""

    FINANCE_LABEL = "Finance"
    INCOME_LABEL = "Income Statement"

    NORMALIZATION_SHARES = "shares"
    NORMALIZATION_REPORTED = "reported"
    NORMALIZATION_SHARE_PRICE = "share_price"
    NORMALIZATION_SHARE_PRICE_PERIOD = "share_price_period"

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
        self.share_count_sources: Dict[str, SegmentSource] = {}
        self.share_price_symbol: Optional[str] = None
        self._period_share_price_map: Optional[Dict[str, float]] = None
        self._share_price_error: Optional[str] = None
        self._period_index: Dict[str, int] = {}
        self._period_token_index: Dict[Tuple[str, int, int, str], int] = {}
        self._normalization_mode = self.NORMALIZATION_REPORTED
        self._load()
        self._refresh_period_indices()
        self._load_metadata()
        self._prefetch_share_prices()

    def _prefetch_share_prices(self) -> None:
        if self._period_share_price_map is not None or self._share_price_error is not None:
            return
        try:
            self._period_share_price_map = self._fetch_period_share_prices()
        except ValueError as exc:
            self._share_price_error = str(exc)

    def _refresh_period_indices(self) -> None:
        self._period_index = {label: idx for idx, label in enumerate(self.periods)}
        token_map: Dict[Tuple[str, int, int, str], int] = {}
        for idx, label in enumerate(self.periods):
            token = _normalized_period_token(label)
            if token not in token_map:
                token_map[token] = idx
        self._period_token_index = token_map

    def _resolve_period_index(self, label: str) -> Optional[int]:
        lookup_label = str(label)
        if lookup_label in self._period_index:
            return self._period_index[lookup_label]
        token = _normalized_period_token(lookup_label)
        return self._period_token_index.get(token)

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

            sort_indices = sorted(
                range(len(self.periods)), key=lambda idx: _period_sort_key(self.periods[idx])
            )
            if sort_indices != list(range(len(self.periods))):
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
            self.NORMALIZATION_SHARE_PRICE_PERIOD,
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
            price_map = self._get_period_share_price_map()
            latest_price: Optional[float] = None
            for label in reversed(self.periods):
                price = price_map.get(str(label))
                if price is not None and math.isfinite(price) and price > 0:
                    latest_price = float(price)
                    break
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
        elif mode == self.NORMALIZATION_SHARE_PRICE_PERIOD:
            if not self.share_counts:
                raise ValueError("combined.csv is missing a share_count row")
            price_map = self._get_period_share_price_map()
            for segment in self.finance_segments + self.income_segments:
                multiples: List[float] = []
                for idx, value in enumerate(segment.raw_values):
                    count = self.share_counts[idx] if idx < len(self.share_counts) else 0.0
                    if idx < len(segment.periods):
                        period_label = str(segment.periods[idx])
                    elif idx < len(self.periods):
                        period_label = str(self.periods[idx])
                    else:
                        period_label = str(idx)
                    price = price_map.get(period_label)
                    if count == 0 or price is None or price <= 0:
                        multiples.append(0.0)
                        continue
                    per_share = value / count
                    if per_share == 0:
                        multiples.append(0.0)
                    else:
                        multiples.append(per_share / price)
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

        ticker_value = (
            payload.get("share_price_symbol")
            or payload.get("share_price_ticker")
            or payload.get("ticker")
            or payload.get("symbol")
        )
        if isinstance(ticker_value, str) and ticker_value.strip():
            self.share_price_symbol = ticker_value.strip()

        rows = payload.get("rows")
        if not isinstance(rows, list):
            return

        metadata_map: Dict[Tuple[str, str, str], Dict[str, SegmentSource]] = {}
        share_count_sources: Dict[str, SegmentSource] = {}

        for entry in rows:
            if not isinstance(entry, dict):
                continue
            type_value = str(entry.get("type", ""))
            category_value = str(entry.get("category", ""))
            item_value = str(entry.get("item", ""))
            key = (type_value, category_value, item_value)
            note_value = str(entry.get("note", "")).strip().lower()
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
                if note_value == NOTE_SHARE_COUNT:
                    share_count_sources = {
                        str(label): source for label, source in period_sources.items()
                    }
                    continue
                metadata_map[key] = period_sources

        if not metadata_map:
            if share_count_sources:
                self.share_count_sources = share_count_sources
            return

        for segment in self.finance_segments + self.income_segments:
            key = (segment.type_value, segment.category, segment.item)
            sources = metadata_map.get(key)
            if sources:
                segment.sources = sources
        if share_count_sources:
            self.share_count_sources = share_count_sources

    def aggregate_totals(
        self, periods: Sequence[str], *, series: str
    ) -> Tuple[List[float], List[bool]]:
        if series == self.FINANCE_LABEL:
            segments = self.finance_segments
        elif series == self.INCOME_LABEL:
            segments = self.income_segments
        else:
            return [0.0] * len(periods), [False] * len(periods)

        if len(self._period_index) != len(self.periods):
            self._refresh_period_indices()

        totals: List[float] = []
        has_data: List[bool] = []

        for label in periods:
            idx = self._resolve_period_index(label)
            if idx is None:
                totals.append(math.nan)
                has_data.append(False)
                continue

            total = 0.0
            present = False
            has_nonzero = False
            for segment in segments:
                if idx >= len(segment.values):
                    continue
                value = segment.values[idx]
                if not math.isfinite(value):
                    continue
                present = True
                total += value
                if not has_nonzero and abs(value) > 1e-9:
                    has_nonzero = True

            if present and has_nonzero:
                totals.append(total)
                has_data.append(True)
            else:
                totals.append(math.nan)
                has_data.append(False)

        return totals, has_data

    def share_counts_for(self, periods: Sequence[str]) -> Tuple[List[float], List[bool]]:
        if not self.share_counts:
            return [math.nan] * len(periods), [False] * len(periods)

        if len(self._period_index) != len(self.periods):
            self._refresh_period_indices()

        values: List[float] = []
        present: List[bool] = []

        for label in periods:
            idx = self._resolve_period_index(label)
            if idx is None or idx >= len(self.share_counts):
                values.append(math.nan)
                present.append(False)
                continue
            value = self.share_counts[idx]
            if not math.isfinite(value) or abs(value) <= 1e-9:
                values.append(math.nan)
                present.append(False)
                continue
            values.append(float(value))
            present.append(True)

        return values, present

    def _try_get_period_share_price_map(self) -> Dict[str, float]:
        if self._period_share_price_map is not None:
            return self._period_share_price_map
        if self._share_price_error is not None:
            return {}
        try:
            self._period_share_price_map = self._fetch_period_share_prices()
        except ValueError as exc:
            self._share_price_error = str(exc)
            return {}
        return self._period_share_price_map or {}

    def _build_share_price_list(
        self, periods: Sequence[str]
    ) -> Tuple[List[float], List[bool]]:
        price_map = self._try_get_period_share_price_map()
        if len(self._period_index) != len(self.periods):
            self._refresh_period_indices()
        values: List[float] = []
        present: List[bool] = []

        for label in periods:
            lookup_label = str(label)
            price = price_map.get(lookup_label)
            if price is None:
                try:
                    price = _lookup_period_value(price_map, lookup_label)
                except KeyError:
                    price = None
            if price is None:
                idx = self._resolve_period_index(lookup_label)
                if idx is not None and idx < len(self.share_prices):
                    fallback_price = self.share_prices[idx]
                    if fallback_price and math.isfinite(fallback_price) and fallback_price > 0:
                        price = float(fallback_price)
            if price is not None and math.isfinite(price) and price > 0:
                values.append(float(price))
                present.append(True)
            else:
                values.append(math.nan)
                present.append(False)

        return values, present

    def share_prices_for(self, periods: Sequence[str]) -> Tuple[List[float], List[bool]]:
        values, present = self._build_share_price_list(periods)
        if not any(present):
            message = self._share_price_error or (
                f"Share price data is not available for {self.company_root.name or 'the selected company'}."
            )
            raise ValueError(message)
        return values, present

    def share_prices_for_optional(
        self, periods: Sequence[str]
    ) -> Tuple[List[float], List[bool]]:
        return self._build_share_price_list(periods)

    def aggregate_raw_totals(
        self, periods: Sequence[str], *, series: str
    ) -> Tuple[List[float], List[bool]]:
        if series == self.FINANCE_LABEL:
            segments = self.finance_segments
        elif series == self.INCOME_LABEL:
            segments = self.income_segments
        else:
            return [math.nan] * len(periods), [False] * len(periods)

        if len(self._period_index) != len(self.periods):
            self._refresh_period_indices()

        totals: List[float] = []
        has_data: List[bool] = []

        for label in periods:
            idx = self._resolve_period_index(label)
            if idx is None:
                totals.append(math.nan)
                has_data.append(False)
                continue

            total = 0.0
            present = False
            has_nonzero = False
            for segment in segments:
                if idx >= len(segment.raw_values):
                    continue
                value = segment.raw_values[idx]
                if not math.isfinite(value):
                    continue
                present = True
                total += value
                if not has_nonzero and abs(value) > 1e-9:
                    has_nonzero = True

            if present and has_nonzero:
                totals.append(total)
                has_data.append(True)
            else:
                totals.append(math.nan)
                has_data.append(False)

        return totals, has_data

    def source_for_series(self, series: str, period: str) -> Optional[SegmentSource]:
        lookup: Sequence[RowSegment]
        if series == self.FINANCE_LABEL:
            lookup = self.finance_segments
        elif series == self.INCOME_LABEL:
            lookup = self.income_segments
        else:
            return None
        for segment in lookup:
            if not segment.sources:
                continue
            try:
                return _lookup_period_value(segment.sources, str(period))
            except KeyError:
                continue
        return None

    def share_count_source_for(self, period: str) -> Optional[SegmentSource]:
        if not self.share_count_sources:
            return None
        try:
            return _lookup_period_value(self.share_count_sources, str(period))
        except KeyError:
            return None

    def latest_share_price(self) -> Optional[float]:
        for value in reversed(self.share_prices):
            if math.isfinite(value) and value > 0:
                return value
        return None

    def _resolve_share_price_symbol(self) -> Optional[str]:
        if self.share_price_symbol and self.share_price_symbol.strip():
            return self.share_price_symbol.strip()
        company_name = self.company_root.name.strip()
        return company_name or None

    def _fetch_period_share_prices(self) -> Dict[str, float]:
        symbol = self._resolve_share_price_symbol()
        if not symbol:
            raise ValueError("A share price symbol could not be determined for this company.")

        period_dates: List[Tuple[str, datetime]] = []
        for label in self.periods:
            parsed = _parse_period_label(label)
            if parsed is None:
                continue
            period_dates.append((label, parsed))

        if not period_dates:
            raise ValueError("The reporting periods could not be translated into calendar dates.")

        period_dates.sort(key=lambda item: item[1])
        earliest = period_dates[0][1]
        latest = period_dates[-1][1]
        start_date = (earliest - timedelta(days=3)).date()
        end_date = (latest + timedelta(days=3)).date()
        start_dt = datetime(start_date.year, start_date.month, start_date.day, tzinfo=timezone.utc)
        end_dt = datetime(end_date.year, end_date.month, end_date.day, tzinfo=timezone.utc) + timedelta(days=1)
        period1 = int(start_dt.timestamp())
        period2 = int(end_dt.timestamp())

        encoded_symbol = quote(symbol)
        url = (
            "https://query1.finance.yahoo.com/v8/finance/chart/"
            f"{encoded_symbol}?interval=1d&period1={period1}&period2={period2}"
        )

        request = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        try:
            with urllib.request.urlopen(request, timeout=15) as response:
                payload = json.loads(response.read().decode("utf-8"))
        except urllib.error.URLError as exc:
            raise ValueError(
                f"Unable to download share price history for {symbol}: {exc.reason if hasattr(exc, 'reason') else exc}"
            ) from exc

        chart = payload.get("chart", {}) if isinstance(payload, dict) else {}
        results = chart.get("result") if isinstance(chart, dict) else None
        if not results:
            error_info = chart.get("error") if isinstance(chart, dict) else None
            description = ""
            if isinstance(error_info, dict):
                description = error_info.get("description") or error_info.get("message") or ""
            raise ValueError(
                "Share price data was not returned by Yahoo Finance"
                + (f": {description}" if description else ".")
            )

        result = results[0]
        timestamps = result.get("timestamp") if isinstance(result, dict) else None
        indicators = result.get("indicators") if isinstance(result, dict) else None
        quotes = indicators.get("quote") if isinstance(indicators, dict) else None
        quote0 = quotes[0] if isinstance(quotes, list) and quotes else None
        closes = quote0.get("close") if isinstance(quote0, dict) else None

        if not timestamps or not closes:
            raise ValueError("Share price history was returned without timestamp information.")

        date_prices: Dict[date, float] = {}
        for index, timestamp_value in enumerate(timestamps):
            if index >= len(closes):
                break
            close_value = closes[index]
            if close_value is None or not math.isfinite(close_value):
                continue
            try:
                ts = float(timestamp_value)
            except (TypeError, ValueError):
                continue
            closing_time = datetime.fromtimestamp(ts, tz=timezone.utc)
            date_prices[closing_time.date()] = float(close_value)

        if not date_prices:
            raise ValueError("No valid share price entries were found in the Yahoo Finance response.")

        available_dates = sorted(date_prices.keys())
        prices: Dict[str, float] = {}
        for label, target_datetime in period_dates:
            target_date = target_datetime.date()
            price = date_prices.get(target_date)
            if price is None:
                index = bisect_left(available_dates, target_date)
                while index < len(available_dates):
                    candidate = available_dates[index]
                    if candidate >= target_date:
                        price = date_prices[candidate]
                        break
                    index += 1
            if price is not None and price > 0 and math.isfinite(price):
                prices[str(label)] = float(price)

        if len(prices) < len(period_dates) and self.share_prices:
            for idx, label in enumerate(self.periods):
                if label in prices:
                    continue
                if idx >= len(self.share_prices):
                    continue
                fallback_price = self.share_prices[idx]
                if fallback_price and math.isfinite(fallback_price) and fallback_price > 0:
                    prices[str(label)] = float(fallback_price)

        if len(prices) < len(period_dates):
            missing_labels = [label for label, _ in period_dates if str(label) not in prices]
            missing_preview = ", ".join(missing_labels[:3])
            if len(missing_labels) > 3:
                missing_preview += ", …"
            raise ValueError(
                "Share price history is unavailable for the following periods: "
                f"{missing_preview}"
            )

        return prices

    def _get_period_share_price_map(self) -> Dict[str, float]:
        if self._period_share_price_map is not None:
            return self._period_share_price_map
        if self._share_price_error is not None:
            raise ValueError(self._share_price_error)
        try:
            price_map = self._fetch_period_share_prices()
        except ValueError as exc:
            self._share_price_error = str(exc)
            raise
        self._period_share_price_map = price_map
        return price_map

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
    VALUE_MODE_SHARE_PRICE = "share_price_values"

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
        self._share_price_ratio_formatter = ScalarFormatter(useOffset=False)
        self._share_price_ratio_formatter.set_scientific(True)
        self._share_price_ratio_formatter.set_powerlimits((0, 0))
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
        self.canvas.draw()

    def _compute_period_metrics(
        self,
        dataset: FinanceDataset,
        periods: Sequence[str],
        *,
        series: str,
    ) -> Dict[str, Dict[str, Any]]:
        if not periods:
            return {}

        share_counts, share_count_presence = dataset.share_counts_for(periods)
        share_prices, share_price_presence = dataset.share_prices_for_optional(periods)
        reported_totals, reported_presence = dataset.aggregate_raw_totals(periods, series=series)

        metrics: Dict[str, Dict[str, Any]] = {}
        for idx, label in enumerate(periods):
            label_key = str(label)

            share_count_value = (
                float(share_counts[idx])
                if idx < len(share_counts)
                and share_count_presence[idx]
                and math.isfinite(share_counts[idx])
                and share_counts[idx] > 0
                else math.nan
            )

            reported_value = (
                float(reported_totals[idx])
                if idx < len(reported_totals)
                and reported_presence[idx]
                and reported_totals[idx] is not None
                and math.isfinite(reported_totals[idx])
                else math.nan
            )

            share_price_value = (
                float(share_prices[idx])
                if idx < len(share_prices)
                and share_price_presence[idx]
                and math.isfinite(share_prices[idx])
                and share_prices[idx] > 0
                else math.nan
            )

            if (
                math.isfinite(reported_value)
                and math.isfinite(share_count_value)
                and share_count_value != 0
            ):
                value_per_share = reported_value / share_count_value
            else:
                value_per_share = math.nan

            if (
                math.isfinite(reported_value)
                and math.isfinite(share_count_value)
                and share_count_value != 0
                and math.isfinite(share_price_value)
                and share_price_value > 0
            ):
                value_share_price_ratio = (
                    reported_value / (share_count_value * share_price_value)
                )
            else:
                value_share_price_ratio = math.nan

            metrics[label_key] = {
                "date": label,
                "share_count": share_count_value,
                "reported_sum": reported_value,
                "value_per_share": value_per_share,
                "value_share_price_ratio": value_share_price_ratio,
                "share_price": share_price_value,
            }

        return metrics

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
            FinanceDataset.NORMALIZATION_SHARE_PRICE_PERIOD,
            self.VALUE_MODE_SHARE_COUNT,
            self.VALUE_MODE_SHARE_PRICE,
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

        if mode == self.VALUE_MODE_SHARE_PRICE:
            if self.normalization_mode == mode:
                return True
            periods = self.visible_periods()
            if not periods:
                periods = self.all_periods()
            if periods and self.datasets:
                for dataset in self.datasets.values():
                    try:
                        dataset.share_prices_for(periods)
                    except ValueError as exc:
                        messagebox.showwarning("Share Price", str(exc))
                        return False
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
        return sorted(periods, key=_period_sort_key)

    @staticmethod
    def _merge_periods(existing: Sequence[str], new_periods: Sequence[str]) -> List[str]:
        ordered_unique: List[str] = []
        seen_tokens: Set[Tuple[str, int, int, str]] = set()
        for label in list(existing) + list(new_periods):
            token = _normalized_period_token(label)
            if token in seen_tokens:
                continue
            seen_tokens.add(token)
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

    def _recalculate_periods_after_mutation(self) -> None:
        if not self.datasets:
            self.periods = None
            self._periods_by_key.clear()
            self._visible_periods = None
            return
        self._periods_by_key.clear()
        for dataset in self.datasets.values():
            self._register_periods(dataset)
        self.periods = self._rebuild_period_sequence()
        if self._visible_periods is None:
            self._visible_periods = list(self.periods)
        else:
            previous = list(self._visible_periods)
            allowed = set(self.periods)
            updated = [label for label in previous if label in allowed]
            if previous and not updated:
                updated = list(self.periods)
            self._visible_periods = updated

    def clear_companies(self) -> None:
        self.datasets.clear()
        self._company_colors.clear()
        self.periods = None
        self._periods_by_key.clear()
        self._visible_periods = None
        self._render_empty()

    def remove_company(self, company: str) -> bool:
        if company not in self.datasets:
            return False
        self.datasets.pop(company, None)
        self._company_colors.pop(company, None)
        self._recalculate_periods_after_mutation()
        if self.datasets:
            self._render()
        else:
            self._render_empty()
        return True

    def has_companies(self) -> bool:
        return bool(self.datasets)

    def current_companies(self) -> List[str]:
        return list(self.datasets.keys())

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
            if self._dataset_normalization_mode in {
                FinanceDataset.NORMALIZATION_SHARE_PRICE,
                FinanceDataset.NORMALIZATION_SHARE_PRICE_PERIOD,
            }:
                normalization_warning = str(exc)
                fallback_mode = FinanceDataset.NORMALIZATION_SHARES
                dataset.set_normalization_mode(fallback_mode)
                for existing in self.datasets.values():
                    existing.set_normalization_mode(fallback_mode)
                self._dataset_normalization_mode = fallback_mode
                if self.normalization_mode in {
                    FinanceDataset.NORMALIZATION_SHARE_PRICE,
                    FinanceDataset.NORMALIZATION_SHARE_PRICE_PERIOD,
                }:
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
            hover_helper.begin_update(self._show_context_menu)
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
                plot_positions = [
                    x_positions[idx]
                    for idx, present in enumerate(share_presence)
                    if present and math.isfinite(share_points[idx])
                ]
                plot_values = [
                    share_points[idx]
                    for idx, present in enumerate(share_presence)
                    if present and math.isfinite(share_points[idx])
                ]
                if not plot_positions:
                    continue
                line = self.axis.plot(
                    plot_positions,
                    plot_values,
                    linestyle=":",
                    marker="o",
                    color=finance_color,
                    markerfacecolor=finance_color,
                    markeredgecolor=finance_color,
                    label=f"{company} Number of shares",
                )
                legend_handles.append(line[0])
                for idx, present in enumerate(share_presence):
                    if not present:
                        continue
                    value = share_points[idx]
                    if not math.isfinite(value):
                        continue
                    period_label = periods[idx] if idx < len(periods) else str(idx)
                    metadata: Dict[str, Any] = {
                        "key": f"{company}_share_count",
                        "type_label": "Number of Shares",
                        "type_value": "Number of Shares",
                        "category": company,
                        "item": "Shares",
                        "period": period_label,
                        "value": value,
                        "company": company,
                        "series": self.VALUE_MODE_SHARE_COUNT,
                    }
                    source = dataset.share_count_source_for(period_label)
                    if source:
                        metadata["pdf_path"] = str(source.pdf_path)
                        metadata["pdf_page"] = source.page
                    hover_helper.add_point_target(
                        float(x_positions[idx]), float(value), metadata
                    )

            if period_indices:
                self.axis.set_xticks(period_indices)
                self.axis.set_xticklabels(periods, rotation=45, ha="right")
                self.axis.set_xlim(-0.5, len(period_indices) - 0.5)
            else:
                self.axis.set_xlim(-0.6, 0.6)
                self.axis.set_xticks([])

        elif self.normalization_mode == self.VALUE_MODE_SHARE_PRICE:
            hover_helper.begin_update(None)
            x_positions = period_indices
            for company, dataset in self.datasets.items():
                finance_color, _ = self._company_colors.setdefault(
                    company, self._next_color_pair()
                )
                try:
                    share_prices, share_presence = dataset.share_prices_for(periods)
                except ValueError:
                    continue
                if not any(share_presence):
                    continue
                price_points = [
                    share_prices[idx] if share_presence[idx] else math.nan
                    for idx in range(len(share_prices))
                ]
                plot_positions = [
                    x_positions[idx]
                    for idx, present in enumerate(share_presence)
                    if present and math.isfinite(price_points[idx])
                ]
                plot_values = [
                    price_points[idx]
                    for idx, present in enumerate(share_presence)
                    if present and math.isfinite(price_points[idx])
                ]
                if not plot_positions:
                    continue
                share_price_metrics = self._compute_period_metrics(
                    dataset,
                    periods,
                    series=FinanceDataset.FINANCE_LABEL,
                )
                line = self.axis.plot(
                    plot_positions,
                    plot_values,
                    linestyle="-",
                    marker="o",
                    color=finance_color,
                    markerfacecolor=finance_color,
                    markeredgecolor=finance_color,
                    linewidth=1.6,
                    markersize=6,
                    label=f"{company} Share Price",
                )
                legend_handles.append(line[0])
                for idx, present in enumerate(share_presence):
                    if not present:
                        continue
                    value = price_points[idx]
                    if not math.isfinite(value):
                        continue
                    period_label = periods[idx] if idx < len(periods) else str(idx)
                    metadata = {
                        "key": f"{company}_share_price",
                        "type_label": "Share Price",
                        "type_value": "Share Price",
                        "category": company,
                        "item": dataset.share_price_symbol or "Share Price",
                        "period": str(period_label),
                        "value": value,
                        "company": company,
                        "series": self.VALUE_MODE_SHARE_PRICE,
                    }
                    extra = share_price_metrics.get(str(period_label))
                    if extra:
                        metadata.update(extra)
                    hover_helper.add_point_target(
                        float(x_positions[idx]), float(value), metadata
                    )

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
                finance_totals_all = [0.0] * len(periods)
                if finance_segment_values:
                    for _, segment_values in finance_segment_values:
                        for idx, value in enumerate(segment_values):
                            finance_totals_all[idx] += value
                finance_positions_active = [
                    finance_positions[idx] for idx in finance_active_indices
                ]
                finance_periods_active = [
                    periods[idx] for idx in finance_active_indices
                ]
                finance_totals_active = [
                    finance_totals_all[idx] for idx in finance_active_indices
                ]
                finance_pos_totals = [0.0] * len(finance_active_indices)
                finance_neg_totals = [0.0] * len(finance_active_indices)
                finance_metrics = self._compute_period_metrics(
                    dataset,
                    finance_periods_active,
                    series=FinanceDataset.FINANCE_LABEL,
                )

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
                            linestyle="-",
                        )
                        hover_helper.add_segment(
                            rectangles,
                            segment,
                            finance_periods_active,
                            company,
                            values=filtered_values,
                            series_label=FinanceDataset.FINANCE_LABEL,
                            totals=finance_totals_active,
                            metadata_overrides=finance_metrics,
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
                income_totals_all = [0.0] * len(periods)
                if income_segment_values:
                    for _, segment_values in income_segment_values:
                        for idx, value in enumerate(segment_values):
                            income_totals_all[idx] += value
                income_positions_active = [
                    income_positions[idx] for idx in income_active_indices
                ]
                income_periods_active = [
                    periods[idx] for idx in income_active_indices
                ]
                income_totals_active = [
                    income_totals_all[idx] for idx in income_active_indices
                ]
                income_pos_totals = [0.0] * len(income_active_indices)
                income_neg_totals = [0.0] * len(income_active_indices)
                income_metrics = self._compute_period_metrics(
                    dataset,
                    income_periods_active,
                    series=FinanceDataset.INCOME_LABEL,
                )

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
                            edgecolor=finance_color,
                            linewidth=2.4,
                            linestyle="--",
                        )
                        hover_helper.add_segment(
                            rectangles,
                            segment,
                            income_periods_active,
                            company,
                            values=filtered_values,
                            series_label=FinanceDataset.INCOME_LABEL,
                            totals=income_totals_active,
                            metadata_overrides=income_metrics,
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
                            linestyle="-",
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
                            edgecolor=finance_color,
                            linewidth=1.0,
                            linestyle="--",
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
            hover_helper.begin_update(self._show_context_menu)
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
                    and math.isfinite(finance_totals[idx])
                    for idx in range(len(finance_totals))
                ]
                finance_points = [
                    value if finance_included[idx] else math.nan
                    for idx, value in enumerate(finance_totals)
                ]
                finance_metrics = self._compute_period_metrics(
                    dataset,
                    periods,
                    series=FinanceDataset.FINANCE_LABEL,
                )
                plot_positions = [
                    x_positions[idx]
                    for idx, included in enumerate(finance_included)
                    if included and math.isfinite(finance_points[idx])
                ]
                plot_values = [
                    finance_points[idx]
                    for idx, included in enumerate(finance_included)
                    if included and math.isfinite(finance_points[idx])
                ]
                if plot_positions:
                    line = self.axis.plot(
                        plot_positions,
                        plot_values,
                        linestyle="-",
                        marker="o",
                        color=company_color,
                        markerfacecolor=company_color,
                        markeredgecolor=company_color,
                        label=f"{company} {FinanceDataset.FINANCE_LABEL}",
                    )
                    legend_handles.append(line[0])
                    for idx, present in enumerate(finance_included):
                        if not present:
                            continue
                        value = finance_points[idx]
                        if not math.isfinite(value):
                            continue
                        period_label = periods[idx] if idx < len(periods) else str(idx)
                        metadata: Dict[str, Any] = {
                            "key": f"{company}_finance_total",
                            "type_label": FinanceDataset.FINANCE_LABEL,
                            "type_value": FinanceDataset.FINANCE_LABEL,
                            "category": company,
                            "item": "Total",
                            "period": str(period_label),
                            "value": float(value),
                            "company": company,
                            "series": FinanceDataset.FINANCE_LABEL,
                        }
                        source = dataset.source_for_series(
                            FinanceDataset.FINANCE_LABEL, str(period_label)
                        )
                        if source:
                            metadata["pdf_path"] = str(source.pdf_path)
                            metadata["pdf_page"] = source.page
                        extra = finance_metrics.get(str(period_label))
                        if extra:
                            metadata.update(extra)
                        hover_helper.add_point_target(
                            float(x_positions[idx]), float(value), metadata
                        )

                income_totals, income_presence = dataset.aggregate_totals(
                    periods, series=FinanceDataset.INCOME_LABEL
                )
                income_included = [
                    income_presence[idx]
                    and math.isfinite(income_totals[idx])
                    for idx in range(len(income_totals))
                ]
                income_points = [
                    value if income_included[idx] else math.nan
                    for idx, value in enumerate(income_totals)
                ]
                income_metrics = self._compute_period_metrics(
                    dataset,
                    periods,
                    series=FinanceDataset.INCOME_LABEL,
                )
                plot_positions = [
                    x_positions[idx]
                    for idx, included in enumerate(income_included)
                    if included and math.isfinite(income_points[idx])
                ]
                plot_values = [
                    income_points[idx]
                    for idx, included in enumerate(income_included)
                    if included and math.isfinite(income_points[idx])
                ]
                if plot_positions:
                    line = self.axis.plot(
                        plot_positions,
                        plot_values,
                        linestyle="--",
                        marker="o",
                        color=company_color,
                        markerfacecolor=company_color,
                        markeredgecolor=company_color,
                        label=f"{company} {FinanceDataset.INCOME_LABEL}",
                    )
                    legend_handles.append(line[0])
                    for idx, present in enumerate(income_included):
                        if not present:
                            continue
                        value = income_points[idx]
                        if not math.isfinite(value):
                            continue
                        period_label = periods[idx] if idx < len(periods) else str(idx)
                        metadata = {
                            "key": f"{company}_income_total",
                            "type_label": FinanceDataset.INCOME_LABEL,
                            "type_value": FinanceDataset.INCOME_LABEL,
                            "category": company,
                            "item": "Total",
                            "period": str(period_label),
                            "value": float(value),
                            "company": company,
                            "series": FinanceDataset.INCOME_LABEL,
                        }
                        source = dataset.source_for_series(
                            FinanceDataset.INCOME_LABEL, str(period_label)
                        )
                        if source:
                            metadata["pdf_path"] = str(source.pdf_path)
                            metadata["pdf_page"] = source.page
                        extra = income_metrics.get(str(period_label))
                        if extra:
                            metadata.update(extra)
                        hover_helper.add_point_target(
                            float(x_positions[idx]), float(value), metadata
                        )

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
        elif self.normalization_mode == self.VALUE_MODE_SHARE_PRICE:
            self.axis.yaxis.set_major_formatter(self._share_price_formatter)
            self.axis.set_ylabel("Share Price")
        elif self.normalization_mode == FinanceDataset.NORMALIZATION_SHARES:
            self.axis.yaxis.set_major_formatter(self._per_share_formatter)
            self.axis.set_ylabel("Value/Share")
        elif self.normalization_mode == FinanceDataset.NORMALIZATION_SHARE_PRICE:
            self.axis.yaxis.set_major_formatter(self._share_price_formatter)
            self.axis.set_ylabel("Share Price Multiple")
        elif self.normalization_mode == FinanceDataset.NORMALIZATION_SHARE_PRICE_PERIOD:
            self.axis.yaxis.set_major_formatter(self._share_price_ratio_formatter)
            self.axis.set_ylabel("Value/(Share*Price)")
        else:
            self.axis.yaxis.set_major_formatter(self._reported_formatter)
            self.axis.set_ylabel("Reported Value")
        if self.datasets:
            companies_list = ", ".join(self.datasets.keys())
            if self.normalization_mode == self.VALUE_MODE_SHARE_COUNT:
                self.axis.set_title(f"Number of Shares — {companies_list}")
            elif self.normalization_mode == self.VALUE_MODE_SHARE_PRICE:
                self.axis.set_title(f"Share Price — {companies_list}")
            else:
                self.axis.set_title(f"Finance vs Income Statement — {companies_list}")
        else:
            if self.normalization_mode == self.VALUE_MODE_SHARE_COUNT:
                self.axis.set_title("Number of Shares")
            elif self.normalization_mode == self.VALUE_MODE_SHARE_PRICE:
                self.axis.set_title("Share Price")
            else:
                self.axis.set_title("Finance vs Income Statement")

        if legend_handles:
            legend_kwargs = {
                "loc": "center left",
                "bbox_to_anchor": (1.02, 0.5),
                "borderaxespad": 0.0,
            }
            self.axis.legend(handles=legend_handles, **legend_kwargs)

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
        metadata = self._context_metadata
        if not metadata:
            return

        def _parse_page(value: Any) -> Optional[int]:
            try:
                number = int(value) if value is not None else None
            except (TypeError, ValueError):
                return None
            if number is None or number <= 0:
                return None
            return number

        def _context_source(path_value: Any, page_value: Any) -> Optional[SegmentSource]:
            if not path_value:
                return None
            pdf_path = Path(str(path_value))
            return SegmentSource(pdf_path=pdf_path, page=_parse_page(page_value))

        company_value = metadata.get("company")
        period_value = metadata.get("period")
        period_label = str(period_value) if period_value is not None else None
        dataset = None
        if company_value is not None:
            dataset = self.datasets.get(str(company_value))
        series_value = metadata.get("series")
        pdf_path_value = metadata.get("pdf_path")
        page_value = metadata.get("pdf_page")

        sources_to_open: List[Tuple[str, SegmentSource, str]] = []
        paired_sources: Optional[
            Tuple[Tuple[str, SegmentSource], Tuple[str, SegmentSource]]
        ] = None

        if series_value == self.VALUE_MODE_SHARE_COUNT:
            if dataset is not None and period_label is not None:
                share_source = dataset.share_count_source_for(period_label)
                if share_source is not None:
                    sources_to_open.append(("Number of Shares", share_source, "full"))
            if not sources_to_open:
                fallback = _context_source(pdf_path_value, page_value)
                if fallback is not None:
                    sources_to_open.append(("Number of Shares", fallback, "full"))
                else:
                    messagebox.showinfo(
                        "Open PDF", "No PDF is associated with this data point."
                    )
                    self._context_metadata = None
                    return
        else:
            finance_source: Optional[SegmentSource] = None
            income_source: Optional[SegmentSource] = None
            if dataset is not None and period_label is not None:
                finance_source = dataset.source_for_series(
                    FinanceDataset.FINANCE_LABEL, period_label
                )
                income_source = dataset.source_for_series(
                    FinanceDataset.INCOME_LABEL, period_label
                )
            if finance_source or income_source:
                if (
                    finance_source is not None
                    and finance_source.page is not None
                    and income_source is not None
                    and income_source.page is not None
                ):
                    paired_sources = (
                        (FinanceDataset.FINANCE_LABEL, finance_source),
                        (FinanceDataset.INCOME_LABEL, income_source),
                    )
                else:
                    if finance_source is not None:
                        sources_to_open.append(
                            (FinanceDataset.FINANCE_LABEL, finance_source, "full")
                        )
                    if income_source is not None:
                        sources_to_open.append(
                            (FinanceDataset.INCOME_LABEL, income_source, "full")
                        )
            else:
                fallback = _context_source(pdf_path_value, page_value)
                if fallback is not None:
                    label_text = (
                        metadata.get("type_label")
                        or metadata.get("type_value")
                        or "Entry"
                    )
                    sources_to_open.append((str(label_text), fallback, "full"))
                else:
                    messagebox.showinfo(
                        "Open PDF", "No PDF is associated with this bar segment."
                    )
                    self._context_metadata = None
                    return

        if paired_sources is not None:
            company_label = metadata.get("company") or "Company"
            period_label = metadata.get("period") or "Period"
            title = f"{company_label} — {period_label}"
            if not self._open_pdf_pair(paired_sources[0], paired_sources[1], title):
                for label_text, source, placement in (
                    (paired_sources[0][0], paired_sources[0][1], "left"),
                    (paired_sources[1][0], paired_sources[1][1], "right"),
                ):
                    sources_to_open.append((label_text, source, placement))

        for label_text, source, placement in sources_to_open:
            self._open_pdf_source(source, str(label_text), placement)

        self._context_metadata = None

    def _open_pdf_source(
        self, source: SegmentSource, label: str, placement: str = "full"
    ) -> None:
        pdf_path = source.pdf_path
        page = source.page
        if not pdf_path.exists():
            messagebox.showwarning(
                "Open PDF",
                f"The PDF for {label} could not be found at {pdf_path}.",
            )
            return
        if page is None:
            messagebox.showinfo(
                "Open PDF",
                f"No page number is available for {label}. The PDF will open to its first page.",
            )
            open_pdf(pdf_path, None)
            return
        if fitz is None:
            messagebox.showwarning(
                "Open PDF",
                (
                    "Viewing individual PDF pages requires the PyMuPDF package (import name 'fitz').\n"
                    f"The full PDF for {label} will be opened instead."
                ),
            )
            open_pdf(pdf_path, page)
            return
        try:
            viewer = PdfPageViewer(self, pdf_path, page, placement=placement)
            viewer.focus_set()
        except Exception as exc:
            messagebox.showwarning(
                "Open PDF",
                f"Could not display page {page} from {pdf_path.name} ({label}): {exc}\n"
                "The full PDF will be opened instead.",
            )
            open_pdf(pdf_path, page)

    def _open_pdf_pair(
        self,
        left: Tuple[str, SegmentSource],
        right: Tuple[str, SegmentSource],
        title: str,
    ) -> bool:
        left_source = left[1]
        right_source = right[1]
        if left_source.page is None or right_source.page is None:
            return False
        if not left_source.pdf_path.exists() or not right_source.pdf_path.exists():
            return False
        if fitz is None:
            return False
        try:
            viewer = PdfSplitViewer(self, left, right, title)
            viewer.focus_set()
            return True
        except Exception as exc:
            messagebox.showwarning(
                "Open PDF",
                (
                    "Could not display the Finance and Income PDFs together. "
                    f"The individual documents will be opened instead.\nDetails: {exc}"
                ),
            )
            return False


class BarHoverHelper:
    """Manage hover annotations for bar segments."""

    def __init__(self, axis) -> None:
        self.axis = axis
        self._rectangles: List[Tuple[Rectangle, Dict[str, Any]]] = []
        self._point_targets: List[Tuple[Tuple[float, float], Dict[str, Any]]] = []
        self._annotation = self._create_annotation()
        self._canvas: Optional[FigureCanvasTkAgg] = None
        self._tk_widget: Optional[tk.Widget] = None
        self._connection_id: Optional[int] = None
        self._button_connection_id: Optional[int] = None
        self._motion_binding_id: Optional[str] = None
        self._leave_binding_id: Optional[str] = None
        self._active_rectangle: Optional[Rectangle] = None
        self._active_point: Optional[Tuple[float, float]] = None
        self._point_hit_radius: float = 12.0
        self._context_callback: Optional[Callable[[Dict[str, Any], Any], None]] = None

    def _create_annotation(self):
        annotation = self.axis.annotate(
            "",
            xy=(0, 0),
            xytext=(12, 12),
            textcoords="offset points",
            bbox={"boxstyle": "round", "fc": "#f9f9f9", "ec": "#555555", "alpha": 0.95},
            arrowprops={"arrowstyle": "->", "color": "#555555"},
        )
        # Ensure the tooltip annotation renders above all bar segments so it is always visible.
        annotation.set_zorder(1000)
        bbox_patch = annotation.get_bbox_patch()
        if bbox_patch is not None:
            bbox_patch.set_zorder(1000)
        if hasattr(annotation, "arrow_patch") and annotation.arrow_patch is not None:
            annotation.arrow_patch.set_zorder(1000)
        annotation.set_visible(False)
        return annotation

    def _ensure_annotation(self) -> None:
        if self._annotation.axes is self.axis:
            return
        self._annotation = self._create_annotation()

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
        self._ensure_annotation()
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
        self._ensure_annotation()
        self._set_context_callback(context_callback)
        self.reset()

    def reset(self) -> None:
        self._ensure_annotation()
        self._rectangles.clear()
        self._point_targets.clear()
        self._active_rectangle = None
        self._active_point = None
        self._annotation.set_visible(False)

    def _on_tk_motion(self, tk_event: tk.Event) -> None:  # type: ignore[name-defined]
        if self._canvas is None:
            return
        canvas_backend = getattr(self._canvas.figure, "canvas", None)
        if canvas_backend is None:
            return

        motion_handler = getattr(canvas_backend, "motion_notify_event", None)
        if motion_handler is None:
            return

        try:
            # Matplotlib 3.8+ accepts the original Tk event and performs the
            # coordinate translation internally, ensuring consistency with the
            # values reported to motion_notify_event callbacks.
            motion_handler(tk_event)
            return
        except TypeError:
            # Older Matplotlib versions expect explicit canvas coordinates. If
            # the call above raises, fall back to the previous manual
            # conversion routine so legacy environments continue to work.
            widget = (
                self._tk_widget if self._tk_widget is not None else getattr(tk_event, "widget", None)
            )
            if widget is None:
                return
            width = widget.winfo_width()
            height = widget.winfo_height()
            if height <= 0 or width <= 0:
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
        series_label: Optional[str] = None,
        totals: Optional[Sequence[float]] = None,
        metadata_overrides: Optional[Dict[str, Dict[str, Any]]] = None,
    ) -> None:
        for index, rect in enumerate(bar_container.patches):
            value_list = values if values is not None else segment.values
            metadata: Dict[str, Any] = {
                "key": segment.key,
                "type_label": segment.type_label,
                "type_value": segment.type_value,
                "category": segment.category,
                "item": segment.item,
                "period": str(periods[index]) if index < len(periods) else str(index),
                "value": value_list[index] if index < len(value_list) else 0.0,
                "company": company,
            }
            if series_label:
                metadata["series"] = series_label
            period_label = metadata["period"]
            if period_label is not None and segment.sources:
                try:
                    source = _lookup_period_value(segment.sources, str(period_label))
                except KeyError:
                    source = None
                if source:
                    metadata["pdf_path"] = str(source.pdf_path)
                    metadata["pdf_page"] = source.page
            if totals is not None and index < len(totals):
                metadata["sum_value"] = totals[index]
            if metadata_overrides:
                extra = metadata_overrides.get(str(period_label))
                if extra:
                    metadata.update(extra)
            self._rectangles.append((rect, metadata))

    def add_point_target(self, x: float, y: float, metadata: Dict[str, Any]) -> None:
        self._point_targets.append(((float(x), float(y)), metadata))

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
        date_value = metadata.get("date") or metadata.get("period") or "—"
        share_count = metadata.get("share_count")
        reported_sum = metadata.get("reported_sum")
        value_per_share = metadata.get("value_per_share")
        value_share_price_ratio = metadata.get("value_share_price_ratio")
        share_price_value = metadata.get("share_price")

        def _format_scientific(raw: Any) -> str:
            if isinstance(raw, (int, float)) and math.isfinite(raw):
                return f"{raw:.3e}"
            return "—"

        def _format_decimal(raw: Any) -> str:
            if isinstance(raw, (int, float)) and math.isfinite(raw):
                return f"{raw:,.2f}"
            return "—"

        def _format_ratio(raw: Any) -> str:
            if isinstance(raw, (int, float)) and math.isfinite(raw):
                return f"{raw:,.6f}"
            return "—"

        lines = [
            f"TYPE: {type_value}",
            f"CATEGORY: {category}",
            f"ITEM: {item}",
            f"DATE: {date_value}",
            f"AMOUNT: {formatted_value}",
            f"REPORTED SUM: {_format_scientific(reported_sum)}",
            f"NUMBER OF SHARES: {_format_scientific(share_count)}",
            f"VALUE/SHARE: {_format_decimal(value_per_share)}",
            f"SHARE PRICE: {_format_decimal(share_price_value)}",
            f"VALUE/(SHARE*PRICE): {_format_ratio(value_share_price_ratio)}",
        ]

        return "\n".join(lines)

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
        self._active_point = None
        if self._canvas is not None:
            self._canvas.draw_idle()

    def _find_point_hit(
        self, display_x: float, display_y: float
    ) -> Optional[Tuple[float, float, Dict[str, Any]]]:
        if not self._point_targets:
            return None
        transform = self.axis.transData.transform
        threshold_sq = self._point_hit_radius * self._point_hit_radius
        closest: Optional[Tuple[float, float, Dict[str, Any]]] = None
        best_distance = threshold_sq
        for (x, y), metadata in self._point_targets:
            if not math.isfinite(x) or not math.isfinite(y):
                continue
            pixel_x, pixel_y = transform((x, y))
            dx = display_x - pixel_x
            dy = display_y - pixel_y
            distance_sq = dx * dx + dy * dy
            if distance_sq <= best_distance:
                closest = (x, y, metadata)
                best_distance = distance_sq
        return closest

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

        point_hit = self._find_point_hit(display_x, display_y)
        if point_hit is not None:
            point_x, point_y, metadata = point_hit
            if (
                self._active_point != (point_x, point_y)
                or not self._annotation.get_visible()
            ):
                self._annotation.set_text(self._format_text(metadata))
                self._annotation.set_visible(True)
            self._active_rectangle = None
            self._active_point = (point_x, point_y)
            raw_value = metadata.get("value")
            if isinstance(raw_value, (int, float)) and math.isfinite(raw_value):
                numeric_value = float(raw_value)
            else:
                numeric_value = 0.0
            self._position_annotation(point_x, point_y, numeric_value)
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
        if event.x is None or event.y is None:
            return None
        point_hit = self._find_point_hit(float(event.x), float(event.y))
        if point_hit is not None:
            return point_hit[2]
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
        self.period_all_var = tk.BooleanVar(value=True)
        self._updating_period_checks = False

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
        self.remove_button = ttk.Button(
            controls, text="Remove", command=self._remove_selected_company
        )
        self.remove_button.pack(side=tk.LEFT, padx=(6, 0))
        self.clear_button = ttk.Button(
            controls, text="Clear All", command=self._clear_all_companies
        )
        self.clear_button.pack(side=tk.LEFT, padx=(6, 0))

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
            text="Value/Share",
            variable=self.normalization_var,
            value=FinanceDataset.NORMALIZATION_SHARES,
            command=self._on_normalization_change,
        ).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Radiobutton(
            values_frame,
            text="Value/(Share*Price)",
            variable=self.normalization_var,
            value=FinanceDataset.NORMALIZATION_SHARE_PRICE_PERIOD,
            command=self._on_normalization_change,
        ).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Radiobutton(
            values_frame,
            text="Share price",
            variable=self.normalization_var,
            value=FinancePlotFrame.VALUE_MODE_SHARE_PRICE,
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
        self.period_all_check = ttk.Checkbutton(
            self.periods_frame,
            text="All dates",
            variable=self.period_all_var,
            command=self._toggle_all_periods,
        )
        self.period_all_check.pack(side=tk.LEFT, padx=(6, 6))
        self.period_checks_frame = ttk.Frame(self.periods_frame)
        self.period_checks_frame.pack(side=tk.LEFT, padx=(6, 0))

        self.plot_frame = FinancePlotFrame(outer)
        self.plot_frame.pack(fill=tk.BOTH, expand=True)
        self.plot_frame.set_normalization_mode(self.normalization_var.get())

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
        if self._updating_period_checks:
            return
        self._apply_period_filters()
        self._update_period_all_state()

    def _apply_period_filters(self) -> None:
        active_periods = [
            label for label, var in self.period_vars.items() if var.get()
        ]
        if not active_periods and self.plot_frame.all_periods():
            self.plot_frame.set_visible_periods([])
            return
        self.plot_frame.set_visible_periods(active_periods)

    def _toggle_all_periods(self) -> None:
        if self._updating_period_checks:
            return
        desired = self.period_all_var.get()
        self._updating_period_checks = True
        for var in self.period_vars.values():
            var.set(desired)
        self._updating_period_checks = False
        self._apply_period_filters()
        self._update_period_all_state()

    def _update_period_all_state(self) -> None:
        if not hasattr(self, "period_all_check"):
            return
        if not self.period_vars:
            self._updating_period_checks = True
            self.period_all_var.set(False)
            self._updating_period_checks = False
            self.period_all_check.state(["disabled"])
            return
        self.period_all_check.state(["!disabled"])
        all_selected = all(var.get() for var in self.period_vars.values())
        self._updating_period_checks = True
        self.period_all_var.set(all_selected)
        self._updating_period_checks = False

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

        self._update_period_all_state()
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

    def _remove_selected_company(self) -> None:
        company = self.company_var.get().strip()
        if not company:
            messagebox.showinfo(
                "Remove Company", "Choose a company to remove from the chart."
            )
            return
        if not self.plot_frame.remove_company(company):
            messagebox.showinfo(
                "Remove Company", f"{company} is not currently on the chart."
            )
            return
        self._refresh_period_controls()

    def _clear_all_companies(self) -> None:
        if not self.plot_frame.has_companies():
            messagebox.showinfo("Clear Chart", "There are no companies to clear from the chart.")
            return
        self.plot_frame.clear_companies()
        self._refresh_period_controls()

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
