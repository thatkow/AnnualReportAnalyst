"""Standalone finance analyst application for combined CSV summaries.

This lightweight Tkinter application lets reviewers choose a company that has a
``combined.csv`` file in ``companies/<company>/``. When a company is opened, the
app creates a new tab containing side-by-side bar charts for Finance and Income
Statement sections. Each bar is stacked by the underlying rows from the
``combined.csv`` file so reviewers can see how individual entries contribute to
each reporting period.
"""

from __future__ import annotations

import csv
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import matplotlib

# Ensure TkAgg is used when embedding plots inside Tkinter widgets.
matplotlib.use("TkAgg")

from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from matplotlib import cm, colors
from matplotlib.patches import Patch, Rectangle

import tkinter as tk
from tkinter import messagebox, ttk


BASE_FIELDS = {"Type", "Category", "Item", "Note"}


@dataclass
class RowSegment:
    """Represent a single CSV row broken into period values for plotting."""

    key: str
    type_label: str
    type_value: str
    category: str
    item: str
    values: List[float]


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
        self.headers: List[str] = []
        self.periods: List[str] = []
        self.finance_segments: List[RowSegment] = []
        self.income_segments: List[RowSegment] = []
        self._color_cache: Dict[str, str] = {}
        self._load()

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

            for row_number, row in enumerate(reader, start=1):
                if type_index >= len(row):
                    continue
                raw_type = row[type_index].strip()
                type_value = raw_type.lower()
                series_label = self._TYPE_TO_LABEL.get(type_value)
                if not series_label:
                    # Skip other sections such as Shares.
                    continue

                values: List[float] = []
                for column_index in data_indices:
                    if column_index >= len(row):
                        values.append(0.0)
                        continue
                    values.append(self._clean_numeric(row[column_index]))

                if not any(value != 0 for value in values):
                    continue

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
                    values=values,
                )
                target = (
                    self.finance_segments
                    if series_label == self.FINANCE_LABEL
                    else self.income_segments
                )
                target.append(segment)

    def has_data(self) -> bool:
        return bool(self.finance_segments or self.income_segments)

    def color_for_key(self, key: str) -> str:
        if key not in self._color_cache:
            palette = cm.get_cmap("tab20").colors
            index = len(self._color_cache) % len(palette)
            self._color_cache[key] = colors.to_hex(palette[index])
        return self._color_cache[key]


class FinanceNotebook(ttk.Notebook):
    """Notebook that renders stacked bar plots for each company."""

    def __init__(self, parent: tk.Widget) -> None:
        super().__init__(parent)
        self._open_tabs: Dict[str, tk.Widget] = {}

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

    def show_company(self, company: str, dataset: FinanceDataset) -> None:
        if company in self._open_tabs:
            self.select(self._open_tabs[company])
            return

        frame = ttk.Frame(self)
        figure = Figure(figsize=(8, 5), dpi=100)
        axis = figure.add_subplot(111)
        periods = dataset.periods
        period_indices = list(range(len(periods)))
        bar_width = 0.38
        finance_positions = [index - bar_width / 2 for index in period_indices]
        income_positions = [index + bar_width / 2 for index in period_indices]

        hover_helper = BarHoverHelper(axis)

        finance_pos_totals = [0.0] * len(periods)
        finance_neg_totals = [0.0] * len(periods)
        income_pos_totals = [0.0] * len(periods)
        income_neg_totals = [0.0] * len(periods)

        for segment in dataset.finance_segments:
            bottoms = self._compute_bottoms(segment.values, finance_pos_totals, finance_neg_totals)
            rectangles = axis.bar(
                finance_positions,
                segment.values,
                width=bar_width,
                bottom=bottoms,
                color=dataset.color_for_key(segment.key),
                edgecolor="#1f77b4",
                linewidth=0.8,
            )
            hover_helper.add_segment(rectangles, segment, periods)

        for segment in dataset.income_segments:
            bottoms = self._compute_bottoms(segment.values, income_pos_totals, income_neg_totals)
            rectangles = axis.bar(
                income_positions,
                segment.values,
                width=bar_width,
                bottom=bottoms,
                color=dataset.color_for_key(segment.key),
                edgecolor="#ff7f0e",
                linewidth=0.8,
            )
            hover_helper.add_segment(rectangles, segment, periods)

        axis.axhline(0, color="#333333", linewidth=0.8)
        axis.set_ylabel("Value")
        axis.set_title(f"{company} Finance vs Income Statement")
        axis.set_xticks(period_indices)
        axis.set_xticklabels(periods, rotation=45, ha="right")
        axis.set_xlim(-0.6, len(periods) - 0.4 if periods else 0.6)

        legend_handles = [
            Patch(facecolor="white", edgecolor="#1f77b4", linewidth=1.0, label=FinanceDataset.FINANCE_LABEL),
            Patch(facecolor="white", edgecolor="#ff7f0e", linewidth=1.0, label=FinanceDataset.INCOME_LABEL),
        ]
        axis.legend(handles=legend_handles, loc="upper left")
        figure.tight_layout()

        canvas = FigureCanvasTkAgg(figure, master=frame)
        canvas.draw()
        canvas_widget = canvas.get_tk_widget()
        canvas_widget.pack(fill=tk.BOTH, expand=True)

        hover_helper.attach(canvas)
        frame._hover_helper = hover_helper  # type: ignore[attr-defined]

        self.add(frame, text=company)
        self._open_tabs[company] = frame
        self.select(frame)


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
        self._annotation.set_visible(False)
        self._canvas: Optional[FigureCanvasTkAgg] = None
        self._connection_id: Optional[int] = None
        self._active_rectangle: Optional[Rectangle] = None

    def add_segment(self, bar_container, segment: RowSegment, periods: Sequence[str]) -> None:
        for index, rect in enumerate(bar_container.patches):
            metadata: Dict[str, Any] = {
                "key": segment.key,
                "type_label": segment.type_label,
                "type_value": segment.type_value,
                "category": segment.category,
                "item": segment.item,
                "period": periods[index] if index < len(periods) else str(index),
                "value": segment.values[index] if index < len(segment.values) else 0.0,
            }
            self._rectangles.append((rect, metadata))

    def attach(self, canvas: FigureCanvasTkAgg) -> None:
        if self._connection_id is not None and self._canvas is not None:
            self._canvas.mpl_disconnect(self._connection_id)
        self._canvas = canvas
        self._connection_id = canvas.mpl_connect("motion_notify_event", self._on_motion)

    def _format_text(self, metadata: Dict[str, Any]) -> str:
        value = metadata.get("value", 0.0)
        formatted_value = f"{value:,.2f}"
        type_value = metadata.get("type_value") or metadata.get("type_label") or "—"
        category = metadata.get("category") or "—"
        item = metadata.get("item") or "—"
        key = metadata.get("key") or "—"
        period = metadata.get("period") or "Period"
        return (
            f"Key: {key}\n"
            f"Type: {type_value}\n"
            f"Category: {category}\n"
            f"Item: {item}\n"
            f"{period}: {formatted_value}"
        )

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
            value = metadata.get("value", 0.0)
            if value >= 0:
                anchor_y = rect.get_y() + rect.get_height()
            else:
                anchor_y = rect.get_y()
            self._annotation.xy = (center_x, anchor_y)
            self._canvas.draw_idle()
            return

        if self._annotation.get_visible():
            self._annotation.set_visible(False)
            self._canvas.draw_idle()
        self._active_rectangle = None


class FinanceAnalystApp:
    """Main window controller for the thatkowfinance_data application."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("thatkowfinance_data")
        self._maximize_window()
        self.base_dir = Path(__file__).resolve().parent
        self.companies_dir = self.base_dir / "companies"

        self.company_var = tk.StringVar()

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

        self.open_button = ttk.Button(controls, text="Open", command=self._open_selected_company)
        self.open_button.pack(side=tk.LEFT)

        self.notebook = FinanceNotebook(outer)
        self.notebook.pack(fill=tk.BOTH, expand=True)

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
        self.notebook.show_company(company, dataset)


def main() -> None:
    root = tk.Tk()
    app = FinanceAnalystApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
