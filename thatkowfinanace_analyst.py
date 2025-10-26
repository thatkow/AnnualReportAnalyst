"""Standalone finance analyst application for combined CSV summaries.

This lightweight Tkinter application lets reviewers choose a company that has a
``combined.csv`` file in ``companies/<company>/``. When a company is opened, the
app creates a new tab containing a stacked bar chart comparing Finance
(``Financial`` rows in the CSV) with the Income Statement (``Income`` rows). The
chart aggregates each numeric column in the CSV so reviewers can see how totals
change across reporting periods.
"""

from __future__ import annotations

import csv
import re
from pathlib import Path
from typing import Dict, List, Sequence

import matplotlib

# Ensure TkAgg is used when embedding plots inside Tkinter widgets.
matplotlib.use("TkAgg")

from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

import tkinter as tk
from tkinter import messagebox, ttk


BASE_FIELDS = {"Type", "Category", "Item", "Note"}


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
        self.finance_totals: List[float] = []
        self.income_totals: List[float] = []
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

    def _init_totals(self, size: int) -> None:
        self.finance_totals = [0.0] * size
        self.income_totals = [0.0] * size

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

            data_indices: List[int] = []
            for idx, column_name in enumerate(header):
                if column_name in BASE_FIELDS:
                    continue
                data_indices.append(idx)
                self.periods.append(column_name)

            if not data_indices:
                raise ValueError("Combined CSV does not contain numeric data columns")

            self._init_totals(len(data_indices))

            for row in reader:
                if type_index >= len(row):
                    continue
                type_value = row[type_index].strip().lower()
                series_label = self._TYPE_TO_LABEL.get(type_value)
                if not series_label:
                    # Skip other sections such as Shares.
                    continue
                target = self.finance_totals if series_label == self.FINANCE_LABEL else self.income_totals
                for position, column_index in enumerate(data_indices):
                    if column_index >= len(row):
                        continue
                    target[position] += self._clean_numeric(row[column_index])

    def has_data(self) -> bool:
        return any(self.finance_totals) or any(self.income_totals)


class FinanceNotebook(ttk.Notebook):
    """Notebook that renders stacked bar plots for each company."""

    def __init__(self, parent: tk.Widget) -> None:
        super().__init__(parent)
        self._open_tabs: Dict[str, tk.Widget] = {}

    def show_company(self, company: str, dataset: FinanceDataset) -> None:
        if company in self._open_tabs:
            self.select(self._open_tabs[company])
            return

        frame = ttk.Frame(self)
        figure = Figure(figsize=(8, 5), dpi=100)
        axis = figure.add_subplot(111)
        periods = dataset.periods
        finance = dataset.finance_totals
        income = dataset.income_totals

        axis.bar(periods, finance, label=FinanceDataset.FINANCE_LABEL, color="#1f77b4")
        axis.bar(
            periods,
            income,
            bottom=finance,
            label=FinanceDataset.INCOME_LABEL,
            color="#ff7f0e",
        )
        axis.set_ylabel("Total")
        axis.set_title(f"{company} Finance vs Income Statement")
        axis.legend()
        axis.tick_params(axis="x", rotation=45)
        figure.tight_layout()

        canvas = FigureCanvasTkAgg(figure, master=frame)
        canvas.draw()
        canvas_widget = canvas.get_tk_widget()
        canvas_widget.pack(fill=tk.BOTH, expand=True)

        self.add(frame, text=company)
        self._open_tabs[company] = frame
        self.select(frame)


class FinanceAnalystApp:
    """Main window controller for the thatkowfinanace_analyst application."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("thatkowfinanace_analyst")
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
