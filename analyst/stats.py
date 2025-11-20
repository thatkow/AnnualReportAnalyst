from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable, Sequence

import matplotlib.pyplot as plt
import pandas as pd

from analyst.data import Company, import_companies

# Base, non-date columns present in the combined dataset
COMBINED_BASE_COLUMNS = [
    "TYPE",
    "CATEGORY",
    "SUBCATEGORY",
    "ITEM",
    "NOTE",
    "Key4Coloring",
]


@dataclass
class BoxPlotFigures:
    """Container for financial and income boxplot figures."""

    fig_fin: plt.Figure
    fig_inc: plt.Figure


def _melt_financial_values(companies: Iterable[Company], type_filter: str) -> pd.DataFrame:
    """Convert wide combined data into long-form values for a specific TYPE."""

    records: list[dict[str, object]] = []
    for company in companies:
        df = company.combined.copy().fillna("")
        excluded_cols = set(COMBINED_BASE_COLUMNS + ["Ticker"])
        date_cols = [c for c in df.columns if c not in excluded_cols]

        # Normalize numeric cells
        for col in date_cols:
            df[col] = pd.to_numeric(df[col].astype(str).str.replace(",", ""), errors="coerce")

        typed_df = df[df["TYPE"].str.lower() == type_filter.lower()]
        for date_col in date_cols:
            for val in typed_df[date_col].dropna():
                records.append({
                    "Ticker": company.ticker,
                    "Date": date_col,
                    "Value": float(val),
                })

    return pd.DataFrame.from_records(records, columns=["Ticker", "Date", "Value"])


def _build_boxplot(df_long: pd.DataFrame, *, title: str) -> plt.Figure:
    """Create a matplotlib boxplot from a long-form dataframe."""

    fig, ax = plt.subplots(figsize=(10, 6))

    if df_long.empty:
        ax.set_title(f"{title} (no data)")
        ax.set_xlabel("Date")
        ax.set_ylabel("Value")
        fig.tight_layout()
        return fig

    grouped = []
    labels: list[str] = []
    for date in sorted(df_long["Date"].unique()):
        values = df_long[df_long["Date"] == date]["Value"].dropna().tolist()
        if not values:
            continue
        labels.append(date)
        grouped.append(values)

    if grouped:
        ax.boxplot(grouped, labels=labels, showmeans=True)

    ax.set_title(title)
    ax.set_xlabel("Date")
    ax.set_ylabel("Value")
    ax.grid(True, linestyle="--", alpha=0.5)
    fig.tight_layout()
    return fig


def financials_boxplots(companies: Sequence[Company]) -> BoxPlotFigures:
    """Build boxplots for financial and income data across companies."""

    if not companies:
        raise ValueError("No companies provided for boxplot generation.")

    df_fin = _melt_financial_values(companies, "financial")
    df_inc = _melt_financial_values(companies, "income")

    fig_fin = _build_boxplot(df_fin, title="Financials by Date")
    fig_inc = _build_boxplot(df_inc, title="Income by Date")

    return BoxPlotFigures(fig_fin=fig_fin, fig_inc=fig_inc)


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Generate financial and income boxplots.")
    parser.add_argument("tickers", nargs="+", help="Ticker symbols to load from the companies directory.")
    parser.add_argument(
        "--companies-dir",
        default="companies",
        help="Base directory containing the per-company data folders (default: companies)",
    )

    args = parser.parse_args()
    company_list = import_companies(args.tickers, companies_dir=args.companies_dir)
    figures = financials_boxplots(company_list)
    figures.fig_fin.show()
    figures.fig_inc.show()
