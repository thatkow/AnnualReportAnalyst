from __future__ import annotations

from pathlib import Path
from typing import Dict

import pandas as pd

from analyst.data import Company
from analyst.stats import FinancialBoxplots, financials_boxplots
from .stackedvisuals import render_stacked_annual_report

# Base, non-date columns present in the combined dataset
COMBINED_BASE_COLUMNS = [
    "TYPE",
    "CATEGORY",
    "SUBCATEGORY",
    "ITEM",
    "NOTE",
    "Key4Coloring",
]


def _extract_multiplier(row_df: pd.DataFrame, num_cols: list[str]) -> Dict[str, float]:
    if row_df.empty:
        return {}

    row = row_df.iloc[0]

    def to_number(val, col):
        if isinstance(val, pd.Series):
            if val.notna().any():
                val = val[val.notna()].iloc[0]
            else:
                raise ValueError(f"Multiplier for column '{col}' is empty or NaN.")

        s = str(val).strip()
        if s == "":
            raise ValueError(f"Multiplier for column '{col}' is blank.")

        try:
            return float(s)
        except Exception as exc:  # pragma: no cover - defensive
            raise ValueError(
                f"Multiplier for column '{col}' must be a number, got: '{val}'"
            ) from exc

    return {col: to_number(row[col], col) for col in num_cols}


def _release_date_map(df_all: pd.DataFrame, num_cols: list[str], company: Company) -> Dict[str, str]:
    release_rows = df_all[
        (df_all["TYPE"].str.lower() == "meta")
        & (df_all["CATEGORY"].str.lower() == "releasedate")
    ]
    release_map: Dict[str, str] = {}
    if not release_rows.empty:
        row = release_rows.iloc[0]
        for col in num_cols:
            val = str(row.get(col, "")).strip()
            if val:
                release_map[col] = val

    # Fallback to legacy ReleaseDates.csv if the combined table lacks release dates
    if not release_map:
        release_csv = company.release_dates_csv
        if release_csv.exists():
            release_map = (
                pd.read_csv(release_csv)
                .set_index("Date")["ReleaseDate"]
                .fillna("")
                .to_dict()
            )
        else:
            raise ValueError(
                "Release dates are missing from the combined table and ReleaseDates.csv."
            )
    return release_map


def _pdf_source_map(df_all: pd.DataFrame, num_cols: list[str]) -> Dict[str, str]:
    """Map each financial date column to the base PDF source name (sans extension)."""

    pdf_rows = df_all[
        (df_all["TYPE"].str.lower() == "meta")
        & (df_all["CATEGORY"].str.lower() == "pdf source")
    ]

    if pdf_rows.empty:
        return {}

    pdf_map: Dict[str, str] = {}
    row = pdf_rows.iloc[0]
    for col in num_cols:
        val = str(row.get(col, "")).strip()
        if val:
            pdf_map[col] = Path(val).stem

    return pdf_map


def plot_stacked_financials(company: Company, *, out_path: str | Path | None = None) -> Path:
    """Plot stacked visuals for a company's combined dataset."""

    combined_df = company.combined
    ticker = company.ticker
    out_path = Path(out_path) if out_path else company.default_visuals_path()

    if combined_df.empty:
        raise ValueError("Combined dataframe is empty; generate data first.")

    df_all = combined_df.copy().fillna("")
    excluded_cols = set(COMBINED_BASE_COLUMNS + ["Ticker"])
    num_cols = [c for c in df_all.columns if c not in excluded_cols]
    for col in num_cols:
        df_all[col] = df_all[col].astype(str).str.replace(",", "", regex=False)

    share_mult = _extract_multiplier(
        df_all[df_all["CATEGORY"].str.lower() == "shares multiplier"], num_cols
    )
    stock_mult = _extract_multiplier(
        df_all[df_all["CATEGORY"].str.lower() == "stock multiplier"], num_cols
    )
    fin_mult = _extract_multiplier(
        df_all[df_all["CATEGORY"].str.lower() == "financial multiplier"], num_cols
    )
    inc_mult = _extract_multiplier(
        df_all[df_all["CATEGORY"].str.lower() == "income multiplier"], num_cols
    )

    df_all["Ticker"] = ticker

    df = df_all[df_all["NOTE"].str.lower() != "excluded"].copy()
    for col in num_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce")
    neg_idx = df["NOTE"].str.lower() == "negated"
    df.loc[neg_idx, num_cols] = df.loc[neg_idx, num_cols].map(
        lambda x: -1.0 * x if pd.notna(x) else x
    )

    def _apply_row_multiplier(mask: pd.Series, factors: Dict[str, float]) -> None:
        for year in num_cols:
            factor = factors.get(year, 1.0)
            if factor == 1.0:
                continue
            df.loc[mask, year] = df.loc[mask, year] * factor

    _apply_row_multiplier(df["TYPE"].str.lower() == "financial", fin_mult)
    _apply_row_multiplier(df["TYPE"].str.lower() == "income", inc_mult)
    _apply_row_multiplier(df["TYPE"].str.lower() == "shares", share_mult)
    _apply_row_multiplier(df["TYPE"].str.lower() == "shares", stock_mult)

    price_rows = df_all[
        (df_all["TYPE"].str.lower() == "stock")
        & (df_all["CATEGORY"].str.lower() == "prices")
    ].copy()
    for year in num_cols:
        price_rows[year] = pd.to_numeric(price_rows[year], errors="coerce")
 

    release_map = _release_date_map(df_all, num_cols, company)
    pdf_map = _pdf_source_map(df_all, num_cols)

    year_cols = [c for c in df.columns if c not in excluded_cols]

    factor_lookup: Dict[str, Dict[str, float]] = {"": {y: 1.0 for y in year_cols}}
    for _, prow in price_rows.iterrows():
        label = str(prow.get("SUBCATEGORY", "")).strip() or "Price"
        factor_lookup[label] = {}
        for year in year_cols:
            price_val = pd.to_numeric(prow.get(year, ""), errors="coerce")
            if pd.isna(price_val) or price_val <= 0:
                factor_lookup[label][year] = float("nan")
            else:
                factor_lookup[label][year] = 1.0 / float(price_val)

    factor_tooltip: Dict[str, list[str]] = {}
    for financial_date in year_cols:
        release_date = release_map.get(financial_date, "")
        release_text = f"{release_date} days" if release_date else ""
        entries = [f"Release Date: {release_text or 'NA'}"]
        for label, lookup in factor_lookup.items():
            if label == "":
                continue
            inv = lookup.get(financial_date, float("nan"))
            if pd.isna(inv) or inv == 0:
                entries.append(f"{label}: NaN")
            else:
                entries.append(f"{label}: {1.0 / inv:.3f}")
        factor_tooltip[financial_date] = entries

    df["Ticker"] = ticker

    share_counts: Dict[str, Dict[str, float]] = {ticker: {}}
    share_rows = df[df["ITEM"].str.lower().str.contains("number of shares", na=False)]
    if not share_rows.empty:
        row = share_rows.iloc[0]
        for year in year_cols:
            raw_val = row.get(year)
            if pd.notna(raw_val):
                numeric_val = float(raw_val)
                share_counts[ticker][year] = round(numeric_val, 2)
            else:
                raise ValueError(f"Number of shares for year '{year}' is missing or NaN.")
    else:
        share_counts[ticker] = {year: 1.0 for year in year_cols}

    df_plot = df[~df["ITEM"].str.lower().str.contains("number of shares", na=False)].copy()

    visuals_dir = Path(out_path).expanduser().resolve().parent
    visuals_dir.mkdir(parents=True, exist_ok=True)
    out_path = visuals_dir / Path(out_path).name

    render_stacked_annual_report(
        df_plot,
        title=f"Financial/Income for {ticker}",
        factor_lookup=factor_lookup,
        factor_label="Release Date Shift",
        factor_tooltip=factor_tooltip,
        factor_tooltip_label="Prices",
        share_counts=share_counts,
        pdf_sources=pdf_map,
        out_path=out_path,
    )

    return out_path


# Backwards compatibility
plot_stacked_visuals = plot_stacked_financials
