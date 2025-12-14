from __future__ import annotations

import math
from pathlib import Path
from typing import Dict

import pandas as pd

from analyst.data import Company
from analyst.stats import (
    FinancialBoxplots,
    FinancialViolins,
    financials_boxplots,
    financials_violin_comparison,
)
from .stackedvisuals import render_stacked_annual_report, render_stacked_comparison
from . import yahoo

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


def plot_stacked_financials(
    company: Company,
    *,
    out_path: str | Path | None = None,
    include_intangibles: bool = True,
) -> Path:
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
        try:
            # Validate that column name is a string
            if not isinstance(col, str):
                raise ValueError(f"Column name '{col}' is not a string.")

            # Confirm column exists
            if col not in df_all.columns:
                raise KeyError(f"Column '{col}' not found in DataFrame.")

            series = df_all[col]

            # Confirm the selection is a Series, not a DataFrame (duplicate column names)
            if not hasattr(series, "dtype"):
                raise TypeError(
                    f"Column '{col}' returned a non-Series object "
                    f"(possible duplicate column names or invalid selection)."
                )

            # Perform string replacement safely
            df_all[col] = series.astype(str).str.replace(",", "", regex=False)

        except Exception as e:
            raise RuntimeError(
                f"Error while processing column '{col}'. "
                f"Column type = {type(series)}. "
                f"Columns in DataFrame = {list(df_all.columns)}. "
                f"Details: {e}"
            )



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

    notes_lower = df_all["NOTE"].astype(str).str.lower()
    mask = notes_lower != "excluded"

    df = df_all[mask].copy()
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

    # Append today's stock price to the latest financial period tooltip if available
    latest_price = None
    try:
        prices = yahoo.get_stock_prices(ticker, years=1)
        if not prices.empty:
            latest_val = pd.to_numeric(prices["Price"].iloc[-1], errors="coerce")
            if isinstance(latest_val, pd.Series):
                latest_val = latest_val.iloc[0]
            if pd.notna(latest_val):
                latest_price = float(latest_val)
    except Exception as exc:  # pragma: no cover - network dependent
        print(f"⚠️ Failed to fetch today's price for {ticker}: {exc}")

    if latest_price is not None and year_cols:
        factor_tooltip.setdefault(year_cols[-1], []).append(f"Today: {latest_price:.3f}")

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
        include_intangibles=include_intangibles,
        latest_price=latest_price,
    )

    return out_path


# Backwards compatibility
plot_stacked_visuals = plot_stacked_financials


def compare_stacked_financials(
    companies: list[Company],
    *,
    out_path: str | Path | None = None,
    include_intangibles: bool = True,
) -> Path:
    if not companies:
        raise ValueError("No companies provided for comparison.")

    required_common_rows = [
        {"TYPE": "Meta", "CATEGORY": "PDF source", "SUBCATEGORY": "", "ITEM": "", "NOTE": "excluded"},
        {"TYPE": "Financial", "CATEGORY": "Financial Multiplier", "SUBCATEGORY": "", "ITEM": "", "NOTE": "excluded"},
        {"TYPE": "Income", "CATEGORY": "Income Multiplier", "SUBCATEGORY": "", "ITEM": "", "NOTE": "excluded"},
        {"TYPE": "Shares", "CATEGORY": "Shares Multiplier", "SUBCATEGORY": "", "ITEM": "", "NOTE": "excluded"},
        {"TYPE": "Meta", "CATEGORY": "Stock Multiplier", "SUBCATEGORY": "", "ITEM": "", "NOTE": "excluded"},
        {"TYPE": "Meta", "CATEGORY": "ReleaseDate", "SUBCATEGORY": "", "ITEM": "", "NOTE": "excluded"},
        {"TYPE": "Stock", "CATEGORY": "Prices", "SUBCATEGORY": "-30", "ITEM": "", "NOTE": "excluded"},
        {"TYPE": "Stock", "CATEGORY": "Prices", "SUBCATEGORY": "-7", "ITEM": "", "NOTE": "excluded"},
        {"TYPE": "Stock", "CATEGORY": "Prices", "SUBCATEGORY": "-1", "ITEM": "", "NOTE": "excluded"},
        {"TYPE": "Stock", "CATEGORY": "Prices", "SUBCATEGORY": "0", "ITEM": "", "NOTE": "excluded"},
        {"TYPE": "Stock", "CATEGORY": "Prices", "SUBCATEGORY": "1", "ITEM": "", "NOTE": "excluded"},
        {"TYPE": "Stock", "CATEGORY": "Prices", "SUBCATEGORY": "7", "ITEM": "", "NOTE": "excluded"},
        {"TYPE": "Stock", "CATEGORY": "Prices", "SUBCATEGORY": "30", "ITEM": "", "NOTE": "excluded"},
    ]

    def _normalize_cell(val: object) -> str:
        if pd.isna(val):
            return ""

        if isinstance(val, (int, float)):
            if isinstance(val, float) and not math.isfinite(val):
                return ""
            return str(int(val)) if float(val).is_integer() else str(val)

        if isinstance(val, str):
            stripped = val.strip()
            if not stripped:
                return ""
            try:
                numeric = float(stripped)
                if math.isfinite(numeric) and numeric.is_integer():
                    return str(int(numeric))
            except ValueError:
                pass
            return stripped.lower()

        return str(val).strip().lower()

    def _signature(row: dict) -> tuple[str, ...]:
        return tuple(_normalize_cell(row.get(col, "")) for col in COMBINED_BASE_COLUMNS)

    required_signatures = {_signature(r) for r in required_common_rows}

    prepared_frames: list[pd.DataFrame] = []
    available_shifts_per_ticker: dict[str, set[str]] = {}
    factor_lookup: Dict[str, Dict[str, Dict[str, float]]] = {}
    pdf_sources: Dict[str, str] = {}
    release_lines: Dict[str, list[str]] = {}
    price_tooltips: Dict[str, Dict[str, str]] = {}
    allowed_shifts = ["-30", "-10", "-1", "0", "1", "10", "30"]

    # Build a combined frame via outer concatenation first
    for company in companies:
        combined_df = company.combined
        ticker = company.ticker

        if combined_df.empty:
            raise ValueError(f"Combined dataframe is empty for {ticker}; generate data first.")

        df_all = combined_df.copy().fillna("")
        df_all["Ticker"] = ticker
        prepared_frames.append(df_all)

        sigs = {_signature(row) for row in df_all.to_dict(orient="records")}
        missing = required_signatures - sigs
        if missing:
            raise ValueError(
                f"Company {ticker} is missing required rows: {missing}."
            )

    combined_all = pd.concat(prepared_frames, ignore_index=True, join="outer", sort=False).fillna("")

    excluded_cols = set(COMBINED_BASE_COLUMNS + ["Ticker"])
    num_cols = [c for c in combined_all.columns if c not in excluded_cols]

    for col in num_cols:
        try:
            series = combined_all[col]
            if not hasattr(series, "dtype"):
                raise TypeError(
                    f"Column '{col}' returned a non-Series object "
                    f"(possible duplicate column names or invalid selection)."
                )

            combined_all[col] = series.astype(str).str.replace(",", "", regex=False)
        except Exception as e:
            raise RuntimeError(
                f"Error while processing column '{col}'. "
                f"Column type = {type(series)}. "
                f"Columns in DataFrame = {list(combined_all.columns)}. "
                f"Details: {e}"
            )

    processed_frames: list[pd.DataFrame] = []
    share_counts: Dict[str, Dict[str, float]] = {}

    for company in companies:
        ticker = company.ticker
        df_all = combined_all[combined_all["Ticker"] == ticker].copy()

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

        notes_lower = df_all["NOTE"].astype(str).str.lower()
        mask = notes_lower != "excluded"

        df = df_all[mask].copy()
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

        year_cols = [c for c in num_cols if df[c].notna().any()]

        share_counts[ticker] = {}
        share_rows = df[df["ITEM"].str.lower().str.contains("number of shares", na=False)]
        if not share_rows.empty:
            row = share_rows.iloc[0]
            for year in year_cols:
                raw_val = row.get(year)
                if pd.notna(raw_val):
                    numeric_val = float(raw_val)
                    share_counts[ticker][year] = round(numeric_val, 2)
                else:
                    raise ValueError(f"Number of shares for year '{year}' is missing or NaN for {ticker}.")
        else:
            share_counts[ticker] = {year: 1.0 for year in year_cols}

        df_plot = df[~df["ITEM"].str.lower().str.contains("number of shares", na=False)].copy()

        for col in year_cols:
            shares = share_counts[ticker].get(col, 1.0)
            if shares == 0 or pd.isna(shares):
                raise ValueError(f"Share count for {ticker} and year '{col}' is invalid: {shares}")
            df_plot[col] = df_plot[col] / shares

        df_plot["Ticker"] = ticker

        for year in year_cols:
            factor_lookup.setdefault("", {}).setdefault(ticker, {})[year] = 1.0
            release_date = release_map.get(year, "")
            release_entry = f"{ticker} Release Date: {release_date + ' days' if release_date else 'NA'}"
            release_lines.setdefault(year, []).append(release_entry)

            for _, prow in price_rows.iterrows():
                label_raw = _normalize_cell(prow.get("SUBCATEGORY", "")) or "Price"
                if label_raw not in allowed_shifts:
                    continue
                price_val = pd.to_numeric(prow.get(year, ""), errors="coerce")
                price_entry = f"{ticker} {label_raw}: "
                if pd.isna(price_val) or price_val <= 0:
                    factor_lookup.setdefault(label_raw, {}).setdefault(ticker, {})[year] = float("nan")
                    price_entry += "NaN"
                else:
                    factor_lookup.setdefault(label_raw, {}).setdefault(ticker, {})[year] = 1.0 / float(price_val)
                    price_entry += f"{price_val:.3f}"
                price_tooltips.setdefault(year, {})[f"{ticker}:{label_raw}"] = price_entry
                available_shifts_per_ticker.setdefault(ticker, set()).add(label_raw)

            pdf_name = pdf_map.get(year)
            if pdf_name:
                pdf_sources[f"{year}:{ticker}"] = pdf_name

        processed_frames.append(df_plot)

    # Determine common shifts across all tickers
    if not available_shifts_per_ticker:
        raise ValueError("No release date shift data available for comparison.")

    common_shifts = set(allowed_shifts)
    for shifts in available_shifts_per_ticker.values():
        common_shifts &= shifts

    if not common_shifts:
        raise ValueError("No common Release Date Shift values across all companies.")

    factor_lookup = {k: v for k, v in factor_lookup.items() if k == "" or k in common_shifts}

    factor_tooltip: Dict[str, list[str]] = {}
    for year, releases in release_lines.items():
        entries = releases.copy()
        price_entries = price_tooltips.get(year, {})
        for ticker, shifts in available_shifts_per_ticker.items():
            for shift in allowed_shifts:
                if shift not in common_shifts or shift not in shifts:
                    continue
                tooltip_line = price_entries.get(f"{ticker}:{shift}")
                if tooltip_line:
                    entries.append(tooltip_line)
        factor_tooltip[year] = entries

    combined_df_plot = pd.concat(processed_frames, ignore_index=True, sort=False)

    if out_path:
        output_path = Path(out_path)
    else:
        base = companies[0].default_visuals_path()
        output_path = base.with_name("stacked_comparison_report.html")

    visuals_dir = output_path.expanduser().resolve().parent
    visuals_dir.mkdir(parents=True, exist_ok=True)
    output_path = visuals_dir / output_path.name

    render_stacked_comparison(
        combined_df_plot,
        title="Financial/Income Comparison",
        factor_lookup=factor_lookup,
        factor_label="Release Date Shift",
        factor_tooltip=factor_tooltip,
        factor_tooltip_label="Prices",
        pdf_sources=pdf_sources,
        out_path=output_path,
        include_intangibles=include_intangibles,
    )

    return output_path
