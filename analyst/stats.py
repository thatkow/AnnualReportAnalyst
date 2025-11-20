from __future__ import annotations

import json
from pathlib import Path

import pandas as pd

from analyst.data import Company
from analyst.plots import COMBINED_BASE_COLUMNS, _extract_multiplier, _release_date_map


def _prepare_company_financials(company: Company):
    ticker = company.ticker
    df_all = company.combined.copy().fillna("")

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

    def _apply_row_multiplier(mask: pd.Series, factors: dict[str, float]) -> None:
        for year in num_cols:
            factor = factors.get(year, 1.0)
            if factor == 1.0:
                continue
            df.loc[mask, year] = df.loc[mask, year] * factor

    _apply_row_multiplier(df["TYPE"].str.lower() == "financial", fin_mult)
    _apply_row_multiplier(df["TYPE"].str.lower() == "income", inc_mult)
    _apply_row_multiplier(df["TYPE"].str.lower() == "shares", share_mult)

    price_rows = df_all[
        (df_all["TYPE"].str.lower() == "stock")
        & (df_all["CATEGORY"].str.lower() == "prices")
    ].copy()
    for year in num_cols:
        factor = stock_mult.get(year, 1.0)
        if factor == 1.0:
            price_rows[year] = pd.to_numeric(price_rows[year], errors="coerce")
        else:
            price_rows[year] = pd.to_numeric(price_rows[year], errors="coerce") * factor

    release_map = _release_date_map(df_all, num_cols, company)
    year_cols = [c for c in df.columns if c not in excluded_cols]

    share_counts: dict[str, dict[str, float]] = {ticker: {}}
    share_rows = df[df["ITEM"].str.lower().str.contains("number of shares", na=False)]
    if not share_rows.empty:
        row = share_rows.iloc[0]
        for year in year_cols:
            raw_val = row.get(year)
            if pd.notna(raw_val):
                share_counts[ticker][year] = float(raw_val)
            else:
                raise ValueError(
                    f"Number of shares for year '{year}' is missing or NaN."
                )
    else:
        share_counts[ticker] = {year: 1.0 for year in year_cols}

    df_plot = df[~df["ITEM"].str.lower().str.contains("number of shares", na=False)].copy()

    return {
        "df": df_plot,
        "price_rows": price_rows,
        "release_map": release_map,
        "share_counts": share_counts,
        "year_cols": year_cols,
    }


def render_release_date_boxplots(
    company: Company, *, out_path: str | Path | None = None
) -> Path:
    """Render box plots grouped by release date differential.

    Values are normalised per share, and corresponding release-date share
    prices are shown alongside the financial series to mirror the
    stacked-visuals context.
    """

    prep = _prepare_company_financials(company)
    df_plot: pd.DataFrame = prep["df"]
    price_rows: pd.DataFrame = prep["price_rows"]
    release_map: dict[str, str] = prep["release_map"]
    share_counts: dict[str, dict[str, float]] = prep["share_counts"]
    year_cols: list[str] = prep["year_cols"]

    ticker = company.ticker
    out_path = (
        Path(out_path)
        if out_path
        else company.visuals_dir / f"ReleaseDateBoxes_{ticker}.html"
    )

    records: list[dict[str, object]] = []

    for _, row in df_plot.iterrows():
        series_label = f"{row.get('TYPE', '')}: {row.get('ITEM', '')}".strip()
        for year in year_cols:
            value = row.get(year)
            if pd.isna(value):
                continue
            shares = share_counts.get(ticker, {}).get(year)
            if shares in (None, 0):
                continue
            release_key = str(release_map.get(year, "Unknown")) or "Unknown"
            records.append(
                {
                    "release": release_key,
                    "series": series_label,
                    "value": float(value) / float(shares),
                }
            )

    for _, row in price_rows.iterrows():
        price_label = str(row.get("SUBCATEGORY", "")).strip() or "Share Price"
        for year in year_cols:
            price_val = pd.to_numeric(row.get(year, ""), errors="coerce")
            if pd.isna(price_val):
                continue
            release_key = str(release_map.get(year, "Unknown")) or "Unknown"
            records.append(
                {
                    "release": release_key,
                    "series": f"Price ({price_label})",
                    "value": float(price_val),
                }
            )

    visuals_dir = out_path.expanduser().resolve().parent
    visuals_dir.mkdir(parents=True, exist_ok=True)
    out_path = visuals_dir / out_path.name

    html = f"""<!DOCTYPE html>
<html lang=\"en\">
<head>
  <meta charset=\"utf-8\" />
  <title>Release Date Box Plots - {ticker}</title>
  <script src=\"https://cdn.plot.ly/plotly-2.31.1.min.js\"></script>
</head>
<body>
  <h2>Release Date Differential Box Plots - {ticker}</h2>
  <div id=\"plot\"></div>
  <script>
    const rawData = {json.dumps(records)};
    const grouped = new Map();

    rawData.forEach(rec => {{
      const key = `${{rec.release}}|${{rec.series}}`;
      if (!grouped.has(key)) grouped.set(key, []);
      grouped.get(key).push(rec.value);
    }});

    const traces = [];
    grouped.forEach((values, key) => {{
      const [release, series] = key.split('|');
      traces.push({{
        type: 'box',
        name: series,
        x: Array(values.length).fill(release),
        y: values,
        boxpoints: 'outliers',
        jitter: 0.4,
        pointpos: -1.8,
        marker: {{ size: 6 }}
      }});
    }});

    const layout = {{
      boxmode: 'group',
      xaxis: {{ title: 'Release Date Differential' }},
      yaxis: {{ title: 'Value (per share)' }},
      legend: {{ orientation: 'h' }},
      margin: {{ t: 40 }}
    }};

    Plotly.newPlot('plot', traces, layout);
  </script>
</body>
</html>
"""

    out_path.write_text(html, encoding="utf-8")
    return out_path

