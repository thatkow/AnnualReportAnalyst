from __future__ import annotations

from pathlib import Path
import json
import os
import tempfile
import webbrowser
from typing import Dict, Iterable, List

import pandas as pd

from analyst.data import Company
from analyst.plots import (
    COMBINED_BASE_COLUMNS,
    _extract_multiplier,
    _pdf_source_map,
    _release_date_map,
)

def _normalize_shift_label(label: str) -> str:
    try:
        numeric = float(label)
    except (TypeError, ValueError):
        return str(label).strip()
    if numeric.is_integer():
        return str(int(numeric))
    return str(numeric)


REQUIRED_COMMON_ROWS = [
    ("Meta", "PDF source", "", "", "excluded"),
    ("Financial", "Financial Multiplier", "", "", "excluded"),
    ("Income", "Income Multiplier", "", "", "excluded"),
    ("Shares", "Shares Multiplier", "", "", "excluded"),
    ("Meta", "Stock Multiplier", "", "", "excluded"),
    ("Meta", "ReleaseDate", "", "", "excluded"),
    ("Stock", "Prices", "-30", "", "excluded"),
    ("Stock", "Prices", "-7", "", "excluded"),
    ("Stock", "Prices", "-1", "", "excluded"),
    ("Stock", "Prices", "0", "", "excluded"),
    ("Stock", "Prices", "1", "", "excluded"),
    ("Stock", "Prices", "7", "", "excluded"),
    ("Stock", "Prices", "30", "", "excluded"),
]


def _validate_required_rows(df: pd.DataFrame, ticker: str) -> pd.DataFrame:
    missing: List[tuple[str, str, str, str, str]] = []
    for row in REQUIRED_COMMON_ROWS:
        type_val, cat_val, sub_val, item_val, note_val = row
        mask = (
            df["TYPE"].astype(str).str.strip().str.lower() == type_val.lower()
        ) & (df["CATEGORY"].astype(str).str.strip().str.lower() == cat_val.lower())
        mask &= df["SUBCATEGORY"].astype(str).str.strip().str.lower() == sub_val.lower()
        mask &= df["ITEM"].astype(str).str.strip().str.lower() == item_val.lower()
        mask &= df["NOTE"].astype(str).str.strip().str.lower() == note_val.lower()
        if not mask.any():
            missing.append(row)
    if not missing:
        return df

    # Fill missing common rows with placeholder entries so comparisons still render
    # even when a source omits the excluded rows. These placeholders remain filtered
    # out of visuals because they retain the "excluded" note.
    placeholders: list[dict[str, object]] = []
    for type_val, cat_val, sub_val, item_val, note_val in missing:
        entry: dict[str, object] = {
            "TYPE": type_val,
            "CATEGORY": cat_val,
            "SUBCATEGORY": sub_val,
            "ITEM": item_val,
            "NOTE": note_val,
        }
        placeholders.append(entry)

    for entry in placeholders:
        for col in df.columns:
            entry.setdefault(col, "")

    df_placeholder = pd.DataFrame(placeholders, columns=df.columns)
    df_placeholder["Ticker"] = ticker
    combined = pd.concat([df, df_placeholder], ignore_index=True)
    return combined


def _prepare_company_dataframe(
    company: Company,
) -> tuple[pd.DataFrame, Dict[str, Dict[str, float]], Dict[str, Dict[str, float]], Dict[str, str]]:
    combined_df = company.combined
    ticker = company.ticker

    if combined_df.empty:
        raise ValueError(f"Combined dataframe is empty for {ticker}; generate data first.")

    df_all = combined_df.copy().fillna("")

    price_mask = (
        df_all["TYPE"].str.lower() == "stock"
    ) & (df_all["CATEGORY"].str.lower() == "prices")
    df_all.loc[price_mask, "SUBCATEGORY"] = df_all.loc[price_mask, "SUBCATEGORY"].map(
        _normalize_shift_label
    )
    df_all = _validate_required_rows(df_all, ticker)

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

    mask = df_all["NOTE"].astype(str).str.lower() != "excluded"
    df = df_all[mask].copy()
    for col in num_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce")
    neg_idx = df["NOTE"].str.lower() == "negated"
    df.loc[neg_idx, num_cols] = df.loc[neg_idx, num_cols].apply(
        lambda col: col.map(lambda x: -1.0 * x if pd.notna(x) else x)
    )

    def _apply_row_multiplier(mask_series: pd.Series, factors: Dict[str, float]) -> None:
        for year in num_cols:
            factor = factors.get(year, 1.0)
            if factor == 1.0:
                continue
            df.loc[mask_series, year] = df.loc[mask_series, year] * factor

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
    _ = _pdf_source_map(df_all, num_cols)  # retained for parity with plot_stacked_financials

    share_counts: Dict[str, float] = {}
    share_rows = df[df["ITEM"].str.lower().str.contains("number of shares", na=False)]
    if not share_rows.empty:
        row = share_rows.iloc[0]
        for year in num_cols:
            raw_val = row.get(year)
            if pd.notna(raw_val):
                share_counts[year] = float(raw_val)
            else:
                share_counts[year] = 1.0
    else:
        share_counts = {year: 1.0 for year in num_cols}

    df_plot = df[~df["ITEM"].str.lower().str.contains("number of shares", na=False)].copy()

    # Divide all numeric values by share counts (per-share values only)
    for year in num_cols:
        divisor = share_counts.get(year, 1.0) or 1.0
        df_plot[year] = df_plot[year] / divisor

    # Build factor lookup using price rows (release date shifts)
    factor_lookup: Dict[str, Dict[str, float]] = {}
    for _, prow in price_rows.iterrows():
        raw_label = str(prow.get("SUBCATEGORY", "")).strip() or "Price"
        label = _normalize_shift_label(raw_label)
        factor_lookup.setdefault(label, {})
        for year in num_cols:
            price_val = pd.to_numeric(prow.get(year, ""), errors="coerce")
            if pd.isna(price_val) or price_val <= 0:
                factor_lookup[label][year] = float("nan")
            else:
                factor_lookup[label][year] = 1.0 / float(price_val)

    df_plot["Ticker"] = ticker

    return df_plot, factor_lookup, {year: release_map.get(year, "") for year in num_cols}, share_counts


def _merge_factor_lookups(factors: Iterable[tuple[str, Dict[str, Dict[str, float]]]]) -> Dict[str, Dict[str, Dict[str, float]]]:
    merged: Dict[str, Dict[str, Dict[str, float]]] = {}
    for ticker, lookup in factors:
        for label, year_map in lookup.items():
            merged.setdefault(label, {})
            merged[label].setdefault(ticker, {})
            for year, factor in year_map.items():
                merged[label][ticker][year] = factor
    return merged


def compare_stacked_financials(
    companies: list[Company],
    *,
    out_path: str | Path | None = None,
    include_intangibles: bool = True,
) -> Path:
    if not companies:
        raise ValueError("At least one company is required for comparison.")

    prepared_frames = []
    factor_entries = []
    release_entries: Dict[str, Dict[str, str]] = {}

    for company in companies:
        df_plot, factors, release_map, shares = _prepare_company_dataframe(company)
        prepared_frames.append(df_plot)
        factor_entries.append((company.ticker, factors))
        release_entries[company.ticker] = release_map

    combined_df = pd.concat(prepared_frames, join="outer", ignore_index=True, sort=False)

    excluded_cols = set(COMBINED_BASE_COLUMNS + ["Ticker"])
    year_cols = [c for c in combined_df.columns if c not in excluded_cols]

    combined_df[COMBINED_BASE_COLUMNS + ["Ticker"]] = combined_df[COMBINED_BASE_COLUMNS + ["Ticker"]].fillna("")
    for col in year_cols:
        combined_df[col] = pd.to_numeric(combined_df[col], errors="coerce").fillna(0)

    tickers = sorted({c.ticker for c in companies})
    types = sorted(combined_df["TYPE"].dropna().unique())

    factor_lookup = _merge_factor_lookups(factor_entries)

    if out_path is None:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".html") as tmp:
            out_path = Path(tmp.name)
    else:
        out_path = Path(out_path)
        visuals_dir = Path(out_path).expanduser().resolve().parent
        visuals_dir.mkdir(parents=True, exist_ok=True)
        out_path = visuals_dir / out_path.name

    def _year_sort_key(val: str) -> tuple[int, float | str]:
        ts = pd.to_datetime(val, errors="coerce", dayfirst=True)
        if pd.notna(ts):
            return (0, ts.timestamp())
        numeric_str = str(val)
        if numeric_str.replace(".", "", 1).lstrip("-+").isdigit():
            return (1, float(numeric_str))
        return (2, str(val))

    sorted_years = sorted(year_cols, key=_year_sort_key)

    year_labels = []
    for year in sorted_years:
        for ticker in tickers:
            release_val = release_entries.get(ticker, {}).get(year, "")
            label_text = f"{year} ({release_val}) - {ticker}" if release_val else f"{year} - {ticker}"
            year_labels.append({"label": label_text, "ticker": ticker, "year": year})

    records = []
    for _, r in combined_df.iterrows():
        rec = {
            "Ticker": r.get("Ticker"),
            "TYPE": r.get("TYPE"),
            "CATEGORY": r.get("CATEGORY"),
            "SUBCATEGORY": r.get("SUBCATEGORY"),
            "ITEM": r.get("ITEM"),
            "NOTE": r.get("NOTE"),
            "Key4Coloring": r.get("Key4Coloring"),
        }
        for y in year_cols:
            rec[y] = float(r.get(y, 0) or 0)
        records.append(rec)

    type_offsets = {t: round((i - (len(types) - 1) / 2) * 0.6, 2) for i, t in enumerate(types)}
    ticker_offsets = {t: (i - ((len(tickers) - 1) / 2)) * 0.25 for i, t in enumerate(tickers)}

    years_json = json.dumps(sorted_years)
    tickers_json = json.dumps(tickers)
    types_json = json.dumps(types)
    records_json = json.dumps(records)
    factor_json = json.dumps(factor_lookup)
    year_labels_json = json.dumps(year_labels)
    release_json = json.dumps(release_entries)
    type_offsets_json = json.dumps(type_offsets)
    ticker_offsets_json = json.dumps(ticker_offsets)

    html = f"""<!DOCTYPE html>
<html lang=\"en\">
<head>
<meta charset=\"utf-8\" />
<title>Stacked Financial Comparison</title>
<script src=\"https://cdn.plot.ly/plotly-2.31.1.min.js\"></script>
<style>
html, body {{ height: 100%; }}
body {{ font-family: sans-serif; margin: 40px; box-sizing: border-box; }}
#controls {{ margin: 15px 0; display: flex; flex-direction: column; gap: 8px; }}
.year-toggle {{ display: inline-flex; align-items: center; gap: 4px; margin-right: 12px; }}
#plotBars {{ width: 100%; min-height: 400px; }}
</style>
</head>
<body>
<h2>Stacked Financial Comparison</h2>
<div id=\"controls\">
  <div>
    <label><b>Release Date Shift:</b></label>
    <select id=\"factorSelector\"></select>
    <label style=\"margin-left:20px;\">
      <input type=\"checkbox\" id=\"intangiblesCheckbox\" { 'checked' if include_intangibles else '' } /> Include intangibles
    </label>
    <label style=\"margin-left:20px;\">
      <input type=\"checkbox\" id=\"hideUncheckedYears\" /> Hide unchecked years from plots
    </label>
  </div>
    <div>
    <details id=\"yearToggleDetails\">
      <summary style=\"font-weight:bold; cursor:pointer;\">Include Years in Stats</summary>
      <div id=\"yearToggleContainer\" style=\"display:flex; flex-wrap:wrap; row-gap:4px; column-gap:12px; padding-top:4px;\"></div>
    </details>
  </div>
</div>
<div id=\"plotBars\"></div>
<script>
const years = {years_json};
const tickers = {tickers_json};
const types = {types_json};
const baseRawData = {records_json};
const includeIntangiblesDefault = {str(include_intangibles).lower()};
let includeIntangibles = includeIntangiblesDefault;
const factorLookup = {factor_json};
const yearLabels = {year_labels_json};
const yearLabelMap = Object.fromEntries(yearLabels.map(y => [`${{y.year}}|${{y.ticker}}`, y.label]));
const releaseMap = {release_json};
const typeOffsets = {type_offsets_json};
const tickerOffsets = {ticker_offsets_json};
const yearToggleState = Object.fromEntries(yearLabels.map((y, idx) => [y.label, idx !== yearLabels.length - 1]));
let hideUncheckedYears = false;

function getPlotHeight() {{
  const controlsHeight = document.getElementById("controls")?.offsetHeight || 0;
  const headerHeight = document.querySelector("h2")?.offsetHeight || 0;
  const padding = 80; // remaining padding/margins
  const desired = window.innerHeight - controlsHeight - headerHeight - padding;
  return Math.max(desired, 500);
  }}

function hashColor(str) {{
  let hash = 0;
  for (let i = 0; i < str.length; i++) hash = str.charCodeAt(i) + ((hash << 5) - hash);
  const hue = Math.abs(hash) % 360;
  return `hsl(${{hue}},70%,55%)`;
}}

function humanReadable(val) {{
  if (val === undefined || val === null || isNaN(val)) return "0";
  const abs = Math.abs(val);
  if (abs >= 1e12) return (val / 1e12).toFixed(2) + 'T';
  if (abs >= 1e9)  return (val / 1e9).toFixed(2) + 'B';
  if (abs >= 1e6)  return (val / 1e6).toFixed(2) + 'M';
  if (abs >= 1e3)  return (val / 1e3).toFixed(2) + 'K';
  return val.toFixed(2);
}}

function filterIntangibles(data, include) {{
  if (include) return data.slice();
  return data.filter(r => (r.NOTE || "").toLowerCase() !== "intangibles");
}}

let rawData = filterIntangibles(baseRawData, includeIntangibles);

const colorMap = {{}};
rawData.forEach(r => {{
  const keyCandidate = (r.Key4Coloring && r.Key4Coloring.trim()) ? r.Key4Coloring.trim() : (r.ITEM || "");
  const fallback = (r.ITEM && r.ITEM.trim()) ? r.ITEM.trim() : "";
  const key4 = keyCandidate || fallback;
  const canonicalKey = `${{r.TYPE}}|${{key4}}`;
  if (!colorMap[canonicalKey]) {{
    colorMap[canonicalKey] = hashColor(canonicalKey);
  }}
  r._CANONICAL_KEY = canonicalKey;
}});

  function initFactorSelector() {{
    const sel = document.getElementById("factorSelector");
    const keys = Object.keys(factorLookup).filter(f => f !== "");
    const sortedKeys = keys
      .sort((a, b) => {{
        const na = Number(a);
        const nb = Number(b);
        const aNum = !isNaN(na);
        const bNum = !isNaN(nb);
        if (aNum && bNum) return na - nb;
        if (aNum) return -1;
        if (bNum) return 1;
        return String(a).localeCompare(String(b));
      }});
    sortedKeys.forEach(f => {{
      const opt = document.createElement("option");
      opt.value = f;
      opt.textContent = f;
      sel.appendChild(opt);
    }});
    const defaultKey = sortedKeys.length ? sortedKeys[sortedKeys.length - 1] : "";
    sel.value = defaultKey;
    sel.addEventListener("change", renderBars);
  }}

function initYearCheckboxes() {{
  const container = document.getElementById("yearToggleContainer");
  container.innerHTML = "";
  yearLabels.forEach(entry => {{
    const label = document.createElement("label");
    label.className = "year-toggle";
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.checked = yearToggleState[entry.label];
    checkbox.dataset.year = entry.year;
    checkbox.dataset.ticker = entry.ticker;
    checkbox.dataset.label = entry.label;
    checkbox.addEventListener("change", (ev) => {{
      const lbl = ev.target.dataset.label;
      yearToggleState[lbl] = ev.target.checked;
      renderBars();
    }});
    const span = document.createElement("span");
    span.textContent = entry.label;
    label.appendChild(checkbox);
    label.appendChild(span);
    container.appendChild(label);
  }});
}}

function getActiveYears() {{
  if (!hideUncheckedYears) return years.slice();
  const active = new Set();
  yearLabels.forEach(entry => {{
    if (yearToggleState[entry.label]) active.add(entry.year);
  }});
  return years.filter(y => active.has(y));
}}

function renderBars() {{
  const factorName = document.getElementById("factorSelector").value || "";
  const factorMap = factorLookup[factorName] || {{}};
  const activeYears = getActiveYears();
  const baseYears = activeYears.map((_, i) => i * 2.0);

  const data = rawData.filter(r => (r.NOTE || "").toLowerCase() !== "excluded");
  const traces = [];
  for (const ticker of tickers) {{
    for (const typ of types) {{
      const subset = data.filter(r => r.Ticker === ticker && r.TYPE === typ);
      for (const row of subset) {{
        const color = colorMap[row._CANONICAL_KEY];
        const yvals = activeYears.map((year) => {{
          const labelKey = `${{year}}|${{ticker}}`;
          const label = yearLabelMap[labelKey] || `${{year}} - ${{ticker}}`;
          if (!yearToggleState[label]) return NaN;
          const factorTickerMap = factorMap[ticker] || {{}};
          const factorVal = factorTickerMap[year];
          if (factorVal === undefined || factorVal === null || isNaN(factorVal)) return NaN;
          const baseVal = row[year] || 0;
          return baseVal * factorVal;
        }});
        if (yvals.every(v => isNaN(v))) continue;
        const xvals = baseYears.map(b => b + (typeOffsets[typ] || 0) + (tickerOffsets[ticker] || 0));
        const releaseDates = activeYears.map(year => releaseMap[ticker]?.[year] || "");
        const maxAbs = Math.max(...yvals.filter(v => !isNaN(v)).map(v => Math.abs(v)));
        const textThreshold = maxAbs / 3;
        const texts = yvals.map(v => {{
          if (isNaN(v)) return "";
          if (Math.abs(v) < textThreshold) return "";
          return humanReadable(v);
        }});
        traces.push({{
          x: xvals,
          y: yvals,
          customdata: releaseDates,
          type: "bar",
          width: 0.33,
          text: texts,
          textposition: "inside",
          textfont: {{ color: "#fff", size: 10 }},
          marker: {{ color, line: {{ width: 0.3, color: "#333" }} }},
          offsetgroup: ticker + "-" + typ + "-" + row._CANONICAL_KEY,
          hovertemplate: `${{typ}}<br>${{row.ITEM || ''}}<br>${{ticker}}<br>%{{customdata}}<br>%{{y}}<extra></extra>` ,
          _orderKey: Math.min(...xvals),
          _ticker: ticker,
          _type: typ,
          _years: activeYears,
        }});
      }}
    }}
  }}

  traces.sort((a, b) => (a._orderKey ?? 0) - (b._orderKey ?? 0));

  const baseSums = new Map();
  const stackTotals = new Map();
  const stackMeta = new Map();
  traces.forEach(tr => {{
    const bases = [];
    tr.x.forEach((xVal, idx) => {{
      const yVal = tr.y[idx];
      if (yVal === null || Number.isNaN(yVal)) {{
        bases.push(NaN);
        return;
      }}
      const key = `${{xVal}}`;
      const current = baseSums.get(key) || {{ pos: 0, neg: 0 }};
      const baseVal = yVal >= 0 ? current.pos : current.neg;
      if (yVal >= 0) {{
        current.pos += yVal;
      }} else {{
        current.neg += yVal;
      }}
      baseSums.set(key, current);
      bases.push(baseVal);

      if (tr._ticker) {{
        const totalKey = `${{xVal}}|${{tr._ticker}}`;
        const totals = stackTotals.get(totalKey) || {{ sum: 0, has: false }};
        totals.sum += yVal;
        totals.has = true;
        stackTotals.set(totalKey, totals);
        if (!stackMeta.has(totalKey)) {{
          const years = Array.isArray(tr._years) ? tr._years : [];
          const releases = Array.isArray(tr.customdata) ? tr.customdata : [];
          stackMeta.set(totalKey, {{
            year: years[idx],
            release: releases[idx],
            type: tr._type,
          }});
        }}
      }}
    }});
    tr.base = bases;
    delete tr._orderKey;
  }});

  const dotByTicker = new Map();
  stackTotals.forEach((totals, key) => {{
    if (!totals.has) return;
    const [xStr, ticker] = key.split("|");
    const xVal = Number(xStr);
    if (!dotByTicker.has(ticker)) dotByTicker.set(ticker, {{ x: [], y: [], year: [], release: [], type: [] }});
    const entry = dotByTicker.get(ticker);
    entry.x.push(xVal);
    entry.y.push(totals.sum);
    const meta = stackMeta.get(key) || {{}};
    entry.year.push(meta.year);
    entry.release.push(meta.release);
    entry.type.push(meta.type);
  }});

  dotByTicker.forEach((vals, ticker) => {{
    const customdata = vals.year.map((yr, idx) => [yr, vals.release[idx], vals.type[idx]]);
    traces.push({{
      x: vals.x,
      y: vals.y,
      customdata,
      type: "scatter",
      mode: "markers",
      marker: {{ color: hashColor(ticker), size: 12 }},
      hovertemplate: `${{ticker}}<br>%{{customdata[2] || ''}}<br>%{{customdata[0] || ''}}<br>Release: %{{customdata[1] || ''}}<br>Total: %{{y}}<extra></extra>`,
      showlegend: false,
    }});
  }});

  const annotations = activeYears.map((year, idx) => {{
    const baseX = baseYears[idx];
    const parts = tickers.map(t => `${{year}}-${{t}}: ${{releaseMap[t]?.[year] || ''}}`).join('<br>');
    return {{
      x: baseX,
      y: 0,
      xref: 'x',
      yref: 'paper',
      text: parts,
      showarrow: false,
      xanchor: 'center',
      yanchor: 'top',
      font: {{ size: 10, color: '#555' }},
      yshift: -30,
    }};
  }});

  Plotly.newPlot('plotBars', traces, {{
    title: 'Financial Values (per share)',
    barmode: 'stack',
    height: getPlotHeight(),
    xaxis: {{
      tickmode: 'array',
      tickvals: baseYears,
      ticktext: activeYears,
    }},
    hovermode: 'closest',
    showlegend: false,
    annotations,
  }});
}}

const intangiblesCheckbox = document.getElementById("intangiblesCheckbox");
if (intangiblesCheckbox) {{
  intangiblesCheckbox.checked = includeIntangiblesDefault;
  intangiblesCheckbox.addEventListener("change", (ev) => {{
    includeIntangibles = ev.target.checked;
    rawData = filterIntangibles(baseRawData, includeIntangibles);
    renderBars();
  }});
}}

const hideUncheckedCheckbox = document.getElementById("hideUncheckedYears");
if (hideUncheckedCheckbox) {{
  hideUncheckedCheckbox.addEventListener("change", (ev) => {{
    hideUncheckedYears = ev.target.checked;
    renderBars();
  }});
}}

window.addEventListener("resize", () => {{
  Plotly.relayout('plotBars', {{ height: getPlotHeight() }});
  Plotly.Plots.resize('plotBars');
}});

initFactorSelector();
initYearCheckboxes();
renderBars();
</script>
</body>
</html>
"""

    out_path.write_text(html, encoding="utf-8")
    try:
        webbrowser.open(f"file://{os.path.abspath(out_path)}")
    except Exception:
        pass
    return out_path
