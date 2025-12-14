from __future__ import annotations

from pathlib import Path
from typing import Dict, Iterable, List

import pandas as pd

from analyst.data import Company
from analyst.plots import _extract_multiplier, _pdf_source_map, _release_date_map, COMBINED_BASE_COLUMNS


def _prepare_company_financials(company: Company, include_intangibles: bool) -> dict:
    combined_df = company.combined
    ticker = company.ticker

    if combined_df.empty:
        raise ValueError(f"Combined dataframe for {ticker} is empty; generate data first.")

    df_all = combined_df.copy().fillna("")
    excluded_cols = set(COMBINED_BASE_COLUMNS + ["Ticker"])
    num_cols = [c for c in df_all.columns if c not in excluded_cols]

    for col in num_cols:
        series = df_all[col]
        df_all[col] = series.astype(str).str.replace(",", "", regex=False)

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

    share_counts: Dict[str, float] = {}
    share_rows = df[df["ITEM"].str.lower().str.contains("number of shares", na=False)]
    if not share_rows.empty:
        row = share_rows.iloc[0]
        for year in num_cols:
            raw_val = row.get(year)
            if pd.notna(raw_val):
                share_counts[year] = float(raw_val)
            else:
                raise ValueError(
                    f"Number of shares for year '{year}' is missing or NaN for {ticker}."
                )
    else:
        share_counts = {year: 1.0 for year in num_cols}

    df_plot = df[~df["ITEM"].str.lower().str.contains("number of shares", na=False)].copy()
    for year in num_cols:
        divisor = share_counts.get(year, 1.0)
        if divisor == 0:
            continue
        df_plot[year] = df_plot[year] / divisor

    if not include_intangibles:
        df_plot = df_plot[df_plot["NOTE"].str.lower() != "intangibles"]

    df_plot["Ticker"] = ticker

    release_map = _release_date_map(df_all, num_cols, company)
    pdf_map = _pdf_source_map(df_all, num_cols)

    return {
        "df": df_plot,
        "share_counts": share_counts,
        "release_map": release_map,
        "pdf_map": pdf_map,
    }


def compare_stacked_financials(
    companies: Iterable[Company], *, out_path: str | Path | None = None, include_intangibles: bool = True
) -> Path:
    companies_list: List[Company] = list(companies)
    if not companies_list:
        raise ValueError("At least one company must be provided for comparison.")

    processed = [_prepare_company_financials(c, include_intangibles) for c in companies_list]

    dfs = [entry["df"] for entry in processed]
    combined = pd.concat(dfs, ignore_index=True)

    year_cols = [c for c in combined.columns if c not in set(COMBINED_BASE_COLUMNS + ["Ticker"])]
    tickers = [c.ticker for c in companies_list]

    pdf_sources = {}
    factor_tooltip = {}
    for idx, company in enumerate(companies_list):
        release_map = processed[idx]["release_map"]
        for year, text in release_map.items():
            label = f"{year}-{company.ticker}"
            factor_tooltip[label] = [f"Release Date: {text} days"]
        for year, src in processed[idx]["pdf_map"].items():
            pdf_sources[f"{year}-{company.ticker}"] = src

    records = []
    for _, row in combined.iterrows():
        rec = {
            "Ticker": row.get("Ticker"),
            "TYPE": row.get("TYPE"),
            "CATEGORY": row.get("CATEGORY"),
            "SUBCATEGORY": row.get("SUBCATEGORY"),
            "ITEM": row.get("ITEM"),
            "NOTE": row.get("NOTE"),
            "Key4Coloring": row.get("Key4Coloring"),
        }
        for year in year_cols:
            rec[f"{year}-{row.get('Ticker')}"] = float(row.get(year)) if pd.notna(row.get(year)) else 0.0
        records.append(rec)

    date_labels = [f"{year}-{ticker}" for year in year_cols for ticker in tickers]

    visuals_dir = Path(out_path).expanduser().resolve().parent if out_path else companies_list[0].default_visuals_path().parent
    visuals_dir.mkdir(parents=True, exist_ok=True)
    out_path = Path(out_path) if out_path else visuals_dir / "stacked_comparison.html"

    _render_comparison_report(
        records,
        date_labels,
        tickers,
        include_intangibles,
        pdf_sources,
        factor_tooltip,
        out_path,
    )

    return Path(out_path)


def _render_comparison_report(
    records: List[dict],
    date_labels: List[str],
    tickers: List[str],
    include_intangibles: bool,
    pdf_sources: Dict[str, str],
    factor_tooltip: Dict[str, list[str]],
    out_path: Path,
):
    import json

    type_offsets = {}
    type_linestyles = {}
    types = sorted({r.get("TYPE", "") for r in records})
    for i, t in enumerate(types):
        type_offsets[t] = round((i - (len(types) - 1) / 2) * 0.6, 2)
        type_linestyles[t] = "solid" if i % 2 == 0 else "dot"

    html = """<!DOCTYPE html>
<html lang=\"en\">
<head>
<meta charset=\"utf-8\" />
<title>Stacked Financial Comparison</title>
<script src=\"https://cdn.plot.ly/plotly-2.31.1.min.js\"></script>
<style>
body {{ font-family: sans-serif; margin: 40px; }}
#controls {{ margin: 15px 0; }}
.year-toggle {{ display: inline-flex; align-items: center; gap: 4px; margin-right: 12px; }}
</style>
</head>
<body>
<h2>Financial Comparison (Per Share)</h2>

<div id=\"controls\">
  <label style=\"margin-right: 20px;\">
    <input type=\"checkbox\" id=\"intangiblesCheckbox\" /> Include intangibles
  </label>
  <label style=\"margin-right:20px;\">
    <input type=\"checkbox\" id=\"hideUncheckedYears\" /> Hide unchecked years from plots
  </label>
  <div id=\"yearToggleContainer\" style=\"margin-top:10px;\"></div>
</div>

<div id=\"plotBars\"></div>

<script>
const dateLabels = {date_labels};
const baseRawData = {records};
const includeIntangiblesDefault = {include_intangibles_default};
const typeOffsets = {type_offsets};
const typeLineStyles = {type_linestyles};
const tickers = {tickers};
const pdfSources = {pdf_sources};
const factorTooltip = {factor_tooltip};
const yearToggleState = Object.fromEntries(dateLabels.map(lbl => [lbl, true]));
let hideUncheckedYears = false;
let includeIntangibles = includeIntangiblesDefault;

function filterIntangibles(data, include) {{
  if (include) return data.slice();
  return data.filter(r => (r.NOTE || "").toLowerCase() !== "intangibles");
}}

function initYearCheckboxes() {{
  const container = document.getElementById('yearToggleContainer');
  container.innerHTML = '';
  dateLabels.forEach(lbl => {{
    const wrapper = document.createElement('label');
    wrapper.className = 'year-toggle';
    const cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.checked = yearToggleState[lbl];
    cb.addEventListener('change', () => {{
      yearToggleState[lbl] = cb.checked;
      renderBars();
    }});
    wrapper.appendChild(cb);
    const span = document.createElement('span');
    span.textContent = lbl;
    wrapper.appendChild(span);
    container.appendChild(wrapper);
  }});
}}

function getActiveLabels() {{
  if (!hideUncheckedYears) return dateLabels.slice();
  return dateLabels.filter(lbl => yearToggleState[lbl]);
}}

function humanReadable(val) {{
  if (val === undefined || val === null || isNaN(val)) return '0';
  const abs = Math.abs(val);
  if (abs >= 1e12) return (val / 1e12).toFixed(2) + 'T';
  if (abs >= 1e9)  return (val / 1e9).toFixed(2) + 'B';
  if (abs >= 1e6)  return (val / 1e6).toFixed(2) + 'M';
  if (abs >= 1e3)  return (val / 1e3).toFixed(2) + 'K';
  return val.toFixed(2);
}}

function hashColor(str) {{
  let hash = 0;
  for (let i = 0; i < str.length; i++) hash = str.charCodeAt(i) + ((hash << 5) - hash);
  const hue = Math.abs(hash) % 360;
  return `hsl(${{hue}},70%,55%)`;
}}

function buildTraces(activeData, activeLabels) {{
  const grouped = {{}};
  activeData.forEach(r => {{
    const keyCandidate = (r.Key4Coloring && r.Key4Coloring.trim()) ? r.Key4Coloring.trim() : (r.ITEM || '');
    const fallback = (r.ITEM && r.ITEM.trim()) ? r.ITEM.trim() : '';
    const key4 = keyCandidate || fallback;
    const canonicalKey = `${{r.TYPE}}|${{key4}}`;
    if (!grouped[canonicalKey]) grouped[canonicalKey] = {{x: [], y: [], text: [], name: key4 || r.ITEM || '', type: r.TYPE, ticker: r.Ticker}};
    activeLabels.forEach(lbl => {{
      grouped[canonicalKey].x.push(`${{lbl}}`);
      grouped[canonicalKey].y.push(r[lbl] || 0);
      const tool = [];
      if (factorTooltip[lbl]) tool.push(...factorTooltip[lbl]);
      if (pdfSources[lbl]) tool.push(`PDF: ${{pdfSources[lbl]}}`);
      grouped[canonicalKey].text.push(tool.join('<br>'));
    }});
  }});

  return Object.entries(grouped).map(([key, val]) => {{
    return {{
      x: val.x,
      y: val.y,
      text: val.text,
      hovertemplate: '%{{x}}<br>%{{y}}<br>%{{text}}<extra></extra>',
      name: `${{val.name}} (${{val.type}})`,
      type: 'bar',
      marker: {{color: hashColor(key)}},
      offsetgroup: val.type,
    }};
  }});
}}

function renderBars() {{
  const activeLabels = getActiveLabels();
  let data = filterIntangibles(baseRawData, includeIntangibles);
  const traces = buildTraces(data, activeLabels);
  const layout = {{
    barmode: 'relative',
    xaxis: {{title: 'Financial Date - Ticker'}},
    yaxis: {{title: 'Per Share Value'}},
    legend: {{orientation: 'h'}},
  }};
  Plotly.newPlot('plotBars', traces, layout, {{responsive: true}});
}}

document.addEventListener('DOMContentLoaded', () => {{
  const intangiblesCheckbox = document.getElementById('intangiblesCheckbox');
  if (intangiblesCheckbox) {{
    intangiblesCheckbox.checked = includeIntangiblesDefault;
    intangiblesCheckbox.addEventListener('change', ev => {{
      includeIntangibles = ev.target.checked;
      renderBars();
    }});
  }}
  const hideUncheckedCheckbox = document.getElementById('hideUncheckedYears');
  if (hideUncheckedCheckbox) {{
    hideUncheckedCheckbox.addEventListener('change', ev => {{
      hideUncheckedYears = ev.target.checked;
      renderBars();
    }});
  }}
  initYearCheckboxes();
  renderBars();
}});

</script>
</body>
</html>
""".format(
        date_labels=json.dumps(date_labels),
        records=json.dumps(records),
        include_intangibles_default=str(include_intangibles).lower(),
        type_offsets=json.dumps(type_offsets),
        type_linestyles=json.dumps(type_linestyles),
        tickers=json.dumps(tickers),
        pdf_sources=json.dumps(pdf_sources),
        factor_tooltip=json.dumps(factor_tooltip),
    )

    out_path = Path(out_path).expanduser().resolve()
    out_path.write_text(html, encoding="utf-8")

