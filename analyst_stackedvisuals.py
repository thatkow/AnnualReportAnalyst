import pandas as pd
import numpy as np
import os
import webbrowser
import json


def render_stacked_annual_report(
    df: pd.DataFrame,
    title: str = "Stacked Annual Report",
    factor_lookup: dict | None = None,
    factor_label: str = "Adjustment Factor",
    share_counts: dict | None = None,
    out_path: str = "stacked_annual_report.html",
):
    """
    Generates an interactive two-tab HTML report:
      1. Financial stacked bars (per-share toggle)
      2. Normalized share counts
    """

    if share_counts is None:
        raise ValueError("❌ 'share_counts' must be provided explicitly.")

    # Identify years and categorical fields
    year_cols = [c for c in df.columns if c[:2].isdigit() or c.startswith("31.")]
    tickers = sorted(df["Ticker"].dropna().unique())
    types = sorted(df["TYPE"].dropna().unique())

    # Default factor lookup
    if factor_lookup is None:
        factor_lookup = {"Normal": {y: 1.0 for y in year_cols}}

    records = []
    for _, r in df.iterrows():
        rec = {
            "Ticker": r.get("Ticker"),
            "TYPE": r.get("TYPE"),
            "CATEGORY": r.get("CATEGORY"),
            "SUBCATEGORY": r.get("SUBCATEGORY"),
            "ITEM": r.get("ITEM"),
        }
        for y in year_cols:
            val = r.get(y)
            rec[y] = float(val) if pd.notna(val) else 0.0
        records.append(rec)

    type_offsets = {t: round((i - (len(types) - 1) / 2) * 0.6, 2) for i, t in enumerate(types)}
    type_linestyles = {t: ("solid" if i % 2 == 0 else "dot") for i, t in enumerate(types)}

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8" />
<title>{title}</title>
<script src="https://cdn.plot.ly/plotly-2.31.1.min.js"></script>
<style>
body {{ font-family: sans-serif; margin: 40px; }}
#tabs {{ display: flex; border-bottom: 2px solid #ccc; margin-bottom: 10px; }}
.tab {{
  padding: 10px 20px;
  cursor: pointer;
  border: 1px solid #ccc;
  border-bottom: none;
  background: #f0f0f0;
  margin-right: 4px;
  border-radius: 6px 6px 0 0;
}}
.tab.active {{ background: white; border-bottom: 2px solid white; font-weight: bold; }}
.tab-content {{ display: none; }}
.tab-content.active {{ display: block; }}
#controls {{ margin: 15px 0; }}
</style>
</head>
<body>
<h2>{title}</h2>

<div id="tabs">
  <div id="tabBars" class="tab active">Financial Values</div>
  <div id="tabShare" class="tab">Share Count</div>
</div>

<div id="contentBars" class="tab-content active">
  <div id="controls">
    <label><b>Factor:</b></label>
    <select id="factorSelector"></select>
    <label style="margin-left:20px;">
      <input type="checkbox" id="perShareCheckbox" checked /> Per Share
    </label>
  </div>
  <div id="plotBars"></div>
</div>

<div id="contentShare" class="tab-content">
  <div id="plotShare"></div>
</div>

<script>
const years = {json.dumps(year_cols)};
const tickers = {json.dumps(tickers)};
const types = {json.dumps(types)};
const rawData = {json.dumps(records)};
const shareCounts = {json.dumps(share_counts)};
const factorLookup = {json.dumps(factor_lookup)};
const typeOffsets = {json.dumps(type_offsets)};
const typeLineStyles = {json.dumps(type_linestyles)};

// Add label for dropdown
const factorLabel = {json.dumps(factor_label)};
document.addEventListener("DOMContentLoaded", () => {{
  document.querySelector('label b').textContent = factorLabel + ":";
}});

// --- Human readable number formatter ---
function humanReadable(val) {{
  if (val === undefined || val === null || isNaN(val)) return "0";
  const abs = Math.abs(val);
  if (abs >= 1e12) return (val / 1e12).toFixed(2) + 'T';
  if (abs >= 1e9)  return (val / 1e9).toFixed(2) + 'B';
  if (abs >= 1e6)  return (val / 1e6).toFixed(2) + 'M';
  if (abs >= 1e3)  return (val / 1e3).toFixed(2) + 'K';
  return val.toFixed(2);
}}

function fmt(val, perShare) {{
  return humanReadable(val);
}}

function hashColor(str) {{
  let hash = 0;
  for (let i = 0; i < str.length; i++) hash = str.charCodeAt(i) + ((hash << 5) - hash);
  const hue = Math.abs(hash) % 360;
  return `hsl(${{hue}},70%,55%)`;
}}

const baseYears = years.map((_, i) => i * 2.0);
const tickerOffsets = Object.fromEntries(tickers.map((t, i) => [t, (i - ((tickers.length - 1) / 2)) * 0.25]));

const colorMap = {{}};
rawData.forEach(r => {{
  const id = r.CATEGORY + '-' + r.SUBCATEGORY + '-' + r.ITEM;
  if (!colorMap[id]) colorMap[id] = hashColor(id);
}});

const sel = document.getElementById("factorSelector");
Object.keys(factorLookup).forEach(f => {{
  const opt = document.createElement("option");
  opt.value = f;
  opt.textContent = f;
  sel.appendChild(opt);
}});
sel.value = Object.keys(factorLookup)[0];

function buildBarTraces(factorName, perShare) {{
  const factorMap = factorLookup[factorName];
  const traces = [];
  for (const ticker of tickers) {{
    for (const typ of types) {{
      const subset = rawData.filter(r => r.TYPE === typ && r.Ticker === ticker);
      for (const row of subset) {{
        const id = row.CATEGORY + '-' + row.SUBCATEGORY + '-' + row.ITEM;
        const color = colorMap[id];
        const yvals = years.map(y => {{
          const baseVal = (row[y] || 0) * (factorMap[y] || 1);
          return perShare && shareCounts[ticker]?.[y] ? baseVal / shareCounts[ticker][y] : baseVal;
        }});
        const xvals = baseYears.map(b => b + (typeOffsets[typ] || 0) + (tickerOffsets[ticker] || 0));
        traces.push({{
          x: xvals,
          y: yvals,
          type: "bar",
          marker: {{ color, line: {{ width: 0.3, color: "#333" }} }},
          offsetgroup: ticker + "-" + typ + "-" + id,
          text: yvals.map(v => fmt(v, perShare)),
          hovertemplate: "TICKER:" + ticker +
                         "<br>YEAR:%{{customdata[0]}}" +
                         "<br>TYPE:" + typ +
                         "<br>CATEGORY:" + row.CATEGORY +
                         "<br>SUBCATEGORY:" + row.SUBCATEGORY +
                         "<br>ITEM:" + row.ITEM +
                         "<br>VALUE:%{{text}}<extra></extra>",
          customdata: years.map(y => [y]),
          legendgroup: ticker
        }});
      }}
    }}
  }}
  return traces;
}}

function buildCumsumLines(factorName, perShare) {{
  const factorMap = factorLookup[factorName];
  const lines = [];
  for (const ticker of tickers) {{
    for (const typ of types) {{
      const perYearTotals = years.map(y => {{
        const subset = rawData.filter(r => r.TYPE === typ && r.Ticker === ticker);
        const sum = subset.reduce((acc, r) => acc + (r[y] || 0) * (factorMap[y] || 1), 0);
        return perShare && shareCounts[ticker]?.[y] ? sum / shareCounts[ticker][y] : sum;
      }});
      const xvals = baseYears.map(b => b + (typeOffsets[typ] || 0) + (tickerOffsets[ticker] || 0));
      const color = hashColor(ticker + typ);
      lines.push({{
        x: xvals,
        y: perYearTotals,
        mode: "lines+markers",
        line: {{ color, dash: typeLineStyles[typ] || "solid", width: 3 }},
        marker: {{ color, size: 8, symbol: "circle" }},
        text: perYearTotals.map(v => fmt(v, perShare)),
        hovertemplate: "TICKER:" + ticker +
                       "<br>YEAR:%{{customdata[0]}}" +
                       "<br>TYPE:" + typ +
                       "<br>TOTAL:%{{text}}<extra></extra>",
        customdata: years.map(y => [y]),
        showlegend: false
      }});
    }}
  }}
  return lines;
}}

function renderBars() {{
  const factorName = sel.value;
  const perShare = document.getElementById("perShareCheckbox").checked;
  const traces = [...buildBarTraces(factorName, perShare), ...buildCumsumLines(factorName, perShare)];
  const layout = {{
    barmode: "relative",
    height: 750,
    title: `Financial Values — Factor: ${{factorName}}`,
    yaxis: {{ title: perShare ? "Value (Per Share)" : "Value", zeroline: true }},
    xaxis: {{ title: "Date", tickvals: baseYears, ticktext: years, tickangle: 45,
              mirror: true, linecolor: "black", linewidth: 4 }},
    hoverlabel: {{ bgcolor: "white", font: {{ family: "Courier New" }} }},
    showlegend: false
  }};
  Plotly.newPlot("plotBars", traces, layout);
}}
renderBars();
sel.addEventListener("change", renderBars);
document.getElementById("perShareCheckbox").addEventListener("change", renderBars);

function buildShareTraces() {{
  const traces = [];
  for (const ticker of tickers) {{
    const sc = shareCounts[ticker];
    if (!sc) continue;
    const latestYear = years[years.length - 1];
    const latest = sc[latestYear];
    const normY = years.map(y => sc[y] / latest);
    const rawY = years.map(y => sc[y]);
    const color = hashColor(ticker);
    traces.push({{
      x: years,
      y: normY,
      mode: "lines+markers+text",
      text: rawY.map(v => humanReadable(v)),
      textposition: "top center",
      line: {{ width: 2, color }},
      marker: {{ color }},
      customdata: years.map((y, i) => [y, rawY[i]]),
      hovertemplate: "TICKER:" + ticker +
                     "<br>YEAR:%{{customdata[0]}}" +
                     "<br>NORMALIZED:%{{y:.2f}}" +
                     "<br>ORIGINAL:%{{text}}<extra></extra>",
      name: ticker
    }});
  }}
  return traces;
}}

// === Dynamically scale Y-axis based on min and max share values ===
const allShareVals = [];
for (const ticker of tickers) {{
  const sc = shareCounts[ticker];
  if (!sc) continue;
  for (const y of years) {{
    const v = sc[y];
    if (v !== undefined && v !== null && !isNaN(v)) allShareVals.push(v);
  }}
}}
let yMin = Math.min(...allShareVals);
let yMax = Math.max(...allShareVals);
const pad = (yMax - yMin) * 0.05;

Plotly.newPlot("plotShare", buildShareTraces(), {{
  height: 600,
  title: "Normalized Share Count",
  yaxis: {{ title: "Normalized Value", range: [yMin - pad, yMax + pad] }},
  xaxis: {{ title: "Date" }},
  hoverlabel: {{ bgcolor: "white", font: {{ family: "Courier New" }} }},
  showlegend: true
}});

function activateTab(tabId) {{
  document.querySelectorAll(".tab").forEach(t => t.classList.remove("active"));
  document.querySelectorAll(".tab-content").forEach(c => c.classList.remove("active"));
  if (tabId === "Bars") {{
    document.getElementById("tabBars").classList.add("active");
    document.getElementById("contentBars").classList.add("active");
  }} else {{
    document.getElementById("tabShare").classList.add("active");
    document.getElementById("contentShare").classList.add("active");
  }}
}}
document.getElementById("tabBars").onclick = () => activateTab("Bars");
document.getElementById("tabShare").onclick = () => activateTab("Share");
</script>
</body>
</html>"""

    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"✅ HTML report written to {out_path}")
    webbrowser.open(f"file://{os.path.abspath(out_path)}")


if __name__ == "__main__":
    df = pd.DataFrame({
        "Ticker": ["DART", "DART", "AGRI", "AGRI"],
        "TYPE": ["A", "B", "A", "B"],
        "CATEGORY": ["Revenue", "Expense", "Revenue", "Expense"],
        "SUBCATEGORY": ["Sales", "COGS", "Sales", "COGS"],
        "ITEM": ["Item1", "Item2", "Item1", "Item2"],
        "31.12.2020": [10, -5, 6, -3],
        "31.12.2021": [12, -6, 8, -4],
        "31.12.2022": [14, -7, 9, -5],
        "31.12.2023": [16, -8, 10, -6],
    })
    share_counts = {
        "DART": {"31.12.2020": 1000, "31.12.2021": 1200, "31.12.2022": 1400, "31.12.2023": 1600},
        "AGRI": {"31.12.2020": 900, "31.12.2021": 1100, "31.12.2022": 1300, "31.12.2023": 1500},
    }
    render_stacked_annual_report(
        df,
        title="Financial/Income Report Example",
        factor_lookup={
            "half": {y: 0.5 for y in ["31.12.2020", "31.12.2021", "31.12.2022", "31.12.2023"]},
            "normal": {y: 1.0 for y in ["31.12.2020", "31.12.2021", "31.12.2022", "31.12.2023"]},
            "double": {y: 2.0 for y in ["31.12.2020", "31.12.2021", "31.12.2022", "31.12.2023"]},
        },
        share_counts=share_counts,
        out_path="stacked_annual_report.html"
    )
