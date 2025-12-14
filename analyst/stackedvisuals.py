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
    factor_tooltip: dict | None = None,
    factor_tooltip_label: str = "Stock Factors",
    share_counts: dict | None = None,
    pdf_sources: dict | None = None,
    out_path: str = "stacked_annual_report.html",
    include_intangibles: bool = True,
    latest_price: float | None = None,
):
    """
    Generates an interactive two-tab HTML report:
      1. Financial stacked bars (per-share toggle)
      2. Normalized share counts
    """

    if share_counts is None:
        raise ValueError("❌ 'share_counts' must be provided explicitly.")

    pdf_sources = pdf_sources or {}

    # Identify years and categorical fields
    year_cols = [c for c in df.columns if c[:2].isdigit() or c.startswith("31.")]
    tickers = sorted(df["Ticker"].dropna().unique())
    types = sorted(df["TYPE"].dropna().unique())

    # Ensure default structures exist
    factor_lookup = factor_lookup or {}
    share_counts = share_counts or {}

    # Default factor lookup
    if factor_lookup is None:
        factor_lookup = {"Normal": {y: 1.0 for y in year_cols}}


    # Convert scalar factor entries (like "") to uniform dicts
    for k, v in list(factor_lookup.items()):
        if isinstance(v, (int, float)):
            # Broadcast scalar to all date columns
            factor_lookup[k] = {y: float(v) for y in df.columns if y[:2].isdigit() or y.startswith("31.")}

    # Guarantee at least one factor for safety
    if not factor_lookup:
        factor_lookup = {"": {y: 1.0 for y in df.columns if y[:2].isdigit() or y.startswith("31.")}}

    records = []
    for _, r in df.iterrows():
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
.year-toggle {{
  display: inline-flex;
  align-items: center;
  gap: 4px;
  margin-right: 12px;
}}
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
    <label style="margin-left:20px;">
      <input type="checkbox" id="intangiblesCheckbox" /> Include intangibles
    </label>
    <div id="yearToggleContainer" style="margin-top:10px;"></div>
    <!-- Per-ticker raw adjustment inputs (one per ticker, applied per TYPE to latest year) -->
    <span
      id="tickerAdjustments"
      style="margin-left:20px; display:inline-flex; gap:16px;
             align-items:flex-end; overflow-x:auto; max-width:100%;
             white-space:nowrap;"
    >
      <!-- filled dynamically -->
    </span>
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
const baseRawData = {json.dumps(records)};
const includeIntangiblesDefault = {str(include_intangibles).lower()};
let includeIntangibles = includeIntangiblesDefault;
let rawData = filterIntangibles(baseRawData, includeIntangibles);
const shareCounts = {json.dumps(share_counts)};
const factorLookup = {json.dumps(factor_lookup)};
const factorTooltip = {json.dumps(factor_tooltip)};
const factorTooltipLabel = {json.dumps(factor_tooltip_label)};
const pdfSources = {json.dumps(pdf_sources)};
const typeOffsets = {json.dumps(type_offsets)};
const typeLineStyles = {json.dumps(type_linestyles)};
const latestPrice = {json.dumps(latest_price)};
const yearToggleState = Object.fromEntries(
  years.map((y, idx) => [y, idx !== years.length - 1])
);
let hideUncheckedYears = true;

let adjustedRawData = rawData;   // rawData + synthetic adjustment rows
let sliderState = {{}};

function filterIntangibles(data, include) {{
  if (include) return data.slice();
  return data.filter(r => (r.NOTE || "").toLowerCase() !== "intangibles");
}}

function getActiveYears() {{
  if (!hideUncheckedYears) return years.slice();
  return years.filter(y => yearToggleState[y]);
}}


// Add label for dropdown
const factorLabel = {json.dumps(factor_label)};
document.addEventListener("DOMContentLoaded", () => {{
  document.querySelector('label b').textContent = factorLabel + ":";
  const intangiblesCheckbox = document.getElementById("intangiblesCheckbox");
  if (intangiblesCheckbox) {{
    intangiblesCheckbox.checked = includeIntangiblesDefault;
    intangiblesCheckbox.addEventListener("change", (ev) => {{
      includeIntangibles = ev.target.checked;
      rawData = filterIntangibles(baseRawData, includeIntangibles);
      renderBars();
    }});
  }}
  initTickerAdjustments();
  initYearCheckboxes();
  updateAdjustmentLabels();
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

const tickerOffsets = Object.fromEntries(tickers.map((t, i) => [t, (i - ((tickers.length - 1) / 2)) * 0.25]));

function buildBaseYears(yearList) {{
  return yearList.map((_, i) => i * 2.0);
}}

// --- Canonical colour mapping using mapped fields ---
const colorMap = {{}};
baseRawData.forEach(r => {{
  const keyCandidate = (r.Key4Coloring && r.Key4Coloring.trim()) ? r.Key4Coloring.trim() : (r.ITEM || "");
  const fallback = (r.ITEM && r.ITEM.trim()) ? r.ITEM.trim() : "";
  const key4 = keyCandidate || fallback;
  const canonicalKey = `${{r.TYPE}}|${{key4}}`;

  if (!colorMap[canonicalKey]) {{
    colorMap[canonicalKey] = hashColor(canonicalKey);
  }}

  // Attach the canonical key for later use
  r._CANONICAL_KEY = canonicalKey;
}});


// Initialize raw adjustment state per ticker
tickers.forEach(t => {{
  if (sliderState[t] === undefined) sliderState[t] = 0;
}});

// Build per-ticker RAW number inputs along the top toolbar
function initTickerAdjustments() {{
  const wrap = document.getElementById("tickerAdjustments");
  if (!wrap) return;
  wrap.innerHTML = "";

  tickers.forEach(ticker => {{
    if (sliderState[ticker] === undefined) sliderState[ticker] = 0;

    const container = document.createElement("span");
    container.style.display = "inline-flex";
    container.style.flexDirection = "column";
    container.style.alignItems = "flex-start";

    const label = document.createElement("span");
    label.textContent = ticker;
    label.style.fontWeight = "bold";

    const row = document.createElement("span");
    row.style.display = "inline-flex";
    row.style.alignItems = "center";
    row.style.gap = "4px";

    const input = document.createElement("input");
    input.type = "number";
    // raw value; no min/max, user can type any adjustment
    input.step = "any";
    input.value = sliderState[ticker];
    input.dataset.ticker = ticker;
    input.style.width = "70px";

    input.addEventListener("input", (ev) => {{
      const t = ev.target.dataset.ticker;
      const v = parseFloat(ev.target.value);
      sliderState[t] = isNaN(v) ? 0 : v;
      renderBars();
    }});

    const delta = document.createElement("span");
    delta.className = "ticker-delta";
    delta.dataset.ticker = ticker;
    delta.style.fontFamily = "monospace";
    delta.style.fontSize = "11px";
    delta.textContent = "Δ: 0";

    row.appendChild(input);
    row.appendChild(delta);
    container.appendChild(label);
    container.appendChild(row);
    wrap.appendChild(container);
  }});
}}

function initYearCheckboxes() {{
  const container = document.getElementById("yearToggleContainer");
  if (!container) return;
  container.innerHTML = "";

  const title = document.createElement("div");
  title.style.fontWeight = "bold";
  title.style.marginBottom = "4px";
  title.textContent = "Include Years in Stats:";
  container.appendChild(title);

  const wrap = document.createElement("div");
  wrap.style.display = "flex";
  wrap.style.flexWrap = "wrap";
  wrap.style.gap = "12px";

  tickers.forEach(ticker => {{
    const column = document.createElement("div");
    column.style.display = "flex";
    column.style.flexDirection = "column";
    column.style.rowGap = "4px";

    const heading = document.createElement("div");
    heading.textContent = ticker;
    heading.style.fontWeight = "bold";
    column.appendChild(heading);

    years.forEach(year => {{
      const label = document.createElement("label");
      label.className = "year-toggle";
      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.checked = yearToggleState[year];
      checkbox.dataset.year = year;
      checkbox.addEventListener("change", (ev) => {{
        const y = ev.target.dataset.year;
        yearToggleState[y] = ev.target.checked;
        container.querySelectorAll(`input[data-year='${{y}}']`).forEach(cb => {{
          if (cb !== ev.target) cb.checked = ev.target.checked;
        }});
        renderBars();
      }});
      const span = document.createElement("span");
      span.textContent = `${{year}} - ${{ticker}}`;
      label.appendChild(checkbox);
      label.appendChild(span);
      column.appendChild(label);
    }});

    wrap.appendChild(column);
  }});

  container.appendChild(wrap);
}}

function updateAdjustmentLabels() {{
  const labels = document.querySelectorAll(".ticker-delta");
  labels.forEach(el => {{
    const ticker = el.dataset.ticker;
    const raw = sliderState[ticker] || 0;
    el.textContent = "Δ: " + humanReadable(raw);
  }});
}}

const sel = document.getElementById("factorSelector");
const factorNames = Object.keys(factorLookup).filter(f => f !== "");
if (!factorNames.length && factorLookup.hasOwnProperty("")) {{
  factorNames.push("");
}}
factorNames.forEach(f => {{
  const opt = document.createElement("option");
  opt.value = f;
  opt.textContent = f;
  sel.appendChild(opt);
}});
const numericFactors = factorNames.map(f => ({{ name: f, num: Number(f) }})).filter(({{ num }}) => !Number.isNaN(num));
if (numericFactors.length) {{
  numericFactors.sort((a, b) => a.num - b.num);
  sel.value = numericFactors[numericFactors.length - 1].name;
}} else if (factorNames.length) {{
  sel.value = factorNames[0];
}}

function buildBarTraces(factorName, perShare, activeYears, activeBaseYears) {{
  const factorMap = factorLookup[factorName];
  const traces = [];
  for (const ticker of tickers) {{
    for (const typ of types) {{
      const subset = adjustedRawData.filter(r => r.TYPE === typ && r.Ticker === ticker);
      for (const row of subset) {{
        // Mapped consistent colour key
        const color = colorMap[row._CANONICAL_KEY];
        // Compute yvals, skipping NaN factor years entirely
        const yvals = activeYears.map(y => {{
          const factor = factorMap[y];
          if (factor === undefined || factor === null || isNaN(factor)) {{
            return NaN;
          }}
          const baseVal = (row[y] || 0) * factor;
          return perShare && shareCounts[ticker]?.[y]
            ? baseVal / shareCounts[ticker][y]
            : baseVal;
        }});

        // If all yvals are NaN, skip this bar entirely
        if (yvals.every(v => isNaN(v))) {{
          continue;
        }}
        const xvals = activeBaseYears.map(b => b + (typeOffsets[typ] || 0) + (tickerOffsets[ticker] || 0));
        traces.push({{
          x: xvals,
          y: yvals,
          type: "bar",
          marker: {{ color, line: {{ width: 0.3, color: "#333" }} }},
          offsetgroup: ticker + "-" + typ + "-" + row._CANONICAL_KEY,
          text: yvals.map(v => fmt(v, perShare)),
          hovertemplate: "TICKER:" + ticker +
                         "<br>YEAR:%{{customdata[0]}}" +
                         "<br>TYPE:" + typ +
                         "<br>CATEGORY:" + row.CATEGORY +
                         "<br>SUBCATEGORY:" + row.SUBCATEGORY +
                         "<br>ITEM:" + row.ITEM +
                         "<br>VALUE:%{{text}}<extra></extra>",
          customdata: activeYears.map(y => [y]),
          legendgroup: ticker
        }});
      }}
    }}
  }}
  return traces;
}}

function buildCumsumLines(factorName, perShare, activeYears, activeBaseYears) {{
  const factorMap = factorLookup[factorName];
  const lines = [];
  for (const ticker of tickers) {{
    for (const typ of types) {{
      const perYearTotals = activeYears.map(y => {{
        const factor = factorMap[y];
        if (factor === undefined || factor === null || isNaN(factor)) {{
          return NaN;
        }}
        const subset = adjustedRawData.filter(r => r.TYPE === typ && r.Ticker === ticker);
        const sum = subset.reduce((acc, r) => acc + (r[y] || 0) * factor, 0);
        return perShare && shareCounts[ticker]?.[y]
          ? sum / shareCounts[ticker][y]
          : sum;
      }});

      // If all totals are NaN, skip this cumsum line
      if (perYearTotals.every(v => isNaN(v))) {{
        continue;
      }}
      const xvals = activeBaseYears.map(b => b + (typeOffsets[typ] || 0) + (tickerOffsets[ticker] || 0));
      lines.push({{
        x: xvals,
        y: perYearTotals,
        mode: "markers",
        marker: {{ color: "#000", size: 8, symbol: "circle" }},
        text: activeYears.map((y, i) => {{
          let tooltipArr = factorTooltip?.[y];
          if (!Array.isArray(tooltipArr)) tooltipArr = tooltipArr ? [tooltipArr] : [];
          const tooltipFormatted = tooltipArr.length
            ? `<br><b>${{factorTooltipLabel}}:</b><br>` + tooltipArr.join("<br>")
            : "";
          const pdfName = pdfSources[`${{y}}:${{ticker}}`] || pdfSources[y];
          const pdfPath = (pdfName && ticker)
            ? encodeURI(`../${{ticker}}/openapiscrape/${{pdfName}}/PDF_FOLDER/${{typ}}.pdf`)
            : "";
          const pdfHint = pdfPath
            ? "<br><i>Double-click point to open PDF</i>"
            : "";
          return (
            "TICKER:" + ticker +
            "<br>YEAR:" + y +
            "<br>TYPE:" + typ +
            tooltipFormatted +
            "<br>TOTAL:" + fmt(perYearTotals[i], perShare) +
            pdfHint
          );
        }}),
        hoverinfo: "text",
        customdata: activeYears.map(y => {{
          const pdfName = pdfSources[`${{y}}:${{ticker}}`] || pdfSources[y];
          const pdfPath = (pdfName && ticker)
            ? encodeURI(`../${{ticker}}/openapiscrape/${{pdfName}}/PDF_FOLDER/${{typ}}.pdf`)
            : "";
          return [y, pdfPath];
        }}),
        showlegend: false
      }});
    }}
  }}
  return lines;
}}

function renderBars() {{
  const factorName = sel.value;
  const perShare = document.getElementById("perShareCheckbox").checked;

  const factorMap = factorLookup[factorName];
  let activeYears = getActiveYears();
  if (activeYears.length === 0) {{
    activeYears = years.slice();
  }}
  const activeBaseYears = buildBaseYears(activeYears);
  const latestYear = activeYears[activeYears.length - 1];

  // === Build synthetic adjustment rows per ticker (latest year only) ===
  // RAW behaviour (Option B): the entered value is applied PER TYPE (no splitting)
  const syntheticRows = [];
  for (const ticker of tickers) {{
    const rawAdj = sliderState[ticker] || 0;
    if (!rawAdj) continue;

    for (const typ of types) {{
      const baseRow = rawData.find(r => r.Ticker === ticker && r.TYPE === typ);
      if (!baseRow) continue;

      const newRow = {{
        Ticker: ticker,
        TYPE: typ,
        CATEGORY: "Adjustment",
        SUBCATEGORY: "Adjustment",
        ITEM: "Adjustment",
        Key4Coloring: baseRow.Key4Coloring,
        _CANONICAL_KEY: baseRow._CANONICAL_KEY
      }};

      years.forEach(y => {{
        newRow[y] = (y === latestYear) ? rawAdj : 0.0;
      }});

      syntheticRows.push(newRow);
    }}
  }}

  // Merge original + synthetic into adjustedRawData for this render
  adjustedRawData = rawData.concat(syntheticRows);

  const barTraces = buildBarTraces(factorName, perShare, activeYears, activeBaseYears);
  const cumsumLines = buildCumsumLines(factorName, perShare, activeYears, activeBaseYears);

  // === Compute per-ticker, per-type cumulative-sum data ===
  const cumsumMap = {{}};
  for (const ticker of tickers) {{
    for (const typ of types) {{
      const key = ticker + "::" + typ;
      const subset = adjustedRawData.filter(r => r.TYPE === typ && r.Ticker === ticker);
      if (subset.length === 0) continue;
      const vals = years.map(y => {{
        const sum = subset.reduce((acc, r) => acc + (r[y] || 0) * (factorMap[y] || 1), 0);
        const adj = perShare && shareCounts[ticker]?.[y] ? sum / shareCounts[ticker][y] : sum;
        return adj;
      }});
      // Exclude NaN-factor years from boxplot stats
      const filteredPairs = [];
      years.forEach((y, i) => {{
        if (!yearToggleState[y]) {{
          return;
        }}
        const factor = factorMap[y];
        const v = vals[i];
        if (factor !== undefined && factor !== null && !isNaN(factor) && v !== undefined && v !== null && !isNaN(v)) {{
          filteredPairs.push({{ year: y, value: v }});
        }}
      }});

      if (filteredPairs.length === 0) {{
        continue;
      }}   

      const cleanedVals = filteredPairs.slice().map(p => p.value);
      if (cleanedVals.length === 0) {{
        continue;
      }}

      cumsumMap[key] = cleanedVals;
    }}
  }}

  // === Build one boxplot per ticker+type ===
  const boxTraces = [];
  const baseX = activeBaseYears[activeBaseYears.length - 1] + 3.0;
  const spacing = 0.4;
  let i = 0;
  for (const key of Object.keys(cumsumMap)) {{
    const [ticker, typ] = key.split("::");   // ensure ticker & typ in scope
    const vals = cumsumMap[key];
    if (!vals || vals.length === 0) continue;
    const sorted = vals.slice().sort((a,b)=>a-b);
    const q1 = sorted[Math.floor(0.25*sorted.length)];
    const q2 = sorted[Math.floor(0.5*sorted.length)];
    const q3 = sorted[Math.floor(0.75*sorted.length)];
    const min = sorted[0];
    const max = sorted[sorted.length-1];
    const mean = vals.reduce((a,b)=>a+b,0)/vals.length;
    const color = hashColor(ticker + typ);

    // Determine latest raw (unfactored) total for ratio
    const subsetRaw = adjustedRawData.filter(r => r.TYPE === typ && r.Ticker === ticker);
    const latestYear = years[years.length - 1];
    let rawTotal = subsetRaw.reduce((acc, r) => acc + (r[latestYear] || 0), 0);

    // Apply per-share adjustment if checkbox ticked
    if (perShare && shareCounts[ticker]?.[latestYear]) {{
      const shares = shareCounts[ticker][latestYear];
      if (shares && !isNaN(shares) && shares !== 0) {{
        rawTotal = rawTotal / shares;
      }}
    }}

    const safeRatio = (stat) => {{
      if (!stat || isNaN(stat) || stat === 0) return {{ value: NaN, display: "–" }};
      if (rawTotal === 0 || isNaN(rawTotal)) return {{ value: NaN, display: "–" }};
      const ratio = rawTotal / stat;
      return {{ value: ratio, display: ratio.toFixed(2) }};
    }};

    const formatStat = (stat) => {{
      const {{ value, display }} = safeRatio(stat);
      const priceValid = latestPrice !== null && latestPrice !== undefined && !isNaN(latestPrice) && latestPrice > 0;
      const percentText = priceValid && !isNaN(value)
        ? `[${{((value / latestPrice) * 100).toFixed(2)}}%]`
        : "";
      return `(${{display}})${{percentText}}`;
    }};

    const tooltip =
      `<b>Ticker:</b> ${{ticker}}<br>` +
      `<b>Type:</b> ${{typ}}<br>` +
      `<b>Count:</b> ${{vals.length}}<br>` +
      `<b>Min:</b> ${{humanReadable(min)}} ${{formatStat(min)}}<br>` +
      `<b>Q1:</b> ${{humanReadable(q1)}} ${{formatStat(q1)}}<br>` +
      `<b>Median:</b> ${{humanReadable(q2)}} ${{formatStat(q2)}}<br>` +
      `<b>Q3:</b> ${{humanReadable(q3)}} ${{formatStat(q3)}}<br>` +
      `<b>Mean:</b> ${{humanReadable(mean)}} ${{formatStat(mean)}}<br>` +
      `<b>Max:</b> ${{humanReadable(max)}} ${{formatStat(max)}}<br>` +
      `<b>Latest raw total${{perShare ? " (per share)" : ""}}:</b> ${{humanReadable(rawTotal)}}` +
      (factorTooltip?.[years[years.length-1]]?.length
        ? "<br><b>Today's Price:</b> " +
          (factorTooltip[years[years.length-1]].find(e => e.startsWith("Today")) || "Today: NaN")
        : "");


    boxTraces.push({{
      y: vals,
      x: Array(vals.length).fill(baseX + i * spacing),
      name: `${{ticker}}-${{typ}} Box`,
      type: "box",
      marker: {{ color, opacity: 0.65 }},
      line: {{ color }},
      boxmean: true,
      boxpoints: "outliers",
      hovertemplate: tooltip + "<extra></extra>"
    }});

    // === NEW: place ❓ icon ABOVE THE BOX PLOT instead of latest stacked bar ===
    const boxX = baseX + i * spacing;
    const boxTop = sorted[sorted.length - 1] * 1.05;   // top whisker * 1.05

    const iconTrace = {{
      x: [boxX],
      y: [boxTop],
      mode: "text",
      text: ["❓"],
      textfont: {{ color: color, size: 20, family: "Arial Black" }},
      hovertemplate: tooltip + "<extra></extra>",
      hoverinfo: "text",
      showlegend: false,
      cliponaxis: false
    }};
    boxTraces.push(iconTrace);

    i++;
  }}
  const traces = [...barTraces, ...cumsumLines, ...boxTraces];

  // === Layout ===
  const layout = {{
    barmode: "relative",
    height: 750,
    title: `Financial Values — Factor: ${{factorName}}`,
    yaxis: {{ title: perShare ? "Value (Per Share)" : "Value", zeroline: true }},
    xaxis: {{
      title: "Date",
      tickvals: [...activeBaseYears, baseX + i * spacing],
      ticktext: [...activeYears, "Box Plots"],
      tickangle: 45,
      mirror: true,
      linecolor: "black",
      linewidth: 4
    }},
    hoverlabel: {{ bgcolor: "white", font: {{ family: "Courier New" }} }},
    showlegend: false,
    cliponaxis: false
  }};

    Plotly.newPlot("plotBars", traces, layout);
    const barsDiv = document.getElementById("plotBars");
    barsDiv.on("plotly_click", evt => {{
      const point = evt?.points?.[0];
      if (!point) return;
      const clickCount = evt.event?.detail || 0;
      if (clickCount < 2) return;
      const pdfPath = point.customdata?.[1];
      if (!pdfPath) return;
      window.open(pdfPath, "_blank", "noopener,noreferrer");
    }});
    // Update Δ labels based on current raw values
    updateAdjustmentLabels();
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

    return out_path


def render_stacked_comparison(
    df: pd.DataFrame,
    title: str = "Stacked Financial Comparison",
    factor_lookup: dict | None = None,
    factor_label: str = "Adjustment Factor",
    factor_tooltip: dict | None = None,
    factor_tooltip_label: str = "Stock Factors",
    pdf_sources: dict | None = None,
    out_path: str = "stacked_comparison_report.html",
    include_intangibles: bool = True,
    latest_price: float | None = None,
):
    if pdf_sources is None:
        pdf_sources = {}

    year_cols = [c for c in df.columns if c[:2].isdigit() or c.startswith("31.")]
    tickers = sorted(df["Ticker"].dropna().unique())
    types = sorted(df["TYPE"].dropna().unique())

    factor_lookup = factor_lookup or {}
    if not factor_lookup:
        factor_lookup = {"": {y: 1.0 for y in year_cols}}

    for k, v in list(factor_lookup.items()):
        if isinstance(v, (int, float)):
            factor_lookup[k] = {y: float(v) for y in year_cols}

    records = []
    for _, r in df.iterrows():
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
            val = r.get(y)
            rec[y] = float(val) if pd.notna(val) else 0.0
        records.append(rec)

    type_offsets = {t: round((i - (len(types) - 1) / 2) * 0.6, 2) for i, t in enumerate(types)}
    type_linestyles = {t: ("solid" if i % 2 == 0 else "dot") for i, t in enumerate(types)}

    html = f"""<!DOCTYPE html>
<html lang=\"en\">
<head>
<meta charset=\"utf-8\" />
<title>{title}</title>
<script src=\"https://cdn.plot.ly/plotly-2.31.1.min.js\"></script>
<style>
body {{ font-family: sans-serif; margin: 40px; }}
#controls {{ margin: 15px 0; }}
.year-toggle {{
  display: inline-flex;
  align-items: center;
  gap: 4px;
  margin-right: 12px;
}}
</style>
</head>
<body>
<h2>{title}</h2>

<div id=\"controls\">
  <label><b>Factor:</b></label>
  <select id=\"factorSelector\"></select>
  <label style=\"margin-left:20px;\">\n    <input type=\"checkbox\" id=\"intangiblesCheckbox\" /> Include intangibles
  </label>
  <div id=\"yearToggleContainer\" style=\"margin-top:10px;\"></div>
</div>
<div id=\"plotBars\"></div>

<script>
const years = {json.dumps(year_cols)};
const tickers = {json.dumps(tickers)};
const types = {json.dumps(types)};
const baseRawData = {json.dumps(records)};
const includeIntangiblesDefault = {str(include_intangibles).lower()};
let includeIntangibles = includeIntangiblesDefault;
let rawData = filterIntangibles(baseRawData, includeIntangibles);
const factorLookup = {json.dumps(factor_lookup)};
const factorTooltip = {json.dumps(factor_tooltip or {})};
const factorTooltipLabel = {json.dumps(factor_tooltip_label)};
const pdfSources = {json.dumps(pdf_sources)};
const typeOffsets = {json.dumps(type_offsets)};
const typeLineStyles = {json.dumps(type_linestyles)};
const latestPrice = {json.dumps(latest_price)};
const yearToggleState = Object.fromEntries(
  years.map((y) => [y, true])
);
let hideUncheckedYears = true;

function filterIntangibles(data, include) {{
  if (include) return data.slice();
  return data.filter(r => (r.NOTE || "").toLowerCase() !== "intangibles");
}}

function getActiveYears() {{
  if (!hideUncheckedYears) return years.slice();
  return years.filter(y => yearToggleState[y]);
}}

const factorLabel = {json.dumps(factor_label)};
document.addEventListener("DOMContentLoaded", () => {{
  document.querySelector('label b').textContent = factorLabel + ":";
  const intangiblesCheckbox = document.getElementById("intangiblesCheckbox");
  if (intangiblesCheckbox) {{
    intangiblesCheckbox.checked = includeIntangiblesDefault;
    intangiblesCheckbox.addEventListener("change", (ev) => {{
      includeIntangibles = ev.target.checked;
      rawData = filterIntangibles(baseRawData, includeIntangibles);
      renderBars();
    }});
  }}
  initYearCheckboxes();
}});

function humanReadable(val) {{
  if (val === undefined || val === null || isNaN(val)) return "0";
  const abs = Math.abs(val);
  if (abs >= 1e12) return (val / 1e12).toFixed(2) + 'T';
  if (abs >= 1e9)  return (val / 1e9).toFixed(2) + 'B';
  if (abs >= 1e6)  return (val / 1e6).toFixed(2) + 'M';
  if (abs >= 1e3)  return (val / 1e3).toFixed(2) + 'K';
  return val.toFixed(2);
}}

function fmt(val) {{
  return humanReadable(val);
}}

function hashColor(str) {{
  let hash = 0;
  for (let i = 0; i < str.length; i++) hash = str.charCodeAt(i) + ((hash << 5) - hash);
  const hue = Math.abs(hash) % 360;
  return `hsl(${{hue}},70%,55%)`;
}}

const tickerOffsets = Object.fromEntries(tickers.map((t, i) => [t, (i - ((tickers.length - 1) / 2)) * 0.25]));

function buildBaseYears(yearList) {{
  return yearList.map((_, i) => i * 2.0);
}}

const colorMap = {{}};
baseRawData.forEach(r => {{
  const keyCandidate = (r.Key4Coloring && r.Key4Coloring.trim()) ? r.Key4Coloring.trim() : (r.ITEM || "");
  const fallback = (r.ITEM && r.ITEM.trim()) ? r.ITEM.trim() : "";
  const key4 = keyCandidate || fallback;
  const canonicalKey = `${{r.TYPE}}|${{key4}}`;

  if (!colorMap[canonicalKey]) {{
    colorMap[canonicalKey] = hashColor(canonicalKey);
  }}

  r._CANONICAL_KEY = canonicalKey;
}});

function initYearCheckboxes() {{
  const container = document.getElementById("yearToggleContainer");
  if (!container) return;
  container.innerHTML = "";

  const title = document.createElement("div");
  title.style.fontWeight = "bold";
  title.style.marginBottom = "4px";
  title.textContent = "Include Years in Stats:";
  container.appendChild(title);

  const wrap = document.createElement("div");
  wrap.style.display = "flex";
  wrap.style.flexWrap = "wrap";
  wrap.style.gap = "12px";

  tickers.forEach(ticker => {{
    const column = document.createElement("div");
    column.style.display = "flex";
    column.style.flexDirection = "column";
    column.style.rowGap = "4px";

    const heading = document.createElement("div");
    heading.textContent = ticker;
    heading.style.fontWeight = "bold";
    column.appendChild(heading);

    years.forEach(year => {{
      const label = document.createElement("label");
      label.className = "year-toggle";
      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.checked = yearToggleState[year];
      checkbox.dataset.year = year;
      checkbox.addEventListener("change", (ev) => {{
        const y = ev.target.dataset.year;
        yearToggleState[y] = ev.target.checked;
        container.querySelectorAll(`input[data-year='${{y}}']`).forEach(cb => {{
          if (cb !== ev.target) cb.checked = ev.target.checked;
        }});
        renderBars();
      }});
      const span = document.createElement("span");
      span.textContent = `${{year}} - ${{ticker}}`;
      label.appendChild(checkbox);
      label.appendChild(span);
      column.appendChild(label);
    }});

    wrap.appendChild(column);
  }});

  container.appendChild(wrap);
}}

const sel = document.getElementById("factorSelector");
const factorNames = Object.keys(factorLookup).filter(f => f !== "");
if (!factorNames.length && factorLookup.hasOwnProperty("")) {{
  factorNames.push("");
}}
factorNames.forEach(f => {{
  const opt = document.createElement("option");
  opt.value = f;
  opt.textContent = f;
  sel.appendChild(opt);
}});
const numericFactors = factorNames.map(f => ({{ name: f, num: Number(f) }})).filter(({{ num }}) => !Number.isNaN(num));
if (numericFactors.length) {{
  numericFactors.sort((a, b) => a.num - b.num);
  sel.value = numericFactors[numericFactors.length - 1].name;
}} else if (factorNames.length) {{
  sel.value = factorNames[0];
}}

function getFactor(factorName, ticker, year) {{
  const factorMap = factorLookup[factorName] || {{}};
  const perTicker = factorMap[ticker] || {{}};
  const v = perTicker[year];
  return (v === undefined || v === null || isNaN(v)) ? NaN : v;
}}

function buildBarTraces(factorName, activeYears, activeBaseYears) {{
  const traces = [];
  for (const ticker of tickers) {{
    for (const typ of types) {{
      const subset = rawData.filter(r => r.TYPE === typ && r.Ticker === ticker);
      for (const row of subset) {{
        const color = colorMap[row._CANONICAL_KEY];
        const yvals = activeYears.map(y => {{
          const factor = getFactor(factorName, ticker, y);
          const baseVal = (row[y] || 0) * factor;
          return baseVal;
        }});

        if (yvals.every(v => isNaN(v))) {{
          continue;
        }}
        const xvals = activeBaseYears.map(b => b + (typeOffsets[typ] || 0) + (tickerOffsets[ticker] || 0));
        traces.push({{
          x: xvals,
          y: yvals,
          type: "bar",
          width: 1.2,
          marker: {{ color, line: {{ width: 0.3, color: "#333" }} }},
          offsetgroup: ticker + "-" + typ + "-" + row._CANONICAL_KEY,
          text: yvals.map(v => fmt(v)),
          hovertemplate: "TICKER:" + ticker +
                         "<br>YEAR:%{{customdata[0]}}" +
                         "<br>TYPE:" + typ +
                         "<br>CATEGORY:" + row.CATEGORY +
                         "<br>SUBCATEGORY:" + row.SUBCATEGORY +
                         "<br>ITEM:" + row.ITEM +
                         "<br>VALUE:%{{text}}<extra></extra>",
          customdata: activeYears.map(y => [y]),
          legendgroup: ticker
        }});
      }}
    }}
  }}
  return traces;
}}

function buildCumsumLines(factorName, activeYears, activeBaseYears) {{
  const lines = [];
  for (const ticker of tickers) {{
    for (const typ of types) {{
      const perYearTotals = activeYears.map(y => {{
        const factor = getFactor(factorName, ticker, y);
        const subset = rawData.filter(r => r.TYPE === typ && r.Ticker === ticker);
        const sum = subset.reduce((acc, r) => acc + (r[y] || 0) * factor, 0);
        return sum;
      }});

      if (perYearTotals.every(v => isNaN(v))) {{
        continue;
      }}
      const xvals = activeBaseYears.map(b => b + (typeOffsets[typ] || 0) + (tickerOffsets[ticker] || 0));
      lines.push({{
        x: xvals,
        y: perYearTotals,
        mode: "markers",
        marker: {{ color: "#000", size: 8, symbol: "circle" }},
        text: activeYears.map((y, i) => {{
          let tooltipArr = factorTooltip?.[y];
          if (!Array.isArray(tooltipArr)) tooltipArr = tooltipArr ? [tooltipArr] : [];
          const tooltipFormatted = tooltipArr.length
            ? `<br><b>${{factorTooltipLabel}}:</b><br>` + tooltipArr.join("<br>")
            : "";
            const pdfName = pdfSources[`${{y}}:${{ticker}}`] || pdfSources[y];
          const pdfPath = (pdfName && ticker)
            ? encodeURI(`../${{ticker}}/openapiscrape/${{pdfName}}/PDF_FOLDER/${{typ}}.pdf`)
            : "";
          const pdfHint = pdfPath
            ? "<br><i>Double-click point to open PDF</i>"
            : "";
          return (
            "TICKER:" + ticker +
            "<br>YEAR:" + y +
            "<br>TYPE:" + typ +
            tooltipFormatted +
            "<br>TOTAL:" + fmt(perYearTotals[i]) +
            pdfHint
          );
        }}),
        hoverinfo: "text",
        customdata: activeYears.map(y => {{
          const pdfName = pdfSources[`${{y}}:${{ticker}}`] || pdfSources[y];
          const pdfPath = (pdfName && ticker)
            ? encodeURI(`../${{ticker}}/openapiscrape/${{pdfName}}/PDF_FOLDER/${{typ}}.pdf`)
            : "";
          return [y, pdfPath];
        }}),
        showlegend: false
      }});
    }}
  }}
  return lines;
}}

function renderBars() {{
  const factorName = sel.value;

  let activeYears = getActiveYears();
  if (activeYears.length === 0) {{
    activeYears = years.slice();
  }}
  const activeBaseYears = buildBaseYears(activeYears);

  const barTraces = buildBarTraces(factorName, activeYears, activeBaseYears);
  const cumsumLines = buildCumsumLines(factorName, activeYears, activeBaseYears);

  const cumsumMap = {{}};
  for (const ticker of tickers) {{
    for (const typ of types) {{
      const key = ticker + "::" + typ;
      const subset = rawData.filter(r => r.TYPE === typ && r.Ticker === ticker);
      if (subset.length === 0) continue;
      const vals = years.map(y => {{
        const sum = subset.reduce((acc, r) => acc + (r[y] || 0) * (getFactor(factorName, ticker, y) || 1), 0);
        return sum;
      }});
      const filteredPairs = [];
      years.forEach((y, i) => {{
        if (!yearToggleState[y]) {{
          return;
        }}
        const factor = getFactor(factorName, ticker, y);
        const v = vals[i];
        if (factor !== undefined && factor !== null && !isNaN(factor) && v !== undefined && v !== null && !isNaN(v)) {{
          filteredPairs.push({{ year: y, value: v }});
        }}
      }});

      if (filteredPairs.length === 0) {{
        continue;
      }}

      const cleanedVals = filteredPairs.slice().map(p => p.value);
      if (cleanedVals.length === 0) {{
        continue;
      }}

      cleanedVals.sort((a, b) => a - b);
      const median = cleanedVals[Math.floor(cleanedVals.length / 2)];
      const q1 = cleanedVals[Math.floor(cleanedVals.length / 4)];
      const q3 = cleanedVals[Math.floor((cleanedVals.length * 3) / 4)];
      const min = Math.min(...cleanedVals);
      const max = Math.max(...cleanedVals);

      cumsumMap[key] = {{
        median,
        q1,
        q3,
        min,
        max,
        years: filteredPairs.map(p => p.year)
      }};
    }}
  }}

  const boxData = Object.entries(cumsumMap).map(([key, stats]) => {{
    const [ticker, typ] = key.split("::");
    const color = hashColor(ticker + typ);
    const label = `${{typ}} (${{ticker}})`;
    return {{
      type: "box",
      boxpoints: "all",
      y: stats.years,
      x: [stats.median, stats.q1, stats.q3, stats.min, stats.max],
      marker: {{ color }},
      line: {{ color }},
      name: label,
      hovertemplate: "TYPE:" + typ +
                     "<br>TICKER:" + ticker +
                     "<br>MEDIAN:%{{x}}<extra></extra>",
      orientation: "h"
    }};
  }});

  const traces = barTraces.concat(cumsumLines).concat(boxData);
  const tickText = activeYears.map(y => `${{y}}<br>${{tickers.join(" / ")}}`);

  const layout = {{
    height: 750,
    barmode: "stack",
    bargap: 0.15,
    title: "Financial Values",
    xaxis: {{ tickvals: buildBaseYears(activeYears), ticktext: tickText, title: "Date" }},
    yaxis: {{ title: "Value" }},
    hoverlabel: {{ bgcolor: "white", font: {{ family: "Courier New" }} }},
    showlegend: false,
    cliponaxis: false,
    margin: {{ l: 80, r: 80, t: 80, b: 80 }}
  }};

  Plotly.newPlot("plotBars", traces, layout);
  const barsDiv = document.getElementById("plotBars");
  barsDiv.on("plotly_click", evt => {{
    const point = evt?.points?.[0];
    if (!point) return;
    const clickCount = evt.event?.detail || 0;
    if (clickCount < 2) return;
    const pdfPath = point.customdata?.[1];
    if (!pdfPath) return;
    window.open(pdfPath, "_blank", "noopener,noreferrer");
  }});
}}
renderBars();
sel.addEventListener("change", renderBars);
</script>
</body>
</html>"""

    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"✅ Comparison HTML written to {out_path}")
    webbrowser.open(f"file://{os.path.abspath(out_path)}")

    return out_path


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
