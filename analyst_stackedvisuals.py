import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import re


def render_stacked_annual_report(df, title="Stacked Annual Report", share_count_note_name="share_count"):
    # --- Infer key dimensions ---
    print("🔍 Inferring configuration from DataFrame...")
    # Detect year-like columns (e.g. "30.06.2024")
    years = [c for c in df.columns if re.match(r"\d{2}\.\d{2}\.\d{4}", str(c))]
    if not years:
        raise ValueError("❌ No year-like columns found (expected format DD.MM.YYYY).")
    # Sort years chronologically (DD.MM.YYYY)
    from datetime import datetime
    try:
        years = sorted(
            years,
            key=lambda y: datetime.strptime(y, "%d.%m.%Y")
        )
    except Exception:
        years = sorted(years)

    print(f"📅 Years (sorted): {years}")
    # Sort years chronologically (DD.MM.YYYY)
    from datetime import datetime
    try:
        years = sorted(
            years,
            key=lambda y: datetime.strptime(y, "%d.%m.%Y")
        )
    except Exception:
        years = sorted(years)

    print(f"📅 Years (sorted): {years}")

    # Infer tickers if present, otherwise default to single
    tickers = sorted(df["Ticker"].dropna().unique().tolist()) if "Ticker" in df.columns else ["Default"]
    print(f"🏷️ Tickers: {tickers}")

    # Infer types
    types = sorted(df["TYPE"].dropna().unique().tolist()) if "TYPE" in df.columns else ["Default"]
    print(f"📂 Types: {types}")

    # --- Extract share_count values ---
    share_count = {}
    sc_rows = df[df["NOTE"].str.lower() == share_count_note_name.lower()] if "NOTE" in df.columns else pd.DataFrame()
    if sc_rows.empty:
        print("⚠️ No rows found with NOTE == 'share_count'. Using 1 for all years.")
        share_count = {tic: {y: 1 for y in years} for tic in tickers}
    else:
        for ticker in tickers:
            share_count[ticker] = {}
            sub = sc_rows[(sc_rows.get("Ticker", ticker) == ticker) | (not "Ticker" in sc_rows.columns)]
            if sub.empty:
                for y in years:
                    share_count[ticker][y] = 1
            else:
                for y in years:
                    try:
                        val = float(sub.iloc[0][y])
                    except Exception:
                        val = 1
                    share_count[ticker][y] = val
        print("✅ Share count extracted for each ticker.")

    # --- Dynamic color and style assignment ---
    base_colors = [
        "#1f77b4", "#2ca02c", "#d62728", "#9467bd", "#8c564b",
        "#e377c2", "#7f7f7f", "#bcbd22", "#17becf"
    ]
    ticker_colors = {tic: base_colors[i % len(base_colors)] for i, tic in enumerate(tickers)}
    style_patterns = ["solid", "dot", "dash", "longdash", "dashdot"]

    # Filter out special placeholder types like 'Shares' before plotting
    types = [ty for ty in types if str(ty).lower() not in ("shares", "share")]

    type_styles = {ty: style_patterns[i % len(style_patterns)] for i, ty in enumerate(types)}

    print(f"🎨 Colors: {ticker_colors}")
    print(f"📈 Styles: {type_styles}")

    # --- Helper functions ---
    def get_value(r, year, norm=False):
        try:
            val = float(r[year])
            if norm:
                t = r.get("Ticker", tickers[0])
                return val / share_count.get(t, {}).get(year, 1)
            return val
        except Exception:
            return 0

    # --- Build traces ---
    def build_main(norm=False):
        traces, xpos, sums = [], [], {}

        # Exclude share_count rows from visual data
        plot_df = df
        if "NOTE" in df.columns:
            plot_df = df[df["NOTE"].str.lower() != share_count_note_name.lower()]

        barw, tgap, ygap = 0.25, 0.05, 0.8
        for i, y in enumerate(years):
            base = i * (len(tickers) * len(types) * (barw + tgap) + ygap)
            for ti, tic in enumerate(tickers):
                for ty_i, ty in enumerate(types):
                    off = base + (ti * len(types) + ty_i) * (barw + tgap)
                    xpos.append((y, tic, ty, off))

        for y, tic, ty, x in xpos:
            sub = plot_df[(plot_df.TYPE == ty) & (plot_df.get("Ticker", tic) == tic)]

            # Use NaN-safe summation to prevent invalid totals
            vals = [get_value(r, y, norm) for _, r in sub.iterrows()]
            if len(vals) == 0:
                total = 0.0
            else:
                total = np.nansum(vals)

            if np.isnan(total):
                print(f"⚠️ Total still NaN for TYPE={ty}, Ticker={tic}, Year={y} (subset={len(sub)})")

            pos = sub[sub[y] > 0].sort_values(by=y, key=lambda s: -s.abs())
            neg = sub[sub[y] < 0].sort_values(by=y, key=lambda s: -s.abs())
            pb, nb = 0, 0
            for _, r in pos.iterrows():
                v = get_value(r, y, norm)
                traces.append(go.Bar(
                    x=[x], y=[v], base=[pb], width=barw,
                    name=f"{ty}-{tic}", meta={"TYPE":ty}, hoverinfo="text",
                    hovertext=(
                        f"<span style='font-family:Courier New;'>"
                        f"YEAR:{y}<br>"
                        f"TICKER:{tic}<br>"
                        f"TYPE:{ty}<br>"
                        f"CATEGORY:{r.CATEGORY}<br>"
                        f"SUBCATEGORY:{r.SUBCATEGORY}<br>"
                        f"ITEM:{r.ITEM}<br>"
                        f"VALUE:{v:,.2f}</span>"
                    )
                ))
                pb += v
            for _, r in neg.iterrows():
                v = get_value(r, y, norm)
                traces.append(go.Bar(
                    x=[x], y=[v], base=[nb], width=barw,
                    name=f"{ty}-{tic}", meta={"TYPE":ty}, hoverinfo="text",
                    hovertext=(
                        f"<span style='font-family:Courier New;'>"
                        f"YEAR:{y}<br>"
                        f"TICKER:{tic}<br>"
                        f"TYPE:{ty}<br>"
                        f"CATEGORY:{r.CATEGORY}<br>"
                        f"SUBCATEGORY:{r.SUBCATEGORY}<br>"
                        f"ITEM:{r.ITEM}<br>"
                        f"VALUE:{v:,.2f}</span>"
                    )
                ))
                nb += v
            sums.setdefault((ty, tic), []).append((x, total))

        for (ty, tic), pts in sums.items():
            pts = sorted(pts, key=lambda p: p[0])
            xs, ys = zip(*pts)
            color = ticker_colors.get(tic, "#000000")
            dash = type_styles.get(ty, "solid")
            traces.append(go.Scatter(
                x=xs, y=ys, mode="lines",
                line=dict(width=1.8, dash=dash, color=color),
                name=f"{ty}-{tic} line", meta={"TYPE":ty}))
            # Format readable numeric labels and tooltips for dots
            def format_val(v):
                if pd.isna(v) or abs(v) < 1e-8:
                    return ""
                av = abs(v)
                if av >= 1_000_000_000:
                    return f"{v/1_000_000_000:.2f}B"
                elif av >= 1_000_000:
                    return f"{v/1_000_000:.2f}M"
                elif av >= 1_000:
                    return f"{v/1_000:.2f}k"
                else:
                    return f"{v:.2f}"

            hover_texts = [
                f"<span style='font-family:Courier New;'>"
                f"YEAR:{years[i]}<br>"
                f"TICKER:{tic}<br>"
                f"TYPE:{ty}<br>"
                f"VALUE:{ys[i]:,.2f}</span>"
                for i in range(len(ys))
            ]

            traces.append(go.Scatter(
                x=xs,
                y=ys,
                mode="markers+text",
                marker=dict(size=8, color=color),
                text=[format_val(y) for y in ys],
                textposition="top center",
                hoverinfo="text",
                hovertext=hover_texts,
                name=f"{ty}-{tic} dots",
                meta={"TYPE": ty}
            ))
        return traces

    def build_share():
        traces = []
        for ticker in tickers:
            color = ticker_colors.get(ticker, "#000000")
            yvals = [share_count[ticker][y] for y in years]
            traces.append(go.Scatter(
                x=years, y=yvals,
                mode="lines+markers+text",
                line=dict(color=color, width=2),
                text=[f"{v:,.0f}" for v in yvals],
                textposition="top center", name=f"ShareCount {ticker}",
                hoverinfo="text",
                hovertext=[f"<span style='font-family:Courier New;'>TICKER:{ticker}<br>YEAR:{y}<br>COUNT:{share_count[ticker][y]:,.0f}</span>" for y in years]
            ))
        return traces
    fig = make_subplots(
        rows=2,
        cols=1,
        shared_xaxes=True,
        # Slightly shrink top graph and expand bottom one
        row_heights=[0.22, 0.78],
        # Increase gap to prevent title/xlabel overlap
        vertical_spacing=0.08,
        subplot_titles=("Share Count", title)
    )

    # Adjust title and label positioning for clarity
    fig.update_layout(
        margin=dict(t=120, b=80),
        title_y=0.97
    )

    main_raw, main_norm, share = build_main(False), build_main(True), build_share()
    for t in share: fig.add_trace(t, row=1, col=1)
    for t in main_raw + main_norm: fig.add_trace(t, row=2, col=1)

    n_share, n_raw, n_norm = len(share), len(main_raw), len(main_norm)
    modes, types_sel, norms = ["both", "bars", "dots"], ["*"] + types, [False, True]

    def vis_mask(norm, mode, typ):
        vis = [True]*n_share
        body = ([True]*n_raw if not norm else [False]*n_raw) + ([False]*n_norm if not norm else [True]*n_norm)
        if typ != "*":
            for i, t in enumerate(fig.data[n_share:], start=n_share):
                tp = t.meta.get("TYPE") if hasattr(t, "meta") and t.meta else None
                if tp not in (typ, None):
                    body[i-n_share] = False
        if mode == "bars":
            for i, t in enumerate(fig.data[n_share:], start=n_share):
                if isinstance(t, go.Scatter): body[i-n_share] = False
        elif mode == "dots":
            for i, t in enumerate(fig.data[n_share:], start=n_share):
                if not isinstance(t, go.Scatter): body[i-n_share] = False
        return vis + body

    combos = {(n, m, t): vis_mask(n, m, t) for n in norms for m in modes for t in types_sel}
    current = (False, "both", "*")

    def button(label, change):
        n, m, t = list(current)
        n = change.get("norm", n)
        m = change.get("mode", m)
        t = change.get("type", t)
        return dict(label=label, method="update", args=[{"visible": combos[(n, m, t)]}])

    fig.update_layout(
        height=950, width=1300, template="plotly_white",
        hoverlabel=dict(bgcolor="white", font_family="Courier New", font_size=12),
        updatemenus=[
            dict(type="buttons", direction="right", x=0.25, y=1.12,
                 buttons=[button("Raw", {"norm": False}),
                          button("Normalize", {"norm": True})]),
            dict(type="dropdown", direction="down", x=0.7, y=1.12,
                 buttons=[button("Bars + Dots/Lines", {"mode": "both"}),
                          button("Bars only", {"mode": "bars"}),
                          button("Dots/Lines only", {"mode": "dots"})]),
            dict(type="dropdown", direction="down", x=0.85, y=1.12,
                 buttons=[button("TYPE * (All)", {"type": "*"})] +
                         [button(f"TYPE {t}", {"type": t}) for t in types])
        ],
        yaxis=dict(title="Share Count"),
        yaxis2=dict(title="Value"),
        showlegend=False
    )

    # Compute approximate x positions for tick labels (midpoints of each year group)
    # These positions are already determined by build_main() ordering logic.
    # We'll re-create those bases consistently.
    barw, tgap, ygap = 0.25, 0.05, 0.8
    per_year_width = len(tickers) * len(types) * (barw + tgap) + ygap
    tick_positions = [
        (i * per_year_width) + (per_year_width / 2) - (ygap / 2)
        for i in range(len(years))
    ]

    fig.update_xaxes(
        tickmode="array",
        tickvals=tick_positions,
        ticktext=years,
        tickangle=45,
        title="Date (Financial Year End)"
    )

    for i, v in enumerate(combos[current]):
        fig.data[i].visible = v

    # Return the Plotly figure object instead of displaying it directly.
    # This allows embedding in Tkinter (e.g., via tkinterweb.HtmlFrame).
    return fig
