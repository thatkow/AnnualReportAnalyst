# === stacked_annual_report.py ===
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots

def render_stacked_annual_report(df, share_count, years, tickers, types, title="Stacked Annual Report"):
    ticker_colors = {"X": "#1f77b4", "Y": "#2ca02c", "Z": "#d62728"}
    type_styles = {"A": "solid", "B": "dash"}

    def get_value(r, year, norm=False):
        return r[year] / share_count[r["Ticker"]][year] if norm else r[year]

    def build_main(norm=False):
        traces, xpos, sums = [], [], {}
        barw, tgap, ygap = 0.25, 0.05, 0.8
        for i, y in enumerate(years):
            base = i * (len(tickers)*len(types)*(barw+tgap) + ygap)
            for ti, tic in enumerate(tickers):
                for ty_i, ty in enumerate(types):
                    off = base + (ti*len(types)+ty_i)*(barw+tgap)
                    xpos.append((y, tic, ty, off))
        for y, tic, ty, x in xpos:
            sub = df[(df.TYPE==ty)&(df.Ticker==tic)]
            total = sum(get_value(r, y, norm) for _, r in sub.iterrows())
            pos, neg = sub[sub[y]>0].sort_values(by=y, key=lambda s:-s.abs()), sub[sub[y]<0].sort_values(by=y, key=lambda s:-s.abs())
            pb, nb = 0, 0
            for _, r in pos.iterrows():
                v = get_value(r, y, norm)
                traces.append(go.Bar(
                    x=[x], y=[v], base=[pb], width=barw,
                    name=f"{ty}-{tic}", meta={"TYPE":ty}, hoverinfo="text",
                    hovertext=f"<span style='font-family:Courier New;'>YEAR:{y}<br>TICKER:{tic}<br>TYPE:{ty}<br>ITEM:{r.ITEM}<br>VALUE:{v:,.2f}</span>"
                ))
                pb += v
            for _, r in neg.iterrows():
                v = get_value(r, y, norm)
                traces.append(go.Bar(
                    x=[x], y=[v], base=[nb], width=barw,
                    name=f"{ty}-{tic}", meta={"TYPE":ty}, hoverinfo="text",
                    hovertext=f"<span style='font-family:Courier New;'>YEAR:{y}<br>TICKER:{tic}<br>TYPE:{ty}<br>ITEM:{r.ITEM}<br>VALUE:{v:,.2f}</span>"
                ))
                nb += v
            sums.setdefault((ty,tic),[]).append((x,total))
        for (ty,tic), pts in sums.items():
            pts = sorted(pts,key=lambda p:p[0]); xs, ys = zip(*pts)
            color = ticker_colors.get(tic, "#000000")
            dash = type_styles.get(ty, "solid")
            traces.append(go.Scatter(
                x=xs, y=ys, mode="lines",
                line=dict(width=1.8, dash=dash, color=color),
                name=f"{ty}-{tic} line", meta={"TYPE":ty}))
            traces.append(go.Scatter(
                x=xs, y=ys, mode="markers+text",
                marker=dict(size=8, color=color),
                text=[f"{y/1000:.1f}k" for y in ys],
                textposition="top center", hoverinfo="skip",
                name=f"{ty}-{tic} dots", meta={"TYPE":ty}))
        return traces

    def build_share():
        traces = []
        for ticker in tickers:
            color = ticker_colors.get(ticker, "#000000")
            traces.append(go.Scatter(
                x=years, y=[share_count[ticker][y] for y in years],
                mode="lines+markers+text",
                line=dict(color=color, width=2),
                text=[f"{share_count[ticker][y]:.1f}" for y in years],
                textposition="top center", name=f"ShareCount {ticker}",
                hoverinfo="text",
                hovertext=[f"<span style='font-family:Courier New;'>TICKER:{ticker}<br>YEAR:{y}<br>COUNT:{share_count[ticker][y]:.2f}</span>" for y in years]
            ))
        return traces

    fig = make_subplots(rows=2, cols=1, shared_xaxes=True,
                        row_heights=[0.25, 0.75], vertical_spacing=0.03,
                        subplot_titles=("Share Count", title))

    main_raw, main_norm, share = build_main(False), build_main(True), build_share()
    for t in share: fig.add_trace(t, row=1, col=1)
    for t in main_raw + main_norm: fig.add_trace(t, row=2, col=1)

    n_share, n_raw, n_norm = len(share), len(main_raw), len(main_norm)
    modes, types_sel, norms = ["both","bars","dots"], ["*","A","B"], [False,True]

    def vis_mask(norm, mode, typ):
        vis = [True]*n_share
        body = ([True]*n_raw if not norm else [False]*n_raw) + ([False]*n_norm if not norm else [True]*n_norm)
        if typ != "*":
            for i, t in enumerate(fig.data[n_share:], start=n_share):
                tp = t.meta["TYPE"] if hasattr(t,"meta") and t.meta else None
                if tp not in (typ,None):
                    body[i-n_share] = False
        if mode == "bars":
            for i, t in enumerate(fig.data[n_share:], start=n_share):
                if isinstance(t, go.Scatter): body[i-n_share] = False
        elif mode == "dots":
            for i, t in enumerate(fig.data[n_share:], start=n_share):
                if not isinstance(t, go.Scatter): body[i-n_share] = False
        return vis + body

    combos = {(n,m,t): vis_mask(n,m,t) for n in norms for m in modes for t in types_sel}
    current = (False,"both","*")

    def button(label, change):
        n,m,t = list(current)
        n = change.get("norm", n)
        m = change.get("mode", m)
        t = change.get("type", t)
        return dict(label=label, method="update",
                    args=[{"visible": combos[(n,m,t)]}])

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
                 buttons=[button("TYPE * (All)", {"type": "*"}),
                          button("TYPE A", {"type": "A"}),
                          button("TYPE B", {"type": "B"})])
        ],
        yaxis=dict(title="Share Count"),
        yaxis2=dict(title="Value"),
        xaxis2=dict(title="Year"),
        showlegend=False
    )

    for i,v in enumerate(combos[current]):
        fig.data[i].visible=v

    fig.show()
    return fig

