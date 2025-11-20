from dataclasses import dataclass
from typing import Iterable, Sequence

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.patches import Patch

from analyst.data import Company


# ------------------------------------------------------------
# Helpers
# ------------------------------------------------------------

def ensure_interactive_backend() -> str:
    """Switch to an interactive backend when possible to allow ``Figure.show()``."""

    backend = plt.get_backend()
    if "agg" not in backend.lower():
        return backend

    for candidate in ("TkAgg", "Qt5Agg", "QtAgg", "MacOSX"):
        try:
            plt.switch_backend(candidate)
            return candidate
        except Exception:
            continue

    return backend


def sort_release_dates(dates: Iterable[str]):
    return sorted(
        dates,
        key=lambda value: pd.to_datetime(
            value, format="%d.%m.%Y", dayfirst=True, errors="coerce"
        ),
    )


def clean_numeric(dfblock):
    out = dfblock.copy().astype(str)
    out = out.map(lambda x: x.strip() if isinstance(x, str) else x)
    out = out.map(lambda x: x.replace(",", "") if isinstance(x, str) else x)
    out = out.replace({"": "0", " ": "0"})
    return out.apply(pd.to_numeric, errors="coerce").fillna(0)


def compute_adjusted_values(ticker, df):
    df = df.copy()
    date_cols = [c for c in df.columns if c[0].isdigit()]

    # --- Extract Stock/Prices divisor rows ---
    stock_rows = df[(df["TYPE"] == "Stock") & (df["CATEGORY"] == "Prices")].copy()
    stock_rows[date_cols] = clean_numeric(stock_rows[date_cols])
    divisors = stock_rows[date_cols].astype(float)

    # Use SUBCATEGORY for labeling
    subcat_labels = stock_rows["SUBCATEGORY"].astype(str).tolist()

    # --- Clean numeric for financial, income, shares + multipliers ---
    allowed = ["Income", "Financial", "Shares"]
    mult_names = [
        "Financial Multiplier", "Income Multiplier",
        "Shares Multiplier", "Stock Multiplier"
    ]

    df_clean = df.copy()
    for c in date_cols:
        mask = (
            (df_clean["TYPE"].isin(allowed)) |
            (df_clean["CATEGORY"].isin(mult_names))
        ) & (df_clean["NOTE"] != "excluded")

        df_clean.loc[mask, c] = clean_numeric(df_clean.loc[mask, [c]])[c]

    # --- Extract multipliers ---
    fin_mult    = df_clean[df_clean["CATEGORY"]=="Financial Multiplier"][date_cols].iloc[0].astype(float)
    inc_mult    = df_clean[df_clean["CATEGORY"]=="Income Multiplier"][date_cols].iloc[0].astype(float)
    shares_mult = df_clean[df_clean["CATEGORY"]=="Shares Multiplier"][date_cols].iloc[0].astype(float)
    stock_mult  = df_clean[df_clean["CATEGORY"]=="Stock Multiplier"][date_cols].iloc[0].astype(float)
    share_count = df_clean[df_clean["NOTE"]=="share_count"][date_cols].iloc[0].astype(float)

    # --- Filter usable rows ---
    df2 = df_clean[
        (~df_clean["CATEGORY"].isin(mult_names)) &
        (df_clean["NOTE"]!="share_count") &
        (df_clean["NOTE"]!="excluded")
    ].copy()

    df2.loc[df2["NOTE"]=="negated", date_cols] *= -1

    # --- Compute adjusted base values ---
    denom = (share_count * shares_mult * stock_mult).replace(0, float("nan"))
    final_df = pd.DataFrame(columns=date_cols)

    for idx, row in df2.iterrows():
        mult = fin_mult if row["TYPE"] == "Financial" else inc_mult
        final_df.loc[idx] = (row[date_cols].astype(float) * mult) / denom

    # --- Build grouped values for Financial or Income ---
    def build_group(type_name):
        sub_idx = df2[df2["TYPE"] == type_name].index
        grouped = []

        # 7 divisors → 7 groups
        for _, divisor_row in divisors.iterrows():
            divided = final_df.loc[sub_idx].divide(
                divisor_row.replace(0, float("nan")),
                axis=1
            )
            sums = divided.sum(skipna=True)
            grouped.append(sums.values)

        return grouped

    return {
        "ticker": ticker,
        "subcats": subcat_labels,
        "financial": build_group("Financial"),
        "income":    build_group("Income")
    }


def get_release_dates(df):
    date_cols = [c for c in df.columns if c[0].isdigit()]
    release = df[df["CATEGORY"] == "ReleaseDate"].iloc[0]
    return [str(release[c]) for c in date_cols[:7]]


# ------------------------------------------------------------
# Interlaced Ticker-Colored Boxplots
# ------------------------------------------------------------

def render_interlaced_boxplots(
    ticker1, groups1, ticker2, groups2, releasedate_labels
):
    """
    7 groups per ticker → 14 interlaced boxplots
    """

    inter_groups = []
    inter_colors = []
    inter_labels = []

    for i in range(7):
        # XRO (ticker1)
        inter_groups.append(groups1[i])
        inter_colors.append(plt.cm.tab10(0))
        inter_labels.append(releasedate_labels[i])

        # SEK (ticker2)
        inter_groups.append(groups2[i])
        inter_colors.append(plt.cm.tab10(1))
        inter_labels.append(releasedate_labels[i])

    fig, ax = plt.subplots(figsize=(16, 6))
    bp = ax.boxplot(inter_groups, labels=inter_labels, patch_artist=True)

    for patch, color in zip(bp['boxes'], inter_colors):
        patch.set_facecolor(color)

    ax.legend(
        handles=[
            Patch(color=plt.cm.tab10(0), label=ticker1.upper()),
            Patch(color=plt.cm.tab10(1), label=ticker2.upper()),
        ]
    )

    plt.xticks(rotation=45, ha='right')
    plt.title(f"Interlaced Boxplots — {ticker1.upper()} vs {ticker2.upper()}")
    plt.tight_layout()

    return fig


@dataclass
class FinancialBoxplots:
    """Container for the generated financial and income boxplots."""

    fig_fin: plt.Figure
    fig_inc: plt.Figure

    def show(self, *, block=True):
        """Display both figures using the active Matplotlib backend."""

        self.fig_fin.show()
        self.fig_inc.show()
        plt.show(block=block)


def financials_boxplots(companies: Sequence[Company]) -> FinancialBoxplots:
    """Generate interlaced boxplots for the first two provided companies."""

    if len(companies) < 2:
        raise ValueError("At least two companies are required to build boxplots.")

    company_a, company_b = companies[0], companies[1]

    ensure_interactive_backend()

    release_dates = sort_release_dates(
        set(
            get_release_dates(company_a.combined)
            + get_release_dates(company_b.combined)
        )
    )
    if len(release_dates) < 7:
        release_dates = sort_release_dates(
            get_release_dates(company_a.combined)
            + get_release_dates(company_b.combined)
        )
    release_labels = release_dates[:7]
    adjusted_a = compute_adjusted_values(company_a.ticker, company_a.combined)
    adjusted_b = compute_adjusted_values(company_b.ticker, company_b.combined)

    fig_fin = render_interlaced_boxplots(
        company_a.ticker,
        adjusted_a["financial"],
        company_b.ticker,
        adjusted_b["financial"],
        release_labels,
    )

    fig_inc = render_interlaced_boxplots(
        company_a.ticker,
        adjusted_a["income"],
        company_b.ticker,
        adjusted_b["income"],
        release_labels,
    )

    return FinancialBoxplots(fig_fin=fig_fin, fig_inc=fig_inc)


# ------------------------------------------------------------
# Equivalent of: int main()
# ------------------------------------------------------------

def main():
    from analyst.data import import_companies

    tickers = ["SEK.AX", "XRO.AX"]
    print(f"Loading data for {', '.join(tickers)}...")

    companies = import_companies(tickers)
    figures = financials_boxplots(companies)

    print("Showing plots...")
    figures.show(block=False)

    print("Done.")


# Python entry point
if __name__ == "__main__":
    main()
