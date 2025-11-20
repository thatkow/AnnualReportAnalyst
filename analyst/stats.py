from dataclasses import dataclass
from typing import Iterable, Sequence

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.patches import Patch

from analyst.data import Company
from . import yahoo


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
        "income":    build_group("Income"),
        "dates": date_cols,
    }


def get_release_dates(df):
    date_cols = [c for c in df.columns if c[0].isdigit()]
    release = df[df["CATEGORY"] == "ReleaseDate"].iloc[0]
    return [str(release[c]) for c in date_cols[:7]]


def get_latest_stock_price(ticker: str) -> float | None:
    """Fetch the most recent closing stock price for the given ticker.

    Returns ``None`` if the data cannot be retrieved.
    """

    try:
        prices = yahoo.get_stock_prices(ticker, years=1)
        if prices.empty:
            print(f"⚠️ No recent stock prices found for {ticker}.")
            return None
        latest = float(prices["Price"].iloc[-1])
        return latest
    except Exception as exc:  # pragma: no cover - network dependent
        print(f"⚠️ Failed to fetch latest stock price for {ticker}: {exc}")
        return None


# ------------------------------------------------------------
# Interlaced Ticker-Colored Boxplots
# ------------------------------------------------------------

def render_interlaced_boxplots(
    ticker1,
    groups1,
    ticker2,
    groups2,
    price_labels,
    *,
    xlabel: str,
    hlines: list[tuple[float, str, str]] | None = None,
):
    """
    7 groups per ticker → 14 interlaced boxplots
    """

    inter_groups = []
    inter_colors = []
    inter_labels = []

    for i in range(len(price_labels)):
        # XRO (ticker1)
        inter_groups.append(groups1[i])
        inter_colors.append(plt.cm.tab10(0))
        inter_labels.append(price_labels[i])

        # SEK (ticker2)
        inter_groups.append(groups2[i])
        inter_colors.append(plt.cm.tab10(1))
        inter_labels.append(price_labels[i])

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

    if hlines:
        for value, color, label in hlines:
            if value is None or np.isnan(value):
                continue
            ax.axhline(value, color=color, linestyle="--", linewidth=1.5, label=label)

    # Combine legend entries if horizontal lines added
    if hlines:
        handles, labels = ax.get_legend_handles_labels()
        ax.legend(handles, labels)

    ax.set_xlabel(xlabel)
    plt.xticks(rotation=45, ha='right')
    plt.title(f"Interlaced Boxplots — {ticker1.upper()} vs {ticker2.upper()}")
    plt.tight_layout()

    return fig


def compute_normalized_latest(groups, dates: Sequence[str], latest_price: float | None):
    """Calculate normalized cumulative sum for the latest available date column."""

    if latest_price is None or latest_price == 0 or pd.isna(latest_price):
        return None

    # Identify the latest date index using day-first parsing
    parsed_dates = [
        pd.to_datetime(value, format="%d.%m.%Y", dayfirst=True, errors="coerce")
        for value in dates
    ]
    if not parsed_dates or all(pd.isna(d) for d in parsed_dates):
        latest_idx = 0
    else:
        latest_idx = int(np.nanargmax(parsed_dates))

    latest_total = np.nansum(
        [grp[latest_idx] for grp in groups if len(grp) > latest_idx]
    )

    if np.isnan(latest_total):
        return None

    return latest_total / latest_price


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

    adjusted_a = compute_adjusted_values(company_a.ticker, company_a.combined)
    adjusted_b = compute_adjusted_values(company_b.ticker, company_b.combined)

    latest_price_a = get_latest_stock_price(company_a.ticker)
    latest_price_b = get_latest_stock_price(company_b.ticker)

    shared_price_labels = [
        label for label in adjusted_a["subcats"] if label in adjusted_b["subcats"]
    ]
    if not shared_price_labels:
        raise ValueError("No shared stock price labels found between companies")

    def filter_groups(groups, available_labels, target_labels):
        index_lookup = {label: i for i, label in enumerate(available_labels)}
        return [groups[index_lookup[label]] for label in target_labels]

    fin_groups_a = filter_groups(
        adjusted_a["financial"], adjusted_a["subcats"], shared_price_labels
    )
    fin_groups_b = filter_groups(
        adjusted_b["financial"], adjusted_b["subcats"], shared_price_labels
    )
    inc_groups_a = filter_groups(
        adjusted_a["income"], adjusted_a["subcats"], shared_price_labels
    )
    inc_groups_b = filter_groups(
        adjusted_b["income"], adjusted_b["subcats"], shared_price_labels
    )

    fin_line_a = compute_normalized_latest(
        fin_groups_a, adjusted_a["dates"], latest_price_a
    )
    fin_line_b = compute_normalized_latest(
        fin_groups_b, adjusted_b["dates"], latest_price_b
    )
    inc_line_a = compute_normalized_latest(
        inc_groups_a, adjusted_a["dates"], latest_price_a
    )
    inc_line_b = compute_normalized_latest(
        inc_groups_b, adjusted_b["dates"], latest_price_b
    )

    fig_fin = render_interlaced_boxplots(
        company_a.ticker,
        fin_groups_a,
        company_b.ticker,
        fin_groups_b,
        shared_price_labels,
        xlabel="Balance Sheet",
        hlines=[
            (fin_line_a, plt.cm.tab10(0), f"{company_a.ticker.upper()} latest norm."),
            (fin_line_b, plt.cm.tab10(1), f"{company_b.ticker.upper()} latest norm."),
        ],
    )

    fig_inc = render_interlaced_boxplots(
        company_a.ticker,
        inc_groups_a,
        company_b.ticker,
        inc_groups_b,
        shared_price_labels,
        xlabel="Income Statement",
        hlines=[
            (inc_line_a, plt.cm.tab10(0), f"{company_a.ticker.upper()} latest norm."),
            (inc_line_b, plt.cm.tab10(1), f"{company_b.ticker.upper()} latest norm."),
        ],
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
