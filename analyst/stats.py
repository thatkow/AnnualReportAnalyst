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


def compute_adjusted_values(ticker, df, include_intangibles: bool = True):
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
    note_lower = df_clean["NOTE"].astype(str).str.lower()
    for c in date_cols:
        mask = (
            (df_clean["TYPE"].isin(allowed)) |
            (df_clean["CATEGORY"].isin(mult_names))
        ) & (note_lower != "excluded")

        if not include_intangibles:
            mask = mask & (note_lower != "intangibles")

        df_clean.loc[mask, c] = clean_numeric(df_clean.loc[mask, [c]])[c]

    # --- Extract multipliers ---
    fin_mult    = df_clean[df_clean["CATEGORY"]=="Financial Multiplier"][date_cols].iloc[0].astype(float)
    inc_mult    = df_clean[df_clean["CATEGORY"]=="Income Multiplier"][date_cols].iloc[0].astype(float)
    shares_mult = df_clean[df_clean["CATEGORY"]=="Shares Multiplier"][date_cols].iloc[0].astype(float)
    stock_mult  = df_clean[df_clean["CATEGORY"]=="Stock Multiplier"][date_cols].iloc[0].astype(float)
    share_count = (
        df_clean[note_lower == "share_count"][date_cols].iloc[0].astype(float)
    )

    # --- Filter usable rows ---
    df2 = df_clean[
        (~df_clean["CATEGORY"].isin(mult_names)) &
        (note_lower!="share_count") &
        (note_lower!="excluded")
    ].copy()

    if not include_intangibles:
        df2 = df2[note_lower.loc[df2.index] != "intangibles"].copy()

    df2.loc[df2["NOTE"].astype(str).str.lower()=="negated", date_cols] *= -1

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
        "divisors": [row.values for _, row in divisors.iterrows()],
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

    def _extract_latest(df: pd.DataFrame) -> float | None:
        if df.empty:
            return None

        latest_value = df["Price"].iloc[-1]

        # ``float`` on a single-element Series is deprecated, so coerce safely
        latest_numeric = pd.to_numeric(latest_value, errors="coerce")
        if isinstance(latest_numeric, pd.Series):
            latest_numeric = latest_numeric.iloc[0]

        if pd.isna(latest_numeric):
            return None

        return float(latest_numeric)

    try:
        prices = yahoo.get_stock_prices(ticker, years=1)
        latest = _extract_latest(prices)
        if latest is not None:
            return latest

        print(f"⚠️ No recent stock prices found for {ticker}; trying Stooq fallback.")
        fallback_prices = yahoo.get_stooq_prices(ticker)
        return _extract_latest(fallback_prices)
    except Exception as exc:  # pragma: no cover - network dependent
        print(f"⚠️ Failed to fetch latest stock price for {ticker}: {exc}")
        return None


# ------------------------------------------------------------
# Interlaced Ticker-Colored Boxplots
# ------------------------------------------------------------

def render_interlaced_boxplots(
    ticker_groups: Sequence[tuple[str, Sequence[Sequence[float]], str]],
    price_labels,
    *,
    xlabel: str,
    hlines: list[tuple[float, str, str]] | None = None,
):
    """Render interlaced boxplots for ``n`` tickers.

    ``ticker_groups`` contains tuples of ``(ticker, groups, color)`` where
    ``groups`` is already filtered to ``price_labels`` order.
    """

    inter_groups: list[Sequence[float]] = []
    inter_colors: list[str] = []
    inter_labels: list[str] = []

    for i in range(len(price_labels)):
        for _, groups, color in ticker_groups:
            inter_groups.append(groups[i])
            inter_colors.append(color)
            inter_labels.append(price_labels[i])

    fig, ax = plt.subplots(figsize=(16, 6))
    bp = ax.boxplot(inter_groups, labels=inter_labels, patch_artist=True)

    for patch, color in zip(bp["boxes"], inter_colors):
        patch.set_facecolor(color)

    legend_handles = [
        Patch(color=color, label=ticker.upper()) for ticker, _, color in ticker_groups
    ]
    ax.legend(handles=legend_handles)

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
    plt.xticks(rotation=45, ha="right")
    title_tickers = ", ".join([ticker.upper() for ticker, _, _ in ticker_groups])
    plt.title(f"Interlaced Boxplots — {title_tickers}")
    plt.tight_layout()

    return fig


def _interleave_groups(
    ticker_groups: Sequence[tuple[str, Sequence[Sequence[float]], str]],
    price_labels: Sequence[str],
):
    inter_groups: list[Sequence[float]] = []
    inter_colors: list[str] = []
    inter_labels: list[str] = []

    for i in range(len(price_labels)):
        for _, groups, color in ticker_groups:
            inter_groups.append(groups[i])
            inter_colors.append(color)
            inter_labels.append(price_labels[i])

    return inter_groups, inter_colors, inter_labels


def render_interlaced_violin(
    ticker_groups_include: Sequence[tuple[str, Sequence[Sequence[float]], str]],
    ticker_groups_exclude: Sequence[tuple[str, Sequence[Sequence[float]], str]],
    price_labels,
    *,
    xlabel: str,
    hlines_include: list[tuple[float, str, str]] | None = None,
    hlines_exclude: list[tuple[float, str, str]] | None = None,
):
    """Render side-by-side interlaced violins for intangible on/off views.

    The left half shows distributions with ``include_intangibles=True`` and the
    right half shows ``include_intangibles=False`` using the same price label
    ordering for easy comparison. Horizontal reference lines can be specified
    independently for each half.
    """

    fig, axes = plt.subplots(1, 2, figsize=(18, 6), sharey=True)

    def plot_half(
        ax, ticker_groups, title: str, hlines: list[tuple[float, str, str]] | None
    ):
        inter_groups, inter_colors, inter_labels = _interleave_groups(
            ticker_groups, price_labels
        )

        positions = np.arange(1, len(inter_groups) + 1)
        vp = ax.violinplot(
            inter_groups,
            positions=positions,
            showmeans=True,
            showextrema=False,
            widths=0.85,
        )

        for body, color in zip(vp["bodies"], inter_colors):
            body.set_facecolor(color)
            body.set_edgecolor("black")
            body.set_alpha(0.6)

        if hlines:
            for value, color, label in hlines:
                if value is None or np.isnan(value):
                    continue
                ax.axhline(value, color=color, linestyle="--", linewidth=1.5, label=label)

        ax.set_xticks(positions)
        ax.set_xticklabels(inter_labels, rotation=45, ha="right")
        ax.set_title(title)
        ax.set_xlabel(xlabel)

    plot_half(axes[0], ticker_groups_include, "Include intangibles", hlines_include)
    plot_half(axes[1], ticker_groups_exclude, "Exclude intangibles", hlines_exclude)

    legend_handles = [
        Patch(color=color, label=ticker.upper())
        for ticker, _, color in ticker_groups_include
    ]
    fig.legend(handles=legend_handles, loc="upper center", ncol=len(legend_handles))
    fig.suptitle("Interlaced Violins — include_intangibles comparison")
    fig.tight_layout(rect=(0, 0, 1, 0.92))

    return fig


def compute_normalized_latest(
    groups: Sequence[Sequence[float]],
    divisors: Sequence[Sequence[float]],
    dates: Sequence[str],
    latest_price: float | None,
):
    """Calculate normalized cumulative sum for the latest available date column.

    This mirrors the adjusted value calculation but swaps the stock price divisor
    for the latest fetched price so the reference line aligns with current
    pricing.
    """

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

    latest_totals: list[float] = []
    for grp, divisor in zip(groups, divisors):
        if len(grp) <= latest_idx or len(divisor) <= latest_idx:
            continue

        base_value = grp[latest_idx] * divisor[latest_idx]
        latest_totals.append(base_value)

    latest_total = latest_totals[len(latest_totals) - 1]  # Sum over all subcategories

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


@dataclass
class FinancialViolins:
    """Container for the generated include/exclude intangibles violin plots."""

    fig_fin: plt.Figure
    fig_inc: plt.Figure

    def show(self, *, block=True):
        """Display both figures using the active Matplotlib backend."""

        self.fig_fin.show()
        self.fig_inc.show()
        plt.show(block=block)


def financials_boxplots(
    companies: Sequence[Company], *, include_intangibles: bool = True
) -> FinancialBoxplots:
    """Generate interlaced boxplots for all provided companies."""

    if not companies:
        raise ValueError("At least one company is required to build boxplots.")

    ensure_interactive_backend()

    # Precompute adjustments and latest prices
    adjusted_list = [
        compute_adjusted_values(
            company.ticker, company.combined, include_intangibles=include_intangibles
        )
        for company in companies
    ]
    latest_prices = [get_latest_stock_price(company.ticker) for company in companies]

    # Normalise subcats so strings '1' and '1.0' match
    def normalise(label):
        try:
            return str(int(float(label)))   # "1.0" -> 1 -> "1"
        except:
            return label
        
    for adjusted in adjusted_list:
        adjusted["subcats"] = [normalise(s) for s in adjusted["subcats"]]

    # Determine shared price labels across all companies preserving first-company order
    first_labels = adjusted_list[0]["subcats"]
    shared_price_labels = [
        label
        for label in first_labels
        if all(label in adjusted["subcats"] for adjusted in adjusted_list[1:])
    ]

    if not shared_price_labels:
        raise ValueError("No shared stock price labels found between companies")

    def filter_groups(groups, available_labels, target_labels):
        index_lookup = {label: i for i, label in enumerate(available_labels)}
        return [groups[index_lookup[label]] for label in target_labels]

    colors = [plt.cm.tab10(i % 10) for i in range(len(companies))]

    fin_ticker_groups = []
    inc_ticker_groups = []
    fin_hlines = []
    inc_hlines = []

    for company, adjusted, price, color in zip(
        companies, adjusted_list, latest_prices, colors
    ):
        fin_groups = filter_groups(
            adjusted["financial"], adjusted["subcats"], shared_price_labels
        )
        inc_groups = filter_groups(
            adjusted["income"], adjusted["subcats"], shared_price_labels
        )
        shared_divisors = filter_groups(
            adjusted["divisors"], adjusted["subcats"], shared_price_labels
        )

        fin_ticker_groups.append((company.ticker, fin_groups, color))
        inc_ticker_groups.append((company.ticker, inc_groups, color))

        fin_line = compute_normalized_latest(
            fin_groups, shared_divisors, adjusted["dates"], price
        )
        inc_line = compute_normalized_latest(
            inc_groups, shared_divisors, adjusted["dates"], price
        )

        fin_hlines.append(
            (fin_line, color, f"{company.ticker.upper()} latest norm.")
        )
        inc_hlines.append(
            (inc_line, color, f"{company.ticker.upper()} latest norm.")
        )

    fig_fin = render_interlaced_boxplots(
        fin_ticker_groups,
        shared_price_labels,
        xlabel="Balance Sheet",
        hlines=fin_hlines,
    )

    fig_inc = render_interlaced_boxplots(
        inc_ticker_groups,
        shared_price_labels,
        xlabel="Income Statement",
        hlines=inc_hlines,
    )

    return FinancialBoxplots(fig_fin=fig_fin, fig_inc=fig_inc)


def financials_violin_comparison(companies: Sequence[Company]) -> FinancialViolins:
    """Compare include_intangibles on/off views via interlaced violins."""

    if not companies:
        raise ValueError("At least one company is required to build violins.")

    ensure_interactive_backend()

    adjusted_include = [
        compute_adjusted_values(company.ticker, company.combined, include_intangibles=True)
        for company in companies
    ]
    adjusted_exclude = [
        compute_adjusted_values(company.ticker, company.combined, include_intangibles=False)
        for company in companies
    ]
    latest_prices = [get_latest_stock_price(company.ticker) for company in companies]

    def normalise(label):
        try:
            return str(int(float(label)))
        except Exception:
            return label

    for adjusted in (*adjusted_include, *adjusted_exclude):
        adjusted["subcats"] = [normalise(s) for s in adjusted["subcats"]]

    first_labels = adjusted_include[0]["subcats"]
    shared_price_labels = [
        label
        for label in first_labels
        if all(label in adjusted["subcats"] for adjusted in adjusted_include[1:])
        and all(label in adjusted["subcats"] for adjusted in adjusted_exclude)
    ]

    if not shared_price_labels:
        raise ValueError(
            "No shared stock price labels found between companies for violin comparison"
        )

    def filter_groups(groups, available_labels, target_labels):
        index_lookup = {label: i for i, label in enumerate(available_labels)}
        return [groups[index_lookup[label]] for label in target_labels]

    colors = [plt.cm.tab10(i % 10) for i in range(len(companies))]

    fin_include_groups = []
    fin_exclude_groups = []
    inc_include_groups = []
    inc_exclude_groups = []
    fin_hlines_include = []
    fin_hlines_exclude = []
    inc_hlines_include = []
    inc_hlines_exclude = []

    for company, adj_inc, adj_exc, price, color in zip(
        companies, adjusted_include, adjusted_exclude, latest_prices, colors
    ):
        fin_groups_inc = filter_groups(
            adj_inc["financial"], adj_inc["subcats"], shared_price_labels
        )
        fin_groups_exc = filter_groups(
            adj_exc["financial"], adj_exc["subcats"], shared_price_labels
        )
        inc_groups_inc = filter_groups(
            adj_inc["income"], adj_inc["subcats"], shared_price_labels
        )
        inc_groups_exc = filter_groups(
            adj_exc["income"], adj_exc["subcats"], shared_price_labels
        )
        shared_divisors_inc = filter_groups(
            adj_inc["divisors"], adj_inc["subcats"], shared_price_labels
        )
        shared_divisors_exc = filter_groups(
            adj_exc["divisors"], adj_exc["subcats"], shared_price_labels
        )

        fin_include_groups.append((company.ticker, fin_groups_inc, color))
        fin_exclude_groups.append((company.ticker, fin_groups_exc, color))
        inc_include_groups.append((company.ticker, inc_groups_inc, color))
        inc_exclude_groups.append((company.ticker, inc_groups_exc, color))

        fin_hlines_include.append(
            (
                compute_normalized_latest(
                    fin_groups_inc, shared_divisors_inc, adj_inc["dates"], price
                ),
                color,
                f"{company.ticker.upper()} latest norm.",
            )
        )
        fin_hlines_exclude.append(
            (
                compute_normalized_latest(
                    fin_groups_exc, shared_divisors_exc, adj_exc["dates"], price
                ),
                color,
                f"{company.ticker.upper()} latest norm.",
            )
        )
        inc_hlines_include.append(
            (
                compute_normalized_latest(
                    inc_groups_inc, shared_divisors_inc, adj_inc["dates"], price
                ),
                color,
                f"{company.ticker.upper()} latest norm.",
            )
        )
        inc_hlines_exclude.append(
            (
                compute_normalized_latest(
                    inc_groups_exc, shared_divisors_exc, adj_exc["dates"], price
                ),
                color,
                f"{company.ticker.upper()} latest norm.",
            )
        )

    fig_fin = render_interlaced_violin(
        fin_include_groups,
        fin_exclude_groups,
        shared_price_labels,
        xlabel="Balance Sheet",
        hlines_include=fin_hlines_include,
        hlines_exclude=fin_hlines_exclude,
    )

    fig_inc = render_interlaced_violin(
        inc_include_groups,
        inc_exclude_groups,
        shared_price_labels,
        xlabel="Income Statement",
        hlines_include=inc_hlines_include,
        hlines_exclude=inc_hlines_exclude,
    )

    return FinancialViolins(fig_fin=fig_fin, fig_inc=fig_inc)

