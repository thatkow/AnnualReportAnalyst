from dataclasses import dataclass
from typing import Iterable, Sequence

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.colors import to_rgba
from matplotlib.patches import Patch
from matplotlib.ticker import MultipleLocator

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

def _fade_color(color: str | tuple, *, alpha: float = 0.35) -> tuple[float, float, float, float]:
    base = to_rgba(color)
    return (base[0], base[1], base[2], alpha)


def _darken_color(
    color: str | tuple, *, factor: float = 0.5, alpha_boost: float = 2.0
) -> tuple[float, float, float, float]:
    """Darken a color towards black while optionally boosting opacity."""

    base = to_rgba(color)
    r, g, b, a = base
    return (
        max(0.0, r * factor),
        max(0.0, g * factor),
        max(0.0, b * factor),
        min(1.0, a * alpha_boost),
    )


def _groups_equal(left: Sequence[float], right: Sequence[float]) -> bool:
    """Return ``True`` when two numeric sequences match elementwise, including NaNs."""

    left_arr = np.asarray(left, dtype=float)
    right_arr = np.asarray(right, dtype=float)

    if left_arr.shape != right_arr.shape:
        return False

    return np.allclose(left_arr, right_arr, equal_nan=True)


def render_interlaced_boxplots(
    ticker_groups: Sequence[tuple[str, Sequence[Sequence[float]], str]],
    price_labels,
    *,
    xlabel: str,
    hlines: list[tuple[float, str, str]] | None = None,
    exclude_ticker_groups: Sequence[tuple[str, Sequence[Sequence[float]], str]] | None = None,
    hline_lookup: dict[str, dict[str, float | None]] | None = None,
    height: float = 6.0,
):
    """Render interlaced boxplots for ``n`` tickers.

    ``ticker_groups`` contains tuples of ``(ticker, groups, color)`` where
    ``groups`` is already filtered to ``price_labels`` order. When
    ``exclude_ticker_groups`` is provided, each ticker's include/exclude groups
    are plotted side-by-side with a faded color for the exclude variant.
    """

    inter_groups: list[Sequence[float]] = []
    inter_colors: list[str | tuple[float, float, float, float]] = []
    inter_labels: list[str] = []
    positions: list[float] = []
    box_meta: list[tuple[str, str]] = []  # (ticker, variant "include"/"exclude")

    exclude_usage: dict[str, bool] = {}

    if exclude_ticker_groups:
        exclude_usage = {ticker: False for ticker, _, _ in exclude_ticker_groups}
        pair_index = 0
        for price_idx, label in enumerate(price_labels):
            for (ticker_inc, groups_inc, color_inc), (
                ticker_exc,
                groups_exc,
                _color_exc,
            ) in zip(ticker_groups, exclude_ticker_groups):
                if ticker_inc != ticker_exc:
                    raise ValueError(
                        "Ticker ordering must match between include/exclude groups"
                    )
                if price_idx >= len(groups_inc) or price_idx >= len(groups_exc):
                    raise ValueError("Price labels must align with all group entries")

                center = pair_index + 1

                if _groups_equal(groups_inc[price_idx], groups_exc[price_idx]):
                    inter_groups.append(groups_inc[price_idx])
                    inter_colors.append(color_inc)
                    inter_labels.append(label)
                    positions.append(center)
                    box_meta.append((ticker_inc, "include"))
                else:
                    inc_pos = center - 0.18
                    exc_pos = center + 0.18

                    inter_groups.extend([groups_inc[price_idx], groups_exc[price_idx]])
                    inter_colors.extend([color_inc, _fade_color(color_inc)])
                    inter_labels.extend([label, label])
                    positions.extend([inc_pos, exc_pos])
                    box_meta.extend([(ticker_inc, "include"), (ticker_inc, "exclude")])
                    exclude_usage[ticker_inc] = True

                pair_index += 1
    else:
        for i in range(len(price_labels)):
            for ticker, groups, color in ticker_groups:
                inter_groups.append(groups[i])
                inter_colors.append(color)
                inter_labels.append(price_labels[i])
                positions.append(len(positions) + 1)
                box_meta.append((ticker, "include"))

    fig, ax = plt.subplots(figsize=(16, height))
    bp = ax.boxplot(
        inter_groups,
        labels=inter_labels,
        patch_artist=True,
        positions=positions,
        widths=0.25,
        showmeans=True,
        showfliers=False,
        meanline=False,
        meanprops={
            "marker": "x",
            "markeredgecolor": "black",
            "markerfacecolor": "black",
        },
    )

    for patch, color in zip(bp["boxes"], inter_colors):
        patch.set_facecolor(color)

    for median in bp.get("medians", []):
        median.set_color("black")
    for mean in bp.get("means", []):
        mean.set_color("black")

    if hline_lookup:
        y_min, y_max = ax.get_ylim()
        line_values = [
            v
            for ticker_lookup in hline_lookup.values()
            for v in ticker_lookup.values()
            if v is not None and not np.isnan(v)
        ]
        if line_values:
            y_max = max(y_max, max(line_values))
            y_min = min(y_min, min(line_values))

        y_pad = 0.12 * (y_max - y_min)
        ax.set_ylim(y_min, y_max + y_pad)
        y_text = y_max + y_pad - 0.02 * (y_max - y_min)

        for pos, data, (ticker, variant), color in zip(
            positions, inter_groups, box_meta, inter_colors
        ):
            hline_value = hline_lookup.get(ticker, {}).get(variant)
            data_arr = np.asarray(data, dtype=float)
            mean_val = np.nanmean(data_arr)
            median_val = np.nanmedian(data_arr)
            n_points = int(np.count_nonzero(~np.isnan(data_arr)))

            def _pct(hline, base):
                if (
                    hline is None
                    or base is None
                    or np.isnan(hline)
                    or np.isnan(base)
                    or base == 0
                ):
                    return "—"
                return f"{(hline / base) * 100:.0f}%"

            mean_pct = _pct(hline_value, mean_val)
            median_pct = _pct(hline_value, median_val)
            ticker_label = ticker.upper()
            includes_intangibles = exclude_ticker_groups is not None and variant == "include"
            if includes_intangibles:
                ticker_label = f"{ticker_label} (I)"

            ax.text(
                pos,
                y_text,
                f"{ticker_label}\n{mean_pct}\n{median_pct}\nn={n_points}",
                ha="center",
                va="top",
                fontsize=8,
                color=_darken_color(color),
            )

            if hline_value is not None and not np.isnan(hline_value):
                ax.scatter(
                    pos,
                    hline_value,
                    color=_darken_color(color),
                    edgecolor="black",
                    marker="*",
                    zorder=5,
                    s=150,
                )

                ax.text(
                    pos + 0.08,
                    hline_value,
                    f"{hline_value:.2f}",
                    ha="left",
                    va="center",
                    fontsize=8,
                    color=_darken_color(color),
                )

    ax.set_xlabel(xlabel)
    plt.xticks(rotation=45, ha="right")
    plt.title(xlabel)
    ax.yaxis.set_major_locator(MultipleLocator(0.1))
    ax.yaxis.set_minor_locator(MultipleLocator(0.5))
    ax.grid(which="major", axis="y", alpha=0.6)
    ax.grid(which="minor", axis="y", alpha=0.3)
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
    """Render split interlaced violins for intangible on/off views in one plot.

    Each price label appears once with a left-half violin for
    ``include_intangibles=True`` and a right-half violin for
    ``include_intangibles=False``. Horizontal reference lines can be specified
    independently for each half.
    """

    paired_groups: list[tuple[Sequence[float], Sequence[float], str, str]] = []

    for price_idx, label in enumerate(price_labels):
        for (ticker_inc, groups_inc, color_inc), (
            ticker_exc,
            groups_exc,
            _color_exc,
        ) in zip(ticker_groups_include, ticker_groups_exclude):
            if ticker_inc != ticker_exc:
                raise ValueError(
                    "Ticker ordering must match between include/exclude groups"
                )

            if price_idx >= len(groups_inc) or price_idx >= len(groups_exc):
                raise ValueError("Price labels must align with all group entries")

            paired_groups.append((groups_inc[price_idx], groups_exc[price_idx], color_inc, label))

    fig, ax = plt.subplots(figsize=(18, 6))
    positions = np.arange(1, len(paired_groups) + 1)

    def _half_violin(data, position, color, *, side: str):
        vp = ax.violinplot(
            [data], positions=[position], showmeans=True, showextrema=False, widths=0.85
        )

        for body in vp["bodies"]:
            verts = body.get_paths()[0].vertices
            if side == "left":
                verts[:, 0] = position - np.abs(verts[:, 0] - position)
            else:
                verts[:, 0] = position + np.abs(verts[:, 0] - position)
            body.set_facecolor(color)
            body.set_edgecolor("black")
            body.set_alpha(0.6)

        return vp

    for pos, (inc_group, exc_group, color, label) in zip(positions, paired_groups):
        _half_violin(inc_group, pos, color, side="left")
        _half_violin(exc_group, pos, color, side="right")

    if hlines_include:
        for value, color, label in hlines_include:
            if value is None or np.isnan(value):
                continue
            ax.axhline(value, color=color, linestyle="--", linewidth=1.5, label=label)

    if hlines_exclude:
        for value, color, label in hlines_exclude:
            if value is None or np.isnan(value):
                continue
            ax.axhline(value, color=color, linestyle="-.", linewidth=1.5, label=label)

    ax.set_xticks(positions)
    ax.set_xticklabels([grp_label for _, _, _, grp_label in paired_groups], rotation=45, ha="right")
    ax.set_xlabel(xlabel)
    ax.set_title("Interlaced Violins — include/exclude intangibles")

    legend_handles = [
        Patch(color=color, label=ticker.upper())
        for ticker, _, color in ticker_groups_include
    ]

    line_handles, line_labels = ax.get_legend_handles_labels()
    combined_handles: list[Patch] = []
    combined_labels: list[str] = []

    for handle, label in list(zip(legend_handles, [h.get_label() for h in legend_handles])) + list(
        zip(line_handles, line_labels)
    ):
        if label in combined_labels:
            continue
        combined_handles.append(handle)
        combined_labels.append(label)

    ax.legend(handles=combined_handles, labels=combined_labels, loc="upper left")
    fig.tight_layout()

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
    companies: Sequence[Company], *, include_intangibles: bool = True, price_labels: Sequence[str | int] | None = None, height: float = 6.0
) -> FinancialBoxplots:
    """Generate interlaced boxplots for all provided companies.

    When ``include_intangibles`` is ``True`` both include/exclude intangibles
    series are plotted side-by-side with a faded style for the exclude variant.

    Args:
        companies: Companies to plot.
        include_intangibles: Whether to display the include/exclude intangibles views.
        price_labels: Specific price columns to include. When ``None`` (default),
            all shared price labels across the given companies are plotted.
        height: Height of the generated boxplots in inches.
    """

    if not companies:
        raise ValueError("At least one company is required to build boxplots.")

    ensure_interactive_backend()

    # Precompute adjustments and latest prices
    adjusted_include = [
        compute_adjusted_values(company.ticker, company.combined, include_intangibles=True)
        for company in companies
    ]
    adjusted_exclude = [
        compute_adjusted_values(company.ticker, company.combined, include_intangibles=False)
        for company in companies
    ]
    latest_prices = [get_latest_stock_price(company.ticker) for company in companies]

    # Normalise subcats so strings '1' and '1.0' match
    def normalise(label):
        try:
            return str(int(float(label)))   # "1.0" -> 1 -> "1"
        except Exception:
            return str(label)

    for adjusted in (*adjusted_include, *adjusted_exclude):
        adjusted["subcats"] = [normalise(s) for s in adjusted["subcats"]]

    def _shared_labels(adjusted_list):
        first_labels = adjusted_list[0]["subcats"]
        return [
            label
            for label in first_labels
            if all(label in adjusted["subcats"] for adjusted in adjusted_list[1:])
        ]

    shared_inc = _shared_labels(adjusted_include)
    shared_exc = _shared_labels(adjusted_exclude)

    shared_available = (
        [label for label in shared_inc if label in shared_exc]
        if include_intangibles
        else shared_exc
    )
    if not shared_available:
        raise ValueError("No shared stock price labels found between companies")

    requested_labels = shared_available if price_labels is None else price_labels
    requested_labels = [normalise(label) for label in requested_labels]

    missing = [label for label in requested_labels if label not in shared_available]
    if missing:
        raise ValueError(
            f"Requested price_labels not found in data: {', '.join(map(str, missing))}"
        )

    shared_price_labels = requested_labels

    def filter_groups(groups, available_labels, target_labels):
        index_lookup = {label: i for i, label in enumerate(available_labels)}
        return [groups[index_lookup[label]] for label in target_labels]

    def omit_latest_date(groups: Sequence[Sequence[float]]):
        trimmed: list[np.ndarray] = []
        for group in groups:
            arr = np.asarray(group)
            trimmed.append(arr[:-1])
        return trimmed

    colors = [plt.cm.tab10(i % 10) for i in range(len(companies))]

    fin_ticker_groups = []
    inc_ticker_groups = []
    fin_ticker_groups_exclude = []
    inc_ticker_groups_exclude = []
    fin_hlines = []
    inc_hlines = []
    fin_line_lookup: dict[str, dict[str, float | None]] = {}
    inc_line_lookup: dict[str, dict[str, float | None]] = {}

    for company, adj_inc, adj_exc, price, color in zip(
        companies, adjusted_include, adjusted_exclude, latest_prices, colors
    ):
        fin_groups_inc = filter_groups(
            adj_inc["financial"], adj_inc["subcats"], shared_price_labels
        )
        inc_groups_inc = filter_groups(
            adj_inc["income"], adj_inc["subcats"], shared_price_labels
        )
        shared_divisors_inc = filter_groups(
            adj_inc["divisors"], adj_inc["subcats"], shared_price_labels
        )

        fin_groups_exc = filter_groups(
            adj_exc["financial"], adj_exc["subcats"], shared_price_labels
        )
        inc_groups_exc = filter_groups(
            adj_exc["income"], adj_exc["subcats"], shared_price_labels
        )
        shared_divisors_exc = filter_groups(
            adj_exc["divisors"], adj_exc["subcats"], shared_price_labels
        )

        fin_ticker_groups.append(
            (company.ticker, omit_latest_date(fin_groups_inc), color)
        )
        inc_ticker_groups.append(
            (company.ticker, omit_latest_date(inc_groups_inc), color)
        )
        fin_ticker_groups_exclude.append(
            (company.ticker, omit_latest_date(fin_groups_exc), color)
        )
        inc_ticker_groups_exclude.append(
            (company.ticker, omit_latest_date(inc_groups_exc), color)
        )

        base_fin_groups = fin_groups_inc if include_intangibles else fin_groups_exc
        base_inc_groups = inc_groups_inc if include_intangibles else inc_groups_exc
        base_fin_divisors = (
            shared_divisors_inc if include_intangibles else shared_divisors_exc
        )
        base_inc_divisors = (
            shared_divisors_inc if include_intangibles else shared_divisors_exc
        )
        base_dates = adj_inc["dates"] if include_intangibles else adj_exc["dates"]

        fin_line_inc = compute_normalized_latest(
            base_fin_groups, base_fin_divisors, base_dates, price
        )
        inc_line_inc = compute_normalized_latest(
            base_inc_groups, base_inc_divisors, base_dates, price
        )

        fin_hlines.append((fin_line_inc, color, f"{company.ticker.upper()} latest"))
        inc_hlines.append((inc_line_inc, color, f"{company.ticker.upper()} latest"))
        fin_line_lookup[company.ticker] = {"include": fin_line_inc}
        inc_line_lookup[company.ticker] = {"include": inc_line_inc}

        if include_intangibles:
            fin_line_exc = compute_normalized_latest(
                fin_groups_exc, shared_divisors_exc, adj_exc["dates"], price
            )
            inc_line_exc = compute_normalized_latest(
                inc_groups_exc, shared_divisors_exc, adj_exc["dates"], price
            )

            fin_hlines.append(
                (
                    fin_line_exc,
                    _fade_color(color),
                    f"{company.ticker.upper()} latest (ex intg)",
                )
            )
            inc_hlines.append(
                (
                    inc_line_exc,
                    _fade_color(color),
                    f"{company.ticker.upper()} latest (ex intg)",
                )
            )
            fin_line_lookup[company.ticker]["exclude"] = fin_line_exc
            inc_line_lookup[company.ticker]["exclude"] = inc_line_exc

    fig_fin = render_interlaced_boxplots(
        fin_ticker_groups,
        shared_price_labels,
        xlabel="Balance Sheet",
        hlines=fin_hlines,
        exclude_ticker_groups=fin_ticker_groups_exclude if include_intangibles else None,
        hline_lookup=fin_line_lookup,
        height=height,
    )

    fig_inc = render_interlaced_boxplots(
        inc_ticker_groups,
        shared_price_labels,
        xlabel="Income Statement",
        hlines=inc_hlines,
        exclude_ticker_groups=inc_ticker_groups_exclude if include_intangibles else None,
        hline_lookup=inc_line_lookup,
        height=height,
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
                f"{company.ticker.upper()} latest",
            )
        )
        fin_hlines_exclude.append(
            (
                compute_normalized_latest(
                    fin_groups_exc, shared_divisors_exc, adj_exc["dates"], price
                ),
                color,
                f"{company.ticker.upper()} latest (ex intg)",
            )
        )
        inc_hlines_include.append(
            (
                compute_normalized_latest(
                    inc_groups_inc, shared_divisors_inc, adj_inc["dates"], price
                ),
                color,
                f"{company.ticker.upper()} latest",
            )
        )
        inc_hlines_exclude.append(
            (
                compute_normalized_latest(
                    inc_groups_exc, shared_divisors_exc, adj_exc["dates"], price
                ),
                color,
                f"{company.ticker.upper()} latest (ex intg)",
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
