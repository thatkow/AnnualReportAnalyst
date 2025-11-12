import pandas as pd
import yfinance as yf
from datetime import datetime, timedelta
import os
import json
import warnings


def get_stock_prices(ticker, years=5, interval="1d"):
    """
    Retrieve historical stock prices for a given ticker symbol.
    """
    df = yf.download(ticker, period=f"{years}y", interval=interval, progress=False)
    if df.empty:
        raise ValueError(f"No data found for ticker {ticker}")
    df = df.reset_index()[['Date', 'Close']].rename(columns={'Close': 'Price'})
    return df


def get_stock_data_for_dates(
    ticker: str,
    dates: list[str],
    days: list[int],
    cache_filepath: str | None = None
) -> pd.DataFrame:
    """
    Retrieve stock prices for a given ticker on each date ± specified day offsets.

    Args:
        ticker (str): Stock ticker symbol (e.g., 'BRK-B')
        dates (list[str]): List of base dates as strings in format 'DD.MM.YYYY'
        days (list[int]): List of integer day offsets (e.g., [-30, -7, 0, 7, 30])
        cache_filepath (str | None): Optional JSON cache file to reuse previously fetched data.

    Returns:
        pd.DataFrame: Columns ['BaseDate', 'OffsetDays', 'Date', 'Price']
    """
    cache = {}
    if cache_filepath and os.path.exists(cache_filepath):
        try:
            with open(cache_filepath, "r", encoding="utf-8") as f:
                cache = json.load(f)
            print(f"📂 Loaded {len(cache)} cached entries from {cache_filepath}")
        except Exception as e:
            print(f"⚠️ Could not read cache file: {e}")

    all_rows = []
    updated = False

    for d in dates:
        base_date = datetime.strptime(d, "%d.%m.%Y").date()
        for offset in days:
            target_date = base_date + timedelta(days=offset)
            key = f"{ticker}|{target_date}"

            if key in cache:
                price = cache[key]
            else:
                with warnings.catch_warnings():
                    warnings.simplefilter("ignore", category=FutureWarning)
                    df = yf.download(
                        ticker,
                        start=target_date - timedelta(days=1),
                        end=target_date + timedelta(days=1),
                        progress=False,
                        interval="1d",
                        auto_adjust=False,
                    )

                if df.empty:
                    price = None
                else:
                    df = df.reset_index()
                    df["Diff"] = (df["Date"].dt.date - target_date).abs()
                    nearest = df.loc[df["Diff"].idxmin()]
                    price = float(
                        nearest["Close"].iloc[0]
                        if hasattr(nearest["Close"], "iloc")
                        else nearest["Close"]
                    )

                # === Retry logic if no data found ===
                if price is None or pd.isna(price):
                    # === Determine search direction ===
                    # For offset == 0 (On release), respect zero_days_search_forward flag
                    search_forward = True
                    try:
                        # Lookup user preference from global or passed param
                        search_forward = zero_days_search_forward  # type: ignore[name-defined]
                    except NameError:
                        # Default fallback if not defined
                        search_forward = False

                    if offset == 0:
                        if search_forward:
                            direction = 1
                            print(f"🔁 On release date: retrying FORWARD up to 30 days for {ticker} near {target_date}")
                        else:
                            direction = -1
                            print(f"🔁 On release date: retrying BACKWARD up to 30 days for {ticker} near {target_date}")
                    else:
                        direction = 1 if offset >= 0 else -1

                    found = False
                    for retry in range(1, 31):  # search up to 30 days
                        try_date = target_date + timedelta(days=retry * direction)
                        with warnings.catch_warnings():
                            warnings.simplefilter("ignore", category=FutureWarning)
                            try_df = yf.download(
                                ticker,
                                start=try_date - timedelta(days=1),
                                end=try_date + timedelta(days=1),
                                progress=False,
                                interval="1d",
                                auto_adjust=False,
                            )

                        if not try_df.empty:
                            try_df = try_df.reset_index()
                            try_df["Diff"] = (try_df["Date"].dt.date - try_date).abs()
                            nearest = try_df.loc[try_df["Diff"].idxmin()]
                            price = float(
                                nearest["Close"].iloc[0]
                                if hasattr(nearest["Close"], "iloc")
                                else nearest["Close"]
                            )
                            found = True
                            if offset == 0:
                                direction_label = "forward" if direction == 1 else "backward"
                                print(
                                    f"✅ Found fallback 'On release' price for {ticker} ({direction_label} {retry:+}d): {price:.2f}"
                                )
                            else:
                                print(
                                    f"🔁 Found substitute price for {ticker} ({offset:+}d → {retry*direction:+}d): {price:.2f}"
                                )
                            break
                            break

                    if not found:
                        print(
                            f"⚠️ No valid price data found for {ticker} near {target_date} "
                            "after 30-day search window — marking as NA"
                        )
                        price = float("nan")

                # Cache even if NA to avoid re-querying endlessly
                cache[key] = price
                updated = True

            all_rows.append({
                "BaseDate": base_date.strftime("%d.%m.%Y"),
                "OffsetDays": offset,
                "Date": target_date.strftime("%d.%m.%Y"),
                "Price": price,
            })

    # Write updated cache back to file
    if cache_filepath and updated:
        try:
            with open(cache_filepath, "w", encoding="utf-8") as f:
                json.dump(cache, f, indent=2)
            print(f"💾 Updated cache saved to {cache_filepath} ({len(cache)} entries)")
        except Exception as e:
            print(f"⚠️ Could not write cache: {e}")

    return pd.DataFrame(all_rows)


# === Main for testing ===
if __name__ == "__main__":
    ticker = "AD8.AX"
    days = [-30, -7, 0, 7, 30]
    dates = ["31.12.2016", "31.12.2017"]

    df = get_stock_data_for_dates(ticker, dates, days, cache_filepath="stock_cache.json")
    print(df)
