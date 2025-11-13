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

        # Define one large window: start = base - min(days) - 30, end = base + max(days) + 30
        start_date = base_date - timedelta(days=(abs(min(days)) + 30))
        end_date = base_date + timedelta(days=(abs(max(days)) + 30))

        with warnings.catch_warnings():
            warnings.simplefilter("ignore", category=FutureWarning)
            df_full = yf.download(
                ticker,
                start=start_date,
                end=end_date,
                progress=False,
                interval="1d",
                auto_adjust=False,
            )

        if df_full.empty:
            print(f"⚠️ No data found for {ticker} in {start_date}–{end_date}. Marking all as NA.")
            df_full = pd.DataFrame(columns=["Date", "Close"])
        else:
            df_full = df_full.reset_index()

            # Flatten multi-level columns if present
            if isinstance(df_full.columns, pd.MultiIndex):
                df_full.columns = [' '.join(c).strip() for c in df_full.columns.values]

            # Try to find a 'Close' column that includes the ticker
            close_cols = [c for c in df_full.columns if "Close" in c]
            if not close_cols:
                print(f"⚠️ Could not find a Close column for {ticker}. Columns: {df_full.columns.tolist()}")
                df_full["Close"] = float("nan")
            else:
                df_full["Close"] = df_full[close_cols[0]]

            df_full["Date"] = pd.to_datetime(df_full["Date"], errors="coerce").dt.date
            df_full["Close"] = pd.to_numeric(df_full["Close"], errors="coerce")

        # Drop invalid rows
        df_full = df_full.dropna(subset=["Date", "Close"])

        print(f"DEBUG: Cleaned df_full shape {df_full.shape}")
        print(df_full.head(5))

        # Build lookup
        price_lookup = {d_: float(v) for d_, v in zip(df_full["Date"], df_full["Close"])}


        # Drop invalid rows
        df_full = df_full.dropna(subset=["Date", "Close"])

        print(f"DEBUG: Cleaned df_full shape {df_full.shape}")
        print(df_full.head(5))

        # Build lookup
        price_lookup = {d_: float(v) for d_, v in zip(df_full["Date"], df_full["Close"])}


        # Drop any invalid rows (non-numeric Close or bad Date)
        df_full = df_full.dropna(subset=["Date", "Close"])

        # Build lookup dict
        try:
            price_lookup = {d_: float(v) for d_, v in zip(df_full["Date"], df_full["Close"])}
        except Exception as e:
            print(f"⚠️ Failed to build price lookup for {ticker}: {e}")
            price_lookup = {}

        for offset in days:
            target_date = base_date + timedelta(days=offset)
            key = f"{ticker}|{target_date}"

            if key in cache:
                price = cache[key]
            else:
                price = price_lookup.get(target_date, float("nan"))

                # If not found, iterate forward (if +offset) or backward (if -offset) to find next available price
                if pd.isna(price):
                    direction = 1 if offset > 0 else -1
                    found = False
                    for retry in range(1, 31):  # search within ±30 days
                        new_date = target_date + timedelta(days=retry * direction)
                        if new_date in price_lookup:
                            price = price_lookup[new_date]
                            found = True
                            print(
                                f"🔁 Found substitute for {ticker} on {target_date} "
                                f"→ {new_date} ({'forward' if direction==1 else 'backward'} {retry:+}d): {price:.2f}"
                            )
                            break
                    if not found:
                        print(
                            f"⚠️ No valid stock price found for {ticker} near {target_date} "
                            f"after 30-day search window — marking as NA."
                        )
                        price = float("nan")

                cache[key] = price

                # Only mark updated and cache if valid price exists
                if not pd.isna(price):
                    updated = True
                else:
                    # Do not store NaN in cache to allow future re-checks
                    cache.pop(key, None)

            if pd.isna(price):
                print(
                    f"⚠️ No price found for {ticker} on {target_date}, "
                    f"even after directional lookup — marking as NA."
                )
            else:
                print(
                    f"✅ Price for {ticker} on {target_date}: {price:.2f} "
                    f"(offset {offset:+}d)"
                )

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
    dates = ["30.10.2017"]

    df = get_stock_data_for_dates(ticker, dates, days, cache_filepath="stock_cache.json")
    print(df)
