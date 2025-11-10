import yfinance as yf
import pandas as pd
from datetime import date, timedelta

def get_stock_prices(ticker: str, years: int = 10, interval: str = "1d") -> pd.DataFrame:
    """
    Fetch historical stock prices from Yahoo Finance and return Date and Price (Close).

    Args:
        ticker (str): Stock symbol (e.g., "AAPL", "BRK-B", "ASX:WOW").
        years (int): Number of years of history to retrieve.
        interval (str): Data interval ("1d", "1wk", "1mo").

    Returns:
        pd.DataFrame: DataFrame with columns ['Date', 'Price'].
    """
    end = date.today()
    start = end - timedelta(days=years * 365)

    df = yf.download(ticker, start=start, end=end, interval=interval, progress=False)
    if df.empty:
        raise ValueError(f"No data returned for {ticker}")

    df = df.reset_index()[['Date', 'Close']].rename(columns={'Close': 'Price'})
    return df

# Example usage:
if __name__ == "__main__":
    data = get_stock_prices("BRK-B", years=10, interval="1wk")
    print(data["Price"].head())
    print(data.head())
