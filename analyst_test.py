# === test_stacked_annual_report.py ===
import pandas as pd
import numpy as np
from analyst_stackedvisuals import render_stacked_annual_report


def generate_test_data():
    """Generate synthetic test data compatible with the new auto-inferring plotter."""
    types = ["A", "B"]
    tickers = ["X", "Y"]
    records = []

    # Create mock data rows for each TYPE and Ticker
    for typ in types:
        for ticker in tickers:
            for i in range(3):
                records.append({
                    "TYPE": typ,
                    "Ticker": ticker,
                    "CATEGORY": f"Cat{i+1}",
                    "SUBCATEGORY": f"Sub{i+1}",
                    "ITEM": f"Item_{typ}{ticker}{i+1}",
                    "NOTE": ""
                })

    df = pd.DataFrame(records)

    # Generate date-formatted year columns (DD.MM.YYYY)
    years = [f"31.12.{year}" for year in range(2021, 2025)]
    for y in years:
        df[y] = np.random.randint(-8, 12, len(df))

    # Add one share_count row per ticker
    for ticker in tickers:
        df = pd.concat([
            df,
            pd.DataFrame([{
                "TYPE": "Shares",
                "Ticker": ticker,
                "CATEGORY": "",
                "SUBCATEGORY": "",
                "ITEM": "NUMBER OF SHARES",
                "NOTE": "share_count",
                **{y: np.random.randint(1000, 2000) for y in years}
            }])
        ], ignore_index=True)

    return df


if __name__ == "__main__":
    df = generate_test_data()
    render_stacked_annual_report(df, title="Test Annual Report (Auto-Infer)")
