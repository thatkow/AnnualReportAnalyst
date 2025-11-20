from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import pandas as pd


@dataclass
class Company:
    """Container for a company's combined dataset and paths."""

    ticker: str
    combined: pd.DataFrame
    company_dir: Path

    @property
    def visuals_dir(self) -> Path:
        """Return the shared visuals directory under the companies root."""
        return self.company_dir.parent / "visuals"

    @property
    def release_dates_csv(self) -> Path:
        return self.company_dir / "ReleaseDates.csv"

    @classmethod
    def from_combined(
        cls, ticker: str, combined_df: pd.DataFrame, *, companies_dir: str | Path = "companies"
    ) -> "Company":
        """Build a :class:`Company` from an in-memory combined DataFrame."""

        company_dir = Path(companies_dir) / ticker
        df = combined_df.copy()
        df = df.fillna("")
        return cls(ticker=ticker, combined=df, company_dir=company_dir)

    def default_visuals_path(self) -> Path:
        return self.visuals_dir / f"ARVisuals_{self.ticker}.html"


def import_company(
    ticker: str, *, companies_dir: str | Path = "companies", combined_filename: str = "Combined.csv"
) -> Company:
    """Load a company's Combined.csv into a :class:`Company` object."""

    company_dir = Path(companies_dir) / ticker
    combined_path = company_dir / combined_filename
    if not combined_path.exists():
        raise FileNotFoundError(f"Combined data not found for {ticker}: {combined_path}")

    df = pd.read_csv(combined_path).fillna("")
    return Company(ticker=ticker, combined=df, company_dir=company_dir)


def import_companies(
    tickers: list[str], *, companies_dir: str | Path = "companies", combined_filename: str = "Combined.csv"
) -> list[Company]:
    """Load multiple companies' Combined.csv files into :class:`Company` objects."""

    return [
        import_company(
            ticker, companies_dir=companies_dir, combined_filename=combined_filename
        )
        for ticker in tickers
    ]
