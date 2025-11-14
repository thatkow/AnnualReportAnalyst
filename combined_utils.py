import os
import pandas as pd
from pathlib import Path


def get_stock_multiplier_path(logger=None, company_dir=None, current_company_name=None):
    """Return the stock_multipliers.csv path for current company under companies/<id>"""
    try:
        # Ensure a valid company name
        company_name = (current_company_name or "").strip() or "UnknownCompany"

        # Base directory defaults to ./companies if not provided
        if company_dir is None:
            company_dir = Path.cwd() / "companies"
        else:
            company_dir = Path(company_dir)

        # Append company subdirectory
        company_path = company_dir / company_name
        company_path.mkdir(parents=True, exist_ok=True)

        fpath = company_path / "stock_multipliers.csv"
        if logger:
            logger.debug(f"📁 Using stock_multipliers.csv path: {fpath}")
        return fpath
    except Exception as e:
        if logger:
            logger.error(f"❌ Could not resolve stock_multipliers.csv path: {e}")
        raise


def ensure_stock_multiplier_file(logger=None, company_dir=None, pdf_paths=None, current_company_name=None):
    """Create the stock_multipliers.csv file if missing with default=1"""
    fpath = get_stock_multiplier_path(logger, company_dir, current_company_name)
    if not fpath.exists():
        # Extract all available dates from the provided data headers
        dates = [d for d in (pdf_paths or []) if isinstance(d, str) and d.strip()]
        df = pd.DataFrame({"Date": dates, "Stock Multiplier": [1] * len(dates)})
        df.to_csv(fpath, index=False)
        if logger:
            logger.info(f"✅ Created new stock_multipliers.csv with default 1s for {len(dates)} dates at {fpath}")
    return fpath


def load_stock_multipliers(logger=None, company_dir=None, pdf_paths=None, current_company_name=None):
    """Load stock_multipliers.csv as dict {pdf_name: multiplier}"""
    fpath = ensure_stock_multiplier_file(logger, company_dir, pdf_paths, current_company_name)
    try:
        df = pd.read_csv(fpath)
        # normalize dates (as strings)
        mults = {}
        for _, r in df.iterrows():
            key = str(r.get("Date", "")).strip()
            val = float(r.get("Stock Multiplier", 1))
            if key:
                mults[key] = val
        if logger:
            logger.info(f"🔢 Loaded {len(mults)} stock multipliers (by Date) from {fpath}")
        return mults
    except Exception as e:
        if logger:
            logger.warning(f"⚠️ Failed to read stock_multipliers.csv: {e}")
        return {}


def open_stock_multipliers_file(logger=None, company_dir=None, pdf_paths=None, current_company_name=None):
    """Open the stock_multipliers.csv file with system default editor"""
    fpath = ensure_stock_multiplier_file(logger, company_dir, pdf_paths, current_company_name)
    try:
        os.startfile(fpath)
    except Exception as e:
        if logger:
            logger.warning(f"⚠️ Could not open {fpath}: {e}")


def reload_stock_multipliers(ui_instance):
    """Reload multipliers and refresh Date Columns by PDF table"""
    mults = load_stock_multipliers(
        logger=getattr(ui_instance, "logger", None),
        company_dir=getattr(ui_instance, "company_dir", None),
        pdf_paths=getattr(ui_instance, "pdf_paths", None),
        current_company_name=getattr(ui_instance, "current_company_name", None),
    )
    ui_instance.stock_multipliers = mults
    ui_instance._populate_date_matrix_table()
    logger = getattr(ui_instance, "logger", None)
    if logger:
        logger.info("🔁 Reloaded stock multipliers into Date Columns view")
def generate_and_open_stock_multipliers(logger=None, company_dir=None, date_columns=None, current_company_name=None):
    ...
    return

def build_release_date_prompt(company: str, dates: list[str]) -> str:
    """
    Build the ReleaseDate research prompt for clipboard use.
    Produces the CSV block:
        Date,ReleaseDate
        30.06.2020,
        30.06.2021,
    And injects COMPANY into the template.
    """

    # Build CSV
    rows = ["Date,ReleaseDate"]
    for d in dates:
        if isinstance(d, str) and d.strip():
            rows.append(f"{d.strip()},")
    csv_block = "\n".join(rows)

    template = f"""
You are a highly accurate filings-research assistant.

I will supply a CSV file with a column named “Date”.
Each row represents the financial year-end date for a company.

Your task:
For each year-end date, conduct a **thorough, exhaustive, multi-source search** to determine
the **public release date** of that year’s Annual Report (or equivalent filing containing
the full audited annual financial statements).

Your search MUST include, at minimum:
• Stock exchange announcements from every exchange the company is or was listed on
• Official company investor-relations website archives
• Regulatory filings databases:
    – SEC EDGAR
    – ASX Announcements
    – HKEXnews
    – SGX Company Announcements
    – LSE RNS
    – TSX/SEDAR+
• PDF annual reports and archived versions
• MarketIndex announcement records (mandatory)
• Third-party filings repositories that store historical annual reports
• Press releases and earnings-release portals if they host annual report PDFs

Rules:
1. Identify the **earliest publicly released document** containing the FULL audited
   annual financial statements for that financial year.
2. Use only verifiable, authoritative sources (including MarketIndex).
3. If a release date cannot be **confidently and conclusively** established,
   leave the “ReleaseDate” field blank.
4. Format all dates as **DD.MM.YYYY**.
5. Do **not** hallucinate or infer dates without evidence.

Output:
• Return only a CSV with the original data and a new column “ReleaseDate”.
• No explanations, no narrative — **CSV only**.

Company: {company}

CSV:
{csv_block}
""".strip()

    return template
    """Force-create a new stock_multipliers.csv using all date columns and open it."""
    fpath = get_stock_multiplier_path(logger, company_dir, current_company_name)
    if not date_columns:
        date_columns = []
    # --- Sort date columns ascending (chronologically if possible) ---
    try:
        import re
        from datetime import datetime

        def parse_date_str(d):
            s = str(d).strip()
            # Try known formats
            for fmt in ("%d.%m.%Y", "%d-%m-%Y", "%Y-%m-%d"):
                try:
                    return datetime.strptime(s, fmt)
                except Exception:
                    pass
            # Fallback: extract numbers
            parts = re.split(r"[./-]", s)
            nums = [int(p) for p in parts if p.isdigit()]
            if len(nums) == 3:
                day, month, year = nums
                if year < 100:  # handle 2-digit years
                    year += 2000
                return datetime(year, month, day)
            return datetime.max

        date_columns = sorted(date_columns, key=parse_date_str)
        if logger:
            logger.info(f"🗓️ Sorted {len(date_columns)} date columns ascending for stock multipliers")
    except Exception as e:
        if logger:
            logger.warning(f"⚠️ Could not sort date columns: {e}")

    df = pd.DataFrame({"Date": date_columns, "Stock Multiplier": [1] * len(date_columns)})
    df.to_csv(fpath, index=False)

    if logger:
        logger.info(f"🆕 Generated new stock_multipliers.csv (sorted {len(date_columns)} dates) at {fpath}")

    try:
        os.startfile(fpath)
    except Exception as e:
        if logger:
            logger.warning(f"⚠️ Could not open generated {fpath}: {e}")
