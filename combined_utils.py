import os
import pandas as pd
from pathlib import Path
from typing import Iterable, Optional


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



def reload_stock_multipliers(ui_instance):
    """
    Reload multipliers and refresh any dependent UI.

    This is wired to work with the CombinedUIMixin-based UI:
      - uses ui_instance.companies_dir
      - uses ui_instance.company_var.get() as the company name
      - uses ui_instance.combined_rename_names / combined_dyn_columns as date list
    """
    logger = getattr(ui_instance, "logger", None)

    # Resolve company directory (CombinedUIMixin uses 'companies_dir')
    company_dir = getattr(ui_instance, "companies_dir", None)

    # Resolve current company name
    current_company_name: Optional[str] = None
    try:
        # CombinedUIMixin style
        if hasattr(ui_instance, "company_var") and ui_instance.company_var is not None:
            current_company_name = ui_instance.company_var.get()
    except Exception:
        current_company_name = None

    # Fallback to any legacy attribute if present
    if not current_company_name:
        current_company_name = getattr(ui_instance, "current_company_name", None)

    # Try to get the date columns used for multipliers (from Combined tab)
    date_cols: Optional[Iterable[str]] = None
    if hasattr(ui_instance, "combined_rename_names"):
        date_cols = getattr(ui_instance, "combined_rename_names", None)

    if (not date_cols) and hasattr(ui_instance, "combined_dyn_columns"):
        try:
            dyn_cols = getattr(ui_instance, "combined_dyn_columns", [])
            date_cols = [dc.get("default_name") for dc in dyn_cols if isinstance(dc, dict)]
        except Exception:
            date_cols = None

    mults = load_stock_multipliers(
        logger=logger,
        company_dir=company_dir,
        pdf_paths=date_cols,
        current_company_name=current_company_name,
    )
    ui_instance.stock_multipliers = mults

    # Best-effort UI refresh (support both old and new APIs)
    if hasattr(ui_instance, "_populate_date_matrix_table"):
        try:
            ui_instance._populate_date_matrix_table()
        except Exception:
            pass
    elif hasattr(ui_instance, "refresh_combined_tab"):
        try:
            ui_instance.refresh_combined_tab()
        except Exception:
            pass

    if logger:
        logger.info("🔁 Reloaded stock multipliers")
def generate_and_open_stock_multipliers(logger=None, company_dir=None, date_columns=None, current_company_name=None):
    """
    Create or extend companies/<company>/stock_multipliers.csv :
      • If file exists: append missing dates (Stock Multiplier = 1)
      • If absent: create new file
      • Always open the file afterwards
    """
    import csv, os
    from datetime import datetime

    # ------------------------------------------------------------------
    # ALWAYS resolve company stock_multipliers.csv path correctly
    # => companies/<company>/stock_multipliers.csv
    # ------------------------------------------------------------------
    fpath = get_stock_multiplier_path(
        logger=logger,
        company_dir=company_dir,
        current_company_name=current_company_name
    )

    # Normalize input
    if not date_columns:
        date_columns = []

    # Try to sort dates in ascending chronological order
    try:
        import re
        def parse_date_str(s):
            s = str(s).strip()
            # Known formats
            for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d-%m-%Y"):
                try:
                    return datetime.strptime(s, fmt)
                except Exception:
                    pass
            # Fallback
            parts = re.split(r"[./-]", s)
            nums = [int(p) for p in parts if p.isdigit()]
            if len(nums) == 3:
                d,m,y = nums
                if y < 100:
                    y += 2000
                return datetime(y,m,d)
            return datetime.max

        date_columns = sorted(date_columns, key=parse_date_str)
    except Exception as e:
        if logger:
            logger.warning(f"⚠️ Failed to sort dates: {e}")

    # ------------------------------------------------------------------
    # CASE A: stock_multipliers.csv EXISTS → merge missing dates
    # ------------------------------------------------------------------
    if fpath.exists():
        existing = {}
        try:
            with fpath.open("r", encoding="utf-8", newline="") as fh:
                reader = csv.DictReader(fh)
                for row in reader:
                    d = row.get("Date", "").strip()
                    mv = row.get("Stock Multiplier", "").strip()
                    if d:
                        existing[d] = mv if mv != "" else "1"
        except Exception as e:
            if logger:
                logger.error(f"❌ Failed to read {fpath}: {e}")
            existing = {}

        changed = False
        for d in date_columns:
            if d not in existing:
                existing[d] = "1"
                changed = True

        if changed:
            try:
                with fpath.open("w", encoding="utf-8", newline="") as fh:
                    writer = csv.writer(fh)
                    writer.writerow(["Date", "Stock Multiplier"])
                    for d in date_columns:
                        writer.writerow([d, existing.get(d, "1")])
            except Exception as e:
                if logger:
                    logger.error(f"❌ Unable to rewrite merged multipliers: {e}")

    else:
        # ------------------------------------------------------------------
        # CASE B: File does NOT exist → create new stock_multipliers.csv
        # ------------------------------------------------------------------
        try:
            df = pd.DataFrame({"Date": date_columns,
                               "Stock Multiplier": [1]*len(date_columns)})
            df.to_csv(fpath, index=False)
        except Exception as e:
            if logger:
                logger.error(f"❌ Failed to create stock_multipliers.csv: {e}")
            return

    # ------------------------------------------------------------------
    # Open the file (Windows/macOS/Linux)
    # ------------------------------------------------------------------
    try:
        os.startfile(str(fpath))
    except Exception as e:
        if logger:
            logger.warning(f"⚠️ Could not open {fpath}: {e}")

    return

def _sort_dates(dates: list[str]) -> list[str]:
    """Return unique date strings sorted chronologically when possible."""

    unique: list[str] = []
    for d in dates:
        s = str(d).strip()
        if not s:
            continue
        if s not in unique:
            unique.append(s)

    if not unique:
        return []

    try:
        import re
        from datetime import datetime

        def parse_date(value: str) -> datetime:
            for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d-%m-%Y"):
                try:
                    return datetime.strptime(value, fmt)
                except Exception:
                    continue
            parts = re.split(r"[./-]", value)
            nums = [int(p) for p in parts if p.isdigit()]
            if len(nums) == 3:
                day, month, year = nums
                if year < 100:
                    year += 2000
                return datetime(year, month, day)
            return datetime.max

        return sorted(unique, key=parse_date)
    except Exception:
        return sorted(unique)


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
    for d in _sort_dates(dates):
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
the full audited annual financial statements). If you cannot confidently find the release
date for the most recent year-end, assume it is a quarterly report and instead determine
the release date for the latest quarterly filing.

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


def build_stock_multiplier_prompt(
    company: str,
    dates: list[str],
    template_path: Optional[Path] = None,
) -> str:
    """Fill the Stock Multiplier template with the company and date table."""

    if template_path is None:
        template_path = Path(__file__).resolve().parent / "prompts" / "Stock_Multipliers.txt"
    else:
        template_path = Path(template_path)

    if not template_path.exists():
        raise FileNotFoundError(f"Stock multiplier prompt template not found: {template_path}")

    template_text = template_path.read_text(encoding="utf-8").strip()

    sorted_dates = _sort_dates(dates)
    table_lines = ["Date\tStock Multiplier"]
    for d in sorted_dates:
        table_lines.append(f"{d}\t")
    if len(table_lines) == 1:
        table_lines.append("<add at least one Date column>\t")
    table_block = "\n".join(table_lines)

    replacements = {
        "{{COMPANY_NAME}}": (company or "<Company>").strip() or "<Company>",
        "{{DATE_TABLE}}": table_block,
    }
    for placeholder, value in replacements.items():
        template_text = template_text.replace(placeholder, value)

    return template_text
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
