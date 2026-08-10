"""
config.py
Org configuration, data folder location, and month/file detection helpers.
"""

import json
import re
from dataclasses import asdict, dataclass
from pathlib import Path

import platformdirs

APP_NAME = 'pta-treasurer'

FISCAL_START_MONTH = 7   # July

MONTH_NAMES = {
    'january': 0, 'february': 1, 'march': 2, 'april': 3, 'may': 4, 'june': 5,
    'july': 6, 'august': 7, 'september': 8, 'october': 9, 'november': 10, 'december': 11,
    'jan': 0, 'feb': 1, 'mar': 2, 'apr': 3, 'jun': 5,
    'jul': 6, 'aug': 7, 'sep': 8, 'oct': 9, 'nov': 10, 'dec': 11,
}


@dataclass
class OrgConfig:
    org_name: str
    fiscal_start_month: int = FISCAL_START_MONTH
    balance_forward: float = 0.0
    # Givebacks auto-download (opt-in, Phase 4). Non-secret identifiers
    # only -- the password is never stored here, see credentials.py.
    givebacks_org_url: str = ''
    givebacks_email: str = ''
    givebacks_cause_id: str = ''
    # AI Assistant (opt-in). Runs against a local Ollama server -- no API
    # key/secret involved, so both fields live here like everything else.
    ollama_host: str = 'http://localhost:11434'
    ollama_model: str = 'llama3.2'
    # Pass-through fund (opt-in, e.g. a fundraiser whose money moves through
    # the org's bank account but isn't the org's own money) -- excluded from
    # Income/Expenses/Budget/YTD entirely, tracked separately. Blank name =
    # feature off. One fund for now, not a list -- see PLAN.md.
    pass_through_fund_name: str = ''
    pass_through_fund_categories: str = ''  # comma-separated QuickBooks category names
    pass_through_fund_balance_forward: float = 0.0


# ── Fiscal month math ──────────────────────────────────────────────────────

def calendar_to_fiscal(cal_month_num: int) -> int:
    """Converts a 1-12 calendar month number to a 0-11 fiscal-year index (0 = July)."""
    return (cal_month_num - FISCAL_START_MONTH) % 12


def month_name_to_fiscal_index(month_name: str) -> int | None:
    cal_idx = MONTH_NAMES.get(month_name.lower())
    if cal_idx is None:
        return None
    return calendar_to_fiscal(cal_idx + 1)


def fiscal_year_start_calendar_year(month_label: str) -> int:
    """The calendar year the fiscal year containing `month_label` (e.g.
    'January 2026') started in -- e.g. 2025, since a July-start fiscal
    year covering Jan 2026 began in July 2025."""
    month_str, year_str = month_label.split()
    cal_idx = MONTH_NAMES.get(month_str.lower())
    if cal_idx is None:
        raise ValueError(f'Unrecognized month name: {month_str!r}')
    year = int(year_str)
    cal_month_num = cal_idx + 1
    return year if cal_month_num >= FISCAL_START_MONTH else year - 1


# ── Month detection from filenames / PDF content ───────────────────────────

def detect_month_from_filename(filepath: Path):
    """
    Detects a fiscal month from a filename like 'quickbooks_july_2025.csv'.
    Returns (month_label, fiscal_index) or (None, None) if no month name found.
    """
    name = filepath.stem.lower()
    parts = re.split(r'[_\-\s]+', name)
    month_str = None
    year_str = None
    for part in parts:
        if part in MONTH_NAMES and month_str is None:
            month_str = part
        if re.match(r'^\d{4}$', part) and year_str is None:
            year_str = part
    if month_str is None:
        return None, None
    fiscal_idx = month_name_to_fiscal_index(month_str)
    month_label = f'{month_str.capitalize()} {year_str}' if year_str else month_str.capitalize()
    return month_label, fiscal_idx


def detect_month_from_pdf(pdf_path: Path):
    """
    Detects a fiscal month for a bank statement PDF: filename first,
    falling back to the statement period printed in the PDF text.
    Returns (month_label, fiscal_index) or (None, None).
    """
    lbl, idx = detect_month_from_filename(pdf_path)
    if lbl:
        return lbl, idx

    import pdfplumber
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = pdf.pages[0].extract_text() or ''
    except Exception:
        return None, None

    m = re.search(
        r'(January|February|March|April|May|June|July|August|'
        r'September|October|November|December)\s+\d{1,2},\s+(\d{4})'
        r'\s*through', text, re.IGNORECASE)
    if m:
        return f'{m.group(1)} {m.group(2)}', month_name_to_fiscal_index(m.group(1))
    return None, None


# ── Data folder location (platformdirs pointer file) ───────────────────────

def _pointer_path() -> Path:
    return Path(platformdirs.user_config_dir(APP_NAME)) / 'data_dir.json'


def get_data_dir() -> Path | None:
    """Returns the user's chosen data folder, or None on first run."""
    p = _pointer_path()
    if not p.exists():
        return None
    try:
        data = json.loads(p.read_text())
    except (json.JSONDecodeError, OSError):
        return None
    path = data.get('data_dir')
    return Path(path) if path else None


def set_data_dir(path: Path) -> None:
    """Remembers the user's chosen data folder for future runs."""
    p = _pointer_path()
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps({'data_dir': str(path)}, indent=2))


# ── Org config, stored inside the data folder itself ───────────────────────

def _config_path(data_dir: Path) -> Path:
    return Path(data_dir) / 'org_config.json'


def load_config(data_dir: Path) -> OrgConfig:
    data = json.loads(_config_path(data_dir).read_text())
    return OrgConfig(**data)


def save_config(config: OrgConfig, data_dir: Path) -> None:
    p = _config_path(data_dir)
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps(asdict(config), indent=2))
