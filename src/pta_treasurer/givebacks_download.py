"""
givebacks_download.py
Opt-in Givebacks auto-download via Playwright (Phase 4). Manual CSV
export/drop-in (parsers.parse_givebacks_files) remains the default path
-- this module is only used if the user explicitly sets up Givebacks
credentials in Settings and clicks "Auto-fetch from Givebacks".

Ported from the original notebook's scraper, with three changes required
to make it safe for a generic, distributable, multi-org app:
  - The notebook hardcoded a single org's Givebacks "cause_id" directly
    in the API call. That can't be baked into shared code -- it's now
    OrgConfig.givebacks_cause_id, set per-org in Settings.
  - The notebook used a blocking `input()` call for one-time-passcode
    entry, which only works in a terminal. `otp_prompt` is now an
    injectable callback so a GUI can supply a dialog instead.
  - The notebook captured an auth token from localStorage/response
    headers but never actually used it (the real API calls only used
    session cookies) -- dropped, since it was dead code that also
    printed truncated tokens to the console for no benefit.

The password is never handled by this module directly except in-memory,
for the duration of the login call -- see credentials.py for how it's
stored/retrieved, and never logged/printed here.
"""

import csv
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from typing import Callable

import httpx

from pta_treasurer.config import OrgConfig
from pta_treasurer.credentials import get_givebacks_password

OtpPrompt = Callable[[], str]


class GivebacksConfigError(Exception):
    """Raised when required Givebacks settings/credentials are missing."""


class GivebacksLoginError(Exception):
    """Raised when login (including OTP) fails."""


def install_chromium(progress_cb: Callable[[str], None] | None = None) -> None:
    """Runs `playwright install chromium`. Call this once per machine
    before the first auto-download -- the Chromium binary is
    deliberately not bundled with the app (see PLAN.md Phase 3/4)."""
    process = subprocess.Popen(
        [sys.executable, '-m', 'playwright', 'install', 'chromium'],
        stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True,
    )
    for line in process.stdout:
        if progress_cb:
            progress_cb(line.rstrip())
    return_code = process.wait()
    if return_code != 0:
        raise RuntimeError(f'playwright install chromium failed (exit {return_code})')


# ── Pure helpers (no browser/network -- unit-testable) ─────────────────────

def select_paid_payout_ids(payouts: list[dict], target_year: int, target_month: int) -> list[str]:
    """Filters raw payout API records down to the paid payout IDs whose
    arrival_date falls in the target month. Order-preserving, de-duped."""
    prefix = f'{target_year}-{target_month:02d}-'
    ids = []
    seen = set()
    for payout in payouts:
        arrival = payout.get('arrival_date', '')[:10]
        pid = payout.get('uuid')
        if pid and arrival.startswith(prefix) and payout.get('status') == 'paid':
            if pid not in seen:
                seen.add(pid)
                ids.append(pid)
    return ids


def parse_summary_table_rows(raw_rows: list[list[str]]) -> list[dict]:
    """Turns raw table cell text (rows of >=3 string cells: item,
    category, transaction count, total) into the CSV row dicts
    parsers.parse_givebacks_files expects. Skips header/Total rows and
    blank items."""
    rows_data = []
    for cells in raw_rows:
        if len(cells) < 3:
            continue
        item = cells[0].strip()
        if not item or item in ('Item', 'Total'):
            continue
        category = cells[1].strip()
        try:
            txns = int(cells[2].strip())
        except ValueError:
            txns = 0
        total = 0.0
        if len(cells) > 3:
            try:
                total = float(cells[3].strip().replace('$', '').replace(',', ''))
            except ValueError:
                total = 0.0
        rows_data.append({
            'Item': item,
            'Categories': category,
            'No. of Transactions': txns,
            'Total': f'${total:,.2f}',
        })
    return rows_data


def write_payout_csv(rows_data: list[dict], dest_folder: Path, month_short: str, payout_id: str) -> Path:
    dest_folder.mkdir(parents=True, exist_ok=True)
    fname = dest_folder / f'givebacks_{month_short}_{payout_id}.csv'
    with open(fname, 'w', newline='') as f:
        writer = csv.DictWriter(
            f, fieldnames=['Item', 'Categories', 'No. of Transactions', 'Total'])
        writer.writeheader()
        writer.writerows(rows_data)
    return fname


# ── Main entry point ────────────────────────────────────────────────────────

async def download_givebacks(
    month_label: str,
    config: OrgConfig,
    dest_folder: Path,
    data_dir: Path,
    otp_prompt: OtpPrompt | None = None,
) -> list[Path]:
    """Downloads Givebacks payout CSVs for `month_label` (e.g. 'July
    1999') into dest_folder, logging into Givebacks with the org's saved
    credentials if no valid session is cached. Returns the list of CSV
    paths written (empty if none already existed and none were found for
    the month)."""
    from playwright.async_api import TimeoutError as PWTimeout
    from playwright.async_api import async_playwright

    if not config.givebacks_org_url or not config.givebacks_email:
        raise GivebacksConfigError(
            'Givebacks org URL and email must be set in Settings before '
            'auto-download can run.')

    password = get_givebacks_password(config.givebacks_email)
    if not password:
        raise GivebacksConfigError(
            'No Givebacks password saved for this email. Set it in Settings.')

    month_short = month_label.split()[0].lower()
    target_month = datetime.strptime(month_label, '%B %Y').month
    target_year = int(month_label.split()[1])
    org_url = config.givebacks_org_url.rstrip('/')

    existing = list(dest_folder.glob(f'givebacks_{month_short}*.csv'))
    if existing:
        return existing

    session_dir = Path(data_dir) / 'data' / 'browser_session'
    session_file = session_dir / 'givebacks_session.json'
    session_dir.mkdir(parents=True, exist_ok=True)

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=True, args=['--no-sandbox', '--disable-dev-shm-usage'])

        if session_file.exists():
            ctx = await browser.new_context(storage_state=str(session_file))
        else:
            ctx = await browser.new_context()

        page = await ctx.new_page()

        # The Givebacks API needs a `cause_id` query param, but it's an
        # internal identifier that never appears anywhere in the UI --
        # there's no field a treasurer could look up and copy in. Instead
        # of requiring it as manual config, capture it from the request
        # the payouts page itself makes when it loads (falling back to
        # OrgConfig.givebacks_cause_id as a manual override, in case
        # auto-discovery ever fails against a future site change).
        discovered_cause_id = None

        def _capture_cause_id(request):
            nonlocal discovered_cause_id
            if discovered_cause_id is None and 'api.givebacks.com/services/payout/payouts' in request.url:
                from urllib.parse import urlparse, parse_qs
                query = parse_qs(urlparse(request.url).query)
                if 'cause_id' in query:
                    discovered_cause_id = query['cause_id'][0]

        page.on('request', _capture_cause_id)

        try:
            await page.goto(f'{org_url}/payouts', timeout=60000)
            await page.wait_for_load_state('domcontentloaded')
            await page.wait_for_timeout(3000)

            current_url = page.url
            login_form = await page.locator('input[type="password"]').count()
            needs_login = ('login' in current_url.lower()
                           or 'sign-in' in current_url.lower()
                           or 'one-time-passcode' in current_url
                           or login_form > 0)

            if needs_login:
                await _login(page, org_url, config.givebacks_email, password, otp_prompt)
                await ctx.storage_state(path=str(session_file))
                await page.goto(f'{org_url}/payouts', timeout=60000)
                await page.wait_for_load_state('domcontentloaded')
                await page.wait_for_timeout(2000)

            all_cookies = await ctx.cookies()
            cookie_dict = {c['name']: c['value'] for c in all_cookies
                           if 'givebacks.com' in c.get('domain', '')}

            if discovered_cause_id is None:
                # Reload once more in case the first load raced the listener.
                await page.goto(f'{org_url}/payouts', timeout=60000)
                await page.wait_for_load_state('domcontentloaded')
                await page.wait_for_timeout(3000)

            cause_id = discovered_cause_id or config.givebacks_cause_id
            if not cause_id:
                raise GivebacksLoginError(
                    "Couldn't determine this org's Givebacks cause ID from "
                    'the payouts page. Set it manually in Settings if you '
                    'know it.')

            all_payouts = []
            offset = 0
            async with httpx.AsyncClient() as client:
                while True:
                    r = await client.get(
                        'https://api.givebacks.com/services/payout/payouts',
                        params={'cause_id': cause_id,
                                'limit': 25, 'offset': offset},
                        cookies=cookie_dict,
                        headers={'Origin': org_url, 'Referer': f'{org_url}/payouts'},
                        timeout=30,
                    )
                    if r.status_code != 200:
                        raise GivebacksLoginError(
                            f'Givebacks API returned {r.status_code} -- '
                            'the saved session may have expired.')
                    data = r.json()
                    all_payouts.extend(data.get('payouts', []))
                    if not data.get('meta', {}).get('has_more', False):
                        break
                    offset += 25

            payout_ids = select_paid_payout_ids(all_payouts, target_year, target_month)
            if not payout_ids:
                return []

            saved_files = []
            for payout_id in payout_ids:
                await page.goto(f'{org_url}/payouts/{payout_id}/summary', timeout=60000)
                await page.wait_for_load_state('domcontentloaded')
                try:
                    await page.wait_for_selector('table tbody tr', timeout=10000)
                except PWTimeout:
                    pass

                raw_rows = []
                for tr in await page.locator('table tbody tr').all():
                    cell_texts = [
                        (await cell.inner_text()) for cell in await tr.locator('td').all()
                    ]
                    raw_rows.append(cell_texts)

                rows_data = parse_summary_table_rows(raw_rows)
                if not rows_data:
                    continue
                saved_files.append(write_payout_csv(rows_data, dest_folder, month_short, payout_id))

            return saved_files

        except Exception as e:
            debug_dir = Path(data_dir) / 'logs'
            debug_dir.mkdir(parents=True, exist_ok=True)
            screenshot_path = debug_dir / f'givebacks_debug_{int(datetime.now().timestamp())}.png'
            try:
                await page.screenshot(path=str(screenshot_path))
            except Exception:
                screenshot_path = None
            suffix = f' (see {screenshot_path})' if screenshot_path else ''
            raise GivebacksLoginError(f'{e}{suffix}') from e
        finally:
            await ctx.close()
            await browser.close()


async def _login(page, org_url: str, email: str, password: str, otp_prompt: OtpPrompt | None) -> None:
    await page.goto(f'{org_url}/sign-in', timeout=30000)
    await page.wait_for_load_state('domcontentloaded')

    await page.wait_for_selector('input[type="email"], input[placeholder*="email" i]', timeout=10000)
    await page.fill('input[type="email"], input[placeholder*="email" i]', email)
    await page.fill('input[type="password"], input[name="password"]', password)

    for btn in await page.locator('button').all():
        if (await btn.inner_text()).strip() == 'Sign In':
            await btn.click()
            break
    await page.wait_for_load_state('domcontentloaded')
    await page.wait_for_timeout(2000)

    if 'one-time-passcode' in page.url:
        if otp_prompt is None:
            raise GivebacksLoginError(
                'Givebacks requires a one-time passcode for this login, '
                'but no OTP entry was available.')
        import asyncio
        code = (await asyncio.get_event_loop().run_in_executor(None, otp_prompt)).strip()

        otp_boxes = [box for box in await page.locator('input').all() if await box.is_visible()]
        for i, digit in enumerate(code[:6]):
            if i < len(otp_boxes):
                await otp_boxes[i].click()
                await otp_boxes[i].type(digit)
                await page.wait_for_timeout(200)
        await page.wait_for_timeout(1000)

        trust = page.locator('input[type="checkbox"]')
        if await trust.count() > 0:
            await trust.click()

        await page.wait_for_selector('button:has-text("Submit"):not([disabled])', timeout=15000)
        await page.click('button:has-text("Submit")')
        await page.wait_for_load_state('domcontentloaded')
        await page.wait_for_timeout(2000)

    if 'one-time-passcode' in page.url or 'login' in page.url.lower():
        raise GivebacksLoginError('Login failed -- credentials or OTP code may be incorrect.')
