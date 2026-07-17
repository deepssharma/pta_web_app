# Convert PTA Treasurer Report Generator into a cross-platform desktop app

## Context

`pta_treasurer` currently works, but only for one technically-savvy user: it requires installing Python, Jupyter, and hand-editing a notebook cell to change org name, fiscal year, and budget categories every year. The goal is a real installable desktop app other PTAs/schools can download and run without touching Python, so any treasurer can generate their monthly report and YTD summary.

Good news from inspecting the current code: `parsers.py` and `builders.py` are already fully portable — every function takes `org_name`/`month_label`/data as parameters, with zero hardcoded org-specific strings. They can move to the new project essentially unchanged. The actual blocker is one large notebook cell (`INCOME_BUDGET`/`EXPENSE_BUDGET` — a ~80-line-item nested dict of this PTA's real budget categories and dollar figures, plus `QB_TO_BUDGET_MAP`) that has to become user-editable data instead of code, plus everything about *how* the tool runs (notebook → real GUI) and *how* it ships (pip install → double-click installer).

Per your direction: this starts as a **new, separate project in a fresh directory** — not nested inside `pta_treasurer` — so the current repo's real financial history never mixes with the generic, distributable codebase. The `input/{Month}_{Year}/`, `output/Treasurer_Report_{Month}_{Year}.xlsx`, and `data/history/` folder conventions carry over unchanged, just relocated to a user-chosen folder an installed app can actually write to (installers can't write into `Program Files`/`Applications`).

**Target directory: `/Users/deepalisharma/pta_web_app`.** You'll continue this work in a fresh Claude Code session started in that directory, so this session's job is just **Step 0: scaffold it now** — create the layout below, port the reusable files verbatim, copy this plan into the new repo so the next session has full context without depending on this machine's global `~/.claude/plans/`, and git-init it. Nothing beyond that gets written now — `config.py`/`budget_io.py`/`pipeline.py`/the GUI are genuinely new code and belong to Build-order phases 1-5 below, done in the new session.

**Confirmed decisions:**
- GUI: **PySide6** (more polished than Tkinter; you accepted the larger installer size)
- Givebacks auto-download (Playwright): **kept as an opt-in advanced feature** — the `playwright` package ships in the app, but the Chromium browser binary is fetched on demand via an in-app "Install browser automation" button (`playwright install chromium`), not bundled in the main installer. Manual CSV export/drop-in (like QuickBooks/Chase already work) is always available as the default path.
- Data location: **user-chosen folder** at first run (default suggestion `~/Documents/PTA Treasurer/`), preserving the exact same `input/`, `output/`, `data/history/` subfolder convention inside it.
- Platforms: **Mac + Windows** for v1.

## New project layout

```
pta_web_app/                        # /Users/deepalisharma/pta_web_app — new git repo
├── PLAN.md                          # this plan, copied in so a fresh session has full context
├── pyproject.toml
├── requirements.txt
├── README.md                        # short: what this is, current status, points at PLAN.md
├── src/pta_treasurer/
│   ├── parsers.py                  # ported from pta_treasurer/parsers.py, ~unchanged  [Step 0]
│   ├── builders.py                 # ported, ~unchanged — becomes the single source of FISCAL_MONTHS  [Step 0]
│   │                                #   (today it's duplicated between builders.py and the notebook)
│   ├── budget_io.py                 # NEW — read/write the budget Excel template (see below)  [Phase 1]
│   ├── config.py                    # NEW — OrgConfig (org name, fiscal start month, balance_forward,
│   │                                #   data_dir, Givebacks creds) as JSON; platformdirs for the
│   │                                #   small pointer file that remembers where the user's data folder is  [Phase 1]
│   ├── pipeline.py                  # NEW — run_month(config, month, year) -> RunResult; single
│   │                                #   orchestration path both the GUI and a CLI call  [Phase 1]
│   ├── givebacks_download.py        # ported from notebook Cell 2 (Playwright), invoked only if opted in  [Phase 4]
│   └── gui/                         # [Phase 2]
│       ├── app.py                   # QApplication entry point
│       ├── setup_wizard.py          # first run: data folder, org name, fiscal start, starting balance
│       ├── main_window.py           # month picker, file status, "Generate Report", progress log
│       └── settings_dialog.py       # edit org config, data folder, Givebacks creds, "Run All Months"
├── packaging/                       # [Phase 3]
│   ├── pyinstaller_mac.spec
│   ├── pyinstaller_win.spec
│   └── README.md                    # build steps per platform
├── sample_data/July_1999/           # copied verbatim from pta_treasurer  [Step 0]
└── tests/
    ├── fixtures/                    # copied verbatim from pta_treasurer/tests/fixtures/  [Step 0]
    ├── test_parsers.py              # ported  [Step 0]
    ├── test_builders.py             # ported  [Step 0]
    ├── test_budget_io.py            # NEW  [Phase 1]
    └── test_pipeline.py             # NEW — runs against sample_data/July_1999/  [Phase 1]
```

Bracketed tags above mark when each piece gets built: `[Step 0]` = this session, right now; `[Phase N]` = the numbered steps under **Build order** below, done in the new Claude session.

### Step 0 — scaffold now, in this session

1. Create the directory tree above under `/Users/deepalisharma/pta_web_app` (just the `[Step 0]`-tagged paths — no stub files for anything tagged `[Phase N]`).
2. Copy verbatim, no edits: `parsers.py`, `builders.py`, `tests/test_parsers.py`, `tests/test_builders.py`, `tests/fixtures/*` (the synthetic fixtures), `sample_data/July_1999/*` — all from `/Users/deepalisharma/Desktop/pta/codes/pta_treasurer`.
3. Write `requirements.txt` for the new project: the current core deps (`openpyxl`, `pdfplumber`, `python-dotenv`, `httpx`) plus the new ones this plan already commits to (`PySide6`, `platformdirs`, `playwright` per the opt-in decision) and dev deps (`pytest`, `pyinstaller`). Pin to currently-installed-equivalent versions where known, otherwise unpinned with a comment that it needs a real pin once installed.
4. Write a minimal `pyproject.toml` (package name `pta_treasurer`, src layout) so `pip install -e .` and `pytest` work out of the box.
5. Write `.gitignore` mirroring `pta_treasurer`'s data-safety rules (`input/`, `output/`, `data/`, `logs/`, `*.xlsx`/`*.pdf`/`*.csv` blanket-ignored, with the same `!tests/fixtures/*`/`!sample_data/**` exceptions) — this app will handle real school financial data too, so the same rules apply from day one.
6. Copy this plan file's content into `PLAN.md` at the new repo's root.
7. Write a short `README.md`: one paragraph on what the project is, current status ("scaffolded, not yet implemented — see PLAN.md"), and how to run tests.
8. Verify the port: `pip install -r requirements.txt` (or at least the core 4) and `pytest` green against the copied tests/fixtures, in the new location.
9. `git init`, initial commit of the scaffold. No remote — that's a separate step for whenever the user's ready to push.

## The budget config problem, and the chosen fix

Today `INCOME_BUDGET`/`EXPENSE_BUDGET` (notebook, "Annual Budget Data" cell) is a nested dict: `section -> item -> (last_year_actual, this_year_budget)`, plus `QB_TO_BUDGET_MAP: qb_category_name -> budget_item_name`. It's edited by hand-editing Python.

Since treasurers already work in QuickBooks/Excel, the simplest and most reliable v1 approach is an **Excel template as the config format**, not a custom in-app table editor:
- `budget_io.generate_template(path)` writes a workbook with an "Income Budget" and "Expense Budget" sheet, columns: `Section | Item | QuickBooks Category Name(s) | Last Year Actual | This Year Budget`.
- `budget_io.load_budget(path) -> (income_budget, expense_budget, qb_to_budget_map)` reads it back into exactly the structures `merge_actuals_into_budget()` (ported unchanged from the notebook) already expects.
- The GUI's "Edit Budget" action just opens this file in the user's default spreadsheet app (`subprocess` / `os.startfile` depending on platform) and offers a "Reload" button — no custom spreadsheet widget to build. A native in-app editor is a reasonable v2 enhancement, not needed for v1.

## Build order

1. **Scaffold + port core, verify safety net first.** Create the new repo, copy `parsers.py`, `builders.py`, `tests/test_parsers.py`, `tests/test_builders.py`, and the synthetic `tests/fixtures/` built earlier this session. Get `pytest` green in the new repo before writing anything new — this is the regression net for every later step.
2. **Headless core (no GUI yet).** Build `config.py`, `budget_io.py`, `pipeline.py` by porting the logic currently spread across notebook cells 1 (config), 3 (find/validate month files), 4 (parse + save history), 5 (load all actuals), 6 (merge budget), and 9 (build workbook). Expose `run_month()` and a thin CLI (`python -m pta_treasurer generate --month July --year 2025`). Copy `sample_data/July_1999/` over and verify the CLI produces a correct 6-tab workbook against it — same check performed manually against the current notebook this session, now automated as `test_pipeline.py`.
3. **PySide6 GUI.** Setup wizard (data folder → org name/fiscal start/balance forward → budget template) then main window (month/year picker, per-file status with Browse/drag-drop, "Generate Report" running `pipeline.run_month()` on a background `QThread` so the UI stays responsive, progress log, "Open Report"/"Open Folder" on success). Settings dialog for editing config after first run, plus a "Run All Months" action mirroring `run_all_months.sh`'s batch behavior.
4. **Packaging.** PyInstaller spec per platform (`--windowed`, bundles PySide6/openpyxl/pdfplumber/python-dotenv/httpx/playwright-the-package but not Chromium). Produce and hand-test a real Mac build and Windows build. Flag but don't block on: Mac notarization/code-signing (unsigned apps get an "unidentified developer" Gatekeeper warning — real UX friction for non-technical users, worth a follow-up pass once the app itself is solid).
5. **Givebacks auto-download (opt-in).** Port the notebook's Playwright scraper into `givebacks_download.py`, wired to Settings' "Install browser automation" button (runs `playwright install chromium` with a progress indicator) and a per-month "Auto-fetch from Givebacks" action. Kept last since it's optional and the highest-risk/most complex piece (async browser automation inside a Qt event loop, OTP login flow).

## What gets dropped / changed vs. today

- **`push_to_github()`** (notebook's git-push cell) — this was specific to your personal workflow of versioning your own reports; a generic multi-school app has no GitHub repo to push to. Not ported.
- **Duplicate `FISCAL_MONTHS`** (currently defined once in `builders.py` and again inline in the notebook config cell) — consolidates to one definition in `builders.py`, imported everywhere else.
- **Hardcoded `INCOME_BUDGET`/`EXPENSE_BUDGET`/`QB_TO_BUDGET_MAP`** — becomes the Excel-template-driven config described above.
- **`ORG_NAME`, fiscal year, `balance_forward`** — move from notebook variables/hardcoded literals into `config.py`'s `OrgConfig`, editable via the setup wizard/Settings dialog instead of code.

## Verification

- `pytest` green in the new repo at each phase (ported tests unchanged in behavior; new tests for `budget_io`/`pipeline`).
- CLI smoke test against the synthetic `sample_data/July_1999/` dataset produces a 6-tab workbook with correct values (same manual check already validated against the current notebook this session — automate it as the `test_pipeline.py` baseline).
- GUI: manual smoke-test checklist (first-run wizard → generate a report against synthetic data → confirm output) on both a Mac and Windows machine before considering v1 done; automated GUI testing (e.g. `pytest-qt`) is a nice-to-have, not required for v1.
- Packaging: install and run the built app on a clean Mac and Windows machine (or VM) with no Python installed, to confirm the bundle is actually self-contained.
