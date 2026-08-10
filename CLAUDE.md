# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A cross-platform desktop app (PySide6/Qt, packaged later with PyInstaller) that generates monthly PTA/school treasurer reports — income/expense vs. budget, bank reconciliation, YTD summary — from QuickBooks, Chase bank statement, and Givebacks exports. It's a from-scratch, generic/distributable rewrite of a private notebook-based tool (`~/Desktop/pta/codes/pta_treasurer`, not part of this repo), aimed at any PTA treasurer with no Python/spreadsheet-formula knowledge.

**Current status: Phases 1-4 built, plus an AI Assistant chat panel, plus two features ported from the private notebook's later work.** Headless core (parsing/budget/pipeline/CLI), the PySide6 GUI, macOS packaging (built + hand-tested; Windows spec written but unverified — needs a Windows machine), opt-in Givebacks auto-download, an opt-in LLM chatbot panel (Q&A about a generated report, troubleshooting input-file errors, runs on local Ollama), a configurable pass-through-fund section on the Treasurer Report, and an opt-in whole-fiscal-year Debits & Credits ledger workbook all exist. See `.claude/skills/verify/SKILL.md` for the verification recipe and known-good numbers against `sample_data/July_1999/`.

The private notebook at `~/Desktop/pta/codes/pta_treasurer` is still actively used for real monthly work and accumulates real bug fixes/features between sessions — when picking this project back up, it's worth diffing against it again (`git log` there since the last port) rather than assuming this repo is fully caught up.

**`PLAN.md`** has the original high-level roadmap/rationale. **`~/.claude/plans/robust-chasing-walrus.md`** (not part of this repo) has the concrete, up-to-date build-order plan with phase-by-phase status — check it before assuming something is or isn't built.

## Hard rules

- **No Jupyter notebooks, ever.** The predecessor project was notebook-based; the entire point of this rewrite is plain, importable `.py` modules run via pytest/CLI/GUI. Never add `.ipynb` files or notebook-cell-style code.
- **Never commit real financial data.** `input/`, `output/`, `data/`, `logs/`, `org_config.json`, and any `*.xlsx`/`*.pdf`/`*.csv` are gitignored by default. The only checked-in data is synthetic: `tests/fixtures/` and `sample_data/July_1999/` (fake org, fake year, no real names/accounts), which have explicit `.gitignore` un-ignore rules. Any new code that writes real output must respect this — don't relax the `.gitignore` patterns.
- **Never store secrets in plaintext.** The Givebacks password is the one real secret in this app. It's stored via the OS-native credential store (`keyring` library, see `credentials.py`) — never in `org_config.json`, never logged/printed. Only the non-secret Givebacks org URL/email/cause ID live in `OrgConfig`. (The AI Assistant has no secret at all — it runs against a local Ollama server, not a cloud API.)

## Commands

```sh
python -m venv .venv && source .venv/bin/activate
pip install -e ".[dev]"             # core + test/packaging tooling
pip install -e ".[dev,givebacks]"   # add playwright for Givebacks auto-download work
pytest                              # run the full suite (173 tests currently green)
pytest tests/test_parsers.py::test_name   # run a single test

python -m pta_treasurer generate --month July --year 2025 --data-dir ...   # CLI
pta-treasurer                       # launch the GUI (installed console script)
```

There is no lint/format tooling configured yet (no ESLint/Prettier equivalent, no ruff/black/mypy config) — this is current state, not an oversight to silently work around.

**Python version note**: dev environment currently runs Python 3.14 — `playwright==1.48.0`'s pinned `greenlet==3.1.1` has no prebuilt wheel for 3.14 and fails to build from source. Use `playwright>=1.50.0` (pinned to `1.62.0` in `pyproject.toml`/`requirements.txt`) if you touch that dependency.

## Architecture

This is a native Python project — no `package.json`/lockfile. `pyproject.toml` (src-layout, package `pta_treasurer` under `src/`) and `requirements.txt` are the dependency manifests. Core deps: `openpyxl` (Excel output), `pdfplumber` (bank PDF parsing), `httpx` (Ollama HTTP calls), `PySide6` (GUI), `platformdirs` (per-user data dir + org config location), `keyring` (OS-native credential storage, for the Givebacks password only). Optional `givebacks` extra: `playwright` (browser automation) — only needed for Givebacks auto-download, not for the default manual-CSV workflow. No extra dependency group for the AI Assistant — it talks to a local Ollama server over plain HTTP using `httpx`, already core.

- **`src/pta_treasurer/parsers.py`** — parsers, no classes, all producing plain dicts/lists:
  - `parse_quickbooks_detail(folder, input_month, fiscal_year, pass_through_categories=None, pass_through_fund_name='')` — QuickBooks "Transaction Detail by Account" CSV export → income/expense category dicts + a flat transaction list. A configured pass-through fund's transactions are diverted to `pass_through_income_total`/`pass_through_expense_total`/`pass_through_net` and excluded from `income`/`expenses` entirely.
  - `parse_givebacks_files(file_info_list)` — merges Givebacks (fundraising platform) CSV exports into item/category/count/total records.
  - `parse_chase_pdf(bank_file)` — Chase bank statement PDF → beginning/ending balance, deposits, checks, withdrawals, fees, daily balances (for reconciliation).
  - `find_bank_statement_month`/`match_credits_to_bank_statement` — match a QuickBooks transaction (or same-date group, for credits) to the bank statement month it actually cleared in, handling the "recorded this month, cleared next month" lag case.
  - `consolidate_givebacks_payouts`/`extract_payout_id` — collapse a month's QuickBooks credit transactions into one row per real Givebacks payout (3-phase matching: whole-date, within-date subset-sum, cross-date combination), and detect the same payout file appearing under more than one month.
  - Data shapes (categories→amount dicts, transaction dict fields, etc.) aren't repeated here — see the fixtures in `tests/conftest.py` for the canonical shape.
- **`src/pta_treasurer/builders.py`** — pure functions building styled `openpyxl` worksheets: `build_treasurer` (takes an optional `pass_through` dict to render the pass-through-fund section), `build_budget`, `build_givebacks`, `build_manifest`, `build_ytd_summary`, and the whole-fiscal-year ledger sheets `build_credits_sheet`/`build_debits_sheet`/`build_memberhub_summary_sheet`. Defines `FISCAL_MONTHS` (fiscal year starts July) as the single source of truth for month ordering, and `MERGE_PAIRS` for special-case line-item merges (e.g. combining "Membership Expenses/Income").
- **`src/pta_treasurer/config.py`** — `OrgConfig` dataclass (org name, fiscal start, balance forward, non-secret Givebacks identifiers, Ollama host/model, and a configurable pass-through fund: `pass_through_fund_name`/`pass_through_fund_categories`/`pass_through_fund_balance_forward` — blank name means the feature is off), JSON-persisted at `{data_dir}/org_config.json`. A separate `platformdirs`-based pointer file remembers which folder the user chose. Also owns month/file-detection helpers (`detect_month_from_filename`, `detect_month_from_pdf`, fiscal-index math, `fiscal_year_start_calendar_year`).
- **`src/pta_treasurer/budget_io.py`** — Excel-template-based budget config (`generate_template`, `load_budget`, `merge_actuals_into_budget`), replacing the old notebook's hardcoded `INCOME_BUDGET`/`EXPENSE_BUDGET`/`QB_TO_BUDGET_MAP` dicts.
- **`src/pta_treasurer/pipeline.py`** — `run_month(config, data_dir, month, year) -> RunResult`, the single orchestration path used by both the CLI (`__main__.py`) and the GUI. Parses inputs, writes/merges history (`data_dir/data/history/*.json`, now also carrying pass-through-fund totals), assembles the 6-tab workbook via `builders.py`. `compute_pass_through_balance_held()` sums the configured fund's net across all history in true calendar order (survives fiscal-year boundaries). `build_debits_credits_ledger(config, data_dir) -> LedgerResult` is a separate, independent orchestration (triggered from Settings, not the month/year picker) that rebuilds a whole-fiscal-year Credits/Debits/MemberHub_Summary ledger from every `input/{Month}_{Year}` folder in the fiscal year containing today's date — always rebuilds from scratch, resilient to a single bad month (warns and continues).
- **`src/pta_treasurer/credentials.py`** — Givebacks password storage via the OS-native credential store (`keyring`), service name `pta-treasurer-givebacks`, keyed by email. This is the *only* place the password is handled; `NoKeyringBackendError` surfaces a clear message if no OS backend is available, rather than falling back to plaintext.
- **`src/pta_treasurer/givebacks_download.py`** — opt-in Playwright scraper (ported from the original notebook, generalized: the old hardcoded `cause_id` is now auto-discovered from the payouts page's own network request, with `OrgConfig.givebacks_cause_id` as an optional manual override, and OTP entry uses an injectable `otp_prompt` callback instead of a blocking terminal `input()`). Manual CSV drop-in (`parsers.parse_givebacks_files`) remains the default path — this is only invoked if the user sets up credentials in Settings and clicks "Auto-fetch". `install_chromium()` runs `playwright install chromium` (the browser binary is never bundled in packaging).
- **`src/pta_treasurer/ai_assistant.py`** — opt-in LLM chatbot backend. Runs against a **local Ollama server** (`OrgConfig.ollama_host`, default `http://localhost:11434`; `OrgConfig.ollama_model`, default `llama3.2`) — no API key, no cloud dependency, nothing leaves the machine. `build_report_context()`/`build_error_context()` are pure functions enforcing a strict allow-list: only aggregated totals (category totals, budget-vs-actual, bank beginning/ending balances) ever get built into the LLM prompt — individual transactions, check numbers, payer names, and per-item Givebacks detail are never included. Don't add a new field to that allow-list without checking it's a total, not a line item.
- **`src/pta_treasurer/gui/`** — PySide6 app: `app.py` (entry point, shows the setup wizard on first run), `setup_wizard.py`, `main_window.py` (month/year picker, file browse/drag-drop + "Auto-fetch…" for Givebacks, background-thread report generation, a `QDockWidget` AI Assistant side panel toggled from the View menu), `settings_dialog.py` (org config, budget template, Givebacks credentials, Ollama host/model, "Run All Months" batch action), `chat_panel.py` (the AI Assistant panel + its `ChatWorker` background thread). All GUI code calls `pipeline.run_month()` — no business logic duplicated in the GUI layer. The Givebacks OTP prompt is the one place a background `QThread` needs a modal dialog on the main thread — see `GivebacksWorker` in `main_window.py` for the signal/queue handshake pattern.
- **`packaging/`** — PyInstaller specs. macOS: built and hand-tested (`pyinstaller_mac.spec`/`build_mac.sh`). Windows: written but unverified, no Windows machine available (`pyinstaller_win.spec`/`build_win.ps1`). See `packaging/README.md`.
- **No database.** Storage is entirely file-based: a user-provided `input/{Month}_{Year}/` folder (with a `givebacks/` subfolder) in, a generated Excel workbook out.
