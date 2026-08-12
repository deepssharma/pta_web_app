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

## Future: Treasurer Copilot (LLM, deferred past v1)

Discussed but **not part of Build order phases 1–5 above** — a v2 idea, noted here so it doesn't get lost or accidentally designed into a corner while v1 is being built.

**The idea:** a text chat panel docked beside whatever report is open in the GUI, so a board member can ask plain-English questions ("why is Book Fair over budget?") and get an answer grounded in the data already on screen. See the concept mockup discussed with the user (navy/teal/gold palette pulled directly from `builders.py`'s existing report styling, not a generic chat-UI look) for what this could look like visually.

**Key decisions already made:**
- **Text, not voice.** Voice adds real complexity (speech-to-text, latency) for little payoff, and it's awkward to dictate financial questions in a shared board setting.
- **No RAG / vector index.** A single PTA's annual data is a few hundred transactions in small JSON files — that easily fits directly in an LLM's context window. Filtering the relevant month/category and passing that structured slice as context is simpler and more reliable than building an embeddings/chunking/vector-search pipeline to solve a scale problem this product doesn't have.
- **Deferred past v1 on purpose.** The core app is designed to be a fully local, offline, no-account installable tool (see "Data location"/"Platforms" above) — that's a meaningfully different product shape than one that needs an API key, network calls, and a cost model. Don't let this feature creep into phases 1–5.

**What v1 should still do, to keep this door open cheaply:** keep the JSON history schema (`data/history/{Month}_{Year}.json`) and budget-line naming (`config.py`/`budget_io.py`) consistent and stable — that's the only thing a future copilot would actually depend on, and it costs nothing extra to keep tidy now.

## Future: agent-driven budget edits (write mode) — Tier 1 + Tier 2 built 2026-08-12

Both tiers below are now implemented (`budget_io.py`'s six write functions
+ `apply_edit`/`describe_edit`, `ai_assistant.py`'s `edit_budget` mode +
`parse_edit_action`, `chat_panel.py`'s Apply/Discard preview flow) and
covered by tests; live-verified against a real local `llama3.2` model,
not just mocks. See the `ai_assistant.py`/`chat_panel.py` entries in
`CLAUDE.md`'s Architecture section for the current shape. Kept below as
the original design rationale.

Extends the read-only Copilot idea above to the question that actually prompted it: *if a treasurer tells an agent "add a category," "remove one," or "move Book Fair under Programs instead of Fundraising," can it just do that?*

**Yes, technically — if `budget_io.py` exposes edits as a handful of structured functions, and no agent (in-app or external) ever touches the Excel template or the Python code directly:**

```python
budget_io.add_item(section, item, last_year_actual=0.0, budget=0.0, qb_names=[])
budget_io.remove_item(item)                             # refuses if the item has nonzero actuals this FY, unless force=True
budget_io.move_item(item, from_section, to_section)      # the "nest it under something else" case — pure re-key, no data at risk
budget_io.rename_item(old_name, new_name)                # keeps QB_TO_BUDGET_MAP pointed at the renamed item
budget_io.map_qb_category(qb_name, budget_item)          # "QuickBooks changed a category's name" case
budget_io.set_budget_amount(item, new_amount)
```

Any agent can drive these — an in-app copilot (below), or a general coding agent (Claude Code/Desktop) pointed straight at the user's data folder, no different from how categories get edited in `pta_treasurer` today, in this very session.

**Two ways this could ship, not mutually exclusive:**
1. **Bring-your-own-agent — effectively free, v1-compatible.** Once `budget_io.py` exists (Build order Phase 1) with clear function signatures/docstrings, a treasurer who already has a coding agent can point it at `~/Documents/PTA Treasurer/` and ask for the edit directly. No GUI work required beyond Phase 1 as already planned.
2. **In-app copilot, write mode — v2, extends the deferred Copilot idea.** The docked chat panel (above) gets write access to the same six functions, gated behind a **preview-diff-then-confirm** step (show old vs. new state, require an explicit "Apply"). This is the piece that reintroduces the API-key/network tradeoff the v1 design deliberately avoids (see "Deferred past v1 on purpose" above), so it inherits the same deferral — just scoped concretely now instead of hand-waved.

**Guardrails either way:**
- All writes go through the structured functions only — never a raw file or code edit by the agent — so a bad LLM guess can't corrupt the template's shape.
- `remove_item`/`rename_item` check the fiscal year's actuals first and refuse/warn rather than silently dropping real transaction history from the budget view.
- Auto-backup the previous template before any agent-driven write (a dated copy or `.bak`), mirroring the undo-via-git safety net this repo already has.

Not deciding now — this is the sketch, per your ask. Say the word if you want tier 1, tier 2, both, or neither turned into actual Build-order phases.

## Verification

- `pytest` green in the new repo at each phase (ported tests unchanged in behavior; new tests for `budget_io`/`pipeline`).
- CLI smoke test against the synthetic `sample_data/July_1999/` dataset produces a 6-tab workbook with correct values (same manual check already validated against the current notebook this session — automate it as the `test_pipeline.py` baseline).
- GUI: manual smoke-test checklist (first-run wizard → generate a report against synthetic data → confirm output) on both a Mac and Windows machine before considering v1 done; automated GUI testing (e.g. `pytest-qt`) is a nice-to-have, not required for v1.
- Packaging: install and run the built app on a clean Mac and Windows machine (or VM) with no Python installed, to confirm the bundle is actually self-contained.
