# Packaging

PyInstaller specs and build scripts that turn `pta_treasurer` into a
double-clickable desktop app. See `PLAN.md` (Phase 3) for the roadmap
context.

## Layout

- `launcher.py` — the PyInstaller entry script (`from
  pta_treasurer.gui.app import main; main()`). PyInstaller analyzes a
  script file, not a module path, so this tiny file is the actual thing
  each spec builds.
- `pyinstaller_mac.spec` / `build_mac.sh` — macOS build. **Built and
  hand-tested** on this repo's dev machine (Apple, macOS, x86_64 Python
  build).
- `pyinstaller_win.spec` / `build_win.ps1` — Windows build. **Written but
  untested** — PyInstaller doesn't cross-compile, so this can only be
  built and verified on an actual Windows machine. There's no CI in this
  repo yet to do that automatically.
- `entitlements.plist` / `notarize_mac.sh` / `build_mac_release.sh` —
  optional signed & notarized macOS release build, for direct
  distribution (not the Mac App Store — see "Signed & Notarized Release
  Build" below). Written but not run against a real Apple Developer
  account from this environment.

## Building on Mac

```sh
packaging/build_mac.sh
```

This creates a **fresh** venv (`.venv-build-mac`, not your dev `.venv`)
and does a non-editable `pip install .` before building — this catches
"works in my dev venv but missing from a real install" bugs that an
editable install would hide, and avoids PyInstaller's import analysis
following editable-install `.pth` indirection. It reuses the same Python
interpreter as `.venv` (currently 3.14.2) rather than whatever `python3`
resolves to on `PATH`, since on this machine that's an unrelated
Anaconda-provided Python 3.8.

Output: `dist/PTA Treasurer.app` (~137MB, onedir/`BUNDLE` build).

First launch will show macOS Gatekeeper's "Apple could not verify..."
warning, since the build isn't codesigned/notarized (deliberately
deferred — see `PLAN.md`'s Phase 3 entry). Right-click the `.app` → Open
to bypass it. This is expected on an unsigned build, not a bug.

## Building on Windows

```powershell
.\packaging\build_win.ps1
```

Same shape as the Mac build: fresh venv, non-editable install, pinned
PyInstaller. Output: `dist\PTA Treasurer\PTA Treasurer.exe`.

**This is unverified.** Known unknowns to check on a real Windows build:
- Whether PySide6's plugin-path handling needs extra spec entries (platform
  plugins, styles) — a common Windows-specific PyInstaller gotcha.
- Whether `pypdfium2_raw`'s Windows DLL gets picked up by the same
  `pyinstaller-hooks-contrib` hook that worked for the macOS `.dylib`.
- General hand-test per the checklist below, run on Windows.

## Hand-test checklist

Run this after every build before considering it done. Mirrors
`.claude/skills/verify/SKILL.md`'s isolation pattern — never touch your
real data folder or the real `platformdirs` pointer.

1. Launch the built app (double-click in Finder/Explorer, not from a
   dev-environment terminal) — confirms the real-world launch path, not
   one with dev env vars still active.
2. First-run wizard: point the data folder picker at a **scratch**
   directory (not a real one). Enter an org name, fiscal start month,
   balance forward. Confirm `budget.xlsx` gets created and the wizard
   completes.
3. Copy `sample_data/July_1999/` into
   `<scratch_dir>/input/July_1999/`.
4. Pick July / 1999 in the main window, click **Generate Report**, wait
   for it to finish.
5. Open the generated `Treasurer_Report_July_1999.xlsx` and confirm:
   - All 6 tabs present: Treasurer Report, Income Budget vs Actuals,
     Expense Budget vs Actuals, Giveback Reconciliation, File Manifest,
     YTD Summary.
   - QB expense_total = 1439.34, income_total = 0.
   - Bank beginning = 32630.10, ending = 31190.76, checks = 181.58,
     withdrawals = 1257.76.
   - Givebacks = 2 items totalling 110.14.
   - Reconciliation "Difference" = 0.
6. Open Settings, confirm the org name/balance forward from step 2
   persisted. Try "Run All Months" against the single scratch month.
7. Delete the scratch directory when done. Confirm your real
   `platformdirs` config location (`~/Library/Application
   Support/pta-treasurer/` on Mac, `%LOCALAPPDATA%\pta-treasurer\` on
   Windows) was never touched.

### Automated proxies used during Mac packaging development

Since this environment has no way to drive GUI clicks on a frozen
binary, the Mac build was verified with these automated checks instead
of (in addition to) the manual checklist above — worth re-running after
any spec change:

- Launch with `QT_QPA_PLATFORM=offscreen` and a pre-seeded scratch
  `HOME` (pointer file + `org_config.json` + `budget.xlsx` already in
  place) — confirms the app reaches `MainWindow` (skipping the wizard)
  with no import errors, which exercises the full bundled module graph
  (`config` → `pipeline` → `budget_io` → `parsers`/`builders`).
- Launch with a stripped environment (`env -i PATH=/usr/bin:/bin`) — a
  single-machine proxy for "runs with no Python/dev-tools installed."
- `otool -L` on the bundled binary + confirming `Contents/Frameworks/`
  contains a full `Python.framework` — proves the app doesn't dynamically
  depend on the Homebrew Python it was built with.
- Confirmed `pypdfium2_raw/libpdfium.dylib` (pdfplumber's native PDF
  backend) is present in the bundle — direct evidence the
  `pyinstaller-hooks-contrib` native-lib hook worked, since a missing
  native lib would surface as a PDF-parsing failure at runtime, not a
  build-time error.

These are *not* a substitute for the full manual checklist (they don't
click through the wizard or verify report numbers) — they were used to
de-risk the packaging-specific unknowns (does freezing break imports?)
given the business logic itself was already covered by Phase 1/2's test
suite and manual verification.

## Deliberately not bundled: `playwright`

`givebacks_download.py` is only ever imported lazily, inside
`GivebacksWorker.run()` (`gui/main_window.py`) — not at module top level —
specifically so PyInstaller's static `Analysis()` never traces `playwright`
into the bundle from `launcher.py`'s import graph. Add it to
`hiddenimports` and re-run the full build+hand-test cycle if that ever
changes (e.g. Givebacks auto-download becomes a required, not opt-in,
feature).

This doesn't contradict `PLAN.md`'s Phase 4 decision that "the
`playwright` package ships in the app... not the Chromium binary" — that
describes what the *finished* product bundles once someone actually
clicks "Auto-fetch" and the lazy import fires at runtime from within the
frozen bundle's own Python environment, which still has `playwright`
installed (it's a normal pip dependency of the built venv) even though
PyInstaller's static analysis doesn't need to trace it ahead of time.

`httpx`, by contrast, **is** bundled automatically now — `gui/chat_panel.py`
imports it at module top level (for `httpx.ConnectError`/`HTTPStatusError`
exception handling around the AI Assistant's Ollama calls), and
`chat_panel.py` is itself imported at module top level by `main_window.py`,
so PyInstaller's static analysis picks it up without needing a
`hiddenimports` entry — same mechanism, opposite outcome, depending on
where in the import graph a module sits.

## Signed & Notarized Release Build

The default `build_mac.sh` produces an **unsigned** `.app` — fine for
local testing, but it triggers macOS Gatekeeper's "Apple could not
verify this app is free of malware" warning for anyone else who
downloads it. Fixing that for real distribution means Developer ID
signing + Apple notarization (not the Mac App Store — that's a much
bigger undertaking involving App Sandbox, which this app doesn't run
under, and would likely require redesigning or dropping Givebacks
auto-download since sandboxed apps can't launch arbitrary subprocesses
like Playwright's Chromium).

### One-time setup (only you can do this — needs your own Apple ID)

1. Enroll in the **Apple Developer Program** ($99/yr) at
   [developer.apple.com](https://developer.apple.com).
2. Create a **"Developer ID Application" certificate**: open Keychain
   Access → Certificate Assistant → Request a Certificate from a
   Certificate Authority (this generates a CSR + private key in your
   login keychain) → upload the CSR at
   developer.apple.com/account/resources/certificates → download the
   issued certificate and double-click it to install into your login
   keychain.
3. Generate an **app-specific password** at
   [appleid.apple.com](https://appleid.apple.com) (Sign-In and Security
   → App-Specific Passwords) — used for notarization, not your regular
   Apple ID password.
4. Run once, in Terminal:
   ```sh
   xcrun notarytool store-credentials "pta-treasurer-notarize" \
     --apple-id "you@example.com" \
     --team-id "YOUR_TEAM_ID" \
     --password "the-app-specific-password"
   ```
   This stores the credentials in your macOS Keychain under the profile
   name `pta-treasurer-notarize` — the same "never plaintext, always the
   OS keychain" pattern the app itself uses for the Givebacks password
   (`credentials.py`). `notarize_mac.sh` only ever references that
   profile *name*, never a secret.

### Building a signed, notarized release

```sh
CODESIGN_IDENTITY="Developer ID Application: Your Name (TEAMID)" \
  packaging/build_mac_release.sh
```

Find your exact identity string with `security find-identity -v -p codesigning`.

This runs `build_mac.sh` as normal, then `notarize_mac.sh`, which:
1. Deep-signs every nested binary with the hardened runtime +
   `entitlements.plist` (PyInstaller's own signing only covers the
   top-level bundle — nested frameworks/dylibs like `Python.framework`
   and `pypdfium2_raw`'s `libpdfium.dylib` need an explicit deep-sign
   pass for notarization specifically, which checks every nested code
   object).
2. Zips the `.app` and submits it via `notarytool` (can take several
   minutes — Apple's service, not a local step).
3. Staples the notarization ticket to the `.app` on success, so
   Gatekeeper can verify it offline.
4. Runs `spctl -a -vvv "dist/PTA Treasurer.app"` — look for `accepted`
   and `source=Notarized Developer ID` in the output. That's the
   concrete signal it worked.

Leaving `CODESIGN_IDENTITY` unset makes `build_mac_release.sh` behave
identically to `build_mac.sh` (unsigned) — the normal dev build loop is
unaffected either way.

`entitlements.plist` is deliberately **not** an App Sandbox entitlements
file — this app isn't sandboxed (that's the Mac-App-Store-only
requirement this whole path avoids). Its four entries exist solely to
satisfy notarization for a PyInstaller-frozen Python app (relaxed
library-validation/executable-memory rules for bundled C extensions and
the Python interpreter itself).

## Version pins

`pyinstaller==6.21.0` and `pyinstaller-hooks-contrib` (currently
`2026.6`) are pinned in `pyproject.toml`'s `dev` extras and
`requirements.txt`, matching the versions the Mac build was verified
against. If you upgrade either, re-run the full build + hand-test cycle
before trusting the new pin.
