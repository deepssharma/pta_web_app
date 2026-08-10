---
name: verify
description: Drive the pta_treasurer CLI and PySide6 GUI end-to-end (not just pytest) to confirm a change actually works.
---

# Verifying pta_treasurer

This is a Python/pytest project with two real user surfaces: a CLI
(`python -m pta_treasurer`) and a PySide6 GUI (`pta_treasurer.gui.app`).
Passing tests do not verify either — always drive the actual surface.

## CLI surface

Isolate in a scratch data dir (never touch a real user's data folder
or the real `platformdirs` pointer file):

```bash
source .venv/bin/activate
mkdir -p /tmp/verify_data/input
cp -r sample_data/July_1999 /tmp/verify_data/input/
python3 -c "
from pathlib import Path
from pta_treasurer.budget_io import generate_template
from pta_treasurer.config import OrgConfig, save_config
data_dir = Path('/tmp/verify_data')
generate_template(data_dir / 'budget.xlsx')
save_config(OrgConfig(org_name='Verify PTA', balance_forward=0.0), data_dir)
"
python3 -m pta_treasurer generate --month July --year 1999 --data-dir /tmp/verify_data
```

Then actually open the workbook and check values, don't just check the
file exists:

```python
import openpyxl
wb = openpyxl.load_workbook('/tmp/verify_data/output/Treasurer_Report_July_1999.xlsx')
assert wb.sheetnames == ['Treasurer Report', 'Income Budget vs Actuals',
                          'Expense Budget vs Actuals', 'Giveback Reconciliation',
                          'File Manifest', 'YTD Summary']
```

Known-good numbers for `sample_data/July_1999` (useful for asserting
against): QB expense_total = 1439.34 (Accounting Expense (Quickbooks)
1257.76 + Picnic 181.58), income_total = 0; bank beginning = 32630.10,
ending = 31190.76, checks = 181.58, withdrawals = 1257.76; Givebacks =
2 items totalling 110.14. Reconciliation "Difference" should be 0.

Worth probing beyond the happy path: a month with no input folder
(`FileNotFoundError`, currently surfaces as a raw traceback — not yet
wrapped by `__main__.py`), no `--data-dir` and no configured pointer
(clean "No data folder set" message, exit 1), and re-running the same
month (should cleanly overwrite history + output, not error).

## GUI surface

Real PySide6 window rendering — use `QT_QPA_PLATFORM=offscreen` and
`widget.grab().save(path)` for screenshots; use `QTest.qWait()` after
state changes. This gives genuine rendered pixels, not a mock.

**Always isolate the data-folder pointer** so a test run can't clobber
the real user's `platformdirs` config:
```python
from pta_treasurer import config as cfg_module
cfg_module._pointer_path = lambda: Path('/tmp/verify_gui_config/data_dir.json')
```

**Drive the real file-intake code paths**, not `shutil.copytree`
bypassing the app entirely:
- Browse button: monkeypatch `PySide6.QtWidgets.QFileDialog.getOpenFileName`
  to return a fixed path, then call `win.files_box._browse('quickbooks')`.
- Drag-and-drop: build a minimal fake event object exposing `.mimeData()`
  (a real `QMimeData`-like object with `.urls()` returning
  `QUrl.fromLocalFile(path)`) and `.acceptProposedAction()`, then call
  `win.files_box.dropEvent(fake_event)` directly — this executes the
  real production handler.

**QThread background work**: after `widget.click()` on Generate Report,
pump the loop until done rather than sleeping a fixed amount:
```python
while win.last_output_path is None and waited < 10:
    app.processEvents(); time.sleep(0.05); waited += 0.05
```

**Modal dialogs will hang a headless script** — monkeypatch
`PySide6.QtWidgets.QMessageBox.critical` (and `.information`) to record
calls instead of blocking, when testing error paths (e.g. missing
`budget.xlsx` mid-session).

Worth probing beyond the happy path: double-clicking Generate rapidly
(button must disable synchronously so only one run fires), missing
`budget.xlsx` deleted mid-session (must show a dialog, not crash —
current message text leaks the internal `budget_io.generate_template()`
call, worth simplifying for a non-technical user eventually), and
"Run All Months" against a mix of valid `Month_Year` folders and a
malformed folder name (must skip the bad one, not abort the batch).

A full worked example of both surfaces (happy path + all the probes
above) lives in conversation history — reconstruct the driver script
following the patterns here if it's not already in the repo.
