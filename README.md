# PTA Treasurer (desktop app)

A cross-platform desktop app for generating monthly PTA/school treasurer
reports (income/expense vs. budget, bank reconciliation, YTD summary) from
QuickBooks, Chase bank statement, and Givebacks exports — no Python or
spreadsheet-formula knowledge required to run it.

This is the generic, installable successor to the notebook-based
[`pta_treasurer`](../Desktop/pta/codes/pta_treasurer) project: same core
parsing/report-building logic, wrapped in a real GUI and packaged as a
double-clickable app instead of requiring Jupyter.

**Status: scaffolded, not yet implemented.** `src/pta_treasurer/parsers.py`
and `builders.py` are ported and tested; everything else (config, budget
import, the pipeline that ties parsing + building together, the GUI, and
packaging) is still to be built. See [`PLAN.md`](PLAN.md) for the full
design and build order.

## Development setup

```sh
python -m venv .venv && source .venv/bin/activate
pip install -e ".[dev]"
pytest
```

## Real financial data

Real input files, generated reports, and history are never committed —
`input/`, `output/`, `data/`, and `logs/` are gitignored. Only synthetic
data lives in git: `tests/fixtures/` (unit test fixtures) and
`sample_data/July_1999/` (a full synthetic demo month — fake org, fake year,
no real names/accounts — for trying the whole pipeline safely).
