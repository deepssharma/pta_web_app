"""
Tests for config.py — OrgConfig persistence, data folder pointer, and
month/file detection helpers.
"""
from pathlib import Path

import pytest

from pta_treasurer.config import (
    OrgConfig, calendar_to_fiscal, detect_month_from_filename,
    get_data_dir, load_config, month_name_to_fiscal_index, save_config,
    set_data_dir,
)


# ── Fiscal month math ─────────────────────────────────────────────────────

def test_calendar_to_fiscal_july_is_zero():
    assert calendar_to_fiscal(7) == 0


def test_calendar_to_fiscal_june_is_eleven():
    assert calendar_to_fiscal(6) == 11


def test_month_name_to_fiscal_index():
    assert month_name_to_fiscal_index('July') == 0
    assert month_name_to_fiscal_index('december') == 5
    assert month_name_to_fiscal_index('June') == 11


def test_month_name_to_fiscal_index_unknown_returns_none():
    assert month_name_to_fiscal_index('Smarch') is None


# ── Filename month detection ──────────────────────────────────────────────

def test_detect_month_from_filename_with_year():
    label, idx = detect_month_from_filename(Path('quickbooks_july_2025.csv'))
    assert label == 'July 2025'
    assert idx == 0


def test_detect_month_from_filename_four_digit_year_before_2000():
    # sample_data uses a synthetic 1999 date — must not be excluded by an
    # over-narrow "20xx" year regex.
    label, idx = detect_month_from_filename(Path('Chase_july_1999.pdf'))
    assert label == 'July 1999'
    assert idx == 0


def test_detect_month_from_filename_no_year():
    label, idx = detect_month_from_filename(Path('givebacks_july_po_demo.csv'))
    assert label == 'July'
    assert idx == 0


def test_detect_month_from_filename_no_month_found():
    label, idx = detect_month_from_filename(Path('mystery_file.csv'))
    assert label is None
    assert idx is None


# ── Data folder pointer ───────────────────────────────────────────────────

def test_get_data_dir_returns_none_before_first_run(tmp_path, monkeypatch):
    monkeypatch.setattr(
        'pta_treasurer.config._pointer_path', lambda: tmp_path / 'nonexistent' / 'data_dir.json'
    )
    assert get_data_dir() is None


def test_set_then_get_data_dir_roundtrip(tmp_path, monkeypatch):
    pointer_path = tmp_path / 'data_dir.json'
    monkeypatch.setattr('pta_treasurer.config._pointer_path', lambda: pointer_path)
    set_data_dir(tmp_path / 'MyData')
    assert get_data_dir() == tmp_path / 'MyData'


# ── OrgConfig persistence ─────────────────────────────────────────────────

def test_save_and_load_config_roundtrip(tmp_path):
    config = OrgConfig(org_name='Demo School PTA', fiscal_start_month=7, balance_forward=1234.56)
    save_config(config, tmp_path)
    loaded = load_config(tmp_path)
    assert loaded == config


def test_load_config_missing_raises(tmp_path):
    with pytest.raises(FileNotFoundError):
        load_config(tmp_path)
