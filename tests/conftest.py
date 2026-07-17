"""
Shared test fixtures for all test files.
Fixtures are automatically available to all tests without importing.
"""
import pytest
import tempfile
import os
from pathlib import Path


@pytest.fixture
def sample_qb_csv(tmp_path):
    """Creates a minimal QuickBooks Transaction Detail CSV for testing."""
    content = '''Setauket School PTA
Transaction Detail by Account
July 1-31, 2025
,,,,,,,,
,,Transaction date,Transaction type,Num,Name,Description,Split,Amount
Checking (4346),,,,,,,,
,,07/01/2025,Check,1077,Michelle Schultz,6th Grade Events,Checking,-181.58
Total for Checking (4346),,,,,,,,
Accounting Expense (Quickbooks),,,,,,,,
,,07/17/2025,Expense,,Quickbooks Online,Monthly Fee,Checking,-1257.76
Total for Accounting Expense (Quickbooks),,,,,,,,
'''
    f = tmp_path / 'quickbooks_july_2025.csv'
    f.write_text(content, encoding='utf-8')
    return f


@pytest.fixture
def sample_givebacks_csv(tmp_path):
    """Creates a minimal Givebacks CSV for testing."""
    content = '''Item,Categories,No. of Transactions,Total
Shop to Give Donation,,1,$95.14
Teacher/Staff,Memberships,1,$15.00
'''
    folder = tmp_path / 'givebacks'
    folder.mkdir()
    f = folder / 'givebacks_july_po_123.csv'
    f.write_text(content, encoding='utf-8')
    return f, folder


@pytest.fixture
def sample_merged_income():
    """Sample income_merged structure for builder testing."""
    return {
        'Fundraising': {
            'Book Fair': (9118.36, 500.0, [0.0]*12),
            'FAST':      (26370.0, 1000.0, [0.0, 0.0, 0.0, 0.0, 0.0, 0.0,
                                             0.0, 0.0, 4340.0, 3360.0, 0.0, 0.0]),
        },
        'Membership': {
            'Teachers': (0.0, 165.0, [0.0]*12),
        }
    }


@pytest.fixture
def sample_merged_expense():
    """Sample expense_merged structure for builder testing."""
    return {
        'Admin/General': {
            'Accounting Quickbooks': (410.66, 1300.0,
                                      [1257.76, 0.0, 0.0, 0.0, 0.0, 0.0,
                                       0.0, 0.0, 0.0, 0.0, 0.0, 0.0]),
            'Bank Services':         (234.94, 200.0,
                                      [0.0, 0.0, 0.0, 0.0, 0.0, 0.0,
                                       0.0, 0.0, 0.0, 30.5, 0.0, 0.0]),
        }
    }
