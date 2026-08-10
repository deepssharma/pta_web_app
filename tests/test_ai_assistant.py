"""
Tests for ai_assistant.py. build_report_context's allow-list is the
most important test here -- it's the one enforcement point for the
privacy decision that individual transactions/check numbers/payer names
must never reach the LLM. Written as a strict allow-list (exact expected
key set + explicit negative assertions on the serialized JSON) so any
future field added to RunResult fails loudly instead of silently
leaking.
"""
import json
from pathlib import Path

from pta_treasurer.ai_assistant import (
    build_error_context, build_messages, build_report_context,
    is_ollama_available, list_ollama_models,
)
from pta_treasurer.pipeline import RunResult

EXPECTED_TOP_LEVEL_KEYS = {
    'month_label', 'org_name', 'income', 'income_total', 'expenses',
    'expense_total', 'net_income', 'bank_period', 'bank_beginning_balance',
    'bank_ending_balance', 'bank_total_deposits', 'bank_total_checks',
    'bank_total_withdrawals', 'bank_total_fees', 'givebacks_total',
    'givebacks_item_count', 'budget_vs_actual', 'warnings', 'report_filename',
}


def _make_result():
    return RunResult(
        output_path=Path('/Users/deepsharma/Documents/PTA Treasurer/output/Treasurer_Report_1999_to_2000_July_1999.xlsx'),
        warnings=['No Givebacks files found for this month'],
        qb={
            'period': 'July 1999',
            'income': {'Membership Income': 100.0},
            'income_total': 100.0,
            'expenses': {'Supplies': 50.0},
            'expense_total': 50.0,
            'net_income': 50.0,
            'transactions': [
                {'date': '2026-07-05', 'payee': 'Jane Doe', 'amount': 25.0,
                 'memo': 'membership dues', 'check_number': '1042'},
            ],
        },
        bank={
            'period': 'July 1 - July 31, 1999',
            'account': '****1234',
            'beginning_balance': 1000.0,
            'ending_balance': 1050.0,
            'total_deposits': 100.0,
            'total_checks': 30.0,
            'total_withdrawals': 20.0,
            'total_fees': 0.0,
            'deposits': [{'date': '2026-07-03', 'amount': 100.0, 'description': 'Deposit'}],
            'checks': [{'check_number': '1042', 'amount': 30.0, 'date': '2026-07-10'}],
            'withdrawals': [{'date': '2026-07-15', 'amount': 20.0}],
            'fees': [],
            'daily_balances': {'2026-07-01': 1000.0, '2026-07-31': 1050.0},
        },
        givebacks=[
            {'item': 'Book Fair', 'category': 'Fundraising', 'count': 2, 'total': 85.20,
             'source_file': 'givebacks_july_po_abc.csv'},
        ],
        income_merged={
            'Membership': {
                'Membership Income': (90.0, 100.0, [0.0] * 12),
            },
        },
        expense_merged={
            'Supplies': {
                'Office Supplies': (40.0, 50.0, [0.0] * 12),
            },
        },
    )


def test_build_report_context_has_exact_allowlisted_keys():
    result = _make_result()
    context = build_report_context(result, 'July 1999', 'Setauket School PTA')
    assert set(context.keys()) == EXPECTED_TOP_LEVEL_KEYS


def test_build_report_context_excludes_line_item_detail():
    result = _make_result()
    context = build_report_context(result, 'July 1999', 'Setauket School PTA')
    serialized = json.dumps(context)

    # Individual transaction/check/payer detail must never appear.
    assert 'Jane Doe' not in serialized
    assert 'membership dues' not in serialized
    assert '1042' not in serialized
    assert 'transactions' not in context
    assert 'deposits' not in context
    assert 'checks' not in context
    assert 'withdrawals' not in context
    assert 'fees' not in context
    assert 'daily_balances' not in context
    assert '2026-07-15' not in serialized  # a withdrawal date, not a real total
    assert '****1234' not in serialized  # bank account identifier


def test_build_report_context_excludes_full_filesystem_path():
    result = _make_result()
    context = build_report_context(result, 'July 1999', 'Setauket School PTA')
    assert context['report_filename'] == 'Treasurer_Report_1999_to_2000_July_1999.xlsx'
    assert 'deepsharma' not in json.dumps(context)
    assert '/Users/' not in json.dumps(context)


def test_build_report_context_includes_expected_aggregates():
    result = _make_result()
    context = build_report_context(result, 'July 1999', 'Setauket School PTA')
    assert context['income_total'] == 100.0
    assert context['expense_total'] == 50.0
    assert context['net_income'] == 50.0
    assert context['bank_beginning_balance'] == 1000.0
    assert context['bank_ending_balance'] == 1050.0
    assert context['givebacks_total'] == 85.20
    assert context['givebacks_item_count'] == 1


def test_build_report_context_flattens_budget_vs_actual_by_section():
    result = _make_result()
    context = build_report_context(result, 'July 1999', 'Setauket School PTA')
    bva = context['budget_vs_actual']
    assert bva['income']['Membership']['Membership Income'] == {
        'budget': 100.0, 'actual_this_month': 0.0,
    }
    assert bva['expenses']['Supplies']['Office Supplies'] == {
        'budget': 50.0, 'actual_this_month': 0.0,
    }
    # raw (last_yr, budget, actuals) tuples must not leak through
    assert '90.0' not in str(bva)


def test_build_error_context_wraps_message_only():
    context = build_error_context('No bank statement PDF found in input/July_2026')
    assert context == {'error_message': 'No bank statement PDF found in input/July_2026'}


def test_build_messages_appends_question_with_context():
    context = {'income_total': 100.0}
    messages = build_messages(context, 'How much income this month?')
    assert len(messages) == 1
    assert messages[0]['role'] == 'user'
    assert 'How much income this month?' in messages[0]['content']
    assert '100.0' in messages[0]['content']


def test_build_messages_preserves_history():
    history = [{'role': 'user', 'content': 'first question'},
               {'role': 'assistant', 'content': 'first answer'}]
    messages = build_messages({}, 'second question', history=history)
    assert len(messages) == 3
    assert messages[0] == history[0]
    assert messages[1] == history[1]
    assert 'second question' in messages[2]['content']


# Port 1 is a reserved/unroutable port -- connections there fail nearly
# instantly, so these behave like fast unit tests despite touching httpx.

def test_is_ollama_available_false_when_unreachable():
    assert is_ollama_available('http://localhost:1') is False


def test_list_ollama_models_empty_when_unreachable():
    assert list_ollama_models('http://localhost:1') == []
