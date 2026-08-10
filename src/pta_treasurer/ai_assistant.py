"""
ai_assistant.py
LLM-powered chatbot: answers questions about a generated report, and
helps troubleshoot input-file problems. Opt-in -- runs entirely against
a local Ollama server (http://localhost:11434 by default), so no API
key/secret and no financial data ever leaves the machine at all -- a
stricter privacy story than the aggregated-totals-only boundary this
module also enforces for the LLM prompt itself.

Privacy boundary: build_report_context() is a strict allow-list of
aggregate figures only (category totals, budget-vs-actual, bank
beginning/ending balances). It never includes individual transactions,
check numbers, payer names, or per-item Givebacks detail. Don't add a
new field here without checking it's a total, not a line item.
"""
import json

import httpx

from pta_treasurer.config import month_name_to_fiscal_index
from pta_treasurer.pipeline import RunResult

SYSTEM_PROMPT_REPORT = (
    'You are a helpful assistant for a PTA/school treasurer using the '
    'pta-treasurer app. You are answering questions about a monthly '
    'treasurer report that was just generated. You only have access to '
    'aggregated totals, not individual transactions -- if asked about a '
    'specific transaction, check number, or payer name, explain that '
    'level of detail is not available to you and suggest opening the '
    'generated Excel report to look it up. Be concise and precise with '
    'dollar figures.'
)

SYSTEM_PROMPT_TROUBLESHOOT = (
    'You are a helpful assistant for a PTA/school treasurer using the '
    'pta-treasurer app. Report generation just failed with the error '
    'message given to you. Explain in plain language (no Python jargon) '
    'what likely went wrong and what specific action would fix it -- '
    'e.g. which file to check, or how a filename/month mismatch should '
    'be corrected. Be concise.'
)


def _budget_vs_actual(merged: dict, fiscal_idx: int | None) -> dict:
    out = {}
    for section, items in merged.items():
        out[section] = {}
        for item, (_last_yr, budget, actuals) in items.items():
            actual = actuals[fiscal_idx] if fiscal_idx is not None else None
            out[section][item] = {'budget': budget, 'actual_this_month': actual}
    return out


def build_report_context(result: RunResult, month_label: str, org_name: str) -> dict:
    """Pure. Strict allow-list of aggregate fields only -- see this
    module's docstring for the privacy boundary this enforces."""
    fiscal_idx = month_name_to_fiscal_index(month_label.split()[0])
    givebacks_total = sum(g.get('total', 0.0) for g in result.givebacks)

    return {
        'month_label': month_label,
        'org_name': org_name,
        'income': dict(result.qb.get('income', {})),
        'income_total': result.qb.get('income_total', 0.0),
        'expenses': dict(result.qb.get('expenses', {})),
        'expense_total': result.qb.get('expense_total', 0.0),
        'net_income': result.qb.get('net_income', 0.0),
        'bank_period': result.bank.get('period', ''),
        'bank_beginning_balance': result.bank.get('beginning_balance', 0.0),
        'bank_ending_balance': result.bank.get('ending_balance', 0.0),
        'bank_total_deposits': result.bank.get('total_deposits', 0.0),
        'bank_total_checks': result.bank.get('total_checks', 0.0),
        'bank_total_withdrawals': result.bank.get('total_withdrawals', 0.0),
        'bank_total_fees': result.bank.get('total_fees', 0.0),
        'givebacks_total': givebacks_total,
        'givebacks_item_count': len(result.givebacks),
        'budget_vs_actual': {
            'income': _budget_vs_actual(result.income_merged, fiscal_idx),
            'expenses': _budget_vs_actual(result.expense_merged, fiscal_idx),
        },
        'warnings': list(result.warnings),
        'report_filename': result.output_path.name,
    }


def build_error_context(error_message: str) -> dict:
    """Pure. Wraps the caught exception's str() -- no file reads, no
    re-parsing."""
    return {'error_message': str(error_message)}


def build_messages(context: dict, question: str, history: list[dict] | None = None) -> list[dict]:
    """Pure. Assembles the chat `messages` list (role/content dicts --
    same shape Ollama's /api/chat expects)."""
    messages = list(history or [])
    messages.append({
        'role': 'user',
        'content': f'Context (JSON):\n{json.dumps(context, indent=2)}\n\nQuestion: {question}',
    })
    return messages


def is_ollama_available(host: str) -> bool:
    """Quick reachability check (local call, short timeout) -- used to
    show/hide the chat panel's 'Ollama not detected' banner."""
    try:
        r = httpx.get(f'{host.rstrip("/")}/api/tags', timeout=2.0)
        return r.status_code == 200
    except httpx.HTTPError:
        return False


def list_ollama_models(host: str) -> list[str]:
    """Names of locally pulled models, for the Settings model picker.
    Returns an empty list if Ollama isn't reachable."""
    try:
        r = httpx.get(f'{host.rstrip("/")}/api/tags', timeout=2.0)
        r.raise_for_status()
        return [m['name'] for m in r.json().get('models', [])]
    except httpx.HTTPError:
        return []


def ask(question: str, context: dict, mode: str, host: str, model: str,
        history: list[dict] | None = None) -> str:
    """Impure. POSTs to a local Ollama server's /api/chat. Raises
    httpx.HTTPError (connection refused, timeout, bad model name, etc.)
    on failure -- ChatWorker turns that into a plain FAILED line in the
    transcript, no crash."""
    system_prompt = SYSTEM_PROMPT_TROUBLESHOOT if mode == 'troubleshoot' else SYSTEM_PROMPT_REPORT
    messages = [{'role': 'system', 'content': system_prompt}] + build_messages(context, question, history)

    response = httpx.post(
        f'{host.rstrip("/")}/api/chat',
        json={'model': model, 'messages': messages, 'stream': False},
        timeout=120.0,
    )
    response.raise_for_status()
    return response.json()['message']['content']
