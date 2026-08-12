"""
chat_panel.py
Persistent side-panel chatbot: answers questions about a generated
report, or helps troubleshoot the last input-file error. Opt-in --
shows a "not detected" banner if the local Ollama server isn't running.
"""
import httpx
from PySide6.QtCore import QThread, Signal
from PySide6.QtWidgets import (
    QHBoxLayout, QLabel, QLineEdit, QPlainTextEdit, QPushButton, QVBoxLayout,
    QWidget,
)


class ChatWorker(QThread):
    finished_ok = Signal(str)
    failed = Signal(str)

    def __init__(self, question, context, mode, host, model, history):
        super().__init__()
        self.question = question
        self.context = context
        self.mode = mode
        self.host = host
        self.model = model
        self.history = history

    def run(self):
        try:
            from pta_treasurer.ai_assistant import ask
            answer = ask(self.question, self.context, self.mode,
                          self.host, self.model, history=self.history)
        except httpx.ConnectError:
            self.failed.emit(
                f"Couldn't reach Ollama at {self.host} — make sure it's "
                "running ('ollama serve' or the Ollama app).")
        except httpx.HTTPStatusError as e:
            self.failed.emit(
                f'Ollama returned an error ({e.response.status_code}) — '
                f"check that the model '{self.model}' is pulled "
                f"('ollama pull {self.model}').")
        except Exception as e:
            self.failed.emit(str(e))
        else:
            self.finished_ok.emit(answer)


class ChatPanel(QWidget):
    open_settings_requested = Signal()

    def __init__(self, config, data_dir, parent=None):
        super().__init__(parent)
        self.config = config
        self.data_dir = data_dir
        self.mode = None
        self.context = None
        self.history = []
        self.worker = None
        self.run_result = None

        # Budget-edit mode state: pending_action holds the last parsed,
        # not-yet-applied edit (see ai_assistant.parse_edit_action) --
        # nothing in budget.xlsx changes until the user clicks Apply.
        self.edit_mode = False
        self.pending_action = None
        self._saved_mode = None
        self._saved_context = None
        self._saved_history = []

        self.transcript = QPlainTextEdit(readOnly=True)
        self.transcript.setPlaceholderText(
            'Generate a report, then ask questions about it here.')

        self.status_label = QLabel()
        self.status_label.setWordWrap(True)
        self.open_settings_btn = QPushButton('Open Settings…')
        self.open_settings_btn.clicked.connect(self.open_settings_requested.emit)

        self.apply_btn = QPushButton('Apply')
        self.apply_btn.clicked.connect(self._apply_pending)
        self.apply_btn.setVisible(False)
        self.discard_btn = QPushButton('Discard')
        self.discard_btn.clicked.connect(self._discard_pending)
        self.discard_btn.setVisible(False)
        self.edit_mode_btn = QPushButton('Edit Budget…')
        self.edit_mode_btn.clicked.connect(self._toggle_edit_mode)

        self.input_edit = QLineEdit()
        self.input_edit.setPlaceholderText('Ask a question…')
        self.input_edit.returnPressed.connect(self._send)
        self.send_btn = QPushButton('Send')
        self.send_btn.clicked.connect(self._send)

        input_row = QHBoxLayout()
        input_row.addWidget(self.input_edit, 1)
        input_row.addWidget(self.send_btn)

        action_row = QHBoxLayout()
        action_row.addWidget(self.apply_btn)
        action_row.addWidget(self.discard_btn)
        action_row.addStretch(1)
        action_row.addWidget(self.edit_mode_btn)

        layout = QVBoxLayout()
        layout.addWidget(self.transcript, 1)
        layout.addWidget(self.status_label)
        layout.addWidget(self.open_settings_btn)
        layout.addLayout(action_row)
        layout.addLayout(input_row)
        self.setLayout(layout)

        self.refresh()

    def refresh(self):
        """Re-checks whether Ollama is reachable. Call after Settings
        changes, since the host/model may have just changed."""
        from pta_treasurer.ai_assistant import is_ollama_available
        available = is_ollama_available(self.config.ollama_host)

        self.status_label.setVisible(not available)
        self.open_settings_btn.setVisible(not available)
        self.input_edit.setEnabled(available)
        self.send_btn.setEnabled(available)
        if not available:
            self.status_label.setText(
                f'Ollama not detected at {self.config.ollama_host}. '
                'Start it, then adjust the host/model in Settings if needed.')

    def set_context(self, report=None, error=None, month_label=None):
        """report: a RunResult from a successful run. error: an error
        message string from a failed run. Exactly one should be given."""
        from pta_treasurer.ai_assistant import build_error_context, build_report_context

        if self.edit_mode:
            # The report/actuals underneath edit mode's data just
            # changed -- drop the pending edit and exit rather than risk
            # applying it against a now-stale picture.
            self._exit_edit_mode(announce=False)

        self.history = []
        self.run_result = report
        if report is not None:
            self.mode = 'report'
            self.context = build_report_context(report, month_label, self.config.org_name)
            self.transcript.appendPlainText(
                f'\n— New report loaded ({month_label}). Ask me anything about it. —')
        elif error is not None:
            self.mode = 'troubleshoot'
            self.context = build_error_context(error)
            self.transcript.appendPlainText(
                '\n— Report generation failed. Ask me for help figuring out why. —')
        else:
            self.mode = None
            self.context = None

    def _toggle_edit_mode(self):
        if self.edit_mode:
            self._exit_edit_mode()
            return

        budget_path = self.data_dir / 'budget.xlsx'
        if not budget_path.exists():
            self.transcript.appendPlainText(
                '\nNo budget.xlsx found yet — open Settings to generate one first.')
            return
        from pta_treasurer.ai_assistant import build_budget_context
        from pta_treasurer.budget_io import load_budget
        try:
            income_budget, expense_budget, _qb_map = load_budget(budget_path)
        except Exception as e:
            self.transcript.appendPlainText(f'\nCould not read budget.xlsx: {e}')
            return

        self._saved_mode, self._saved_context, self._saved_history = (
            self.mode, self.context, self.history)
        self.edit_mode = True
        self.mode = 'edit_budget'
        self.context = build_budget_context(income_budget, expense_budget)
        self.history = []
        self.pending_action = None
        self.apply_btn.setVisible(False)
        self.discard_btn.setVisible(False)
        self.edit_mode_btn.setText('Exit Budget Editing')
        self.transcript.appendPlainText(
            '\n— Budget editing mode. Describe a change (e.g. "add an '
            'expense category called Robotics Club under Programs, budget '
            '$500") and I\'ll propose an edit for you to review before '
            "it's applied. —")

    def _exit_edit_mode(self, announce=True):
        self.edit_mode = False
        self.mode, self.context, self.history = (
            self._saved_mode, self._saved_context, self._saved_history)
        self.pending_action = None
        self.apply_btn.setVisible(False)
        self.discard_btn.setVisible(False)
        self.edit_mode_btn.setText('Edit Budget…')
        if announce:
            self.transcript.appendPlainText('\n— Exited budget editing. —')

    def _send(self):
        question = self.input_edit.text().strip()
        if not question or self.context is None:
            return

        self.input_edit.clear()
        self.input_edit.setEnabled(False)
        self.send_btn.setEnabled(False)
        self.transcript.appendPlainText(f'\nYou: {question}')

        self.worker = ChatWorker(
            question, self.context, self.mode,
            self.config.ollama_host, self.config.ollama_model, self.history)
        self.worker.finished_ok.connect(lambda answer: self._on_answer(question, answer))
        self.worker.failed.connect(self._on_failure)
        self.worker.start()

    def _on_answer(self, question, answer):
        self.input_edit.setEnabled(True)
        self.send_btn.setEnabled(True)
        self.input_edit.setFocus()
        self.history.append({'role': 'user', 'content': question})
        self.history.append({'role': 'assistant', 'content': answer})

        if self.edit_mode:
            from pta_treasurer.ai_assistant import parse_edit_action
            from pta_treasurer.budget_io import describe_edit
            action = parse_edit_action(answer)
            if action['action'] == 'clarify':
                self.transcript.appendPlainText(f"Assistant: {action['message']}")
                self.pending_action = None
                self.apply_btn.setVisible(False)
                self.discard_btn.setVisible(False)
                return
            self.pending_action = action
            self.transcript.appendPlainText(
                f'Assistant proposes: {describe_edit(action)}\n'
                f'Click Apply to make this change in budget.xlsx, or Discard to cancel.')
            self.apply_btn.setVisible(True)
            self.discard_btn.setVisible(True)
            return

        self.transcript.appendPlainText(f'Assistant: {answer}')

    def _discard_pending(self):
        self.pending_action = None
        self.apply_btn.setVisible(False)
        self.discard_btn.setVisible(False)
        self.transcript.appendPlainText('Discarded — budget.xlsx was not changed.')

    def _apply_pending(self):
        if not self.pending_action:
            return
        from pta_treasurer.budget_io import actuals_total_for_item, apply_edit
        action = self.pending_action
        budget_path = self.data_dir / 'budget.xlsx'

        current_actuals_total = None
        if action['action'] == 'remove_item' and self.run_result is not None:
            merged = (self.run_result.income_merged if action['sheet'] == 'Income Budget'
                      else self.run_result.expense_merged)
            current_actuals_total = actuals_total_for_item(merged, action['item'])

        try:
            apply_edit(budget_path, action, current_actuals_total=current_actuals_total)
        except ValueError as e:
            self.transcript.appendPlainText(f'FAILED: {e}')
            return

        self.transcript.appendPlainText(
            'Applied — budget.xlsx updated. Regenerate the report to see it reflected.')
        self.pending_action = None
        self.apply_btn.setVisible(False)
        self.discard_btn.setVisible(False)

        # Refresh the edit-mode context so a follow-up edit in the same
        # session sees the new structure, not the pre-edit one.
        from pta_treasurer.ai_assistant import build_budget_context
        from pta_treasurer.budget_io import load_budget
        income_budget, expense_budget, _qb_map = load_budget(budget_path)
        self.context = build_budget_context(income_budget, expense_budget)

    def _on_failure(self, message):
        self.transcript.appendPlainText(f'FAILED: {message}')
        self.input_edit.setEnabled(True)
        self.send_btn.setEnabled(True)
