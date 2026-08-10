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

    def __init__(self, config, parent=None):
        super().__init__(parent)
        self.config = config
        self.mode = None
        self.context = None
        self.history = []
        self.worker = None

        self.transcript = QPlainTextEdit(readOnly=True)
        self.transcript.setPlaceholderText(
            'Generate a report, then ask questions about it here.')

        self.status_label = QLabel()
        self.status_label.setWordWrap(True)
        self.open_settings_btn = QPushButton('Open Settings…')
        self.open_settings_btn.clicked.connect(self.open_settings_requested.emit)

        self.input_edit = QLineEdit()
        self.input_edit.setPlaceholderText('Ask a question…')
        self.input_edit.returnPressed.connect(self._send)
        self.send_btn = QPushButton('Send')
        self.send_btn.clicked.connect(self._send)

        input_row = QHBoxLayout()
        input_row.addWidget(self.input_edit, 1)
        input_row.addWidget(self.send_btn)

        layout = QVBoxLayout()
        layout.addWidget(self.transcript, 1)
        layout.addWidget(self.status_label)
        layout.addWidget(self.open_settings_btn)
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

        self.history = []
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
        self.transcript.appendPlainText(f'Assistant: {answer}')
        self.history.append({'role': 'user', 'content': question})
        self.history.append({'role': 'assistant', 'content': answer})
        self.input_edit.setEnabled(True)
        self.send_btn.setEnabled(True)
        self.input_edit.setFocus()

    def _on_failure(self, message):
        self.transcript.appendPlainText(f'FAILED: {message}')
        self.input_edit.setEnabled(True)
        self.send_btn.setEnabled(True)
