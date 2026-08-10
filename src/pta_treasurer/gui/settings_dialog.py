"""
settings_dialog.py
Edit org config, jump to the data folder/budget template, and configure
the opt-in Givebacks auto-download (Phase 4) and AI Assistant (local
Ollama, no credentials involved). Batch actions ("Run All Months", "Build
Debits & Credits Ledger") live on the main window, not here -- see
main_window.py.
"""
from pathlib import Path

from PySide6.QtCore import QThread, Signal
from PySide6.QtWidgets import (
    QComboBox, QDialog, QDialogButtonBox, QDoubleSpinBox, QFileDialog,
    QFormLayout, QGroupBox, QLineEdit, QMessageBox, QPlainTextEdit,
    QPushButton, QVBoxLayout,
)

from pta_treasurer.budget_io import generate_template
from pta_treasurer.config import OrgConfig, save_config, set_data_dir
from pta_treasurer.credentials import (
    NoKeyringBackendError, get_givebacks_password, set_givebacks_password,
)
from pta_treasurer.gui.utils import open_in_default_app


class ChromiumInstallWorker(QThread):
    line = Signal(str)
    finished_ok = Signal()
    failed = Signal(str)

    def run(self):
        from pta_treasurer.givebacks_download import install_chromium
        try:
            install_chromium(progress_cb=self.line.emit)
        except Exception as e:
            self.failed.emit(str(e))
        else:
            self.finished_ok.emit()


class SettingsDialog(QDialog):
    def __init__(self, config: OrgConfig, data_dir: Path, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Settings')
        self.config = config
        self.data_dir = Path(data_dir)

        self.org_name_edit = QLineEdit(config.org_name)
        self.balance_spin = QDoubleSpinBox()
        self.balance_spin.setRange(-1_000_000, 1_000_000)
        self.balance_spin.setDecimals(2)
        self.balance_spin.setPrefix('$')
        self.balance_spin.setValue(config.balance_forward)

        form = QFormLayout()
        form.addRow('Organization name:', self.org_name_edit)
        form.addRow('Balance forward:', self.balance_spin)

        # ── Givebacks (optional auto-download) ──────────────────────────
        self.gb_url_edit = QLineEdit(config.givebacks_org_url)
        self.gb_url_edit.setPlaceholderText('https://your-org.givebacks.com')
        self.gb_email_edit = QLineEdit(config.givebacks_email)
        self.gb_cause_id_edit = QLineEdit(config.givebacks_cause_id)
        self.gb_cause_id_edit.setPlaceholderText('Usually auto-detected — leave blank')

        self.gb_password_edit = QLineEdit()
        self.gb_password_edit.setEchoMode(QLineEdit.Password)
        self._refresh_password_placeholder()
        self.gb_email_edit.textChanged.connect(self._refresh_password_placeholder)

        gb_remove_btn = QPushButton('Remove Saved Password')
        gb_remove_btn.clicked.connect(self._remove_givebacks_password)

        gb_form = QFormLayout()
        gb_form.addRow('Org URL:', self.gb_url_edit)
        gb_form.addRow('Email:', self.gb_email_edit)
        gb_form.addRow('Cause ID (optional):', self.gb_cause_id_edit)
        gb_form.addRow('Password:', self.gb_password_edit)

        self.gb_install_btn = QPushButton('Install Browser Automation…')
        self.gb_install_btn.clicked.connect(self._install_chromium)
        self.gb_install_log = QPlainTextEdit(readOnly=True)
        self.gb_install_log.setFixedHeight(80)
        self.gb_install_log.hide()

        gb_box = QGroupBox('Givebacks Auto-Download (optional)')
        gb_layout = QVBoxLayout()
        gb_layout.addLayout(gb_form)
        gb_layout.addWidget(gb_remove_btn)
        gb_layout.addWidget(self.gb_install_btn)
        gb_layout.addWidget(self.gb_install_log)
        gb_box.setLayout(gb_layout)

        # ── AI Assistant (optional chatbot, runs on local Ollama) ────────
        self.ollama_host_edit = QLineEdit(config.ollama_host)
        self.ollama_model_combo = QComboBox()
        self.ollama_model_combo.setEditable(True)
        self.ollama_model_combo.setEditText(config.ollama_model)
        self._refresh_ollama_models()

        ai_refresh_btn = QPushButton('Refresh Available Models')
        ai_refresh_btn.clicked.connect(self._refresh_ollama_models)

        ai_form = QFormLayout()
        ai_form.addRow('Ollama host:', self.ollama_host_edit)
        ai_form.addRow('Model:', self.ollama_model_combo)

        ai_box = QGroupBox('AI Assistant (optional — runs on local Ollama, no API key needed)')
        ai_layout = QVBoxLayout()
        ai_layout.addLayout(ai_form)
        ai_layout.addWidget(ai_refresh_btn)
        ai_box.setLayout(ai_layout)

        # ── Pass-through fund (optional, e.g. a fundraiser whose money
        # moves through the org's bank account but isn't the org's own) ──
        self.pt_name_edit = QLineEdit(config.pass_through_fund_name)
        self.pt_name_edit.setPlaceholderText('e.g. READTHON — leave blank to disable')
        self.pt_categories_edit = QLineEdit(config.pass_through_fund_categories)
        self.pt_categories_edit.setPlaceholderText('Comma-separated QuickBooks category names')
        self.pt_balance_spin = QDoubleSpinBox()
        self.pt_balance_spin.setRange(-1_000_000, 1_000_000)
        self.pt_balance_spin.setDecimals(2)
        self.pt_balance_spin.setPrefix('$')
        self.pt_balance_spin.setValue(config.pass_through_fund_balance_forward)

        pt_form = QFormLayout()
        pt_form.addRow('Fund name:', self.pt_name_edit)
        pt_form.addRow('QuickBooks categories:', self.pt_categories_edit)
        pt_form.addRow('Opening balance held:', self.pt_balance_spin)

        pt_box = QGroupBox('Pass-Through Fund (optional)')
        pt_layout = QVBoxLayout()
        pt_layout.addLayout(pt_form)
        pt_box.setLayout(pt_layout)

        edit_budget_btn = QPushButton('Edit Budget…')
        edit_budget_btn.clicked.connect(self._edit_budget)

        change_folder_btn = QPushButton('Change Data Folder…')
        change_folder_btn.clicked.connect(self._change_data_folder)

        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self._save)
        buttons.rejected.connect(self.reject)

        layout = QVBoxLayout()
        layout.addLayout(form)
        layout.addWidget(gb_box)
        layout.addWidget(ai_box)
        layout.addWidget(pt_box)
        layout.addWidget(edit_budget_btn)
        layout.addWidget(change_folder_btn)
        layout.addWidget(buttons)
        self.setLayout(layout)

    def _refresh_password_placeholder(self):
        email = self.gb_email_edit.text().strip()
        has_saved = False
        if email:
            try:
                has_saved = bool(get_givebacks_password(email))
            except NoKeyringBackendError:
                pass
        self.gb_password_edit.setPlaceholderText(
            '•••••• (saved — leave blank to keep)' if has_saved else 'Not set')

    def _remove_givebacks_password(self):
        email = self.gb_email_edit.text().strip()
        if not email:
            return
        from pta_treasurer.credentials import clear_givebacks_password
        try:
            clear_givebacks_password(email)
        except NoKeyringBackendError as e:
            QMessageBox.critical(self, 'Could not remove password', str(e))
            return
        self.gb_password_edit.clear()
        self._refresh_password_placeholder()
        QMessageBox.information(self, 'Password removed', f'Saved password for {email} removed.')

    def _refresh_ollama_models(self):
        from pta_treasurer.ai_assistant import list_ollama_models
        current_text = self.ollama_model_combo.currentText()
        models = list_ollama_models(self.ollama_host_edit.text().strip())
        self.ollama_model_combo.clear()
        self.ollama_model_combo.addItems(models)
        self.ollama_model_combo.setEditText(current_text)

    def _install_chromium(self):
        self.gb_install_btn.setEnabled(False)
        self.gb_install_log.show()
        self.gb_install_log.clear()
        self._chromium_worker = ChromiumInstallWorker()
        self._chromium_worker.line.connect(self.gb_install_log.appendPlainText)
        self._chromium_worker.finished_ok.connect(self._on_chromium_installed)
        self._chromium_worker.failed.connect(self._on_chromium_install_failed)
        self._chromium_worker.start()

    def _on_chromium_installed(self):
        self.gb_install_btn.setEnabled(True)
        self.gb_install_log.appendPlainText('Done.')

    def _on_chromium_install_failed(self, message):
        self.gb_install_btn.setEnabled(True)
        self.gb_install_log.appendPlainText(f'FAILED: {message}')

    def _save(self):
        self.config.org_name = self.org_name_edit.text().strip()
        self.config.balance_forward = self.balance_spin.value()
        self.config.givebacks_org_url = self.gb_url_edit.text().strip()
        self.config.givebacks_email = self.gb_email_edit.text().strip()
        self.config.givebacks_cause_id = self.gb_cause_id_edit.text().strip()

        new_password = self.gb_password_edit.text()
        if new_password and self.config.givebacks_email:
            try:
                set_givebacks_password(self.config.givebacks_email, new_password)
            except NoKeyringBackendError as e:
                QMessageBox.critical(self, 'Could not save password', str(e))
                return

        self.config.ollama_host = self.ollama_host_edit.text().strip()
        self.config.ollama_model = self.ollama_model_combo.currentText().strip()

        self.config.pass_through_fund_name = self.pt_name_edit.text().strip()
        self.config.pass_through_fund_categories = self.pt_categories_edit.text().strip()
        self.config.pass_through_fund_balance_forward = self.pt_balance_spin.value()

        save_config(self.config, self.data_dir)
        self.accept()

    def _edit_budget(self):
        budget_path = self.data_dir / 'budget.xlsx'
        if not budget_path.exists():
            generate_template(budget_path)
        open_in_default_app(budget_path)

    def _change_data_folder(self):
        chosen = QFileDialog.getExistingDirectory(self, 'Choose a data folder', str(self.data_dir))
        if not chosen:
            return
        set_data_dir(Path(chosen))
        QMessageBox.information(
            self, 'Data folder changed',
            'Restart the app for the new data folder to take effect.'
        )

