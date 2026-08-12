"""
main_window.py
Month/year picker, input-file status with browse/drag-drop, "Generate
Report" (run on a background thread so the UI stays responsive), and
Open Report/Open Folder actions on success.
"""
import queue
import shutil
from datetime import date
from pathlib import Path

from PySide6.QtCore import QThread, Qt, Signal
from PySide6.QtWidgets import (
    QComboBox, QDockWidget, QFileDialog, QFormLayout, QGroupBox, QHBoxLayout,
    QInputDialog, QLabel, QLineEdit, QMainWindow, QMessageBox,
    QPlainTextEdit, QProgressBar, QPushButton, QSpinBox, QVBoxLayout,
    QWidget,
)

from pta_treasurer.config import OrgConfig, load_config
from pta_treasurer.gui.chat_panel import ChatPanel
from pta_treasurer.gui.settings_dialog import SettingsDialog
from pta_treasurer.gui.utils import open_in_default_app
from pta_treasurer.pipeline import list_available_fiscal_years, run_month

MONTH_LABELS = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December',
]


class RunWorker(QThread):
    finished_ok = Signal(object)
    failed = Signal(str)

    def __init__(self, config, data_dir, month, year):
        super().__init__()
        self.config = config
        self.data_dir = data_dir
        self.month = month
        self.year = year

    def run(self):
        try:
            result = run_month(self.config, self.data_dir, self.month, self.year)
        except Exception as e:
            self.failed.emit(str(e))
        else:
            self.finished_ok.emit(result)


class GivebacksWorker(QThread):
    """Runs the opt-in Givebacks auto-download (Phase 4) on a background
    thread. OTP entry needs a modal dialog, which can only be shown on
    the main thread -- otp_requested is emitted (Qt auto-queues it across
    threads) and _otp_queue.get() blocks this thread until MainWindow's
    slot puts the entered code back."""
    finished_ok = Signal(list)
    failed = Signal(str)
    otp_requested = Signal()

    def __init__(self, month_label, config, dest_folder, data_dir):
        super().__init__()
        self.month_label = month_label
        self.config = config
        self.dest_folder = dest_folder
        self.data_dir = data_dir
        self._otp_queue = queue.Queue()

    def submit_otp(self, code: str) -> None:
        self._otp_queue.put(code)

    def run(self):
        import asyncio

        def otp_prompt():
            self.otp_requested.emit()
            return self._otp_queue.get()

        try:
            from pta_treasurer.givebacks_download import download_givebacks
            saved = asyncio.run(download_givebacks(
                self.month_label, self.config, self.dest_folder, self.data_dir,
                otp_prompt=otp_prompt))
        except ImportError:
            self.failed.emit(
                "Givebacks auto-download needs the 'givebacks' extra "
                '(playwright + keyring) — install with:\n'
                'pip install -e ".[givebacks]"')
        except Exception as e:
            self.failed.emit(str(e))
        else:
            self.finished_ok.emit(saved)


class InputFilesBox(QGroupBox):
    """Shows QuickBooks/Bank/Givebacks file status for the selected month,
    and accepts files dropped anywhere on it (routed by filename)."""

    def __init__(self, parent_window):
        super().__init__('Input Files')
        self.parent_window = parent_window
        self.setAcceptDrops(True)

        self.qb_label = QLabel()
        self.bank_label = QLabel()
        self.gb_label = QLabel()

        qb_btn = QPushButton('Browse…')
        qb_btn.clicked.connect(lambda: self._browse('quickbooks'))
        bank_btn = QPushButton('Browse…')
        bank_btn.clicked.connect(lambda: self._browse('bank'))
        gb_btn = QPushButton('Browse…')
        gb_btn.clicked.connect(lambda: self._browse('givebacks'))
        self.gb_auto_btn = QPushButton('Auto-fetch…')
        self.gb_auto_btn.setToolTip('Download Givebacks CSVs automatically (needs credentials set in Settings)')
        self.gb_auto_btn.clicked.connect(lambda: self.parent_window._auto_fetch_givebacks())

        form = QFormLayout()
        form.addRow('QuickBooks CSV:', self._row(self.qb_label, qb_btn))
        form.addRow('Bank Statement PDF:', self._row(self.bank_label, bank_btn))
        form.addRow('Givebacks CSV(s):', self._row(self.gb_label, gb_btn, self.gb_auto_btn))

        hint = QLabel('You can also drag and drop files here.')
        hint.setStyleSheet('color: gray; font-style: italic;')

        layout = QVBoxLayout()
        layout.addLayout(form)
        layout.addWidget(hint)
        self.setLayout(layout)

    @staticmethod
    def _row(label, *buttons):
        row = QHBoxLayout()
        row.addWidget(label, 1)
        for button in buttons:
            row.addWidget(button)
        container = QWidget()
        container.setLayout(row)
        return container

    def _browse(self, kind):
        filters = {
            'quickbooks': 'CSV files (*.csv)',
            'bank': 'PDF files (*.pdf)',
            'givebacks': 'CSV files (*.csv)',
        }
        path, _ = QFileDialog.getOpenFileName(self, 'Choose a file', '', filters[kind])
        if path:
            self._place_file(Path(path), kind)

    def _place_file(self, src: Path, kind: str):
        month_folder = self.parent_window.current_month_folder()
        dest_folder = month_folder / 'givebacks' if kind == 'givebacks' else month_folder
        dest_folder.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dest_folder / src.name)
        self.parent_window.refresh_file_status()

    @staticmethod
    def _classify(filename: str):
        lower = filename.lower()
        if 'givebacks' in lower:
            return 'givebacks'
        if lower.endswith('.pdf'):
            return 'bank'
        if lower.endswith('.csv'):
            return 'quickbooks'
        return None

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        for url in event.mimeData().urls():
            src = Path(url.toLocalFile())
            kind = self._classify(src.name)
            if kind:
                self._place_file(src, kind)
        event.acceptProposedAction()


class MainWindow(QMainWindow):
    def __init__(self, config: OrgConfig, data_dir: Path):
        super().__init__()
        self.config = config
        self.data_dir = Path(data_dir)
        self.last_output_path = None
        self.last_run_result = None
        self.last_error = None
        self.worker = None

        self.setWindowTitle('PTA Treasurer')
        self._build_menu()

        self.header = QLabel(config.org_name)
        self.header.setStyleSheet('font-size: 18px; font-weight: bold;')

        today = date.today()
        self.month_combo = QComboBox()
        self.month_combo.addItems(MONTH_LABELS)
        self.month_combo.setCurrentIndex(today.month - 1)
        self.month_combo.currentIndexChanged.connect(self.refresh_file_status)

        self.year_spin = QSpinBox()
        self.year_spin.setRange(1990, today.year + 5)
        self.year_spin.setValue(today.year)
        self.year_spin.valueChanged.connect(self.refresh_file_status)

        month_row = QHBoxLayout()
        month_row.addWidget(QLabel('Month:'))
        month_row.addWidget(self.month_combo)
        month_row.addWidget(QLabel('Year:'))
        month_row.addWidget(self.year_spin)
        month_row.addStretch()

        self.files_box = InputFilesBox(self)

        self.generate_btn = QPushButton('Generate Report')
        self.generate_btn.clicked.connect(self._generate_report)

        self.run_all_btn = QPushButton('Run All Months')
        self.run_all_btn.setToolTip('Generates the Treasurer Report for every Month_Year folder found under input/')
        self.run_all_btn.clicked.connect(self._run_all_months)

        self.ledger_fy_combo = QComboBox()
        self._refresh_ledger_fiscal_years()

        self.build_ledger_btn = QPushButton('Build Debits & Credits Ledger')
        self.build_ledger_btn.setToolTip(
            'Rebuilds a whole-fiscal-year check-register-style ledger from every month found under input/')
        self.build_ledger_btn.clicked.connect(self._build_ledger)

        batch_row = QHBoxLayout()
        batch_row.addWidget(self.run_all_btn)
        batch_row.addWidget(QLabel('Fiscal year:'))
        batch_row.addWidget(self.ledger_fy_combo)
        batch_row.addWidget(self.build_ledger_btn, 1)

        batch_box = QGroupBox('Batch Actions (all months)')
        batch_layout = QVBoxLayout()
        batch_layout.addLayout(batch_row)
        batch_box.setLayout(batch_layout)

        self.progress = QProgressBar()
        self.progress.setRange(0, 0)   # indeterminate
        self.progress.hide()

        self.log = QPlainTextEdit(readOnly=True)

        self.open_report_btn = QPushButton('Open Report')
        self.open_report_btn.clicked.connect(self._open_report)
        self.open_report_btn.setEnabled(False)

        self.open_folder_btn = QPushButton('Open Report Folder')
        self.open_folder_btn.clicked.connect(self._open_folder)
        self.open_folder_btn.setEnabled(False)

        result_row = QHBoxLayout()
        result_row.addWidget(self.open_report_btn)
        result_row.addWidget(self.open_folder_btn)
        result_row.addStretch()

        central = QWidget()
        layout = QVBoxLayout()
        layout.addWidget(self.header)
        layout.addLayout(month_row)
        layout.addWidget(self.files_box)
        layout.addWidget(self.generate_btn)
        layout.addWidget(batch_box)
        layout.addWidget(self.progress)
        layout.addWidget(self.log, 1)
        layout.addLayout(result_row)
        central.setLayout(layout)
        self.setCentralWidget(central)

        self.chat_panel = ChatPanel(self.config, self.data_dir)
        self.chat_panel.open_settings_requested.connect(self._open_settings)
        self.chat_dock = QDockWidget('AI Assistant', self)
        self.chat_dock.setWidget(self.chat_panel)
        self.addDockWidget(Qt.RightDockWidgetArea, self.chat_dock)
        self.chat_dock.hide()  # optional feature, off by default
        self._view_menu.addAction(self.chat_dock.toggleViewAction())

        self.refresh_file_status()

    def _build_menu(self):
        file_menu = self.menuBar().addMenu('File')

        settings_action = file_menu.addAction('Settings…')
        settings_action.triggered.connect(self._open_settings)

        file_menu.addSeparator()
        quit_action = file_menu.addAction('Quit')
        quit_action.triggered.connect(self.close)

        self._view_menu = self.menuBar().addMenu('View')

    def current_month_folder(self) -> Path:
        month = MONTH_LABELS[self.month_combo.currentIndex()]
        year = str(self.year_spin.value())
        return self.data_dir / 'input' / f'{month}_{year}'

    def refresh_file_status(self):
        month = MONTH_LABELS[self.month_combo.currentIndex()].lower()
        month_folder = self.current_month_folder()

        if not month_folder.exists():
            self.files_box.qb_label.setText('— not found')
            self.files_box.bank_label.setText('— not found')
            self.files_box.gb_label.setText('— none')
            return

        qb_files = [f for f in month_folder.glob('*.csv')
                    if month in f.name.lower() and 'givebacks' not in f.name.lower()]
        bank_files = [f for f in month_folder.glob('*.pdf') if month in f.name.lower()]
        gb_files = list((month_folder / 'givebacks').glob('*.csv')) \
            if (month_folder / 'givebacks').exists() else []

        self.files_box.qb_label.setText(f'✓ {qb_files[0].name}' if qb_files else '— not found')
        self.files_box.bank_label.setText(f'✓ {bank_files[0].name}' if bank_files else '— not found')
        self.files_box.gb_label.setText(f'✓ {len(gb_files)} file(s)' if gb_files else '— none')

    def _generate_report(self):
        self.generate_btn.setEnabled(False)
        self.progress.show()
        self.log.appendPlainText('Generating report…')

        month = MONTH_LABELS[self.month_combo.currentIndex()]
        year = str(self.year_spin.value())
        self.worker = RunWorker(self.config, self.data_dir, month, year)
        self.worker.finished_ok.connect(self._on_success)
        self.worker.failed.connect(self._on_failure)
        self.worker.start()

    def _on_success(self, result):
        self.progress.hide()
        self.generate_btn.setEnabled(True)
        self.last_output_path = result.output_path
        self.log.appendPlainText(f'Saved -> {result.output_path}')
        for w in result.warnings:
            self.log.appendPlainText(f'WARNING: {w}')
        self.open_report_btn.setEnabled(True)
        self.open_folder_btn.setEnabled(True)

        self.last_run_result = result
        self.last_error = None
        month = MONTH_LABELS[self.month_combo.currentIndex()]
        year = str(self.year_spin.value())
        self.chat_panel.set_context(report=result, month_label=f'{month} {year}')
        self._refresh_ledger_fiscal_years()

    def _on_failure(self, message):
        self.progress.hide()
        self.generate_btn.setEnabled(True)
        self.log.appendPlainText(f'FAILED: {message}')
        QMessageBox.critical(self, 'Report generation failed', message)

        self.last_run_result = None
        self.last_error = message
        self.chat_panel.set_context(error=message)

    def _run_all_months(self):
        input_dir = self.data_dir / 'input'
        month_folders = sorted(f for f in input_dir.iterdir() if f.is_dir()) if input_dir.exists() else []

        self.log.appendPlainText('\nRunning all months…')
        if not month_folders:
            self.log.appendPlainText('No month folders found under input/.')
            return

        for folder in month_folders:
            parts = folder.name.split('_')
            if len(parts) != 2 or not parts[1].isdigit():
                self.log.appendPlainText(f'{folder.name}: skipped (expected "Month_Year" folder name)')
                continue
            month, year = parts
            try:
                result = run_month(self.config, self.data_dir, month, year)
                self.log.appendPlainText(f'{folder.name}: OK -> {result.output_path.name}')
                for w in result.warnings:
                    self.log.appendPlainText(f'  WARNING: {w}')
            except Exception as e:
                self.log.appendPlainText(f'{folder.name}: FAILED — {e}')

        self._refresh_ledger_fiscal_years()

    def _refresh_ledger_fiscal_years(self):
        years = list_available_fiscal_years(self.data_dir)
        self.ledger_fy_combo.clear()
        for y in years:
            self.ledger_fy_combo.addItem(f'{y}-{str(y + 1)[-2:]}', userData=y)

    def _build_ledger(self):
        from pta_treasurer.pipeline import build_debits_credits_ledger

        fiscal_year_start = self.ledger_fy_combo.currentData()

        self.log.appendPlainText('\nBuilding Debits & Credits ledger…')
        try:
            result = build_debits_credits_ledger(self.config, self.data_dir,
                                                  fiscal_year_start=fiscal_year_start)
        except Exception as e:
            self.log.appendPlainText(f'FAILED: {e}')
            return
        self.log.appendPlainText(f'Saved -> {result.output_path}')
        for w in result.warnings:
            self.log.appendPlainText(f'  WARNING: {w}')

    def _auto_fetch_givebacks(self):
        month = MONTH_LABELS[self.month_combo.currentIndex()]
        year = str(self.year_spin.value())
        month_label = f'{month} {year}'
        dest_folder = self.current_month_folder() / 'givebacks'

        self.files_box.gb_auto_btn.setEnabled(False)
        self.log.appendPlainText(f'Fetching Givebacks data for {month_label}…')

        self.gb_worker = GivebacksWorker(month_label, self.config, dest_folder, self.data_dir)
        self.gb_worker.finished_ok.connect(self._on_givebacks_success)
        self.gb_worker.failed.connect(self._on_givebacks_failure)
        self.gb_worker.otp_requested.connect(self._prompt_for_otp)
        self.gb_worker.start()

    def _prompt_for_otp(self):
        code, ok = QInputDialog.getText(
            self, 'One-Time Passcode',
            'Givebacks sent a one-time passcode to your email.\nEnter it here:',
            QLineEdit.Normal, '')
        self.gb_worker.submit_otp(code if ok else '')

    def _on_givebacks_success(self, saved_files):
        self.files_box.gb_auto_btn.setEnabled(True)
        if saved_files:
            self.log.appendPlainText(f'Givebacks: saved {len(saved_files)} file(s).')
        else:
            self.log.appendPlainText('Givebacks: no new payouts found for this month.')
        self.refresh_file_status()

    def _on_givebacks_failure(self, message):
        self.files_box.gb_auto_btn.setEnabled(True)
        self.log.appendPlainText(f'Givebacks fetch FAILED: {message}')
        QMessageBox.critical(self, 'Givebacks auto-fetch failed', message)

    def _open_report(self):
        if self.last_output_path:
            open_in_default_app(self.last_output_path)

    def _open_folder(self):
        if self.last_output_path:
            open_in_default_app(self.last_output_path.parent)

    def _open_settings(self):
        dialog = SettingsDialog(self.config, self.data_dir, parent=self)
        if dialog.exec():
            self.config = load_config(self.data_dir)
            self.header.setText(self.config.org_name)
            self.chat_panel.config = self.config
            self.chat_panel.refresh()
