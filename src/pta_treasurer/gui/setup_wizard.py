"""
setup_wizard.py
First-run wizard: choose a data folder, enter org info, and create a blank
budget template ready for the treasurer to fill in.
"""
from pathlib import Path

from PySide6.QtWidgets import (
    QDoubleSpinBox, QFileDialog, QHBoxLayout, QLabel, QLineEdit,
    QMessageBox, QPushButton, QVBoxLayout, QWizard, QWizardPage,
)

from pta_treasurer.budget_io import generate_template
from pta_treasurer.config import OrgConfig, save_config, set_data_dir

DEFAULT_DATA_DIR = Path.home() / 'Documents' / 'PTA Treasurer'


class DataFolderPage(QWizardPage):
    def __init__(self):
        super().__init__()
        self.setTitle('Choose a data folder')
        self.setSubTitle(
            'This is where your monthly reports, budget, and input files will '
            'live. It can be backed up or moved independently of the app.'
        )
        self.path_edit = QLineEdit(str(DEFAULT_DATA_DIR))
        browse_btn = QPushButton('Browse…')
        browse_btn.clicked.connect(self._browse)

        row = QHBoxLayout()
        row.addWidget(self.path_edit)
        row.addWidget(browse_btn)
        layout = QVBoxLayout()
        layout.addLayout(row)
        self.setLayout(layout)

        self.registerField('data_dir*', self.path_edit)

    def _browse(self):
        chosen = QFileDialog.getExistingDirectory(
            self, 'Choose a data folder', self.path_edit.text())
        if chosen:
            self.path_edit.setText(chosen)


class OrgInfoPage(QWizardPage):
    def __init__(self):
        super().__init__()
        self.setTitle('Organization info')
        self.org_name_edit = QLineEdit()
        self.org_name_edit.setPlaceholderText('e.g. Example Elementary PTA')

        self.balance_spin = QDoubleSpinBox()
        self.balance_spin.setRange(-1_000_000, 1_000_000)
        self.balance_spin.setDecimals(2)
        self.balance_spin.setPrefix('$')

        layout = QVBoxLayout()
        layout.addWidget(QLabel('Organization name:'))
        layout.addWidget(self.org_name_edit)
        layout.addWidget(QLabel(
            "Balance forward (checking account balance at the start of "
            "this fiscal year, July 1):"
        ))
        layout.addWidget(self.balance_spin)
        self.setLayout(layout)

        self.registerField('org_name*', self.org_name_edit)


class SetupWizard(QWizard):
    """On completion, self.data_dir is set to the chosen/created data folder."""

    def __init__(self):
        super().__init__()
        self.setWindowTitle('PTA Treasurer — First-Run Setup')
        self.data_folder_page = DataFolderPage()
        self.org_info_page = OrgInfoPage()
        self.addPage(self.data_folder_page)
        self.addPage(self.org_info_page)
        self.data_dir = None

    def accept(self):
        data_dir = Path(self.field('data_dir')).expanduser()
        org_name = self.field('org_name')
        balance_forward = self.org_info_page.balance_spin.value()

        try:
            for sub in ('input', 'output', 'data/history'):
                (data_dir / sub).mkdir(parents=True, exist_ok=True)

            set_data_dir(data_dir)
            save_config(OrgConfig(org_name=org_name, balance_forward=balance_forward), data_dir)

            budget_path = data_dir / 'budget.xlsx'
            if not budget_path.exists():
                generate_template(budget_path)
        except OSError as e:
            QMessageBox.critical(self, 'Setup failed', str(e))
            return

        self.data_dir = data_dir
        super().accept()
