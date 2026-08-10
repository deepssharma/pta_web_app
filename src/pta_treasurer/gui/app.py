"""
app.py
QApplication entry point. Runs the first-run setup wizard if no data folder
is set yet, then shows the main window.
"""
import sys

from PySide6.QtWidgets import QApplication

from pta_treasurer.config import get_data_dir, load_config
from pta_treasurer.gui.main_window import MainWindow
from pta_treasurer.gui.setup_wizard import SetupWizard


def main():
    app = QApplication(sys.argv)

    data_dir = get_data_dir()
    if data_dir is None:
        wizard = SetupWizard()
        if not wizard.exec() or wizard.data_dir is None:
            return 0
        data_dir = wizard.data_dir

    config = load_config(data_dir)
    window = MainWindow(config, data_dir)
    window.resize(640, 720)
    window.show()

    return app.exec()


if __name__ == '__main__':
    sys.exit(main())
