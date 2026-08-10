"""
utils.py
Small cross-platform helpers shared by the GUI windows/dialogs.
"""
import subprocess
import sys
from pathlib import Path


def open_in_default_app(path: Path) -> None:
    """Opens a file or folder in the OS's default application/file browser."""
    path = str(path)
    if sys.platform == 'darwin':
        subprocess.run(['open', path])
    elif sys.platform == 'win32':
        import os
        os.startfile(path)
    else:
        subprocess.run(['xdg-open', path])
