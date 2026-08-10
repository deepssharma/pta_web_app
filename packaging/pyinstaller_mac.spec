# PyInstaller spec for the macOS build of PTA Treasurer.
#
# Build with: packaging/build_mac.sh (see packaging/README.md for details).
# Produces: dist/PTA Treasurer.app
#
# Notes:
# - onedir (not onefile): launches instantly, and the unpacked folder is
#   easy to inspect when debugging what got bundled.
# - `hiddenimports` starts empty. PyInstaller + pyinstaller-hooks-contrib
#   already ship hooks for pypdfium2/pypdfium2_raw (pdfplumber's PDF
#   backend, including its native libpdfium binary), pdfminer, openpyxl,
#   platformdirs, and every PySide6 submodule. Only add entries here if a
#   real build proves something is missing (a ModuleNotFoundError at
#   runtime) -- don't pre-guess.
# - `excludes` drops PySide6 submodules the app never imports (GUI code
#   only touches QtCore/QtWidgets/QtGui/QtPrintSupport). PySide6 alone is
#   ~1.1GB installed and is the dominant bundle-size lever, not pdfplumber.
# - `playwright` and `httpx` are intentionally NOT bundled. Neither has an
#   import site anywhere in src/ yet (Phase 4, unbuilt) -- PyInstaller only
#   bundles what it can trace from the entry script's import graph, so
#   there is nothing to force in. Add playwright to hiddenimports and
#   re-verify the build when Phase 4 lands givebacks_download.py. This
#   doesn't contradict PLAN.md's "ships the playwright package, not the
#   Chromium binary" decision -- that's about what the finished v1 product
#   bundles, not when to wire it into this spec.
# - No `datas`: budget_io.generate_template() builds the workbook
#   programmatically, it doesn't read a template file off disk. No icon:
#   no .icns exists yet in this repo (nice-to-have, not a blocker).
# - Codesigning: off by default (CODESIGN_IDENTITY unset), same as the
#   original unsigned dev build -- unsigned builds show macOS Gatekeeper's
#   "unidentified developer" warning on first launch, right-click the
#   .app -> Open to bypass. Set CODESIGN_IDENTITY (e.g. "Developer ID
#   Application: Your Name (TEAMID)") to sign the top-level bundle at
#   build time; packaging/notarize_mac.sh does the deep-sign +
#   notarization pass PyInstaller's own signing here doesn't cover (nested
#   frameworks/dylibs need hardened-runtime entitlements applied
#   separately). See packaging/README.md's "Signed & Notarized Release
#   Build" section.

import os
import sys
from pathlib import Path

block_cipher = None

repo_root = Path(SPECPATH).parent

codesign_identity = os.environ.get('CODESIGN_IDENTITY') or None
entitlements_file = (str(repo_root / 'packaging' / 'entitlements.plist')
                      if codesign_identity else None)

unused_qt_excludes = [
    'PySide6.QtWebEngineCore', 'PySide6.QtWebEngineQuick', 'PySide6.QtWebEngineWidgets',
    'PySide6.QtQml', 'PySide6.QtQuick', 'PySide6.QtQuick3D', 'PySide6.QtQuickWidgets',
    'PySide6.QtQuickControls2',
    'PySide6.Qt3DAnimation', 'PySide6.Qt3DCore', 'PySide6.Qt3DExtras',
    'PySide6.Qt3DInput', 'PySide6.Qt3DLogic', 'PySide6.Qt3DRender',
    'PySide6.QtMultimedia', 'PySide6.QtMultimediaWidgets', 'PySide6.QtSpatialAudio',
    'PySide6.QtBluetooth', 'PySide6.QtNfc', 'PySide6.QtSerialPort', 'PySide6.QtSerialBus',
    'PySide6.QtSensors', 'PySide6.QtPositioning', 'PySide6.QtLocation',
    'PySide6.QtCharts', 'PySide6.QtDataVisualization', 'PySide6.QtGraphs',
    'PySide6.QtGraphsWidgets',
    'PySide6.QtTextToSpeech', 'PySide6.QtRemoteObjects', 'PySide6.QtScxml',
    'PySide6.QtStateMachine', 'PySide6.QtNetworkAuth', 'PySide6.QtHttpServer',
    'PySide6.QtSql',
]

a = Analysis(
    [str(repo_root / 'packaging' / 'launcher.py')],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=unused_qt_excludes,
    noarchive=False,
    cipher=block_cipher,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='PTA Treasurer',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=codesign_identity,
    entitlements_file=entitlements_file,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='PTA Treasurer',
)

app = BUNDLE(
    coll,
    name='PTA Treasurer.app',
    icon=None,
    bundle_identifier='com.ptatreasurer.app',
    info_plist={
        'CFBundleName': 'PTA Treasurer',
        'CFBundleDisplayName': 'PTA Treasurer',
        'CFBundleShortVersionString': '0.1.0',
        'NSHighResolutionCapable': True,
    },
)
