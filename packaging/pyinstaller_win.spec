# PyInstaller spec for the Windows build of PTA Treasurer.
#
# STATUS: written and structurally mirrors pyinstaller_mac.spec (which has
# been built and hand-tested), but this spec itself has NOT been run on
# Windows -- PyInstaller does not cross-compile, so a .exe can only be
# produced by running PyInstaller on an actual Windows machine. There is
# no CI in this repo yet to do this automatically. Treat this file as
# "ready to try", not "verified", until someone runs
# packaging/build_win.ps1 on Windows and works through the hand-test
# checklist in packaging/README.md.
#
# Build with: packaging/build_win.ps1 (on Windows; see packaging/README.md)
# Output: dist/PTA Treasurer/PTA Treasurer.exe
#
# Notes (see pyinstaller_mac.spec for the full rationale on each point --
# repeated briefly here since the two specs are meant to be read
# independently):
# - onedir, not onefile.
# - `hiddenimports` starts empty, relying on PyInstaller +
#   pyinstaller-hooks-contrib's existing hooks for pypdfium2/pdfminer/
#   openpyxl/platformdirs/PySide6. Add entries here only if a real
#   Windows build proves something is missing.
# - `excludes` drops the same unused PySide6 submodules as the Mac spec --
#   this app only uses QtCore/QtWidgets/QtGui/QtPrintSupport.
# - `playwright`/`httpx` intentionally NOT bundled -- no import site in
#   src/ yet (Phase 4, unbuilt). Add to hiddenimports when
#   givebacks_download.py lands.
# - No `datas` (budget template is generated programmatically). No icon
#   (no .ico exists yet in this repo).
# - Windows-specific unknowns flagged in packaging/README.md: whether
#   PySide6's plugin-path handling needs extra spec entries on Windows,
#   and whether pypdfium2_raw's Windows DLL gets picked up by the same
#   contrib hook that worked for the macOS .dylib -- both plausible but
#   unverified until a real Windows build is attempted.

from pathlib import Path

block_cipher = None

repo_root = Path(SPECPATH).parent

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
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
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
