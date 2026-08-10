# Builds the Windows distributable for PTA Treasurer.
#
# STATUS: UNTESTED. This script has not been run on a Windows machine --
# it was written on macOS alongside pyinstaller_win.spec, mirroring the
# already-verified packaging/build_mac.sh. Treat it as a documented
# starting point, not a proven build path, until someone actually runs it
# on Windows and works through the hand-test checklist in
# packaging/README.md.
#
# Usage (from a PowerShell prompt, repo root):
#   .\packaging\build_win.ps1
# Output: dist\PTA Treasurer\PTA Treasurer.exe

$ErrorActionPreference = "Stop"
Set-Location (Join-Path $PSScriptRoot "..")

$BuildVenv = ".venv-build-win"

if (Test-Path $BuildVenv) { Remove-Item -Recurse -Force $BuildVenv }
if (Test-Path "build") { Remove-Item -Recurse -Force "build" }
if (Test-Path "dist") { Remove-Item -Recurse -Force "dist" }

python -m venv $BuildVenv
& "$BuildVenv\Scripts\Activate.ps1"

pip install --upgrade pip
pip install .
pip install "pyinstaller==6.21.0" "pyinstaller-hooks-contrib"

pyinstaller packaging\pyinstaller_win.spec --noconfirm --clean

deactivate
Write-Host ""
Write-Host "Built: dist\PTA Treasurer\PTA Treasurer.exe"
