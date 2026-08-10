#!/usr/bin/env bash
# Deep-signs, notarizes, and staples an already-built dist/PTA Treasurer.app.
#
# Prerequisites (one-time, see packaging/README.md "Signed & Notarized
# Release Build" -- these require a paid Apple Developer account and
# can't be done by an automated script):
#   - A "Developer ID Application" certificate installed in your login
#     keychain (Keychain Access -> Certificate Assistant -> Request a
#     Certificate from a Certificate Authority, then upload/download via
#     developer.apple.com).
#   - `xcrun notarytool store-credentials "pta-treasurer-notarize" \
#       --apple-id "<your Apple ID email>" --team-id "<your Team ID>" \
#       --password "<an app-specific password from appleid.apple.com>"`
#     run once, storing credentials in the macOS Keychain under that
#     profile name -- never in a file this repo controls.
#
# Usage:
#   CODESIGN_IDENTITY="Developer ID Application: Your Name (TEAMID)" \
#     packaging/notarize_mac.sh
#
# Normally invoked via packaging/build_mac_release.sh, not run directly.

set -euo pipefail
cd "$(dirname "$0")/.."

APP_PATH="dist/PTA Treasurer.app"
ZIP_PATH="dist/PTA Treasurer.zip"
KEYCHAIN_PROFILE="pta-treasurer-notarize"

if [ -z "${CODESIGN_IDENTITY:-}" ]; then
    echo "ERROR: CODESIGN_IDENTITY is not set. Example:" >&2
    echo '  CODESIGN_IDENTITY="Developer ID Application: Your Name (TEAMID)" packaging/notarize_mac.sh' >&2
    exit 1
fi

if [ ! -d "$APP_PATH" ]; then
    echo "ERROR: $APP_PATH not found -- run packaging/build_mac.sh first." >&2
    exit 1
fi

echo "== Deep-signing with hardened runtime =="
# PyInstaller's own codesign_identity (wired in pyinstaller_mac.spec) signs
# the top-level bundle, but nested frameworks/dylibs (Python.framework,
# pypdfium2_raw's libpdfium.dylib, Qt frameworks) need an explicit
# deep-sign pass with hardened runtime -- notarization checks every nested
# code object, not just the outer bundle.
codesign --deep --force --options runtime \
    --entitlements packaging/entitlements.plist \
    --sign "$CODESIGN_IDENTITY" \
    "$APP_PATH"

echo "== Verifying signature =="
codesign --verify --deep --strict --verbose=2 "$APP_PATH"

echo "== Zipping for submission =="
rm -f "$ZIP_PATH"
ditto -c -k --keepParent "$APP_PATH" "$ZIP_PATH"

echo "== Submitting for notarization (this can take several minutes) =="
xcrun notarytool submit "$ZIP_PATH" \
    --keychain-profile "$KEYCHAIN_PROFILE" \
    --wait

echo "== Stapling notarization ticket =="
xcrun stapler staple "$APP_PATH"

echo "== Gatekeeper assessment (should say 'accepted' / 'Notarized Developer ID') =="
spctl -a -vvv "$APP_PATH"

rm -f "$ZIP_PATH"
echo ""
echo "Signed, notarized, and stapled: $APP_PATH"
