#!/usr/bin/env bash
# Builds a release macOS .app -- runs the same steps as build_mac.sh
# (unsigned dev build), then signs + notarizes it if CODESIGN_IDENTITY is
# set. build_mac.sh itself is untouched by this -- the fast, no-account-
# needed dev loop stays exactly as it was.
#
# Unsigned (same as build_mac.sh):
#   packaging/build_mac_release.sh
#
# Signed + notarized (needs the one-time Apple Developer setup in
# packaging/README.md's "Signed & Notarized Release Build" section):
#   CODESIGN_IDENTITY="Developer ID Application: Your Name (TEAMID)" \
#     packaging/build_mac_release.sh

set -euo pipefail
cd "$(dirname "$0")/.."

packaging/build_mac.sh

if [ -n "${CODESIGN_IDENTITY:-}" ]; then
    echo ""
    packaging/notarize_mac.sh
else
    echo ""
    echo "CODESIGN_IDENTITY not set -- built unsigned (same as build_mac.sh)."
    echo "For a signed, notarized release build, see packaging/README.md."
fi
