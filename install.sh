#!/bin/bash
# Installs MarkerTracker.app without the Gatekeeper quarantine flag.
# Usage: curl -fsSL https://raw.githubusercontent.com/Tboy662/FSV-Marker/main/install.sh | bash
set -euo pipefail

APP_NAME="MarkerTracker.app"
ZIP_URL="https://github.com/Tboy662/FSV-Marker/raw/main/MarkerTracker.zip"

DEST="$HOME/Applications"
mkdir -p "$DEST"

TMP="$(mktemp -d)"
trap 'rm -rf "$TMP"' EXIT

echo "Downloading $APP_NAME..."
curl -fL --progress-bar -o "$TMP/app.zip" "$ZIP_URL"

echo "Installing to $DEST..."
ditto -x -k "$TMP/app.zip" "$TMP/unzipped"
rm -rf "$DEST/$APP_NAME"
mv "$TMP/unzipped/$APP_NAME" "$DEST/"

xattr -dr com.apple.quarantine "$DEST/$APP_NAME" 2>/dev/null || true

# Ad-hoc re-sign so macOS doesn't reject a missing or broken signature
codesign --force --deep --sign - "$DEST/$APP_NAME" 2>/dev/null || true

echo "Installed: $DEST/$APP_NAME"
open "$DEST/$APP_NAME"
