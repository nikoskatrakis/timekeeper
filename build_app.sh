#!/bin/bash
# Build Timekeeper.app for macOS
# Run this once from the Timekeeper folder: bash build_app.sh

set -e
cd "$(dirname "$0")"

echo "▶ Installing build dependencies..."
pip3 install pyinstaller PySide6 openpyxl pyobjc-framework-Cocoa --quiet

echo "▶ Building icon..."
iconutil -c icns Timekeeper.iconset -o Timekeeper.icns

echo "▶ Building Timekeeper.app..."
pyinstaller --clean --noconfirm timekeeper_app.spec

echo "▶ Copying to /Applications..."
cp -r dist/timekeeper.app /Applications/timekeeper.app

echo ""
echo "✓ Done! timekeeper.app is in /Applications."
echo "  If macOS blocks it: System Settings → Privacy & Security → Open Anyway"
