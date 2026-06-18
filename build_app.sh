#!/bin/bash
# Build Timekeeper.app for macOS
# Run this once from the Timekeeper folder: bash build_app.sh

set -e
cd "$(dirname "$0")"

echo "▶ Incrementing version..."
VNUM=$(grep 'APP_VERSION' timekeeper_app.py | grep -o '[0-9]*"' | tr -d '"')
NEWNUM=$((VNUM + 1))
NEWVER=$(printf '%05d' $NEWNUM)
sed -i '' "s/APP_VERSION        = \"v0\.[0-9]*/APP_VERSION        = \"v0.$NEWVER/" timekeeper_app.py
echo "  Version: v0.$NEWVER"

echo "▶ Installing build dependencies..."
pip3 install pyinstaller PySide6 openpyxl pyobjc-framework-Cocoa --quiet

echo "▶ Building icon..."
iconutil -c icns Timekeeper.iconset -o Timekeeper.icns

echo "▶ Clearing build cache..."
rm -rf build/ dist/

echo "▶ Building Timekeeper.app..."
pyinstaller --noconfirm timekeeper_app.spec

echo "▶ Removing old app from /Applications..."
sudo rm -rf /Applications/timekeeper.app || true

echo "▶ Copying to /Applications..."
cp -r dist/timekeeper.app /Applications/timekeeper.app

echo ""
echo "✓ Done! timekeeper.app is in /Applications."
echo "  If macOS blocks it: System Settings → Privacy & Security → Open Anyway"
