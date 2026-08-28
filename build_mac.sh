#!/bin/bash
# Build, sign, and package the macOS app. Run from the project root:
#   ./build_mac.sh
set -euo pipefail

IDENTITY="Think Academy Automation"
APP="dist/Teacher Feedback Desk.app"

./.venv/bin/python -m PyInstaller TeacherFeedbackDesk.spec --noconfirm

# PyInstaller signs the launcher but leaves python.org's own signature on the
# bundled Python.framework. A cert-signed (non-ad-hoc) launcher may not load a
# library carrying a different Team ID, so the app dies instantly at startup
# ("different Team IDs" in dyld's error). Deep re-signing puts one identity on
# every nested Mach-O.
codesign --force --deep --sign "$IDENTITY" "$APP"
codesign --verify --deep --strict "$APP"

STAGE=$(mktemp -d)
cp -R "$APP" "$STAGE/"
ln -s /Applications "$STAGE/Applications"
rm -f "dist/Teacher Feedback Desk.dmg"
hdiutil create -volname "Teacher Feedback Desk" -srcfolder "$STAGE" -ov -format UDZO "dist/Teacher Feedback Desk.dmg"
rm -rf "$STAGE"

echo ""
echo "Done: $APP"
echo "      dist/Teacher Feedback Desk.dmg"
echo "Install locally with:  rm -rf '/Applications/Teacher Feedback Desk.app' && ditto \"$APP\" '/Applications/Teacher Feedback Desk.app'"
