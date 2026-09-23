#!/bin/bash
# Build, sign, and package the macOS app. Run from the project root:
#   ./build_mac.sh
set -euo pipefail

IDENTITY="Think Academy Automation"
APP="dist/Teacher Feedback Desk.app"
VOLNAME="Teacher Feedback Desk"
DMG="dist/Teacher Feedback Desk.dmg"

./.venv/bin/python -m PyInstaller TeacherFeedbackDesk.spec --noconfirm

# PyInstaller signs the launcher but leaves python.org's own signature on the
# bundled Python.framework. A cert-signed (non-ad-hoc) launcher may not load a
# library carrying a different Team ID, so the app dies instantly at startup
# ("different Team IDs" in dyld's error). Deep re-signing puts one identity on
# every nested Mach-O.
codesign --force --deep --sign "$IDENTITY" "$APP"
codesign --verify --deep --strict "$APP"

# Styled DMG: an icon-view window whose background literally tells the teacher
# to drag the app onto Applications. Without it, people double-click the app
# inside the mounted DMG and their permission grants never line up. create-dmg
# handles the Finder layout; plain AppleScript view-option setting broke on
# recent macOS (-10006).
STAGE=$(mktemp -d)
cp -R "$APP" "$STAGE/"
rm -f "$DMG"
create-dmg \
  --volname "$VOLNAME" \
  --background dmg_background.png \
  --window-size 600 440 \
  --icon-size 96 \
  --icon "Teacher Feedback Desk.app" 150 195 \
  --app-drop-link 450 195 \
  --hide-extension "Teacher Feedback Desk.app" \
  --no-internet-enable \
  "$DMG" "$STAGE" >/dev/null
rm -rf "$STAGE"

echo ""
echo "Done: $APP"
echo "      $DMG"
echo "Install locally with:  rm -rf '/Applications/Teacher Feedback Desk.app' && ditto \"$APP\" '/Applications/Teacher Feedback Desk.app'"
