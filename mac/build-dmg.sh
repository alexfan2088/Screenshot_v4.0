#!/bin/sh
set -e

APP_NAME="Screenshot"
VERSION="4.0.0"
ARCH="${1:-arm64}"
ROOT_DIR="$(cd "$(dirname "$0")/.." && pwd)"
OUT_DIR="$ROOT_DIR/mac/bin"
APP_DIR="$OUT_DIR/$APP_NAME.app"
DMG_TMP="$ROOT_DIR/mac/build/${APP_NAME}-${ARCH}.tmp.dmg"
DMG_PATH="$OUT_DIR/$APP_NAME-$VERSION-$ARCH.dmg"
DMG_DIR="$ROOT_DIR/mac/build/dmg-$ARCH"
BG_SRC="$ROOT_DIR/mac/dmg-background.png"

if [ ! -d "$APP_DIR" ]; then
  echo "App not found: $APP_DIR" >&2
  exit 1
fi

rm -rf "$DMG_DIR"
mkdir -p "$DMG_DIR"
cp -R "$APP_DIR" "$DMG_DIR/"
ln -s /Applications "$DMG_DIR/Applications"

rm -f "$DMG_TMP" "$DMG_PATH"
hdiutil create -volname "$APP_NAME" -srcfolder "$DMG_DIR" -ov -format UDRW "$DMG_TMP" >/dev/null

ATTACH_INFO=$(hdiutil attach "$DMG_TMP" -readwrite -nobrowse -noverify -owners on)
MOUNT_DIR=$(echo "$ATTACH_INFO" | awk '/Volumes/ {print $3; exit}')
DEVICE=$(echo "$ATTACH_INFO" | awk '/^\/dev\// {print $1; exit}')
if [ -z "$MOUNT_DIR" ]; then
  echo "Failed to mount DMG." >&2
  exit 1
fi

WRITABLE=1
touch "$MOUNT_DIR/.writable_test" 2>/dev/null || WRITABLE=0
if [ "$WRITABLE" -eq 1 ]; then
  rm -f "$MOUNT_DIR/.writable_test"
  mkdir -p "$MOUNT_DIR/.background"
  if [ -f "$BG_SRC" ]; then
    cp "$BG_SRC" "$MOUNT_DIR/.background/background.png" || true
  fi

  osascript <<EOF
tell application "Finder"
  tell disk "$APP_NAME"
    open
    set current view of container window to icon view
    set toolbar visible of container window to false
    set statusbar visible of container window to false
    set the bounds of container window to {100, 100, 600, 420}
    set viewOptions to the icon view options of container window
    set arrangement of viewOptions to not arranged
    set icon size of viewOptions to 96
    set background picture of viewOptions to file ".background:background.png"
    set position of item "$APP_NAME.app" to {160, 210}
    set position of item "Applications" to {440, 210}
    close
    open
    update without registering applications
    delay 1
  end tell
end tell
EOF
else
  echo "Warning: DMG mounted read-only, skipping background/layout." >&2
fi

hdiutil detach "${DEVICE:-$MOUNT_DIR}" -quiet -force || true
sleep 1
hdiutil convert "$DMG_TMP" -format UDZO -o "$DMG_PATH" >/dev/null
rm -f "$DMG_TMP"

echo "Done: $DMG_PATH"
