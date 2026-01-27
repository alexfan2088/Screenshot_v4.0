#!/bin/sh
set -e

APP_NAME="Screenshot"
VERSION="4.0.0"
ARCH="${1:-arm64}"
ROOT_DIR="$(cd "$(dirname "$0")/.." && pwd)"
OUT_DIR="$ROOT_DIR/mac/bin"
APP_DIR="$OUT_DIR/$APP_NAME.app"
DMG_DIR="/Users/fwp-mac/Documents/dmg file"
DMG_PATH="/Users/fwp-mac/Documents/可分发dmg/$APP_NAME-$VERSION-$ARCH.dmg"
DMG_RW="/tmp/$APP_NAME-$VERSION-$ARCH.rw.dmg"
DMG_OUT_TMP="/tmp/$APP_NAME-$VERSION-$ARCH.dmg"

if [ ! -d "$APP_DIR" ]; then
  echo "App not found: $APP_DIR" >&2
  exit 1
fi

if [ ! -d "$DMG_DIR" ]; then
  echo "DMG staging folder not found: $DMG_DIR" >&2
  exit 1
fi

rm -f "$DMG_PATH" "$DMG_OUT_TMP" "$DMG_RW"
hdiutil detach "/Volumes/$APP_NAME" -quiet || true
hdiutil detach "/Volumes/$APP_NAME 1" -quiet || true
hdiutil create -size 300m -fs HFS+ -volname "$APP_NAME" -ov "$DMG_RW" >/dev/null
ATTACH_INFO=$(hdiutil attach "$DMG_RW" -readwrite -noverify -owners on)
MOUNT_DIR=$(echo "$ATTACH_INFO" | awk '/Volumes/ {print $3; exit}')
DEVICE=$(echo "$ATTACH_INFO" | awk '/^\/dev\// {print $1; exit}')
if [ -z "$MOUNT_DIR" ]; then
  echo "Failed to mount DMG." >&2
  exit 1
fi

cp -R "$APP_DIR" "$MOUNT_DIR/"
ln -s /Applications "$MOUNT_DIR/Applications"

if [ -f "$DMG_DIR/.background/background.png" ]; then
  mkdir -p "$MOUNT_DIR/.background"
  cp "$DMG_DIR/.background/background.png" "$MOUNT_DIR/.background/background.png"
fi

osascript <<EOF
tell application "Finder"
  tell disk "$APP_NAME"
    open
    delay 1
    set current view of container window to icon view
    set toolbar visible of container window to false
    set statusbar visible of container window to false
    set the bounds of container window to {80, 80, 1060, 720}
    set viewOptions to the icon view options of container window
    set arrangement of viewOptions to not arranged
    set icon size of viewOptions to 180
    set text size of viewOptions to 14
    try
      set background picture of viewOptions to file ".background:background.png"
    end try
    set position of item "$APP_NAME.app" to {490, 510}
    set position of item "Applications" to {490, 210}
    delay 1
    close
    delay 1
  end tell
end tell
EOF

sync
hdiutil detach "${DEVICE:-$MOUNT_DIR}" -quiet || true
sync
sleep 1
hdiutil convert "$DMG_RW" -format UDZO -o "$DMG_OUT_TMP" >/dev/null
rm -f "$DMG_RW"
mv "$DMG_OUT_TMP" "$DMG_PATH"

echo "Done: $DMG_PATH"
