#!/bin/sh
set -e

APP_NAME="Screenshot"
VERSION="4.0.0"
ARCH="${1:-arm64}"
ROOT_DIR="$(cd "$(dirname "$0")/.." && pwd)"
OUT_DIR="$ROOT_DIR/mac/bin"
APP_DIR="$OUT_DIR/$APP_NAME.app"
DMG_DIR="/Users/fwp-mac/Documents/dmg file"
DMG_PATH="$DMG_DIR/$APP_NAME-$VERSION-$ARCH.dmg"
DMG_OUT_TMP="/tmp/$APP_NAME-$VERSION-$ARCH.dmg"

if [ ! -d "$APP_DIR" ]; then
  echo "App not found: $APP_DIR" >&2
  exit 1
fi

if [ ! -d "$DMG_DIR" ]; then
  echo "DMG staging folder not found: $DMG_DIR" >&2
  exit 1
fi

rm -rf "$DMG_DIR/$APP_NAME.app"
cp -R "$APP_DIR" "$DMG_DIR/"
if [ ! -e "$DMG_DIR/Applications" ]; then
  ln -s /Applications "$DMG_DIR/Applications"
fi

rm -f "$DMG_PATH" "$DMG_OUT_TMP"
hdiutil create -volname "$APP_NAME" -srcfolder "$DMG_DIR" -ov -format UDZO "$DMG_OUT_TMP" >/dev/null
mv "$DMG_OUT_TMP" "$DMG_PATH"

echo "Done: $DMG_PATH"
