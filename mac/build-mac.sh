#!/bin/sh
set -e

APP_NAME="Screenshot"
BUNDLE_ID="com.screenshot.app"
VERSION="4.0.0"
ARCH="${1:-arm64}"
ROOT_DIR="$(cd "$(dirname "$0")/.." && pwd)"
APP_PROJ="$ROOT_DIR/src/Screenshot.App/Screenshot.App.csproj"
HELPER_DIR="$ROOT_DIR/src/Screenshot.Mac/RecorderHelper"
HELPER_BIN="$HELPER_DIR/.build/release/RecorderHelper"
OUT_DIR="$ROOT_DIR/mac/bin"
PUBLISH_DIR="$ROOT_DIR/mac/build/publish-$ARCH"
APP_DIR="$OUT_DIR/$APP_NAME.app"
MACOS_DIR="$APP_DIR/Contents/MacOS"
RES_DIR="$APP_DIR/Contents/Resources"
PLIST="$APP_DIR/Contents/Info.plist"

mkdir -p "$OUT_DIR"
mkdir -p "$ROOT_DIR/mac/build"

if ! command -v dotnet >/dev/null 2>&1; then
  echo "dotnet not found. Install .NET SDK 8+ first." >&2
  exit 1
fi

if ! command -v swift >/dev/null 2>&1; then
  echo "swift not found. Install Xcode Command Line Tools first." >&2
  exit 1
fi

echo "[1/4] Build RecorderHelper (Swift)"
cd "$HELPER_DIR"
swift build -c release

if [ ! -f "$HELPER_BIN" ]; then
  echo "RecorderHelper build failed: $HELPER_BIN not found" >&2
  exit 1
fi

cd "$ROOT_DIR"

echo "[2/4] Publish Avalonia app ($ARCH)"
DOTNET_RUNTIME="osx-$ARCH"
dotnet publish "$APP_PROJ" -c Release -r "$DOTNET_RUNTIME" --self-contained true -o "$PUBLISH_DIR"

if [ ! -d "$PUBLISH_DIR" ]; then
  echo "Publish output not found: $PUBLISH_DIR" >&2
  exit 1
fi

APP_HOST=$(ls "$PUBLISH_DIR" | grep -E '^Screenshot\.App(\.dll)?$' | head -n 1)
if [ -z "$APP_HOST" ]; then
  APP_HOST=$(ls "$PUBLISH_DIR" | grep -E '^Screenshot\.App$' | head -n 1)
fi

if [ -z "$APP_HOST" ]; then
  echo "App host not found in publish output." >&2
  exit 1
fi

echo "[3/4] Assemble .app bundle"
rm -rf "$APP_DIR"
mkdir -p "$MACOS_DIR" "$RES_DIR"

cp "$PUBLISH_DIR/$APP_HOST" "$MACOS_DIR/$APP_NAME"
chmod +x "$MACOS_DIR/$APP_NAME"

cp -R "$PUBLISH_DIR"/* "$RES_DIR/" >/dev/null 2>&1 || true

cp "$HELPER_BIN" "$RES_DIR/RecorderHelper"
chmod +x "$RES_DIR/RecorderHelper"

echo "[4/4] Write Info.plist"
cat > "$PLIST" <<PLIST
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>CFBundleName</key>
  <string>$APP_NAME</string>
  <key>CFBundleDisplayName</key>
  <string>$APP_NAME</string>
  <key>CFBundleIdentifier</key>
  <string>$BUNDLE_ID</string>
  <key>CFBundleVersion</key>
  <string>$VERSION</string>
  <key>CFBundleShortVersionString</key>
  <string>$VERSION</string>
  <key>CFBundleExecutable</key>
  <string>$APP_NAME</string>
  <key>LSMinimumSystemVersion</key>
  <string>13.0</string>
  <key>NSHighResolutionCapable</key>
  <true/>
  <key>LSEnvironment</key>
  <dict>
    <key>SCREENSHOT_MAC_HELPER</key>
    <string>@executable_path/../Resources/RecorderHelper</string>
  </dict>
</dict>
</plist>
PLIST

echo "Done: $APP_DIR"
