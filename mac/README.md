mac packaging

- Build mac .app into `mac/bin` and bundle RecorderHelper into the app.
- Run: `mac/build-mac.sh [arm64|x64]`

Notes
- Requires .NET SDK 8+ and Xcode Command Line Tools.
- The app is placed at `mac/bin/Screenshot.app`.
- RecorderHelper is copied into `Contents/Resources/RecorderHelper`.
- Update `BUNDLE_ID` and `VERSION` inside `mac/build-mac.sh` as needed.
