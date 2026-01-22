macOS helper

- Build the helper:
  - `cd src/Screenshot.Mac/RecorderHelper`
  - `swift build -c release`
- The binary will be at:
  - `src/Screenshot.Mac/RecorderHelper/.build/release/RecorderHelper`
- Pass the helper path into `MacRecordingBackend`.

Notes
- Requires macOS 13+ and Screen Recording permission.
- Uses ScreenCaptureKit system audio (native mode only). Virtual audio device mode is still TODO.
