---
'@pptx-studio/cli': minor
---

`render` on macOS lays text out on the box CoreText reads: `hhea` whatever `fsSelection` bit 7 says, rounded half up from the em ratio held in single precision, measured 247/247 on `macos-latest`. `FontBackend` gains `coretext` and `FaceMetrics` says which backend it was read for.
