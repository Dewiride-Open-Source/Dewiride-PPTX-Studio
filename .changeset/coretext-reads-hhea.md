---
'@pptx-studio/cli': minor
---

`render` on macOS lays text out on the box CoreText reads: `hhea` whatever `fsSelection` bit 7 says, rounded half up from the em ratio held as 16.16 fixed point, measured 248/248 on `macos-latest`. `FontBackend` gains `coretext` and `FaceMetrics` says which backend it was read for.
