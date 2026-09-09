---
'@pptx-studio/cli': minor
---

`renderDeck` takes only what it draws.

Its options were the `render` verb's own, so a caller that wanted SVG strings passed
`out`, `json` and `quiet` — three fields the function ignores — and five more it had no
opinion about. `RenderDeckOptions` is the shape now: `slide`, `width`, `fontDirs`,
`systemFonts` and `text`, all optional. `renderDeck(bytes)` is a whole call, and
`renderDeck(bytes, { slide: 1, width: 640 })` is a thumbnail.

`RenderOptions` is the verb's: it extends `RenderDeckOptions` with `out`, `json` and
`quiet`, and `runRender` is what takes it. `RENDER_DEFAULTS` is gone, because the
defaults are in `renderDeck` and `DEFAULT_WIDTH` is still exported. A width that is not
a positive whole number now throws `CLI_WIDTH` instead of drawing an empty picture.

Breaking: a call written against 0.1.0 passes three properties this shape does not have.
Delete them, and delete any field you were only passing to satisfy the type.
