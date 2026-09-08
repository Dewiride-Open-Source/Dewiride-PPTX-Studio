# 0039 — Upright glyphs in vertical text, and the insets that never turned

- **Status** accepted
- **Sub-phase** 3.6, reopened
- **Experiment** T10 — 87 probes across 14 slides, PowerPoint driven over COM,
  read back through an EMF drawing stream and a bitmap at four pixels to the point
- **Fixture** [`corpus/ground-truth/vertical-glyphs.json`](../../../corpus/ground-truth/vertical-glyphs.json)
- **Code** `packages/text/src/frames/frame.ts`, `packages/text/src/lines/baseline.ts`,
  `packages/render-svg/src/text/{layout,emit}.ts`, `packages/render-dom/src/text/layer.ts`

## Context

3.6 was recorded as accepted with one shortfall: `eaVert` was laid out and then
refused rather than drawn, so three corpus slides scored nothing. ADR 0032 has
its axes — it stacks leftward from the right edge and aligns from the top,
exactly as `vert` does — and 12 of its 46 probes are numerically identical to
their `vert` twins. What nobody had measured is the only thing that differs:
what happens to the **glyphs**.

Going to look cost one afternoon and turned up two things, one of them a bug in
code that has been scoring 152 slides since 3.9.

## `eaVert` is `vert` with the East Asian glyphs stood back up

`一` is one horizontal stroke, so its ink answers the whole orientation question
without a model: upright it is 81.75 by 5.25 points at 96pt, turned it is 5.25 by
81.75, and there is nothing in between. Scored against the transpose rather than
against a tolerance, **45 of 45**:

| `@vert`         | East Asian glyph | Latin glyph |
| --------------- | ---------------- | ----------- |
| `horz`          | upright          | upright     |
| `vert`          | turned           | turned      |
| `vert270`       | turned           | turned      |
| `eaVert`        | **upright**      | turned      |
| `mongolianVert` | **upright**      | turned      |

The drawing stream says how: PowerPoint asks GDI for the **`@`-prefixed vertical
variant** of the face for that stretch and for the plain face beside it. `一一`
comes back as `@Yu Gothic` and the `L` next to it as `Yu Gothic`, on the same
slide, under the same world transform. **The choice is per script run, not per
frame** — 32 of 32 over eight mixed strings, against the rival that gives one
frame one face. It is the `a:ea` slot 3.5 already computes, which is why
`splitScriptRuns` needed no new table.

Everything else is untouched. The line pitch is **1.2 of the em on all five
faces**, read from two pens rather than from a bounding box, so it is exact
rather than within the five points `BoundWidth` overstates by. The cell pitch is
the face's own advance, the same in every direction. So the block, the wrap, the
anchor and the advance are `vert`'s, and only the glyph turns back.

### What a browser has to do instead

There is no `@` face in a browser, so the renderer places each East Asian glyph
itself and turns it a quarter back out of the line. That costs the shaping across
that one boundary and buys the only rendering of `eaVert` that is not simply
`vert` wearing its name.

Two numbers place the glyph. Along the run, its alphabetic baseline sits
`1 - ideographicBaseline` down its cell. Across the line, its advance box is
centred. Measured against PowerPoint, as fractions of the em:

| face            | cell   | baseline along | error      | origin across | error      |
| --------------- | ------ | -------------- | ---------- | ------------- | ---------- |
| Yu Gothic       | 1.0002 | 0.8801         | **0.0002** | 1.0609        | −0.0391    |
| MS Gothic       | 1.0000 | 0.8600         | **0.0010** | 1.0994        | **0.0006** |
| SimSun          | 0.9994 | 0.8594         | **0.0004** | 1.0987        | **0.0013** |
| Microsoft YaHei | 1.1027 | 0.8505         | 0.1125     | 1.2498        | 0.1498     |
| Malgun Gothic   | 1.1200 | 0.9200         | 0.1620     | 1.2237        | 0.0924     |

The `cell` column is the boundary and it is measurable: where the em box is the
cell, the baseline reading is exact to a thousandth. Where it is not — the
Chinese and the Korean face — neither reading fits, and the fixture records the
residual rather than a rule. `ideographicBaseline` is the only route to that
metric a browser has, and on those two faces Chromium and PowerPoint disagree
about it by 0.11 and 0.16 of the em.

Rivals scored and refuted: the baseline at `1.2 x` the face box's own share,
which is where a **horizontal** line puts it and is 0.11 of the em out on Yu
Gothic; the em box centred rather than the advance box, which is the same
function for a full-width glyph and differs for every other one; the `@` face
having metrics of its own, which GDI denies — `@Yu Gothic` reports the identical
`TEXTMETRIC` to `Yu Gothic`.

**Yu Gothic is 0.039 of the em out across the line and nothing measured explains
it.** It is 0.55pt at 14pt. It is recorded, asserted in a test so that nobody
narrows it away, and left open.

## The insets never turned with the frame, and that was a bug

Laying `eaVert` out meant reading the turn code, and the turn code hands the
laid-out frame the insets as the file states them. ADR 0032 measured that
**the insets do not rotate with the text** — `lIns` insets from the physical left
edge whatever `@vert` says, 7 of 7 — and the layout has to be handed them the
other way round for that to come out true. It was not.

`vert-ins-vert` is in the fixture and has been since 3.6: a 200pt box, insets 13
/ 20 / 3 / 5, block at `dLeft` 175.4. The layout put it at **158.4**. `vert270`
was out by the same construction. Every corpus slide with vertical text either
had symmetric insets or was one of the three that refused to draw, so nothing
caught it: the probes existed, and no test read them.

Now `turnedInsets` sends each edge to the one the quarter turn makes it, and a
test lays out **every** `vert` probe of a drawn direction and compares against
PowerPoint's own corner. 18 of 18 within 0.05pt, and the un-turned reading is
asserted to be the 17pt miss it is.

## `mongolianVert` came free

Its axes were measured in T6 and it stacks rightward from the left where `vert`
stacks leftward from the right — the same quarter turn with the anchor axis
reversed. One row in the turn table, and the same upright rule, and the fixture
places it to 0.05pt.

`wordArtVert` and `wordArtVertRtl` stay refused. They put each glyph on a line of
its own and their column is 1.303 of the em for Arial and 1.167 for Yu Gothic,
which is a font metric nothing has measured; the typed error names it.

## Verification

- **The renderer beside PowerPoint's own bitmap.** `horz` and `vert` come out
  **pixel-identical** — the ink box of `一` at 96pt matches to the quarter point
  on both axes. `eaVert` matches on the run axis exactly and is out across the
  line by the residual the table predicts: −0.25pt on MS Gothic, −3.75pt on Yu
  Gothic. The second glyph, `ト`, agrees independently.
- `pnpm fidelity`: **154 slides scored where 152 were**, 1 not drawn where 3
  were, corpus mean 9767 basis points. `b02-layouts-10` scores 9917 and
  `b02-layouts-11` 9851.
- 3216 → **3256 tests**, 91 files. The frame tests are driven from `frames.json`
  and `vertical-glyphs.json`, never from the implementation.
- **13 of 13 mutants killed**: the inset permutation and its two directions, the
  baseline's ascent and descent, the advance box centred, the cross offset's
  sign, `eaVert`'s turn, `mongolianVert`'s anchor edge, the slot test, the split
  guard, the glyph's own quarter, and the ideographic sign.
- `pnpm check` green.

## Deviations from the plan

| plan item                     | disposition | why                                                                                                              |
| ----------------------------- | ----------- | ---------------------------------------------------------------------------------------------------------------- |
| `eaVert` in v1                | **done**    | Refused since 3.6; two of the three slides it blocked now draw.                                                  |
| Mongolian renders approximate | **exact**   | Its axes were measured in T6 and the upright rule is the same one, so approximating it would have been a choice. |
| WordArt renders approximate   | **refused** | Its per-glyph column is unmeasured. A plausible value is the one thing this repository will not ship.            |

## Open questions

1. **Yu Gothic sits 0.039 of the em off across the line.** Exact on MS Gothic
   and SimSun, so the centring reading is right about something; what makes Yu
   Gothic different is not in its GDI metrics, its font box or its `@` variant.
2. **Microsoft YaHei and Malgun Gothic have a cell of 1.10 and 1.12 of the em**,
   not one, and Chromium's `ideographicBaseline` for them is not the number
   PowerPoint drew with. Chinese and Korean vertical text is out by 0.11 to 0.16
   of the em until that is understood.
3. **No corpus slide scores an upright glyph.** `a10-rtl-cjk-03` carries the only
   one and is still refused, for the `wordArtVertRtl` frame beside it. The
   verification above is a one-off against PowerPoint's bitmap, which is weaker
   than a committed oracle, and closing it means either the WordArt cell or a new
   probe deck.
4. **An underline under upright text throws.** Vertical Japanese draws it beside
   the column rather than under the glyph, and nothing measured which.
5. **Halfwidth katakana is drawn with `ETO_GLYPH_INDEX`** — the one stretch in
   the experiment whose record carries glyph ids. The `@` face has its own
   presentation forms for it, so its advance is not the horizontal one; no probe
   separates that yet.
