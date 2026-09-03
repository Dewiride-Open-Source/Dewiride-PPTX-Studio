# `@pptx-studio/paint`

DrawingML colour and fills: the **six colour bases**, the **twenty-eight transforms** applied in
document order, and resolution against a **theme** and a **colour map** (2.6); then **gradients**
and the **54 preset pattern tiles** (2.7). Dashes, arrowheads and effects are 2.8 and are not here
yet.

Every rule in this package was measured against Microsoft PowerPoint rather than derived from the
standard, because on the question that matters most the standard says nothing at all: ECMA-376
defines `a:tint` as "10% of the input colour combined with 90% white" and never says which colour
space the combining happens in. The credible implementations disagree by up to **73 units out of
255** — on every themed fill in every deck.

```ts
import { resolveColor, toCss } from '@pptx-studio/paint';

// <a:schemeClr val="accent1"><a:lumMod val="60000"/><a:lumOff val="40000"/></a:schemeClr>
const color = {
  space: 'scheme',
  name: 'accent1',
  transforms: [
    { op: 'lumMod', val: 60000 },
    { op: 'lumOff', val: 40000 },
  ],
} as const;

toCss(resolveColor(color, { scheme, map })); // '#8FAADC'
```

That is PowerPoint's own "Accent 1, Lighter 40%" swatch, and `8FAADC` is the number it paints.

## The three spaces

A renderer that does all of this in one colour space is wrong on more than half of the operators.

| space         | transfer                 | operators                                          |
| ------------- | ------------------------ | -------------------------------------------------- |
| linear sRGB   | the sRGB piecewise curve | `tint` `shade` `inv` `red*` `green*` `blue*`       |
| sRGB          | none                     | `gamma` `invGamma` `gray`                          |
| HSL over sRGB | none                     | `hue*` `sat*` `lum*` `comp`                        |
| —             | —                        | `alpha` `alphaMod` `alphaOff` are a fourth channel |

Scoreboard for the first row, worst error against PowerPoint over the 102 `tint` and `shade`
swatches of sub-phase 0.7-C:

| model                                   | worst  | exact   |
| --------------------------------------- | ------ | ------- |
| sRGB piecewise linearisation            | 0/255  | 102/102 |
| power law γ=2.3 (LibreOffice)           | 8/255  | 41/102  |
| Apache POI's asymmetric pair            | 73/255 | 51/102  |
| arithmetic straight on the 0–255 values | 73/255 | 0/102   |

Apache POI is the interesting loser: it linearises `shade` and computes `tint` straight on the sRGB
values, so it is exactly right on half the question.

The middle row of the first table is the strongest evidence in the set. **`a:gamma` _is_ the sRGB
de-linearisation** and `a:invGamma` is its exact inverse — the file format exposes the curve
directly, which independently confirms the curve the top row uses.

## Things that are easy to get wrong, and are pinned by measurement

- **Transforms apply in document order, and order changes the colour.** On `#4472C4`, `lumMod 60%`
  then `lumOff 40%` is `8FAADC`; the reverse is `517CC8`. `tint 50%` then `shade 50%` is `8D93A7`;
  the reverse is `BEC2D1`. A renderer that sorts transforms into a canonical order gets one of every
  such pair badly wrong, and the first pair is PowerPoint's own gallery.

- **Every transform hands the next one a _clamped_ colour.** `lumOff 60%` then `lumOff -60%` is
  `666666`, not the colour you started with: the first drives lightness to 1.12, which paints white,
  and taking 0.6 back off white is mid-grey. Carrying the un-clamped state across the boundary is
  94/255 wrong. An intervening `gamma`/`invGamma` round trip changes nothing, and neither does an
  intervening `alpha`.

- **But the arithmetic _inside_ one transform is not clamped.** `satMod 200%` and `satMod 300%` on
  `#4472C4` are different colours — `0460FF` and `004EFF`. Pinning saturation at 1 gives one answer
  for both.

- **`a:scrgbClr` is linear light, not sRGB percentages.** `r="50000"` is `BC`, not `80`. This is the
  single largest divergence in colour and the name is the only clue: scRGB is a linear-light space.

- **`a:inv` is a complement in _linear_ light** — `#4472C4` inverts to `F8EBB3`, not to the channel
  complement `BB8D3B`. **`a:comp` is something else entirely**: half a turn of hue with saturation
  and lightness untouched.

- **`a:gray` is the Rec.709 luma of the _sRGB_ values**, with no linearisation. `#FF0000` becomes
  `363636`, not `7F7F7F`.

- **`dk1`/`lt1`/`dk2`/`lt2` bypass the `clrMap`; `bg1`/`tx1`/`bg2`/`tx2` and the accents go through
  it.** Every deck ships an identity-ish map, under which both readings agree — which is exactly why
  this bug survives to production. Measured against a master whose map was crossed on purpose.

- **`ST_Percentage` has two spellings and both are accepted.** `val="60000"` and `val="60%"` mean
  the same thing inside the same Transitional part, and `val="60.5%"` works too. `parseInt("150%")`
  is `150`, so a 150% transform silently becomes 0.15%.

- **`a:prstClr`'s 147 names are not CSS colour names.** Not one of them: DrawingML abbreviates
  (`dkBlue`, `ltGray`, `medOrchid`) where CSS spells the words out, so `color: dkBlue` is not the
  wrong colour, it is no colour at all.

## `a:sysClr`, where we differ from PowerPoint on purpose

The plan said `a:sysClr` "prefers `@lastClr`". It does not — 2.6 wrote
`<a:sysClr val="windowText" lastClr="FF00FF"/>` and PowerPoint painted it black, from the machine's
palette, ignoring the attribute. `@lastClr` is what its name says: a cache of what the writer's
machine resolved, for a reader that has no machine to ask.

This package is that reader. A browser has no Windows system palette, and the viewer's theme is not
the author's, so `resolveColor` takes `@lastClr` first and falls back to a table of one Windows 11
light theme. That reproduces what the deck looked like where it was written, which is the closer
answer to what its author meant.

## Gradients

`a:gradFill` is two questions — what curve joins the stops, and where the ramp runs — and
PowerPoint answers both in ways the standard does not mention.

**Some gradients are not linear.** One with **two distinct colours and a stop at each end** goes
through a curve: black to white paints `BABABA` at its midpoint, not `808080`. Anything else is
linear in sRGB.

Note _colours_, not _stops_. `[0 white][50% black][100% white]` has three stops, two colours, and is
curved; `[0 black][50% 808080][100% white]` has the same three positions, three colours, and is
linear. That distinction is not academic — variants 3 and 4 of every gradient in PowerPoint's own
gallery are three stops in two colours, so "two stops" would be the wrong rule for a great many real
decks. The two conditions together are exactly the gradient GDI+ can express as a two-colour brush
with a blend array, which is the only one of its two constructors that gamma-corrects.

The curve blends in a plain **2.2 power** — not the sRGB piecewise transfer, not linear light — and
it ships as a table rather than a formula, because no closed form fits better than eight of 255. The
plan predicted `1 - (1 - t)^1.875`; that is the closest one-parameter fit through the middle of the
range and it misses the dark end by thirteen.

`resolveGradientStops` therefore returns **fifteen** stops where the file had two. SVG interpolates
stop colours linearly in sRGB and `color-interpolation` is not implemented for gradients anywhere,
so pre-sampling is the only way to draw the curve at all. Fifteen adaptive knots hold PowerPoint's
own export to 1.9 of 255 in both directions.

**Where the ramp runs.** `@ang` is clockwise from +x with y down. `@scaled="1"` bends the angle
_towards_ the long axis — `atan2(w·sin a, h·cos a)` — because what the unit square stretches is the
bands, not the arrow; on a 3:1 shape a written 45° renders at 71.6° and the inverted reading predicts
18.4°. The ramp spans the shape's full projection onto that direction, so the extreme corner is
exactly the first stop.

Use `linearGradientVector` and emit `gradientUnits="userSpaceOnUse"`. `objectBoundingBox` is wrong
for **both** values of `@scaled`: it reproduces the bend but shears the ramp, putting the corner of a
3:1 shape at 0.21 where PowerPoint puts it at 0.01.

**Path gradients.** Stop 0 sits at the focus and the last at the edge — the same way round as SVG, so
nothing is reversed. `a:fillToRect` behaves as a point at its centre; its extent is not a flat
region. `path="circle"` is a true circle reaching its last stop at the corner, not an ellipse fitted
to the box. `path="rect"` is Chebyshev normalised per axis to the edge the ray faces.

## Patterns

The 54 `ST_PresetPatternVal` tiles are **measured**, not transcribed. PowerPoint was asked to author
one shape per `MsoPatternType`; it accepted exactly 54 and spelled each name itself, and the pixels
were read from a 96 DPI export.

The tile is eight pixels at 96 DPI — 6 pt, 76200 EMU — and that is a fixed **physical** size. Export
the same shape at 640, 1280, 1920 and 2560 pixels wide and the period is 4, 8, 12, 16. So
`patternUnits="userSpaceOnUse"` with a size in slide units; `objectBoundingBox` would make a wide
shape's hatching coarser than a narrow one's.

Tiles are stored as coverage rather than bits, because three of them — `dnDiag`, `upDiag`,
`diagCross` — are drawn as antialiased lines. `lgGrid` and `cross` measure byte-identical, which is
the only such pair and matches GDI+ aliasing `LargeGrid` to `Cross`. The `pctNN` names are labels:
`pct75` covers 88% of its tile and `pct20` covers 13%.

A missing `a:fgClr` paints black and a missing `a:bgClr` paints white, whatever the theme.

## What is not here

- Dashes, arrowheads and effects — 2.8.
- `a:blipFill` and `a:grpFill` have a place in the `Fill` union and nothing behind them. A `grpFill`
  child has to show the _slice_ of the enclosing group's gradient that falls under it, which needs
  the group's rectangle: 2.10.
- `path="shape"` falls back to the bounding rectangle. Following a real outline needs
  `@pptx-studio/geometry`, which is deliberately not a dependency of this package.
- The `clrMapOvr` chain and `phClr`'s value: this package takes a `ClrMap` and a `phClr` and does
  not know how to find them. That is the resolver in 2.9.
- Chart and diagram colour styles (`colors1.xml`, `cs:variation`) — Phase 9.
- Anything that composites. `Rgba.a` is opacity; what to do with it is the renderer's decision.

## Ground truth

`corpus/ground-truth/color-transforms.json` — 214 swatches, sub-phase 0.7-C, the transforms.
`corpus/ground-truth/color-bases.json` — 259 swatches, sub-phase 2.6, the bases, alpha, the colour
map and the percentage grammar.
`corpus/ground-truth/fills.json` — 144 probes in 23 packages, sub-phase 2.7, gradients and patterns.

The colour fixtures were read back two independent ways — PowerPoint's object model and the centre
pixel of its own bitmap export — and the two agree on all 464 opaque swatches. `color.test.ts` runs
every one of them through `resolveColor`; it is exact on 454 of the 455 it predicts, and off by one
unit in one channel on the last, which lands on an exact rounding tie.

A gradient has no single colour, so for C3 the bitmap is the oracle and the object model only
corroborates. `fill.test.ts` predicts every ramp in that fixture to within two bytes, with two named
exceptions where PowerPoint's rasteriser softens a corner in the ramp profile.

See `docs/adr/0007-ground-truth.md`, `docs/adr/0021-colour.md` and `docs/adr/0022-fills.md`.
