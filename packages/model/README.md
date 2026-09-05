# @pptx-studio/model

The document model: the sheets, the chain between them, and the resolver everything else reads.

A `.pptx` renders correctly or it does not, and almost all of the difference is here rather than in
any renderer. A slide placeholder that PowerPoint wrote has no geometry, no fill, no line and no
style of its own; every one of those comes down a chain of two hops whose rules are asymmetric,
undocumented, and different from what every implementation this project has read assumes.

Sub-phase 2.9. Everything below was measured against Microsoft PowerPoint rather than argued from
the standard; the fixture is `corpus/ground-truth/sheets.json` and the reasoning is
`docs/adr/phase-2-geometry-and-paint/0024-model-parse-and-resolve.md`.

## `undefined` means the file did not say

Every visual property is `T | undefined`, and absence is a fact the file states rather than a gap to
be filled in.

PowerPoint proves this about its own output. Create a slide from a layout and its title placeholder
has no `a:xfrm` at all. Type into it and it still has none. Nudge it one point to the right and the
whole resolved rectangle appears at once — inherited two hops, from the master, baked in:

```xml
<!-- untouched, and after typing -->
<p:sp><p:nvSpPr>…<p:ph type="title"/>…</p:nvSpPr><p:spPr/>…</p:sp>

<!-- after moving it one point -->
<p:sp>…<p:spPr><a:xfrm><a:off x="850900" y="365125"/><a:ext cx="10515600" cy="1325563"/></a:xfrm>…
```

So a parser that defaults geometry at parse time destroys the only record of which shapes are still
bound to their layout. That record is what Change Layout reads, what the inspector's provenance
chips read, and what tells a theme swap which shapes to invalidate.

## The two hops do not use the same key, and neither uses the pair

This is the finding of the sub-phase, and it contradicts the plan that led to it.

| hop             | matches on                                   | when nothing matches |
| --------------- | -------------------------------------------- | -------------------- |
| slide → layout  | **`@idx` alone.** `@type` is never consulted | orphan               |
| layout → master | **the folded `@type`**, taking the first     | orphan               |

Measured over 46 probes, each a slide placeholder with no geometry of its own whose rendered
position names the layout placeholder it matched. Scored against the 34 that had a candidate set:

| candidate rule                    | right |
| --------------------------------- | ----- |
| `@idx` alone                      | 34/34 |
| family (title/content) and `@idx` | 24/34 |
| `(type, idx)`                     | 13/34 |
| `@type` alone                     | 6/34  |

A slide `title` at idx 0 takes a layout `body` at idx 0. A slide `body` at idx 2 takes a layout
`sldNum` at idx 2. A slide `title` at idx 9, against a layout that holds a title at idx 0, matches
**nothing** and renders at the origin with zero size. No type is privileged — not the title family,
not the header-and-footer trio, which is what the shuffled-type deck was built to find.

The asymmetry is not arbitrary:

- Every stock PowerPoint layout numbers its date, footer and slide-number placeholders **10, 11 and
  12**, and every stock master numbers the same three **2, 3 and 4**. A second hop that looked at
  `@idx` would inherit nothing in the templates Office ships.
- A slide master's placeholder vocabulary is only `title`, `body`, `dt`, `ftr`, `sldNum` and `hdr`.
  A master carrying `ctrTitle`, `subTitle`, `obj` or `pic` is **repaired on open**. So a second hop
  that looked at the type as written would find nothing for the `ctrTitle` in layout 1 or the
  `<p:ph idx="1"/>` in "Title and Content".

`ctrTitle` folds to `title` and every content type folds to `body` on the way up. And each hop asks
with the placeholder of the sheet it is _leaving_: a slide `obj` at idx 0 reaching a layout
`ctrTitle` lands on the master's **title**, not its body.

`<p:ph/>` with no attributes is `type="obj" idx="0"` — read back through PowerPoint's own object
model as `ppPlaceholderObject`. ECMA's default for the attribute is `body`. The two agree at the
first hop because it ignores the type, and at the second because `obj` folds to `body`; they agree
by construction rather than by luck, which is worth knowing.

## Every binding comes from the part's own `.rels`

A slide part does not name its layout, a layout does not name its master, and a master does not name
its theme. Each binding lives in exactly one place, and every other candidate is a trap:

- **`p:sldLayoutIdLst` is not the binding.** A two-master deck saved by PowerPoint puts all 22
  layouts in one flat `ppt/slideLayouts/` folder numbered 1 to 22, and which master owns which is
  recorded only in the two masters' relationship parts.
- **`ppt/presentation.xml.rels` has a `theme` relationship**, it always points at `theme1.xml`, and
  reading it gives every slide the first master's palette. On a one-master deck — which is every
  deck most people ever test with — that is indistinguishable from correct. Measured: two masters,
  two themes, and `accent1` resolves to a different colour on each.
- **A `.rels` file is not in `rId` order.** PowerPoint routinely writes `rId8` first.

A missing binding is not an error. PowerPoint repairs a slide with no layout, a layout with no
master and a master with no theme rather than refusing them, so `loadDocument` records a problem and
leaves `parent` or `theme` as `null`. Refusing to _write_ one is `@pptx-studio/validate`'s job.

## The style matrix, `phClr`, and the background share one rule

`a:fillRef` and `p:bgRef` index the same two lists with the same offset, and they were measured
independently against the same six theme entries. All six agreed.

| `@idx`       | means                          |
| ------------ | ------------------------------ |
| `0`          | nothing at all                 |
| `1`…`999`    | `a:fillStyleLst[idx − 1]`      |
| `1000`       | nothing at all                 |
| `1001`…      | `a:bgFillStyleLst[idx − 1001]` |
| out of range | clamps to the last entry       |

Out of range **clamps**: a `fillRef idx="4"` against a three-entry list paints the third, and a
`bgRef idx="9999"` paints the last background entry. PowerPoint opens both without repairing, so a
renderer that throws refuses a deck PowerPoint shows.

Each entry is a _function_ of a colour. `<a:solidFill><a:schemeClr val="phClr"><a:lumMod
val="60000"/></a:schemeClr></a:solidFill>` is the body and the reference is the call, with the
reference's own colour child as the argument — and that argument may carry transforms of its own.
PowerPoint's shape-style gallery writes `<a:schemeClr val="accent1"><a:shade val="15000"/></a:schemeClr>`
on seven of its forty-two entries. Miss `phClr` and every themed shape renders black.

`a:fontRef/@idx` is **not** a number. It is `major`, `minor` or `none`, and PowerPoint repairs a
numeric one.

The background takes the nearest sheet that declares a `p:bg` — slide, else layout, else master. An
absent `p:bg` is the only thing that inherits: a slide declaring `<p:bgPr><a:noFill/>` paints nothing
and does _not_ fall through, which is why `Background` is a value rather than a nullable fill.

## `masterClrMapping` means _parent_, not master

`a:masterClrMapping` says "the map already in force". For a slide, that is the **layout's** map. A
layout carrying an `a:overrideClrMapping` changes the colours of every slide bound to it, and a
slide's own override beats its layout's — measured with two overrides that disagreed, because two
identical ones cannot tell you which applied. `dk1`/`lt1`/`dk2`/`lt2` bypass the map either way, as
2.6 established.

And the map that resolves a `p:bgRef`'s colour is the one in force on the sheet being _asked about_,
not on the sheet the `bgRef` was written on. A slide whose `clrMapOvr` remaps `bg1` gets a different
background out of the same master.

## The resolver reports where it looked

```ts
resolve(shape, sheet, (s) => s.xfrm);
// → { value, origin: 'masterPh', explicit: false, sheet, shape }
```

Carrying the origin costs one field and buys four features from one function — provenance chips,
per-property Reset, layout-compatibility scoring, and correct theme invalidation — all reading the
same function the renderer reads, so none of them can drift from what is on the screen. `explicit`
is not the same as `origin === 'shape'`: a layout placeholder that declares a fill is explicit about
that fill, and the slide inheriting it is not.

Geometry is not special. Fills, lines, effects and `p:style` all travel the same chain, and all four
were measured arriving from the layout placeholder and from the master placeholder through a layout
that declared nothing.

## What is not here

- Nothing is painted. The renderers are 2.10.
- The ten-source **text** cascade is 3.1. `Origin` already declares the five members only that
  cascade can produce, so 3.1 adds cases rather than widening a type every consumer has switched on.
- Commands, undo and history are Phase 5; Change Layout is 7.4. What PowerPoint rewrites when a
  layout changes is recorded in the fixture and not acted on.
- `a:effectDag` is not modelled — a directed graph of effect primitives PowerPoint has never been
  observed to write. An `undefined` there means the shape re-emits byte for byte.
- Group child coordinate spaces are parsed (`a:chOff`/`a:chExt`) and not applied; that is 2.10.

## Fixtures

`corpus/ground-truth/sheets.json` — 114 probes in 37 packages, sub-phase 2.9.

The measurement is a **position**, and PowerPoint reports it directly. C3 and C4 sampled bitmaps
because a fill and a stroke are pictures; an inheritance is not. A placeholder with no `a:xfrm` of
its own still has a position, `Shape.Left` reports it in points, and that is the resolver's own
answer read out of the resolver rather than reconstructed from what it painted. So every candidate
parent sits in a rectangle no other candidate shares, and the rectangle a probe lands in names the
parent it matched — one number, no fitting, no error bars.

23 of the 37 packages are hostile and each is alone in its own file. 16 were refused or repaired.

`model.test.ts` re-derives the matcher from the fixture's 34 recorded cases rather than comparing
against a summary of them, so a rule that drifts fails on the measurements.

See `docs/adr/phase-0-foundation/0007-ground-truth.md`, `docs/adr/phase-2-geometry-and-paint/0021-colour.md`, `docs/adr/phase-2-geometry-and-paint/0022-fills.md`,
`docs/adr/phase-2-geometry-and-paint/0023-lines.md` and `docs/adr/phase-2-geometry-and-paint/0024-model-parse-and-resolve.md`.
