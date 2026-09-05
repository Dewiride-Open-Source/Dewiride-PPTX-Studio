# 0024 — Model: parse and resolve

Status: accepted
Sub-phase: 2.9
Date: 2026-09-04
Fixture: `corpus/ground-truth/sheets.json`
Supersedes nothing. **Overturns the approved plan's placeholder matcher** — both
the plan's 2.9 line and the five tiers it hardens into in 7.1 — and adds two
package-level rules the plan does not mention.

## Context

Everything up to here paints one shape from markup that says what to paint. This
sub-phase is about the markup that says nothing.

A slide placeholder that PowerPoint wrote has no geometry, no fill, no line and
no style. Not "has defaults" — has _nothing_. Every one of those arrives down a
chain of two hops, and the rules for those hops are not in ECMA-376 in any form
you could implement from. §19.3.1.36 defines `p:ph` and its attributes; it does
not say what matching means, which attribute is the key, or what happens when
nothing matches.

So every implementation guesses, and they all guess the same way: match on
`(type, idx)`, with a ladder of fallbacks for titles and for the
date/footer/slide-number trio. The plan for this project guessed that too, in
more detail than most — five tiers, and a 121-case matrix in 7.1 to harden them.

That guess is wrong, and this is the sub-phase that found out.

## The experiment

Two halves, cheapest first, as in 2.7 and 2.8.

**PowerPoint authored first.** `tools/ground-truth/model/author.ps1` drives the
object model to _make_ the structures and then reads the files it saved. One
script, and it settled the sub-phase's central architectural question plus four
others before a probe existed:

- **Three slides on one layout: untouched, typed into, and nudged one point.**
  Only the nudged one has an `a:xfrm`, and it has the _whole_ resolved rectangle
  — `x="850900"`, which is the master's `838200` plus one point, with the
  master's `cx` and `cy` alongside. Typing text does not create geometry. This is
  architectural bet 3 confirmed by the format's own author.
- **The eleven stock layouts and their master.** Every layout numbers its date,
  footer and slide-number placeholders `10`, `11`, `12`; the master numbers the
  same three `2`, `3`, `4`. Three of the eleven layouts carry a placeholder with
  **no `@type` at all** — `<p:ph idx="1"/>` — including "Title and Content", the
  most common layout in the world.
- **`Shape.ShapeStyle` over its whole range.** Forty-two theme styles, whose
  `p:style` elements use `fillRef`/`lnRef` indices 0–3 and `effectRef` 0–3, and
  whose `a:fontRef/@idx` is `minor` — a _name_, never a number.
- **A second design.** Two masters, 22 layouts in one flat folder numbered 1–22,
  and `ppt/presentation.xml.rels` carrying a `theme` relationship to `theme1.xml`
  that is a trap for anyone who reads it.
- **A slide with no background opinion carries no `p:bg` at all**, and
  `FollowMasterBackground = False` writes `p:bgPr`, never `p:bgRef`.

**Then the probes.** 114 of them across 37 packages
(`tools/ground-truth/model/probes.ts`), read back through
`tools/ground-truth/model/read.ps1` and reduced by `tools/ground-truth/model/analyse.ts`.

## The measurement is a position, and it is free

C3 and C4 sampled bitmaps because a fill and a stroke are pictures: no COM
property says what colour is at a point. An inheritance is not a picture.

A placeholder with no `a:xfrm` still has a position. PowerPoint knows it.
`Shape.Left` reports it in points, exactly, with no antialiased edge to
interpolate through. That is the resolver's own answer read out of the resolver
rather than reconstructed from what it painted.

So every candidate parent was given a rectangle no other candidate shares — five
bands of seven boxes, 80 points apart in `y` and 130 in `x` — and **the box a
probe lands in names the parent it matched**. One number, no fitting, no error
bars, and the whole experiment runs in about a minute. Bitmaps are still exported
for a human to look at; nothing depends on them.

## What the probes said

### The first hop matches on `@idx` alone

Not `(type, idx)`. Not a family. Not a ladder of tiers. `@type` is not consulted
at the slide-to-layout hop at all — not as a key, not as a tiebreak, not as a
fallback.

Scored over the 34 first-hop cases:

| candidate rule                  | right |
| ------------------------------- | ----- |
| `@idx` alone                    | 34/34 |
| title/content family and `@idx` | 24/34 |
| `(type, idx)`                   | 13/34 |
| `@type` alone                   | 6/34  |

The cases that separate them are not exotic:

- A slide `title` at idx 0 takes a layout **`body`** at idx 0.
- A slide `body` at idx 2 takes a layout **`sldNum`** at idx 2.
- A slide `ftr` at idx 0 takes a layout **`body`** at idx 0.
- A slide `title` at idx 9, against a layout holding a title at idx 0, matches
  **nothing** and renders at the origin with zero size.

That last one kills the plan's tier 2 ("if title, any title regardless of idx")
and tier 3 ("if `sldNum`/`dt`/`ftr`, raw type only ignoring idx") outright. Both
were probed directly, with a stray `@idx` on each of the four types, and all four
orphaned.

A deck built by PowerPoint never exercises this, because PowerPoint writes the
slide placeholder as an exact copy of the layout's. It is Change Layout — the
verb this whole product exists for — that puts a slide placeholder against a
layout placeholder it does not match.

### The second hop matches on the folded type, first wins

Layout to master is `@type` only, and `@idx` is not even a tiebreak: a layout
`body` at idx 5, against a master holding bodies at idx 1 **and** idx 5, takes
the one at idx 1.

`ctrTitle` folds to `title`; every content type folds to `body`. Measured on
`pic`, `obj`, `subTitle` and `ctrTitle`.

And **each hop asks with the placeholder of the sheet it is leaving**. A slide
`obj` at idx 0 reaching a layout `ctrTitle` lands on the master's _title_, not
its body — which is the discriminating case, because using the slide's own type
would fold to `body` and land elsewhere.

### Why the asymmetry is not arbitrary

Two facts about masters make a symmetric matcher impossible:

- Every stock layout numbers dt/ftr/sldNum 10/11/12 and every stock master
  numbers them 2/3/4. An `@idx`-keyed second hop inherits nothing in the
  templates Office ships.
- **A slide master's placeholder vocabulary is `title`, `body`, `dt`, `ftr`,
  `sldNum` and `hdr`, and nothing else.** A master carrying `ctrTitle`,
  `subTitle`, `obj` or `pic` is repaired on open. So a type-keyed second hop that
  did not fold would find nothing for the `ctrTitle` in layout 1 or the
  `<p:ph idx="1"/>` in "Title and Content".

That second rule was found the expensive way. Four otherwise-clean probe decks
came back REPAIRED, and the readings from a repaired deck are readings of a file
PowerPoint rewrote. Bisecting eight variants localised it to one attribute value.

### A bare `<p:ph/>` is `obj` at idx 0

Read back through the object model as `ppPlaceholderObject`. The plan's 2.9 line
says `obj`; its 7.1 line says `type ??= 'body'`. The plan contradicts itself and
2.9 is the half that is right.

The two agree in practice — at the first hop because it ignores the type, at the
second because `obj` folds to `body` — but they agree by construction rather than
by luck, and that is worth knowing before someone relies on it.

### An orphan renders at the origin with zero size

Not refused, not dropped, not given a default rectangle. Four probes: a title
where the layout has none, a body at an index nobody has, and both against a
layout with no placeholders at all.

### `a:fillRef` and `p:bgRef` are one rule over one table

Measured independently against a theme whose six style entries paint six
distinguishable colours. All six agreed:

| `@idx`       | means                          |
| ------------ | ------------------------------ |
| `0`          | nothing at all                 |
| `1`…`999`    | `a:fillStyleLst[idx − 1]`      |
| `1000`       | nothing at all                 |
| `1001`…      | `a:bgFillStyleLst[idx − 1001]` |
| out of range | clamps to the last entry       |

The plan states the 1000-offset for `bgRef` and says nothing about `fillRef`;
they are the same rule and the fixture now says so from both directions.

Out of range **clamps** rather than failing — `fillRef idx="4"` against a
three-entry list paints the third, `bgRef idx="9999"` paints the last background
entry, and both packages open with no repair. A renderer that throws there
refuses a deck PowerPoint shows.

### `phClr` is a fully transformed colour

The entry is a function of a colour and the reference is the call. The argument
may carry its own transforms: PowerPoint's own gallery writes
`<a:schemeClr val="accent1"><a:shade val="15000"/></a:schemeClr>` on seven of its
forty-two shape styles. Measured on three different scheme colours through the
same entry, on a literal `srgbClr`, and on a transformed `schemeClr`.

### Fills, lines and `p:style` descend the same chain as geometry

All three arrive from the layout placeholder, and all three arrive from the
**master** placeholder through a layout placeholder that declares nothing. An
explicit `spPr` fill beats an inherited one; an explicit `a:noFill` beats it too.
A shape with no fill, no style and no placeholder paints nothing — not black.

### The background takes the nearest sheet that has an opinion

Slide, else layout, else master. An absent `p:bg` is the only thing that
inherits: a slide declaring `<p:bgPr><a:noFill/>` paints nothing and does not
fall through. That is why `Background` is a value and not a nullable fill.

### `masterClrMapping` means _parent_, not master

A layout carrying an `a:overrideClrMapping` changes the colours of every slide
bound to it that says `masterClrMapping`, and a slide's own override beats its
layout's. Probed with two overrides that **disagree**, because the first attempt
used two identical ones and could not tell which had applied.

And a `p:bgRef`'s colour resolves through the map in force on the sheet being
asked about, not on the sheet the `bgRef` was written on: the same master
background comes out differently on a slide whose `clrMapOvr` remaps `bg1`.

### Every master owns its theme part

Two masters pointing at the same `theme1.xml` is repaired on open; six masters
with six themes is not. Found by bisection after a six-master probe deck came
back repaired.

And the theme a slide resolves against is its **master's**, reached through its
own chain — not the one `ppt/presentation.xml.rels` names, which always points at
`theme1.xml` and is right on exactly the decks that have one master.

## Decision

`@pptx-studio/model` at layer 2, with:

- `parseSheet` / `parseTheme`, which default nothing the file could have said.
- `matchInLayout` (by `@idx`) and `matchInMaster` (by folded `@type`, first
  wins), and `inheritanceChain`, which asks each hop with the placeholder of the
  sheet it is leaving.
- `resolve(shape, sheet, pick)`, returning `{ value, origin, explicit, sheet,
shape }`. One function; provenance, Reset, layout scoring and theme
  invalidation all read it.
- `styleMatrixTarget` / `styleMatrixFill` / `styleMatrixLine` / `phClrOf`, and
  `resolveAppearance`, which searches the `spPr` route down the whole chain
  before consulting the `p:style` route.
- `resolveBackground`, sharing the matrix and the offset with the above.
- `loadDocument`, which builds every binding from the owning part's own `.rels`
  and records a broken one as a problem rather than throwing.

`Origin` declares all eight members, including the five only the 3.1 text cascade
can produce, so that sub-phase adds cases rather than widening a type every
consumer has already switched on.

## Consequences

**The plan's 7.1 changes.** The five-tier matcher and its 121-case matrix were
designed against a rule that is not the rule. What 7.1 should harden is the
`@idx` match plus the fold, and its matrix is still worth building — as a
regression net over the eleven stock layouts crossed with themselves, which is
exactly the configuration where a wrong matcher looks right.

**Change Layout gets simpler and more honest.** `matchQuality` in 7.2 was to be
scored from which tier fired. There is one tier. Quality is now a property of the
_target_: whether the layout has a placeholder at that index at all, and whether
its type is the same family — the second being advisory, since PowerPoint will
bind them regardless.

**Two new rules for `validate` in 1.2's firewall**, both repair-inducing and
neither in the appendix: a master's placeholder vocabulary, and one theme part
per master.

**`a:blipFill` is still unassigned.** It has a place in paint's `Fill` union and
nothing behind it, and Gate 2 asks for cropped image fills. Flagged again.

## Open questions

- Does the first hop's "first wins" for duplicate indices hold on a _slide_ with
  two placeholders at one index, or only on a layout? Probed on the layout side
  only.
- What `showMasterSp="0"` hides. The plan says non-placeholder master shapes only
  (7.9); not probed here, because master shapes do not appear in `Slide.Shapes`
  and it needs the bitmap.
- Whether an orphan's zero rectangle is a rendering decision or a resolution one
  — i.e. whether PowerPoint's own UI would let you move it.
- What a `p:ph` on a _master_ shape inherits from, if anything.
- Whether `@sz` and `@orient` affect anything at all. Both are recorded and
  neither participates in matching; no probe found a use for either.
- Whether the second hop's fold is exactly "anything a master may not carry
  becomes `body`", or whether `sldImg` and `dgm` behave differently. Neither was
  probed; both are inferred.
