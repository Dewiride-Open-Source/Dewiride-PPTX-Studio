# 0021 — Colour, and the second time we asked PowerPoint

Date: 2026-09-02
Status: **accepted** — 473 measured swatches across two experiments, 455 of them predicted by this
package and 454 of those exactly
Sub-phase: 2.6 — colour

---

## Context

Phase 2 has produced shapes with no colour in them since 2.1. `packages/paint` is where that stops.
2.6 is the colour part of it: the six things a `EG_ColorChoice` can be, the twenty-eight transforms
that can hang off one, and what a theme and a colour map do to a name. Gradients, patterns, dashes,
arrowheads and effects are 2.7 and 2.8.

Sub-phase 0.7-C had already settled the part of this that was actively contested. ECMA-376 defines
`a:tint` as "10% of the input colour combined with 90% white" and never says which colour space the
combining happens in; LibreOffice uses a power law with γ=2.3, Apache POI the sRGB piecewise curve
for `shade` and no curve at all for `tint`, and doing the arithmetic on the bytes is a third answer.
214 swatches in front of PowerPoint settled it: **sRGB piecewise linearisation, exact on 102 of
102**, with the losers 8, 73 and 73 units out of 255.

What 0.7-C did not ask is what this sub-phase ran into immediately. Every one of its swatches was
built on `a:srgbClr` or `a:schemeClr`, deliberately, so that a disagreement could be attributed to
the transform rather than to the base. That left four of the six bases, all three alpha operators,
any non-identity `clrMap` and the whole percentage grammar unmeasured — and a renderer cannot ship
"probably linear" for a base that appears in files.

So 2.6 ran a second experiment against the same PowerPoint, using the same committed machinery.

---

## The second experiment, and the design change it forced on itself

**C2 is 259 swatches in 21 packages, one question per package.**

The first run was six decks grouped by topic. One came back **repaired**, with the only sentence
PowerPoint ever says — _"The file or directory is corrupted and unreadable"_ — and no indication
which of sixty-seven swatches caused it. A repaired deck is not a measurement: PowerPoint may have
rewritten anything in it, and the readback then describes the repair.

Splitting to one question per package made the refusal name its own cause in a single run: **20 of
the 21 opened with no repair, and the one that did not was `hsl-range`** — two swatches writing an
`a:hslClr/@hue` outside `ST_PositiveFixedAngle`. Those two swatches carry `fromRepairedDeck: true`
in the fixture and are excluded from every assertion, which the test asserts rather than assumes.

That is a small process finding worth keeping: when the oracle's only diagnostic is "no", the unit
of experiment has to be the unit of the question.

Both experiments were read back two independent ways — `Shape.Fill.ForeColor.RGB` from the object
model, and the centre pixel of PowerPoint's own bitmap export. **They agree on all 464 opaque
swatches across the two.** Where they could differ the pixels win, because they are what a user
sees; the fixtures record the pixels. Alpha is the one thing only the object model can answer: a
translucent swatch's pixel is a composite with the swatch behind it.

---

## Decisions

### 1. The base union stays lazy

`Color` keeps the space it was written in — `srgb`, `scrgb`, `hsl`, `scheme`, `sys`, `prst` — and
`resolveColor` is the only place it becomes a number. Flattening at parse time is the same mistake
`xfrm === undefined` exists to avoid one layer up: it destroys the fact that the deck said
`accent1`, and with it the palette swap in 7.6, the inspector's provenance chip in 6.5, and
re-resolution when a `clrMapOvr` changes.

### 2. Values are in the units the file uses, and there are two grammars

A percentage transform carries hundred-thousandths and an angle sixtieths of a degree, because that
is what the attribute says. The string half is `parsePercentage`, and it has to accept **both**
spellings: `val="60000"` and `val="60%"`.

That is not a theoretical concern about Strict-versus-Transitional. 2.6 wrote
`<a:lumMod val="60%"/><a:lumOff val="40%"/>` into a Transitional slide; PowerPoint opened it with no
repair and painted `8FAADC` — the same colour, to the byte, as the `60000`/`40000` control beside
it. `val="60.5%"` works too, which the integer form cannot express at that precision. The failure
this prevents is `parseInt("150%") === 150`, which silently turns 150% into 0.15%.

### 3. Three spaces, per the measurement

`tint`, `shade`, `inv` and the nine per-channel operators in **linear light**; `gamma`, `invGamma`
and `gray` on the **sRGB** values as they stand; the hue, saturation and lightness families and
`comp` in **HSL over sRGB**; the three alpha operators in no colour space at all. A renderer that
picks one space is wrong on more than half of them.

`a:gamma` **is** the sRGB de-linearisation and `a:invGamma` its exact inverse — the format exposes
the curve directly, which is the strongest independent confirmation that 0.7-C picked the right one.

### 4. Every transform hands the next one a clamped colour

**This is the finding that changed the design, and it arrived after the module was written and
passing.**

The first implementation grouped consecutive operators in the same space into one excursion, on the
theory that converting per operator was waste. It reproduced all 214 of 0.7-C's swatches. It is
wrong by 94 units out of 255:

| chain on `#4472C4`                  | PowerPoint | carrying the state |
| ----------------------------------- | ---------- | ------------------ |
| `lumOff 60%` then `lumOff -60%`     | `666666`   | `4472C4`           |
| `satMod 300%` then `satMod 33.333%` | `556FAA`   | `4472C4`           |

The first chain drives lightness to 1.12, which paints white. Taking 0.6 back off _white_ is
mid-grey — not the colour you started with. 0.7-C could not see this because no chain in it drove a
value out of range and then back.

Two things that could have made the boundary invisible were probed and do not: an intervening
`a:gamma`/`a:invGamma` round trip changes nothing, and neither does an intervening `a:alpha`.

A stronger hypothesis — that the intermediate is quantised to eight bits per channel between
transforms — was fitted against all 269 predictable swatches and **loses**: 256 exact against 266
for a float clamp. So the clamp is to the range, not to the byte.

### 5. But the arithmetic inside one transform is not clamped

0.7-C's finding survives intact and is not in tension with the one above. `satMod 200%` and
`satMod 300%` are different colours — `0460FF` and `004EFF` — because a saturation of 1.56 is
meaningful _while the conversion back to RGB is happening_. It stops being meaningful when that
conversion finishes. Pinning saturation at 1 first gives one answer for both.

### 6. `a:scrgbClr` is linear light

`r="50000"` is `BC`, not `80`. **14 of 14 exact under the linear reading, 5 of 14 under the
percentage reading** — and the five are the ones where the two agree, 0% and 100%.

This is the single largest divergence in colour: 61 units out of 255 at the midpoint of every
channel. The name is the only clue the format gives, and it is a clue everyone ignores because
`r="50000"` reads so obviously as "50%". Out-of-range values clamp before de-linearising, which two
swatches confirm.

### 7. `a:hslClr` is the same bi-hexcone the `lum` and `sat` transforms use

So a deck can express a colour either way and get the same answer. 16 of 16 exact apart from the two
rounding ties in decision 11.

The plan's note to "force `S=0` when `L` hits 0 or 100%" needs no code: at `l=0` the formula gives
`q=0, p=0` and at `l=1` it gives `q=1, p=1`, so every channel collapses on its own. Measured —
`lum="0"` is `000000` and `lum="100000"` is `FFFFFF` at a saturation of 52%.

### 8. `a:prstClr` is a measured table of 147 names

Not a transcribed one. Every name got a swatch and the deck opened with no repair, which is how we
know the enumeration carries **both** the `Gray` and the `Grey` spelling of the seven grey names —
nobody on this project has read the schema. The table carries nobody's copyright, because it is a
measurement of a program.

The plan says these "are not CSS colour names" and that is right, though not for the reason a value
table would suggest: **not one of the 147 is a name a browser accepts**. DrawingML abbreviates —
`dkBlue`, `ltGray`, `medOrchid`, `ltGoldenrodYellow` — so `color: dkBlue` is not the wrong colour,
it is no colour at all. On values, `gray` is `808080`, the HTML value rather than X11's `BEBEBE`;
whether any of the other 146 differs from the X11 value of its expanded name is unmeasured, because
this package does not need to know.

### 9. `a:sysClr` prefers `@lastClr`, and PowerPoint does not

**The plan is wrong here, and 0.7-C recorded it as confirmed when it was a coincidence.**

2.6 wrote three swatches with a `@lastClr` no system theme would produce —
`<a:sysClr val="windowText" lastClr="FF00FF"/>` — and PowerPoint painted all three the _machine's_
colour, ignoring the attribute. 0.7-C only looked like a confirmation because the theme's
`windowText` had a `lastClr` of `000000`, which is also what this machine says. Two answers, one
number.

`@lastClr` is what its name says: a cache of what the writer's machine resolved, for a reader with
no machine to ask.

This package is that reader, so it takes `@lastClr` first anyway, and falls back to a committed
table of one Windows 11 light theme. That is a deliberate divergence from PowerPoint and it is the
right one: a browser has no Windows system palette, and the viewer's theme is not the author's, so
the cached value reproduces what the deck looked like where it was written. Asserted in the test as
a divergence — the same rows are checked to equal `lastClr` and to differ from what PowerPoint
painted — so it cannot be quietly "fixed".

### 10. `dk1`/`lt1`/`dk2`/`lt2` bypass the `clrMap`

Measured against a master whose map was crossed on purpose: `bg1="dk1"`, `tx1="lt1"`,
`accent1="accent2"`. Under it, `bg1` painted black and `accent1` painted `accent2`'s orange, while
`dk1` and `lt2` were untouched. On the identity-ish map every real deck ships, both readings give
the same answer — which is exactly why this bug survives to production and why it needed a master
written for the purpose.

### 11. The value written is a byte, and the tie-breaking is a fit

Across the 269 swatches whose value a correct model predicts, seven land on **exactly** `x.5` in
exact arithmetic. PowerPoint takes the lower byte four times and the higher one three times:

| value | PowerPoint | where                               |
| ----- | ---------- | ----------------------------------- |
| 127.5 | 127        | `#FF0000` `lumMod 50%`              |
| 127.5 | 127        | `#FF0000` `lumOff 25%`              |
| 153.5 | 153        | `#808080` `lumOff 10%`              |
| 93.5  | 93         | `hslClr` hue 218, sat 100%, lum 50% |
| 110.5 | 111        | `hslClr` hue 218, sat 50%, lum 50%  |
| 129.5 | 130        | `#4472C4` `satOff -50%`             |
| 134.5 | 135        | `#4472C4` `satOff -50%`             |

Half-down fits four, half-up three, half-even one. Holding the intermediate in single precision does
not explain it either — three of the seven move to the _wrong_ side under float32. Whatever
PowerPoint does here is not a rounding mode applied to this number.

`toByte` breaks a tie towards **odd**, which fits six of the seven and agrees with half-down on
every case 0.7-C had, so nothing that experiment settled is disturbed. The one it misses — `129.5`,
where PowerPoint paints 130 — is asserted in the test rather than smoothed away.

0.7-C concluded "PowerPoint takes the lower value" from three samples and said so cautiously. Four
more samples were enough to show it was a coincidence. That is an argument for measuring more of
them rather than for trusting the rule.

### 12. A missing context throws, and nothing defaults to black

`schemeClr` with no theme, `phClr` outside a style invocation, `sysClr` with neither a `lastClr` nor
a table entry, a `prstClr` name outside the 147, a theme slot that resolves to itself: five codes,
five different fixes. A themed deck rendering entirely black is the hardest colour bug to see in a
screenshot and the easiest to see in an exception.

---

## What building it found

### The plan's four colour claims: three hold, one does not

- "transforms applied in document order (`lumMod` then `lumOff` ≠ the reverse)" — **holds**, 0.7-C.
- "`dk1`/`lt1`/`dk2`/`lt2` bypass `clrMap`" — **holds**, and now against a non-identity map.
- "`a:prstClr` names are not CSS colour names" — **holds**, all 147 of them.
- "`a:sysClr` prefers `@lastClr`" — **does not hold** of PowerPoint. See decision 9.

### PowerPoint enforces `ST_PositiveFixedAngle` on `a:hslClr/@hue`

`hue="24000000"` and `hue="-3600000"` were refused outright and opened only after a repair, which
reset the hue to zero. `parseAngle` does **not** enforce it: we are a reader, and refusing to
display a file PowerPoint would repair is a worse outcome than displaying it wrapped.

### Two things about compositing, neither of them ours

Both fell out of checking the alpha swatches against the bitmap, and both are recorded because 2.10
will want them.

PowerPoint composites in **sRGB**, `a*c + (1-a)*bg`, even though it blends `tint` and `shade` in
linear light. And it composites the **quantised** colour: the `lumMod`/`lumOff` swatch resolves to a
green a hair above `AA`, and compositing the real number gives `D5` where PowerPoint gives `D4` —
which is what compositing `AA` gives. Its rasteriser also rounds ties down on all three channels of
the 50% swatch, where the colour engine's ties go both ways.

### Mutation testing

Twelve mutations, each applied by literal replacement to a `cmp`-verified baseline, type-checked and
run. Counts are tests failed out of 31.

| mutation                                                | killed |
| ------------------------------------------------------- | -----: |
| `toByte` rounds halves up                               |      5 |
| saturation clamped inside a transform                   |      5 |
| the clamped state carried across a transform boundary   |      3 |
| `clrMap` consulted for `dk1`/`lt1`/`dk2`/`lt2`          |      3 |
| `parsePercentage` implemented with `parseInt`           |      3 |
| `a:scrgbClr` read as sRGB percentages                   |      2 |
| `tint` and `shade` blended in sRGB                      |      2 |
| `a:gray` computed in linear light                       |      2 |
| `a:inv` as `255 - c` in sRGB                            |      2 |
| `a:comp` as a linear-light complement                   |      2 |
| transforms sorted rather than applied in document order |      2 |
| `a:sysClr` consults the table before `@lastClr`         |      1 |

The last row is the divergence in decision 9, and one test failing is the right answer: exactly one
assertion exists to pin it.

---

## Consequences

- 2.9's resolver supplies `ClrScheme`, `ClrMap` and `phClr`; this package does not know how to find
  them and will throw rather than guess.
- 2.7's gradient stops and 2.8's line and effect colours are `Color` values and go through
  `resolveColor` unchanged. The `phClr` substitution the plan flags for 2.9 is already a context
  field here.
- 7.6's palette swap is a `ClrScheme` replacement and nothing else, because nothing downstream has
  flattened a scheme name.
- 7.10's theme-ification must match against **this** transform math, which the plan already
  requires. `applyTransforms` is exported for that.
- The 147-name preset table and the 30-name system table are committed data with a test that
  re-derives them from the fixture, so the two cannot drift.

---

## What is not done

- **Nothing is painted.** This produces `Rgba`. No fill, no stroke, no compositing.
- **Gradients, patterns, dashes, arrowheads, effects** — 2.7 and 2.8.
- **Chart and diagram colour styles** (`colors1.xml`, `cs:variation`) — Phase 9.1, and a different
  question.
- **`a:duotone`, `a:alphaModFix` and the blip-level colour changes** (`a:clrChange`) are image
  effects, not `EG_ColorChoice`, and belong with 10.7.

### Open questions, in the 0.7 style

1. **What decides PowerPoint's rounding at an exact half?** Seven samples, no rule. A sweep designed
   to hit ties on purpose — a few hundred of them — would either find the rule or establish that
   there is not one to find. Costs an hour and would remove the last residual.
2. **Is `a:sysClr` resolved per-viewer in PowerPoint for the web?** If it is, the divergence in
   decision 9 is a divergence from the desktop only, and the web behaviour is the one to match.
3. **Does a `clrMapOvr` on a layout or slide reach a colour inside a theme style?** 2.9's question,
   but it is a colour question and it is cheap to probe here.
4. **Is `a:scrgbClr` clamped before or after its transforms?** We clamp at the base. A chain of
   `scrgbClr r="150000"` plus `shade` would tell the two apart.
5. **What does PowerPoint do with a `schemeClr` naming a slot the theme does not define?** Our
   answer is `COLOR_SCHEME_NAME`; a theme with eleven `clrScheme` children would settle whether
   PowerPoint repairs it or substitutes.

---

## Verification

- `pnpm check` green.
- 259 new swatches committed as `corpus/ground-truth/color-bases.json`, with a manifest entry
  establishing CC0 provenance; 20 of 21 probe packages accepted by PowerPoint with no repair, and
  the one that was not is flagged in the data and excluded in the test.
- Object model and painted pixels agree on all 464 opaque swatches across both experiments.
- Of the 455 swatches this package predicts, **454 are exact** and one is off by a single unit in a
  single channel, at a rounding tie the fixture names.
- 12 mutations, each killed by at least one test.
