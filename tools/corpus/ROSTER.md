# The corpus roster

What sub-phase 1.1's corpus holds, now fifty-four decks, deck by deck, and which experiment gates the ones that
cannot be written yet. ADR 0009 decides _where a fixture may come from_; this decides _what is in
the corpus_.

It is a plan, not a manifest. `corpus/decks/manifest.json` is what exists; a slot named here and
missing there is a deck still to write. `C-COV` is the rule that refuses to let the difference go
unnoticed — every census key covered by at least one deck, or named in a manifest's `uncovered`
array with why nothing covers it and what would close it. A declared gap is a signed statement; an
undeclared one is how a 54/54 badge becomes a lie. It is `C019` in `check.ts`, and it is now
closed: forty-six of forty-seven keys covered, one declared.

## Status

| Tier  | Producer                              | Planned | Built  |
| ----- | ------------------------------------- | ------- | ------ |
| **A** | `tools/corpus/tiers/a-generated`      | 44      | **44** |
| **B** | Microsoft PowerPoint 365 (16.0.20326) | 9       | **9**  |
| **C** | `packages/opc`'s own writer           | 1       | **1**  |

**All three tiers are complete**, and so are all four named rules. `C-LEX` was waiting on the
second and third producers (`tools/corpus/lexical/lexical.test.ts`), so the round-trip gate is no longer a
proof of idempotence; `C-COV` was waiting on `a28-model3d`, which is **cut**. E4 is blocked on a
download nobody has approved, and the roster's own instruction for that case was to declare the key
rather than drop it quietly, so `model3d` is now the single entry in `corpus/decks/manifest.json`'s
`uncovered` array.

The three do not contribute equally, and the rule now counts the difference rather than describing
it. PowerPoint is a second producer of everything — part names, element order, lexical form and ZIP
headers alike. Our own writer is a second producer of the **container only**: `c01-opc-writer` is
`b01-blank` rewritten, so its XML is Microsoft's byte for byte and only its thirty-seven ZIP
headers are ours. `C018` makes each manifest declare both halves, and `corpus/written` names
PowerPoint as its XML serializer using the same string `corpus/authored` uses, so the two collide
to one producer instead of counting twice.

The result is a corpus that is well evidenced at one layer and thin at the other:

| layer                    | serializers | forms | covered by two |
| ------------------------ | ----------- | ----- | -------------- |
| the ZIP container        | **3**       | 24    | **17**         |
| the XML inside the parts | **2**       | 36    | **12**         |

Forty-four, and forty-three of them keep the tier's rule that a probe is about one thing.
`a43-kitchen-sink` is the exception, added at Gate 1 and added deliberately: the gate does not ask
whether five features survive a round trip — five decks here answer that — but whether they survive
it **together**, which has failure modes none of the five can reach. Four `Default` content types
from three feature families sharing one `[Content_Types].xml`; a macro-enabled main part beside four
chart parts; and a `p:timing` tree whose `p:spTgt/@spid` names a `p:graphicFrame` rather than a
`p:sp`, which is where renumbering shape ids on export silently unhooks an animation. It is built by
merging `a21`, `a23`, `a26` and `a32` rather than by re-authoring them, so it cannot drift from the
decks it is made of — and it is the second deck in the corpus with a `vbaProject.bin`, which took
`macros` off `C-COV`'s single-probe list.

Forty-one was forty-two until `a28` was cut, and forty until `a18-slide-sizes` was built and the
obvious became unavoidable: `p:sldSz` is one element on `p:presentation`, so a package has exactly
one slide size and the roster's "4:3, 16:9, A4 and a custom size" is four decks, not one. `a18` is
the 4:3 one, `a41-a4` and `a42-custom-size` are the other two, and 16:9 is every other deck in the
tier — which turned out to be four decks about `p:sldSz` rather than the three that split implies,
because measuring the sixteen `ST_SlideSizeType` values found `screen16x9` naming two different
extents.

`a28-model3d` is the one slot that went the other way, and it went there by the route the roster
wrote down in advance: a slot that turns out to be genuinely uncoverable goes into the manifest's
declared `uncovered` array rather than being quietly dropped. Whoever runs E4 deletes that entry in
the same commit that adds the deck — `C019` fires on a declaration that has been outlived, so the
file cannot go on apologising for a hole that was filled.

## Tier A — the generated probes

Every deck is one to three slides, committed as bytes, and reproduced by
`node tools/corpus/tiers/a-generated/build-probes.ts --out corpus/decks --id <id>`.

`a01-minimal` is the **subtrahend**: its census is the chassis alone — a master with two
placeholders, one layout with a third, and the two `a:gradFill` entries in the theme's
`fillStyleLst` that `fillRef/@idx` needs. Any other deck's `features` map minus that one is what
that deck's probe contributes.

### Fills, effects, geometry — built

| id               | probes                                                                               | first to emit                                              |
| ---------------- | ------------------------------------------------------------------------------------ | ---------------------------------------------------------- |
| `a01-minimal`    | the floor: one slide, one title, nothing optional                                    | —                                                          |
| `a03-fills`      | six colour models, 7 gradients, 8 picture fills, 54 pattern presets                  | `blipFill`, `groupFill`, `group`                           |
| `a04-effects`    | all 8 `a:effectLst` children, `a:scene3d`/`a:sp3d`, `a:effectDag`                    | `innerShadow`, `glow`, `softEdge`, `reflection`, `scene3d` |
| `a05-geometry`   | six path commands, `gdLst`/`ahLst`/`cxnLst`/`rect`, 24 adjusted presets              | `customGeom`                                               |
| `a06-lines`      | widths, caps, `@cmpd`, 11 dashes, `custDash`, joins, arrowheads, `p:cxnSp`           | `connector`                                                |
| `a44-transforms` | 12 rotations and flips on one asymmetric preset, and 6 groups including a nested one | first `@rot`, `@flipH`, nested `p:grpSp`                   |

### Text and inheritance — built

| id                 | probes                                                                          | first to emit |
| ------------------ | ------------------------------------------------------------------------------- | ------------- |
| `a02-placeholders` | 14 of the 16 `ST_PlaceholderType` values, `p:hf`, the 5-tier matcher's cases    | —             |
| `a07-text-cascade` | the 10-source cascade, `lstStyle` at every level, `ST_Percentage` in both forms | —             |
| `a08-bullets`      | `buChar`/`buAutoNum`/`buBlip`, the 41 autonumber schemes, PUA symbol bullets    | —             |
| `a09-fields`       | the 15 reserved `a:fld` types with their `ST_Guid` and cached text              | `field`       |
| `a10-rtl-cjk`      | `a:rtl`, `a:ea`/`a:cs`/`a:sym` runs, `p:kinsoku`, `vert`/`vert270`/`eaVert`     | —             |
| `a11-autofit`      | `normAutofit` at ladder steps, `spAutoFit`, `spcFirstLastPara`, `a:endParaRPr`  | —             |

`a10` is where `@lang` returns to the generator, threaded properly. It was added and removed in
one sitting during the recipe refactor because it was declared and honoured nowhere.

Two corrections came out of building these, and both live in the decks rather than only here.
`a:fld` has **fifteen** reserved types, not the fourteen the plan says: `slidenum`, `datetime`, and
`datetime1` through `datetime13`. And `a02` covers fourteen of the sixteen placeholder types
because PowerPoint refuses the other two outright — see **What PowerPoint refuses**.

`a11` carries a spread of `normAutofit` scales from 100% down to 25% but is deliberately **not**
the source of truth for the ladder's steps. Sub-phase 3.4 measures those from PowerPoint's own
output, and a fixture that guessed them would turn that measurement into a tautology. What `a11`
is authoritative about is that stored values are applied verbatim in view mode.

### Structure — built

| id                | probes                                                                                   | first to emit                    |
| ----------------- | ---------------------------------------------------------------------------------------- | -------------------------------- |
| `a12-masters`     | three masters, three themes, three `p:clrMap`s, `a:overrideClrMapping`                   | —                                |
| `a13-sections`    | `p14:sectionLst`, `p:custShowLst`, a slide in two shows and one listed twice             | `section`, `customShow`          |
| `a14-notes`       | notes master, handout master, three notes slides, and the `hdr`/`sldImg` pair            | `notesSlide`                     |
| `a15-comments`    | `p:cmLst` and `p:cmAuthorLst`, **and** the 2018 `p188:cmLst` PowerPoint writes           | `comment`                        |
| `a16-transitions` | all 21 `p:transition` effects across a master, 18 layouts and 3 slides                   | `transition`, `alternateContent` |
| `a17-animations`  | `p:timing` nine levels deep, every behaviour but audio and video, `p:bldLst`             | `animation`                      |
| `a18-slide-sizes` | `screen4x3`, and a `p:notesSz` in a different aspect ratio from the slide                | —                                |
| `a19-decorative`  | `adec:decorative`, `@descr`, `@title`, `@hidden`, and `p:spTree` reading order           | `decorative`                     |
| `a45-backgrounds` | `p:bgRef` across the 1000 offset, an explicit `p:bgPr`, and a layout-supplied background | first `p:bgPr`                   |

Four corrections came out of building these, and all four are measurements rather than opinions.

**`p:transition` has twenty-one effects, not twenty-two.** `CT_SlideTransition`'s choice group holds
blinds, checker, circle, comb, cover, cut, diamond, dissolve, fade, newsflash, plus, pull, push,
random, randomBar, split, strips, wedge, wheel, wipe and zoom. Its other two children — `p:sndAc`
and `p:extLst` — are not effects. `a16` covers all twenty-one, which needs more sheets than a
three-slide deck has: a transition is legal on `p:sldMaster` and `p:sldLayout` too, and a slide
with none of its own inherits its layout's.

**PowerPoint no longer writes `p:cmLst`.** A comment added through `Slide.Comments.Add` on
2026-08-27 produced `ppt/comments/modernComment_100_4A8498BC.xml` with content type
`application/vnd.ms-powerpoint.comments+xml`, plus `ppt/authors.xml`, both rooted in `p188:` —
`http://schemas.microsoft.com/office/powerpoint/2018/8/main`. **The census's `comment` rule matches
the content type ending `presentationml.comments+xml` and therefore reports zero comments for a
deck commented in PowerPoint today.** `a15` carries both formats so the gap is stated rather than
unnoticed; whether the census grows a rule for the 2018 form is an open decision, not a silent one.

**A notes slide binds to its notes master on type, not on `(type, idx)`.** Measured from the same
deck: the notes master's body placeholder is `idx="3"` and the notes slide's is `idx="1"`, while
the slide's `sldImg` carries no `@idx` at all against a master that says `idx="2"`. This is the
same shape as sub-phase 7.1's layout-to-master rule, and a matcher that requires `idx` to agree
orphans the notes text of every deck PowerPoint has ever written.

**PowerPoint writes `&quot;`.** The open question from the text group is closed. A shape renamed
through COM to `a "quoted" name` was saved as `name="a &quot;quoted&quot; name"` — two occurrences
and no other entity anywhere in the package. `escapeAttribute` escapes rather than throwing,
`powerpoint-conventions.json` records the fourth entity beside E8's three, and `a19` exercises it
in `@descr`, the attribute that holds free text a user typed.

Two smaller facts from the same measurement, both in `a14` and `a19`: **there is no handout master
until somebody edits one** — a deck with speaker notes gets `notesMaster1.xml` and `theme2.xml` and
no handout master at all — and **decorative and alt text are mutually exclusive**, because marking a
shape decorative clears its `@descr` and `@title`.

### Hard content — built

| id                   | probes                                                                                      | first to emit                          |
| -------------------- | ------------------------------------------------------------------------------------------- | -------------------------------------- |
| `a20-tables`         | `gridSpan`/`hMerge`/`rowSpan`/`vMerge`, every `a:tcPr` child, both table-style sources      | `table`, `graphicFrame`, `patternFill` |
| `a21-charts`         | bar, combo on a secondary axis, pie from literals, scatter with a fit, sparse `c:pt/@idx`   | `chart`                                |
| `a22-chartex`        | seven of the eight ChartEx `layoutId` values, and the switch PowerPoint wraps them in       | `chartEx`                              |
| `a23-smartart`       | `dgm:relIds`, four parts, and two different kinds of missing `dsp:` drawing                 | `smartArt`, `smartArtDrawing`          |
| `a24-media`          | embedded audio and video, a linked video, the `p14:media` double relationship               | `audio`, `video`, `media`              |
| `a25-svg-blips`      | `asvg:svgBlip` with and without a raster, in a picture and in a fill, and a hostile SVG     | `svgBlip`                              |
| `a26-ole`            | the E6 shape: `mc:Choice Requires="v"` with no VML part, plus a transitional frame with one | `oleObject`, `embeddedPackage`         |
| `a27-ink`            | `inkml:trace` inside `p:contentPart`, with `p14:xfrm` and bare                              | `ink`, `contentPart`                   |
| ~~`a28-model3d`~~    | ~~`am3d:model3d` with its baked `am3d:raster` preview~~                                     | **cut — declared `uncovered`**         |
| `a29-math`           | `m:oMath` inside a text body, at both MCE positions                                         | `math`                                 |
| `a30-vml`            | a `vmlDrawing` part reached by `@spid`, and what VML can express                            | `vml`                                  |
| `a31-embedded-fonts` | `p:embeddedFontLst`, the six export artifacts, the CC0 probe font                           | `embeddedFont`                         |

`a26-ole` emits two frames, and its `oleObject` count is **2 per logical object**: the census reads
raw pre-MCE markup, so the `mc:Choice` and the `mc:Fallback` both count. That is the right policy
for a census and `features` must say so.

`a28-model3d` is **cut**, and `model3d` is the corpus's one declared gap. Experiment E4 needs the
`a:ext` GUID that carries `am3d:model3d`, and it is in no public specification: Microsoft's own
Open XML SDK documents the element and its namespace — `DocumentFormat.OpenXml.Office2019.Drawing`
`.Model3D`, serialized `am3d:model3d`, which at least corroborates the census rule's spelling — but
not the extension URI that hosts it, nor the relationship type and content type of the model part
beside it. The only way to obtain the set is to insert a 3-D model in PowerPoint and read it back,
and PowerPoint's gallery fetches a Microsoft-licensed model over the network. That is a download and
has not been approved.

Guessing the four strings was considered and rejected, and `a34-extlst` is the deck that settles
it. What that deck measured is that **an unknown `a:ext/@uri` is carried through untouched** — which
is the contract `extLst` exists for, and here it cuts the wrong way. A fabricated GUID would not be
refused, would not be stripped, and would not degrade on the fifth resave: PowerPoint would open the
deck, render the baked raster, and preserve the extension indefinitely as somebody else's. Every
gate this corpus has would pass, permanently, on a probe that probes nothing. A silent failure that
never surfaces is worse than a hole that is written down.

Two things paid for this group. Four decks — charts, ChartEx, SmartArt and media — were written
from decks PowerPoint 16.0.20326 authored on 2026-08-27 and read back, so their relationship types,
content types and extension URIs are measured rather than transcribed. And **five of the eleven
were refused by PowerPoint on the first attempt**, which is five findings the corpus would not
otherwise have. All of them are below.

#### Six refusals, all found by bisection

- **`c:tx` takes `c:strRef` or `c:v`, and a `c:strLit` there is a whole-package refusal.**
  `c:cat` accepts five branches including `c:strLit`, and `c:val` accepts `c:numLit`, so the
  literal forms look general. `CT_SerTx` is not one of them.
- **A `cs:chartStyle` with a subset of its thirty-one entries is a whole-package refusal.**
  `a21` shipped four of them — chartArea, dataPoint, legend, plotArea, which is everything the
  corpus probes — and PowerPoint refused the file. With all thirty-one present it opens. Nothing
  says the sequence is required rather than optional, and the message names no part.
- **A `cx:chartSpace` part with no relationships of its own is a whole-package refusal.** Proved
  the strong way: `chartEx1.xml` was lifted **byte for byte** out of the deck PowerPoint had just
  written, put in a package with nothing else changed, and refused; adding back its `chartStyle`
  and `chartColorStyle` relationships made the same package open. A classic `c:chartSpace` has no
  such requirement — `a21`'s charts 2, 3 and 4 carry none and open.
- **One WAV cannot be both a media object's target and a transition's `p:snd`.** Slides 1 and 2 of
  `a24` opened together, slides 2 and 3 opened together, and slides 1 and 3 — the two that shared
  the part — did not. The message differs from every other refusal here: "PowerPoint could not open
  the file", rather than "the file or directory is corrupted and unreadable". PowerPoint's own
  output keeps them apart too, writing `media1.wav` for an object and `audio1.wav` for a sound.
- **A `p:contentPart` whose relationship type is `…/officeDocument/2006/relationships/customXml`
  is a whole-package refusal**, and PowerPoint is not being strict in general: with only the
  `@Type` changed, `…/office/2010/relationships/customXml` opens, and so do `…/image` and
  `…/slide`, which are nonsense here. It validates that one type, because the ECMA `customXml`
  relationship means the document's custom XML data store and comes with a shape it knows.
- **`p:control` is a whole-package refusal in every form tried.** Eight variants, one change each:
  an empty `<p:controls/>` opens; `<p:control/>` with no attributes at all does not, and neither
  does name-only, `@spid` without `r:id`, minimal `@r:id` + `@spid`, the full form with
  `@imgW`/`@imgH` and a `p:pic` preview, the same with an ActiveX part using `persistStreamInit`
  and its `.bin`, or the same package retyped as a macro-enabled presentation. So it is the
  element, not its attributes, not its target, and not the absence of a VBA project. `a30-vml`
  reaches its VML through `p:oleObj/@spid` instead, and says why.

#### Five things PowerPoint does silently, found by re-saving

A deck that opens is only half the gate. Each of these eleven was opened, saved through PowerPoint
into the session scratchpad, and diffed against the original — which is the only way to see a
repair that does not announce itself.

- **`p:contentPart` is discarded on save**, and so is the `mc:Fallback` picture beside it:
  `a27`'s three content parts come back as zero, with no `mc:AlternateContent` and no `p:pic`
  left. The deck opens, PowerPoint does not surface the parts as shapes, and saving loses them.
  Whether that is because our ink markup is wrong or because PowerPoint reconstructs ink only from
  its own files cannot be told apart from here, and the declared gap below says so.
- **VML that is not an OLE site is converted to DrawingML.** `a30`'s `v:group` — reached only
  through the slide's `vmlDrawing` relationship and named by nothing in the slide — arrives as a
  native `p:grpSp`, which is why PowerPoint reports four shapes on that slide where the generator
  wrote three. Content added, not lost. On save the VML parts are gone.
- **A transitional `p:oleObj` is upgraded to the `mc:AlternateContent` form**, so `a26`'s four
  become six, and the embedding is renamed to `Microsoft_Excel_Worksheet.xlsx`.
- **A missing SmartArt drawing is generated.** `a23` slide 3 has no `dsp:dataModelExt` and
  therefore no drawing part; PowerPoint lays the diagram out and writes one, so two `dsp:drawing`
  parts come back as three. That is the deck's own claim — that only the second kind of missing
  fallback is repairable — confirmed by measurement.
- **Parts PowerPoint did not write are dropped.** `a21`'s `colors1.xml`, `style1.xml` and the
  chart's `.rels` are gone after a round trip; `a25`'s PNG fallback and second SVG are gone and
  the survivor is renamed `image1.svg`.

None of that is a defect in these decks — every one opens with no repair prompt and every probed
element is present in the file we wrote. It is a list of the things **our** writer must not do, and
it is why sub-phase 1.4 compares canonical XML and a relationship graph rather than bytes.

#### And ten measurements that are not refusals

- **ChartEx arrives inside `mc:AlternateContent`.** The `mc:Choice` requires `cx1` —
  `http://schemas.microsoft.com/office/drawing/2015/9/8/chartex`, which is **not** the `cx`
  namespace the chart is in and exists only to be named in `@Requires` — and the `mc:Fallback`
  holds a `p:pic` of a PNG PowerPoint rasterised at save time.
- **`chartEx1.xml` writes its self-closing tags spaced.** Twelve of twelve, `<cx:title … />`,
  while the other 3,046 self-closing tags across the same sixty parts have no space at all. There
  is more than one XML serializer inside PowerPoint. See `a36-spaced-tags`.
- **An SVG inserted through COM produces an `a:blip` with no `@r:embed` at all**, whose only
  reference is the `asvg:svgBlip` in its extension list, and no PNG part is written beside it.
- **`dsp:dataModelExt/@relId` resolves against the slide's rels**, not the data part's, and its
  `a:ext/@uri` is a plain URI rather than a GUID. Every `dsp:cNvPr` carries `id="0" name=""`;
  identity is `dsp:sp/@modelId`.
- **Media is related to twice.** The 2006 `audio`/`video` relationship named by
  `a:audioFile`/`a:videoFile`, and the 2007 `media` relationship named by `p14:media`, both
  pointing at the same part. `@r:link` is used even when the target is internal, so only
  `TargetMode` distinguishes an embedded clip from a linked one.
- **Every media shape carries `<a:hlinkClick r:id="" action="ppaction://media"/>`** — a required
  `r:id` holding the empty string. Sub-phase 1.2's rule that every `r:id` resolves must exempt it.
- **`<Default Extension="wav" ContentType="audio/x-wav"/>`**, and `p:stSnd/p:snd/@name` is the
  original file name rather than the part name.
- **PowerPoint refuses to insert a video it cannot decode** — an ISO-BMFF container with no tracks
  is rejected at insert time with a codec message — but opens a package that references one.
- **The extension URIs**, all measured: `a16:creationId`
  `{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}`, `p14:modId`
  `{D42A27DB-BD31-4B8C-83A1-F6EECF244321}`, `p14:media`
  `{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}`, `asvg:svgBlip`
  `{96DAC541-7B7A-43D3-8B79-37D633B846F1}`.
- **A ChartEx's dimension type is not fixed.** `numDim type="val"` for waterfall, funnel and box
  and whisker; `type="size"` for treemap and sunburst; and histogram and Pareto carry no category
  dimension at all. A hierarchy is several `cx:lvl` children of one `cx:strDim`, deepest first.

### The awkward ones — built

| id                 | probes                                                                                               |
| ------------------ | ---------------------------------------------------------------------------------------------------- |
| `a32-macros`       | a `.pptm` with a real VBA project whose only body is `Sub Noop()`                                    |
| `a33-thumbnail`    | `docProps/thumbnail.jpeg`, which PowerPoint writes and no _part_ references                          |
| `a34-extlst`       | `a:extLst`/`p:extLst` with unknown URIs, kept as an ordered opaque list                              |
| `a35-zip-shapes`   | mixed compression methods, the general-purpose flag bits, and the 520-byte `0xa220` growth hint      |
| `a36-spaced-tags`  | self-closing tags written `<x />`, and six more forms no producer here emits                         |
| `a37-mce`          | `mc:AlternateContent` nested, `mc:Ignorable`, `mc:ProcessContent`, `mc:MustUnderstand`               |
| `a38-degenerate`   | zero extents, `chExt` of 0, an empty group, a zero path space, an empty `p:spTree`                   |
| `a39-large-ids`    | the four id spaces at their edges, and the one PowerPoint does not implement                         |
| `a40-unicode`      | part names and text at the edges of what OPC and XML 1.0 permit                                      |
| `a41-a4`           | `p:sldSz type="A4"` — a second package, because one package holds one slide size                     |
| `a42-custom-size`  | a `p:sldSz` with no `@type`, at an extent no enumeration names                                       |
| `a43-kitchen-sink` | charts, SmartArt, animations, OLE and macros **together** — the one deck that is not about one thing |

`a35-zip-shapes` is the **only** corpus deck that deflates. Everything else is stored, so its bytes
are a pure function of our own XML and `C-REGEN` does not depend on which zlib built it. Making the
exception the deck that is _about_ compression keeps the dependency where it can be reasoned about.
It is therefore also the one deck whose committed hash can change on a toolchain upgrade with no
source change; the fix when that happens is to check that nothing but the deflate streams moved and
re-pin it.

`a32-macros` is the only deck that is not a `.pptx`. PowerPoint checks the file extension against
the content type of the main part and refuses the pair when they disagree, so a macro-enabled
package has to be written `.pptm` — which is why `ProbeDeck` grew an `extension` field, and why the
manifest's `format` and `path` both derive from it rather than being two independent claims.

Three of these eleven were refused on the first attempt, and bisecting them produced the three
findings below. Two more decks needed a chassis option each; the rest are markup.

#### Three more refusals, all found by bisection

- **A slide has exactly one `slideLayout` relationship, and it resolves to a slide layout part.**
  Four variants, one change each, all rejected with "PowerPoint could not open the file": no `.rels`
  part at all, a `.rels` part holding one relationship of a different type, a `slideLayout`
  relationship whose target is a **master**, and one whose target does not exist. **Two** of them is
  a refusal as well. `CT_Slide` requires none of this and neither does OPC — the binding lives only
  in the rels part, so the schema has nothing to say about it. `a38-degenerate` was written with the
  unbound slide and now records why it has not got one.
- **`p:cNvPr/@id` is not the space `ST_DrawingElementId` says it is.** Swept one id per package:
  `0 … 2147483647` open, **`2147483648 … 4294967294` are whole-package refusals**, and `4294967295`
  opens. PowerPoint reads a schema-declared `xsd:unsignedInt` as a **signed 32-bit integer** and
  rejects everything that comes out negative, except `0xFFFFFFFF` — which is minus one, and which it
  keeps as a sentinel: re-saving renumbers it away while `2147483647` survives untouched. Two
  billion schema-legal values are refused and the two survivors sit at opposite ends of the declared
  range.
- **A percent-escape of an _unreserved_ character in a part name is a refusal**, with `0x808D1005` —
  a code this project had not seen, from the packaging layer rather than the presentation one. Nine
  one-name packages, nine for nine: `%2D` `%41` `%5F` `%7E` refused; `%24` `%2C` `%3A` `%20` `%23`
  open. The line is exactly RFC 3986 §6.2.2.2 — an escape that is _required_ is fine, an escape that
  is _gratuitous_ is fatal. **`packages/opc` rates that case a warning**, rule `M1.8`, so
  `toPartName` accepts what PowerPoint rejects. That is a severity to fix in sub-phase 1.2 rather
  than a rule to add: nothing in the code was wrong and the number beside it was.

#### Nine things the re-save pass found this time

Each of the eleven was opened, saved through PowerPoint into the scratchpad, and diffed. Every deck
gains a `p:extLst` on two slides and three `p:ext` children — `p14:creationId` and friends — which
is the baseline; everything below is on top of it.

- **An `a:extLst` survives whole.** All thirteen `a:ext` children of `a34`'s hazardous list come
  back **in the original order**, including the URI used **twice** and the two GUIDs differing only
  in **case**, and including an `a:ext` with no payload at all. That is the contract stated in the
  must-not-break rules, confirmed by measurement rather than by argument.
- **But only the _last_ child element of an `a:ext` survives.** An extension holding `a`, `b`, `c`
  came back holding `c`; one holding four elements came back holding the fourth. PowerPoint treats
  `CT_OfficeArtExtension` as carrying one element however many the schema's `xsd:any` permits — so
  a writer that adds a second child to an existing extension loses the first.
- **A comment inside an extension is dropped; a CDATA section is not.** Both are things a typed
  model would lose and only one of them is something PowerPoint loses.
- **`mc:Ignorable` and everything it names are dropped**, and the switches PowerPoint understands
  are resolved: `a37`'s nine `mc:AlternateContent` elements come back as six, `mc:ProcessContent`
  as one of two, and the `zz:`-namespace markup all but disappears. Its shape counts on open were
  5, 5, 5 — which is how `mc:ProcessContent` was confirmed to be **implemented**: the shape inside
  the named wrapper survives and the one inside the unnamed wrapper does not.
- **Every spaced self-closing tag is re-serialized tight.** `a36`'s thirty-five come back as zero.
  PowerPoint normalises the lexical form of any part it rewrites, which is exactly why our own
  serializer must not: a deck that has been through PowerPoint is no longer a fixture for `C-LEX`.
- **A carriage return inside an `a:t` is structure, not text.** One run written
  `before&#13;after` came back as **two `a:p` elements**. A numeric character reference to `#xD`
  survives XML 1.0's line-ending normalisation and then does not survive PowerPoint's text model.
- **Astral characters, combining marks and bidi controls all survive** as literal UTF-8: U+1D54F,
  U+1F4C4, U+20BB7, a combining acute, U+202E/U+202C, U+FFFD, ZWJ, ZWNJ and NBSP came back
  unchanged.
- **`p:sldId` values and their order survive**, including 2147483647 first and 256 second, as do
  `p:sldMasterId` and `p:sldLayoutId` at the very top of their range.
- **Media part names are renumbered** whatever they were: `image%201.png`, `image+2,~!()$@'.png`,
  `IMAGE-3.PNG` and `image%234.png` all came back as `image1..4`, and the mixed-case one kept its
  extension's case as `image3.PNG`.

#### And six measurements that are not refusals

- **`ST_SlideSizeType` is a hint that cannot be inverted.** Read out of `PageSetup` one value at a
  time: `onScreen`, `letter` and `overhead` all write **9144000 x 6858000**, and `screen16x9` names
  **two** extents — 9144000 x 5143500 from Page Setup, and 12192000 x 6858000 for a new deck, which
  is what every other deck in this corpus is. `A4` is 9906000 x 6858000, the printable area rather
  than the 297 x 210 mm sheet.
- **PowerPoint never writes `type="custom"`.** Setting a custom size makes `PageSetup.SlideSize`
  report `ppSlideSizeCustom`, and the saved file omits `@type` entirely.
- **Macro-enabled is one content type and nothing else.** A `.pptm` saved with no code is identical
  to the `.pptx` but for the `Override` on `/ppt/presentation.xml`, and carries **no
  `vbaProject.bin` at all**. The content type and the part are independent signals.
- **`docProps/thumbnail.jpeg` is 256 x 144, baseline, JFIF 1.1 at 96 dpi, 4:2:0**, with the four
  **standard ITU-T T.81 Annex K Huffman tables** and two quantization tables of PowerPoint's own.
  `tools/corpus/tiers/a-generated/assets/jpeg.ts`'s tables were read out of its `DHT` segments rather than recalled.
- **`_rels/.rels` is written out of rId order** — `rId3`, `rId2`, `rId1`, `rId4`.
  `CT_Relationships` is a sequence, so a writer that regenerates the part sorted by id changes bytes
  it was not asked to change. `a33` records it and does not reproduce it; it closes with a Tier B
  deck.
- **The thumbnail is reached only from `_rels/.rels`.** Nothing under `ppt/` mentions it, so it is
  the one part where "reachable" and "referenced by the document" come apart — and the case that
  decides whether sub-phase 1.3's media collector walks package relationships or only part ones.

`a36-spaced-tags` exists because of a measurement, and the measurement has since been corrected.
ADR 0005 counted **70,822 of 98,777** self-closing tags in the real-world corpus written spaced,
`<a:off … />`, against **zero** of 1,305 from PowerPoint 16.0.20326, and concluded that no producer
available to this project emits the spaced form.

That conclusion was wrong, and building `a22-chartex` is what found it. In the deck PowerPoint
wrote on 2026-08-27, `ppt/charts/chartEx1.xml` writes **12 of its 12** self-closing tags spaced
while the other 3,046 across the same sixty parts write none. PowerPoint has more than one XML
serializer and the ChartEx one is the odd member, so the spaced form _is_ reachable here — from a
producer, in a part, on demand.

The deck stays, because one part of one file type is a thin sample and `a36` is where the form is
exercised deliberately. What changes is the claim: `C-LEX` now has a second producer for this
lexical form rather than none, and the rule stands that it checks against forms known to exist in
the wild rather than merely against "two distinct producers", because two producers that agree
prove nothing.

Building it widened the deck from one form to seven. Whitespace before `/>` spelled with a space, a
tab and two spaces; whitespace before `>` on a start tag and on an end tag; whitespace either side
of `=`; and a single-quoted attribute value. Every one is a plain XML 1.0 production, none is
emitted anywhere in this repository, and the single-quoted attribute is the sharpest of them —
the only form here that a serializer cannot reproduce by remembering one boolean, because the
quote character has to be recorded per attribute rather than per part.

The forms are mixed **within** a part rather than uniform across one, which is what distinguishes a
per-node serializer from a per-part one and is therefore the fixture architectural bet 2 actually
needs. `chartEx1.xml` is uniformly spaced and cannot tell the two apart.

## Tier B — PowerPoint-authored — built

Nine decks, authored by `tools/corpus/tiers/b-authored/build-tier-b.ps1` so the provenance is auditable
line by line, and carrying **no recipe**: PowerPoint stamps `dcterms:created` into
`docProps/core.xml` on every save, so two runs one second apart differ, and the monthly build
changes the markup underneath. A recipe claiming to reproduce these bytes would be a lie with a
shelf life. They live in `corpus/authored/` rather than `corpus/decks/` because
`build-probes.ts --manifest` rewrites the latter in full, so an entry parked there would survive
exactly until the next Tier A change.

| id                | slides | probes                                                                                                |
| ----------------- | ------ | ----------------------------------------------------------------------------------------------------- |
| `b01-blank`       | 1      | the floor — `a01-minimal` as Microsoft writes it, and the file every lexical claim is checked against |
| `b02-layouts`     | 11     | the eleven built-in layouts, slide _i_ bound to layout _i_, for sub-phase 7.1's matrix                |
| `b03-text`        | 3      | `a:normAutofit`'s `@fontScale` as the ladder that defines it produced them                            |
| `b04-table`       | 1      | a built-in style GUID whose definition is nowhere in the package                                      |
| `b05-chart`       | 2      | one `c:plotArea` holding a `c:barChart` **and** a `c:lineChart`                                       |
| `b06-smartart`    | 3      | three families, five parts each, and the `dsp:` fallback                                              |
| `b07-ole`         | 1      | E6's own deck, committed — `mc:AlternateContent` with no `vmlDrawing` part                            |
| `b08-transitions` | 5      | `p14:honeycomb` in a `mc:Choice` whose `mc:Fallback` is a **different** transition                    |
| `b09-picture`     | 2      | a four-sided crop, four effects at once, and a filled `pic` placeholder                               |

`b01-blank` is Tier B's subtrahend the way `a01-minimal` is Tier A's, and it is much bigger:
**37 ZIP entries against our 18**, because all eleven layouts ship whether a slide uses them or
not and `docProps/thumbnail.jpeg` is written unasked. Every Tier B deck therefore reports at least
`{placeholder: 65, shape: 65, field: 24, presetGeom: 5, gradientFill: 3, shadow: 1, thumbnail: 1}`,
and the interesting part of any entry is what exceeds that.

### What the pair of floors is worth

`b01` and `a01` are the same deck by two producers, and diffing them is where the lexical
divergences show up smallest. Three, measured on 2026-08-28:

- **`p:sldSz` carries no `@type`.** PowerPoint writes `<p:sldSz cx="12192000" cy="6858000"/>` for
  a default widescreen deck; our chassis writes `type="screen16x9"` on the identical extent. Both
  are legal. The consequence is that a renderer keyed on `@type` rather than on `cx`/`cy` works on
  all 41 Tier A decks and fails on every file PowerPoint ever wrote. `PageSetup.SlideSize` reports
  `ppSlideSizeCustom` (7) for this deck, so the object model agrees the attribute is absent rather
  than defaulted — and `a42-custom-size` omits `@type` too, but at a portrait custom extent, so it
  never covered this case.
- **`_rels/.rels` is written `rId3`, `rId2`, `rId1`, `rId4`** — core properties, thumbnail,
  officeDocument, extended properties. This closes the declared gap that asked for it.
- **The thumbnail is the only STORED entry** in an archive that deflates everything else.
  `b05` and `b07` extend the rule: `ppt/embeddings/Microsoft_Excel_Worksheet.xlsx` is stored too.
  PowerPoint appears to store what is already compressed and deflate the rest.

`p:presentation` also carries `saveSubsetFonts="1"`, and `removePersonalInfo="1"` — see below.

### The four constraints, enforced rather than trusted

Blank Presentation only and never a shipped design template; no `ppt/fonts/` part; no Designer
suggestion and no stock imagery; and no author name, company or machine name anywhere in the
package. `decks.test.ts` asserts the last two directly — a whole-archive sweep for the account,
the organisation and the machine, plus the positive form on `dc:creator` and `cp:lastModifiedBy`.

**The scrub does not work the way this file used to say it did.** `BuiltInDocumentProperties` is
unreachable from PowerShell on this machine: the collection is non-null and enumerates to 34
items, but `.GetType()`, `.Item('Author')`, `.Name` and `.Value` all throw
`NullReferenceException`, and `[System.__ComObject].InvokeMember` reports
`Method 'System.__ComObject.Item' not found` because the Office type library is not bound in this
host. A scrub written that way runs, reports nothing and changes nothing — the first build of
`b01` was saved carrying the author's name twice, and the check that caught it was itself passing
vacuously until a field-name bug in it was fixed.

`RemoveDocumentInformation(ppRDIRemovePersonalInformation)` works, in one call, before the save
and without editing the saved file. It **empties** `dc:creator` and `cp:lastModifiedBy` rather
than removing the elements, which is why the test asserts `<dc:creator></dc:creator>` and not
absence. Its one visible cost is recorded rather than hidden: it writes `removePersonalInfo="1"`
onto `p:presentation`.

### Three findings from authoring

- **`PpEntryEffect` is one enumeration serving two features, and neither half accepts all of it.**
  `ppEffectFlashOnceFast` (3841) is rejected for a slide transition with "this enumeration value is
  not valid for transitions". The transition-valid values used here are `ppEffectFade` 1793,
  `ppEffectDissolve` 1537, `ppEffectWipeRight` 2819, `ppEffectPushUp` 3855, `ppEffectHoneycomb` 3898.
- **`Presentations.Add` returns a presentation with zero slides, not one.** Deleting "the
  auto-created first slide" costs a real one; `b02` first came out with ten layouts instead of
  eleven that way.
- **A slide bound to a layout inherits only that layout's content placeholders.** Every layout
  carries `dt`, `ftr` and `sldNum`, and none of the three reaches a slide unless `p:hf` turns it
  on — which is why a slide on Title Slide reports two shapes and not five. `b02`'s shape counts
  read `[2, 2, 2, 3, 5, 1, 0, 3, 3, 2, 2]`, with Blank at zero.

## Tier C — our own writer — built

`c01-opc-writer`, one deck round-tripped through `packages/opc`'s writer: `b01-blank` read through
`PartStore.open` and written straight back out, with nothing touched in between. It gives
`cli roundtrip` a case against the writer this project ships, and it is the committed evidence for
architectural bet 1 at package scale — 37 entries in, 37 out, in the same order, every part
byte-identical after inflation, and no entry recompressed.

| id               | source                           | bytes           | entries |
| ---------------- | -------------------------------- | --------------- | ------- |
| `c01-opc-writer` | `corpus/authored/b01-blank.pptx` | 33 456 (−1 832) | 37      |

Opens in PowerPoint 16.0.20326 with no repair prompt, two shapes on its one slide, which is
`b01`'s Title Slide count unchanged.

**Tier C has a `C-REGEN` and Tier B cannot**, which looks backwards until you see why. A deflating
writer normally cannot promise reproducible bytes at all — deflated output depends on which zlib
produced it, and that is exactly why Tier A stores its entries rather than compressing them. It
holds here because a no-op write touches no part, so `passthroughEntry` moves every already-
compressed stream across without inflating it and `fflate` never runs. Writing `c01` again is
byte-identical; writing `c01` itself is a fixpoint. Both are asserted.

### What our writer normalises away

The whole contribution of this tier is the diff against its source, because its census is `b01`'s
key for key and adds nothing to `C-COV`. Three header fields differ, every one of them legal to
drop, none of them a bug — which is precisely why they are asserted in
`written/decks.test.ts` rather than described here and checked nowhere. A reader ignores all three,
so nothing else in this repository would ever notice if our writer started emitting them or if
PowerPoint stopped.

| field           | PowerPoint 16.0.20326              | `packages/opc` |
| --------------- | ---------------------------------- | -------------- |
| version made by | 45 (host 0, i.e. ZIP 4.5)          | 20             |
| general-purpose | `0x0006` deflated, `0x0000` stored | `0x0000`       |
| extra field     | `0xA220` growth hint × 5           | none           |

- **Version made by 45, not 20.** `zip-writer.ts` said 20 was "measured from a file PowerPoint 365
  wrote on this machine". It was not; nothing here needs above 2.0 and 20 is the honest value for
  what we emit, but the sentence claiming it came from a measurement was wrong and has been fixed.
- **`0x0006` is bits 1 and 2**, which for method 8 are a compression _level hint_ and mean nothing
  to a decompressor — `0b11` is "super fast". We deflate at level 6, whose encoding is `0b00`.
- **The `0xA220` growth hint is 512 bytes of padding** so an editor can rewrite a part slightly
  larger _in place_ without moving every entry after it. PowerPoint writes five — 520 bytes on each
  of the first two entries and 264 on three more — and 1832 bytes is the **entire** size difference
  between the two files. We rewrite the whole archive on every save, so it would buy nothing.

The same three sat in `zip-writer.ts`'s header table as "as PowerPoint writes", which
`corpus/ground-truth/powerpoint-conventions.json` had already contradicted on two of them since E8.
Nothing in the code was wrong; the documentation beside it was, in the one direction that matters —
a divergence recorded as an agreement.

### What Tier C does not cover

`packages/opc` has two XML serializers of its own — `ContentTypes.serialize()` and
`Relationships.serialize()` — and a no-op write invokes neither, because passthrough means nothing
is dirty. So no committed deck in this corpus carries a part serialized by `packages/opc`. Their
lexical form is not unmeasured (`flat-xml.test.ts` pins the declaration to Office's exact
`standalone="yes"` + CRLF + no BOM, and the tight `/>`), but it is unrepresented in the corpus, and
`C-LEX` has to say so rather than count `c01` as a third XML producer. It is not.

The cost of stopping at one deck is recorded honestly in ADR 0009: for XML lexical form there are
still only two producers, so `C-LEX` rests almost entirely on Tier B, and if Tier B were ever
removed the round-trip gate would degrade to a proof of idempotence.

## What `C-LEX` found

Sixty forms declared in `lexical-forms.ts`, each with the evidence that it occurs outside this
repository — or the counter-evidence, which is a finding of its own. Twenty-nine are covered by two
or more serializers. The thirty-one that are not divide into three kinds, and the difference between
them is the useful part:

**Closable by one more deck — one item, and it is the important one.** Every entity reference in
the corpus was written by `tools/corpus/tiers/a-generated`: the nine PowerPoint-authored decks contain **no
ampersand at all** across their 387 XML parts. So `&amp;`, `&lt;`, `&gt;` and `&quot;` are each
checked only against the escaper that wrote them, which is precisely the failure `C-LEX` was written
to name. It is not a limit of the producer: E8 measured PowerPoint writing all four, and a follow-up
caught `name="a &quot;quoted&quot; name"` on a renamed shape. Those decks were never committed. A
tenth Tier B deck whose title and shape names carry `& < > "` closes four rows at once.

**Closable, but only by a deck of a kind we do not have.** `xml.selfClosing=spaced` is the largest:
70 822 of ADR 0004's 98 777 real self-closing tags are written `<x />`, and the only witness here is
`a36-spaced-tags`, which we wrote. There is a second producer and it is inside PowerPoint —
`ppt/charts/chartEx1.xml` spaced 12 of its 12 in the deck measured on 2026-08-27 — so a Tier B deck
carrying a **ChartEx** chart would supply it. `b05-chart` has `c:` charts and no ChartEx part.

**Not closable at all, and the row says so instead of implying a deck is owed.** ADR 0004 opened 62
lexical variants in the installed PowerPoint and re-saved the ones it accepted: comments and
processing instructions are discarded, a CDATA section is rewritten as plain text, `&#72;` resolves
to `H`, a single-quoted attribute is rewritten double, a byte order mark is stripped. For those
forms our second producer is not silent but incapable. Three container forms — ZIP64, data
descriptors, directory entries — are in the same class for a different reason: ADR 0002 found Office
writes none of them, `tools/ground-truth/lib/zip.ts` cannot write two of them, and all four live in
`packages/opc/src/zip/zip-reader.test.ts` where they belong.

### Two things it found that nothing else had

- **PowerPoint does not always write an XML declaration.** Four of the corpus's 1422 XML parts have
  none: `ppt/charts/style1.xml`, `style2.xml`, `colors1.xml` and `colors2.xml` in `b05-chart`
  begin at `<cs:chartStyle` and `<cs:colorStyle`. ADR 0004 scanned 2834 real parts and E8 scanned
  38, and no part in either lacked one. It is the same fact as the ChartEx spacing seen from another
  angle — there is more than one XML serializer inside PowerPoint, and the chart ones are the odd
  members.
- **The commonest real-world form of three conventions is absent here entirely.** 1037 of ADR 0004's
  2834 parts carry a BOM and none of ours do. 2161 of them put nothing between the declaration and
  the root element; all 1418 declared parts here put a CRLF there. Two of its three declarations
  spell the encoding `utf-8`; neither of our producers does.

### What `C-LEX` may not assert

**Entry order.** The two E8 decks and `b01` disagree about whether `slide1.xml` precedes its own
`.rels`, and `b05` interleaves layouts with their rels in no pattern at all. The order is unstable
_within_ a single producer, so there is no convention there to check — which is a rule about the
rule, not a gap in it.

## What `C-COV` found

**Forty-six of the census's forty-seven feature keys are exercised by a deck**, and the forty-
seventh is `model3d`, declared above. That is the number sub-phase 1.1 set out to reach and the
rule that keeps it honest is `C019`, which runs in `pnpm corpus` on a bare clone.

Two things the number hides, and `coverage.test.ts` pins both so they cannot drift unnoticed:

- **Fourteen of the forty-six rest on a single deck each** — `ink` and `contentPart` on `a27-ink`;
  `audio`, `video` and `media` on `a24-media`; `innerShadow` and `scene3d` on `a04-effects`; and
  `chartEx`, `connector`, `decorative`, `groupFill`, `macros`, `math` and `svgBlip` one apiece. That
  is what a probe corpus looks like when it is built one feature at a time. It is not a violation of
  anything, and it is the list to read first before deleting a deck.
- **A count of zero is not coverage, and neither is a `pinned` deck.** `a01-minimal` uses zeroes to
  say the census looked and found none, which is a statement worth recording and the opposite of
  exercising the feature; a `pinned` entry is the hash of a file that never left the machine that
  made it. Reading either as coverage would let the corpus claim breadth no clone can reproduce.

## Experiments

| id      | question                                                    | status                                                                              |
| ------- | ----------------------------------------------------------- | ----------------------------------------------------------------------------------- |
| **E4**  | which `a:ext` GUID carries `am3d:model3d`                   | **blocked** — needs an online model fetch, not approved; `a28` cut, key declared    |
| **E6**  | what a modern build writes for an OLE object                | ✔ answered — see ADR 0009                                                           |
| **E8**  | the byte conventions: BOM, declaration, escaping, ZIP flags | ✔ answered — `corpus/ground-truth/powerpoint-conventions.json`                      |
| **E11** | where `a:rtl` sits in `CT_TextCharacterProperties`          | ✔ closed by the schema order table: `…sym, hlinkClick, hlinkMouseOver, rtl, extLst` |

## Declared gaps

Not everything a roster row asks for can be measured yet, and a gap that is written down is a
different thing from one nobody noticed.

**Tier B being complete does not close the gaps below that say "Tier B".** Five of them said that;
only `_rels/.rels` was covered by one of the planned nine, and it is struck through. The other four
— a shape-anchored comment, a VBA project that compiles, a video with coded samples, and ink — each
need a tenth deck the roster never planned, and two of them additionally need something this project
may not do on its own: a Trust Center setting, and a machine with a pen. Closing them is a decision
about widening Tier B, not work that got skipped.

Writing `C-LEX` added two more of the same kind, and they are cheaper than any of the four: neither
needs a setting, a device or a network, only a decision to author a tenth deck.

- **Text and a shape name that need escaping.** Every entity reference in this corpus was written by
  `tools/corpus/tiers/a-generated`; the nine PowerPoint-authored decks contain **no ampersand at all** across
  their 387 XML parts, so `&amp;`, `&lt;`, `&gt;` and `&quot;` are each checked only against the
  escaper that wrote them. E8 already measured PowerPoint writing all four — the decks it measured
  were simply never committed. One deck whose title, body text and a shape's `@name` carry
  `& < > "` closes four `C-LEX` rows at once, and it is the largest single thing this corpus is
  missing.
- **A ChartEx chart.** `xml.selfClosing=spaced` is the dominant real-world form — 70 822 of ADR
  0004's 98 777 — and `a36-spaced-tags` is its only witness here, which is to say we are the only
  producer of the commonest form in the wild. PowerPoint does write it, from the one serializer that
  behaves differently: `ppt/charts/chartEx1.xml` spaced 12 of its 12 in the deck measured on
  2026-08-27. `b05-chart` carries `c:` charts and no ChartEx part, so a deck with a waterfall or a
  treemap would supply a genuinely independent producer for the form the round-trip gate most needs
  one for.

- **A modern comment anchored to a shape.** `a15-comments` writes the 2018 `p188:cm` with the
  `pc:sldMkLst` document marker that was measured, and no `pc:spMkLst`. Nothing in any public
  specification describes the shape marker, and `Slide.Comments.Add` produces only the slide one,
  so writing one would be inventing markup. It closes with a Tier B deck.
- **A VBA project the VBA engine would compile.** `a32-macros` carries a real compound file holding
  a real MS-OVBA project, and both layers are checked structurally — `cfb.test.ts` reads the
  container back through an independently written reader, the compressor round-trips through an
  independently written decompressor, and the `dir` record walk lands exactly on the end of the
  stream. None of that is evidence that VBA would **load** it: PowerPoint does not read a VBA
  project when it opens a file, so the deck opening proves the package and not the project. Six
  `Reserved` constants inside `dir`, the `_VBA_PROJECT` version word, and whether a project with no
  `PROJECTREFERENCES` is acceptable are all asserted from the specification and checked by nothing.
  Closing it needs "Trust access to the VBA project object model", which is a Trust Center setting
  on the user's machine rather than something this project may change. Tier B.
- ~~**`_rels/.rels` in the order PowerPoint writes it.**~~ **Closed by `b01-blank`**, which is
  written `rId3`, `rId2`, `rId1`, `rId4`. The gap was never that the order was unmeasured — it was
  that no committed deck exhibited it, because the chassis writes its own three package
  relationships first and appends. Now one does.
- **The eighth ChartEx layout, `regionMap`.** `a22` covers the other seven, every one read back
  from PowerPoint. A region map's `cx:layoutPr/cx:geography` is resolved against an online map
  service, so authoring one is a download and has not been approved.
- **A video with coded samples.** `a24`'s MP4 is a structurally valid ISO-BMFF container with no
  tracks, which is what this repository can author without an encoder; PowerPoint opens a package
  that references it and refuses to _insert_ it. The markup — `a:videoFile`, the `p14:media`
  double relationship, the poster frame — is complete and measured. A clip that plays is Tier B.
- **Ink as PowerPoint writes it.** No COM entry point creates an ink annotation, so `a27`'s
  relationship type was found by bisection rather than measured, and PowerPoint **discards every
  `p:contentPart` on save**. Whether that is our markup or PowerPoint's reluctance to reconstruct
  ink it did not create cannot be told apart without a machine with a pen. Tier B.
- **Where PowerPoint puts an equation.** Same reason: there is no `Shapes.AddEquation`, so
  `a29`'s two MCE placements are both legal and neither is measured. The `m:oMath` content itself
  is transcribed from ECMA-376 Part 1 §22.1, which is complete and precise.
- **A 3-D model.** The only gap that is also a `C-COV` `uncovered` entry, and the only one this
  roster answers by cutting a slot rather than adding one. Everything above is a gap in evidence;
  this is a gap in the corpus's coverage of the census itself, which is why it is declared in a
  manifest and not only here. See `a28-model3d` above for why guessing the four undocumented
  strings would be worse than the hole.
- **An EMF with text and raster blits.** `tools/corpus/tiers/a-generated/assets/emf.ts` writes a five-record metafile
  that genuinely draws, which is enough to be an OLE preview and nothing like enough for sub-phase
  10.6: the hard part is the ~70 record types `rtf.js` routes into a no-op, and a real Visio or
  Excel metafile is tens of kilobytes of exactly those.

## What PowerPoint refuses

Findings from opening probe decks, which is the only feedback loop that exists — PowerPoint emits
no diagnostic log and its refusal message names no part, element or line.

- **A guide name referenced but never defined is a whole-package refusal.** Not a repair, not a
  dropped shape: "PowerPoint could not open the file". It applies both to an `a:gdLst` formula
  naming an adjust value that no `a:avLst` defines, and to an `a:pt` coordinate naming a guide that
  no `a:gdLst` defines. `ST_GeomGuideName` is an unconstrained token, so this is schema-legal
  markup rejected for a reason no schema states, and it belongs in sub-phase 1.2's rules.
- **`p:ph type="hdr"` and `p:ph type="sldImg"` are a whole-package refusal**, on a slide layout
  and on a slide alike, either one alone with no other change. The other seven content types were
  built as one-type packages in the same bisection and every one opens: `obj`, `chart`, `tbl`,
  `clipArt`, `dgm`, `media`, `pic`. Nothing in the schema says so — `CT_Placeholder` is one
  complex type shared by masters, layouts, slides, notes slides and handout masters, and
  `ST_PlaceholderType` is one enumeration holding all sixteen values. The restriction is real and
  unwritten: those two belong to the notes and handout families, and `a14-notes` is where they
  will be exercised legally.
- **A `p:sldMasterId/@id` colliding with a `p:sldLayoutId/@id` is a whole-package refusal**, and
  the two are the same number space. `ST_SlideMasterId` and `ST_SlideLayoutId` are both
  "2147483648 and up" and no schema says they may not overlap; PowerPoint allocates from one running
  counter — master, its layouts, next master, its layouts — and refuses a package that does not.
  Found by bisection when `a12-masters` would not open: the chassis numbered masters and layouts
  from two counters, which is indistinguishable from correct with one master and collides on the
  second, so **every** two-master package was refused. Single-master, inverted `p:clrMap`, a second
  `a:clrScheme` and `a:overrideClrMapping` all opened; two masters with any content did not.
- **`a:ahXY` or `a:cxn` as a direct child of `a:custGeom`** — missing its `a:ahLst`/`a:cxnLst`
  wrapper — is also a whole-package refusal, with the same message.
- **A `c:strLit` inside `c:tx` is a whole-package refusal.** `CT_SerTx` is a choice of
  `c:strRef` or `c:v` and nothing else, even though `c:cat` and `c:val` both accept the literal
  forms two elements away.
- **A `cs:chartStyle` carrying fewer than its thirty-one entries is a whole-package refusal**, and
  a chart with no chart-style relationship at all is fine. Four entries refused; thirty-one opened.
- **A `cx:chartSpace` part with no `.rels` is a whole-package refusal**, including one PowerPoint
  wrote itself. It needs a `chartStyle` and a `chartColorStyle` relationship; the classic
  `c:chartSpace` needs neither.
- **A `.wav` that is both a media object's target and a transition's `p:snd` is a refusal**, and
  the only one so far whose message is "PowerPoint could not open the file" rather than "the file
  or directory is corrupted and unreadable". Two copies of the same bytes open.
- **A `p:contentPart` reached by the ECMA `…/relationships/customXml` type is a refusal.** The
  2010 Microsoft type of the same name opens, and so do two relationship types that make no sense
  at all — PowerPoint validates that one type rather than types in general.
- **A slide without exactly one `slideLayout` relationship, resolving to a layout part, is a
  refusal.** Zero, two, one pointing at a master and one pointing at a missing part were each
  built as a single change and each rejected. Nothing in `CT_Slide` or in OPC requires it, because
  the binding lives only in the rels part.
- **`p:cNvPr/@id` between 2147483648 and 4294967294 is a whole-package refusal**, while every value
  from 0 to 2147483647 opens and 4294967295 opens. `ST_DrawingElementId` is `xsd:unsignedInt`;
  PowerPoint reads it as a signed 32-bit integer and keeps `0xFFFFFFFF` as a sentinel it renumbers
  away on save.
- **A percent-escape of an unreserved character in a part name is a refusal**, with `0x808D1005`.
  RFC 3986 §6.2.2.2 normalisation, enforced: `%2D`, `%41`, `%5F` and `%7E` are rejected while
  `%20`, `%23`, `%24`, `%2C` and `%3A` open. `packages/opc` reports this as a **warning** (`M1.8`);
  the severity is wrong and sub-phase 1.2 should raise it.
- **`p:control` is a whole-package refusal**, in all eight forms tried: bare, name-only, with and
  without `r:id`, with a `p:pic` preview, with an ActiveX part and its `.bin`, and in a
  macro-enabled package. An empty `p:controls` is accepted, which places it precisely on the
  child element.
- **`a:effectDag` is accepted.** PowerPoint never authors one, preferring `a:effectLst` for
  everything a user can build, but it reads a nested `a:cont` graph without complaint and the shape
  survives with its effects.
- Accepted without complaint, all of which are spec features PowerPoint does not itself write:
  `a:hslClr`, `a:scrgbClr`, `a:prstClr`, path gradients with `a:fillToRect`, `a:tileRect`,
  `a:tile/@flip`, `a:blip` colour effects, `a:grpFill`, `a:ahPolar`, and all 54
  `ST_PresetPatternVal` values.

`corpus/reject/` exists as of sub-phase 1.2. Seventeen of the refusals above are now fixtures —
the smallest package we can build carrying the markup that was measured — and `C-REJECT` asserts
the validator catches every one, by the rule its manifest names, before an export is handed over.
Each entry records what PowerPoint said and where the finding is written down.

One limitation, stated where it belongs rather than in a footnote: the **markup** in each fixture
is the markup that was refused, but the **package** around it is new and much smaller, and nobody
has opened those seventeen files in PowerPoint. `C-REJECT` asserts only what is checkable without
doing so - that we refuse them, and for the right reason. Confirming that the minimal package
reproduces the refusal is a run of `gen/open-in-powerpoint.ps1` away, on a machine with PowerPoint,
by somebody who chooses to.
