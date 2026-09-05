import { basePackage, OD, rel, relsPart, type PartOverrides } from './deck.ts';

/**
 * The packages PowerPoint refuses, one change each.
 *
 * ## What each of these is, and what it is not
 *
 * Every fixture below reproduces a refusal recorded in
 * `tools/corpus/ROSTER.md` under **What PowerPoint refuses**. Those findings
 * were measured the only way they can be: build a package, open it in
 * PowerPoint 16.0.20326, and read the one sentence it gives back. There is no
 * log, no part named, no line number - `refusal` on each entry is the whole of
 * the diagnostic channel, which is why it is recorded verbatim.
 *
 * **The fixtures themselves were measured in sub-phase 1.5**, and the result
 * changes one word on most of these entries. All seventeen were opened in
 * PowerPoint 16.0.20326 on 2026-08-29 through
 * `packages/cli/scripts/powerpoint-oracle.ps1`:
 *
 *   - With `OpenAndRepair` **off**, every one of the seventeen fails to open.
 *     So each minimal package does reproduce a failure, which is what was owed.
 *   - With `OpenAndRepair` **on**, every one of the seventeen *opens*, repaired,
 *     down to a single slide holding a single shape - two for `r17`.
 *
 * That makes them **repairs rather than outright refusals**, and several entries
 * below say "whole-package refusal". Those words are not being corrected,
 * because they record what was measured on the original probe packages, which
 * were larger; this note records what the minimal fixtures do. The two can
 * legitimately differ, and the honest summary is that the minimal fixtures land
 * in the repair bucket. Nothing here has reproduced PowerPoint's "could not open
 * the file" path, which is a third outcome again.
 *
 * Nothing in `C-REJECT` depends on that confirmation. What the rule asserts is
 * that **our validator refuses each of these before the bytes are handed over**,
 * which is a claim about us and is checkable here. Whether PowerPoint agrees
 * with our reason for refusing is the claim `open-in-powerpoint.ps1` settles,
 * and the two are worth keeping apart.
 *
 * ## Why `kind: "fixture"` and not `kind: "deck"`
 *
 * These are `.pptx` files, so calling them decks would be the obvious choice
 * and it would corrupt two other rules. `C-COV` counts a census feature as
 * covered when some deck reports a non-zero count for it - and a package
 * PowerPoint refuses is not evidence that a feature works. `C-LEX` counts
 * distinct XML producers, and these share ours. So they are fixtures with
 * `tags`, which is the shape for a committed artefact that makes no census
 * claim.
 */

export interface RejectFixture {
  /** Kebab-case, matching the manifest entry id and the file name. */
  readonly id: string;
  /** The validator rule this package must trip, fatally. */
  readonly rule: string;
  /** What the fixture is: the one change, in a sentence. */
  readonly description: string;
  /** What PowerPoint said, and where the measurement is written down. */
  readonly refusal: string;
  /** Parts replacing or added to `basePackage()`. `null` removes one. */
  readonly parts: PartOverrides;
}

const base = basePackage();

/** Replace inside a base part, failing loudly if the anchor has drifted. */
function edit(part: string, from: string, to: string): string {
  const source = base[part];
  if (source === undefined) throw new Error('no base part named ' + part);
  if (!source.includes(from)) {
    throw new Error('the anchor ' + JSON.stringify(from) + ' is not in ' + part);
  }
  return source.replace(from, to);
}

const SLIDE = 'ppt/slides/slide1.xml';
const SLIDE_RELS = 'ppt/slides/_rels/slide1.xml.rels';
const TITLE_PH = '<p:ph type="title"/>';
const RECT = '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>';
const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

/** A `custGeom` with the guide lists filled in as given. */
function custGeom(inner: string): string {
  return '<a:custGeom>' + inner + '</a:custGeom>';
}

const CHART_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const CHART_STYLE_NS = 'http://schemas.microsoft.com/office/drawing/2012/chartStyle';
const CHART_EX_NS = 'http://schemas.microsoft.com/office/drawing/2014/chartex';

/** An extra `<Override>` on the end of the content-type stream. */
function withOverride(partName: string, contentType: string): string {
  return edit(
    '[Content_Types].xml',
    '</Types>',
    `<Override PartName="${partName}" ContentType="${contentType}"/></Types>`,
  );
}

export const REJECT_FIXTURES: readonly RejectFixture[] = [
  {
    id: 'r01-ph-hdr',
    rule: 'V022',
    description:
      'A slide placeholder typed `hdr`, which belongs to the notes and handout families.',
    refusal:
      'Whole-package refusal, on a slide and on a layout alike, either one alone with no other ' +
      'change. The other seven content types were built as one-type packages in the same ' +
      'bisection and every one opens: obj, chart, tbl, clipArt, dgm, media, pic. ' +
      'ROSTER.md, "What PowerPoint refuses".',
    parts: { [SLIDE]: edit(SLIDE, TITLE_PH, '<p:ph type="hdr"/>') },
  },
  {
    id: 'r02-ph-sldimg',
    rule: 'V022',
    description:
      'The same, typed `sldImg`. The second of the two placeholder types PowerPoint refuses.',
    refusal:
      'Whole-package refusal. Nothing in the schema says so: CT_Placeholder is one complex type ' +
      'shared by every sheet family and ST_PlaceholderType is one enumeration holding all ' +
      'sixteen values. ROSTER.md, "What PowerPoint refuses".',
    parts: { [SLIDE]: edit(SLIDE, TITLE_PH, '<p:ph type="sldImg"/>') },
  },
  {
    id: 'r03-guide-in-formula',
    rule: 'V023',
    description: 'An `a:gd` formula naming `adj1`, which no `a:avLst` in the shape defines.',
    refusal:
      'Whole-package refusal - "PowerPoint could not open the file". Not a repair and not a ' +
      'dropped shape. ST_GeomGuideName is an unconstrained token, so this is schema-legal markup ' +
      'refused for a reason no schema states. ROSTER.md, "What PowerPoint refuses".',
    parts: {
      [SLIDE]: edit(
        SLIDE,
        RECT,
        custGeom(
          '<a:avLst/><a:gdLst><a:gd name="x1" fmla="*/ w adj1 100000"/></a:gdLst>' +
            '<a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst/>',
        ),
      ),
    },
  },
  {
    id: 'r04-guide-in-point',
    rule: 'V023',
    description: 'An `a:pt` coordinate naming a guide no `a:gdLst` defines.',
    refusal:
      'The other half of the same finding, and it was measured separately: a coordinate naming a ' +
      'missing guide is refused just as a formula naming one is. ROSTER.md, ' +
      '"What PowerPoint refuses".',
    parts: {
      [SLIDE]: edit(
        SLIDE,
        RECT,
        custGeom(
          '<a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/>' +
            '<a:pathLst><a:path w="21600" h="21600">' +
            '<a:moveTo><a:pt x="notAGuide" y="0"/></a:moveTo>' +
            '<a:lnTo><a:pt x="r" y="b"/></a:lnTo>' +
            '</a:path></a:pathLst>',
        ),
      ),
    },
  },
  {
    id: 'r05-ahxy-unwrapped',
    rule: 'V012',
    description: 'An `a:ahXY` as a direct child of `a:custGeom`, without its `a:ahLst` wrapper.',
    refusal:
      'Whole-package refusal, same message. The wrapper is not decoration: the adjust-handle list ' +
      'is where the schema puts handles, and a handle outside it is markup with nowhere to be. ' +
      'ROSTER.md, "What PowerPoint refuses".',
    parts: {
      [SLIDE]: edit(
        SLIDE,
        RECT,
        custGeom(
          '<a:avLst/><a:gdLst/><a:ahXY><a:pos x="0" y="0"/></a:ahXY><a:cxnLst/>' +
            '<a:rect l="l" t="t" r="r" b="b"/><a:pathLst/>',
        ),
      ),
    },
  },
  {
    id: 'r06-cxn-unwrapped',
    rule: 'V012',
    description: 'An `a:cxn` as a direct child of `a:custGeom`, without its `a:cxnLst` wrapper.',
    refusal:
      'Whole-package refusal, measured alongside `a:ahXY` and refused the same way. ROSTER.md, ' +
      '"What PowerPoint refuses".',
    parts: {
      [SLIDE]: edit(
        SLIDE,
        RECT,
        custGeom(
          '<a:avLst/><a:gdLst/><a:ahLst/><a:cxn ang="0"><a:pos x="0" y="0"/></a:cxn>' +
            '<a:rect l="l" t="t" r="r" b="b"/><a:pathLst/>',
        ),
      ),
    },
  },
  {
    id: 'r07-ser-strlit',
    rule: 'V024',
    description:
      'A chart series whose `c:tx` holds a `c:strLit` rather than a `c:strRef` or `c:v`.',
    refusal:
      'Whole-package refusal. CT_SerTx is a choice of c:strRef or c:v and nothing else, even ' +
      'though c:cat and c:val - two elements away in the same series - both accept the literal ' +
      'forms. ROSTER.md, "Six refusals, all found by bisection".',
    parts: {
      '[Content_Types].xml': withOverride(
        '/ppt/charts/chart1.xml',
        'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
      ),
      'ppt/charts/chart1.xml':
        DECL +
        `<c:chartSpace xmlns:c="${CHART_NS}"><c:chart><c:plotArea><c:barChart><c:ser>` +
        '<c:idx val="0"/><c:order val="0"/>' +
        '<c:tx><c:strLit><c:ptCount val="1"/><c:pt idx="0"><c:v>Series 1</c:v></c:pt></c:strLit>' +
        '</c:tx></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>',
    },
  },
  {
    id: 'r08-chartstyle-subset',
    rule: 'V026',
    description: 'A `cs:chartStyle` carrying four of its thirty-one entries.',
    refusal:
      'Whole-package refusal - "The file or directory is corrupted and unreadable." Four entries ' +
      'refused; thirty-one opened; a chart with no chart-style relationship at all is fine, so a ' +
      'subset is worse than an absence. ROSTER.md, "Six refusals, all found by bisection".',
    parts: {
      '[Content_Types].xml': withOverride(
        '/ppt/charts/style1.xml',
        'application/vnd.ms-office.chartstyle+xml',
      ),
      'ppt/charts/style1.xml':
        DECL +
        `<cs:chartStyle xmlns:cs="${CHART_STYLE_NS}" id="201">` +
        '<cs:chartArea/><cs:dataPoint/><cs:legend/><cs:plotArea/>' +
        '</cs:chartStyle>',
    },
  },
  {
    id: 'r09-chartex-no-rels',
    rule: 'V009',
    description: 'A `cx:chartSpace` part with no relationship part of its own.',
    refusal:
      'Whole-package refusal, including for a chartEx1.xml PowerPoint wrote itself: lifted byte ' +
      'for byte out of a saved deck, put in a package with nothing else changed, and refused. ' +
      'Adding back its chartStyle and chartColorStyle relationships made the same package open. ' +
      'A classic c:chartSpace needs neither. ROSTER.md, "Six refusals, all found by bisection".',
    parts: {
      '[Content_Types].xml': withOverride(
        '/ppt/charts/chartEx1.xml',
        'application/vnd.ms-office.chartex+xml',
      ),
      'ppt/charts/chartEx1.xml':
        DECL +
        `<cx:chartSpace xmlns:cx="${CHART_EX_NS}"><cx:chartData><cx:data id="0">` +
        '<cx:strDim type="cat"><cx:lvl ptCount="1"><cx:pt idx="0">A</cx:pt></cx:lvl></cx:strDim>' +
        '</cx:data></cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion>' +
        '<cx:series layoutId="treemap" uniqueId="{1}"/>' +
        '</cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>',
    },
  },
  {
    id: 'r10-slide-no-layout',
    rule: 'V009',
    description: 'A slide with no `slideLayout` relationship at all.',
    refusal:
      'Whole-package refusal. Zero, two, one pointing at a master and one pointing at a missing ' +
      'part were each built as a single change and each rejected. Nothing in CT_Slide or in OPC ' +
      'requires it, because the binding lives only in the rels part - which is why "change ' +
      'layout" is a relationship rewrite. ROSTER.md, "Three more refusals, all found by ' +
      'bisection".',
    parts: { [SLIDE_RELS]: relsPart('') },
  },
  {
    id: 'r11-slide-layout-is-master',
    rule: 'V009',
    description: 'A slide whose one `slideLayout` relationship resolves to the master.',
    refusal:
      'Whole-package refusal, one of the four variants in the same bisection. The relationship ' +
      'type is right and the target part exists; it is the wrong kind of part. ROSTER.md, ' +
      '"Three more refusals, all found by bisection".',
    parts: {
      [SLIDE_RELS]: relsPart(rel('rId1', OD + 'slideLayout', '../slideMasters/slideMaster1.xml')),
    },
  },
  {
    id: 'r12-sheet-id-collision',
    rule: 'V019',
    description: 'A `p:sldLayoutId` carrying the same number as the `p:sldMasterId`.',
    refusal:
      'Whole-package refusal. The two are one number space and no schema says so; PowerPoint ' +
      'allocates from a single running counter - master, its layouts, next master, its layouts. ' +
      'Found by bisection when a12-masters would not open, because two counters are ' +
      'indistinguishable from correct with one master and collide on the second. ROSTER.md, ' +
      '"What PowerPoint refuses".',
    parts: {
      'ppt/slideMasters/slideMaster1.xml': edit(
        'ppt/slideMasters/slideMaster1.xml',
        '<p:sldLayoutId id="2147483649"',
        '<p:sldLayoutId id="2147483648"',
      ),
    },
  },
  {
    id: 'r13-cnvpr-negative',
    rule: 'V020',
    description: 'A `p:cNvPr/@id` of 3000000000, inside the range PowerPoint reads as negative.',
    refusal:
      'Whole-package refusal. ST_DrawingElementId is xsd:unsignedInt and every value from 0 to ' +
      '2147483647 opens; 2147483648 to 4294967294 are refused; 4294967295 opens, because it is ' +
      'minus one and PowerPoint keeps it as a sentinel it renumbers away on save. a39-large-ids ' +
      'is the deck that measured the boundaries. ROSTER.md, "What PowerPoint refuses".',
    parts: {
      [SLIDE]: edit(
        SLIDE,
        '<p:cNvPr id="2" name="Title 1"/>',
        '<p:cNvPr id="3000000000" name="Title 1"/>',
      ),
    },
  },
  {
    id: 'r14-percent-escape',
    rule: 'V004',
    description: 'A part name percent-escaping an underscore, which is an unreserved character.',
    refusal:
      'Whole-package refusal with 0x808D1005 - RFC 3986 §6.2.2.2 normalisation, enforced rather ' +
      'than recommended. Measured nine for nine on one-name packages: %2D, %41, %5F and %7E ' +
      'refused; %20, %23, %24, %2C and %3A open. packages/opc rates this a warning (M1.8) ' +
      'because a reader must open what it is given. ROSTER.md, "What PowerPoint refuses".',
    parts: {
      'ppt/tags/tag%5F1.xml':
        DECL + '<p:tagLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>',
    },
  },
  {
    id: 'r15-control',
    rule: 'V025',
    description: 'A `p:control` inside an otherwise empty `p:controls`.',
    refusal:
      'Whole-package refusal in all eight forms tried: bare, name-only, with and without r:id, ' +
      'with a p:pic preview, with an ActiveX part and its .bin, and inside a macro-enabled ' +
      'package. An empty p:controls is accepted, which places the refusal precisely on the child ' +
      'element. ROSTER.md, "What PowerPoint refuses".',
    parts: {
      [SLIDE]: edit(
        SLIDE,
        '</p:cSld>',
        '<p:controls><p:control name="Button1"/></p:controls></p:cSld>',
      ),
    },
  },
  {
    id: 'r16-shared-sound',
    rule: 'V009',
    description: 'One `.wav` that is both a transition sound and a media object’s target.',
    refusal:
      'A refusal, and the only one so far whose message is "PowerPoint could not open the file" ' +
      'rather than "the file or directory is corrupted and unreadable". Two copies of the same ' +
      'bytes open; one shared copy does not. ROSTER.md, "What PowerPoint refuses".',
    parts: {
      '[Content_Types].xml': withOverride('/ppt/media/media1.wav', 'audio/wav'),
      // Not a real RIFF file. Nothing in the package reads it, and a fixture
      // whose point is a relationship graph should not carry a payload whose
      // only job is to look convincing.
      'ppt/media/media1.wav': 'RIFF----WAVEfmt ',
      [SLIDE]: edit(
        SLIDE,
        '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>',
        '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>' +
          '<p:transition><p:sndAc><p:stSnd>' +
          '<p:snd r:embed="rId2" name="media1.wav"/>' +
          '</p:stSnd></p:sndAc></p:transition>',
      ).replace(
        '<p:nvPr><p:ph type="title"/></p:nvPr>',
        '<p:nvPr><p:ph type="title"/><a:audioFile r:link="rId2"/></p:nvPr>',
      ),
      [SLIDE_RELS]: relsPart(
        rel('rId1', OD + 'slideLayout', '../slideLayouts/slideLayout1.xml') +
          rel('rId2', OD + 'audio', '../media/media1.wav'),
      ),
    },
  },
  {
    id: 'r17-contentpart-customxml',
    rule: 'V009',
    description: 'A `p:contentPart` reached by the ECMA `customXml` relationship type.',
    refusal:
      'A refusal. The Microsoft 2010 type of the same name opens, and so do two relationship ' +
      'types that make no sense at all - PowerPoint validates that one type rather than types in ' +
      'general. ROSTER.md, "What PowerPoint refuses".',
    parts: {
      'ppt/ink/ink1.xml': DECL + '<inkml:ink xmlns:inkml="http://www.w3.org/2003/InkML"/>',
      [SLIDE]: edit(SLIDE, '</p:spTree>', '<p:contentPart r:id="rId2"/></p:spTree>'),
      [SLIDE_RELS]: relsPart(
        rel('rId1', OD + 'slideLayout', '../slideLayouts/slideLayout1.xml') +
          rel('rId2', OD + 'customXml', '../ink/ink1.xml'),
      ),
    },
  },
];

/** `<id>.pptx`, the only name a reject fixture is written under. */
export function outputName(fixture: RejectFixture): string {
  return fixture.id + '.pptx';
}
