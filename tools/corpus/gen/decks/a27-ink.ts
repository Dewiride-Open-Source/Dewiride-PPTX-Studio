import { DECLARATION, REL, type ProbePart, type ProbeRel } from '../package.ts';
import { quadrantPng } from '../png.ts';
import { picture, shape } from '../shapes.ts';
import { textLine, txBody } from '../text.ts';
import type { ProbeDeck } from '../types.ts';

/**
 * Ink: `p:contentPart`, which in the Transitional schema has no children at all.
 *
 * `CT_Rel` is the whole of `p:contentPart` in ECMA-376 - one `r:id` attribute
 * and nothing else. Everything that makes an ink annotation drawable is
 * extension markup in the `p14` namespace hanging off an element the schema
 * says is empty, which is why slide 2 here is a bare
 * `<p:contentPart r:id="rId3"/>`: it is what the standard actually defines, it
 * is legal, and it carries no position, no size and no name.
 *
 * Slide 1 is what a real pen stroke looks like:
 *
 * ```xml
 * <mc:AlternateContent>
 *   <mc:Choice xmlns:p14="…/powerpoint/2010/main" Requires="p14">
 *     <p:contentPart r:id="rId2" p14:bwMode="auto">
 *       <p14:nvContentPartPr>…</p14:nvContentPartPr>
 *       <p14:xfrm><a:off …/><a:ext …/></p14:xfrm>
 *     </p:contentPart>
 *   </mc:Choice>
 *   <mc:Fallback><p:pic>… a raster of the strokes …</p:pic></mc:Fallback>
 * </mc:AlternateContent>
 * ```
 *
 * Note where the transform lives: `p14:xfrm`, a third namespace for the same
 * `a:off`/`a:ext` pair that `a:xfrm` and `p:xfrm` also hold. Three spellings,
 * identical children.
 *
 * ## The traces are polylines with a header explaining the columns
 *
 * `inkml:traceFormat` declares the channels in order - X, Y, and usually F for
 * pen force - and every point in every `inkml:trace` is that many numbers.
 * Read the channel list, map X and Y to a polyline with round caps, and ink is
 * done; the rest of InkML is metadata.
 *
 * The catch is the encoding. A trace value may carry a prefix:
 *
 * | prefix | means                                    |
 * | ------ | ---------------------------------------- |
 * | none   | an absolute value                        |
 * | `'`    | a first difference from the previous point |
 * | `"`    | a second difference                      |
 *
 * so `10 20 100, '5 '0 '0` is two points, at (10,20) and (15,20). Both forms
 * are here - trace 1 absolute, trace 2 differenced - because a parser that
 * assumes absolute reads the second one as a stroke that jumps to the origin.
 *
 * Slide 3 adds `inkml:traceGroup` with an `inkml:annotationXML`, which is how a
 * recogniser attaches "this is the word *corpus*" to a set of strokes, and two
 * brushes so a renderer has to read `@brushRef` per trace rather than once.
 *
 * ## The relationship type, which bisection had to find
 *
 * Every other part of the Hard content group was read back from a file
 * PowerPoint wrote. Ink could not be: there is no COM entry point that creates
 * an ink annotation - `AddInk` does not exist and `AddPolyline` produces a
 * `p:sp` with `a:custGeom`, not a content part. So the relationship type was
 * written from the specification, as
 * `…/officeDocument/2006/relationships/customXml`, and **PowerPoint refused the
 * package**.
 *
 * The bisection is worth recording because the result is narrow and useful.
 * With the ink parts present and nothing pointing at them: opens. With a
 * relationship to an ink part and no `p:contentPart`: opens. With a
 * `p:contentPart` naming that relationship: refused - and retyping the ink part
 * as `application/xml` changed nothing, so it is not about the content type.
 * Then the same package with the relationship's `@Type` changed, and nothing
 * else:
 *
 * | `@Type`                                              | result   |
 * | ---------------------------------------------------- | -------- |
 * | `…/officeDocument/2006/relationships/customXml`      | refused  |
 * | `…/office/2010/relationships/customXml`              | opens    |
 * | `…/officeDocument/2006/relationships/image`          | opens    |
 * | `…/officeDocument/2006/relationships/slide`          | opens    |
 *
 * So PowerPoint does not validate the type in general - two nonsense types open
 * fine - it validates **that one**. The ECMA `customXml` type means the
 * document's custom XML data store, which comes with an `itemProps` companion
 * and a shape PowerPoint knows; handed an InkML part instead, it refuses the
 * whole package rather than the part. The 2010 Microsoft type is what this deck
 * uses, and it is very probably what PowerPoint writes, but that half is still
 * inference: `ROSTER.md` keeps the gap open for a Tier B deck on a machine with
 * a pen.
 */

const NS_INKML = 'http://www.w3.org/2003/InkML';
const NS_P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main';
const NS_A14 = 'http://schemas.microsoft.com/office/drawing/2010/main';
const NS_MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';

/**
 * Ink parts are related as custom XML - the **2010 Microsoft** one.
 *
 * Not `REL + 'customXml'`. That is the ECMA type for the document's custom XML
 * data store, and pointing it at an InkML part is a whole-package refusal. See
 * the file comment.
 */
const REL_INK = 'http://schemas.microsoft.com/office/2010/relationships/customXml';
const CT_INK = 'application/inkml+xml';

const FRAME = { x: 914400, y: 1600200, cx: 4572000, cy: 3048000 } as const;

// -------------------------------------------------------------- the ink part

/**
 * `inkml:traceFormat` and one brush.
 *
 * HIMETRIC is the unit PowerPoint's ink uses: hundredths of a millimetre, so
 * 2540 to the inch and 360 EMU to the unit.
 */
const context = (id: string, brushes: string): string =>
  '<inkml:definitions>' +
  `<inkml:context xml:id="${id}">` +
  '<inkml:inkSource xml:id="inkSrc0">' +
  '<inkml:traceFormat>' +
  '<inkml:channel name="X" type="integer" max="32767" units="himetric"/>' +
  '<inkml:channel name="Y" type="integer" max="32767" units="himetric"/>' +
  '<inkml:channel name="F" type="integer" max="32767" units="dev"/>' +
  '</inkml:traceFormat>' +
  '<inkml:channelProperties>' +
  '<inkml:channelProperty channel="X" name="resolution" value="1" units="1/himetric"/>' +
  '<inkml:channelProperty channel="Y" name="resolution" value="1" units="1/himetric"/>' +
  '<inkml:channelProperty channel="F" name="resolution" value="1" units="1/dev"/>' +
  '</inkml:channelProperties>' +
  '</inkml:inkSource>' +
  brushes +
  '</inkml:context>' +
  '</inkml:definitions>';

const brush = (id: string, width: number, colour: string): string =>
  `<inkml:brush xml:id="${id}">` +
  `<inkml:brushProperty name="width" value="${String(width)}" units="himetric"/>` +
  `<inkml:brushProperty name="height" value="${String(width)}" units="himetric"/>` +
  `<inkml:brushProperty name="color" value="${colour}"/>` +
  '<inkml:brushProperty name="ignorePressure" value="0"/>' +
  '</inkml:brush>';

const trace = (contextRef: string, brushRef: string, data: string, extra = ''): string =>
  `<inkml:trace contextRef="#${contextRef}" brushRef="#${brushRef}"${extra}>${data}</inkml:trace>`;

const ink = (body: string): string =>
  DECLARATION + `<inkml:ink xmlns:inkml="${NS_INKML}">` + body + '</inkml:ink>';

/** Absolute values throughout: three numbers a point, comma between points. */
const STROKE_ABSOLUTE =
  '1000 3000 800, 1400 2400 900, 1900 2000 1000, 2500 1900 1000, ' +
  '3100 2100 900, 3600 2600 800, 3900 3200 700';

/** The same shape, written as first differences. See the table above. */
const STROKE_DIFFERENCE =
  "1000 4200 800, '400 '-300 '100, '500 '-200 '100, '600 '-100 '0, " +
  "'600 '200 '-100, '500 '500 '-100, '300 '600 '-100";

const INK_1 = ink(
  context('ctx0', brush('br0', 105, '#C00000') + brush('br1', 210, '#0070C0')) +
    trace('ctx0', 'br0', STROKE_ABSOLUTE) +
    trace('ctx0', 'br1', STROKE_DIFFERENCE),
);

const INK_2 = ink(
  context('ctx0', brush('br0', 105, '#7030A0')) +
    trace('ctx0', 'br0', '500 500 900, 4000 500 900, 4000 2500 900, 500 2500 900, 500 500 900'),
);

/**
 * A recognised group.
 *
 * `inkml:traceGroup` collects strokes that mean one thing, and
 * `inkml:annotationXML` is where a handwriting recogniser puts its answer. It
 * is a different namespace inside an InkML element inside a PresentationML
 * package, three deep, and it is preserved rather than read.
 */
const INK_3 = ink(
  context('ctx0', brush('br0', 105, '#538135') + brush('br1', 53, '#BF8F00')) +
    '<inkml:traceGroup xml:id="tg0">' +
    '<inkml:annotation type="writingRegion">corpus</inkml:annotation>' +
    '<inkml:annotationXML type="text" encoding="text">corpus</inkml:annotationXML>' +
    trace(
      'ctx0',
      'br0',
      '600 2200 900, 900 1900 900, 1300 1900 900, 1500 2200 900',
      ' xml:id="t0"',
    ) +
    trace('ctx0', 'br0', '1800 1900 900, 1800 2500 900', ' xml:id="t1"') +
    '</inkml:traceGroup>' +
    trace('ctx0', 'br1', '600 2900 700, 3800 2900 700', ' xml:id="t2"'),
);

// ------------------------------------------------------------ the shape tree

/**
 * `p:contentPart` with everything PowerPoint hangs off it.
 *
 * `p14:nvContentPartPr` mirrors a shape's `nvSpPr` exactly - a `cNvPr` with an
 * id and a name, a locks element, and an empty `nvPr` - except that all three
 * are in `p14`, so a reader keyed on the PresentationML names finds none of
 * them.
 */
const contentPart = (spec: {
  readonly id: number;
  readonly name: string;
  readonly relId: string;
  readonly x: number;
  readonly y: number;
  readonly cx: number;
  readonly cy: number;
}): string =>
  `<p:contentPart r:id="${spec.relId}" p14:bwMode="auto">` +
  '<p14:nvContentPartPr>' +
  `<p14:cNvPr id="${String(spec.id)}" name="${spec.name}"/>` +
  `<p14:cNvContentPartPr><a14:cpLocks xmlns:a14="${NS_A14}" noRot="1" noChangeAspect="1"` +
  ' noMove="1" noResize="1" noEditPoints="1" noAdjustHandles="1" noChangeArrowheads="1"' +
  ' noChangeShapeType="1"/></p14:cNvContentPartPr>' +
  '<p14:nvPr/>' +
  '</p14:nvContentPartPr>' +
  // A third namespace for the same two children a:xfrm and p:xfrm hold.
  `<p14:xfrm xmlns:a="${NS_A}">` +
  `<a:off x="${String(spec.x)}" y="${String(spec.y)}"/>` +
  `<a:ext cx="${String(spec.cx)}" cy="${String(spec.cy)}"/>` +
  '</p14:xfrm>' +
  '</p:contentPart>';

const switchedInk = (id: number, name: string, relId: string, rasterRelId: string): string =>
  `<mc:AlternateContent xmlns:mc="${NS_MC}">` +
  `<mc:Choice xmlns:p14="${NS_P14}" Requires="p14">` +
  contentPart({ id, name, relId, ...FRAME }) +
  '</mc:Choice>' +
  '<mc:Fallback>' +
  picture({ id, name, relId: rasterRelId, ...FRAME, description: 'A raster of the ink strokes' }) +
  '</mc:Fallback>' +
  '</mc:AlternateContent>';

// ------------------------------------------------------------------ package

const PARTS: readonly ProbePart[] = [
  { name: 'ppt/ink/ink1.xml', bytes: INK_1, contentType: { kind: 'override', type: CT_INK } },
  { name: 'ppt/ink/ink2.xml', bytes: INK_2, contentType: { kind: 'override', type: CT_INK } },
  { name: 'ppt/ink/ink3.xml', bytes: INK_3, contentType: { kind: 'override', type: CT_INK } },
  {
    name: 'ppt/media/image1.png',
    bytes: quadrantPng(),
    contentType: { kind: 'default', extension: 'png', type: 'image/png' },
  },
];

const inkRel = (relId: string, number: number): ProbeRel => ({
  id: relId,
  type: REL_INK,
  target: `../ink/ink${String(number)}.xml`,
});

const PNG_REL: ProbeRel = { id: 'rId3', type: REL + 'image', target: '../media/image1.png' };

const caption = (id: number, lines: readonly string[]): string =>
  shape({
    id,
    name: 'Caption',
    x: 5943600,
    y: 1600200,
    cx: 5562600,
    cy: 3048000,
    textBody: txBody({
      bodyPr: '<a:bodyPr wrap="square" anchor="t"/>',
      paras: lines.map((line) => textLine(line, { sz: 1300 })).join(''),
    }),
  });

export const a27Ink: ProbeDeck = {
  id: 'a27-ink',
  title: 'PPTX Studio corpus: a27 ink',
  description:
    'Ink annotations in all three shapes they take: the mc:AlternateContent PowerPoint writes, ' +
    'with p14:nvContentPartPr, p14:xfrm and a rasterised p:pic fallback; a bare p:contentPart with ' +
    'nothing but an r:id, which is the whole of CT_Rel and the only form ECMA-376 defines; and a ' +
    'part whose traces are grouped under inkml:traceGroup with a recogniser annotation. Traces are ' +
    'written both absolute and as first differences, because a parser that assumes absolute reads ' +
    'the second stroke as a jump to the origin.',
  features: {
    // 6 chassis + one caption a slide.
    shape: 9,
    placeholder: 6,
    gradientFill: 2,
    contentPart: 3,
    // Two traces, one, and three - two of them inside a traceGroup.
    ink: 6,
    picture: 1,
    presetGeom: 4,
    // Slide 1, with a fallback, and slide 3, deliberately without one.
    alternateContent: 2,
  },
  build: () => ({
    title: 'PPTX Studio corpus: a27 ink',
    slides: [
      {
        title: 'a27 — a content part, and the picture of it underneath',
        rels: [inkRel('rId2', 1), PNG_REL],
        body:
          switchedInk(10, 'Ink 1', 'rId2', 'rId3') +
          caption(11, [
            'p14:nvContentPartPr mirrors nvSpPr - cNvPr, locks, nvPr - in p14.',
            'The transform is p14:xfrm: a third namespace holding a:off and a:ext.',
            'The fallback is an ordinary p:pic, so a reader that discards the',
            'Choice still shows the strokes, flattened and unselectable.',
          ]),
      },
      {
        title: 'a27 — the whole of p:contentPart, as ECMA-376 defines it',
        rels: [inkRel('rId2', 2)],
        body:
          '<p:contentPart xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
          ' r:id="rId2"/>' +
          caption(10, [
            'CT_Rel is one attribute. No position, no size, no name, no locks.',
            'Everything else is p14 extension markup on an element the schema',
            'says is empty - which is legal, and is why a strict reader that',
            'refuses unknown children cannot render ink at all.',
            'Where does this draw? Nothing in the package says. That is the point.',
          ]),
      },
      {
        title: 'a27 — traces grouped, annotated, and drawn with two brushes',
        rels: [inkRel('rId2', 3)],
        body:
          `<mc:AlternateContent xmlns:mc="${NS_MC}">` +
          `<mc:Choice xmlns:p14="${NS_P14}" Requires="p14">` +
          contentPart({ id: 10, name: 'Ink 3', relId: 'rId2', ...FRAME }) +
          '</mc:Choice>' +
          // No fallback branch at all. `mc:Fallback` is optional, and a switch
          // without one means "if you cannot do p14, draw nothing".
          '</mc:AlternateContent>' +
          caption(11, [
            'inkml:traceGroup collects the strokes a recogniser read as one word,',
            'and inkml:annotationXML carries what it decided they said.',
            'Two brushes, so @brushRef has to be read per trace.',
            'This switch has no mc:Fallback, which is legal: it means a reader',
            'that cannot do p14 should draw nothing rather than something wrong.',
          ]),
      },
    ],
    parts: PARTS,
  }),
};
