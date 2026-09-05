import { cNvPrXml, type ProbePart } from '../package.ts';
import { quadrantPng } from '../png.ts';
import { shape } from '../shapes.ts';
import { textLine, txBody } from '../text.ts';
import type { ProbeDeck } from '../types.ts';

/**
 * SVG blips: the form PowerPoint wrote, the form with a fallback, and the SVG
 * a renderer must not trust.
 *
 * ## The measurement, which was not what the plan assumed
 *
 * On 2026-08-27 an SVG was inserted through `Shapes.AddPicture`. PowerPoint
 * 16.0.20326 wrote this, and **nothing else**:
 *
 * ```xml
 * <p:blipFill>
 *   <a:blip>
 *     <a:extLst>
 *       <a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">
 *         <asvg:svgBlip xmlns:asvg="…/2016/SVG/main" r:embed="rId2"/>
 *       </a:ext>
 *     </a:extLst>
 *   </a:blip>
 *   <a:stretch><a:fillRect/></a:stretch>
 * </p:blipFill>
 * ```
 *
 * **`a:blip` has no `@r:embed` at all**, and no PNG part was written. The only
 * reference to any image in the whole picture is the one inside the extension.
 * So a renderer that reads `a:blip/@r:embed`, finds nothing and stops draws an
 * empty rectangle for a picture that is perfectly well formed - and the
 * extension it skipped was not an enhancement, it was the entire content.
 *
 * That is the opposite of the shape sub-phase 10.7 was written against, where
 * the blip carries a raster and the extension carries the vector original. Both
 * exist. Slide 1 has one of each so neither can be assumed.
 *
 * `@uri` is `{96DAC541-7B7A-43D3-8B79-37D633B846F1}` and the namespace is
 * `http://schemas.microsoft.com/office/drawing/2016/SVG/main` - both measured
 * rather than transcribed.
 *
 * ## A blip is not only a picture
 *
 * Slide 2 puts the same extension inside `a:blipFill` used as a **shape fill**,
 * which is a different element in a different namespace from `p:pic`'s
 * `p:blipFill` and the one the census counts. A renderer that special-cases SVG
 * in its picture path and not in its fill path gets the second one wrong, and
 * the second one is where `a:tile` and `a:srcRect` also live.
 *
 * ## The third slide is a security fixture
 *
 * An SVG is an XML document the package hands to whatever draws it, and
 * `<img src=blobURL>` is the only safe way to draw one - inline `<svg>` in the
 * page runs its script with the application's origin. `hostile.svg` here
 * carries, deliberately and inertly:
 *
 * - a `<script>` element,
 * - an `onload` attribute on the root,
 * - a `<use href="https://…">` pointing off-origin,
 * - an `<image href="https://…">` doing the same, and
 * - an `<a href="javascript:…">`.
 *
 * None of them does anything - the script body is a comment - and every one is
 * a thing sub-phase 10.7's sanitizer must remove. A corpus with only
 * well-behaved SVGs cannot tell whether the sanitizer runs at all.
 */

const SVG_EXT_URI = '{96DAC541-7B7A-43D3-8B79-37D633B846F1}';
const NS_ASVG = 'http://schemas.microsoft.com/office/drawing/2016/SVG/main';

const REL_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

/** The `a:extLst` that turns a blip into a vector one. */
const svgExt = (relId: string): string =>
  `<a:extLst><a:ext uri="${SVG_EXT_URI}">` +
  `<asvg:svgBlip xmlns:asvg="${NS_ASVG}" r:embed="${relId}"/>` +
  '</a:ext></a:extLst>';

/**
 * A `p:pic` whose blip may or may not carry a raster.
 *
 * The chassis's `picture()` always writes `@r:embed`, which is the one thing
 * this deck needs to be able to leave out, so the markup is written here.
 */
function svgPicture(spec: {
  readonly id: number;
  readonly name: string;
  readonly descr: string;
  readonly svgRelId: string;
  readonly rasterRelId?: string;
  readonly x: number;
  readonly y: number;
  readonly cx: number;
  readonly cy: number;
}): string {
  return (
    '<p:pic><p:nvPicPr>' +
    cNvPrXml({ id: spec.id, name: spec.name, descr: spec.descr }) +
    '<p:cNvPicPr><a:picLocks/></p:cNvPicPr><p:nvPr/>' +
    '</p:nvPicPr>' +
    '<p:blipFill>' +
    `<a:blip${spec.rasterRelId === undefined ? '' : ` r:embed="${spec.rasterRelId}"`}>` +
    svgExt(spec.svgRelId) +
    '</a:blip>' +
    '<a:stretch><a:fillRect/></a:stretch>' +
    '</p:blipFill>' +
    '<p:spPr>' +
    `<a:xfrm><a:off x="${String(spec.x)}" y="${String(spec.y)}"/>` +
    `<a:ext cx="${String(spec.cx)}" cy="${String(spec.cy)}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
    '</p:spPr></p:pic>'
  );
}

/** `a:blipFill` as a shape fill - the DrawingML one the census counts. */
const svgShapeFill = (svgRelId: string, rasterRelId: string, tile: boolean): string =>
  '<a:blipFill rotWithShape="1">' +
  `<a:blip r:embed="${rasterRelId}">${svgExt(svgRelId)}</a:blip>` +
  (tile
    ? '<a:tile tx="0" ty="0" sx="50000" sy="50000" flip="none" algn="tl"/>'
    : '<a:stretch><a:fillRect/></a:stretch>') +
  '</a:blipFill>';

// --------------------------------------------------------------- the assets

/** A shape whose orientation is obvious, so a flipped or cropped draw shows. */
const PROBE_SVG = `<svg xmlns="http://www.w3.org/2000/svg" width="240" height="160" viewBox="0 0 240 160">
<title>PPTX Studio corpus probe</title>
<rect x="0" y="0" width="240" height="160" fill="#F2F2F2"/>
<circle cx="70" cy="80" r="50" fill="#C00000"/>
<rect x="130" y="40" width="80" height="80" fill="#0070C0"/>
<path d="M 10 150 L 230 150" stroke="#7030A0" stroke-width="6" fill="none"/>
<text x="12" y="28" font-family="sans-serif" font-size="18" fill="#7030A0">top left</text>
</svg>
`;

/**
 * Every construct sub-phase 10.7's sanitizer has to remove, all inert.
 *
 * The script body is a comment, the handler assigns to a local, and both remote
 * URLs point at `example.invalid`, a name reserved by RFC 2606 that cannot
 * resolve. Nothing here does anything if it runs; the point is that it must
 * not be given the chance.
 */
const HOSTILE_SVG = `<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" width="240" height="160" viewBox="0 0 240 160" onload="void 0">
<title>PPTX Studio corpus: everything a sanitizer must strip</title>
<script type="application/ecmascript">/* corpus probe: a renderer must remove this element */</script>
<rect x="0" y="0" width="240" height="160" fill="#FFF2CC"/>
<use href="https://example.invalid/off-origin.svg#icon" x="10" y="10"/>
<image href="https://example.invalid/off-origin.png" x="60" y="20" width="60" height="60"/>
<a href="javascript:void 0"><rect x="140" y="20" width="80" height="60" fill="#C55A11"/></a>
<text x="12" y="150" font-family="sans-serif" font-size="14" fill="#833C00">sanitize me</text>
</svg>
`;

const PARTS: readonly ProbePart[] = [
  {
    name: 'ppt/media/image1.png',
    bytes: quadrantPng(),
    contentType: { kind: 'default', extension: 'png', type: 'image/png' },
  },
  {
    // The extension is what types an SVG, and PowerPoint writes it as a
    // `Default` rather than an `Override` - measured, same as png and jpeg.
    name: 'ppt/media/image2.svg',
    bytes: PROBE_SVG,
    contentType: { kind: 'default', extension: 'svg', type: 'image/svg+xml' },
  },
  { name: 'ppt/media/image3.svg', bytes: HOSTILE_SVG },
];

const SVG_REL = { id: 'rId2', type: REL_IMAGE, target: '../media/image2.svg' };
const PNG_REL = { id: 'rId3', type: REL_IMAGE, target: '../media/image1.png' };
const HOSTILE_REL = { id: 'rId4', type: REL_IMAGE, target: '../media/image3.svg' };

const caption = (
  id: number,
  name: string,
  x: number,
  y: number,
  cx: number,
  lines: readonly string[],
): string =>
  shape({
    id,
    name,
    x,
    y,
    cx,
    cy: 1143000,
    textBody: txBody({
      bodyPr: '<a:bodyPr wrap="square" anchor="t"/>',
      paras: lines.map((line) => textLine(line, { sz: 1200 })).join(''),
    }),
  });

export const a25SvgBlips: ProbeDeck = {
  id: 'a25-svg-blips',
  title: 'PPTX Studio corpus: a25 svg blips',
  description:
    'asvg:svgBlip in the three places it occurs. Slide 1 has the form PowerPoint 16.0.20326 ' +
    'actually wrote - an a:blip with no @r:embed at all, whose only reference is the SVG inside ' +
    'the extension - beside the documented form that carries a PNG fallback as well. Slide 2 puts ' +
    'the same extension inside a:blipFill used as a shape fill, stretched and tiled, which is a ' +
    'different element from p:pic’s p:blipFill. Slide 3 references an SVG carrying a script ' +
    'element, an onload attribute, off-origin use and image references and a javascript: link, ' +
    'all inert, so sub-phase 10.7’s sanitizer has something to be tested against.',
  features: {
    // 6 chassis + one caption a slide + the two filled shapes on slide 2.
    shape: 11,
    placeholder: 6,
    picture: 3,
    // Every picture and every shape here is a plain rectangle.
    presetGeom: 8,
    gradientFill: 2,
    blipFill: 2,
    svgBlip: 5,
  },
  build: () => ({
    title: 'PPTX Studio corpus: a25 svg blips',
    parts: PARTS,
    slides: [
      {
        title: 'a25 — a blip with no raster, and a blip with one',
        rels: [SVG_REL, PNG_REL],
        body:
          svgPicture({
            id: 10,
            name: 'Vector only, as PowerPoint wrote it',
            descr: 'a:blip with no r:embed; the svgBlip extension is the only reference',
            svgRelId: 'rId2',
            x: 685800,
            y: 1600200,
            cx: 3657600,
            cy: 2438400,
          }) +
          svgPicture({
            id: 11,
            name: 'Vector with a raster fallback',
            descr: 'a:blip r:embed points at a PNG; the extension carries the vector original',
            svgRelId: 'rId2',
            rasterRelId: 'rId3',
            x: 5029200,
            y: 1600200,
            cx: 3657600,
            cy: 2438400,
          }) +
          caption(12, 'Slide 1 caption', 685800, 4343400, 8001000, [
            'Left: no @r:embed. A reader that requires one renders nothing.',
            'Right: @r:embed is the PNG; the extension is the vector to prefer.',
            'Both are legal, and only one of them was written by PowerPoint.',
          ]),
      },
      {
        title: 'a25 — the same extension inside a shape fill',
        rels: [SVG_REL, PNG_REL],
        body:
          shape({
            id: 10,
            name: 'a:blipFill, stretched',
            x: 685800,
            y: 1600200,
            cx: 3657600,
            cy: 2438400,
            fill: svgShapeFill('rId2', 'rId3', false),
          }) +
          shape({
            id: 11,
            name: 'a:blipFill, tiled at 50%',
            x: 5029200,
            y: 1600200,
            cx: 3657600,
            cy: 2438400,
            fill: svgShapeFill('rId2', 'rId3', true),
          }) +
          caption(12, 'Slide 2 caption', 685800, 4343400, 8001000, [
            'a:blipFill is DrawingML and p:blipFill is PresentationML.',
            'They are different elements, and only this one tiles and crops.',
          ]),
      },
      {
        title: 'a25 — an SVG a renderer must not hand to the page',
        rels: [HOSTILE_REL],
        body:
          svgPicture({
            id: 10,
            name: 'Hostile SVG',
            descr: 'script, onload, off-origin use and image, and a javascript: link',
            svgRelId: 'rId4',
            x: 685800,
            y: 1600200,
            cx: 3657600,
            cy: 2438400,
          }) +
          caption(11, 'Slide 3 caption', 5029200, 1600200, 6096000, [
            'The SVG behind this picture carries five things a sanitizer removes:',
            'a script element, an onload attribute, a use and an image pointing',
            'off-origin, and an anchor with a javascript: href. All inert - the',
            'script body is a comment and both URLs are under example.invalid,',
            'which RFC 2606 reserves so that it cannot resolve.',
          ]),
      },
    ],
  }),
};
