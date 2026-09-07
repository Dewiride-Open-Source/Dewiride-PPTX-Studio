/**
 * Shape markup for the Tier A probes.
 *
 * Every builder here writes `a:spPr`'s children in the one order the schema
 * allows, taken from `packages/xml/src/edit/schema-order.gen.ts` rather than from
 * memory:
 *
 * ```
 * spPr = xfrm, custGeom|prstGeom, <fill>, ln, effectDag|effectLst, scene3d, sp3d, extLst
 * ```
 *
 * That table is generated from the ECMA-376 Transitional XSDs, so it is the
 * same source the runtime's `insertInOrder()` uses. Getting this wrong is not a
 * theoretical risk: Microsoft's own Open XML SDK shipped a regression that
 * swapped two elements in a `slideMaster1.xml` and PowerPoint refused the file.
 *
 * Every probe shape carries its own caption in its `p:txBody`. A separate
 * caption text box would read better but would double the shape count in every
 * `features` map, and those maps are the thing a reviewer reads.
 */

import { cNvPrXml, escapeXml, NS_A, NS_R, type DrawingProps } from './chassis.ts';

/** The content area under the title: everything below y, across the full width. */
const CONTENT = { x: 457200, y: 1188720, width: 11277600, height: 5181600 } as const;

export interface Cell {
  readonly x: number;
  readonly y: number;
  readonly cx: number;
  readonly cy: number;
}

/**
 * A grid over the content area.
 *
 * Probe decks are read by a person looking for the one shape that rendered
 * wrong, so the layout is a grid rather than the benchmark generator's
 * pseudo-random scatter. Nothing here affects the markup being probed.
 */
export function grid(columns: number, rows: number, inset = 45720): (index: number) => Cell {
  const cellWidth = Math.floor(CONTENT.width / columns);
  const cellHeight = Math.floor(CONTENT.height / rows);
  return (index: number): Cell => ({
    x: CONTENT.x + (index % columns) * cellWidth + inset,
    y: CONTENT.y + Math.floor(index / columns) * cellHeight + inset,
    cx: cellWidth - 2 * inset,
    cy: cellHeight - 2 * inset,
  });
}

export interface ShapeSpec extends DrawingProps {
  /** `@name` is also the caption, so a human opening the deck sees what failed. */
  readonly x: number;
  readonly y: number;
  readonly cx: number;
  readonly cy: number;
  /** `a:prstGeom` or `a:custGeom` markup. Defaults to a plain rectangle. */
  readonly geometry?: string;
  /** One of the six fill elements. Defaults to `a:noFill`. */
  readonly fill?: string;
  /** `a:ln` markup, or omitted for no line element at all. */
  readonly line?: string;
  /** `a:effectLst` or `a:effectDag` markup. */
  readonly effects?: string;
  readonly scene3d?: string;
  readonly sp3d?: string;
  /**
   * `a:extLst` markup, always last inside `a:spPr`.
   *
   * Named apart from `DrawingProps.extLst`, which is the *other* extension
   * list on the same shape - the one inside `p:cNvPr`, where
   * `adec:decorative` lives. A shape has two and they are not interchangeable.
   */
  readonly spPrExtLst?: string;
  readonly rotation?: number;
  readonly flipH?: boolean;
  readonly flipV?: boolean;
  /** Suppress the caption, for shapes whose whole point is an empty text body. */
  readonly caption?: false;
  /**
   * A whole `p:txBody`, replacing the caption.
   *
   * The text decks - `a07` through `a11` - are about what is *in* the text
   * body, so a one-line centred caption is not a probe, it is the thing being
   * probed. Everything else keeps the caption, which is why this is an opt-in
   * rather than the shape builder growing a mode.
   */
  readonly textBody?: string;
}

export function prstGeom(
  preset: string,
  adjustments: Readonly<Record<string, number>> = {},
): string {
  const entries = Object.entries(adjustments);
  const avLst =
    entries.length === 0
      ? '<a:avLst/>'
      : '<a:avLst>' +
        entries
          .map(([name, value]) => `<a:gd name="${name}" fmla="val ${String(value)}"/>`)
          .join('') +
        '</a:avLst>';
  return `<a:prstGeom prst="${preset}">${avLst}</a:prstGeom>`;
}

function xfrmXml(spec: ShapeSpec): string {
  const attributes =
    (spec.rotation === undefined ? '' : ` rot="${String(spec.rotation)}"`) +
    (spec.flipH === true ? ' flipH="1"' : '') +
    (spec.flipV === true ? ' flipV="1"' : '');
  return (
    `<a:xfrm${attributes}>` +
    `<a:off x="${String(spec.x)}" y="${String(spec.y)}"/>` +
    `<a:ext cx="${String(spec.cx)}" cy="${String(spec.cy)}"/>` +
    '</a:xfrm>'
  );
}

/** `a:spPr` with its children in schema order. Shared by shapes and connectors. */
function spPrXml(spec: ShapeSpec): string {
  return (
    '<p:spPr>' +
    xfrmXml(spec) +
    (spec.geometry ?? prstGeom('rect')) +
    (spec.fill ?? '<a:noFill/>') +
    (spec.line ?? '') +
    (spec.effects ?? '') +
    (spec.scene3d ?? '') +
    (spec.sp3d ?? '') +
    (spec.spPrExtLst ?? '') +
    '</p:spPr>'
  );
}

function captionXml(spec: ShapeSpec): string {
  if (spec.textBody !== undefined) return spec.textBody;
  if (spec.caption === false) {
    return '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-GB"/></a:p></p:txBody>';
  }
  return (
    '<p:txBody>' +
    '<a:bodyPr wrap="square" lIns="18288" tIns="9144" rIns="18288" bIns="9144" anchor="b"/>' +
    '<a:lstStyle/>' +
    '<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-GB" sz="700" dirty="0"/>' +
    `<a:t>${escapeXml(spec.name)}</a:t></a:r></a:p>` +
    '</p:txBody>'
  );
}

/** `p:sp`. `CT_Shape` is nvSpPr, spPr, style, txBody, extLst - txBody is last but one. */
export function shape(spec: ShapeSpec): string {
  return (
    '<p:sp><p:nvSpPr>' +
    cNvPrXml(spec) +
    '<p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
    spPrXml(spec) +
    captionXml(spec) +
    '</p:sp>'
  );
}

export interface ConnectorSpec extends ShapeSpec {
  /** `a:stCxn` - the shape id and connection-site index this end is glued to. */
  readonly start?: { readonly id: number; readonly idx: number };
  readonly end?: { readonly id: number; readonly idx: number };
}

/**
 * `p:cxnSp`.
 *
 * A connector is not a shape with a line preset: `CT_Connector` has its own
 * `nvCxnSpPr`, and `a:stCxn`/`a:endCxn` inside `p:cNvCxnSpPr` are what make it
 * follow the shapes it is glued to. It also has no `p:txBody` at all, so there
 * is nowhere to put a caption.
 */
export function connector(spec: ConnectorSpec): string {
  const site = (tag: 'stCxn' | 'endCxn', to: { id: number; idx: number } | undefined): string =>
    to === undefined ? '' : `<a:${tag} id="${String(to.id)}" idx="${String(to.idx)}"/>`;
  return (
    '<p:cxnSp><p:nvCxnSpPr>' +
    cNvPrXml(spec) +
    '<p:cNvCxnSpPr>' +
    '<a:cxnSpLocks/>' +
    site('stCxn', spec.start) +
    site('endCxn', spec.end) +
    '</p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr>' +
    spPrXml(spec) +
    '</p:cxnSp>'
  );
}

export interface GroupSpec extends DrawingProps {
  readonly x: number;
  readonly y: number;
  readonly cx: number;
  readonly cy: number;
  /** Child coordinate space. Omitted means it equals the group's own extent. */
  readonly childOffsetX?: number;
  readonly childOffsetY?: number;
  readonly childWidth?: number;
  readonly childHeight?: number;
  /** `a:grpSpPr`'s fill. The one place `a:grpFill` on a child means anything. */
  readonly fill?: string;
  /** Sixtieths of a degree on the group's own `a:xfrm`, applied after the flip. */
  readonly rotation?: number;
  readonly flipH?: boolean;
  readonly flipV?: boolean;
  readonly children: string;
}

/**
 * `p:grpSp`.
 *
 * `a:chOff`/`a:chExt` are the child coordinate space, and they are the reason a
 * resized group is the most common group bug: a child at `x` renders at
 * `off.x + (x - chOff.x) * (ext.cx / chExt.cx)`. Writing them explicitly here,
 * even when they equal the group's own extent, keeps every probe deck honest
 * about which space its numbers are in.
 */
export function group(spec: GroupSpec): string {
  const chOffX = spec.childOffsetX ?? spec.x;
  const chOffY = spec.childOffsetY ?? spec.y;
  const chExtX = spec.childWidth ?? spec.cx;
  const chExtY = spec.childHeight ?? spec.cy;
  const transform =
    (spec.rotation === undefined ? '' : ` rot="${String(spec.rotation)}"`) +
    (spec.flipH === true ? ' flipH="1"' : '') +
    (spec.flipV === true ? ' flipV="1"' : '');
  return (
    '<p:grpSp><p:nvGrpSpPr>' +
    cNvPrXml(spec) +
    '<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
    '<p:grpSpPr>' +
    `<a:xfrm${transform}>` +
    `<a:off x="${String(spec.x)}" y="${String(spec.y)}"/>` +
    `<a:ext cx="${String(spec.cx)}" cy="${String(spec.cy)}"/>` +
    `<a:chOff x="${String(chOffX)}" y="${String(chOffY)}"/>` +
    `<a:chExt cx="${String(chExtX)}" cy="${String(chExtY)}"/>` +
    '</a:xfrm>' +
    (spec.fill ?? '') +
    '</p:grpSpPr>' +
    spec.children +
    '</p:grpSp>'
  );
}

/** `p:pic`. The PresentationML `p:blipFill`, which is *not* the census's `blipFill`. */
export function picture(
  spec: DrawingProps & {
    readonly relId: string;
    readonly x: number;
    readonly y: number;
    readonly cx: number;
    readonly cy: number;
    /** Written as `@descr` whether or not it is empty, which is what PowerPoint does. */
    readonly description: string;
  },
): string {
  return (
    '<p:pic><p:nvPicPr>' +
    cNvPrXml({ ...spec, descr: spec.description }) +
    '<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/>' +
    '</p:nvPicPr>' +
    `<p:blipFill><a:blip r:embed="${spec.relId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
    '<p:spPr>' +
    `<a:xfrm><a:off x="${String(spec.x)}" y="${String(spec.y)}"/><a:ext cx="${String(spec.cx)}" cy="${String(spec.cy)}"/></a:xfrm>` +
    prstGeom('rect') +
    '</p:spPr></p:pic>'
  );
}

/**
 * `p:graphicFrame` - the host every piece of hard content arrives in.
 *
 * Four things live behind this one element and nothing in the frame says which
 * until you read `a:graphicData/@uri`: a table (inline, in the slide part), a
 * chart, a ChartEx, a SmartArt diagram or an OLE object (all four in parts of
 * their own, reached by relationship). A renderer that switches on the frame
 * gets nowhere; it has to switch on the URI.
 *
 * Two traps, both of them in the schema rather than in the content:
 *
 * - The transform is **`p:xfrm`**, in the PresentationML namespace, not the
 *   `a:xfrm` every other shape carries. It is required, not optional, and it
 *   is one of the must-not-break rules for that reason.
 * - **Office ignores `rot`, `flipH` and `flipV` here.** They are writable
 *   and PowerPoint will not honour them, which is why sub-phase 5.3 disables
 *   the rotate handle on frames rather than writing a rotation no user can see.
 */
export function graphicFrame(
  spec: DrawingProps & {
    readonly x: number;
    readonly y: number;
    readonly cx: number;
    readonly cy: number;
    /** `a:graphicData/@uri`. The only thing that says what this frame holds. */
    readonly uri: string;
    /** The `a:graphicData` children. */
    readonly content: string;
    /** Extra `p:nvPr` markup, e.g. the `p:ph` a chart placeholder carries. */
    readonly nvPr?: string;
    /** `a:graphicFrameLocks` attributes, e.g. `noGrp="1"`. */
    readonly locks?: string;
  },
): string {
  return (
    '<p:graphicFrame><p:nvGraphicFramePr>' +
    cNvPrXml(spec) +
    '<p:cNvGraphicFramePr' +
    (spec.locks === undefined
      ? '/>'
      : `><a:graphicFrameLocks ${spec.locks}/></p:cNvGraphicFramePr>`) +
    `<p:nvPr>${spec.nvPr ?? ''}</p:nvPr>` +
    '</p:nvGraphicFramePr>' +
    // p:xfrm, not a:xfrm. Same children, different namespace, and required.
    `<p:xfrm><a:off x="${String(spec.x)}" y="${String(spec.y)}"/><a:ext cx="${String(spec.cx)}" cy="${String(spec.cy)}"/></p:xfrm>` +
    `<a:graphic><a:graphicData uri="${spec.uri}">${spec.content}</a:graphicData></a:graphic>` +
    '</p:graphicFrame>'
  );
}

/** The five `a:graphicData/@uri` values a `p:graphicFrame` can carry. */
export const GRAPHIC_URI = {
  table: 'http://schemas.openxmlformats.org/drawingml/2006/table',
  chart: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
  chartEx: 'http://schemas.microsoft.com/office/drawing/2014/chartex',
  diagram: 'http://schemas.openxmlformats.org/drawingml/2006/diagram',
  ole: 'http://schemas.openxmlformats.org/presentationml/2006/ole',
} as const;

// ------------------------------------------------------------------- colours

/** `a:srgbClr`, optionally with transforms. Transforms apply in document order. */
export function srgb(hex: string, transforms = ''): string {
  return transforms === ''
    ? `<a:srgbClr val="${hex}"/>`
    : `<a:srgbClr val="${hex}">${transforms}</a:srgbClr>`;
}

export function scheme(name: string, transforms = ''): string {
  return transforms === ''
    ? `<a:schemeClr val="${name}"/>`
    : `<a:schemeClr val="${name}">${transforms}</a:schemeClr>`;
}

export const solidFill = (color: string): string => `<a:solidFill>${color}</a:solidFill>`;

/** `a:ln`. Children in schema order: fill, dash, join, headEnd, tailEnd, extLst. */
export function line(options: {
  readonly width?: number;
  readonly cap?: 'rnd' | 'sq' | 'flat';
  readonly compound?: 'sng' | 'dbl' | 'thickThin' | 'thinThick' | 'tri';
  readonly align?: 'ctr' | 'in';
  readonly fill?: string;
  readonly dash?: string;
  readonly join?: string;
  readonly headEnd?: string;
  readonly tailEnd?: string;
}): string {
  const attributes =
    (options.width === undefined ? '' : ` w="${String(options.width)}"`) +
    (options.cap === undefined ? '' : ` cap="${options.cap}"`) +
    (options.compound === undefined ? '' : ` cmpd="${options.compound}"`) +
    (options.align === undefined ? '' : ` algn="${options.align}"`);
  return (
    `<a:ln${attributes}>` +
    (options.fill ?? '') +
    (options.dash ?? '') +
    (options.join ?? '') +
    (options.headEnd ?? '') +
    (options.tailEnd ?? '') +
    '</a:ln>'
  );
}

/** A run-level hyperlink, for the decks that need one. */
export const hlinkClick = (relId: string): string =>
  `<a:hlinkClick xmlns:r="${NS_R}" r:id="${relId}"/>`;

export { NS_A };
