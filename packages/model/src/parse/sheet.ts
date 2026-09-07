/**
 * A slide, layout or master part into a `Sheet`; a theme part into a `Theme`.
 *
 * Nothing here resolves anything and nothing here defaults anything that the
 * file could have said. The one exception is `showMasterShapes`, whose default
 * is `true` and which has no third state to preserve.
 */

import {
  MAPPED_COLOR_NAMES,
  SCHEME_SLOTS,
  type ClrMap,
  type ClrScheme,
  type Color,
  type Effect,
  type Fill,
  type Line,
  type MappedColorName,
  type SchemeSlot,
} from '@pptx-studio/paint';
import { attributeValue, childElements, firstChild, type XElement } from '@pptx-studio/xml';

import { parseGeometry } from './geometry.js';

import { ModelError } from '../errors.js';
import {
  parseColorElement,
  parseEffects,
  parseFill,
  parseFillElement,
  parsePictureFill,
  parseLine,
  parseLineElement,
} from './paint.js';
import { parseTextBodyChild, parseTextStyles } from './text.js';
import {
  PLACEHOLDER_TYPES,
  type Background,
  type ColorMapOverride,
  type FontCollection,
  type FontScheme,
  type FormatScheme,
  type Placeholder,
  type PlaceholderSize,
  type PlaceholderType,
  type Shape,
  type ShapeKind,
  type ShapeStyle,
  type Sheet,
  type SheetKind,
  type StyleRef,
  type Theme,
  type Xfrm,
} from '../types.js';

/* -------------------------------------------------------------------------- */
/* the colour map                                                             */
/* -------------------------------------------------------------------------- */

const SLOT_SET: ReadonlySet<string> = new Set<string>(SCHEME_SLOTS);

/**
 * `p:clrMap` or `a:overrideClrMapping`: all twelve attributes, every time.
 *
 * Eleven is a repair prompt, so a map that is missing one is a broken file
 * rather than a partial opinion, and there is nothing sensible to default the
 * missing one to.
 */
export function parseClrMap(element: XElement, partName: string): ClrMap {
  const out: Partial<Record<MappedColorName, SchemeSlot>> = {};
  for (const name of MAPPED_COLOR_NAMES) {
    const raw = attributeValue(element, name);
    if (raw === undefined) {
      throw new ModelError(
        'MODEL_CLRMAP_INCOMPLETE',
        `${element.qname} has no @${name}`,
        partName,
        name,
      );
    }
    if (!SLOT_SET.has(raw)) {
      throw new ModelError(
        'MODEL_CLRMAP_SLOT',
        `${element.qname}/@${name} is "${raw}", which is not a clrScheme slot`,
        partName,
        raw,
      );
    }
    out[name] = raw as SchemeSlot;
  }
  return out as ClrMap;
}

function parseClrMapOvr(parent: XElement, partName: string): ColorMapOverride | undefined {
  const element = firstChild(parent, 'p:clrMapOvr');
  if (element === undefined) return undefined;
  const override = firstChild(element, 'a:overrideClrMapping');
  if (override !== undefined) return { kind: 'override', map: parseClrMap(override, partName) };
  return { kind: 'inherit' };
}

/* -------------------------------------------------------------------------- */
/* geometry                                                                   */
/* -------------------------------------------------------------------------- */

function intOf(element: XElement, name: string, partName: string): number {
  const raw = attributeValue(element, name);
  if (raw === undefined) return 0;
  const value = Number(raw);
  if (!Number.isInteger(value)) {
    throw new ModelError(
      'MODEL_XFRM_NUMBER',
      `${element.qname}/@${name} is "${raw}"`,
      partName,
      raw,
    );
  }
  return value;
}

/**
 * The transform, from whichever element holds it.
 *
 * A `p:graphicFrame` keeps its geometry in a **PresentationML** `p:xfrm` and
 * has no `p:spPr` at all - it is a different element in a different namespace
 * from the `a:xfrm` every other shape uses, and looking for the wrong one gives
 * every chart, table, diagram and OLE object a size of zero.
 */
function parseXfrm(parent: XElement, partName: string, qname = 'a:xfrm'): Xfrm | undefined {
  const element = firstChild(parent, qname);
  if (element === undefined) return undefined;
  const off = firstChild(element, 'a:off');
  const ext = firstChild(element, 'a:ext');
  const chOff = firstChild(element, 'a:chOff');
  const chExt = firstChild(element, 'a:chExt');
  return {
    x: off === undefined ? 0 : intOf(off, 'x', partName),
    y: off === undefined ? 0 : intOf(off, 'y', partName),
    cx: ext === undefined ? 0 : intOf(ext, 'cx', partName),
    cy: ext === undefined ? 0 : intOf(ext, 'cy', partName),
    rot: intOf(element, 'rot', partName),
    flipH: attributeValue(element, 'flipH') === '1',
    flipV: attributeValue(element, 'flipV') === '1',
    child:
      chOff === undefined && chExt === undefined
        ? null
        : {
            x: chOff === undefined ? 0 : intOf(chOff, 'x', partName),
            y: chOff === undefined ? 0 : intOf(chOff, 'y', partName),
            cx: chExt === undefined ? 0 : intOf(chExt, 'cx', partName),
            cy: chExt === undefined ? 0 : intOf(chExt, 'cy', partName),
          },
  };
}

/* -------------------------------------------------------------------------- */
/* placeholders and styles                                                    */
/* -------------------------------------------------------------------------- */

const TYPE_SET: ReadonlySet<string> = new Set<string>(PLACEHOLDER_TYPES);
const SIZE_SET: ReadonlySet<string> = new Set<string>(['full', 'half', 'quarter']);

/** `2^32 - 1`, the top of `xsd:unsignedInt`, which is what `@idx` is. */
const MAX_IDX = 4294967295;

function parsePlaceholder(nvPr: XElement | undefined, partName: string): Placeholder | null {
  if (nvPr === undefined) return null;
  const ph = firstChild(nvPr, 'p:ph');
  if (ph === undefined) return null;

  const rawType = attributeValue(ph, 'type');
  const rawIdx = attributeValue(ph, 'idx');
  let idx: number | null = null;
  if (rawIdx !== undefined) {
    const value = Number(rawIdx);
    if (!Number.isInteger(value) || value < 0 || value > MAX_IDX) {
      throw new ModelError(
        'MODEL_PLACEHOLDER_IDX',
        `p:ph/@idx is "${rawIdx}", which is not an unsignedInt`,
        partName,
        rawIdx,
      );
    }
    idx = value;
  }

  const rawSize = attributeValue(ph, 'sz');
  const orient = attributeValue(ph, 'orient');
  return {
    // An unknown `@type` is recorded as written rather than dropped, so a
    // round trip keeps it. PowerPoint repairs the file and normalises it to
    // `obj`; the matcher below treats anything it does not know the same way.
    type: rawType !== undefined && TYPE_SET.has(rawType) ? (rawType as PlaceholderType) : null,
    idx,
    size: rawSize !== undefined && SIZE_SET.has(rawSize) ? (rawSize as PlaceholderSize) : null,
    orient: orient === 'vert' ? 'vert' : orient === 'horz' ? 'horz' : null,
    hasCustomPrompt: attributeValue(ph, 'hasCustomPrompt') === '1',
  };
}

function parseStyleRef(element: XElement | undefined, partName: string): StyleRef {
  if (element === undefined) return { idx: 0, color: null };
  const raw = attributeValue(element, 'idx');
  const idx = raw === undefined ? 0 : Number(raw);
  if (!Number.isInteger(idx) || idx < 0) {
    throw new ModelError(
      'MODEL_STYLE_IDX',
      `${element.qname}/@idx is "${String(raw)}"`,
      partName,
      raw ?? '',
    );
  }
  let color: Color | null = null;
  for (const child of childElements(element)) {
    if (child.qname.startsWith('a:') && child.qname.endsWith('Clr')) {
      color = parseColorElement(child);
      break;
    }
  }
  return { idx, color };
}

function parseStyle(parent: XElement, partName: string): ShapeStyle | undefined {
  const element = firstChild(parent, 'p:style');
  if (element === undefined) return undefined;

  const fontRefElement = firstChild(element, 'a:fontRef');
  const rawFont = fontRefElement === undefined ? undefined : attributeValue(fontRefElement, 'idx');
  if (rawFont !== undefined && rawFont !== 'major' && rawFont !== 'minor' && rawFont !== 'none') {
    // `a:fontRef/@idx` is `ST_FontCollectionIndex`, a name, and never a number.
    // PowerPoint repairs a numeric one; a parser that reads it as an integer
    // gets `NaN` and silently loses the shape's font.
    throw new ModelError(
      'MODEL_FONT_COLLECTION',
      `a:fontRef/@idx is "${rawFont}", not major, minor or none`,
      partName,
      rawFont,
    );
  }
  let fontColor: Color | null = null;
  if (fontRefElement !== undefined) {
    for (const child of childElements(fontRefElement)) {
      if (child.qname.startsWith('a:') && child.qname.endsWith('Clr')) {
        fontColor = parseColorElement(child);
        break;
      }
    }
  }

  return {
    lnRef: parseStyleRef(firstChild(element, 'a:lnRef'), partName),
    fillRef: parseStyleRef(firstChild(element, 'a:fillRef'), partName),
    effectRef: parseStyleRef(firstChild(element, 'a:effectRef'), partName),
    fontRef: { idx: rawFont ?? 'minor', color: fontColor },
  };
}

/* -------------------------------------------------------------------------- */
/* shapes                                                                     */
/* -------------------------------------------------------------------------- */

const SHAPE_KINDS: Readonly<Record<string, ShapeKind>> = {
  'p:sp': 'sp',
  'p:pic': 'pic',
  'p:grpSp': 'grpSp',
  'p:graphicFrame': 'graphicFrame',
  'p:cxnSp': 'cxnSp',
  'p:contentPart': 'contentPart',
};

/** Where each shape kind keeps its non-visual properties. */
const NV_CONTAINERS: Readonly<Record<ShapeKind, string>> = {
  sp: 'p:nvSpPr',
  pic: 'p:nvPicPr',
  grpSp: 'p:nvGrpSpPr',
  graphicFrame: 'p:nvGraphicFramePr',
  cxnSp: 'p:nvCxnSpPr',
  contentPart: 'p:nvContentPartPr',
};

function parseShape(element: XElement, kind: ShapeKind, partName: string): Shape {
  const nv = firstChild(element, NV_CONTAINERS[kind]);
  const cNvPr = nv === undefined ? undefined : firstChild(nv, 'p:cNvPr');

  // `p:graphicFrame` keeps its geometry in a **PresentationML** `p:xfrm`, not
  // in an `a:xfrm` inside `p:spPr`, and has no `p:spPr` at all. Reading it the
  // other way gives every chart, table and diagram a size of zero.
  const spPr = firstChild(element, kind === 'grpSp' ? 'p:grpSpPr' : 'p:spPr');
  const geometryHost = kind === 'graphicFrame' ? element : spPr;

  const prstGeom = spPr === undefined ? undefined : firstChild(spPr, 'a:prstGeom');

  const children: Shape[] = [];
  if (kind === 'grpSp') {
    for (const child of childElements(element)) {
      const childKind = SHAPE_KINDS[child.qname];
      if (childKind !== undefined) children.push(parseShape(child, childKind, partName));
    }
  }

  // A `p:pic` draws the image in its `p:blipFill`, which sits beside `p:spPr`
  // rather than in it, so reading only `spPr` leaves every picture blank.
  const picture = kind === 'pic' ? parsePictureFill(element, partName) : undefined;
  const rawId = cNvPr === undefined ? undefined : attributeValue(cNvPr, 'id');
  return {
    kind,
    cNvPrId: rawId === undefined ? 0 : Number(rawId),
    name: (cNvPr === undefined ? undefined : attributeValue(cNvPr, 'name')) ?? '',
    descr: (cNvPr === undefined ? undefined : attributeValue(cNvPr, 'descr')) ?? null,
    hidden: cNvPr !== undefined && attributeValue(cNvPr, 'hidden') === '1',
    placeholder: parsePlaceholder(
      nv === undefined ? undefined : firstChild(nv, 'p:nvPr'),
      partName,
    ),
    xfrm:
      geometryHost === undefined
        ? undefined
        : parseXfrm(geometryHost, partName, kind === 'graphicFrame' ? 'p:xfrm' : 'a:xfrm'),
    fill: picture ?? (spPr === undefined ? undefined : parseFill(spPr, partName)),
    line: spPr === undefined ? undefined : parseLine(spPr, partName),
    effects: spPr === undefined ? undefined : parseEffects(spPr),
    style: parseStyle(element, partName),
    prstGeom: prstGeom === undefined ? undefined : (attributeValue(prstGeom, 'prst') ?? ''),
    geometry: parseGeometry(spPr, partName),
    text: parseTextBodyChild(element, partName),
    children,
    node: element,
  };
}

function parseShapeTree(cSld: XElement, partName: string): Shape[] {
  const tree = firstChild(cSld, 'p:spTree');
  if (tree === undefined) return [];
  const out: Shape[] = [];
  for (const child of childElements(tree)) {
    const kind = SHAPE_KINDS[child.qname];
    if (kind !== undefined) out.push(parseShape(child, kind, partName));
  }
  return out;
}

/* -------------------------------------------------------------------------- */
/* the background                                                             */
/* -------------------------------------------------------------------------- */

function parseBackground(cSld: XElement, partName: string): Background | undefined {
  const bg = firstChild(cSld, 'p:bg');
  if (bg === undefined) return undefined;

  const ref = firstChild(bg, 'p:bgRef');
  if (ref !== undefined) return { kind: 'ref', ref: parseStyleRef(ref, partName) };

  const pr = firstChild(bg, 'p:bgPr');
  if (pr === undefined) return undefined;
  const fill = parseFill(pr, partName);
  // `p:bgPr` requires a fill. An `a:noFill` inside one is honoured rather than
  // treated as absent: measured in 2.9, a slide declaring it paints nothing and
  // does not fall through to its layout.
  return { kind: 'fill', fill: fill ?? { type: 'none' }, effects: parseEffects(pr) };
}

/* -------------------------------------------------------------------------- */
/* the theme                                                                  */
/* -------------------------------------------------------------------------- */

function parseScheme(element: XElement, partName: string): ClrScheme {
  const out: Partial<Record<SchemeSlot, Color>> = {};
  for (const slot of SCHEME_SLOTS) {
    const child = firstChild(element, `a:${slot}`);
    if (child === undefined) continue;
    for (const grandchild of childElements(child)) {
      if (grandchild.qname.startsWith('a:') && grandchild.qname.endsWith('Clr')) {
        out[slot] = parseColorElement(grandchild);
        break;
      }
    }
  }
  const missing = SCHEME_SLOTS.filter((slot) => out[slot] === undefined);
  if (missing.length > 0) {
    throw new ModelError(
      'MODEL_CLRMAP_INCOMPLETE',
      `a:clrScheme is missing ${missing.join(', ')}`,
      partName,
      missing.join(','),
    );
  }
  return out as ClrScheme;
}

function parseFontScheme(element: XElement | undefined): FontScheme {
  // All three scripts, because `a:rPr` has four typeface slots and three of them
  // take a theme reference: `+mn-ea` on an `a:ea` is how a CJK run follows the
  // theme, and reading only `a:latin` leaves it unresolvable.
  const collection = (which: string): FontCollection => {
    const font = element === undefined ? undefined : firstChild(element, which);
    const face = (script: string): string | null => {
      if (font === undefined) return null;
      const child = firstChild(font, script);
      return child === undefined ? null : (attributeValue(child, 'typeface') ?? null);
    };
    return { latin: face('a:latin'), ea: face('a:ea'), cs: face('a:cs') };
  };
  return {
    name: element === undefined ? null : (attributeValue(element, 'name') ?? null),
    major: collection('a:majorFont'),
    minor: collection('a:minorFont'),
  };
}

function parseFormatScheme(element: XElement | undefined, partName: string): FormatScheme | null {
  if (element === undefined) return null;
  const fills = (name: string): Fill[] => {
    const list = firstChild(element, name);
    if (list === undefined) return [];
    return childElements(list).map((child) => parseFillElement(child, partName));
  };
  const lineList = firstChild(element, 'a:lnStyleLst');
  const lines: Line[] =
    lineList === undefined
      ? []
      : childElements(lineList)
          .filter((child) => child.qname === 'a:ln')
          .map((child) => parseLineElement(child, partName));
  const effectList = firstChild(element, 'a:effectStyleLst');
  const effects: (readonly Effect[])[] =
    effectList === undefined
      ? []
      : childElements(effectList)
          .filter((child) => child.qname === 'a:effectStyle')
          .map((child) => parseEffects(child) ?? []);

  return {
    name: attributeValue(element, 'name') ?? null,
    fillStyles: fills('a:fillStyleLst'),
    lineStyles: lines,
    effectStyles: effects,
    bgFillStyles: fills('a:bgFillStyleLst'),
  };
}

export function parseTheme(root: XElement, partName: string): Theme {
  const elements = firstChild(root, 'a:themeElements');
  const scheme = elements === undefined ? undefined : firstChild(elements, 'a:clrScheme');
  if (elements === undefined || scheme === undefined) {
    throw new ModelError('MODEL_PART_KIND', 'a theme with no a:clrScheme', partName);
  }
  return {
    partName,
    name: attributeValue(root, 'name') ?? null,
    scheme: parseScheme(scheme, partName),
    fonts: parseFontScheme(firstChild(elements, 'a:fontScheme')),
    format: parseFormatScheme(firstChild(elements, 'a:fmtScheme'), partName),
    node: root,
  };
}

/* -------------------------------------------------------------------------- */
/* the sheet                                                                  */
/* -------------------------------------------------------------------------- */

const ROOTS: Readonly<Record<string, SheetKind>> = {
  'p:sld': 'slide',
  'p:sldLayout': 'layout',
  'p:sldMaster': 'master',
};

/**
 * One part into one sheet. The parent and the theme are filled in by
 * `document.ts`, which is the only thing that can follow a relationship.
 */
export function parseSheet(root: XElement, partName: string): Omit<Sheet, 'parent' | 'theme'> {
  const kind = ROOTS[root.qname];
  if (kind === undefined) {
    throw new ModelError('MODEL_PART_KIND', `${root.qname} is not a sheet root`, partName);
  }
  const cSld = firstChild(root, 'p:cSld');
  if (cSld === undefined) {
    throw new ModelError('MODEL_PART_KIND', `${root.qname} has no p:cSld`, partName);
  }
  const clrMapElement = kind === 'master' ? firstChild(root, 'p:clrMap') : undefined;
  // Only a master carries `p:txStyles`, and whether it carries one at all is a
  // load-bearing distinction: PowerPoint substitutes its whole built-in set for
  // a master that declares none, measured in 3.1.
  const txStylesElement = kind === 'master' ? firstChild(root, 'p:txStyles') : undefined;

  return {
    kind,
    partName,
    name: attributeValue(cSld, 'name') ?? null,
    layoutType: kind === 'layout' ? (attributeValue(root, 'type') ?? null) : null,
    matchingName: kind === 'layout' ? (attributeValue(root, 'matchingName') ?? null) : null,
    shapes: parseShapeTree(cSld, partName),
    background: parseBackground(cSld, partName),
    clrMap: clrMapElement === undefined ? undefined : parseClrMap(clrMapElement, partName),
    txStyles:
      txStylesElement === undefined ? undefined : parseTextStyles(txStylesElement, partName),
    clrMapOvr: kind === 'master' ? undefined : parseClrMapOvr(root, partName),
    // `@showMasterSp` sits on `p:sld` and `p:sldLayout` themselves, not inside
    // `p:cSld` where the rest of this comes from.
    showMasterShapes: attributeValue(root, 'showMasterSp') !== '0',
    node: root,
  };
}
