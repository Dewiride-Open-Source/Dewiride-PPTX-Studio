/**
 * DrawingML markup into the types `@pptx-studio/paint` defines.
 *
 * Paint depends on nothing but arithmetic - it has never seen an `XElement` and
 * must not, or it would inherit the XML layer's whole surface for the sake of
 * six element names. So the crossing lives here, at layer 2, which is the first
 * place that knows about both.
 *
 * Every function returns `undefined` when the element is absent, and `undefined`
 * is load-bearing: "this shape declares no fill" and "this shape declares
 * `a:noFill`" are different files with different resolutions, and only the first
 * one inherits.
 */

import {
  COLOR_TRANSFORM_OPS,
  PaintError,
  parseAngle,
  parsePercentage,
  transformShape,
  WHOLE_RECT,
  type Color,
  type ColorTransform,
  type ColorTransformOp,
  type DashSegment,
  type Effect,
  type BlipEffect,
  type BlipStretch,
  type BlipTile,
  type Fill,
  type GradientShade,
  type GradientStop,
  type Line,
  type LineCap,
  type LineDash,
  type LineEnd,
  type LineEndSize,
  type LineJoin,
  type LineJoinKind,
  type PenAlignment,
  type RectAlignment,
  type RelativeRect,
  type SchemeColorName,
  type ShadowGeometry,
  type TileAlign,
  type TileFlipMode,
} from '@pptx-studio/paint';
import { attributeValue, childElements, firstChild, type XElement } from '@pptx-studio/xml';

import { ModelError } from '../errors.js';

/* -------------------------------------------------------------------------- */
/* small readers                                                              */
/* -------------------------------------------------------------------------- */

const COLOR_ELEMENTS = new Set([
  'a:srgbClr',
  'a:scrgbClr',
  'a:hslClr',
  'a:schemeClr',
  'a:sysClr',
  'a:prstClr',
]);

const FILL_ELEMENTS = new Set([
  'a:noFill',
  'a:solidFill',
  'a:gradFill',
  'a:blipFill',
  'a:pattFill',
  'a:grpFill',
]);

/** An integer attribute, or `fallback` when absent. Throws on a non-integer. */
function intAttr(element: XElement, name: string, fallback: number, what: string): number {
  const raw = attributeValue(element, name);
  if (raw === undefined) return fallback;
  const value = Number(raw);
  if (!Number.isInteger(value)) {
    throw new ModelError('MODEL_XFRM_NUMBER', `${what}/@${name} is "${raw}"`, null, raw);
  }
  return value;
}

/**
 * `ST_Percentage`, which may legally be written `"40000"` or `"40%"`.
 *
 * The second spelling is Strict's and a Transitional part accepts it anyway -
 * measured in 2.6. `parseInt` on `"40%"` yields 40, which is 0.04%, and that is
 * how a line-spacing of 150% becomes a collapsed slide.
 */
function percentAttr(element: XElement, name: string, fallback: number): number {
  const raw = attributeValue(element, name);
  if (raw === undefined) return fallback;
  return parsePercentage(raw);
}

function boolAttr(element: XElement, name: string, fallback: boolean): boolean {
  const raw = attributeValue(element, name);
  if (raw === undefined) return fallback;
  return raw === '1' || raw === 'true';
}

/* -------------------------------------------------------------------------- */
/* colour                                                                     */
/* -------------------------------------------------------------------------- */

function parseTransforms(element: XElement): ColorTransform[] {
  const out: ColorTransform[] = [];
  for (const child of childElements(element)) {
    const local = child.qname.startsWith('a:') ? child.qname.slice(2) : child.qname;
    const shape = transformShape(local);
    if (shape === null) continue;
    const op = local as ColorTransformOp;
    const raw = attributeValue(child, 'val');
    if (shape === 'flag') {
      out.push({ op } as ColorTransform);
      continue;
    }
    if (raw === undefined) {
      throw new PaintError('COLOR_NUMBER', `a:${local} has no @val`, local);
    }
    const val = shape === 'angle' ? parseAngle(raw) : parsePercentage(raw);
    out.push({ op, val } as ColorTransform);
  }
  return out;
}

/** One of the six `EG_ColorChoice` members. */
export function parseColorElement(element: XElement): Color {
  const transforms = parseTransforms(element);
  switch (element.qname) {
    case 'a:srgbClr':
      return { space: 'srgb', hex: attributeValue(element, 'val') ?? '', transforms };
    case 'a:scrgbClr':
      return {
        space: 'scrgb',
        r: percentAttr(element, 'r', 0),
        g: percentAttr(element, 'g', 0),
        b: percentAttr(element, 'b', 0),
        transforms,
      };
    case 'a:hslClr':
      return {
        space: 'hsl',
        hue: Number(attributeValue(element, 'hue') ?? '0'),
        sat: percentAttr(element, 'sat', 0),
        lum: percentAttr(element, 'lum', 0),
        transforms,
      };
    case 'a:schemeClr':
      return {
        space: 'scheme',
        name: (attributeValue(element, 'val') ?? '') as SchemeColorName,
        transforms,
      };
    case 'a:sysClr':
      return {
        space: 'sys',
        name: attributeValue(element, 'val') ?? '',
        lastClr: attributeValue(element, 'lastClr') ?? null,
        transforms,
      };
    case 'a:prstClr':
      return { space: 'prst', name: attributeValue(element, 'val') ?? '', transforms };
    default:
      throw new PaintError(
        'COLOR_SCHEME_NAME',
        `${element.qname} is not a colour element`,
        element.qname,
      );
  }
}

/** The first colour child of a wrapper such as `a:solidFill` or `a:fillRef`. */
export function parseColorChild(parent: XElement | undefined): Color | null {
  if (parent === undefined) return null;
  for (const child of childElements(parent)) {
    if (COLOR_ELEMENTS.has(child.qname)) return parseColorElement(child);
  }
  return null;
}

/* -------------------------------------------------------------------------- */
/* fills                                                                      */
/* -------------------------------------------------------------------------- */

function parseRelativeRect(element: XElement | undefined): RelativeRect {
  if (element === undefined) return WHOLE_RECT;
  return {
    l: percentAttr(element, 'l', 0),
    t: percentAttr(element, 't', 0),
    r: percentAttr(element, 'r', 0),
    b: percentAttr(element, 'b', 0),
  };
}

function parseGradientShade(element: XElement): GradientShade | null {
  const lin = firstChild(element, 'a:lin');
  if (lin !== undefined) {
    return {
      kind: 'linear',
      ang: Number(attributeValue(lin, 'ang') ?? '0'),
      scaled: boolAttr(lin, 'scaled', false),
    };
  }
  const path = firstChild(element, 'a:path');
  if (path === undefined) return null;
  // An absent `@path` paints the box ramp, measured against 27 probes: ADR 0022.
  const raw = attributeValue(path, 'path') ?? 'rect';
  const kind = raw === 'circle' || raw === 'rect' ? raw : 'shape';
  const rect = firstChild(path, 'a:fillToRect');
  return {
    kind: 'path',
    path: kind,
    fillToRect: rect === undefined ? null : parseRelativeRect(rect),
  };
}

const TILE_ALIGNS = new Set<string>(['tl', 't', 'tr', 'l', 'ctr', 'r', 'bl', 'b', 'br']);
const TILE_FLIPS = new Set<string>(['none', 'x', 'y', 'xy']);

/** `a:blip`'s colour effects, in document order, which is the order they apply. */
function parseBlipEffects(blip: XElement): BlipEffect[] {
  const effects: BlipEffect[] = [];
  for (const child of childElements(blip)) {
    switch (child.qname) {
      case 'a:grayscl':
        effects.push({ kind: 'grayscale' });
        break;
      case 'a:biLevel':
        effects.push({ kind: 'biLevel', thresh: percentAttr(child, 'thresh', 0) });
        break;
      case 'a:lum':
        effects.push({
          kind: 'lum',
          bright: percentAttr(child, 'bright', 0),
          contrast: percentAttr(child, 'contrast', 0),
        });
        break;
      case 'a:duotone': {
        const colors = childElements(child)
          .filter((c) => COLOR_ELEMENTS.has(c.qname))
          .map(parseColorElement);
        const [from, to] = colors;
        if (from === undefined || to === undefined) {
          throw new ModelError('BLIP_DUOTONE', 'a:duotone needs two colours', 'a:duotone');
        }
        effects.push({ kind: 'duotone', from, to });
        break;
      }
      case 'a:alphaModFix':
        effects.push({ kind: 'alphaModFix', amt: percentAttr(child, 'amt', 100000) });
        break;
      case 'a:clrChange': {
        const from = parseColorChild(firstChild(child, 'a:clrFrom'));
        const to = parseColorChild(firstChild(child, 'a:clrTo'));
        if (from === null || to === null) {
          throw new ModelError('BLIP_CLR_CHANGE', 'a:clrChange needs both colours', 'a:clrChange');
        }
        effects.push({ kind: 'clrChange', from, to, useAlpha: boolAttr(child, 'useA', true) });
        break;
      }
      default:
        break;
    }
  }
  return effects;
}

/**
 * `a:blipFill`.
 *
 * The relationship id is carried, never resolved: a fill in a layout resolves
 * against the layout's rels and the same markup inherited onto a slide still
 * means the layout's image. Same contract as `a:buBlip`.
 */
function parseBlipFill(element: XElement, partName: string): Fill {
  const blip = firstChild(element, 'a:blip');
  const embed = blip === undefined ? undefined : attributeValue(blip, 'r:embed');
  if (blip === undefined || embed === undefined) {
    throw new ModelError(
      'BLIP_NO_EMBED',
      'a:blipFill has no a:blip/@r:embed to resolve',
      'a:blipFill',
    );
  }

  const tile = firstChild(element, 'a:tile');
  const stretch = firstChild(element, 'a:stretch');
  let mode: BlipStretch | BlipTile;
  if (tile !== undefined) {
    const algn = attributeValue(tile, 'algn') ?? 'tl';
    const flip = attributeValue(tile, 'flip') ?? 'none';
    if (!TILE_ALIGNS.has(algn)) {
      throw new ModelError('BLIP_TILE_ALIGN', `a:tile/@algn is "${algn}"`, 'a:tile');
    }
    if (!TILE_FLIPS.has(flip)) {
      throw new ModelError('BLIP_TILE_FLIP', `a:tile/@flip is "${flip}"`, 'a:tile');
    }
    mode = {
      kind: 'tile',
      tx: intAttr(tile, 'tx', 0, 'a:tile'),
      ty: intAttr(tile, 'ty', 0, 'a:tile'),
      sx: percentAttr(tile, 'sx', 100000),
      sy: percentAttr(tile, 'sy', 100000),
      flip: flip as TileFlipMode,
      algn: algn as TileAlign,
    };
  } else {
    // No `a:tile` and no `a:stretch` is legal and behaves as a bare stretch.
    mode = {
      kind: 'stretch',
      fillRect: parseRelativeRect(
        stretch === undefined ? undefined : firstChild(stretch, 'a:fillRect'),
      ),
    };
  }

  return {
    type: 'blip',
    embed,
    part: partName,
    srcRect: parseRelativeRect(firstChild(element, 'a:srcRect')),
    mode,
    effects: parseBlipEffects(blip),
    dpi: intAttr(element, 'dpi', 0, 'a:blipFill'),
    rotWithShape: boolAttr(element, 'rotWithShape', true),
  };
}

/** One member of `EG_FillProperties`. */
export function parseFillElement(element: XElement, partName: string): Fill {
  switch (element.qname) {
    case 'a:noFill':
      return { type: 'none' };
    case 'a:grpFill':
      return { type: 'group' };
    case 'a:blipFill':
      return parseBlipFill(element, partName);
    case 'a:solidFill': {
      const color = parseColorChild(element);
      if (color === null) {
        // A `solidFill` with no colour child is not something PowerPoint writes,
        // and black is the wrong guess: it is indistinguishable from a real
        // black and would paint over a themed shape silently.
        throw new PaintError('COLOR_SCHEME_NAME', 'a:solidFill has no colour child', 'a:solidFill');
      }
      return { type: 'solid', color };
    }
    case 'a:pattFill':
      return {
        type: 'pattern',
        prst: attributeValue(element, 'prst') ?? '',
        fg: parseColorChild(firstChild(element, 'a:fgClr')),
        bg: parseColorChild(firstChild(element, 'a:bgClr')),
      };
    case 'a:gradFill': {
      const stops: GradientStop[] = [];
      const list = firstChild(element, 'a:gsLst');
      if (list !== undefined) {
        for (const gs of childElements(list)) {
          if (gs.qname !== 'a:gs') continue;
          const color = parseColorChild(gs);
          if (color === null) continue;
          stops.push({ pos: percentAttr(gs, 'pos', 0), color });
        }
      }
      const flip = (attributeValue(element, 'flip') ?? 'none') as TileFlipMode;
      return {
        type: 'gradient',
        stops,
        shade: parseGradientShade(element),
        tileRect:
          firstChild(element, 'a:tileRect') === undefined
            ? null
            : parseRelativeRect(firstChild(element, 'a:tileRect')),
        flip,
        rotWithShape: boolAttr(element, 'rotWithShape', true),
      };
    }
    default:
      throw new PaintError(
        'FILL_PATTERN_UNKNOWN',
        `${element.qname} is not a fill element`,
        element.qname,
      );
  }
}

/** The fill a container declares, or `undefined` when it declares none. */
export function parseFill(parent: XElement, partName: string): Fill | undefined {
  for (const child of childElements(parent)) {
    if (FILL_ELEMENTS.has(child.qname)) return parseFillElement(child, partName);
  }
  return undefined;
}

/* -------------------------------------------------------------------------- */
/* lines                                                                      */
/* -------------------------------------------------------------------------- */

const JOIN_KINDS: Readonly<Record<string, LineJoinKind>> = {
  'a:round': 'round',
  'a:bevel': 'bevel',
  'a:miter': 'miter',
};

function parseLineEnd(element: XElement | undefined): LineEnd | null {
  if (element === undefined) return null;
  const size = (name: string): LineEndSize => {
    const raw = attributeValue(element, name);
    return raw === 'sm' || raw === 'med' || raw === 'lg' ? raw : 'med';
  };
  return { type: attributeValue(element, 'type') ?? 'none', w: size('w'), len: size('len') };
}

function parseDash(element: XElement): LineDash | null {
  const preset = firstChild(element, 'a:prstDash');
  if (preset !== undefined) {
    return { kind: 'preset', val: attributeValue(preset, 'val') ?? 'solid' };
  }
  const custom = firstChild(element, 'a:custDash');
  if (custom === undefined) return null;
  const stops: DashSegment[] = [];
  for (const ds of childElements(custom)) {
    if (ds.qname !== 'a:ds') continue;
    // `@d` and `@sp` are percentages of the line width, so a `1000` is ten
    // widths and not a thousand of anything.
    stops.push({ d: percentAttr(ds, 'd', 0) / 100000, sp: percentAttr(ds, 'sp', 0) / 100000 });
  }
  return { kind: 'custom', stops };
}

/** `a:ln`, lazily: every field is `null` when the attribute was absent. */
export function parseLineElement(element: XElement, partName: string): Line {
  let join: LineJoin | null = null;
  for (const child of childElements(element)) {
    const kind = JOIN_KINDS[child.qname];
    if (kind === undefined) continue;
    const lim = attributeValue(child, 'lim');
    join = { kind, limit: lim === undefined ? null : parsePercentage(lim) / 100000 };
    break;
  }

  const cap = attributeValue(element, 'cap');
  const algn = attributeValue(element, 'algn');
  const w = attributeValue(element, 'w');

  return {
    w: w === undefined ? null : intAttr(element, 'w', 0, 'a:ln'),
    cap: cap === 'flat' || cap === 'sq' || cap === 'rnd' ? (cap satisfies LineCap) : null,
    cmpd: attributeValue(element, 'cmpd') ?? null,
    algn: algn === 'ctr' || algn === 'in' ? (algn satisfies PenAlignment) : null,
    fill: parseFill(element, partName) ?? null,
    dash: parseDash(element),
    join,
    headEnd: parseLineEnd(firstChild(element, 'a:headEnd')),
    tailEnd: parseLineEnd(firstChild(element, 'a:tailEnd')),
  };
}

export function parseLine(parent: XElement, partName: string): Line | undefined {
  const element = firstChild(parent, 'a:ln');
  return element === undefined ? undefined : parseLineElement(element, partName);
}

/* -------------------------------------------------------------------------- */
/* effects                                                                    */
/* -------------------------------------------------------------------------- */

const RECT_ALIGNMENTS = new Set(['tl', 't', 'tr', 'l', 'ctr', 'r', 'bl', 'b', 'br']);

function alignAttr(element: XElement, fallback: RectAlignment): RectAlignment {
  const raw = attributeValue(element, 'algn');
  return raw !== undefined && RECT_ALIGNMENTS.has(raw) ? (raw as RectAlignment) : fallback;
}

function shadowGeometry(element: XElement): ShadowGeometry {
  return {
    blurRad: intAttr(element, 'blurRad', 0, element.qname),
    dist: intAttr(element, 'dist', 0, element.qname),
    dir: intAttr(element, 'dir', 0, element.qname),
    sx: percentAttr(element, 'sx', 100000),
    sy: percentAttr(element, 'sy', 100000),
    kx: intAttr(element, 'kx', 0, element.qname),
    ky: intAttr(element, 'ky', 0, element.qname),
    // Measured in 2.8: bottom-centre, not the centre. `sy="-100000"` mirrors
    // about the bottom edge, which is what a floor reflection wants.
    algn: alignAttr(element, 'b'),
    rotWithShape: boolAttr(element, 'rotWithShape', true),
  };
}

function requireColor(element: XElement): Color {
  const color = parseColorChild(element);
  if (color === null) {
    throw new PaintError(
      'COLOR_SCHEME_NAME',
      `${element.qname} has no colour child`,
      element.qname,
    );
  }
  return color;
}

/** One member of `a:effectLst`, or `null` for an element this does not model. */
function parseEffectElement(element: XElement): Effect | null {
  switch (element.qname) {
    case 'a:outerShdw':
      return { kind: 'outerShdw', color: requireColor(element), ...shadowGeometry(element) };
    case 'a:innerShdw':
      return {
        kind: 'innerShdw',
        color: requireColor(element),
        ...shadowGeometry(element),
      };
    case 'a:prstShdw':
      return {
        kind: 'prstShdw',
        prst: attributeValue(element, 'prst') ?? '',
        dist: intAttr(element, 'dist', 0, 'a:prstShdw'),
        dir: intAttr(element, 'dir', 0, 'a:prstShdw'),
        color: requireColor(element),
      };
    case 'a:glow':
      return {
        kind: 'glow',
        rad: intAttr(element, 'rad', 0, 'a:glow'),
        color: requireColor(element),
      };
    case 'a:softEdge':
      return { kind: 'softEdge', rad: intAttr(element, 'rad', 0, 'a:softEdge') };
    case 'a:blur':
      return {
        kind: 'blur',
        rad: intAttr(element, 'rad', 0, 'a:blur'),
        grow: boolAttr(element, 'grow', true),
      };
    case 'a:reflection':
      return {
        kind: 'reflection',
        blurRad: intAttr(element, 'blurRad', 0, 'a:reflection'),
        stA: percentAttr(element, 'stA', 100000),
        stPos: percentAttr(element, 'stPos', 0),
        endA: percentAttr(element, 'endA', 0),
        endPos: percentAttr(element, 'endPos', 100000),
        dist: intAttr(element, 'dist', 0, 'a:reflection'),
        dir: intAttr(element, 'dir', 0, 'a:reflection'),
        fadeDir: intAttr(element, 'fadeDir', 5400000, 'a:reflection'),
        sx: percentAttr(element, 'sx', 100000),
        sy: percentAttr(element, 'sy', 100000),
        kx: intAttr(element, 'kx', 0, 'a:reflection'),
        ky: intAttr(element, 'ky', 0, 'a:reflection'),
        algn: alignAttr(element, 'b'),
        rotWithShape: boolAttr(element, 'rotWithShape', true),
      };
    default:
      return null;
  }
}

/**
 * `a:effectLst`, in the order the file wrote it.
 *
 * Which is not the order it is painted in: 2.8 measured a glow painting *over*
 * an outer shadow, the reverse of the schema's sequence. `effectFilter` in
 * `paint` owns that, and this owns only the reading.
 *
 * `a:effectDag` is not modelled. It is a directed graph of the same primitives
 * that PowerPoint has never been observed to write, and an `undefined` here
 * means the shape is re-emitted byte for byte rather than approximated.
 */
export function parseEffects(parent: XElement): readonly Effect[] | undefined {
  const list = firstChild(parent, 'a:effectLst');
  if (list === undefined) return undefined;
  const out: Effect[] = [];
  for (const child of childElements(list)) {
    const effect = parseEffectElement(child);
    if (effect !== null) out.push(effect);
  }
  return out;
}

/** The 28 transform names, re-exported so a caller can test a tag without paint. */
export { COLOR_TRANSFORM_OPS };
