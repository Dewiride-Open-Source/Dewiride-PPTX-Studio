/**
 * `a:prstGeom` and `a:custGeom`, into the one type `@pptx-studio/geometry` takes.
 *
 * ## Why this is in the model rather than in the renderer
 *
 * Because geometry inherits. A slide placeholder that states no `a:spPr/a:xfrm`
 * usually states no `a:prstGeom` either, and takes both from its layout - so
 * "which shape is this" is a question for the resolver, and the resolver only
 * works over fields on `Shape`. Reading the geometry off `shape.node` in a
 * renderer would put half of one answer in the model and the other half
 * somewhere that cannot see the chain.
 *
 * 2.9 recorded only the preset's *name*, which is enough to say what a shape is
 * and not enough to draw it: `a:avLst` is where a rounded rectangle's corner
 * radius and a chevron's notch depth live, and a preset drawn with its defaults
 * when the file overrode them is a visibly wrong shape.
 *
 * ## The two are one type, and that is 2.4's finding rather than a convenience
 *
 * `a:custGeom` carries `avLst`, `gdLst`, `ahLst`, `cxnLst`, `rect` and
 * `pathLst`: the same six children with the same meanings as a preset
 * definition, written inline instead of looked up by name. So both parse into
 * `Geometry` and the evaluator has one code path. The only difference is that a
 * preset has a name and a `custGeom` does not.
 */

import {
  custGeom,
  type Geometry,
  type PresetAdjustHandle,
  type PresetCommand,
  type PresetConnectionSite,
  type PresetGuide,
  type PresetPath,
  type PresetPathFill,
  type PresetPoint,
  type PresetTextRect,
} from '@pptx-studio/geometry';
import { attributeValue, childElements, firstChild, type XElement } from '@pptx-studio/xml';

import { ModelError } from '../errors.js';

/**
 * What a shape says it is.
 *
 * `preset` keeps the adjust values as guides rather than numbers, because
 * `a:avLst` is a `CT_GeomGuideList` and a guide may in principle carry any
 * formula - `evaluateGuides` takes either form and the file's form is the one
 * that round-trips.
 */
export type ShapeGeometry =
  | { readonly kind: 'preset'; readonly prst: string; readonly adjust: readonly PresetGuide[] }
  | { readonly kind: 'custom'; readonly geometry: Geometry };

function guides(parent: XElement | undefined, qname: string): PresetGuide[] {
  const list = parent === undefined ? undefined : firstChild(parent, qname);
  if (list === undefined) return [];
  const out: PresetGuide[] = [];
  for (const child of childElements(list)) {
    if (child.qname !== 'a:gd') continue;
    const name = attributeValue(child, 'name');
    const fmla = attributeValue(child, 'fmla');
    if (name === undefined || fmla === undefined) continue;
    // Split on whitespace and keep every token: the arity is not fixed, and
    // `OVER_LONG_FORMULAS` in the preset buckets is the evidence.
    out.push({ name, fmla: fmla.trim().split(/\s+/) });
  }
  return out;
}

function point(element: XElement, partName: string): PresetPoint {
  const x = attributeValue(element, 'x');
  const y = attributeValue(element, 'y');
  if (x === undefined || y === undefined) {
    throw new ModelError('MODEL_GEOMETRY', `<${element.qname}> has no x or no y`, partName);
  }
  return { x, y };
}

const PATH_FILLS = new Set<string>([
  'none',
  'norm',
  'lighten',
  'lightenLess',
  'darken',
  'darkenLess',
]);

function pathFill(raw: string | undefined): PresetPathFill {
  return raw !== undefined && PATH_FILLS.has(raw) ? (raw as PresetPathFill) : 'norm';
}

/** `@w`/`@h` on an `a:path`. Absent means zero, which means no path space. */
function pathExtent(element: XElement, name: string, partName: string): number {
  const raw = attributeValue(element, name);
  if (raw === undefined) return 0;
  const value = Number(raw);
  if (!Number.isFinite(value)) {
    throw new ModelError('MODEL_GEOMETRY', `a:path/@${name} is "${raw}"`, partName);
  }
  return value;
}

function commands(path: XElement, partName: string): PresetCommand[] {
  const out: PresetCommand[] = [];
  for (const child of childElements(path)) {
    switch (child.qname) {
      case 'a:moveTo':
      case 'a:lnTo': {
        const pt = firstChild(child, 'a:pt');
        if (pt === undefined) break;
        out.push({
          kind: child.qname === 'a:moveTo' ? 'moveTo' : 'lnTo',
          to: point(pt, partName),
        });
        break;
      }
      case 'a:quadBezTo': {
        const pts = childElements(child).filter((c) => c.qname === 'a:pt');
        if (pts.length < 2) break;
        out.push({
          kind: 'quadBezTo',
          c1: point(pts[0]!, partName),
          to: point(pts[1]!, partName),
        });
        break;
      }
      case 'a:cubicBezTo': {
        const pts = childElements(child).filter((c) => c.qname === 'a:pt');
        if (pts.length < 3) break;
        out.push({
          kind: 'cubicBezTo',
          c1: point(pts[0]!, partName),
          c2: point(pts[1]!, partName),
          to: point(pts[2]!, partName),
        });
        break;
      }
      case 'a:arcTo': {
        const wR = attributeValue(child, 'wR');
        const hR = attributeValue(child, 'hR');
        const stAng = attributeValue(child, 'stAng');
        const swAng = attributeValue(child, 'swAng');
        if (wR === undefined || hR === undefined || stAng === undefined || swAng === undefined) {
          throw new ModelError('MODEL_GEOMETRY', 'a:arcTo is missing an attribute', partName);
        }
        out.push({ kind: 'arcTo', wR, hR, stAng, swAng });
        break;
      }
      case 'a:close':
        out.push({ kind: 'close' });
        break;
      default:
        break;
    }
  }
  return out;
}

function paths(parent: XElement, partName: string): PresetPath[] {
  const list = firstChild(parent, 'a:pathLst');
  if (list === undefined) return [];
  const out: PresetPath[] = [];
  for (const child of childElements(list)) {
    if (child.qname !== 'a:path') continue;
    out.push({
      w: pathExtent(child, 'w', partName),
      h: pathExtent(child, 'h', partName),
      fill: pathFill(attributeValue(child, 'fill')),
      // `@stroke` defaults to true, which is not the same as "absent means no".
      stroke: attributeValue(child, 'stroke') !== '0',
      extrusionOk: attributeValue(child, 'extrusionOk') === '1',
      commands: commands(child, partName),
    });
  }
  return out;
}

function handles(parent: XElement, partName: string): PresetAdjustHandle[] {
  const list = firstChild(parent, 'a:ahLst');
  if (list === undefined) return [];
  const out: PresetAdjustHandle[] = [];
  for (const child of childElements(list)) {
    const pos = firstChild(child, 'a:pos');
    if (pos === undefined) continue;
    const at = point(pos, partName);
    if (child.qname === 'a:ahXY') {
      out.push({
        kind: 'xy',
        pos: at,
        gdRefX: attributeValue(child, 'gdRefX') ?? null,
        minX: attributeValue(child, 'minX') ?? null,
        maxX: attributeValue(child, 'maxX') ?? null,
        gdRefY: attributeValue(child, 'gdRefY') ?? null,
        minY: attributeValue(child, 'minY') ?? null,
        maxY: attributeValue(child, 'maxY') ?? null,
      });
    } else if (child.qname === 'a:ahPolar') {
      out.push({
        kind: 'polar',
        pos: at,
        gdRefAng: attributeValue(child, 'gdRefAng') ?? null,
        minAng: attributeValue(child, 'minAng') ?? null,
        maxAng: attributeValue(child, 'maxAng') ?? null,
        gdRefR: attributeValue(child, 'gdRefR') ?? null,
        minR: attributeValue(child, 'minR') ?? null,
        maxR: attributeValue(child, 'maxR') ?? null,
      });
    }
  }
  return out;
}

function sites(parent: XElement, partName: string): PresetConnectionSite[] {
  const list = firstChild(parent, 'a:cxnLst');
  if (list === undefined) return [];
  const out: PresetConnectionSite[] = [];
  for (const child of childElements(list)) {
    if (child.qname !== 'a:cxn') continue;
    const pos = firstChild(child, 'a:pos');
    const ang = attributeValue(child, 'ang');
    if (pos === undefined || ang === undefined) continue;
    out.push({ ang, pos: point(pos, partName) });
  }
  return out;
}

function textRect(parent: XElement): PresetTextRect | null {
  const rect = firstChild(parent, 'a:rect');
  if (rect === undefined) return null;
  const l = attributeValue(rect, 'l');
  const t = attributeValue(rect, 't');
  const r = attributeValue(rect, 'r');
  const b = attributeValue(rect, 'b');
  return l === undefined || t === undefined || r === undefined || b === undefined
    ? null
    : { l, t, r, b };
}

/**
 * The geometry a shape declares, or `undefined` when it declares none.
 *
 * `undefined` is the whole point: a slide placeholder with no `a:prstGeom` has
 * not chosen to be a rectangle, it has said nothing, and `resolve` is what turns
 * that into the layout's answer.
 */
export function parseGeometry(
  spPr: XElement | undefined,
  partName: string,
): ShapeGeometry | undefined {
  if (spPr === undefined) return undefined;

  const preset = firstChild(spPr, 'a:prstGeom');
  if (preset !== undefined) {
    const prst = attributeValue(preset, 'prst');
    if (prst === undefined) {
      throw new ModelError('MODEL_GEOMETRY', 'a:prstGeom has no @prst', partName);
    }
    return { kind: 'preset', prst, adjust: guides(preset, 'a:avLst') };
  }

  const custom = firstChild(spPr, 'a:custGeom');
  if (custom === undefined) return undefined;
  return {
    kind: 'custom',
    geometry: custGeom({
      avLst: guides(custom, 'a:avLst'),
      gdLst: guides(custom, 'a:gdLst'),
      ahLst: handles(custom, partName),
      cxnLst: sites(custom, partName),
      rect: textRect(custom),
      pathLst: paths(custom, partName),
    }) satisfies Geometry,
  };
}
