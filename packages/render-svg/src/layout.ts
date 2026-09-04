/**
 * A sheet into a flat, positioned tree - the one pass both renderers share.
 *
 * `render-svg` turns this into a string and `render-dom` into elements. Neither
 * of them resolves anything: by the time a `Placed` exists, every inheritance
 * has been followed, every group transform composed, and every `a:grpFill`
 * chased up to the fill it actually shows. That is what keeps the two renderers
 * from drifting, and it is the whole reason this module is separate from the
 * emitter that happens to live beside it.
 *
 * ## A group paints nothing of its own
 *
 * Measured in C6, and the deck says it plainly: a `p:grpSp` with a solid fill
 * over a rectangle nothing else covers shows white. A group's fill exists only
 * to be asked for by a child's `a:grpFill`. So `Placed.fill` is `null` for every
 * group, and a renderer that paints one is painting something PowerPoint does
 * not.
 *
 * ## And `a:grpFill` is a slice of a rectangle nobody writes down
 *
 * A `grpFill` child shows the enclosing group's fill laid out over **the
 * bounding box of everything that group contains** - not over the group's own
 * `a:off`/`a:ext`, which is what one would guess and what the file makes
 * obvious. Solved from the bitmap on four probes: a group declaring
 * `off.x=100, ext.cx=600` whose two children sit at 150..250 and 550..650 lays
 * its ramp over exactly 150..650, and moving one child moves the ramp with it.
 * A sibling that does not use `grpFill` still counts, and a sibling outside the
 * group's declared rectangle still counts.
 */

import {
  resolveGeometry,
  getPreset,
  type Geometry,
  type PresetGuide,
  type ResolvedGeometry,
} from '@pptx-studio/geometry';
import {
  ORPHAN_RECT,
  colorContextOf,
  resolve,
  resolveAppearance,
  resolveXfrm,
  type ResolvedAppearance,
  type Shape,
  type ShapeGeometry,
  type Sheet,
  type Xfrm,
} from '@pptx-studio/model';
import type { ColorContext, Fill, Rgba } from '@pptx-studio/paint';

import { RenderError } from './errors.js';
import {
  UNIT_CHILD_SPACE,
  childSpace,
  frameOf,
  placeChild,
  unionBox,
  type Box,
  type Frame,
} from './transform.js';

/** A shape with its position, outline and paint all resolved. */
export interface Placed {
  readonly shape: Shape;
  readonly sheet: Sheet;
  readonly frame: Frame;
  /**
   * The outline, evaluated at the frame's extent.
   *
   * `null` for a shape that declares no geometry and inherits none - a group, a
   * `p:graphicFrame`, or a text box that is nothing but its text.
   */
  readonly geometry: ResolvedGeometry | null;
  /**
   * The definition `geometry` was evaluated from, and the adjust values in
   * force.
   *
   * `geometry` is the answer at one size; this is what produced it. The adjust
   * handles need the definition itself - `resolveHandles` and `dragHandle` both
   * take a `Geometry` - and re-deriving it in the overlay would mean redoing the
   * inheritance walk somewhere that cannot see the sheet chain.
   */
  readonly geometrySource: GeometrySource | null;
  readonly appearance: ResolvedAppearance;
  /** What is painted, after `a:grpFill` has been followed. `null` paints nothing. */
  readonly fill: Fill | null;
  /** The colour `phClr` takes inside `fill`. */
  readonly fillPhClr: Rgba | null;
  /**
   * The rectangle the fill is laid out over, in slide EMU.
   *
   * The shape's own frame, except for an `a:grpFill`, where it is the content
   * bounds of the group the fill came from.
   */
  readonly fillBox: Box;
  readonly colorContext: ColorContext;
  readonly children: readonly Placed[];
  /** For a group: the bounding box of every leaf inside it. `null` otherwise. */
  readonly contentBounds: Box | null;
}

/** A geometry definition, with whatever `a:avLst` the file put over it. */
export interface GeometrySource {
  readonly geometry: Geometry;
  readonly adjust: readonly PresetGuide[];
}

/** How deep a group tree may nest before it is treated as hostile. */
export const MAX_GROUP_DEPTH = 64;

/**
 * The definition behind a shape's geometry.
 *
 * A `custGeom` is its own definition and carries no separate adjust list - its
 * `a:avLst` is already inside it - which is 2.4's finding restated: the two are
 * one type, and a preset is only a definition looked up by name.
 */
function sourceFor(spec: ShapeGeometry): GeometrySource {
  if (spec.kind === 'custom') return { geometry: spec.geometry, adjust: [] };
  const preset = getPreset(spec.prst);
  if (preset === undefined) {
    throw new RenderError(
      'RENDER_UNKNOWN_PRESET',
      `a:prstGeom prst="${spec.prst}" is not one of the 187 preset shapes`,
      spec.prst,
    );
  }
  return { geometry: preset satisfies Geometry, adjust: spec.adjust };
}

function geometryFor(source: GeometrySource, frame: Frame, name: string): ResolvedGeometry {
  return resolveGeometry(
    source.geometry,
    { w: frame.cx, h: frame.cy },
    { adjust: source.adjust, name },
  );
}

interface Descent {
  readonly sheet: Sheet;
  readonly frame: Frame;
  readonly space: ReturnType<typeof childSpace>;
  readonly depth: number;
}

function placeOne(shape: Shape, parent: Descent): Placed {
  const resolved = resolveXfrm(shape, parent.sheet);
  // A placeholder that matched nothing on its layout has no rectangle at all,
  // and PowerPoint renders it at the origin with no size rather than refusing
  // the file or inventing one. Measured in 2.9 on four probes.
  const own = frameOf(
    resolved?.value ?? { ...ORPHAN_RECT, rot: 0, flipH: false, flipV: false, child: null },
  );
  const frame = placeChild(parent.frame, parent.space, own);

  const appearance = resolveAppearance(shape, parent.sheet);
  const colorContext = colorContextOf(
    parent.sheet,
    appearance.fillPhClr === null ? undefined : appearance.fillPhClr,
  );

  const geometrySpec = resolve(shape, parent.sheet, (s) => s.geometry);
  const geometrySource = geometrySpec === undefined ? null : sourceFor(geometrySpec.value);
  const geometry = geometrySource === null ? null : geometryFor(geometrySource, frame, shape.name);

  let children: readonly Placed[] = [];
  let contentBounds: Box | null = null;
  if (shape.kind === 'grpSp') {
    if (parent.depth >= MAX_GROUP_DEPTH) {
      throw new RenderError(
        'RENDER_GROUP_DEPTH',
        `groups nest more than ${String(MAX_GROUP_DEPTH)} deep`,
        shape.name,
      );
    }
    const xfrm: Xfrm | undefined = resolved?.value;
    const descent: Descent = {
      sheet: parent.sheet,
      frame,
      space: xfrm === undefined ? UNIT_CHILD_SPACE : childSpace(frame, xfrm),
      depth: parent.depth + 1,
    };
    children = shape.children.map((child) => placeOne(child, descent));
    for (const child of children) {
      contentBounds = unionBox(contentBounds, child.contentBounds ?? child.frame);
    }
  }

  return {
    shape,
    sheet: parent.sheet,
    frame,
    geometry,
    geometrySource,
    appearance,
    // A group's own fill is never painted; `resolveGroupFills` gives the
    // children that ask for it their answer.
    fill: shape.kind === 'grpSp' ? null : appearance.fill,
    fillPhClr: appearance.fillPhClr,
    fillBox: frame,
    colorContext,
    children,
    contentBounds,
  };
}

/**
 * Give every `a:grpFill` the fill it actually shows, and the box it shows it in.
 *
 * A second pass rather than part of the first, because the answer needs the
 * group's *content bounds*, which are only known once its children are placed.
 *
 * The walk goes up to the nearest ancestor that states something other than
 * `a:grpFill` - so three groups deep, all saying `grpFill`, all reach the same
 * fill, and all reach it over the outermost one's content bounds. Measured, and
 * it is why the leaf in that probe read the middle of the ramp rather than a
 * tenth of the way along it.
 */
function resolveGroupFills(placed: Placed, ancestors: readonly Placed[]): Placed {
  const chain = [...ancestors, placed];
  let self = placed;

  if (placed.fill !== null && placed.fill.type === 'group') {
    let fill: Fill | null = null;
    let phClr: Rgba | null = null;
    let box: Box | null = null;
    for (let i = ancestors.length - 1; i >= 0; i--) {
      const group = ancestors[i]!;
      const candidate = group.appearance.fill;
      if (candidate === undefined || candidate === null) break;
      if (candidate.type === 'group') continue;
      fill = candidate;
      phClr = group.appearance.fillPhClr;
      box = group.contentBounds ?? group.frame;
      break;
    }
    self = { ...placed, fill, fillPhClr: phClr, fillBox: box ?? placed.frame };
  }

  return {
    ...self,
    children: self.children.map((child) => resolveGroupFills(child, chain)),
  };
}

/**
 * The whole shape tree of a sheet, placed.
 *
 * The top level is not wrapped in the `p:spTree`'s own `p:grpSpPr/a:xfrm`.
 * Every deck writes that as four zeros, and the measured rules make four zeros
 * the identity anyway - a `chExt` of zero is no scaling and a `chOff` of zero is
 * no offset - so the two readings agree on every file that exists. Placing the
 * top level directly is the one that also agrees on a file that does not.
 */
export function layoutSheet(sheet: Sheet): readonly Placed[] {
  const root: Descent = {
    sheet,
    frame: { x: 0, y: 0, cx: 0, cy: 0, rot: 0, flipH: false, flipV: false },
    space: UNIT_CHILD_SPACE,
    depth: 0,
  };
  return sheet.shapes
    .filter((shape) => !shape.hidden)
    .map((shape) => resolveGroupFills(placeOne(shape, root), []));
}

/**
 * The sheets whose own shapes appear beneath this one's, furthest away first.
 *
 * A slide shows its master's shapes and its layout's, and `@showMasterSp` is
 * what turns that off - but not in the way its name suggests. Measured in C6
 * with three squares, one on each sheet:
 *
 * ```text
 *   slide showMasterSp   layout showMasterSp   master   layout   slide
 *   (absent, so true)    (absent, so true)     shown    shown    shown
 *   0                    (true)                hidden   hidden   shown
 *   (true)               0                     hidden   shown    shown
 *   0                    0                     hidden   hidden   shown
 * ```
 *
 * So the attribute hides *everything inherited*, not the master's contribution
 * specifically: on a slide it takes the layout's furniture with it, and on a
 * layout it takes the master's. One recursive definition covers both rows, and
 * sub-phase 7.9 sells this as "hide layout graphics" - the name the behaviour
 * actually has.
 *
 * Placeholders are not in it. A master placeholder and a layout placeholder
 * that no slide shape matched were both invisible on the slide in all four
 * cases: an unmatched placeholder is a hole for content, not a drawing.
 */
export function inheritedSheets(sheet: Sheet): readonly Sheet[] {
  if (!sheet.showMasterShapes || sheet.parent === null) return [];
  return [...inheritedSheets(sheet.parent), sheet.parent];
}

/**
 * Everything drawn on a sheet: the inherited furniture, then its own shapes.
 *
 * Each inherited shape is placed and resolved against **the sheet it lives on**,
 * so a master's `schemeClr` reads the master's own colour map and theme.
 */
export function layoutSlide(sheet: Sheet): readonly Placed[] {
  const out: Placed[] = [];
  for (const ancestor of inheritedSheets(sheet)) {
    const root: Descent = {
      sheet: ancestor,
      frame: { x: 0, y: 0, cx: 0, cy: 0, rot: 0, flipH: false, flipV: false },
      space: UNIT_CHILD_SPACE,
      depth: 0,
    };
    for (const shape of ancestor.shapes) {
      if (shape.hidden || shape.placeholder !== null) continue;
      out.push(resolveGroupFills(placeOne(shape, root), []));
    }
  }
  out.push(...layoutSheet(sheet));
  return out;
}

/** Every placed shape in a tree, depth first, groups included. */
export function flatten(placed: readonly Placed[]): readonly Placed[] {
  const out: Placed[] = [];
  const walk = (nodes: readonly Placed[]): void => {
    for (const node of nodes) {
      out.push(node);
      walk(node.children);
    }
  };
  walk(placed);
  return out;
}
