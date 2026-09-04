/**
 * The debug overlay: what a shape's geometry actually resolved to.
 *
 * Sub-phase 2.11. The plan calls this "simultaneously the debugging tool and
 * 90% of the eventual adjust-handle editing UI", and that is the reason it sits
 * here rather than in `render-dom`. Everything it draws is a fact about the
 * geometry - where the paths went, where the text goes, where a connector may
 * attach, how far a handle can travel - so it is built as the same `SvgNode`
 * tree the shapes are, `render-dom` mounts it with the identical code path, and
 * the two cannot disagree about where a handle is. `render-dom` adds only the
 * pointer.
 *
 * ## Five things, and each of them answers a question you actually ask
 *
 * - **The paths, one colour each.** 60 of the 187 presets have paths that
 *   disagree about fill or stroke, and a wrong `d` on the second path of
 *   `smileyFace` is invisible in the painted shape and obvious here.
 * - **`a:rect`.** The text rectangle is not the shape's box - on a chevron or a
 *   pentagon it is inset by a whole arrowhead - and 3.6 will lay text into it.
 *   Five presets declare none at all, which the overlay says rather than
 *   inventing one.
 * - **`cxnLst`.** Where a connector may attach, with the angle it leaves at.
 * - **`ahLst`, with the locus.** Not just the handle: the actual curve its
 *   position traces as its guide sweeps its declared range, sampled through the
 *   same evaluator that draws the shape. A handle whose range is degenerate, or
 *   whose bounds had to be widened to admit the shape's own value, is visible
 *   at a glance instead of after an afternoon.
 * - **`gdLst`.** Every guide, built-in and computed, at the size on screen.
 *
 * ## Units
 *
 * Everything is EMU, like the rest of the renderer. Chrome that should stay one
 * size on screen - a handle square, a dashed rule - is therefore sized in EMU
 * per screen pixel, which the caller knows and this does not: pass `unit`.
 */

import {
  builtinGuideNames,
  resolveHandles,
  type Point,
  type PresetAdjustHandle,
  type ResolvedConnectionSite,
  type ResolvedGeometry,
  type ResolvedHandle,
  type ResolvedHandleAxis,
  type ResolvedRect,
} from '@pptx-studio/geometry';

import type { GeometrySource, Placed } from './layout.js';
import { element, num, type SvgElement, type SvgNode } from './node.js';
import { frameTransform } from './transform.js';

/** One EMU per screen pixel at 100% on a 10-inch-wide slide is 12700. */
export const DEFAULT_UNIT = 12700;

/** How many positions of a handle's locus are drawn. */
export const LOCUS_SAMPLES = 25;

/**
 * A path colour per subpath, cycled.
 *
 * Chosen to stay apart from each other on both a white and a dark ground and
 * from the blue every other piece of chrome here uses. Six is enough: the most
 * paths any preset declares is five (`actionButtonMovie`).
 */
export const PATH_COLORS: readonly string[] = [
  '#E8453C',
  '#1B9E4B',
  '#F09300',
  '#9C27B0',
  '#00838F',
  '#C2185B',
];

const CHROME = '#0B62B0';
const TEXT_RECT = '#7B1FA2';
const SITE = '#00838F';
const HANDLE = '#E8453C';
const WIDENED = '#B26A00';

export interface OverlayOptions {
  /** EMU per screen pixel, so chrome stays one thickness at any zoom. */
  readonly unit?: number;
  readonly paths?: boolean;
  readonly textRect?: boolean;
  readonly connectionSites?: boolean;
  readonly handles?: boolean;
  /** Sample each handle axis across its range and draw the curve. Default true. */
  readonly locus?: boolean;
}

/** One guide, for the live `gdLst` dump. */
export interface OverlayGuide {
  readonly name: string;
  readonly value: number;
  /** True for one of the 44 seeded built-ins rather than the shape's own. */
  readonly builtin: boolean;
  readonly finite: boolean;
}

/** One handle, with everything the pointer and the panel both need. */
export interface OverlayHandle {
  readonly index: number;
  readonly handle: PresetAdjustHandle;
  readonly resolved: ResolvedHandle;
  /** The curve `resolved.pos` traces as each axis sweeps its range, in shape space. */
  readonly locus: readonly (readonly Point[])[];
}

export interface ShapeOverlay {
  /** The whole overlay for this shape, already carrying the shape's transform. */
  readonly node: SvgElement;
  readonly guides: readonly OverlayGuide[];
  readonly handles: readonly OverlayHandle[];
  readonly textRect: ResolvedRect | null;
  readonly connectionSites: readonly ResolvedConnectionSite[];
  /** Paths whose evaluation reached a non-finite number, by index. */
  readonly brokenPaths: readonly number[];
}

const EMPTY: ShapeOverlay = {
  node: element('g', { 'data-role': 'overlay' }),
  guides: [],
  handles: [],
  textRect: null,
  connectionSites: [],
  brokenPaths: [],
};

/**
 * Where a handle's position goes as one axis sweeps its declared range.
 *
 * Sampled rather than solved, and through `resolveHandles` rather than any
 * private path, so the curve is drawn by the same evaluator that draws the
 * shape. A polar handle's locus is a genuine arc and a `blockArc`'s radius axis
 * is not a straight line, so a two-point min-to-max rule would draw a chord
 * through the outside of the shape and call it a range.
 */
function locusOf(
  source: GeometrySource,
  size: { w: number; h: number },
  index: number,
  axis: ResolvedHandleAxis,
  base: Readonly<Record<string, number>>,
): Point[] {
  if (axis.max === axis.min) return [];
  const points: Point[] = [];
  for (let step = 0; step < LOCUS_SAMPLES; step++) {
    const t = step / (LOCUS_SAMPLES - 1);
    const adjust = { ...base, [axis.guide]: axis.min + (axis.max - axis.min) * t };
    const at = resolveHandles(source.geometry, size, { adjust })[index];
    if (at === undefined || !at.finite) continue;
    points.push(at.pos);
  }
  return points;
}

const finitePoint = (point: Point): boolean => Number.isFinite(point.x) && Number.isFinite(point.y);

/**
 * Path data, through `num`.
 *
 * Every coordinate that reaches a `d` goes through it, and not for tidiness: a
 * `d` holding a `NaN` is invalid, and a browser handed one drops the whole
 * element without a word. `num` is the single place that is guaranteed never to
 * emit one - the attribute path gets it for free because the serializer calls
 * it, and a string built here has to ask.
 */
function polyline(points: readonly Point[]): string {
  return points.map((p, i) => `${i === 0 ? 'M' : 'L'}${num(p.x, 1)} ${num(p.y, 1)}`).join('');
}

/** A diamond, because a circle at this size is a blob and a square is a handle. */
function diamond(at: Point, radius: number): string {
  return (
    `M${num(at.x)} ${num(at.y - radius)}` +
    `L${num(at.x + radius)} ${num(at.y)}` +
    `L${num(at.x)} ${num(at.y + radius)}` +
    `L${num(at.x - radius)} ${num(at.y)}Z`
  );
}

/** The resolved outlines, one colour each. Non-finite paths are counted, not drawn. */
function outlineNodes(
  geometry: ResolvedGeometry,
  unit: number,
  broken: number[],
): readonly SvgNode[] {
  const nodes: SvgNode[] = [];
  geometry.paths.forEach((path, index) => {
    if (!path.finite) {
      broken.push(index);
      return;
    }
    if (path.d === '') return;
    nodes.push(
      element('path', {
        'data-role': 'outline',
        'data-path': index,
        d: path.d,
        fill: 'none',
        stroke: PATH_COLORS[index % PATH_COLORS.length] ?? CHROME,
        'stroke-width': unit * 1.75,
        'stroke-linejoin': 'round',
        // A path the definition marks unstroked is still drawn here - it is an
        // outline, not a rendering - but dashed, so the two are told apart.
        ...(path.stroke ? {} : { 'stroke-dasharray': `${num(unit * 4)} ${num(unit * 3)}` }),
      }),
    );
  });
  return nodes;
}

/**
 * The text rectangle and the connection sites.
 *
 * Both are dropped when their coordinates are not finite, which a zero extent
 * reaches: the guides divide by the size and everything downstream is `NaN`.
 * Drawing them at the origin would be an invention, and the broken path list
 * already says what went wrong.
 */
function anchorNodes(
  geometry: ResolvedGeometry,
  unit: number,
  options: OverlayOptions,
): readonly SvgNode[] {
  const hair = unit * 1.25;
  const nodes: SvgNode[] = [];
  const rect = geometry.textRect;
  const rectFinite =
    rect !== null && [rect.l, rect.t, rect.r, rect.b].every((value) => Number.isFinite(value));

  if (options.textRect !== false && rect !== null && rectFinite) {
    nodes.push(
      element('rect', {
        'data-role': 'text-rect',
        x: rect.l,
        y: rect.t,
        width: Math.max(0, rect.r - rect.l),
        height: Math.max(0, rect.b - rect.t),
        fill: 'none',
        stroke: TEXT_RECT,
        'stroke-width': hair,
        'stroke-dasharray': `${num(unit * 5)} ${num(unit * 4)}`,
      }),
    );
  }

  const sites = geometry.connectionSites.filter((site) => finitePoint(site.pos));
  if (options.connectionSites !== false && sites.length > 0) {
    nodes.push(
      element(
        'g',
        { 'data-role': 'connection-sites' },
        sites.map((site, index) =>
          element('path', {
            'data-site': index,
            // `@ang` stays in 60000ths of a degree, as `geometry` hands it over.
            'data-ang': site.ang,
            d: diamond(site.pos, unit * 3.5),
            fill: SITE,
            'fill-opacity': 0.85,
            stroke: '#FFFFFF',
            'stroke-width': hair * 0.8,
          }),
        ),
      ),
    );
  }
  return nodes;
}

/** One handle: its locus per axis, then the square that is grabbed. */
function handleNodes(handle: OverlayHandle, unit: number): readonly SvgNode[] {
  const hair = unit * 1.25;
  const half = unit * 4;
  const { resolved, index } = handle;
  const nodes: SvgNode[] = [];

  handle.locus.forEach((points, axisAt) => {
    if (points.length < 2) return;
    nodes.push(
      element('path', {
        'data-role': 'handle-locus',
        'data-handle': index,
        'data-axis': axisAt,
        d: polyline(points),
        fill: 'none',
        // An axis whose declared bounds had to be widened to admit the shape's
        // own value is marked, because it means the definition and the file
        // disagree and a drag will clamp somewhere unexpected.
        stroke: resolved.axes[axisAt]?.widened === true ? WIDENED : HANDLE,
        'stroke-opacity': 0.55,
        'stroke-width': hair,
        'stroke-dasharray': `${num(unit * 3)} ${num(unit * 2)}`,
      }),
    );
  });

  nodes.push(
    element('rect', {
      'data-role': 'handle',
      'data-handle': index,
      'data-kind': resolved.kind,
      // Empty when the handle declares no `gdRef` at all, which is a handle
      // that cannot be dragged. No preset has one; an `a:custGeom` may.
      'data-axes': resolved.axes.map((axis) => axis.guide).join(' '),
      x: resolved.pos.x - half,
      y: resolved.pos.y - half,
      width: half * 2,
      height: half * 2,
      fill: resolved.axes.length === 0 ? '#FFFFFF' : HANDLE,
      stroke: '#FFFFFF',
      'stroke-width': hair,
    }),
  );
  return nodes;
}

/** Every handle the definition declares, resolved at this size. */
function resolveOverlayHandles(
  source: GeometrySource,
  size: { w: number; h: number },
  withLocus: boolean,
): readonly OverlayHandle[] {
  const resolved = resolveHandles(source.geometry, size, { adjust: source.adjust });
  const base: Record<string, number> = {};
  for (const handle of resolved) {
    for (const axis of handle.axes) base[axis.guide] = axis.value;
  }
  const out: OverlayHandle[] = [];
  resolved.forEach((entry, index) => {
    const declared = source.geometry.ahLst[index];
    if (declared === undefined) return;
    out.push({
      index,
      handle: declared,
      resolved: entry,
      locus:
        withLocus && entry.finite
          ? entry.axes.map((axis) => locusOf(source, size, index, axis, base))
          : [],
    });
  });
  return out;
}

/**
 * The overlay for one placed shape.
 *
 * Returns an empty overlay for anything with no geometry - a group, a
 * `p:graphicFrame` - rather than throwing, because a selection is allowed to
 * land on one and an overlay is not the place to be strict about it.
 */
export function shapeOverlay(placed: Placed, options: OverlayOptions = {}): ShapeOverlay {
  const geometry = placed.geometry;
  const source = placed.geometrySource;
  if (geometry === null || source === null) return EMPTY;

  const unit = options.unit ?? DEFAULT_UNIT;
  const size = { w: placed.frame.cx, h: placed.frame.cy };
  const builtins = new Set(builtinGuideNames());

  const guides: OverlayGuide[] = [...geometry.guides].map(([name, value]) => ({
    name,
    value,
    builtin: builtins.has(name),
    finite: Number.isFinite(value),
  }));

  const brokenPaths: number[] = [];
  const outlines = options.paths === false ? [] : outlineNodes(geometry, unit, brokenPaths);
  const anchors = anchorNodes(geometry, unit, options);

  const handles =
    options.handles === false ? [] : resolveOverlayHandles(source, size, options.locus !== false);
  const drawn = handles.filter((handle) => handle.resolved.finite);
  const handleGroup =
    drawn.length === 0
      ? []
      : [
          element(
            'g',
            { 'data-role': 'handles' },
            drawn.flatMap((handle) => handleNodes(handle, unit)),
          ),
        ];

  const transform = frameTransform(placed.frame);
  return {
    node: element(
      'g',
      {
        'data-role': 'overlay',
        'data-shape': String(placed.shape.cNvPrId),
        ...(transform === '' ? {} : { transform }),
      },
      [...outlines, ...anchors, ...handleGroup],
    ),
    guides,
    handles,
    textRect: geometry.textRect,
    connectionSites: geometry.connectionSites,
    brokenPaths,
  };
}
