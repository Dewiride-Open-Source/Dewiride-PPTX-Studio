/**
 * `a:ln` - the stroke.
 *
 * Everything here is a measurement, and four of the numbers contradict either
 * the standard or the obvious reading of it. They are called out where they are
 * defined; the evidence is `corpus/ground-truth/lines.json` and the reasoning is
 * in `docs/adr/0023-lines.md`.
 *
 * The shape of the model follows 2.6's `Color` and 2.7's `Fill`: every property
 * that a file may leave unstated is `null`, and the schema default is applied
 * once, in `resolveLine`, rather than at parse time. A renderer needs the
 * resolved value; an editor needs to know whether the file said it.
 */

import { PaintError } from './errors.js';
import { COMPOUND_RUNS_60, isCompoundName } from './compound-table.js';
import { PRESET_DASHES, type DashSegment } from './dash-table.js';
import { MARKERS, MARKER_SIZE, isMarkerSizeName, isMarkerTypeName } from './marker-table.js';
import type { Fill } from './fill.js';

/** English Metric Units in one point. A stroke width is always EMU. */
export const EMU_PER_POINT = 12700;

/**
 * The width of an `a:ln` that does not say.
 *
 * 9525 EMU is three quarters of a point, and it is what `Shape.Line.Weight`
 * reports for a stroke with no `@w` - confirmed against a stroke that declares
 * `w="9525"` explicitly and paints identically.
 *
 * Worth knowing what this is *not*. A shape PowerPoint's own UI has not styled
 * carries no `a:ln` element at all; its stroke comes from
 * `p:style/a:lnRef idx="2"` into the theme's `a:lnStyleLst`. So this default
 * governs a hand-written or programmatically-generated `a:ln`, not the ordinary
 * unstyled shape, whose width is a theme question that 2.9 answers.
 */
export const DEFAULT_LINE_WIDTH = 9525;

/**
 * The cap of an `a:ln` that does not say.
 *
 * ECMA-376 gives `ST_LineEndCap` a default of `sq`. Office uses `flat`, and the
 * difference is half a stroke width of ink at each end of every unstyled line in
 * every deck: a 200pt line at 24pt wide measured 300..500 with no `@cap` and
 * with `cap="flat"`, and 288..512 with `cap="sq"`.
 *
 * This is not an edge case hidden behind an unusual attribute. PowerPoint's own
 * authoring UI writes `@cap` for exactly two of its twelve dash styles - the two
 * it calls square dot and round dot - and nothing else. Every other line it has
 * ever written relies on this default.
 */
export const DEFAULT_LINE_CAP: LineCap = 'flat';

/**
 * The join of an `a:ln` with no join element.
 *
 * Round, measured: a stroked corner with no join element and one with `a:round`
 * both reach exactly half a stroke width past the corner, and a bevel reaches
 * 0.33. ECMA-376 states no default. SVG's is `miter`, so a renderer that simply
 * omits `stroke-linejoin` gets a different picture on every sharp corner.
 */
export const DEFAULT_LINE_JOIN: LineJoinKind = 'round';

/**
 * The miter limit of an `a:miter` with no `@lim`, as a multiple of the width.
 *
 * Eight, bracketed to a tenth: corners whose miter ratio was 2, 4, 5, 6, 7 and
 * 7.91 all painted a full miter, and 8.01, 8.11, 12 and 20 all fell back to a
 * bevel. `@lim` itself is a percentage of the whole stroke width and not of half
 * of it - on a corner of ratio 5, `lim="400000"` clips and `lim="600000"` does
 * not.
 *
 * SVG's default `stroke-miterlimit` is 4, so every corner between four and eight
 * widths renders differently unless this is written out explicitly.
 */
export const DEFAULT_MITER_LIMIT = 8;

export type LineCap = 'flat' | 'sq' | 'rnd';
export type LineJoinKind = 'round' | 'bevel' | 'miter';
export type PenAlignment = 'ctr' | 'in';
export type LineEndSize = 'sm' | 'med' | 'lg';

export interface LineJoin {
  readonly kind: LineJoinKind;
  /** A multiple of the stroke width. `null` means the attribute was absent. */
  readonly limit: number | null;
}

/** `a:headEnd` or `a:tailEnd`. */
export interface LineEnd {
  readonly type: string;
  readonly w: LineEndSize;
  readonly len: LineEndSize;
}

export type LineDash =
  | { readonly kind: 'preset'; readonly val: string }
  | { readonly kind: 'custom'; readonly stops: readonly DashSegment[] };

/**
 * `CT_LineProperties`, lazily.
 *
 * `null` everywhere means the file did not say. `w` is EMU.
 */
export interface Line {
  readonly w: number | null;
  readonly cap: LineCap | null;
  readonly cmpd: string | null;
  readonly algn: PenAlignment | null;
  /** `null` is a genuine absence; an explicit `a:noFill` is `{ type: 'none' }`. */
  readonly fill: Fill | null;
  readonly dash: LineDash | null;
  readonly join: LineJoin | null;
  readonly headEnd: LineEnd | null;
  readonly tailEnd: LineEnd | null;
}

export interface ResolvedLine {
  /** EMU. */
  readonly width: number;
  readonly cap: LineCap;
  readonly join: LineJoin & { readonly limit: number };
  readonly algn: PenAlignment;
  readonly cmpd: string;
  readonly fill: Fill;
  readonly dash: readonly DashSegment[] | null;
  readonly headEnd: LineEnd | null;
  readonly tailEnd: LineEnd | null;
}

/** Every property the file left unstated, filled in with what Office does. */
export function resolveLine(line: Line): ResolvedLine {
  const cmpd = line.cmpd ?? 'sng';
  if (!isCompoundName(cmpd)) {
    throw new PaintError('LINE_COMPOUND_UNKNOWN', `a:ln/@cmpd is "${cmpd}"`, cmpd);
  }
  const join = line.join ?? { kind: DEFAULT_LINE_JOIN, limit: null };
  return {
    width: line.w ?? DEFAULT_LINE_WIDTH,
    cap: line.cap ?? DEFAULT_LINE_CAP,
    join: { kind: join.kind, limit: join.limit ?? DEFAULT_MITER_LIMIT },
    algn: line.algn ?? 'ctr',
    cmpd,
    fill: line.fill ?? { type: 'none' },
    dash: line.dash === null ? null : dashSegments(line.dash),
    headEnd: line.headEnd,
    tailEnd: line.tailEnd,
  };
}

/** The on/off pattern for a dash, in multiples of the stroke width. */
export function dashSegments(dash: LineDash): readonly DashSegment[] {
  if (dash.kind === 'custom') return dash.stops;
  const preset = PRESET_DASHES[dash.val];
  if (preset === undefined) {
    throw new PaintError(
      'LINE_DASH_UNKNOWN',
      `a:prstDash/@val is "${dash.val}", which is not one of the eleven`,
      dash.val,
    );
  }
  return preset;
}

/**
 * The dash pattern in EMU, ready for SVG's `stroke-dasharray`.
 *
 * ## The cap compensation, which is the whole point of this function
 *
 * DrawingML's dash array describes the ink PowerPoint **paints**. SVG's
 * describes the path the cap is then added to, so a round or square cap extends
 * every dash by half a width at each end and eats the same from each gap. Copy
 * the array across unchanged with `stroke-linecap: round` and every dot is twice
 * as long as it should be.
 *
 * Measured rather than assumed: `dot` is `[1, 3]`, and with `cap="rnd"` it still
 * paints ink of exactly one width and a gap of exactly three, the same as
 * `cap="flat"`. So PowerPoint compensates, and to match it we shorten each dash
 * by one width and lengthen each gap by one - which preserves the period, and
 * turns `dot` into a zero-length segment that the round cap renders as the disc
 * it is meant to be.
 *
 * A segment that would go negative is clamped to zero rather than dropped; the
 * period has to survive or the pattern drifts along the line.
 */
export function dashArray(segments: readonly DashSegment[], width: number, cap: LineCap): number[] {
  const shrink = cap === 'flat' ? 0 : width;
  const out: number[] = [];
  for (const { d, sp } of segments) {
    out.push(Math.max(0, d * width - shrink), sp * width + shrink);
  }
  return out;
}

/** One rail of a compound stroke, as offsets from the geometry in EMU. */
export interface StrokeRail {
  /** Distance from the path to the rail's centre. Negative is outward. */
  readonly offset: number;
  /** EMU. */
  readonly width: number;
}

/**
 * A compound stroke's rails, measured out from the geometry.
 *
 * The five compound forms subdivide `@w`; they do not add to it. A `dbl` at 36pt
 * is two 12pt rails with a 12pt gap, spanning 36pt in total - not two 36pt rails
 * and not a 72pt band. `tri` is a sixth, a sixth of gap, a third, a sixth of gap
 * and a sixth.
 *
 * `algn` decides where that band sits. Centred, it straddles the geometry by
 * half a width either side; inset, it sits wholly inside, and "inside" means the
 * direction the shape's interior is in - which for a closed path is the side the
 * winding says and for an open one is a convention this function cannot know.
 * So the offsets are signed with negative meaning outward, and the caller
 * orients them.
 */
export function compoundRails(cmpd: string, width: number, algn: PenAlignment): StrokeRail[] {
  const runs = COMPOUND_RUNS_60[cmpd];
  if (runs === undefined) {
    throw new PaintError('LINE_COMPOUND_UNKNOWN', `a:ln/@cmpd is "${cmpd}"`, cmpd);
  }
  // Sixtieths, divided last. A third of a 36000 EMU stroke is 12000 exactly this
  // way and 11999.988 if the table carries a rounded decimal instead - and the
  // rails of a compound stroke have to close on the width, not nearly close.
  const scale = width / 60;
  // Centred, the band runs from -w/2 to +w/2 about the path; inset, from 0 to w.
  let at = algn === 'in' ? 0 : -width / 2;
  const rails: StrokeRail[] = [];
  runs.forEach((run, i) => {
    const length = run * scale;
    // Runs alternate rail, gap, rail, ... starting with a rail.
    if (i % 2 === 0) rails.push({ offset: at + length / 2, width: length });
    at += length;
  });
  return rails;
}

/**
 * The marker for one line end, in EMU.
 *
 * `length` and `width` are the marker's box; `refX` is how far along that box
 * the line's own endpoint falls, which is not the same for every type - a
 * diamond and an oval are centred on the endpoint and the other three put their
 * tip on it. Getting that wrong displaces a large head by half its length.
 */
export interface MarkerGeometry {
  readonly type: string;
  /** EMU, along the line. */
  readonly length: number;
  /** EMU, across it. */
  readonly width: number;
  /** EMU from the marker's back edge to the line's endpoint. */
  readonly refX: number;
  readonly filled: boolean;
  /** The outline in a unit box; scale by `length` and `width`. */
  readonly path: string;
}

export function markerGeometry(end: LineEnd | null, strokeWidth: number): MarkerGeometry | null {
  if (end === null || end.type === 'none') return null;
  if (!isMarkerTypeName(end.type)) {
    throw new PaintError(
      'LINE_END_UNKNOWN',
      `a:headEnd/@type or a:tailEnd/@type is "${end.type}", which is not one of the six`,
      end.type,
    );
  }
  if (!isMarkerSizeName(end.len) || !isMarkerSizeName(end.w)) {
    const bad = isMarkerSizeName(end.len) ? end.w : end.len;
    throw new PaintError('LINE_END_SIZE', `a line end size is "${bad}", not sm, med or lg`, bad);
  }
  const marker = MARKERS[end.type]!;
  const length = MARKER_SIZE[end.len]! * strokeWidth;
  const width = MARKER_SIZE[end.w]! * strokeWidth;
  return {
    type: end.type,
    length,
    width,
    refX: marker.anchor === 'centre' ? length / 2 : length,
    filled: marker.filled,
    path: marker.path,
  };
}

/** What a marker adds to the shape's ink extent, in EMU beyond the endpoint. */
export function markerOvershoot(end: LineEnd | null, strokeWidth: number): number {
  const geometry = markerGeometry(end, strokeWidth);
  if (geometry === null) return 0;
  return geometry.length - geometry.refX;
}

/** SVG stroke attributes for a single-rail stroke, in EMU. */
export interface SvgStroke {
  readonly 'stroke-width': number;
  readonly 'stroke-linecap': 'butt' | 'square' | 'round';
  readonly 'stroke-linejoin': 'round' | 'bevel' | 'miter';
  readonly 'stroke-miterlimit': number;
  readonly 'stroke-dasharray': readonly number[] | null;
}

const SVG_CAP: Readonly<Record<LineCap, 'butt' | 'square' | 'round'>> = {
  flat: 'butt',
  sq: 'square',
  rnd: 'round',
};

/**
 * The stroke attributes for one rail.
 *
 * Both `stroke-linejoin` and `stroke-miterlimit` are always written, never left
 * to SVG's own defaults - which are `miter` and `4` against Office's `round` and
 * `8`, so an omitted attribute is a wrong attribute twice over.
 */
export function svgStroke(line: ResolvedLine, railWidth = line.width): SvgStroke {
  return {
    'stroke-width': railWidth,
    'stroke-linecap': SVG_CAP[line.cap],
    'stroke-linejoin': line.join.kind,
    'stroke-miterlimit': line.join.limit,
    'stroke-dasharray': line.dash === null ? null : dashArray(line.dash, railWidth, line.cap),
  };
}
