/**
 * What Gate 3 claims, as checks with no tolerance in them.
 *
 * The gate's wording is "a 100-slide deck rendered faithfully at any zoom, entirely
 * client-side". Written out: the deck has a hundred slides and every one draws at every zoom
 * the page offers; the SVG at any zoom is the SVG at 100 % apart from what the stroke and grid
 * rules own - the two stroke attributes, a crisp path's half-pixel translate and a line end's
 * numbers (F2 and F3, ADR 0054); a 2x display's 100 % markup
 * is the 200 % markup to the byte; nothing leaves the page once the
 * deck is in hand; and our own raster of every slide at every zoom is the one recorded. The
 * scores against PowerPoint are reported beside all that and gate nothing (ADR 0035).
 */

import { FidelityError } from '../fidelity/errors.ts';

/** The deck the gate is run on: a hundred slides, and the one the oracle holds at every zoom. */
export const GATE_DECK_ID = 'a46-hundred-slides';

export const GATE_SLIDES = 100;

/** The stage's zooms, 100 % first because it is what the others are compared with. */
export const STAGE_ZOOMS: readonly number[] = [1, 0.25, 2, 4];

/** The strip's thumbnails: an eighth, which is the 120-pixel export the oracle holds. */
export const STRIP_ZOOM = 0.125;

/** The width a 960-point slide is drawn at for a zoom, which is what the oracle is keyed by. */
export function widthAt(zoom: number, slideWidthPt = 960): number {
  return Math.round(zoom * slideWidthPt);
}

/** The cell that keeps a grid 120 wide at a width, which must come out whole. */
export function cellAt(width: number): number {
  const cell = (8 * width) / 960;
  if (!Number.isInteger(cell) || cell < 1) {
    throw new FidelityError(
      'FID_RASTER_GEOMETRY',
      `${String(width)} px wide reduces by ${String(cell)} px cells, which is not a raster`,
    );
  }
  return cell;
}

/** `<deck>-<NN>@<width>`: one slide at one width, the key a zoomed digest is recorded under. */
export function zoomKey(slideKey: string, width: number): string {
  return `${slideKey}@${String(width)}`;
}

/** The SVG without the root's `width` and `height`: all a display's ratio may change. */
export function maskRootSize(svg: string): string {
  if (!svg.startsWith('<svg')) {
    throw new FidelityError('FID_SVG_INTRINSIC_SIZE', 'not an <svg> root', svg.slice(0, 40));
  }
  const end = svg.indexOf('>');
  return svg.slice(0, end).replace(/ (?:width|height)="[^"]*"/g, '') + svg.slice(end);
}

/**
 * The SVG with what a zoom is allowed to change taken out.
 *
 * The root's size; `stroke-width` and `stroke-dasharray` (F2); a crisp path's half-pixel
 * translate (F3); a line end's numbers, its outline's commands kept (F2's M8). ADR 0054.
 */
export function maskZoom(svg: string): string {
  return maskRootSize(svg)
    .replace(/ (stroke-width|stroke-dasharray)="[^"]*"/g, ' $1="*"')
    .replace(/<path\b[^>]*shape-rendering="crispEdges"[^>]*>/g, (path) =>
      path.replace(/ transform="translate\([^"]*\)"/, ''),
    )
    .replace(/<marker\b[^>]*>[\s\S]*?<\/marker>/g, (marker) =>
      marker
        .replace(/ (markerWidth|markerHeight|refX|refY|stroke-width)="[^"]*"/g, ' $1="*"')
        .replace(/ d="[^"]*"/g, (d) => d.replace(/-?[\d.]+/g, '*')),
    );
}

/** The display ratio the gate switches the page to, unannounced, once the zooms are done. */
export const DISPLAY_RATIO = 2;

/** Where a 2x display's markup must be the zoom's with the same device pixels, and was not. */
export interface RatioBreak {
  readonly key: string;
  readonly what: 'stage svg' | 'strip svg';
}

/** Every `<from>-<n>` id, reference or `url(#…)` renamed to `<to>-<n>`; `thumb3` never touches `thumb30`. */
export function renameIdPrefix(svg: string, from: string, to: string): string {
  const escaped = from.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return svg.replace(new RegExp(`(["#])${escaped}-(?=\\d)`, 'g'), `$1${to}-`);
}

export interface ZoomEntry {
  readonly zoom: number;
  readonly masked: string;
}

/** The zooms whose masked SVG is not the one at 100 %; empty is invariance. */
export function invarianceOf(entries: readonly ZoomEntry[]): readonly number[] {
  const reference = entries.find((entry) => entry.zoom === 1);
  if (reference === undefined) {
    throw new FidelityError('FID_RASTER_GEOMETRY', 'no 100 % entry to compare the zooms with');
  }
  return entries
    .filter((entry) => entry.zoom !== 1 && entry.masked !== reference.masked)
    .map((entry) => entry.zoom);
}

export type RequestKind = 'static' | 'deck' | 'other';

/** What a request was for: the page and its bundle under `/dist/`, the deck, or something the gate refuses. */
export function classifyRequest(url: string, origin: string, deckPath: string): RequestKind {
  let parsed: URL;
  try {
    parsed = new URL(url);
  } catch {
    return 'other';
  }
  if (parsed.origin !== origin) return 'other';
  if (parsed.pathname === deckPath) return 'deck';
  if (parsed.pathname === '/' || parsed.pathname.startsWith('/dist/')) return 'static';
  return 'other';
}

export interface LoggedRequest {
  readonly url: string;
  /** Made after the context went offline, which nothing client-side ever does. */
  readonly afterOffline: boolean;
}

export interface RequestVerdict {
  readonly total: number;
  readonly static: number;
  readonly deck: number;
  readonly other: readonly string[];
  readonly afterOffline: readonly string[];
}

export function requestVerdict(
  log: readonly LoggedRequest[],
  origin: string,
  deckPath: string,
): RequestVerdict {
  let staticCount = 0;
  let deckCount = 0;
  const other: string[] = [];
  for (const entry of log) {
    const kind = classifyRequest(entry.url, origin, deckPath);
    if (kind === 'static') staticCount += 1;
    else if (kind === 'deck') deckCount += 1;
    else other.push(entry.url);
  }
  return {
    total: log.length,
    static: staticCount,
    deck: deckCount,
    other,
    afterOffline: log.filter((entry) => entry.afterOffline).map((entry) => entry.url),
  };
}

/** The facts the verdict is read off; every one is a count and zero is the only pass. */
export interface GateFacts {
  readonly slides: number;
  readonly notDrawn: number;
  readonly pageErrors: number;
  readonly breaks: number;
  readonly changed: number;
  readonly vanished: number;
  readonly offenders: number;
  readonly afterOffline: number;
  /** Disagreements between the 2x display and the zoom with the same device pixels. */
  readonly ratio: number;
}

export function gateHolds(facts: GateFacts): boolean {
  return (
    facts.slides >= GATE_SLIDES &&
    facts.notDrawn === 0 &&
    facts.pageErrors === 0 &&
    facts.breaks === 0 &&
    facts.changed === 0 &&
    facts.vanished === 0 &&
    facts.offenders === 0 &&
    facts.afterOffline === 0 &&
    facts.ratio === 0
  );
}

export interface Box {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

/** A page box as a screenshot clip: every edge on a whole pixel, or the raster is a resample. */
export function integerClip(box: Box, what: string): Box {
  for (const [name, value] of Object.entries(box)) {
    if (!Number.isInteger(value)) {
      throw new FidelityError(
        'FID_RASTER_GEOMETRY',
        `${what} sits at a fractional ${name} of ${String(value)} px`,
        what,
      );
    }
  }
  return box;
}
