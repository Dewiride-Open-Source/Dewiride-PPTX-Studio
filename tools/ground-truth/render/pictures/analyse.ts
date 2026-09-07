/**
 * Experiment C7, step 3 - score every reading of `p:pic` against the pixels.
 *
 * ```
 * node tools/ground-truth/render/pictures/analyse.ts <work-dir>
 * ```
 *
 * Four questions - twenty-four whole readings of the container - each predicting
 * an exact colour for all 54 samples. The verdict throws unless exactly one
 * survives: a reading nothing refutes is a guess, and two readings that both fit
 * mean the probes failed to separate them.
 *
 * The outline is then read a second way, off an edge profile, because a band's
 * placement is a distance and a colour sample can only ever bracket it.
 */

import { readFileSync } from 'node:fs';
import { join } from 'node:path';

import { readBmp } from '../../lib/bmp.ts';
import { SLIDE_HEIGHT, SLIDE_WIDTH } from '../../lib/pptx.ts';
import {
  BACKGROUND,
  LINE_COLOR,
  OUTLINE_PT,
  pictureProbes,
  PT,
  quadPixel,
  QUAD_PX,
  SAMPLES,
  shapeControl,
  SIDE_PT,
  SPPR_FILL,
  type PictureProbe,
  type Rect,
} from './probes.ts';

const dir = process.argv[2];
if (dir === undefined) {
  throw new Error('usage: node tools/ground-truth/render/pictures/analyse.ts <work-dir>');
}

/* ----------------------------------------------------------- the candidates */

/** Whether a non-rect `prstGeom` clips the image or merely bounds it. */
type Clip = 'geometry' | 'bbox';
/** Which fill paints when `p:blipFill` and a `p:spPr` fill are both present. */
type Precedence = 'blip' | 'spPr';
/** Where the band of `a:ln` falls relative to the picture's own box. */
type Outline = 'outside' | 'centred' | 'inside';
/** Whether `flipH`/`flipV` mirror the image or only the outline. */
type Mirror = 'image' | 'outline';

interface Reading {
  readonly clip: Clip;
  readonly precedence: Precedence;
  readonly outline: Outline;
  readonly mirror: Mirror;
}

function readings(): readonly Reading[] {
  const out: Reading[] = [];
  for (const clip of ['geometry', 'bbox'] as const) {
    for (const precedence of ['blip', 'spPr'] as const) {
      for (const outline of ['outside', 'centred', 'inside'] as const) {
        for (const mirror of ['image', 'outline'] as const) {
          out.push({ clip, precedence, outline, mirror });
        }
      }
    }
  }
  return out;
}

function name(reading: Reading): string {
  return `${reading.clip}/${reading.precedence}/${reading.outline}/${reading.mirror}`;
}

/* ------------------------------------------------------------- the geometry */

/** Whether `(u, v)` in the unit square is inside the probe's own outline. */
function insideGeometry(prst: string, u: number, v: number): boolean {
  if (prst === 'ellipse') return (u - 0.5) ** 2 + (v - 0.5) ** 2 <= 0.25;
  // `triangle`'s default adjust puts the apex at the top centre, so its two
  // sloping edges are v = |2u - 1|.
  if (prst === 'triangle') return v >= Math.abs(2 * u - 1);
  return true;
}

/** How far into the box a reading says the outline reaches, as a fraction of it. */
function reach(outline: Outline): number {
  if (outline === 'outside') return 0;
  return (outline === 'centred' ? OUTLINE_PT / 2 : OUTLINE_PT) / SIDE_PT;
}

/* ----------------------------------------------------------- the prediction */

/** The image's own colour at a point of the source, in the unit square. */
function imageAt(u: number, v: number): number {
  const x = Math.min(QUAD_PX - 1, Math.max(0, Math.floor(u * QUAD_PX)));
  const y = Math.min(QUAD_PX - 1, Math.max(0, Math.floor(v * QUAD_PX)));
  return quadPixel(x, y);
}

/** What a reading says PowerPoint painted at one sample of one probe. */
function predict(reading: Reading, probe: PictureProbe, u: number, v: number): number {
  if (reading.clip === 'geometry' && !insideGeometry(probe.prst, u, v)) return BACKGROUND;
  if (probe.line !== undefined && Math.min(u, 1 - u, v, 1 - v) < reach(reading.outline)) {
    return LINE_COLOR;
  }
  if (probe.spPrFill !== undefined && reading.precedence === 'spPr') return SPPR_FILL;

  // A mirror is applied to the shape's own box, so the source point that a
  // device point lands on is the mirrored one.
  const su = reading.mirror === 'image' && probe.flipH ? 1 - u : u;
  const sv = reading.mirror === 'image' && probe.flipV ? 1 - v : v;
  return imageAt(su, sv);
}

/* --------------------------------------------------------------- the pixels */

function channels(rgb: number): [number, number, number] {
  return [(rgb >> 16) & 0xff, (rgb >> 8) & 0xff, rgb & 0xff];
}

function distance(rgb: readonly [number, number, number], wanted: number): number {
  const [r, g, b] = channels(wanted);
  return Math.max(Math.abs(rgb[0] - r), Math.abs(rgb[1] - g), Math.abs(rgb[2] - b));
}

interface Found {
  readonly probe: string;
  readonly sample: string;
  readonly u: number;
  readonly v: number;
  readonly rgb: [number, number, number];
}

const bitmap = readBmp(readFileSync(join(dir, 'shots', 'pictures-01-1920.bmp')));
const scale = bitmap.width / SLIDE_WIDTH;
if (Math.abs(bitmap.height / SLIDE_HEIGHT - scale) > 1e-6) {
  throw new Error('the export is not the slide aspect ratio');
}

const probes = pictureProbes();
const byId = new Map(probes.map((probe) => [probe.id, probe]));

const found: Found[] = [];
for (const probe of probes) {
  for (const sample of SAMPLES) {
    const x = Math.round((probe.rect.x + sample.u * probe.rect.cx) * scale);
    const y = Math.round((probe.rect.y + sample.v * probe.rect.cy) * scale);
    found.push({
      probe: probe.id,
      sample: sample.id,
      u: sample.u,
      v: sample.v,
      rgb: bitmap.pixel(x, y),
    });
  }
}

/* -------------------------------------------------------------- the verdict */

/** Anti-aliasing only: every sample sits in flat colour well away from a seam. */
const TOLERANCE = 8;

interface Score {
  readonly reading: Reading;
  readonly worst: number;
  readonly at: string;
}

const scores: Score[] = readings()
  .map((reading) => {
    let worst = 0;
    let at = '-';
    for (const entry of found) {
      const probe = byId.get(entry.probe);
      if (probe === undefined) throw new Error(`no probe called ${entry.probe}`);
      const d = distance(entry.rgb, predict(reading, probe, entry.u, entry.v));
      if (d > worst) {
        worst = d;
        at = `${entry.probe}/${entry.sample}`;
      }
    }
    return { reading, worst, at };
  })
  .sort((a, b) => a.worst - b.worst);

console.log(`${String(found.length)} samples over ${String(probes.length)} probes\n`);
console.log('| reading | worst channel | first refuted at |');
console.log('| --- | ---: | --- |');
for (const score of scores) {
  console.log(`| ${name(score.reading)} | ${String(score.worst)} | ${score.at} |`);
}

const fits = scores.filter((score) => score.worst <= TOLERANCE);
const best = scores[0];
if (best === undefined) throw new Error('no readings were scored');
if (fits.length === 0) {
  throw new Error(
    `no reading of p:pic fits: the closest is ${name(best.reading)} at ` +
      `${String(best.worst)} on ${best.at}`,
  );
}
if (fits.length > 1) {
  throw new Error(
    `${String(fits.length)} readings fit - the probes do not separate them: ` +
      fits.map((fit) => name(fit.reading)).join(', '),
  );
}

/* ------------------------------------------------------- the outline, twice */

/** How far a line's band reaches either side of one vertical edge, in points. */
function band(rect: Rect): { readonly outside: number; readonly inside: number } {
  const y = Math.round((rect.y + 0.3 * rect.cy) * scale);
  const left = Math.round(rect.x * scale);
  const perPoint = PT * scale;
  const isLine = (x: number): boolean => distance(bitmap.pixel(x, y), LINE_COLOR) <= TOLERANCE;
  let outside = 0;
  while (isLine(left - outside - 1)) outside++;
  let inside = 0;
  while (isLine(left + inside)) inside++;
  return { outside: outside / perPoint, inside: inside / perPoint };
}

const picture = byId.get('outline-thick');
if (picture === undefined) throw new Error('no outline-thick probe');
const pictureBand = band(picture.rect);
const shapeBand = band(shapeControl().rect);

console.log(`\n| a ${String(OUTLINE_PT)}pt outline on | outside (pt) | inside (pt) |`);
console.log('| --- | ---: | ---: |');
console.log(`| a picture | ${pictureBand.outside.toFixed(1)} | ${pictureBand.inside.toFixed(1)} |`);
console.log(`| a shape | ${shapeBand.outside.toFixed(1)} | ${shapeBand.inside.toFixed(1)} |`);

const winner = fits[0];
if (winner === undefined) throw new Error('unreachable: exactly one reading fits');
const claimed = reach(winner.reading.outline) * SIDE_PT;
if (Math.abs(pictureBand.inside - claimed) > 1) {
  throw new Error(
    `the colour samples say the outline reaches ${String(claimed)}pt into the picture, ` +
      `the profile says ${pictureBand.inside.toFixed(1)}pt`,
  );
}
// 2.8 measured a shape's default stroke as centred; if that has stopped being
// true, the picture rule below is not a contrast with anything.
if (Math.abs(shapeBand.inside - OUTLINE_PT / 2) > 1) {
  throw new Error(
    `a shape's outline is meant to be centred, but reaches ${shapeBand.inside.toFixed(1)}pt in`,
  );
}

const runnerUp = scores[1];
if (runnerUp === undefined) throw new Error('unreachable: there are 24 readings');
console.log(
  `\nexactly one reading fits: ${name(winner.reading)} at ${String(winner.worst)}; ` +
    `the next is ${name(runnerUp.reading)} at ${String(runnerUp.worst)}`,
);
