/**
 * Experiment C6, step 3 - score the candidate models against the readings.
 *
 * ```
 * node tools/ground-truth/render/analyse.ts <dir> [--fixture corpus/ground-truth/transforms.json]
 * ```
 *
 * Nothing here asserts a sentence. Four candidate rules for the child
 * coordinate map and three for the orientation are computed from each probe's
 * declared chain and scored against what PowerPoint reported; the fixture
 * carries the scores and the raw cases, so `render-svg`'s tests re-derive the
 * rule rather than trusting a comment. If no rule fits perfectly the script
 * throws, because a fixture that records a contradiction is worse than no
 * fixture.
 */

import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

import { hex, readBmp } from '../lib/bmp.ts';
import { entry, readZip } from '../lib/zip.ts';
import type { Frame, Rect, Sample, TransformProbe } from './probes.ts';

/* -------------------------------------------------------------------------- */
/* loading                                                                    */
/* -------------------------------------------------------------------------- */

const dirArg = process.argv[2];
if (dirArg === undefined)
  throw new Error('usage: tools/ground-truth/render/analyse.ts <dir> [--fixture <path>]');
// Narrowed into its own binding: the closures below are compiled before the
// check runs, so a `string | undefined` captured from `process.argv` stays one.
const dir: string = dirArg;
const fixtureFlag = process.argv.indexOf('--fixture');
const fixturePath = fixtureFlag === -1 ? null : process.argv[fixtureFlag + 1];

/**
 * PowerShell's `ConvertTo-Json` writes a one-element array as a bare object.
 * Every list crossing that boundary goes through here, because the shape of the
 * JSON depends on how many probes happened to be in a deck.
 */
function list<T>(value: unknown): T[] {
  if (Array.isArray(value)) return value as T[];
  return value === null || value === undefined ? [] : [value as T];
}

function readJson(name: string): unknown {
  return JSON.parse(readFileSync(join(dir, name), 'utf8'));
}

interface Inputs {
  readonly slidePoints: { readonly w: number; readonly h: number };
  readonly emuPerPoint: number;
  readonly exportPixels: { readonly w: number; readonly h: number };
  readonly decks: readonly { deck: string; file: string; hostile: boolean; slides: number }[];
  readonly probes: readonly TransformProbe[];
}

interface ShapeReading {
  readonly name: string | null;
  readonly left: number | null;
  readonly top: number | null;
  readonly width: number | null;
  readonly height: number | null;
  readonly rotation: number | null;
  readonly flipH: number | null;
  readonly flipV: number | null;
  readonly type: number | null;
  readonly fillType: number | null;
  readonly fillRgb: number | null;
  readonly lineVisible: number | null;
  readonly lineWeight: number | null;
}

interface SlideReading {
  readonly index: number;
  readonly inPlace: unknown;
  readonly ungrouped: unknown;
  readonly bitmap: string | null;
}

interface DeckReading {
  readonly deck: string;
  readonly hostile: boolean;
  readonly opened: boolean;
  readonly repaired: boolean | null;
  readonly slides: unknown;
}

const inputs = readJson('transform-inputs.json') as Inputs;
const readings = readJson('transform-readings.json') as { decks: unknown };
const decks = list<DeckReading>(readings.decks);

const PX_PER_PT = inputs.exportPixels.w / inputs.slidePoints.w;

/* -------------------------------------------------------------------------- */
/* the two-by-two                                                             */
/* -------------------------------------------------------------------------- */

/** `x' = a x + c y`, `y' = b x + d y`. No translation - that is handled apart. */
interface Mat {
  readonly a: number;
  readonly b: number;
  readonly c: number;
  readonly d: number;
}

const IDENTITY: Mat = { a: 1, b: 0, c: 0, d: 1 };

/**
 * A rotation, clockwise on screen because y points down.
 *
 * Measured, and worth stating because half the confusion about `@rot` is here:
 * `rot="5400000"` is 90 degrees and turns the top edge to the right.
 */
function rotM(degrees: number): Mat {
  const t = (degrees * Math.PI) / 180;
  const cos = Math.cos(t);
  const sin = Math.sin(t);
  return { a: cos, b: sin, c: -sin, d: cos };
}

function flipM(flipH: boolean, flipV: boolean): Mat {
  return { a: flipH ? -1 : 1, b: 0, c: 0, d: flipV ? -1 : 1 };
}

function mul(m: Mat, n: Mat): Mat {
  return {
    a: m.a * n.a + m.c * n.b,
    b: m.b * n.a + m.d * n.b,
    c: m.a * n.c + m.c * n.d,
    d: m.b * n.c + m.d * n.d,
  };
}

function apply(m: Mat, x: number, y: number): { x: number; y: number } {
  return { x: m.a * x + m.c * y, y: m.b * x + m.d * y };
}

function sameMat(m: Mat, n: Mat, tol = 1e-6): boolean {
  return (
    Math.abs(m.a - n.a) < tol &&
    Math.abs(m.b - n.b) < tol &&
    Math.abs(m.c - n.c) < tol &&
    Math.abs(m.d - n.d) < tol
  );
}

/* -------------------------------------------------------------------------- */
/* the candidate models                                                       */
/* -------------------------------------------------------------------------- */

interface MapOptions {
  /** `chExt` of zero means no scaling on that axis, rather than a collapse. */
  readonly chExtZeroIsOne: boolean;
  /** A group with no `a:chOff`/`a:chExt` uses `chOff = 0` and no scaling. */
  readonly missingChildIsOrigin: boolean;
  /** `chOff` is subtracted at all. */
  readonly useChOff: boolean;
  /** The scale is applied at all. */
  readonly useScale: boolean;
  /** The turn is `rotate . flip` rather than `flip . rotate`. */
  readonly flipFirst: boolean;
}

const REFERENCE: MapOptions = {
  chExtZeroIsOne: true,
  missingChildIsOrigin: true,
  useChOff: true,
  useScale: true,
  flipFirst: true,
};

const MAP_MODELS: Readonly<Record<string, MapOptions>> = {
  // The plan's formula, with the two zero rules measured here.
  reference: REFERENCE,
  // `chExt="0"` read as a divisor rather than as "no child scaling".
  chExtZeroCollapses: { ...REFERENCE, chExtZeroIsOne: false },
  // The most common group bug: the offset without the scale factor.
  noScale: { ...REFERENCE, useScale: false },
  // The second most common: forgetting that the children are in their own space.
  noChOff: { ...REFERENCE, useChOff: false },
  // A group with no child transform read as `chOff = off`, `chExt = ext`.
  missingChildIsFrame: { ...REFERENCE, missingChildIsOrigin: false },
  // Rotate, then mirror.
  rotBeforeFlip: { ...REFERENCE, flipFirst: false },
};

function childSpace(
  frame: Frame,
  opts: MapOptions,
): { sx: number; sy: number; ox: number; oy: number } {
  if (frame.noChildTransform === true && opts.missingChildIsOrigin) {
    return { sx: 1, sy: 1, ox: 0, oy: 0 };
  }
  const zero = opts.chExtZeroIsOne ? 1 : 0;
  const sx = !opts.useScale ? 1 : frame.child.w === 0 ? zero : frame.rect.w / frame.child.w;
  const sy = !opts.useScale ? 1 : frame.child.h === 0 ? zero : frame.rect.h / frame.child.h;
  return {
    sx,
    sy,
    ox: opts.useChOff ? frame.child.x : 0,
    oy: opts.useChOff ? frame.child.y : 0,
  };
}

/** One rectangle, from a frame's child space into the frame's parent's. */
function mapUp(frame: Frame, r: Rect, opts: MapOptions): Rect {
  const { sx, sy, ox, oy } = childSpace(frame, opts);
  const x = frame.rect.x + (r.x - ox) * sx;
  const y = frame.rect.y + (r.y - oy) * sy;
  const w = r.w * sx;
  const h = r.h * sy;

  const fcx = frame.rect.x + frame.rect.w / 2;
  const fcy = frame.rect.y + frame.rect.h / 2;
  const turn = opts.flipFirst
    ? mul(rotM(frame.rot ?? 0), flipM(frame.flipH === true, frame.flipV === true))
    : mul(flipM(frame.flipH === true, frame.flipV === true), rotM(frame.rot ?? 0));
  const v = apply(turn, x + w / 2 - fcx, y + h / 2 - fcy);
  return { x: fcx + v.x - w / 2, y: fcy + v.y - h / 2, w, h };
}

/** The composed frame a leaf occupies on the slide, under one model. */
function composeFrame(probe: TransformProbe, opts: MapOptions): Rect {
  let r: Rect = probe.leaf;
  for (let i = probe.chain.length - 1; i >= 0; i--) r = mapUp(probe.chain[i]!, r, opts);
  return r;
}

type TurnModel = 'rotAfterFlip' | 'flipAfterRot' | 'sumOfAngles';

function factor(rot: number, flipH: boolean, flipV: boolean, model: TurnModel): Mat {
  const R = rotM(rot);
  const F = flipM(flipH, flipV);
  return model === 'flipAfterRot' ? mul(F, R) : mul(R, F);
}

/** The composed orientation of a leaf, under one model. */
function composeTurn(probe: TransformProbe, model: TurnModel): Mat {
  if (model === 'sumOfAngles') {
    let angle = probe.leafRot;
    let flipH = probe.leafFlipH;
    let flipV = probe.leafFlipV;
    for (const frame of probe.chain) {
      angle += frame.rot ?? 0;
      if (frame.flipH === true) flipH = !flipH;
      if (frame.flipV === true) flipV = !flipV;
    }
    return factor(angle, flipH, flipV, 'rotAfterFlip');
  }
  let m = IDENTITY;
  for (const frame of probe.chain) {
    m = mul(m, factor(frame.rot ?? 0, frame.flipH === true, frame.flipV === true, model));
  }
  return mul(m, factor(probe.leafRot, probe.leafFlipH, probe.leafFlipV, model));
}

/* -------------------------------------------------------------------------- */
/* the readings, per probe                                                    */
/* -------------------------------------------------------------------------- */

interface Case {
  readonly probe: TransformProbe;
  readonly repaired: boolean;
  readonly inPlace: ShapeReading | null;
  readonly ungrouped: ShapeReading | null;
}

const cases: Case[] = [];
for (const probe of inputs.probes) {
  const deck = decks.find((d) => d.deck === probe.deck);
  if (deck === undefined) throw new Error(`no readings for deck ${probe.deck}`);
  const slide = list<SlideReading>(deck.slides)[probe.slide - 1];
  if (slide === undefined) throw new Error(`no readings for ${probe.id}`);
  cases.push({
    probe,
    repaired: deck.repaired === true,
    inPlace: list<ShapeReading>(slide.inPlace).find((s) => s.name === probe.shape) ?? null,
    ungrouped: list<ShapeReading>(slide.ungrouped).find((s) => s.name === probe.shape) ?? null,
  });
}

const missing = cases.filter((c) => c.inPlace === null);
if (missing.length > 0) {
  throw new Error(`PowerPoint did not report: ${missing.map((c) => c.probe.id).join(', ')}`);
}

/* -------------------------------------------------------------------------- */
/* control                                                                    */
/* -------------------------------------------------------------------------- */

const bitmaps = new Map<string, ReturnType<typeof readBmp>>();
function bitmapFor(deckName: string, slide: number): ReturnType<typeof readBmp> {
  const key = `${deckName}-s${String(slide).padStart(2, '0')}`;
  const cached = bitmaps.get(key);
  if (cached !== undefined) return cached;
  const bmp = readBmp(new Uint8Array(readFileSync(join(dir, `${key}.bmp`))));
  bitmaps.set(key, bmp);
  return bmp;
}

interface SampleReading {
  readonly id: string;
  readonly probe: string;
  readonly x: number;
  readonly y: number;
  readonly asks: string;
  readonly expect: string | null;
  /** The pixel itself. */
  readonly rgb: string;
  /** The most common colour of the 5x5 block around it, to survive an edge. */
  readonly block: string;
  /**
   * The fraction of a 21x21 block that is not white.
   *
   * A single pixel cannot answer "is a pattern painted here" - `pct50` covers
   * half the pixels and a line pattern one in eight - so ink density answers it
   * instead, without also having to solve for the tile phase.
   */
  readonly ink: number;
  /**
   * The mean colour of a 21x21 block, as hex.
   *
   * The comparison against our own renderer runs on this rather than on the
   * pixel: a `pct50` tile is half ink and half ground, so "the colour here" is
   * a coin flip at any one pixel and a stable number over a block.
   */
  readonly mean: string;
}

function readSample(probe: TransformProbe, sample: Sample): SampleReading {
  const bmp = bitmapFor(probe.deck, probe.slide);
  const px = Math.round(sample.x * PX_PER_PT);
  const py = Math.round(sample.y * PX_PER_PT);
  const counts = new Map<string, number>();
  for (let dy = -2; dy <= 2; dy++) {
    for (let dx = -2; dx <= 2; dx++) {
      const h = hex(bmp.pixel(px + dx, py + dy));
      counts.set(h, (counts.get(h) ?? 0) + 1);
    }
  }
  let block = '';
  let best = -1;
  for (const [h, n] of counts) {
    if (n > best) {
      best = n;
      block = h;
    }
  }
  let inked = 0;
  let total = 0;
  let red = 0;
  let green = 0;
  let blue = 0;
  for (let dy = -10; dy <= 10; dy++) {
    for (let dx = -10; dx <= 10; dx++) {
      total++;
      const pixel = bmp.pixel(px + dx, py + dy);
      red += pixel[0];
      green += pixel[1];
      blue += pixel[2];
      if (hex(pixel) !== 'FFFFFF') inked++;
    }
  }
  return {
    id: sample.id,
    probe: probe.id,
    x: sample.x,
    y: sample.y,
    asks: sample.asks,
    expect: sample.expect ?? null,
    rgb: hex(bmp.pixel(px, py)),
    block,
    ink: Math.round((inked / total) * 1000) / 1000,
    mean: hex([Math.round(red / total), Math.round(green / total), Math.round(blue / total)]),
  };
}

const samples: SampleReading[] = [];
for (const probe of inputs.probes) {
  for (const sample of probe.samples ?? []) samples.push(readSample(probe, sample));
}

const sampleById = new Map(samples.map((s) => [s.id, s]));
function sampleColor(id: string): string {
  const found = sampleById.get(id);
  if (found === undefined) throw new Error(`no sample ${id}`);
  return found.block;
}

/* -------------------------------------------------------------------------- */
/* turning a colour back into a position                                      */
/* -------------------------------------------------------------------------- */

function colorDistance(a: string, b: string): number {
  let sum = 0;
  for (let i = 0; i < 3; i++) {
    const d = parseInt(a.slice(i * 2, i * 2 + 2), 16) - parseInt(b.slice(i * 2, i * 2 + 2), 16);
    sum += d * d;
  }
  return sum;
}

/**
 * Where along a reference ramp a sampled colour sits, as a fraction.
 *
 * The reference is an ordinary shape carrying the identical gradient over a
 * known rectangle, drawn behind everything on the same slide. Scanning it turns
 * "this pixel is FB0037" into "this is a tenth of the way along", which is the
 * only form in which the question - over what rectangle is a group's fill laid
 * out - can be answered.
 */
interface RampReading {
  readonly id: string;
  readonly probe: string;
  /** Where on the slide the sample was taken, in points. */
  readonly at: number;
  readonly rgb: string;
  /** The fraction along the reference ramp, or `null` if nothing matched. */
  readonly t: number | null;
  /** The reference position that colour occurs at, in points. */
  readonly refAt: number | null;
}

const ramps: RampReading[] = [];
for (const probe of inputs.probes) {
  const ref = probe.reference;
  if (ref === undefined) continue;
  const bmp = bitmapFor(probe.deck, probe.slide);
  const scan: { at: number; rgb: string }[] = [];
  const steps = Math.round((ref.x1 - ref.x0) * PX_PER_PT);
  for (let i = 0; i <= steps; i++) {
    const at = ref.x0 + i / PX_PER_PT;
    const px = ref.vertical === true ? Math.round(ref.y * PX_PER_PT) : Math.round(at * PX_PER_PT);
    const py = ref.vertical === true ? Math.round(at * PX_PER_PT) : Math.round(ref.y * PX_PER_PT);
    if (px < 0 || py < 0 || px >= bmp.width || py >= bmp.height) continue;
    scan.push({ at, rgb: hex(bmp.pixel(px, py)) });
  }
  for (const sample of probe.samples ?? []) {
    const reading = sampleById.get(sample.id);
    if (reading === undefined) continue;
    let best: { at: number; rgb: string } | null = null;
    for (const point of scan) {
      if (
        best === null ||
        colorDistance(point.rgb, reading.block) < colorDistance(best.rgb, reading.block)
      ) {
        best = point;
      }
    }
    const matched = best !== null && colorDistance(best.rgb, reading.block) <= 12;
    ramps.push({
      id: sample.id,
      probe: probe.id,
      at: ref.vertical === true ? sample.y : sample.x,
      rgb: reading.block,
      t: matched ? (best!.at - ref.x0) / (ref.x1 - ref.x0) : null,
      refAt: matched ? best!.at : null,
    });
  }
}

/**
 * Two samples of one ramp give the rectangle it was laid out over.
 *
 * `t = (p - x0) / (x1 - x0)` at two points is two equations in two unknowns,
 * so a pair of readings names the span outright rather than testing a guess.
 */
function spanFrom(a: RampReading, b: RampReading): { x0: number; x1: number } | null {
  if (a.t === null || b.t === null) return null;
  const dt = b.t - a.t;
  if (Math.abs(dt) < 1e-6) return null;
  const width = (b.at - a.at) / dt;
  const x0 = a.at - a.t * width;
  return { x0: Math.round(x0 * 100) / 100, x1: Math.round((x0 + width) * 100) / 100 };
}

const rampById = new Map(ramps.map((r) => [r.id, r]));
function span(a: string, b: string): { x0: number; x1: number } | null {
  const ra = rampById.get(a);
  const rb = rampById.get(b);
  return ra === undefined || rb === undefined ? null : spanFrom(ra, rb);
}

// Read the controls first. C's write-up says so and it cost that experiment a
// day: a deck can open cleanly, report plausible numbers, and be measuring
// something else entirely.
for (const sample of samples) {
  if (sample.expect !== null && sample.block !== sample.expect) {
    throw new Error(
      `control ${sample.id} expected ${sample.expect} and read ${sample.block} - ` +
        'the sampling or the deck is wrong, and nothing else here can be trusted',
    );
  }
}

/* -------------------------------------------------------------------------- */
/* scoring                                                                    */
/* -------------------------------------------------------------------------- */

const TOL_PT = 0.01;

function sameRect(a: Rect, b: Rect): boolean {
  return (
    Math.abs(a.x - b.x) < TOL_PT &&
    Math.abs(a.y - b.y) < TOL_PT &&
    Math.abs(a.w - b.w) < TOL_PT &&
    Math.abs(a.h - b.h) < TOL_PT
  );
}

function readingRect(r: ShapeReading): Rect {
  return { x: r.left ?? NaN, y: r.top ?? NaN, w: r.width ?? NaN, h: r.height ?? NaN };
}

/**
 * Every probe but one, and the exception is named rather than dropped.
 *
 * `t-shear` puts a rotated child inside a group that scales the two axes
 * differently. The true image of a rotated rectangle under that map is a
 * parallelogram, and no `a:xfrm` can hold one, so PowerPoint has to approximate
 * and the frame it reports is not the frame any of these models computes. It is
 * a finding of its own, below, and scoring it here would only make every model
 * look one worse.
 */
const SHEAR_PROBE = 't-shear';
const isShear = (c: Case): boolean => c.probe.id === SHEAR_PROBE || c.probe.deck === 'shear';

const frameCases = cases.filter((c) => !c.repaired && !isShear(c));
const turnCases = cases.filter((c) => !c.repaired && c.probe.kind === 'turn' && !isShear(c));

interface Score {
  readonly model: string;
  readonly right: number;
  readonly of: number;
  readonly misses: readonly string[];
}

const mapScores: Score[] = Object.entries(MAP_MODELS).map(([model, opts]) => {
  const misses: string[] = [];
  for (const c of frameCases) {
    if (!sameRect(composeFrame(c.probe, opts), readingRect(c.inPlace!))) misses.push(c.probe.id);
  }
  return { model, right: frameCases.length - misses.length, of: frameCases.length, misses };
});

const turnScores: Score[] = (['rotAfterFlip', 'flipAfterRot', 'sumOfAngles'] as const).map(
  (model) => {
    const misses: string[] = [];
    for (const c of turnCases) {
      const reading = c.ungrouped ?? c.inPlace!;
      const reported = factor(
        reading.rotation ?? 0,
        (reading.flipH ?? 0) !== 0,
        (reading.flipV ?? 0) !== 0,
        model === 'sumOfAngles' ? 'rotAfterFlip' : model,
      );
      if (!sameMat(composeTurn(c.probe, model), reported)) misses.push(c.probe.id);
    }
    return { model, right: turnCases.length - misses.length, of: turnCases.length, misses };
  },
);

function winner(scores: readonly Score[], what: string): Score {
  const perfect = scores.filter((s) => s.right === s.of);
  if (perfect.length === 0) {
    const table = scores.map((s) => `${s.model} ${String(s.right)}/${String(s.of)}`).join(', ');
    throw new Error(`no ${what} model fits every case (${table}) - the experiment is inconclusive`);
  }
  return perfect[0]!;
}

const mapWinner = winner(mapScores, 'child-map');
const turnWinner = winner(turnScores, 'orientation');

/* -------------------------------------------------------------------------- */
/* the questions the object model cannot answer                               */
/* -------------------------------------------------------------------------- */

/** Two colours that are the same to within a couple of levels of quantisation. */
function near(a: string, b: string, tol = 6): boolean {
  for (let i = 0; i < 3; i++) {
    const av = parseInt(a.slice(i * 2, i * 2 + 2), 16);
    const bv = parseInt(b.slice(i * 2, i * 2 + 2), 16);
    if (Math.abs(av - bv) > tol) return false;
  }
  return true;
}

const WHITE = 'FFFFFF';

const gradientLeft = sampleColor('f-grad-left');
const gradientRight = sampleColor('f-grad-right');
const groupLeft = sampleColor('f-grad-ref-left');
const groupRight = sampleColor('f-grad-ref-right');

/**
 * What a group actually contains, in slide points.
 *
 * The union of every leaf shape's composed frame. Groups themselves contribute
 * nothing of their own, which is the point: a group's declared rectangle turns
 * out not to be what its fill is laid out over.
 */
function contentBounds(probe: TransformProbe): { x0: number; x1: number; y0: number; y1: number } {
  const rects: Rect[] = [composeFrame(probe, REFERENCE)];
  for (const sibling of probe.siblings ?? []) {
    rects.push(composeFrame({ ...probe, leaf: sibling.rect }, REFERENCE));
  }
  return {
    x0: Math.min(...rects.map((r) => r.x)),
    x1: Math.max(...rects.map((r) => r.x + r.w)),
    y0: Math.min(...rects.map((r) => r.y)),
    y1: Math.max(...rects.map((r) => r.y + r.h)),
  };
}

/** The declared rectangle of the outermost group, on the axis being read. */
function declaredBox(probe: TransformProbe, vertical: boolean): { x0: number; x1: number } | null {
  const frame = probe.chain[0];
  if (frame === undefined) return null;
  return vertical
    ? { x0: frame.rect.y, x1: frame.rect.y + frame.rect.h }
    : { x0: frame.rect.x, x1: frame.rect.x + frame.rect.w };
}

/**
 * The four probes that can name the rectangle outright.
 *
 * Two samples of one ramp solve for its span, so each of these pairs a reading
 * near one end with a reading near the other. The pairs are named rather than
 * taken from the ends of the sample list, because two of these slides also
 * sample the reference shape and a reference reading would solve for the
 * reference's own rectangle - which is a different, and trivially true, answer.
 */
const SPAN_PAIRS: readonly (readonly [string, string, string])[] = [
  ['f-gradient', 'f-grad-left', 'f-grad-right'],
  ['g-one', 'g-one-a', 'g-one-c'],
  ['g-deep', 'g-deep-a', 'g-deep-c'],
  ['g-vert', 'g-vert-top', 'g-vert-bot'],
];

const spanChecks = SPAN_PAIRS.map(([id, first, last]) => {
  const probe = inputs.probes.find((p) => p.id === id)!;
  const bounds = contentBounds(probe);
  const vertical = probe.reference?.vertical === true;
  const measured = span(first, last);
  const content = vertical ? { x0: bounds.y0, x1: bounds.y1 } : { x0: bounds.x0, x1: bounds.x1 };
  const declared = declaredBox(probe, vertical);
  const matches = (box: { x0: number; x1: number } | null): boolean =>
    measured !== null &&
    box !== null &&
    Math.abs(measured.x0 - box.x0) < 2 &&
    Math.abs(measured.x1 - box.x1) < 2;
  return {
    probe: id,
    axis: vertical ? 'y' : 'x',
    measured,
    contentBounds: content,
    declaredBox: declared,
    matchesContent: matches(content),
    matchesDeclared: matches(declared),
  };
});

const groupFill = {
  /** A `grpFill` child over a solid group fill reads that colour. */
  solidReachesGroup: near(sampleColor('f-solid-leaf'), 'E53935'),
  /** Is a group's own fill painted where no child covers it? */
  groundIsPainted: !near(sampleColor('f-solid-bare'), WHITE),
  groundIsPaintedAlone: !near(sampleColor('f-alone-inside'), WHITE),
  gradientLeft,
  gradientRight,
  groupLeft,
  groupRight,
  /**
   * A slice, or a copy?
   *
   * Two congruent children at opposite ends of the same group. A copy gives
   * each of them its own ramp and so the same colour at the same relative
   * position; a slice gives them different colours. Which rectangle the slice
   * runs over is `spanChecks`, and it is not the one anybody would guess.
   */
  gradientIsSlice: !near(gradientLeft, gradientRight),
  spanChecks,
  /**
   * The rectangle the group's ramp is laid out over, solved from two samples.
   *
   * The group in every one of these declares `off=(100,100) ext=(600,300)`, so
   * a span of 100..700 would mean "the group's own rectangle" and anything else
   * means something else. What comes back is the bounding box of what the group
   * actually contains, which is not the same thing and is not written anywhere.
   */
  extents: {
    twoChildren: span('f-grad-left', 'f-grad-right'),
    oneChild: span('g-one-a', 'g-one-c'),
    deep: span('g-deep-a', 'g-deep-c'),
    vertical: span('g-vert-top', 'g-vert-bot'),
    /**
     * The one-sample probes cannot solve a span, so they report the fraction
     * they landed at. `g-sib-a` and `g-out-a` are the same grpFill child at
     * x 300..400 in a group at 100..700; the only difference is a sibling that
     * does not use `grpFill`, once inside the group rectangle and once well
     * outside it. If the fraction moves, siblings count.
     */
    solidSiblingT: rampById.get('g-sib-a')?.t ?? null,
    outsideSiblingT: rampById.get('g-out-a')?.t ?? null,
    onlyChildT: rampById.get('g-one-b')?.t ?? null,
  },
  ramps,
  /** Ink density, because a single pixel cannot see a half-tone tile. */
  patternLeafInk: sampleById.get('f-patt-leaf')?.ink ?? null,
  patternGroundInk: sampleById.get('f-patt-group')?.ink ?? null,
  patternReferenceInk: sampleById.get('f-patt-ref')?.ink ?? null,
  /** `a:grpFill` over an explicit `a:noFill`, and over a group that says nothing. */
  overNoFill: sampleColor('f-nofill-leaf'),
  overSilence: sampleColor('f-silent-leaf'),
  /** `a:grpFill` on a shape that is in no group at all. */
  atTopLevel: sampleColor('f-top-leaf'),
  /** A `grpFill` inside a group that itself says `grpFill`. */
  nested: sampleColor('f-nested-leaf'),
  nestedReachesOuter: near(sampleColor('f-nested-leaf'), groupLeft),
};

/**
 * What a slide draws that it did not write, and what `@showMasterSp` hides.
 *
 * `Slide.Shapes` lists only the slide's own, so the inherited furniture is on
 * the screen without being in the collection and the bitmap is the only witness.
 * Three squares in three colours, one per sheet, and two placeholders nothing on
 * the slide matches.
 */
const compositing = [
  'inherit-default',
  'inherit-slide-off',
  'inherit-layout-off',
  'inherit-both-off',
].map((deck) => ({
  deck,
  masterShape: sampleColor(`${deck}-master`),
  layoutShape: sampleColor(`${deck}-layout`),
  slideShape: sampleColor(`${deck}-slide`),
  masterPlaceholder: sampleColor(`${deck}-master-ph`),
  layoutPlaceholder: sampleColor(`${deck}-layout-ph`),
  masterVisible: sampleColor(`${deck}-master`) !== WHITE,
  layoutVisible: sampleColor(`${deck}-layout`) !== WHITE,
  slideVisible: sampleColor(`${deck}-slide`) !== WHITE,
  placeholdersVisible:
    sampleColor(`${deck}-master-ph`) !== WHITE || sampleColor(`${deck}-layout-ph`) !== WHITE,
}));

const widthCase = cases.find((c) => c.probe.id === 't-width');
const strokeScale = {
  /** `Line.Weight` in points, after the group has been composed away. */
  weightAfterUngroup: widthCase?.ungrouped?.lineWeight ?? null,
  declaredWidth: widthCase?.probe.lineWidth ?? null,
  groupScale: 2,
  /** A pixel that only a doubled stroke would cover. */
  scaledRow: sampleColor('t-width-scaled'),
  strokeRow: sampleColor('t-width-stroke'),
  strokeIsScaled: !near(sampleColor('t-width-scaled'), WHITE),
};

/**
 * A rotated child in a group that scales the axes differently.
 *
 * `sx` and `sy` here are the group's, `baked` is the rectangle "scale it and
 * keep the angle" produces, and `reported` is what PowerPoint's object model
 * says - which the bitmap confirms is also what it paints. The rule is left for
 * the reader of the fixture to state, because one angle produced a swap and a
 * swap is not a rule.
 */
const shearCases = cases.filter(isShear).map((c) => {
  const frame = c.probe.chain[0]!;
  const sx = frame.child.w === 0 ? 1 : frame.rect.w / frame.child.w;
  const sy = frame.child.h === 0 ? 1 : frame.rect.h / frame.child.h;
  const baked = composeFrame(c.probe, REFERENCE);
  const reported = readingRect(c.inPlace!);
  return {
    id: c.probe.id,
    rot: c.probe.leafRot,
    flipH: c.probe.leafFlipH,
    sx,
    sy,
    leaf: c.probe.leaf,
    baked,
    reported,
    reportedRotation: (c.ungrouped ?? c.inPlace!).rotation,
    reportedFlipH: ((c.ungrouped ?? c.inPlace!).flipH ?? 0) !== 0,
    reportedFlipV: ((c.ungrouped ?? c.inPlace!).flipV ?? 0) !== 0,
    /** Did the two extents come out swapped relative to the baked rectangle? */
    swapped: Math.abs(reported.w - baked.h) < TOL_PT && Math.abs(reported.h - baked.w) < TOL_PT,
    /** Or unswapped, which is the case whenever there is nothing to approximate? */
    baked_: Math.abs(reported.w - baked.w) < TOL_PT && Math.abs(reported.h - baked.h) < TOL_PT,
    /** The centre is the naive one either way; this records that it is. */
    centreAgrees:
      Math.abs(reported.x + reported.w / 2 - (baked.x + baked.w / 2)) < TOL_PT &&
      Math.abs(reported.y + reported.h / 2 - (baked.y + baked.h / 2)) < TOL_PT,
  };
});

/**
 * Three readings of when the two scale factors change places.
 *
 * The extents come out swapped for exactly those angles whose nearest quadrant
 * is a quarter or three quarters of a turn. Rounding is what separates the
 * candidates: at 45 degrees the swap happens and at 135 it does not, so the tie
 * goes upward in both cases and a comparison of `|sin|` against `|cos|` - the
 * form anyone would write first - is wrong at one end or the other.
 */
const SWAP_RULES: Readonly<Record<string, (rot: number) => boolean>> = {
  quadrantRoundHalfUp: (rot) => Math.floor(((((rot % 360) + 360) % 360) + 45) / 90) % 2 === 1,
  sinStrictlyGreater: (rot) =>
    Math.abs(Math.sin((rot * Math.PI) / 180)) > Math.abs(Math.cos((rot * Math.PI) / 180)),
  sinAtLeast: (rot) =>
    Math.abs(Math.sin((rot * Math.PI) / 180)) >= Math.abs(Math.cos((rot * Math.PI) / 180)) - 1e-9,
};

const swapScores: Score[] = Object.entries(SWAP_RULES).map(([model, rule]) => {
  const misses = shearCases.filter((c) => rule(c.rot) !== c.swapped).map((c) => c.id);
  return { model, right: shearCases.length - misses.length, of: shearCases.length, misses };
});

const swapWinner = winner(swapScores, 'extent-swap');

const shear = {
  cases: shearCases,
  swapRule: { winner: swapWinner.model, scores: swapScores },
  /** Every case's centre lands where scaling the rectangle would put it. */
  centreAlwaysAgrees: shearCases.every((c) => c.centreAgrees),
  /** And the angle is carried through untouched. */
  rotationAlwaysKept: shearCases.every((c) => (c.reportedRotation ?? 0) === c.rot),
  /** The four bitmap samples of the original 45-degree case. */
  paintsTheReportedFrame: [
    sampleColor('t-shear-left') === WHITE,
    sampleColor('t-shear-right') === WHITE,
    sampleColor('t-shear-top') === WHITE,
    sampleColor('t-shear-centre') !== WHITE,
  ],
  samples: samples.filter((s) => s.probe === SHEAR_PROBE),
};

/* -------------------------------------------------------------------------- */
/* what PowerPoint authored                                                   */
/* -------------------------------------------------------------------------- */

interface AuthoredXfrm {
  readonly name: string;
  readonly rot: number | null;
  readonly flipH: boolean;
  readonly flipV: boolean;
  readonly off: { x: number; y: number } | null;
  readonly ext: { cx: number; cy: number } | null;
  readonly chOff: { x: number; y: number } | null;
  readonly chExt: { cx: number; cy: number } | null;
}

/**
 * Names and `a:xfrm`s out of a slide part, in document order.
 *
 * A lexical scan rather than a parse, for the reason `tools/corpus/lexical/inventory.ts`
 * gives: `tools/` does not link the workspace packages, and this reads two
 * files PowerPoint wrote on this machine ten minutes ago.
 */
function authoredSlide(archive: Uint8Array, slide: number): AuthoredXfrm[] {
  const bytes = entry(readZip(archive), `ppt/slides/slide${String(slide)}.xml`);
  if (bytes === undefined) throw new Error(`no slide${String(slide)} in the authored deck`);
  const xml = new TextDecoder().decode(bytes);
  const out: AuthoredXfrm[] = [];
  const names = /<p:cNvPr id="\d+" name="([^"]*)"/g;
  let match: RegExpExecArray | null;
  const starts: { name: string; at: number }[] = [];
  while ((match = names.exec(xml)) !== null) starts.push({ name: match[1]!, at: match.index });
  for (const { name, at } of starts) {
    if (name === '') continue;
    const open = xml.indexOf('<a:xfrm', at);
    if (open === -1) continue;
    const close = xml.indexOf('</a:xfrm>', open);
    const block = xml.slice(open, close === -1 ? open : close);
    const num = (re: RegExp): number | null => {
      const m = re.exec(block);
      return m === null ? null : Number(m[1]);
    };
    const offX = num(/<a:off x="(-?\d+)"/);
    const offY = num(/<a:off x="-?\d+" y="(-?\d+)"/);
    const extCx = num(/<a:ext cx="(-?\d+)"/);
    const extCy = num(/<a:ext cx="-?\d+" cy="(-?\d+)"/);
    const chX = num(/<a:chOff x="(-?\d+)"/);
    const chY = num(/<a:chOff x="-?\d+" y="(-?\d+)"/);
    const chCx = num(/<a:chExt cx="(-?\d+)"/);
    const chCy = num(/<a:chExt cx="-?\d+" cy="(-?\d+)"/);
    out.push({
      name,
      rot: num(/<a:xfrm[^>]*\brot="(-?\d+)"/),
      flipH: /<a:xfrm[^>]*\bflipH="1"/.test(block),
      flipV: /<a:xfrm[^>]*\bflipV="1"/.test(block),
      off: offX === null || offY === null ? null : { x: offX, y: offY },
      ext: extCx === null || extCy === null ? null : { cx: extCx, cy: extCy },
      chOff: chX === null || chY === null ? null : { x: chX, y: chY },
      chExt: chCx === null || chCy === null ? null : { cx: chCx, cy: chCy },
    });
  }
  return out;
}

interface Authored {
  readonly flipCases: readonly {
    name: string;
    angle: number;
    op: string;
    writtenRot: number | null;
    flipH: boolean;
    flipV: boolean;
  }[];
  readonly fresh: readonly AuthoredXfrm[];
  readonly resized: readonly AuthoredXfrm[];
  readonly turned: readonly AuthoredXfrm[];
  readonly turnedUngrouped: readonly AuthoredXfrm[];
  readonly nested: readonly AuthoredXfrm[];
  readonly nestedUngrouped: readonly AuthoredXfrm[];
}

let authored: Authored | null = null;
try {
  const groups = new Uint8Array(readFileSync(join(dir, 'pp-groups.pptx')));
  const ungrouped = new Uint8Array(readFileSync(join(dir, 'pp-ungrouped.pptx')));
  const slide1 = authoredSlide(groups, 1);
  authored = {
    flipCases: slide1
      .filter((s) => s.name.startsWith('flip-'))
      .map((s) => {
        const [, angle, op] = s.name.split('-');
        return {
          name: s.name,
          angle: Number(angle),
          op: op ?? '',
          writtenRot: s.rot,
          flipH: s.flipH,
          flipV: s.flipV,
        };
      }),
    fresh: authoredSlide(groups, 2),
    resized: authoredSlide(groups, 3),
    turned: authoredSlide(groups, 4),
    turnedUngrouped: authoredSlide(ungrouped, 4),
    nested: authoredSlide(groups, 6),
    nestedUngrouped: authoredSlide(ungrouped, 6),
  };
} catch {
  console.log(
    '  (no authored decks in the directory - run tools/ground-truth/render/author.ps1 for those)',
  );
}

/**
 * The flip cases, restated as the question they answer.
 *
 * PowerPoint was handed a shape at `angle` and asked to mirror it. Under
 * `rotate . flip` it must write `-angle`; under `flip . rotate` it must write
 * `+angle`. Both flips at once is a half turn, which commutes, so those cases
 * must write `+angle` under either.
 */
const flipVerdict =
  authored === null
    ? null
    : authored.flipCases
        .filter((c) => c.op === 'H' || c.op === 'V')
        .filter((c) => c.angle !== 0)
        .map((c) => ({
          name: c.name,
          angle: c.angle,
          written: c.writtenRot === null ? 0 : c.writtenRot / 60000,
          negated: ((c.writtenRot ?? 0) / 60000 + c.angle) % 360 === 0,
        }));

/* -------------------------------------------------------------------------- */
/* the fixture                                                                */
/* -------------------------------------------------------------------------- */

const fixture = {
  experiment: 'C6 - group transforms, flip order and a:grpFill',
  generatedBy: 'tools/ground-truth/render/analyse.ts',
  slidePoints: inputs.slidePoints,
  emuPerPoint: inputs.emuPerPoint,
  exportPixels: inputs.exportPixels,
  decks: decks.map((d) => ({
    deck: d.deck,
    hostile: d.hostile,
    opened: d.opened,
    repaired: d.repaired,
  })),
  probes: cases.map((c) => ({
    id: c.probe.id,
    deck: c.probe.deck,
    slide: c.probe.slide,
    kind: c.probe.kind,
    question: c.probe.question,
    fromRepairedDeck: c.repaired,
    chain: c.probe.chain,
    leaf: c.probe.leaf,
    leafRot: c.probe.leafRot,
    leafFlipH: c.probe.leafFlipH,
    leafFlipV: c.probe.leafFlipV,
    reading: {
      frame: readingRect(c.inPlace!),
      rotation: (c.ungrouped ?? c.inPlace!).rotation,
      flipH: ((c.ungrouped ?? c.inPlace!).flipH ?? 0) !== 0,
      flipV: ((c.ungrouped ?? c.inPlace!).flipV ?? 0) !== 0,
      lineWeight: c.ungrouped?.lineWeight ?? null,
    },
  })),
  samples,
  authored,
  findings: {
    childMap: { winner: mapWinner.model, scores: mapScores, options: MAP_MODELS[mapWinner.model] },
    orientation: { winner: turnWinner.model, scores: turnScores },
    flipVerdict,
    groupFill,
    compositing,
    strokeScale,
    shear,
    chExtZero: {
      probe: 'h-chext-zero-x',
      reading: readingRect(cases.find((c) => c.probe.id === 'h-chext-zero-x')!.inPlace!),
      meansNoScaling: true,
    },
    extZero: {
      probe: 'h-ext-zero-x',
      reading: readingRect(cases.find((c) => c.probe.id === 'h-ext-zero-x')!.inPlace!),
    },
    missingChildTransform: {
      probe: 'h-no-chxfrm',
      reading: readingRect(cases.find((c) => c.probe.id === 'h-no-chxfrm')!.inPlace!),
    },
    negativeChExt: {
      probe: 'h-chext-negative',
      repaired: decks.find((d) => d.deck === 'h-chext-negative')?.repaired ?? null,
    },
  },
};

if (fixturePath !== undefined && fixturePath !== null) {
  writeFileSync(fixturePath, `${JSON.stringify(fixture, null, 2)}\n`);
  console.log(`wrote ${fixturePath}`);
}

/* -------------------------------------------------------------------------- */
/* the report                                                                 */
/* -------------------------------------------------------------------------- */

console.log('');
console.log('child coordinate map');
for (const s of mapScores) {
  console.log(
    `  ${s.model.padEnd(22)} ${String(s.right).padStart(2)}/${String(s.of)}` +
      (s.misses.length > 0 ? `   misses ${s.misses.slice(0, 4).join(' ')}` : ''),
  );
}
console.log('');
console.log('orientation composition');
for (const s of turnScores) {
  console.log(
    `  ${s.model.padEnd(22)} ${String(s.right).padStart(2)}/${String(s.of)}` +
      (s.misses.length > 0 ? `   misses ${s.misses.slice(0, 4).join(' ')}` : ''),
  );
}
console.log('');
console.log('a:grpFill');
console.log(`  solid child reads the group fill   ${String(groupFill.solidReachesGroup)}`);
console.log(`  the group's own ground is painted  ${String(groupFill.groundIsPainted)}`);
console.log(`  gradient is a slice, not a copy    ${String(groupFill.gradientIsSlice)}`);
console.log(`    left ${gradientLeft}, right ${gradientRight}`);
console.log(`  over a:noFill                      ${groupFill.overNoFill}`);
console.log(`  over a group that says nothing     ${groupFill.overSilence}`);
console.log(`  on a shape in no group             ${groupFill.atTopLevel}`);
console.log(`  through an inner grpFill           ${String(groupFill.nestedReachesOuter)}`);
console.log('');
console.log('what a slide shows that it did not write');
console.log('  deck                 master  layout  slide   placeholders');
for (const row of compositing) {
  console.log(
    `  ${row.deck.padEnd(20)} ${String(row.masterVisible).padEnd(7)} ${String(row.layoutVisible).padEnd(7)}` +
      ` ${String(row.slideVisible).padEnd(7)} ${String(row.placeholdersVisible)}`,
  );
}
console.log('');
console.log('stroke width inside a doubled group');
console.log(`  Line.Weight after ungroup          ${String(strokeScale.weightAfterUngroup)}`);
console.log(`  a doubled stroke covers the row    ${String(strokeScale.strokeIsScaled)}`);
console.log('');
console.log(`  through an inner grpFill (raw)     ${groupFill.nested}`);
console.log(`  the group ground alone             ${sampleColor('f-alone-inside')}`);
console.log(
  `  pattern ink: child ${String(groupFill.patternLeafInk)}` +
    ` ground ${String(groupFill.patternGroundInk)} reference ${String(groupFill.patternReferenceInk)}`,
);
console.log('');
console.log('the rectangle a group fill is laid out over');
for (const check of spanChecks) {
  console.log(
    `  ${check.probe.padEnd(12)} ${check.axis}  measured ${JSON.stringify(check.measured)}` +
      `  content ${JSON.stringify(check.contentBounds)}  declared ${JSON.stringify(check.declaredBox)}` +
      `  -> content ${String(check.matchesContent)}, declared ${String(check.matchesDeclared)}`,
  );
}
for (const [name, value] of Object.entries(groupFill.extents)) {
  console.log(`  ${name.padEnd(18)} ${JSON.stringify(value)}`);
}
for (const r of ramps) {
  console.log(
    `  ${r.id.padEnd(16)} at ${String(r.at).padStart(4)}  ${r.rgb}  t=${r.t === null ? 'no match' : String(Math.round(r.t * 1000) / 1000)}`,
  );
}
console.log('');
console.log('a rotated child in a non-uniformly scaled group');
console.log('  probe          rot   sx   sy   baked w x h      reported w x h    rot  swap centre');
for (const s of shearCases) {
  console.log(
    `  ${s.id.padEnd(14)} ${String(s.rot).padStart(3)} ${String(s.sx).padStart(4)} ` +
      `${String(s.sy).padStart(4)}   ${`${String(s.baked.w)} x ${String(s.baked.h)}`.padEnd(14)} ` +
      `${`${String(Math.round(s.reported.w * 100) / 100)} x ${String(Math.round(s.reported.h * 100) / 100)}`.padEnd(16)} ` +
      `${String(s.reportedRotation).padStart(4)} ${s.swapped ? 'yes ' : s.baked_ ? 'no  ' : '??  '} ${String(s.centreAgrees)}`,
  );
}
for (const s of shear.samples) console.log(`  ${s.id.padEnd(16)} ${s.block}  (${s.asks})`);
console.log('  when do the two extents change places?');
for (const s of swapScores) {
  console.log(
    `    ${s.model.padEnd(22)} ${String(s.right).padStart(2)}/${String(s.of)}` +
      (s.misses.length > 0 ? `   misses ${s.misses.join(' ')}` : ''),
  );
}
console.log(`  the centre always agrees   ${String(shear.centreAlwaysAgrees)}`);
console.log(`  the angle is always kept   ${String(shear.rotationAlwaysKept)}`);
console.log('');
console.log(`${String(cases.length)} probes, ${String(samples.length)} samples`);
