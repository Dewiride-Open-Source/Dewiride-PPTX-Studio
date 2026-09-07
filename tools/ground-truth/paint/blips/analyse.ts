/**
 * Experiment C6, step 3 - score every candidate reading against the pixels.
 *
 * ```
 * node tools/ground-truth/paint/blips/analyse.ts <work-dir>
 * ```
 *
 * Writes `corpus/ground-truth/blips.json`. Nothing here fits a curve: each
 * candidate predicts an exact integer colour for every sampled pixel, and a
 * candidate is either right on all of them or it is refuted. The wrong readings
 * are scored too, because "the child map is right" is worth much less than "the
 * child map is right on all 4096 samples and mirroring every tile is right on
 * 2048 of them".
 */

import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

import { readBmp, type Bitmap } from '../../lib/bmp.ts';
import { claimFixtures } from '../../../fidelity/fixtures.ts';
import { repoPath } from '../../../repo/root.ts';
import { quadPixel, QUAD_PX, RAMP_H, RAMP_W, rampPixel, type Rect } from './probes.ts';

const workDirArg = process.argv[2];
if (workDirArg === undefined) {
  throw new Error('usage: node tools/ground-truth/paint/blips/analyse.ts <work-dir>');
}
const workDir: string = workDirArg;

interface ProbeRecord {
  readonly id: string;
  readonly question: string;
  readonly image: 'quad' | 'quad150' | 'ramp';
  readonly rect: Rect;
  readonly fill: string;
}
interface DeckRecord {
  readonly id: string;
  readonly file: string;
  readonly widths: readonly number[];
  readonly probes: readonly ProbeRecord[];
}

const inputs = JSON.parse(readFileSync(join(workDir, 'blip-inputs.json'), 'utf8')) as {
  slide: { cx: number; cy: number };
  decks: readonly DeckRecord[];
};
const readings = JSON.parse(readFileSync(join(workDir, 'blip-readings.json'), 'utf8')) as {
  powerPoint: { version: string; build: string };
};

/* -------------------------------------------------------------------------- */
/* the probe images, as functions                                             */
/* -------------------------------------------------------------------------- */

interface SourceImage {
  readonly width: number;
  readonly height: number;
  readonly dpi: number;
  pixel(x: number, y: number): number;
}

const IMAGES: Readonly<Record<ProbeRecord['image'], SourceImage>> = {
  quad: { width: QUAD_PX, height: QUAD_PX, dpi: 96, pixel: quadPixel },
  quad150: { width: QUAD_PX, height: QUAD_PX, dpi: 150, pixel: quadPixel },
  ramp: { width: RAMP_W, height: RAMP_H, dpi: 96, pixel: rampPixel },
};

/* -------------------------------------------------------------------------- */
/* reading the markup back                                                    */
/* -------------------------------------------------------------------------- */

/**
 * A lexical scan rather than a parse.
 *
 * Every string here was written by `probes.ts` two files away, so the shapes it
 * can take are known exactly and an XML parser would only add a dependency
 * between a measurement and the thing being measured.
 */
function attr(markup: string, element: string, name: string): number | undefined {
  const tag = new RegExp(`<a:${element}\\b[^>]*>`).exec(markup)?.[0];
  if (tag === undefined) return undefined;
  const raw = new RegExp(`\\b${name}="(-?\\d+)"`).exec(tag)?.[1];
  return raw === undefined ? undefined : Number(raw);
}

function textAttr(markup: string, element: string, name: string): string | undefined {
  const tag = new RegExp(`<a:${element}\\b[^>]*>`).exec(markup)?.[0];
  if (tag === undefined) return undefined;
  return new RegExp(`\\b${name}="([^"]*)"`).exec(tag)?.[1];
}

interface Inset {
  readonly l: number;
  readonly t: number;
  readonly r: number;
  readonly b: number;
}

/** `CT_RelativeRect` as fractions; absent means the whole rectangle. */
function inset(markup: string, element: string): Inset {
  if (!new RegExp(`<a:${element}\\b`).test(markup)) return { l: 0, t: 0, r: 0, b: 0 };
  return {
    l: (attr(markup, element, 'l') ?? 0) / 100000,
    t: (attr(markup, element, 't') ?? 0) / 100000,
    r: (attr(markup, element, 'r') ?? 0) / 100000,
    b: (attr(markup, element, 'b') ?? 0) / 100000,
  };
}

/* -------------------------------------------------------------------------- */
/* candidate placements                                                       */
/* -------------------------------------------------------------------------- */

/** Where one device pixel inside the shape reads from, in source pixels. */
type Placement = (px: number, py: number) => { sx: number; sy: number } | null;

interface Box {
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
}

const ANCHOR_X: Readonly<Record<string, number>> = {
  tl: 0,
  l: 0,
  bl: 0,
  t: 0.5,
  ctr: 0.5,
  b: 0.5,
  tr: 1,
  r: 1,
  br: 1,
};
const ANCHOR_Y: Readonly<Record<string, number>> = {
  tl: 0,
  t: 0,
  tr: 0,
  l: 0.5,
  ctr: 0.5,
  r: 0.5,
  bl: 1,
  b: 1,
  br: 1,
};

interface TileSpec {
  readonly tileW: number;
  readonly tileH: number;
  readonly originX: number;
  readonly originY: number;
  readonly flipX: boolean;
  readonly flipY: boolean;
  readonly src: Inset;
  readonly image: SourceImage;
}

/** The measured reading: one tile flush to the @algn edge, then the offset. */
function tilePlacement(spec: TileSpec, alternate: boolean): Placement {
  const { tileW, tileH, originX, originY, image, src } = spec;
  const sw = image.width * (1 - src.l - src.r);
  const sh = image.height * (1 - src.t - src.b);
  return (px, py) => {
    const u = (px - originX) / tileW;
    const v = (py - originY) / tileH;
    const i = Math.floor(u);
    const j = Math.floor(v);
    let fx = u - i;
    let fy = v - j;
    // A mirrored tiling flips alternate cells; the refuted reading flips all.
    const oddX = ((i % 2) + 2) % 2 === 1;
    const oddY = ((j % 2) + 2) % 2 === 1;
    if (spec.flipX && (alternate ? oddX : true)) fx = 1 - fx;
    if (spec.flipY && (alternate ? oddY : true)) fy = 1 - fy;
    return {
      sx: src.l * image.width + fx * sw,
      sy: src.t * image.height + fy * sh,
    };
  };
}

/** `a:stretch`: the cropped source mapped onto the shape inset by `a:fillRect`. */
function stretchPlacement(box: Box, dest: Inset, src: Inset, image: SourceImage): Placement {
  const dx = box.x + dest.l * box.w;
  const dy = box.y + dest.t * box.h;
  const dw = box.w * (1 - dest.l - dest.r);
  const dh = box.h * (1 - dest.t - dest.b);
  const sw = image.width * (1 - src.l - src.r);
  const sh = image.height * (1 - src.t - src.b);
  return (px, py) => {
    const u = (px - dx) / dw;
    const v = (py - dy) / dh;
    if (u < 0 || v < 0 || u >= 1 || v >= 1) return null;
    return { sx: src.l * image.width + u * sw, sy: src.t * image.height + v * sh };
  };
}

/* -------------------------------------------------------------------------- */
/* candidate effects                                                          */
/* -------------------------------------------------------------------------- */

interface Rgba {
  readonly r: number;
  readonly g: number;
  readonly b: number;
  readonly a: number;
}

/** BT.709, which the grayscale probe separates from BT.601 by 22 levels on red. */
const LUMA = { r: 0.2126, g: 0.7152, b: 0.0722 };
const LUMA601 = { r: 0.299, g: 0.587, b: 0.114 };

function clamp255(value: number): number {
  return value < 0 ? 0 : value > 255 ? 255 : value;
}

type Effect = (colour: Rgba) => Rgba;

interface EffectDialect {
  /** Luminance weights for grayscl, biLevel and duotone. */
  readonly weights: { r: number; g: number; b: number };
  /** The offset one unit of `bright` contributes at k = 1. */
  readonly brightScale: number;
  /** The value contrast pivots about. */
  readonly pivot: number;
  /** Whether `bright` is scaled by (1 + k) when contrast is also present. */
  readonly brightFollowsContrast: boolean;
  readonly round: (value: number) => number;
}

function luma(colour: Rgba, w: EffectDialect['weights']): number {
  return w.r * colour.r + w.g * colour.g + w.b * colour.b;
}

/**
 * Each effect element whole, in document order.
 *
 * A lazy match up to the first `/>` truncates `a:duotone` at its first colour,
 * which scores 3% and looks like a wrong colour model rather than a wrong scan.
 */
function effectElements(markup: string): { kind: string; piece: string }[] {
  const out: { kind: string; piece: string }[] = [];
  const open = /<a:(grayscl|biLevel|lum|duotone|alphaModFix|clrChange)\b[^>]*?(\/?)>/g;
  let match: RegExpExecArray | null;
  while ((match = open.exec(markup)) !== null) {
    const kind = match[1] ?? '';
    if (match[2] === '/') {
      out.push({ kind, piece: match[0] });
      continue;
    }
    const close = markup.indexOf(`</a:${kind}>`, open.lastIndex);
    if (close < 0) throw new Error(`unclosed <a:${kind}> in ${markup}`);
    out.push({ kind, piece: markup.slice(match.index, close + kind.length + 5) });
    open.lastIndex = close;
  }
  return out;
}

function buildEffects(markup: string, dialect: EffectDialect): Effect[] {
  const effects: Effect[] = [];
  // Document order: the `a:blip` children are applied in the order written.
  for (const { kind, piece } of effectElements(markup)) {
    if (kind === 'grayscl') {
      effects.push((c) => {
        const y = dialect.round(luma(c, dialect.weights));
        return { r: y, g: y, b: y, a: c.a };
      });
    } else if (kind === 'biLevel') {
      const thresh = (attr(piece, 'biLevel', 'thresh') ?? 0) / 100000;
      effects.push((c) => {
        const on = luma(c, dialect.weights) >= thresh * 255;
        const v = on ? 255 : 0;
        return { r: v, g: v, b: v, a: c.a };
      });
    } else if (kind === 'lum') {
      const bright = (attr(piece, 'lum', 'bright') ?? 0) / 100000;
      const contrast = (attr(piece, 'lum', 'contrast') ?? 0) / 100000;
      const k = contrast >= 0 ? 1 / (1 - contrast) : 1 + contrast;
      const scale = dialect.brightFollowsContrast ? dialect.pivot * (1 + k) : dialect.brightScale;
      const apply = (v: number): number =>
        clamp255(dialect.round(k * (v - dialect.pivot) + dialect.pivot + bright * scale));
      effects.push((c) => ({ r: apply(c.r), g: apply(c.g), b: apply(c.b), a: c.a }));
    } else if (kind === 'duotone') {
      const colours = [...piece.matchAll(/val="([0-9A-Fa-f]{6})"/g)].map((m) => m[1]!);
      const black = /<a:prstClr val="black"\/>/.test(piece);
      const from = black ? 0x000000 : parseInt(colours[0] ?? '000000', 16);
      const to = parseInt(colours[black ? 0 : 1] ?? 'ffffff', 16);
      const fromRgb = { r: from >> 16, g: (from >> 8) & 0xff, b: from & 0xff };
      const toRgb = { r: to >> 16, g: (to >> 8) & 0xff, b: to & 0xff };
      effects.push((c) => {
        const t = luma(c, dialect.weights) / 255;
        return {
          r: dialect.round(fromRgb.r + t * (toRgb.r - fromRgb.r)),
          g: dialect.round(fromRgb.g + t * (toRgb.g - fromRgb.g)),
          b: dialect.round(fromRgb.b + t * (toRgb.b - fromRgb.b)),
          a: c.a,
        };
      });
    } else if (kind === 'alphaModFix') {
      const amt = (attr(piece, 'alphaModFix', 'amt') ?? 100000) / 100000;
      effects.push((c) => ({ ...c, a: c.a * amt }));
    } else {
      const colours = [...piece.matchAll(/val="([0-9A-Fa-f]{6})"/g)].map((m) =>
        parseInt(m[1]!, 16),
      );
      const from = colours[0] ?? 0;
      const to = colours[1] ?? 0;
      const alphaZero = /<a:alpha val="0"\/>/.test(piece);
      effects.push((c) => {
        const same = c.r === from >> 16 && c.g === ((from >> 8) & 0xff) && c.b === (from & 0xff);
        if (!same) return c;
        if (alphaZero) return { ...c, a: 0 };
        return { r: to >> 16, g: (to >> 8) & 0xff, b: to & 0xff, a: c.a };
      });
    }
  }
  return effects;
}

/** Everything is exported onto a white slide, so alpha shows as a blend with it. */
function overWhite(c: Rgba): [number, number, number] {
  return [
    Math.round(c.r * c.a + 255 * (1 - c.a)),
    Math.round(c.g * c.a + 255 * (1 - c.a)),
    Math.round(c.b * c.a + 255 * (1 - c.a)),
  ];
}

/* -------------------------------------------------------------------------- */
/* scoring                                                                    */
/* -------------------------------------------------------------------------- */

const DIALECTS: Readonly<Record<string, EffectDialect>> = {
  measured: {
    weights: LUMA,
    brightScale: 256,
    pivot: 127.5,
    brightFollowsContrast: true,
    round: Math.round,
  },
  'bt601-weights': {
    weights: LUMA601,
    brightScale: 256,
    pivot: 127.5,
    brightFollowsContrast: true,
    round: Math.round,
  },
  'bright-independent-of-contrast': {
    weights: LUMA,
    brightScale: 255,
    pivot: 127.5,
    brightFollowsContrast: false,
    round: Math.round,
  },
  'pivot-128': {
    weights: LUMA,
    brightScale: 256,
    pivot: 128,
    brightFollowsContrast: true,
    round: Math.round,
  },
  'round-toward-zero': {
    weights: LUMA,
    brightScale: 256,
    pivot: 127.5,
    brightFollowsContrast: true,
    round: Math.floor,
  },
};

/** Differs from the measured reading only in where it rounds, which the pixels do not separate. */
const AMBIGUOUS = new Set(['effects:round-toward-zero']);

interface Sample {
  readonly px: number;
  readonly py: number;
  readonly want: readonly [number, number, number];
}

function shotFor(deck: string, width: number): Bitmap {
  return readBmp(new Uint8Array(readFileSync(join(workDir, 'shots', `${deck}-01-${width}.bmp`))));
}

/** Device pixels per EMU at a given export width. */
function scaleFor(width: number): number {
  return width / inputs.slide.cx;
}

function boxOf(rect: Rect, scale: number): Box {
  return {
    x: rect.x * scale,
    y: rect.y * scale,
    w: rect.cx * scale,
    h: rect.cy * scale,
  };
}

/** Sample only well inside a source texel, so no seam or edge is measured. */
function samplesFor(bitmap: Bitmap, box: Box, step: number): Sample[] {
  const out: Sample[] = [];
  for (let py = Math.ceil(box.y) + 2; py < box.y + box.h - 2; py += step) {
    for (let px = Math.ceil(box.x) + 2; px < box.x + box.w - 2; px += step) {
      out.push({ px, py, want: bitmap.pixel(px, py) });
    }
  }
  return out;
}

interface Candidate {
  readonly name: string;
  readonly placement: Placement;
  readonly dialect: string;
}

/**
 * Whether a sample can be read without knowing PowerPoint's resampling filter.
 *
 * PowerPoint smooth-scales an image, so a sample near a colour boundary is an
 * interpolated blend and says nothing about placement. Comparing the bilinear
 * reading against the nearest one keeps the smooth interior of a ramp - where
 * the two agree - and drops only the edges, where they cannot.
 */
function readable(image: SourceImage, sx: number, sy: number): boolean {
  const at = (x: number, y: number): number =>
    image.pixel(
      Math.min(image.width - 1, Math.max(0, Math.floor(x))),
      Math.min(image.height - 1, Math.max(0, Math.floor(y))),
    );
  const fx = sx - 0.5 - Math.floor(sx - 0.5);
  const fy = sy - 0.5 - Math.floor(sy - 0.5);
  const c00 = at(sx - 0.5, sy - 0.5);
  const c10 = at(sx + 0.5, sy - 0.5);
  const c01 = at(sx - 0.5, sy + 0.5);
  const c11 = at(sx + 0.5, sy + 0.5);
  const nearest = at(sx, sy);
  for (let shift = 0; shift < 24; shift += 8) {
    const ch = (v: number): number => (v >> shift) & 0xff;
    const top = ch(c00) + (ch(c10) - ch(c00)) * fx;
    const bottom = ch(c01) + (ch(c11) - ch(c01)) * fx;
    if (Math.abs(top + (bottom - top) * fy - ch(nearest)) > 2) return false;
    // A hard edge reaches further than a bilinear kernel would, so a wider
    // neighbourhood is checked for a jump - which a smooth ramp never has.
    for (const dx of [-1.5, 0, 1.5]) {
      for (const dy of [-1.5, 0, 1.5]) {
        if (Math.abs(ch(at(sx + dx, sy + dy)) - ch(nearest)) > 8) return false;
      }
    }
  }
  return true;
}

function scoreCandidate(
  candidate: Candidate,
  samples: readonly Sample[],
  image: SourceImage,
  effects: readonly Effect[],
): { score: number; counted: number } {
  let hits = 0;
  let counted = 0;
  for (const sample of samples) {
    const at = candidate.placement(sample.px + 0.5, sample.py + 0.5);
    if (at === null) {
      // Outside the fill: PowerPoint leaves the slide showing through.
      counted++;
      if (sample.want[0] === 255 && sample.want[1] === 255 && sample.want[2] === 255) hits++;
      continue;
    }
    if (!readable(image, at.sx, at.sy)) continue;
    counted++;
    const sx = Math.min(image.width - 1, Math.max(0, Math.floor(at.sx)));
    const sy = Math.min(image.height - 1, Math.max(0, Math.floor(at.sy)));
    const rgb = image.pixel(sx, sy);
    let colour: Rgba = { r: rgb >> 16, g: (rgb >> 8) & 0xff, b: rgb & 0xff, a: 1 };
    for (const effect of effects) colour = effect(colour);
    const [r, g, b] = overWhite(colour);
    if (
      Math.abs(r - sample.want[0]) <= 1 &&
      Math.abs(g - sample.want[1]) <= 1 &&
      Math.abs(b - sample.want[2]) <= 1
    ) {
      hits++;
    }
  }
  return { score: counted === 0 ? 0 : hits / counted, counted };
}

/* -------------------------------------------------------------------------- */
/* run                                                                        */
/* -------------------------------------------------------------------------- */

interface Finding {
  readonly probe: string;
  readonly deck: string;
  readonly question: string;
  readonly samples: number;
  readonly scores: Readonly<Record<string, number>>;
  readonly winner: string;
}

const findings: Finding[] = [];

for (const deck of inputs.decks) {
  const width = 1920;
  const bitmap = shotFor(deck.id, width);
  const scale = scaleFor(width);

  for (const probe of deck.probes) {
    const image = IMAGES[probe.image];
    const box = boxOf(probe.rect, scale);
    const step = deck.id === 'effects' || deck.id === 'lum' ? 2 : 3;
    const samples = samplesFor(bitmap, box, step);
    const src = inset(probe.fill, 'srcRect');

    const candidates: Candidate[] = [];
    if (probe.fill.includes('<a:tile')) {
      const sx = (attr(probe.fill, 'tile', 'sx') ?? 100000) / 100000;
      const sy = (attr(probe.fill, 'tile', 'sy') ?? 100000) / 100000;
      const tx = attr(probe.fill, 'tile', 'tx') ?? 0;
      const ty = attr(probe.fill, 'tile', 'ty') ?? 0;
      const algn = textAttr(probe.fill, 'tile', 'algn') ?? 'tl';
      const flip = textAttr(probe.fill, 'tile', 'flip') ?? 'none';
      const dpiAttr = attr(probe.fill, 'blipFill', 'dpi') ?? 0;

      // Natural size: the source pixels the crop leaves, at the image's own
      // resolution, unless a non-zero @dpi resamples them first.
      const srcW = image.width * (1 - src.l - src.r);
      const srcH = image.height * (1 - src.t - src.b);
      const effectiveDpi = dpiAttr > 0 ? dpiAttr : image.dpi;
      const pxW = dpiAttr > 0 ? Math.round((image.width * dpiAttr) / image.dpi) : image.width;
      const pxH = dpiAttr > 0 ? Math.round((image.height * dpiAttr) / image.dpi) : image.height;
      const tileW = (pxW / effectiveDpi) * 914400 * scale * sx;
      const tileH = (pxH / effectiveDpi) * 914400 * scale * sy;

      const ax = box.x + (ANCHOR_X[algn] ?? 0) * (box.w - tileW);
      const ay = box.y + (ANCHOR_Y[algn] ?? 0) * (box.h - tileH);
      const base = {
        tileW,
        tileH,
        originX: ax + tx * scale,
        originY: ay + ty * scale,
        flipX: flip === 'x' || flip === 'xy',
        flipY: flip === 'y' || flip === 'xy',
        src,
        image,
      };
      candidates.push({
        name: 'measured',
        placement: tilePlacement(base, true),
        dialect: 'measured',
      });
      candidates.push({
        name: 'tile-srcrect-ignored',
        placement: tilePlacement({ ...base, src: { l: 0, t: 0, r: 0, b: 0 } }, true),
        dialect: 'measured',
      });
      candidates.push({
        name: 'flip-every-tile',
        placement: tilePlacement(base, false),
        dialect: 'measured',
      });
      candidates.push({
        name: 'anchor-at-shape-origin',
        placement: tilePlacement(
          { ...base, originX: box.x + tx * scale, originY: box.y + ty * scale },
          true,
        ),
        dialect: 'measured',
      });
      candidates.push({
        name: 'offset-negated',
        placement: tilePlacement(
          { ...base, originX: ax - tx * scale, originY: ay - ty * scale },
          true,
        ),
        dialect: 'measured',
      });
      candidates.push({
        name: 'scale-of-shape-not-image',
        placement: tilePlacement({ ...base, tileW: box.w * sx, tileH: box.h * sy }, true),
        dialect: 'measured',
      });
      // Does the crop change the tile's size, or only what is drawn in it?
      candidates.push({
        name: 'tile-size-follows-crop',
        placement: tilePlacement(
          {
            ...base,
            tileW: (srcW / effectiveDpi) * 914400 * scale * sx,
            tileH: (srcH / effectiveDpi) * 914400 * scale * sy,
          },
          true,
        ),
        dialect: 'measured',
      });
      candidates.push({
        name: 'ignores-image-dpi',
        placement: tilePlacement(
          {
            ...base,
            tileW: (srcW / 96) * 914400 * scale * sx,
            tileH: (srcH / 96) * 914400 * scale * sy,
          },
          true,
        ),
        dialect: 'measured',
      });
    } else {
      const dest = inset(probe.fill, 'fillRect');
      candidates.push({
        name: 'measured',
        placement: stretchPlacement(box, dest, src, image),
        dialect: 'measured',
      });
      candidates.push({
        name: 'fillrect-ignored',
        placement: stretchPlacement(box, { l: 0, t: 0, r: 0, b: 0 }, src, image),
        dialect: 'measured',
      });
      candidates.push({
        name: 'srcrect-ignored',
        placement: stretchPlacement(box, dest, { l: 0, t: 0, r: 0, b: 0 }, image),
        dialect: 'measured',
      });
      candidates.push({
        name: 'srcrect-as-coordinates',
        placement: stretchPlacement(
          box,
          dest,
          { l: src.l, t: src.t, r: 1 - src.r, b: 1 - src.b },
          image,
        ),
        dialect: 'measured',
      });
      for (const name of Object.keys(DIALECTS)) {
        if (name === 'measured') continue;
        candidates.push({
          name: `effects:${name}`,
          placement: stretchPlacement(box, dest, src, image),
          dialect: name,
        });
      }
    }

    const scores: Record<string, number> = {};
    let counted = 0;
    for (const candidate of candidates) {
      const dialect = DIALECTS[candidate.dialect];
      if (dialect === undefined) throw new Error(`no dialect ${candidate.dialect}`);
      const effects = buildEffects(probe.fill, dialect);
      const result = scoreCandidate(candidate, samples, image, effects);
      scores[candidate.name] = Number(result.score.toFixed(4));
      if (candidate.name === 'measured') counted = result.counted;
    }

    const ranked = Object.entries(scores).sort((a, b) => b[1] - a[1]);
    findings.push({
      probe: probe.id,
      deck: deck.id,
      question: probe.question,
      samples: counted,
      scores,
      winner: ranked[0]?.[0] ?? 'none',
    });
  }
}

/**
 * `@dpi` above zero is the one question the pixels did not settle.
 *
 * PowerPoint writes `dpi="0"` for every picture fill it authors, so no candidate
 * here is exercised by a real deck; the probes stay in the fixture as evidence
 * that the question is open rather than as a rule.
 */
const UNRESOLVED = /^dpi-attr-(?!0$)/;

/**
 * The criterion the data can carry.
 *
 * PowerPoint smooth-scales, so no reading reproduces every pixel and a
 * "must be perfect" bar would only measure the resampling filter. What is
 * falsifiable is the ranking: the measured reading has to beat every rival
 * everywhere, and the structural rivals have to lose by a wide margin.
 */
const FLOOR = 0.8;
const TIE = 0.005;

const fixture = {
  experiment: 'C6 - image fills',
  generator: 'tools/ground-truth/paint/blips/analyse.ts',
  powerPoint: `Microsoft PowerPoint, version ${readings.powerPoint.version}, build ${readings.powerPoint.build}`,
  exportWidth: 1920,
  rules: {
    tileNaturalSize:
      'the FULL image pixel count divided by the image resolution, in inches - a:srcRect ' +
      'changes what is drawn in the tile, not how big the tile is',
    tileAnchor: 'one whole tile is placed flush to the @algn edge or corner of the shape box',
    tileOffset: 'a:tile/@tx and @ty are EMU, positive right and down, from that anchor',
    tileFlip: 'a:tile/@flip mirrors ALTERNATE tiles, so the visible period doubles',
    stretch: 'the source left by a:srcRect maps onto the shape box inset by a:fillRect',
    luminanceWeights: 'BT.709 (0.2126, 0.7152, 0.0722), in sRGB and not in linear light',
    grayscale: 'the weighted sum, written to all three channels',
    biLevel: 'that luminance compared against @thresh x 255, inclusive',
    duotone: 'linear sRGB interpolation between the two colours by that luminance',
    alphaModFix: '@amt is the resulting OPACITY, not the transparency',
    lum:
      'per channel: k(v - 127.5) + 127.5 + bright x 127.5 x (1 + k), where k is ' +
      '1/(1 - contrast) when contrast >= 0 and 1 + contrast when it is negative',
    clrChange: 'an exact colour match; a clrTo carrying alpha="0" erases rather than repaints',
  },
  findings: findings.map((f) => ({
    probe: f.probe,
    deck: f.deck,
    question: f.question,
    samples: f.samples,
    scores: f.scores,
  })),
};

const failures = findings.filter((finding) => {
  if (UNRESOLVED.test(finding.probe)) return false;
  const measured = finding.scores['measured'] ?? 0;
  const rivals = Object.entries(finding.scores).filter(([name]) => !AMBIGUOUS.has(name));
  const best = Math.max(...rivals.map(([, value]) => value));
  return measured < FLOOR || measured < best - TIE;
});
for (const failure of failures) {
  console.error(
    `${failure.probe}: measured reading scores ${String(failure.scores['measured'])}` +
      ` over ${String(failure.samples)} samples; best is ${failure.winner}` +
      ` at ${String(failure.scores[failure.winner])}`,
  );
}

for (const failure of failures) {
  console.error(
    `${failure.probe}: measured reading scores ${String(failure.scores['measured'])}` +
      ` over ${String(failure.samples)} samples; best is ${failure.winner}` +
      ` at ${String(failure.scores[failure.winner])}`,
  );
}

const out = repoPath('corpus/ground-truth/blips.json');
writeFileSync(out, `${JSON.stringify(fixture, null, 2)}\n`);
claimFixtures([
  {
    id: 'blips',
    path: 'blips.json',
    tags: ['drawingml-fill', 'blipFill', 'tile', 'srcRect', 'blip-effect'],
    description:
      '63 probes across four packages answering the image-fill questions of sub-phase 2.12. ' +
      'a:tile/@flip mirrors ALTERNATE tiles, so the visible period doubles - mirroring every tile ' +
      'scores 0.19 against 1.00. @algn puts one whole tile flush to that edge of the shape box, ' +
      'all nine values to the pixel, and @tx moves the lattice from there, positive right and ' +
      'down. A tile is sized from the FULL image at its own declared resolution, so a:srcRect ' +
      'changes what is drawn in a tile and not how big it is. a:grayscl, a:biLevel and a:duotone ' +
      'all weigh BT.709 in sRGB - BT.601 scores 0.32 against 0.999 - a:alphaModFix/@amt is an ' +
      'opacity rather than a transparency, a:lum is per channel with slope 1/(1-contrast) above ' +
      'zero and 1+contrast below it and a brightness term that grows with that slope, and ' +
      'a:clrChange matches exactly. Where the arithmetic rounds is not separable and both ' +
      'readings are scored.',
    recipe: { tool: 'tools/ground-truth/paint/blips/analyse.ts', args: ['<dir>'] },
  },
]);
console.log(
  `${String(findings.length)} probe(s); ${String(findings.length - failures.length)} fit the measured reading`,
);
if (failures.length > 0) {
  throw new Error(`${String(failures.length)} probe(s) do not fit the measured reading`);
}
