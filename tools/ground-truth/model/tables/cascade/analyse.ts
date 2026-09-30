/**
 * Experiment C9, step 3 - score every reading of the table-style cascade against what PowerPoint
 * reported and drew.
 *
 * ```
 * node tools/ground-truth/model/tables/cascade/analyse.ts <dir> [--fixture]
 * ```
 *
 * Throws unless exactly one reading, or one group no drawable table tells apart, fits every row of
 * each question, and unless every instrument guard holds.
 */

import { createHash } from 'node:crypto';
import { existsSync, mkdirSync, readdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { deflateRawSync } from 'node:zlib';

import {
  BUILTIN_TEXT_STYLES,
  TEXT_FLOOR,
  type TableStyle,
  type TableStylePartName,
} from '../../../../../packages/model/dist/index.js';
import { repoPath } from '../../../../repo/root.ts';
import { hex, readBmp } from '../../../lib/bmp.ts';
import { fixtureJson } from '../../../lib/json.ts';
import { readPng, type PackedBitmap } from '../../../lib/png.ts';
import { themeXml } from '../../../lib/sheet-pptx.ts';
import { entry, readZip } from '../../../lib/zip.ts';
import {
  allDecks,
  APPLICABILITY,
  cellsInRow,
  BASELINE,
  CELL,
  EDGE_MODELS,
  edgeKey,
  edgeLines,
  edgesOf,
  FILL_PAINTS,
  fillStack,
  fillWinner,
  lineText,
  MERGES,
  ORDERS,
  ORIGIN,
  PAGE,
  paletteOf,
  partsGrid,
  predictCom,
  predictFill,
  PROFILE_REACH,
  SCALE,
  styledTables,
  cellKey,
  colorHex,
  edgeLinesBetween,
  floorText,
  FLOOR_DEFAULT_TEXT_STYLE,
  FLOOR_TX_STYLES,
  THEMES,
  LADDER,
  LEVELS,
  maskOf,
  styleById,
  TEXT_PACKAGES,
  type DeckSpec,
  type FillReport,
  type TextLevel,
  type DrawnLine,
  type Edge,
  type FillPaint,
  type Palette,
  type Scenario,
  type Shape,
  type TableSpec,
} from './probes.ts';

/* -------------------------------------------------------------------------- */
/* the readings                                                               */
/* -------------------------------------------------------------------------- */

interface SlideReading {
  readonly style: string;
  readonly frame: readonly number[];
  readonly cells: string;
  readonly png: string;
  readonly bmp: string | null;
}

interface DeckReading {
  readonly id: string;
  readonly deck: string;
  readonly opened: boolean;
  readonly repaired: boolean | null;
  readonly error: string | null;
  readonly slides: readonly SlideReading[];
  readonly session: {
    readonly pid: number;
    readonly started: string;
    readonly version: string;
    readonly build: string;
  };
}

/** One cell as COM reported it; `sides` only in `full` mode. */
export interface ComCell {
  readonly fill: string;
  readonly sides: Readonly<Record<string, string>> | null;
  /** `rgb/b/face`, plus `/italic/size/bullet` in `full` mode; `empty` for a cell with no text. */
  readonly text: string;
  readonly rect: readonly number[] | null;
}

export interface ComTable {
  readonly rows: number;
  readonly cols: number;
  readonly background: string | null;
  readonly cells: readonly (readonly ComCell[])[];
}

function parseCom(text: string, where: string): ComTable {
  const [head = '', ...lines] = text.split('\n').filter((l) => l !== '');
  const [dims = '', extra] = head.split('|');
  const [rows, cols] = dims.split('x').map(Number) as [number, number];
  const background = extra === undefined ? null : (/background=([^;]*);/.exec(extra)?.[1] ?? null);
  const cells: ComCell[][] = Array.from({ length: rows }, () => []);
  for (const line of lines) {
    const [at = '', body = '', rect] = line.split('|');
    const [r, c] = at.split(',').map(Number) as [number, number];
    const fields = new Map<string, string>();
    for (const kv of body.split(';')) {
      const eq = kv.indexOf('=');
      if (eq > 0) fields.set(kv.slice(0, eq), kv.slice(eq + 1));
    }
    const sideNames = ['top', 'left', 'bottom', 'right', 'down', 'up'];
    const sides = sideNames.every((s) => fields.has(s))
      ? Object.fromEntries(sideNames.map((s) => [s, fields.get(s) ?? '']))
      : null;
    const fill = fields.get('fill');
    const txt = fields.get('text');
    if (fill === undefined || txt === undefined)
      throw new Error(`${where}: cell ${at} lacks a fill or text`);
    if (/raw/.test(fill) || /raw/.test(txt))
      throw new Error(`${where}: cell ${at} reports a mixed value: ${body}`);
    const row = cells[r];
    if (row === undefined) throw new Error(`${where}: cell ${at} is outside ${dims}`);
    row[c] = {
      fill,
      sides,
      text: txt,
      rect: rect === undefined ? null : rect.slice('rect='.length).split(',').map(Number),
    };
  }
  return { rows, cols, background, cells };
}

/* -------------------------------------------------------------------------- */
/* the pixels                                                                 */
/* -------------------------------------------------------------------------- */

const at = (pt: number): number => Math.round(pt * SCALE);
const hexOf = (v: number): string => v.toString(16).toUpperCase().padStart(6, '0');

/** A cell's paint: a 4 x 4 block below its text and clear of every edge, which must be one colour. */
export function interiorOf(image: PackedBitmap, r: number, c: number, where: string): string {
  const x0 = at(ORIGIN.x + c * CELL.w + 36);
  const y0 = at(ORIGIN.y + r * CELL.h + 30);
  const first = image.rgb(x0, y0);
  for (let y = y0; y < y0 + 4; y++) {
    for (let x = x0; x < x0 + 4; x++) {
      if (image.rgb(x, y) !== first)
        throw new Error(`${where}: cell ${String(r)},${String(c)} is not one colour`);
    }
  }
  return hexOf(first);
}

/** The pixels across a grid edge at its midpoint, from outside-up/left to inside-down/right. */
function profilePixels(edge: Edge): [number, number][] {
  const out: [number, number][] = [];
  const [x, y] =
    edge.axis === 'h'
      ? [at(ORIGIN.x + edge.c * CELL.w + 36), at(ORIGIN.y + edge.r * CELL.h)]
      : [at(ORIGIN.x + edge.c * CELL.w), at(ORIGIN.y + edge.r * CELL.h + 30)];
  for (let k = -PROFILE_REACH; k < PROFILE_REACH; k++)
    out.push(edge.axis === 'h' ? [x, y + k] : [x + k, y]);
  return out;
}

export function profileOf(image: PackedBitmap, edge: Edge): string[] {
  return profilePixels(edge).map(([x, y]) => hexOf(image.rgb(x, y)));
}

/**
 * What an edge shows once the two paints either side are stripped: where the first other pixel
 * starts, relative to the grid line, and the runs of colour from there. `0|` is no line at all.
 */
export function describe(profile: readonly string[]): string {
  const a = profile[0];
  const b = profile.at(-1);
  let s = 0;
  while (s < profile.length && profile[s] === a) s++;
  let e = profile.length;
  while (e > s && profile[e - 1] === b) e--;
  if (s >= e) return `${String(s - PROFILE_REACH)}|`;
  const runs: string[] = [];
  for (let k = s; k < e;) {
    let n = 1;
    while (k + n < e && profile[k + n] === profile[k]) n++;
    runs.push(`${profile[k] ?? ''}*${String(n)}`);
    k += n;
  }
  return `${String(s - PROFILE_REACH)}|${runs.join(',')}`;
}

/** Per width in points, how much of each pixel across a double line its strokes cover: measured. */
export type Compound = ReadonlyMap<number, readonly number[]>;

/** `c * line + (1 - c) * under`, the sum rounded with ties down: how the 4-pt double line's edges blend. */
function blend(line: string, under: string, c: number): string {
  const l = rgbOf(line);
  const u = rgbOf(under);
  const k = Math.round(c * 1000);
  return hexRgb(u.map((v, i) => roundDown(k * (l[i] ?? 0) + (1000 - k) * v)) as unknown as Rgb);
}

/**
 * The pixels a list of lines would paint over paints `a` and `b` meeting at the grid line. A double
 * line takes its measured coverage; a compound one never measured paints a marker nothing matches.
 */
export function render(
  lines: readonly DrawnLine[],
  a: string,
  b: string,
  double: Compound,
): { readonly pixels: string[]; readonly soft: (Soft | null)[] } {
  const pixels = Array.from({ length: 2 * PROFILE_REACH }, (_, k) => (k < PROFILE_REACH ? a : b));
  const soft = pixels.map((): Soft | null => null);
  for (const line of lines) {
    const half = (line.weight * SCALE) / 2;
    const from = PROFILE_REACH - Math.ceil(half);
    const to = PROFILE_REACH + Math.ceil(half);
    const coverage = line.cmpd === 'dbl' ? double.get(line.weight) : undefined;
    for (let k = from; k < to; k++) {
      if (line.cmpd !== 'sng' && coverage === undefined) {
        pixels[k] = `?${line.cmpd}${String(line.weight)}`;
        continue;
      }
      // A single line covers each pixel by its overlap with [-half, half) about the grid line.
      const overlap = Math.min(k + 1, PROFILE_REACH + half) - Math.max(k, PROFILE_REACH - half);
      const c = coverage === undefined ? overlap : (coverage[k - from] ?? 0);
      if (c <= 0) continue;
      soft[k] = c < 1 ? { line: line.rgb, under: pixels[k] ?? '' } : null;
      pixels[k] = c >= 1 ? line.rgb : blend(line.rgb, pixels[k] ?? '', c);
    }
  }
  return { pixels, soft };
}

/** The profile a descriptor stands for over paints `a` and `b`: `describe` undone. */
export function profileFrom(descriptor: string, a: string, b: string): string[] {
  const [offset = '0', runs = ''] = descriptor.split('|');
  const out: string[] = Array.from({ length: Number(offset) + PROFILE_REACH }, () => a);
  for (const run of runs === '' ? [] : runs.split(',')) {
    const [colour = '', n = '0'] = run.split('*');
    for (let k = 0; k < Number(n); k++) out.push(colour);
  }
  while (out.length < 2 * PROFILE_REACH) out.push(b);
  return out;
}

/** A pixel a line only partly covers: the line's colour, and what it was painted over. */
interface Soft {
  readonly line: string;
  readonly under: string;
}

/** `x` is `line` over `under` at some coverage, to within one unit a channel. */
function blendOf(line: string, under: string, x: string): boolean {
  if (![line, under, x].every((h) => /^[0-9A-F]{6}$/.test(h))) return false;
  const l = rgbOf(line);
  const u = rgbOf(under);
  const o = rgbOf(x);
  const spread = l.map((v, i) => Math.abs(v - (u[i] ?? 0)));
  const i = spread.indexOf(Math.max(...spread));
  const span = (l[i] ?? 0) - (u[i] ?? 0);
  if (Math.abs(span) <= 2) return near(x, line) || near(x, under);
  // Partly covered: strictly between the two, so a whole or a bare pixel never passes for one.
  if (x === line || x === under) return false;
  const c = ((o[i] ?? 0) - (u[i] ?? 0)) / span;
  return (
    c > 0 && c < 1 && l.every((v, j) => Math.abs(c * v + (1 - c) * (u[j] ?? 0) - (o[j] ?? 0)) <= 1)
  );
}

/** Within one unit on every channel. */
function near(x: string, y: string): boolean {
  if (!/^[0-9A-F]{6}$/.test(x) || !/^[0-9A-F]{6}$/.test(y)) return false;
  const p = rgbOf(x);
  const q = rgbOf(y);
  return p.every((v, i) => Math.abs(v - (q[i] ?? -9)) <= 1);
}

/**
 * The descriptor `lines` would draw over `a` and `b`, a partly covered pixel taken as drawn where it
 * is a blend of its line over what was under it: how much a pixel is covered is the rasteriser's
 * (4.4), and which pixels are whole, partial or bare stays exact.
 */
export function drawnAs(
  lines: readonly DrawnLine[],
  a: string,
  b: string,
  double: Compound,
  observed: string,
): string {
  const { pixels, soft } = render(lines, a, b, double);
  const seen = profileFrom(observed, a, b);
  pixels.forEach((p, k) => {
    const drawn = seen[k] ?? '';
    const partial = soft[k];
    if (partial !== null && partial !== undefined && blendOf(partial.line, partial.under, drawn))
      pixels[k] = drawn;
  });
  return describe(pixels);
}

/** The glyph's colour: the commonest colour in the text's box that is not the cell's paint. */
export function glyphOf(
  image: PackedBitmap,
  r: number,
  c: number,
  paint: string,
): { rgb: string | null; ink: number } {
  const counts = new Map<number, number>();
  const skip = Number.parseInt(paint, 16);
  const x0 = at(ORIGIN.x + c * CELL.w + 3);
  const x1 = at(ORIGIN.x + (c + 1) * CELL.w - 3);
  const y0 = at(ORIGIN.y + r * CELL.h + 3);
  const y1 = at(ORIGIN.y + r * CELL.h + 26);
  for (let y = y0; y < y1; y++) {
    for (let x = x0; x < x1; x++) {
      const v = image.rgb(x, y);
      if (v !== skip) counts.set(v, (counts.get(v) ?? 0) + 1);
    }
  }
  const sorted = [...counts].sort((p, q) => q[1] - p[1]);
  const [top, next] = sorted;
  const ink = sorted.reduce((n, [, k]) => n + k, 0);
  if (top === undefined || top[1] < 40 || (next !== undefined && top[1] < 2 * next[1]))
    return { rgb: null, ink };
  return { rgb: hexOf(top[0]), ink };
}

/* -------------------------------------------------------------------------- */
/* scoring                                                                    */
/* -------------------------------------------------------------------------- */

/** One observation, and what a candidate says it should have been. */
export interface Row<C> {
  readonly key: string;
  readonly observed: string;
  readonly predict: (candidate: C) => string;
}

export interface Score {
  readonly name: string;
  readonly fits: number;
  readonly of: number;
  readonly wrong: readonly string[];
}

export interface Finding {
  readonly question: string;
  readonly winner: readonly string[];
  readonly candidates: readonly Score[];
  readonly rows: number;
  readonly informative: number;
}

/**
 * Scores every candidate on the rows where they do not all agree; throws unless the perfect ones are
 * exactly one group that agrees on every row, and fits the rows where all agree as well.
 */
export function score<C>(
  question: string,
  candidates: readonly [string, C][],
  rows: readonly Row<C>[],
): Finding {
  const tallies = candidates.map(([name]) => ({
    name,
    fits: 0,
    wrong: [] as string[],
    sig: [] as string[],
  }));
  let informative = 0;
  const unanimousMisses: string[] = [];
  for (const row of rows) {
    const predicted = candidates.map(([, c]) => row.predict(c));
    const agree = predicted.every((p) => p === predicted[0]);
    if (agree && predicted[0] !== row.observed) {
      unanimousMisses.push(
        `${row.key}: every candidate says ${String(predicted[0])}, observed ${row.observed}`,
      );
    }
    predicted.forEach((p, i) => {
      const tally = tallies[i];
      if (tally === undefined) return;
      tally.sig.push(p);
      if (agree) return;
      if (p === row.observed) tally.fits += 1;
      else if (tally.wrong.length < 4)
        tally.wrong.push(`${row.key}: predicted ${p}, observed ${row.observed}`);
    });
    if (!agree) informative += 1;
  }
  if (unanimousMisses.length > 0) {
    throw new Error(
      `${question}: ${String(unanimousMisses.length)} rows no candidate fits\n  ${unanimousMisses.slice(0, 6).join('\n  ')}`,
    );
  }
  const perfect = tallies.filter((t) => t.fits === informative);
  if (perfect.length === 0) {
    const ranked = [...tallies].sort((p, q) => q.fits - p.fits);
    throw new Error(
      `${question}: no candidate fits every informative row\n` +
        ranked
          .map(
            (t) => `${t.name} ${String(t.fits)}/${String(informative)}\n  ${t.wrong.join('\n  ')}`,
          )
          .join('\n'),
    );
  }
  const first = perfect[0];
  const apart = perfect.filter((t) => t.sig.join('\u0001') !== first?.sig.join('\u0001'));
  if (apart.length > 0) {
    throw new Error(
      `${question}: ${perfect.map((t) => t.name).join(', ')} all fit, and some rows tell them apart`,
    );
  }
  return {
    question,
    winner: perfect.map((t) => t.name),
    candidates: tallies.map((t) => ({
      name: t.name,
      fits: t.fits,
      of: informative,
      wrong: t.wrong,
    })),
    rows: rows.length,
    informative,
  };
}

/* -------------------------------------------------------------------------- */
/* the run                                                                    */
/* -------------------------------------------------------------------------- */

interface Inputs {
  readonly decks: readonly {
    readonly id: string;
    readonly slides: readonly { readonly mask: number }[];
  }[];
  readonly twins: readonly string[];
  readonly separability: readonly unknown[];
}

/** One slide: its spec, what COM said, and its pixels, decoded once. */
export interface Observed {
  readonly deck: DeckSpec;
  readonly slide: number;
  readonly table: TableSpec;
  readonly com: ComTable;
  readonly pixels: Pixels;
}

/** Everything the analysis reads off a slide's export, taken once so the image can be dropped. */
export interface Pixels {
  /** By visual position; `null` where the paint is no one colour (theme C's gradients). */
  readonly paints: readonly (readonly (string | null)[])[];
  readonly glyphs: readonly (readonly { readonly rgb: string | null; readonly ink: number }[])[];
  /** By visual grid edge, `h0,0`. */
  readonly profiles: ReadonlyMap<string, readonly string[]>;
  readonly samples: Record<string, unknown> | null;
  /** The centre pixels along each grid edge a written line dashes, end to end, by visual edge. */
  readonly along: Readonly<Record<string, readonly string[]>>;
  /** The colour at each candidate point on a diagonal a written line could draw, by `x,y` in pixels. */
  readonly diagonal: Readonly<Record<string, string>>;
}

/** A slide's pixels as cached, with the hash of the PNG they were decoded from. */
type CachedPixels = Omit<Pixels, 'profiles' | 'along' | 'diagonal'> &
  Partial<Pick<Pixels, 'along' | 'diagonal'>> & {
    readonly source: string;
    readonly profiles: [string, readonly string[]][];
  };

function pixelsOf(image: PackedBitmap, deck: DeckSpec, table: TableSpec): Pixels {
  const where = deck.id;
  const background = deck.group === 'background';
  const paints = Array.from({ length: table.rows }, (_, r) =>
    Array.from({ length: table.cols }, (_, c) =>
      background ? null : interiorOf(image, r, c, where),
    ),
  );
  const glyphs = paints.map((row, r) =>
    row.map((paint, c) => (paint === null ? { rgb: null, ink: 0 } : glyphOf(image, r, c, paint))),
  );
  const profiles = new Map<string, string[]>();
  for (const edge of edgesOf(table))
    profiles.set(`${edge.axis}${String(edge.r)},${String(edge.c)}`, profileOf(image, edge));
  let samples: Record<string, unknown> | null = null;
  if (background) {
    const y = at(ORIGIN.y + 2 * CELL.h + 30);
    samples = {
      across: Array.from({ length: table.cols * 8 }, (_, k) =>
        hexOf(image.rgb(at(ORIGIN.x) + Math.round(((k + 0.5) * at(CELL.w)) / 8), y)),
      ),
      right: hexOf(image.rgb(at(ORIGIN.x + table.cols * CELL.w + 3), y)),
      below: hexOf(image.rgb(at(ORIGIN.x + 36), at(ORIGIN.y + table.rows * CELL.h + 3))),
      beyondGrid: hexOf(image.rgb(at(ORIGIN.x + table.cols * CELL.w + 20), y)),
    };
  }
  const along: Record<string, string[]> = {};
  for (const edge of edgesOf(table)) {
    if (dashedAt(table, edge))
      along[`${edge.axis}${String(edge.r)},${String(edge.c)}`] = alongOf(image, edge);
  }
  const diagonal: Record<string, string> = {};
  for (const [x, y] of diagonalPoints(table))
    diagonal[`${String(x)},${String(y)}`] = hexOf(image.rgb(x, y));
  return { paints, glyphs, profiles, samples, along, diagonal };
}

/** The paint at visual position (r, c). */
export function paintAt(o: Observed, r: number, c: number): string {
  const paint = o.pixels.paints[r]?.[c];
  if (paint === null || paint === undefined)
    throw new Error(`${o.deck.id}#${String(o.slide)}: no one paint at ${String(r)},${String(c)}`);
  return paint;
}

/** The pixels across visual grid edge `edge`. */
export function profileAt(o: Observed, edge: Edge): readonly string[] {
  const profile = o.pixels.profiles.get(`${edge.axis}${String(edge.r)},${String(edge.c)}`);
  if (profile === undefined)
    throw new Error(
      `${o.deck.id}#${String(o.slide)}: no edge ${edge.axis}${String(edge.r)},${String(edge.c)}`,
    );
  return profile;
}

function load(dir: string): {
  inputs: Inputs;
  slides: Observed[];
  readings: Map<string, DeckReading>;
} {
  const inputs = JSON.parse(readFileSync(join(dir, 'cascade-inputs.json'), 'utf8')) as Inputs;
  const wanted = new Set(inputs.decks.map((d) => d.id));
  const decks = allDecks().filter((d) => wanted.has(d.id));
  if (decks.length !== wanted.size)
    throw new Error('cascade-inputs.json names decks probes.ts does not build');
  const readings = new Map<string, DeckReading>();
  const controls = readdirSync(join(dir, 'readings'))
    .map((f) => /^(control-[a-z0-9]+(?:-again)?)\.json$/.exec(f)?.[1])
    .filter((id): id is string => id !== undefined);
  for (const id of new Set([...wanted, 'control-again', ...controls])) {
    const path = join(dir, 'readings', `${id}.json`);
    if (!existsSync(path)) throw new Error(`no readings for ${id}: run read.ps1 to the end`);
    const reading = JSON.parse(readFileSync(path, 'utf8')) as DeckReading;
    if (!reading.opened || reading.repaired !== false) {
      throw new Error(
        `${id} was ${reading.opened ? 'repaired' : 'refused'}: ${String(reading.error)}`,
      );
    }
    readings.set(id, reading);
  }
  // Every session that read a deck read the control first and last under its own names.
  const sessionOf = (id: string): string => {
    const s = readings.get(id)?.session;
    if (s === undefined) throw new Error(`no reading ${id}`);
    return `${String(s.pid)}@${s.started}`;
  };
  const pairs = [
    ['control', 'control-again'],
    ...controls.filter((c) => !c.endsWith('-again')).map((c) => [c, `${c}-again`]),
  ] as const;
  const bracketed = new Map<string, string>();
  for (const [first, last] of pairs) {
    if (!readings.has(last))
      throw new Error(`${first} has no ${last}: that session did not finish`);
    if (sessionOf(first) !== sessionOf(last))
      throw new Error(`${first} and ${last} are from different sessions`);
    bracketed.set(sessionOf(first), first);
  }
  for (const id of wanted) {
    if (!bracketed.has(sessionOf(id)))
      throw new Error(`${id} was read in a session that read no control first and last`);
  }

  const slides: Observed[] = [];
  for (const deck of decks) {
    const reading = readings.get(deck.id);
    if (reading === undefined || reading.slides.length !== deck.slides.length) {
      throw new Error(
        `${deck.id}: ${String(reading?.slides.length)} slides read, ${String(deck.slides.length)} built`,
      );
    }
    // Decoding every export takes minutes, so each deck's pixels are kept beside its readings.
    const cachePath = join(dir, 'pixels', `${deck.id}.json`);
    const cached = existsSync(cachePath)
      ? (JSON.parse(readFileSync(cachePath, 'utf8')) as CachedPixels[])
      : null;
    const fresh: CachedPixels[] = [];
    let stale = false;
    deck.slides.forEach((table, slide) => {
      const read = reading.slides[slide];
      if (read === undefined) throw new Error(`${deck.id}#${String(slide)} has no reading`);
      const where = `${deck.id}#${String(slide)}`;
      const com = parseCom(read.cells, where);
      if (com.rows !== table.rows || com.cols !== table.cols)
        throw new Error(`${where}: COM reads ${String(com.rows)}x${String(com.cols)}`);
      const expected =
        table.style !== null && table.style.startsWith('{C9') ? '' : (table.style ?? '');
      if (read.style.toUpperCase() !== expected.toUpperCase())
        throw new Error(`${where}: COM names style ${read.style}`);
      const [left, top, width, height] = read.frame;
      if (
        left !== ORIGIN.x ||
        top !== ORIGIN.y ||
        width !== table.cols * CELL.w ||
        height !== table.rows * CELL.h
      ) {
        throw new Error(
          `${where}: the table was drawn at ${read.frame.join(',')}, not on the grid the pixels assume`,
        );
      }
      const png = readFileSync(join(dir, read.png));
      const source = createHash('sha256').update(png).digest('hex');
      const kept = cached?.[slide];
      // A cached slide without dash or diagonal samples stands where its table has neither.
      const sampled =
        (kept?.along !== undefined && kept.diagonal !== undefined) ||
        (!edgesOf(table).some((edge) => dashedAt(table, edge)) &&
          diagonalPoints(table).length === 0);
      const reuse = kept !== undefined && kept.source === source && sampled;
      if (!reuse || kept.along === undefined) stale = true;
      const pixels: Pixels = reuse
        ? {
            ...kept,
            along: kept.along ?? {},
            diagonal: kept.diagonal ?? {},
            profiles: new Map(kept.profiles),
          }
        : pixelsOf(readPng(png), deck, table);
      fresh.push({ ...pixels, source, profiles: [...pixels.profiles] });
      slides.push({ deck, slide, table, com, pixels });
    });
    if (stale || cached?.length !== deck.slides.length) {
      mkdirSync(join(dir, 'pixels'), { recursive: true });
      writeFileSync(cachePath, JSON.stringify(fresh));
    }
  }
  return { inputs, slides, readings };
}

/** The instruments agree with themselves and with each other, or nothing is scored. */
function guards(
  dir: string,
  inputs: Inputs,
  slides: readonly Observed[],
  readings: Map<string, DeckReading>,
): void {
  // Every read of the control, first and last in every session, is the same COM and the same export.
  const control = readings.get('control');
  if (control === undefined) throw new Error('no control readings');
  const reads = [...readings.keys()].filter((id) => id === 'control' || id.startsWith('control-'));
  for (const id of reads) {
    const other = readings.get(id);
    control.slides.forEach((s, k) => {
      const later = other?.slides[k];
      if (later?.cells !== s.cells)
        throw new Error(`control #${String(k)}: COM differs in ${id} from the first read`);
      const a = readFileSync(join(dir, s.png));
      const b = readFileSync(join(dir, later.png));
      if (Buffer.compare(a, b) !== 0)
        throw new Error(`control #${String(k)}: the export differs in ${id} from the first read`);
    });
  }
  for (const id of inputs.twins) {
    const first = readings.get(id)?.slides[0];
    if (first?.bmp === null || first?.bmp === undefined) throw new Error(`${id} has no BMP twin`);
    const png = readPng(readFileSync(join(dir, first.png)));
    const bmp = readBmp(readFileSync(join(dir, first.bmp)));
    for (let y = 0; y < bmp.height; y += 1) {
      for (let x = 0; x < bmp.width; x += 1) {
        if (hexOf(png.rgb(x, y)) !== hex(bmp.pixel(x, y)))
          throw new Error(`${id}: PNG and BMP differ at ${String(x)},${String(y)}`);
      }
    }
  }
  // Theme C's gradients have no one paint per cell; the background question samples them itself.
  for (const o of slides.filter((x) => x.deck.group !== 'background')) {
    const where = `${o.deck.id}#${String(o.slide)}`;
    for (let r = 0; r < o.table.rows; r++) {
      for (let c = 0; c < o.table.cols; c++) {
        const cell = o.com.cells[r]?.[c];
        if (cell === undefined) throw new Error(`${where}: no COM cell ${String(r)},${String(c)}`);
        if ((cell.text === 'empty') !== c >= cellsInRow(o.table, r))
          throw new Error(`${where}: cell ${String(r)},${String(c)} is ${cell.text} to COM`);
        // An rtl table is drawn mirrored: logical column c at visual column cols - 1 - c.
        const v = o.table.rtl === true ? o.table.cols - 1 - c : c;
        const paint = paintAt(o, r, v);
        const [rgb, transparency] = cell.fill.split('/');
        if (transparency === '0' && rgb !== paint) {
          throw new Error(
            `${where}: cell ${String(r)},${String(c)} COM fill ${cell.fill} is drawn ${paint}`,
          );
        }
        const spec = o.table.cells?.[`${String(r)},${String(c)}`];
        // A diagonal crosses the text's box, and a covered position draws no text of its own.
        if (/lnTlToBr|lnBlToTr/.test(spec?.tcPr ?? '') || /Merge=/.test(spec?.attrs ?? ''))
          continue;
        // A position no `a:tc` reached is padded empty (ADR 0056): it has no text to draw.
        if (c >= cellsInRow(o.table, r)) continue;
        const glyph = o.pixels.glyphs[r]?.[v] ?? { rgb: null, ink: 0 };
        const text = cell.text.split('/')[0];
        if (glyph.rgb === null ? text !== paint : glyph.rgb !== text) {
          throw new Error(
            `${where}: cell ${String(r)},${String(c)} COM text ${String(text)} is drawn ${String(glyph.rgb)} on ${paint}`,
          );
        }
      }
    }
  }
}

/* -------------------------------------------------------------------------- */
/* the cascade: fills, text and edges of styled tables                        */
/* -------------------------------------------------------------------------- */

const MAIN_GROUPS = new Set(['control', 'sweep', 'ends', 'sides']);

/** A slide PowerPoint drew from a built-in with nothing direct on it: what the cascade alone decides. */
function isStyled(o: Observed): boolean {
  if (!MAIN_GROUPS.has(o.deck.group)) return false;
  for (const styled of styledTables([{ ...o.deck, slides: [o.table] }]))
    return styled.table === o.table;
  return false;
}

function styleOf(
  o: Observed,
): ReturnType<typeof styledTables> extends Generator<infer T> ? T : never {
  for (const styled of styledTables([{ ...o.deck, slides: [o.table] }])) return styled;
  throw new Error(`${o.deck.id}#${String(o.slide)} is not styled`);
}

/** COM's fill, or the first character's colour, bold and face, per cell of every styled slide. */
export function comRows(slides: readonly Observed[], unit: 'fill' | 'text'): Row<Scenario>[] {
  const rows: Row<Scenario>[] = [];
  for (const o of slides) {
    if (!isStyled(o)) continue;
    const { style, table } = styleOf(o);
    const palette = paletteOf(o.deck.theme);
    for (let r = 0; r < table.rows; r++) {
      for (let c = 0; c < table.cols; c++) {
        const cell = o.com.cells[r]?.[c];
        if (cell === undefined) throw new Error('unreachable: guards checked every cell');
        rows.push({
          key: cellKey(o.deck.id, o.slide, r, c, unit),
          observed: unit === 'fill' ? cell.fill : cell.text.split('/').slice(0, 3).join('/'),
          predict: (s) => predictCom(style, table, r, c, unit, s.reading, s.source, palette),
        });
      }
    }
  }
  return rows;
}

/**
 * How much of each pixel across a double line its strokes cover, per width, from every edge where
 * the baseline draws one double line alone: quarters, each checked to reproduce its pixel exactly,
 * and one pattern per width wherever such a line was drawn.
 */
export function measureDouble(slides: readonly Observed[]): Compound {
  const found = new Map<number, (number | null)[]>();
  for (const o of slides) {
    if (!isStyled(o)) continue;
    const { style, table } = styleOf(o);
    const palette = paletteOf(o.deck.theme);
    for (const edge of edgesOf(table)) {
      const lines = edgeLines(style, table, edge, BASELINE.reading, BASELINE.edges, palette);
      const [only] = lines;
      if (lines.length !== 1 || only?.cmpd !== 'dbl') continue;
      const profile = profileAt(o, edge);
      const half = (only.weight * SCALE) / 2;
      const from = PROFILE_REACH - half;
      const where = `${o.deck.id}#${String(o.slide)} ${edge.axis}${String(edge.r)},${String(edge.c)}`;
      const coverage = Array.from({ length: 2 * half }, (_, k): number | null => {
        const under = rgbOf(profile[from + k < PROFILE_REACH ? 0 : profile.length - 1] ?? '');
        const drawn = profile[from + k] ?? '';
        const ink = rgbOf(only.rgb);
        const spread = under.map((u, i) => Math.abs(u - (ink[i] ?? 0)));
        const i = spread.indexOf(Math.max(...spread));
        if ((spread[i] ?? 0) === 0) return null;
        const c =
          Math.round(
            (4 * ((under[i] ?? 0) - (rgbOf(drawn)[i] ?? 0))) / ((under[i] ?? 0) - (ink[i] ?? 0)),
          ) / 4;
        const side = hexRgb(under);
        const whole = c === 0 || c === 1;
        const predicted = c === 0 ? side : c === 1 ? only.rgb : blend(only.rgb, side, c);
        if (whole ? predicted !== drawn : !near(predicted, drawn)) {
          throw new Error(
            `${where}: pixel ${String(k)} of ${lineText(lines)} is ${drawn}, which no quarter coverage draws`,
          );
        }
        return c;
      });
      const known = found.get(only.weight) ?? coverage;
      coverage.forEach((c, k) => {
        if (c === null) return;
        if (known[k] !== null && known[k] !== c)
          throw new Error(
            `${where}: a ${String(only.weight)}-pt double line covers pixel ${String(k)} differently`,
          );
        known[k] = c;
      });
      found.set(only.weight, known);
    }
  }
  const out = new Map<number, readonly number[]>();
  for (const [weight, coverage] of found) {
    if (coverage.includes(null))
      throw new Error(`a ${String(weight)}-pt double line was never seen whole`);
    out.set(weight, coverage as number[]);
  }
  return out;
}

/** Each grid edge's pixels, stripped of the paints either side, per styled slide. */
export function edgeRows(slides: readonly Observed[], double: Compound): Row<Scenario>[] {
  const rows: Row<Scenario>[] = [];
  const memo = new Map<string, string>();
  for (const o of slides) {
    if (!isStyled(o)) continue;
    const { style, table } = styleOf(o);
    const palette = paletteOf(o.deck.theme);
    for (const edge of edgesOf(table)) {
      const profile = profileAt(o, edge);
      const a = profile[0] ?? '';
      const b = profile.at(-1) ?? '';
      const observed = describe(profile);
      rows.push({
        key: edgeKey(o.deck.id, o.slide, edge),
        observed,
        predict: (s) => {
          const lines = edgeLines(style, table, edge, s.reading, s.edges, palette);
          const id = `${lineText(lines)}|${a}|${b}|${observed}`;
          let seen = memo.get(id);
          if (seen === undefined) {
            seen = drawnAs(lines, a, b, double, observed);
            memo.set(id, seen);
          }
          return seen;
        },
      });
    }
  }
  return rows;
}

interface Axis {
  readonly name: string;
  readonly values: readonly string[];
  readonly set: (s: Scenario, value: string) => Scenario;
}

const AXES: readonly Axis[] = [
  {
    name: 'applicability',
    values: APPLICABILITY,
    set: (s, v) => ({ ...s, reading: { ...s.reading, applicability: v as never } }),
  },
  {
    name: 'order',
    values: Object.keys(ORDERS),
    set: (s, v) => ({ ...s, reading: { ...s.reading, order: v as never } }),
  },
  {
    name: 'merge',
    values: MERGES,
    set: (s, v) => ({ ...s, reading: { ...s.reading, merge: v as never } }),
  },
  { name: 'edges', values: EDGE_MODELS, set: (s, v) => ({ ...s, edges: v as never }) },
  {
    name: 'source',
    values: ['explicit', 'fontRef'],
    set: (s, v) => ({ ...s, source: v as never }),
  },
];

const valueOf = (s: Scenario, axis: string): string =>
  axis === 'edges'
    ? s.edges
    : axis === 'source'
      ? s.source
      : String(s.reading[axis as keyof Scenario['reading']]);

/**
 * The reading that fits every row, found one axis at a time from the baseline, and each axis
 * re-scored with the others held at it: the finding per axis, and what each rival got wrong.
 */
export function cascadeFindings(
  property: string,
  rows: readonly Row<Scenario>[],
  axes: readonly string[],
): Finding[] {
  let best = BASELINE;
  for (let pass = 0; pass < 3; pass++) {
    let moved = false;
    for (const axis of AXES.filter((a) => axes.includes(a.name))) {
      const tallies = axis.values.map((v) => {
        const scenario = axis.set(best, v);
        return {
          scenario,
          fits: rows.filter((row) => row.predict(scenario) === row.observed).length,
        };
      });
      const top = [...tallies].sort((p, q) => q.fits - p.fits)[0];
      if (top !== undefined && valueOf(top.scenario, axis.name) !== valueOf(best, axis.name)) {
        const current = tallies.find(
          (t) => valueOf(t.scenario, axis.name) === valueOf(best, axis.name),
        );
        if (top.fits > (current?.fits ?? -1)) {
          best = top.scenario;
          moved = true;
        }
      }
    }
    if (!moved) break;
  }
  return AXES.filter((a) => axes.includes(a.name)).map((axis) =>
    score(
      `${property}: ${axis.name}`,
      axis.values.map((v): [string, Scenario] => [v, axis.set(best, v)]),
      rows,
    ),
  );
}

/* -------------------------------------------------------------------------- */
/* paint: the table background, and translucent fills over it                 */
/* -------------------------------------------------------------------------- */

type Rgb = readonly [number, number, number];

const rgbOf = (h: string): Rgb => [
  Number.parseInt(h.slice(0, 2), 16),
  Number.parseInt(h.slice(2, 4), 16),
  Number.parseInt(h.slice(4, 6), 16),
];
const hexRgb = (c: Rgb): string =>
  c.map((v) => v.toString(16).toUpperCase().padStart(2, '0')).join('');

/** Thousandths rounded to a whole, ties down (ADR 0021). */
const roundDown = (n: number): number => {
  const q = Math.floor(n / 1000);
  return n % 1000 > 500 ? q + 1 : q;
};

export type Compositing = 'whole' | 'perTerm';

/** sRGB `a*c + (1-a)*bg` on the quantised colour: the sum rounded (ADR 0021), or each term. */
export function over(under: Rgb, fill: FillReport, compositing: Compositing = 'perTerm'): Rgb {
  if (!fill.visible) return under;
  const top = rgbOf(fill.rgb);
  const a = Math.round((1 - fill.transparency) * 1000);
  return under.map((u, i) =>
    compositing === 'whole'
      ? roundDown(a * (top[i] ?? 0) + (1000 - a) * u)
      : roundDown(a * (top[i] ?? 0)) + roundDown((1000 - a) * u),
  ) as unknown as Rgb;
}

export interface Paint {
  readonly fills: FillPaint;
  readonly background: 'drawn' | 'ignored';
  readonly compositing: Compositing;
}

/** What a styled cell's paint composites to over the page, for one reading of how it is painted. */
export function paintOf(o: Observed, r: number, c: number, paint: Paint): string {
  const { style, table } = styleOf(o);
  const palette = paletteOf(o.deck.theme);
  const parts = partsGrid(BASELINE.reading.applicability, table)[r]?.[c] ?? new Set();
  let colour = rgbOf(PAGE);
  if (paint.background === 'drawn' && style.background?.fill !== undefined) {
    colour = over(colour, predictFill(style.background.fill, palette), paint.compositing);
  }
  const fills =
    paint.fills === 'stacked'
      ? fillStack(style, parts, BASELINE.reading.order)
      : [fillWinner(style, parts, BASELINE.reading)?.fill].filter((f) => f !== undefined);
  for (const fill of fills) colour = over(colour, predictFill(fill, palette), paint.compositing);
  return hexRgb(colour);
}

export function paintRows(slides: readonly Observed[]): Row<Paint>[] {
  const rows: Row<Paint>[] = [];
  for (const o of slides) {
    if (!isStyled(o)) continue;
    for (let r = 0; r < o.table.rows; r++) {
      for (let c = 0; c < o.table.cols; c++) {
        rows.push({
          key: cellKey(o.deck.id, o.slide, r, c, 'paint'),
          observed: paintAt(o, r, c),
          predict: (p) => paintOf(o, r, c, p),
        });
      }
    }
  }
  return rows;
}

export const PAINTS: readonly [string, Paint][] = FILL_PAINTS.flatMap((fills) =>
  (['drawn', 'ignored'] as const).flatMap((background) =>
    (['whole', 'perTerm'] as const).map((compositing): [string, Paint] => [
      `${fills}/${background}/${compositing}`,
      { fills, background, compositing },
    ]),
  ),
);

/* -------------------------------------------------------------------------- */
/* the default grid: a table naming no built-in                               */
/* -------------------------------------------------------------------------- */

type Grid = 'layer' | 'floor' | 'endsBold';

/** The null style's text: `tx1` in the minor face above the floor, the floor, or bold at the ends. */
export function gridRows(slides: readonly Observed[]): Row<Grid>[] {
  const rows: Row<Grid>[] = [];
  for (const o of slides.filter(
    (x) => x.deck.id === 'sweep-none' || x.deck.id === 'sweep-unknown',
  )) {
    const palette = paletteOf(o.deck.theme);
    const tx1 = colorHex({ space: 'scheme', name: 'tx1', transforms: [] }, palette);
    const minor = palette.theme.fonts.minor.latin ?? '';
    const floor = floorText(palette);
    const ends = partsGrid('spec', o.table);
    for (let r = 0; r < o.table.rows; r++) {
      for (let c = 0; c < o.table.cols; c++) {
        const cell = o.com.cells[r]?.[c];
        if (cell === undefined) continue;
        const atEnd = [...(ends[r]?.[c] ?? [])].some((p) => /^(first|last)(Row|Col)$/.test(p));
        rows.push({
          key: cellKey(o.deck.id, o.slide, r, c, 'text'),
          observed: cell.text,
          predict: (g) =>
            g === 'floor'
              ? `${floor.rgb}/-/${floor.face}`
              : `${tx1}/${g === 'endsBold' && atEnd ? 'b' : '-'}/${minor}`,
        });
      }
    }
  }
  return rows;
}

/** The null style's lines: a fixed 1-pt black grid, one in `tx1`, or none. */
export function gridEdgeRows(slides: readonly Observed[]): Row<'black' | 'tx1' | 'none'>[] {
  const rows: Row<'black' | 'tx1' | 'none'>[] = [];
  for (const o of slides.filter(
    (x) => x.deck.id === 'sweep-none' || x.deck.id === 'sweep-unknown',
  )) {
    const palette = paletteOf(o.deck.theme);
    const tx1 = colorHex({ space: 'scheme', name: 'tx1', transforms: [] }, palette);
    for (const edge of edgesClear(o.table)) {
      const profile = profileAt(o, edge);
      const [a, b] = [profile[0] ?? '', profile.at(-1) ?? ''];
      const observed = describe(profile);
      rows.push({
        key: edgeKey(o.deck.id, o.slide, edge),
        observed,
        predict: (g) =>
          drawnAs(
            g === 'none' ? [] : [{ rgb: g === 'black' ? '000000' : tx1, weight: 1, cmpd: 'sng' }],
            a,
            b,
            new Map(),
            observed,
          ),
      });
    }
  }
  return rows;
}

/* -------------------------------------------------------------------------- */
/* where the table's text sits in the text cascade                            */
/* -------------------------------------------------------------------------- */

export const POSITIONS = ['aboveRun', 'belowRun', 'belowPara', 'belowList', 'floor'] as const;
export type Position = (typeof POSITIONS)[number];
export const TERMINI = [
  'otherThenDefault',
  'defaultThenOther',
  'otherOnly',
  'defaultOnly',
  'chain',
] as const;
export type Terminus = (typeof TERMINI)[number];
export const REACHES = ['everyLevel', 'firstLevel'] as const;
export type Reach = (typeof REACHES)[number];
export const FLOORS = ['other', 'frame'] as const;
export type Floor = (typeof FLOORS)[number];
export interface TextPlace {
  readonly position: Position;
  readonly terminus: Terminus;
  /** Whether the table's text reaches a paragraph below the first level. */
  readonly reach: Reach;
  /** Under a master with no `p:txStyles`: the built-in `other` floor, or the frame's own bucket's. */
  readonly floor: Floor;
}

/** What a level declares: colour, size in hundredths, face, bold; `bullet` for p:bodyStyle alone. */
interface Declared {
  readonly rgb?: string;
  readonly sz?: number;
  readonly face?: string;
  readonly b?: boolean;
  readonly bullet?: boolean;
}

/** The deck's `tx1`, and its theme's two faces: what a floor names by reference. */
export interface FloorBase {
  readonly rgb: string;
  readonly minor: string;
  readonly major: string;
}

/** What PowerPoint lays text out with when no level declares it (ADR 0027), for `ph`'s frame. */
function floorDeclared(
  place: TextPlace,
  ph: string | undefined,
  txStyles: boolean,
  lvl: number,
  base: FloorBase,
): Declared {
  let level = TEXT_FLOOR;
  if (!txStyles) {
    let bucket: 'title' | 'body' | 'other' = 'other';
    if (place.floor === 'frame' && ph !== undefined)
      bucket = /type="(title|ctrTitle)"/.test(ph) ? 'title' : 'body';
    level = BUILTIN_TEXT_STYLES[bucket][Math.min(lvl, 8)] ?? TEXT_FLOOR;
  }
  const face = level.typeface === '+mj-lt' ? base.major : base.minor;
  return { rgb: base.rgb, face, sz: level.sz, ...(level.bullet ? { bullet: true } : {}) };
}

export function ladderPrediction(
  cellLevels: readonly TextLevel[],
  packageLevels: readonly TextLevel[],
  lvl: number,
  layer: Declared,
  place: TextPlace,
  ph: string | undefined,
  base: FloorBase,
): string {
  const declared = (level: TextLevel): Declared => {
    const { hex: h, sz, face } = LEVELS[level];
    return { rgb: h, sz, face, b: false, ...(level === 'bodyStyle' ? { bullet: true } : {}) };
  };
  // The package's levels declare lvl1 only; the cell's list declares the paragraph's own level.
  const packaged = (level: TextLevel): boolean => lvl === 0 && packageLevels.includes(level);
  const terminus: TextLevel[] = (() => {
    switch (place.terminus) {
      case 'otherThenDefault':
        return ['otherStyle', 'defaultTextStyle'];
      case 'defaultThenOther':
        return ['defaultTextStyle', 'otherStyle'];
      case 'otherOnly':
        return ['otherStyle'];
      case 'defaultOnly':
        return ['defaultTextStyle'];
      case 'chain':
        return ph !== undefined ? ['layoutPh', 'masterPh', 'bodyStyle'] : ['defaultTextStyle'];
    }
  })();
  const cell: Declared[] = cellLevels.map(declared);
  const pkg: Declared[] = terminus.filter(packaged).map(declared);
  const at =
    place.reach === 'firstLevel' && lvl > 0
      ? -1
      : { aboveRun: 0, belowRun: 1, belowPara: 2, belowList: 3, floor: 99 }[place.position];
  const cellOrder = (['run', 'para', 'list'] as const).filter((l) => cellLevels.includes(l));
  const walk: Declared[] = [];
  let inserted = false;
  cellOrder.forEach((level, k) => {
    const rank = { run: 0, para: 1, list: 2 }[level];
    if (!inserted && at >= 0 && rank >= at) {
      walk.push(layer);
      inserted = true;
    }
    walk.push(cell[k] ?? {});
  });
  if (!inserted && at >= 0 && at !== 99) walk.push(layer);
  walk.push(...pkg);
  if (!inserted && at === 99) walk.push(layer);
  const txStyles = packageLevels.includes('otherStyle') || packageLevels.includes('bodyStyle');
  walk.push(floorDeclared(place, ph, txStyles, lvl, base));
  const first = <K extends keyof Declared>(key: K): Declared[K] | undefined =>
    walk.find((d) => d[key] !== undefined)?.[key];
  return [
    first('rgb') ?? '?',
    first('b') === true ? 'b' : '-',
    first('face') ?? '?',
    String((first('sz') ?? 0) / 100),
    first('bullet') === true ? 'bullet' : '-',
  ].join('/');
}

export function textRows(slides: readonly Observed[]): Row<TextPlace>[] {
  const rows: Row<TextPlace>[] = [];
  for (const o of slides.filter((x) => x.deck.group === 'text')) {
    const pkg = TEXT_PACKAGES.find((p) => p.id === o.deck.id);
    if (pkg === undefined) throw new Error(`no text package ${o.deck.id}`);
    const palette = paletteOf(o.deck.theme);
    const minor = palette.theme.fonts.minor.latin ?? '';
    const base: FloorBase = {
      rgb: colorHex({ space: 'scheme', name: 'tx1', transforms: [] }, palette),
      minor,
      major: palette.theme.fonts.major.latin ?? '',
    };
    const named = (slot: 'lt1' | 'dk1' | 'tx1'): string =>
      colorHex({ space: 'scheme', name: slot, transforms: [] }, palette);
    for (let r = 0; r < o.table.rows; r++) {
      LADDER.forEach((rung, c) => {
        const cell = o.com.cells[r]?.[c];
        if (cell === undefined) return;
        // The style's text for this cell: MS2-A1's header and body, and tx1 for both grids.
        const layer: Declared =
          o.slide === 0
            ? r === 0
              ? { rgb: named('lt1'), b: true, face: minor }
              : { rgb: named('dk1'), face: minor }
            : { rgb: named('tx1'), face: minor };
        const [rgb, b, face, , size, bullet] = cell.text.split('/');
        rows.push({
          key: cellKey(o.deck.id, o.slide, r, c, 'text'),
          observed: [rgb, b, face, size, bullet].join('/'),
          predict: (place) =>
            ladderPrediction(rung.levels, pkg.levels, rung.lvl, layer, place, pkg.ph, base),
        });
      });
    }
  }
  return rows;
}

/* -------------------------------------------------------------------------- */
/* direct formatting, merged cells, rtl, tblPr fill, the table background     */
/* -------------------------------------------------------------------------- */

/** Every edge and interior of every slide in `group`, as the pixels show them, for the fixture. */
export function recordPixels(o: Observed): {
  interiors: string[][];
  edges: Record<string, string>;
} {
  const table: Shape = o.table;
  const interiors = Array.from({ length: table.rows }, (_, r) =>
    Array.from({ length: table.cols }, (_, c) => paintAt(o, r, c)),
  );
  const edges: Record<string, string> = {};
  for (const edge of edgesOf(table))
    edges[`${edge.axis}${String(edge.r)},${String(edge.c)}`] = describe(profileAt(o, edge));
  return { interiors, edges };
}

/** A direct `a:tcPr` line as written: `none` for an explicit `a:noFill`, and whatever attributes it has. */
interface Written {
  readonly none: boolean;
  readonly w: number | undefined;
  readonly rgb: string | undefined;
  readonly cmpd: string | undefined;
  readonly dash: boolean;
}

function written(tcPr: string, tag: string): Written | undefined {
  const m = new RegExp(`<a:${tag}( [^>]*)?(/>|>(.*?)</a:${tag}>)`).exec(tcPr);
  if (m === null) return undefined;
  const attrs = m[1] ?? '';
  const body = m[3] ?? '';
  const w = /w="(\d+)"/.exec(attrs)?.[1];
  return {
    none: /<a:noFill\/>/.test(body),
    w: w === undefined ? undefined : Number(w),
    rgb: /<a:srgbClr val="([0-9A-F]{6})"/.exec(body)?.[1],
    cmpd: /cmpd="(\w+)"/.exec(attrs)?.[1],
    dash: /<a:prstDash val="(?!solid)/.test(body),
  };
}

export type Completion = 'replace' | 'merge';

/** The width a line states nowhere: the shape default, 9525 EMU (paint's `DEFAULT_LINE_WIDTH`). */
const LINE_DEFAULT_EMU = 9525;

/** A written line completed: on its own, or attribute by attribute over the style's line there. */
function completed(line: Written, style: DrawnLine | undefined, how: Completion): DrawnLine | null {
  if (line.none) return null;
  const rgb = line.rgb ?? (how === 'merge' ? style?.rgb : undefined);
  if (rgb === undefined) return null;
  const weight =
    line.w === undefined
      ? how === 'merge' && style !== undefined
        ? style.weight
        : LINE_DEFAULT_EMU / 12700
      : line.w / 12700;
  const cmpd = line.cmpd ?? (how === 'merge' ? (style?.cmpd ?? 'sng') : 'sng');
  return { rgb, weight, cmpd };
}

export type Direct = 'owner' | 'ownerThenOther' | 'later' | 'earlier' | 'heavier' | 'style';
export interface DirectReading {
  readonly rule: Direct;
  readonly completion: Completion;
}

/** Edges where a cell writes a line on the side facing it, and what each reading draws there. */
export function directRows(slides: readonly Observed[], double: Compound): Row<DirectReading>[] {
  const rows: Row<DirectReading>[] = [];
  for (const o of slides.filter((x) => x.deck.group === 'direct')) {
    const palette = paletteOf(o.deck.theme);
    const style = o.table.style === null ? null : styleById(o.table.style);
    const grid = partsGrid('spec', o.table);
    const inside = (rc: readonly number[]): boolean =>
      (rc[0] ?? -1) >= 0 &&
      (rc[1] ?? -1) >= 0 &&
      (rc[0] ?? 0) < o.table.rows &&
      (rc[1] ?? 0) < o.table.cols;
    const tcPr = (rc: readonly number[]): string =>
      o.table.cells?.[`${String(rc[0])},${String(rc[1])}`]?.tcPr ?? '';
    for (const edge of edgesClear(o.table)) {
      const before = edge.axis === 'h' ? [edge.r - 1, edge.c] : [edge.r, edge.c - 1];
      const after = [edge.r, edge.c];
      const [tagA, tagB] = edge.axis === 'h' ? ['lnB', 'lnT'] : ['lnR', 'lnL'];
      const lineA = inside(before) ? written(tcPr(before), tagA ?? '') : undefined;
      const lineB = inside(after) ? written(tcPr(after), tagB ?? '') : undefined;
      if (lineA === undefined && lineB === undefined) continue;
      // A dash is not in an edge's descriptor: whether the midpoint lands on a dash is chance.
      if (lineA?.dash === true || lineB?.dash === true) continue;
      const profile = profileAt(o, edge);
      const [a, b] = [profile[0] ?? '', profile.at(-1) ?? ''];
      const styleLines =
        style === null
          ? [{ rgb: '000000', weight: 1, cmpd: 'sng' }]
          : edgeLinesBetween(
              style,
              inside(before) ? (grid[before[0] ?? 0]?.[before[1] ?? 0] ?? null) : null,
              inside(after) ? (grid[after[0] ?? 0]?.[after[1] ?? 0] ?? null) : null,
              edge.axis,
              BASELINE.reading,
              palette,
            );
      const observed = describe(profile);
      rows.push({
        key: edgeKey(o.deck.id, o.slide, edge),
        observed,
        predict: ({ rule, completion }) => {
          const side = (line: Written | undefined): DrawnLine[] | undefined => {
            if (line === undefined) return undefined;
            const done = completed(line, styleLines[0], completion);
            return done === null ? [] : [done];
          };
          const A = side(lineA);
          const B = side(lineB);
          let lines: DrawnLine[];
          if (rule === 'style') lines = styleLines;
          else if (rule === 'owner') lines = (inside(before) ? A : B) ?? styleLines;
          else if (rule === 'ownerThenOther') {
            lines = inside(before)
              ? (A ?? (styleLines.length === 0 ? (B ?? []) : styleLines))
              : (B ?? styleLines);
          } else if (rule === 'later') lines = B ?? A ?? styleLines;
          else if (rule === 'earlier') lines = A ?? B ?? styleLines;
          else {
            const wa = A?.[0]?.weight ?? -1;
            const wb = B?.[0]?.weight ?? -1;
            lines = (wb > wa ? B : A) ?? B ?? styleLines;
          }
          return drawnAs(lines, a, b, double, observed);
        },
      });
    }
  }
  return rows;
}

/** Cells with a direct fill: `a:tcPr` wins over the style, or the style over it. */
export function directFillRows(slides: readonly Observed[]): Row<'tcPr' | 'style'>[] {
  const rows: Row<'tcPr' | 'style'>[] = [];
  for (const o of slides.filter((x) => x.deck.group === 'direct')) {
    const palette = paletteOf(o.deck.theme);
    for (const [at, spec] of Object.entries(o.table.cells ?? {})) {
      const tcPr = spec.tcPr ?? '';
      const own = /<a:tcPr>(<a:solidFill>.*?<\/a:solidFill>|<a:noFill\/>)/.exec(tcPr)?.[1];
      if (own === undefined) continue;
      const [r, c] = at.split(',').map(Number) as [number, number];
      const cell = o.com.cells[r]?.[c];
      if (cell === undefined) continue;
      const direct =
        own === '<a:noFill/>' ? 'none' : `${/val="([0-9A-F]{6})"/.exec(own)?.[1] ?? '?'}/0`;
      const styled =
        o.table.style === null
          ? 'none'
          : predictCom(
              styleById(o.table.style),
              o.table,
              r,
              c,
              'fill',
              BASELINE.reading,
              'explicit',
              palette,
            );
      rows.push({
        key: cellKey(o.deck.id, o.slide, r, c, 'fill'),
        observed: cell.fill,
        predict: (d) => (d === 'tcPr' ? direct : styled),
      });
    }
  }
  return rows;
}

/** A cell's paint over the page and the style's `tblBg` (the paint question's winner). */
function underCell(
  style: TableStyle,
  palette: Palette,
  parts: ReadonlySet<TableStylePartName>,
): string {
  let colour = rgbOf(PAGE);
  if (style.background?.fill !== undefined)
    colour = over(colour, predictFill(style.background.fill, palette));
  return hexRgb(
    over(colour, predictFill(fillWinner(style, parts, BASELINE.reading)?.fill ?? null, palette)),
  );
}

export const MERGED = [
  'anchor',
  'position',
  'union',
  'unionByKind',
  'endsPooled',
  'endsPooledByKind',
] as const;
/** Which parts a merged cell takes: its anchor's, each position's, all pooled, or the ends pooled. */
export type Merged = (typeof MERGED)[number];
/** The merged readings whose end parts offer the edge their kind faces. */
const BY_KIND: ReadonlySet<Merged> = new Set(['unionByKind', 'endsPooledByKind']);

/** The parts position (r, c) takes under reading `m`, or `null` beyond the table. */
function mergedParts(
  table: TableSpec,
  grid: ReturnType<typeof partsGrid>,
  owner: readonly (readonly [number, number])[][],
  r: number,
  c: number,
  m: Merged,
): ReadonlySet<TableStylePartName> | null {
  if (r < 0 || c < 0 || r >= table.rows || c >= table.cols) return null;
  const [ar, ac] = owner[r]?.[c] ?? [r, c];
  if (m === 'position') return grid[r]?.[c] ?? null;
  if (m === 'anchor') return grid[ar]?.[ac] ?? null;
  const endsOnly = m === 'endsPooled' || m === 'endsPooledByKind';
  const pooled = new Set<TableStylePartName>(endsOnly ? (grid[ar]?.[ac] ?? []) : []);
  owner.forEach((line, y) =>
    line.forEach(([oy, ox], x) => {
      if (oy !== ar || ox !== ac) return;
      for (const part of grid[y]?.[x] ?? [])
        if (!endsOnly || !part.startsWith('band')) pooled.add(part);
    }),
  );
  return pooled;
}

/** A merged cell's paints and segments: its anchor's parts, each position's own, or all of them pooled. */
export function mergeRows(
  slides: readonly Observed[],
  double: Compound,
  which: 'paint' | 'edge',
): Row<Merged>[] {
  const rows: Row<Merged>[] = [];
  for (const o of slides.filter((x) => x.deck.group === 'merge')) {
    const style = styleById(o.table.style ?? '');
    const palette = paletteOf(o.deck.theme);
    const grid = partsGrid('spec', o.table);
    const owner = anchorsOf(o.table);
    const plain = (o.table.cells ? Object.values(o.table.cells) : []).every(
      (s) => s.tcPr === undefined,
    );
    const partsFor = (r: number, c: number, m: Merged): ReadonlySet<TableStylePartName> | null =>
      mergedParts(o.table, grid, owner, r, c, m);
    for (let r = 0; r < o.table.rows && which === 'paint'; r++) {
      for (let c = 0; c < o.table.cols; c++) {
        rows.push({
          key: cellKey(o.deck.id, o.slide, r, c, 'paint'),
          observed: paintAt(o, r, c),
          predict: (m) => {
            const parts = partsFor(r, c, m) ?? new Set<TableStylePartName>();
            return underCell(style, palette, parts);
          },
        });
      }
    }
    if (!plain || which === 'paint') continue;
    for (const edge of edgesClear(o.table)) {
      const profile = profileAt(o, edge);
      const [a, b] = [profile[0] ?? '', profile.at(-1) ?? ''];
      const [br, bc] = edge.axis === 'h' ? [edge.r - 1, edge.c] : [edge.r, edge.c - 1];
      const same =
        br >= 0 &&
        bc >= 0 &&
        edge.r < o.table.rows &&
        edge.c < o.table.cols &&
        String(owner[br]?.[bc]) === String(owner[edge.r]?.[edge.c]);
      const observed = describe(profile);
      rows.push({
        key: edgeKey(o.deck.id, o.slide, edge),
        observed,
        predict: (m) =>
          drawnAs(
            same
              ? []
              : edgeLinesBetween(
                  style,
                  partsFor(br, bc, m),
                  partsFor(edge.r, edge.c, m),
                  edge.axis,
                  BASELINE.reading,
                  palette,
                  BY_KIND.has(m) ? 'kind' : 'geometry',
                ),
            a,
            b,
            double,
            observed,
          ),
      });
    }
  }
  return rows;
}

export type Segment = 'whole' | 'along' | 'anchorPosition' | 'position';
export type Beneath =
  'style' | 'afterIfCovered' | 'afterUnclaimed' | 'afterUnclaimedElseOwn' | 'afterLevelElseOwn';
export interface MergedLine {
  readonly segment: Segment;
  readonly beneath: Beneath;
  readonly after: Exclude<Segment, 'anchorPosition'>;
}

/**
 * Every edge of the merge slides that write a line. `segment`: which segments a merged cell before
 * the edge claims. `beneath`: what an unclaimed or unwritten segment draws. `after`: which of the
 * cell after's segments its own line reaches.
 */
export function mergedLineRows(
  slides: readonly Observed[],
  double: Compound,
  pooling: Merged,
): Row<MergedLine>[] {
  const rows: Row<MergedLine>[] = [];
  for (const o of slides.filter((x) => x.deck.group === 'merge')) {
    const cells = o.table.cells ?? {};
    if (Object.values(cells).every((s) => s.tcPr === undefined)) continue;
    const style = styleById(o.table.style ?? '');
    const palette = paletteOf(o.deck.theme);
    const grid = partsGrid('spec', o.table);
    const owner = anchorsOf(o.table);
    const inside = (r: number, c: number): boolean =>
      r >= 0 && c >= 0 && r < o.table.rows && c < o.table.cols;
    const tcPr = (r: number, c: number): string => cells[`${String(r)},${String(c)}`]?.tcPr ?? '';
    const isCovered = (r: number, c: number): boolean =>
      /Merge=/.test(cells[`${String(r)},${String(c)}`]?.attrs ?? '');
    for (const edge of edgesClear(o.table)) {
      const profile = profileAt(o, edge);
      const [a, b] = [profile[0] ?? '', profile.at(-1) ?? ''];
      const [br, bc] = edge.axis === 'h' ? [edge.r - 1, edge.c] : [edge.r, edge.c - 1];
      const [tagA, tagB] = edge.axis === 'h' ? ['lnB', 'lnT'] : ['lnR', 'lnL'];
      const observed = describe(profile);
      const [ar, ac] = owner[br]?.[bc] ?? [br, bc];
      const same =
        inside(br, bc) &&
        inside(edge.r, edge.c) &&
        String([ar, ac]) === String(owner[edge.r]?.[edge.c]);
      const styleLines = same
        ? []
        : edgeLinesBetween(
            style,
            mergedParts(o.table, grid, owner, br, bc, pooling),
            mergedParts(o.table, grid, owner, edge.r, edge.c, pooling),
            edge.axis,
            BASELINE.reading,
            palette,
            BY_KIND.has(pooling) ? 'kind' : 'geometry',
          );
      const drawn = (line: Written | undefined): DrawnLine[] | undefined => {
        if (line === undefined) return undefined;
        const done = completed(line, styleLines[0], 'replace');
        return done === null ? [] : [done];
      };
      // Whether the cell at position (r, c) claims this segment under `segment`, and its line there.
      const claimOf = (
        segment: Segment,
        r: number,
        c: number,
        tag: string,
      ): {
        readonly claims: boolean;
        readonly line: Written | undefined;
        readonly anchor: Written | undefined;
      } => {
        const [or, oc] = owner[r]?.[c] ?? [r, c];
        const anchor = written(tcPr(or, oc), tag);
        let claims = true;
        if (segment === 'along') claims = edge.axis === 'v' ? edge.r === or : edge.c === oc;
        else if (segment === 'anchorPosition') claims = r === or && c === oc;
        const line = segment === 'position' ? written(tcPr(r, c), tag) : anchor;
        return { claims, line: claims ? line : undefined, anchor };
      };
      rows.push({
        key: edgeKey(o.deck.id, o.slide, edge),
        observed,
        predict: ({ segment, beneath, after }) => {
          if (same) return drawnAs([], a, b, double, observed);
          const later = inside(edge.r, edge.c)
            ? claimOf(after, edge.r, edge.c, tagB ?? '').line
            : undefined;
          if (!inside(br, bc)) return drawnAs(drawn(later) ?? styleLines, a, b, double, observed);
          const own = claimOf(segment, br, bc, tagA ?? '');
          const outer = !inside(edge.r, edge.c);
          let lines: DrawnLine[];
          if (beneath === 'afterLevelElseOwn' && !own.claims) {
            const level = outer ? undefined : claimOf('along', edge.r, edge.c, tagB ?? '');
            lines = (level?.claims === true ? drawn(level.line) : drawn(own.anchor)) ?? styleLines;
          } else if (beneath === 'afterUnclaimedElseOwn' && !own.claims)
            lines = (outer ? drawn(own.anchor) : drawn(later)) ?? styleLines;
          else if (beneath === 'afterUnclaimed' && !own.claims) lines = drawn(later) ?? styleLines;
          else if (beneath === 'afterIfCovered' && isCovered(br, bc))
            lines = drawn(own.line) ?? drawn(later) ?? styleLines;
          else lines = drawn(own.line) ?? styleLines;
          return drawnAs(lines, a, b, double, observed);
        },
      });
    }
  }
  return rows;
}

/** Each position's anchor, from the spans written: C7's occupancy (ADR 0056) for these tables. */
function anchorsOf(table: TableSpec): [number, number][][] {
  const out: [number, number][][] = Array.from({ length: table.rows }, (_, r) =>
    Array.from({ length: table.cols }, (_, c): [number, number] => [r, c]),
  );
  for (const [at, spec] of Object.entries(table.cells ?? {})) {
    const [r, c] = at.split(',').map(Number) as [number, number];
    if (/Merge=/.test(spec.attrs ?? '')) continue;
    const rows = Number(/rowSpan="(\d+)"/.exec(spec.attrs ?? '')?.[1] ?? 1);
    const cols = Number(/gridSpan="(\d+)"/.exec(spec.attrs ?? '')?.[1] ?? 1);
    for (let y = r; y < r + rows; y++) for (let x = c; x < c + cols; x++) out[y]![x] = [r, c];
  }
  return out;
}

export type Mirror = 'logical' | 'visual';

/** An rtl table: drawn as its logical table mirrored, or with the parts laid over visual columns. */
export function rtlRows(slides: readonly Observed[], double: Compound): Row<Mirror>[] {
  const rows: Row<Mirror>[] = [];
  for (const o of slides.filter((x) => x.deck.group === 'rtl' && x.table.cells === undefined)) {
    const style = styleById(o.table.style ?? '');
    const palette = paletteOf(o.deck.theme);
    const grid = partsGrid('spec', o.table);
    const cols = o.table.cols;
    for (let r = 0; r < o.table.rows; r++) {
      for (let v = 0; v < cols; v++) {
        rows.push({
          key: cellKey(o.deck.id, o.slide, r, v, 'paint'),
          observed: paintAt(o, r, v),
          predict: (m) => {
            const c = m === 'logical' ? cols - 1 - v : v;
            const parts = grid[r]?.[c] ?? new Set<TableStylePartName>();
            return underCell(style, palette, parts);
          },
        });
      }
    }
    for (const edge of edgesClear(o.table)) {
      const profile = profileAt(o, edge);
      const [a, b] = [profile[0] ?? '', profile.at(-1) ?? ''];
      const observed = describe(profile);
      rows.push({
        key: edgeKey(o.deck.id, o.slide, edge),
        observed,
        predict: (m) => {
          // Mirrored, visual column c is logical column cols - 1 - c, and vertical line c is cols - c.
          let logical: Edge = edge;
          if (m === 'logical') {
            logical =
              edge.axis === 'h' ? { ...edge, c: cols - 1 - edge.c } : { ...edge, c: cols - edge.c };
          }
          return drawnAs(
            edgeLines(style, o.table, logical, BASELINE.reading, 'perEdge', palette),
            a,
            b,
            double,
            observed,
          );
        },
      });
    }
  }
  return rows;
}

/**
 * An rtl table's direct lines: owned and placed on logical columns and then mirrored, a cell's
 * `a:lnL` drawn on its visual right; or owned on visual columns, its `a:lnL` on its visual left.
 */
export function rtlDirectRows(slides: readonly Observed[], double: Compound): Row<Mirror>[] {
  const rows: Row<Mirror>[] = [];
  for (const o of slides.filter((x) => x.deck.group === 'rtl' && x.table.cells !== undefined)) {
    const style = styleById(o.table.style ?? '');
    const palette = paletteOf(o.deck.theme);
    const { rows: height, cols } = o.table;
    const tcPr = (r: number, c: number): string =>
      o.table.cells?.[`${String(r)},${String(c)}`]?.tcPr ?? '';
    const inside = (r: number, c: number): boolean => r >= 0 && c >= 0 && r < height && c < cols;
    for (const edge of edgesClear(o.table)) {
      const profile = profileAt(o, edge);
      const [a, b] = [profile[0] ?? '', profile.at(-1) ?? ''];
      const observed = describe(profile);
      const logical: Edge =
        edge.axis === 'h' ? { ...edge, c: cols - 1 - edge.c } : { ...edge, c: cols - edge.c };
      const styleLines = edgeLines(style, o.table, logical, BASELINE.reading, 'perEdge', palette);
      const [tagBefore, tagAfter] = edge.axis === 'h' ? ['lnB', 'lnT'] : ['lnR', 'lnL'];
      rows.push({
        key: edgeKey(o.deck.id, o.slide, edge),
        observed,
        predict: (m) => {
          // The edge in the columns the reading decides on, and the logical column of each.
          const e = m === 'logical' ? logical : edge;
          const column = (c: number): number => (m === 'logical' ? c : cols - 1 - c);
          const [br, bc] = e.axis === 'h' ? [e.r - 1, e.c] : [e.r, e.c - 1];
          const direct = inside(br, bc)
            ? written(tcPr(br, column(bc)), tagBefore ?? '')
            : written(tcPr(e.r, column(e.c)), tagAfter ?? '');
          if (direct === undefined) return drawnAs(styleLines, a, b, double, observed);
          const done = completed(direct, styleLines[0], 'replace');
          return drawnAs(done === null ? [] : [done], a, b, double, observed);
        },
      });
    }
  }
  return rows;
}

export type TablePaint = 'underCells' | 'ignored' | 'overCells';

/** `a:tblPr`'s own fill: under the cells in place of `tblBg`, ignored, or over the cells. */
export function tblPrRows(slides: readonly Observed[]): Row<TablePaint>[] {
  const rows: Row<TablePaint>[] = [];
  for (const o of slides.filter((x) => x.deck.group === 'tblpr')) {
    const palette = paletteOf(o.deck.theme);
    const style = o.table.style === null ? null : styleById(o.table.style);
    const grid = partsGrid('spec', o.table);
    const own = /<a:srgbClr val="([0-9A-F]{6})"/.exec(o.table.tblPrFill ?? '')?.[1];
    const tableFill: FillReport =
      own === undefined
        ? { visible: false, rgb: '', transparency: 1 }
        : { visible: true, rgb: own, transparency: 0 };
    for (let r = 0; r < o.table.rows; r++) {
      for (let c = 0; c < o.table.cols; c++) {
        const cellFill =
          style === null
            ? null
            : (fillWinner(style, grid[r]?.[c] ?? new Set(), BASELINE.reading)?.fill ?? null);
        const background = style?.background?.fill;
        rows.push({
          key: cellKey(o.deck.id, o.slide, r, c, 'paint'),
          observed: paintAt(o, r, c),
          predict: (t) => {
            let colour = rgbOf(PAGE);
            if (t === 'underCells') colour = over(colour, tableFill);
            else if (background !== undefined)
              colour = over(colour, predictFill(background, palette));
            if (t === 'ignored' && background === undefined) colour = rgbOf(PAGE);
            colour = over(colour, predictFill(cellFill, palette));
            if (t === 'overCells') colour = over(colour, tableFill);
            return hexRgb(colour);
          },
        });
      }
    }
  }
  return rows;
}

/** Theme C: pixel samples across the table and beyond its edges, recorded rather than predicted. */
export function backgroundSamples(slides: readonly Observed[]): Record<string, unknown>[] {
  return slides
    .filter((o) => o.deck.group === 'background')
    .map((o) => ({
      deck: o.deck.id,
      slide: o.slide,
      style: o.table.style,
      mask: maskOf(o.table.flags),
      tblPr: o.table.tblPrFill ?? o.table.tblPrEffects ?? null,
      extScale: o.table.extScale ?? 1,
      com: o.com.background,
      ...o.pixels.samples,
    }));
}

/* -------------------------------------------------------------------------- */
/* dashes, diagonals, the background's effect, and what a save keeps          */
/* -------------------------------------------------------------------------- */

/** Clear of any line across an edge's ends: more than half the widest line a probe draws. */
const ALONG_CLEAR = 12;

/** Whether a cell either side of edge `edge` writes a dashed line facing it. */
function dashedAt(table: TableSpec, edge: Edge): boolean {
  const tcPr = (r: number, c: number): string =>
    table.cells?.[`${String(r)},${String(c)}`]?.tcPr ?? '';
  const [before, after] =
    edge.axis === 'h'
      ? [written(tcPr(edge.r - 1, edge.c), 'lnB'), written(tcPr(edge.r, edge.c), 'lnT')]
      : [written(tcPr(edge.r, edge.c - 1), 'lnR'), written(tcPr(edge.r, edge.c), 'lnL')];
  const dashed = before?.dash === true || after?.dash === true;
  if (dashed && table.rtl === true) throw new Error('a dashed line on an rtl probe');
  return dashed;
}

/** The pixels on grid edge `edge`'s centre line, from one end to the other. */
export function alongOf(image: PackedBitmap, edge: Edge): string[] {
  const out: string[] = [];
  const [fixed, from, to] =
    edge.axis === 'h'
      ? [ORIGIN.y + edge.r * CELL.h, ORIGIN.x + edge.c * CELL.w, ORIGIN.x + (edge.c + 1) * CELL.w]
      : [ORIGIN.x + edge.c * CELL.w, ORIGIN.y + edge.r * CELL.h, ORIGIN.y + (edge.r + 1) * CELL.h];
  for (let k = at(from) + ALONG_CLEAR; k < at(to) - ALONG_CLEAR; k++) {
    out.push(hexOf(edge.axis === 'h' ? image.rgb(k, at(fixed)) : image.rgb(at(fixed), k)));
  }
  return out;
}

/** A run-length form of a pixel row: `RRGGBB*n`, comma-separated. */
export function runsOf(pixels: readonly string[]): string {
  const runs: string[] = [];
  for (let k = 0; k < pixels.length;) {
    let n = 1;
    while (k + n < pixels.length && pixels[k + n] === pixels[k]) n++;
    runs.push(`${pixels[k] ?? ''}*${String(n)}`);
    k += n;
  }
  return runs.join(',');
}

export type Dash = 'dashed' | 'solid' | 'nothing';

/** What an edge's centre line shows: the line and the paints in turn, the line alone, or no line. */
export function dashClass(
  pixels: readonly string[],
  line: string | null,
  paints: readonly string[],
): string {
  const lit = line === null ? 0 : pixels.filter((p) => p === line).length;
  const bare = pixels.filter((p) => paints.includes(p)).length;
  if (lit + bare < pixels.length * 0.8) return 'mixed';
  if (lit === 0) return 'nothing';
  return bare === 0 ? 'solid' : 'dashed';
}

/**
 * Edges a written dashed line faces, under the owner rule and whole replacement C9 measured: is the
 * dash drawn, dropped for a solid line, or the line not drawn at all.
 */
export function dashRows(slides: readonly Observed[]): (Row<Dash> & { runs: string })[] {
  const rows: (Row<Dash> & { runs: string })[] = [];
  for (const o of slides) {
    for (const [key, pixels] of Object.entries(o.pixels.along)) {
      const axis = key[0] === 'h' ? 'h' : 'v';
      const [r, c] = key.slice(1).split(',').map(Number) as [number, number];
      const edge: Edge = { axis, r, c };
      const palette = paletteOf(o.deck.theme);
      const style = o.table.style === null ? null : styleById(o.table.style);
      const grid = partsGrid('spec', o.table);
      const tcPr = (rr: number, cc: number): string =>
        o.table.cells?.[`${String(rr)},${String(cc)}`]?.tcPr ?? '';
      const [before, after] =
        axis === 'h'
          ? [
              [r - 1, c],
              [r, c],
            ]
          : [
              [r, c - 1],
              [r, c],
            ];
      const inside = (rc: readonly number[]): boolean =>
        (rc[0] ?? -1) >= 0 &&
        (rc[1] ?? -1) >= 0 &&
        (rc[0] ?? 0) < o.table.rows &&
        (rc[1] ?? 0) < o.table.cols;
      const styleLines =
        style === null
          ? [{ rgb: '000000', weight: 1, cmpd: 'sng' }]
          : edgeLinesBetween(
              style,
              inside(before) ? (grid[before[0] ?? 0]?.[before[1] ?? 0] ?? null) : null,
              inside(after) ? (grid[after[0] ?? 0]?.[after[1] ?? 0] ?? null) : null,
              axis,
              BASELINE.reading,
              palette,
            );
      const owner = inside(before)
        ? written(tcPr(before[0] ?? 0, before[1] ?? 0), axis === 'h' ? 'lnB' : 'lnR')
        : written(tcPr(after[0] ?? 0, after[1] ?? 0), axis === 'h' ? 'lnT' : 'lnL');
      // Only the owner's line draws (C9): a dash its neighbour writes is not this question's.
      if (owner?.dash !== true) continue;
      const line = completed(owner, styleLines[0], 'replace');
      const profile = profileAt(o, edge);
      const paints = [profile[0] ?? '', profile.at(-1) ?? ''];
      rows.push({
        key: edgeKey(o.deck.id, o.slide, edge),
        observed: dashClass(pixels, line?.rgb ?? null, paints),
        runs: runsOf(pixels),
        predict: (d) => (line === null || d === 'nothing' ? 'nothing' : d),
      });
    }
  }
  return rows;
}

/** A diagonal a written line could draw: its rectangle in pixels, its direction, half its width. */
interface Stroke {
  readonly x0: number;
  readonly y0: number;
  readonly x1: number;
  readonly y1: number;
  /** From the rectangle's top left to its bottom right, as drawn. */
  readonly down: boolean;
  readonly half: number;
  readonly rgb: string;
}

const DIAGONAL_TAGS = ['lnTlToBr', 'lnBlToTr'] as const;

export interface DiagonalReading {
  /** The written direction, the other one, or no diagonal at all. */
  readonly drawn: 'asWritten' | 'swapped' | 'none';
  /** Across a merged cell's whole span, or its anchor's position only. */
  readonly extent: 'merged' | 'position';
  /** A covered position's own diagonal. */
  readonly covered: 'ignored' | 'drawn';
  /** In an rtl table: mirrored with the table, or drawn as written in the mirrored position. */
  readonly rtl: 'mirrored' | 'logical';
}

/** The diagonals a table's written lines draw under `reading`, in pixels. */
function strokesOf(table: TableSpec, reading: DiagonalReading | 'all'): Stroke[] {
  const out: Stroke[] = [];
  for (const [key, spec] of Object.entries(table.cells ?? {})) {
    const [r, c] = key.split(',').map(Number) as [number, number];
    const covered = /Merge=/.test(spec.attrs ?? '');
    const spans: [number, number][] = [[1, 1]];
    const rows = Number(/rowSpan="(\d+)"/.exec(spec.attrs ?? '')?.[1] ?? 1);
    const cols = Number(/gridSpan="(\d+)"/.exec(spec.attrs ?? '')?.[1] ?? 1);
    if (!covered && (rows > 1 || cols > 1)) spans.push([rows, cols]);
    for (const tag of DIAGONAL_TAGS) {
      const line = written(spec.tcPr ?? '', tag);
      if (line?.rgb === undefined || line.w === undefined) continue;
      const half = ((line.w / 12700) * SCALE) / 2;
      for (const [h, w] of spans) {
        let downs = [true, false];
        if (reading !== 'all') {
          if (reading.drawn === 'none' || (covered && reading.covered === 'ignored')) continue;
          if ((reading.extent === 'merged') !== (h > 1 || w > 1) && spans.length > 1) continue;
          let down = tag === 'lnTlToBr';
          if (reading.drawn === 'swapped') down = !down;
          if (table.rtl === true && reading.rtl === 'mirrored') down = !down;
          downs = [down];
        }
        const left = table.rtl === true ? table.cols - c - w : c;
        for (const down of downs) {
          out.push({
            x0: at(ORIGIN.x + left * CELL.w),
            y0: at(ORIGIN.y + r * CELL.h),
            x1: at(ORIGIN.x + (left + w) * CELL.w),
            y1: at(ORIGIN.y + (r + h) * CELL.h),
            down,
            half,
            rgb: line.rgb,
          });
        }
      }
    }
  }
  return out;
}

/** A table's grid edges, less those whose profile a written diagonal crosses and so reads instead. */
export function edgesClear(table: TableSpec): Edge[] {
  const strokes = strokesOf(table, 'all');
  return edgesOf(table).filter(
    (edge) =>
      !profilePixels(edge).some(([x, y]) =>
        strokes.some((s) => distanceTo(x, y, s) < s.half + OFF_STROKE),
      ),
  );
}

/** How far a pixel's centre lies from a stroke's centre line. */
function distanceTo(x: number, y: number, s: Stroke): number {
  const [ax, ay, bx, by] = s.down ? [s.x0, s.y0, s.x1, s.y1] : [s.x0, s.y1, s.x1, s.y0];
  const [px, py] = [x + 0.5, y + 0.5];
  const [dx, dy] = [bx - ax, by - ay];
  const t = Math.max(0, Math.min(1, ((px - ax) * dx + (py - ay) * dy) / (dx * dx + dy * dy)));
  return Math.hypot(px - (ax + t * dx), py - (ay + t * dy));
}

/** A pixel this close to a stroke's centre is wholly covered; this much past its edge, untouched. */
const ON_STROKE = 1.5;
const OFF_STROKE = 2.5;
/** Clear of every grid line's pixels. */
const GRID_CLEAR = 8;

/**
 * Points on every diagonal some reading draws, each wholly on one of them or clear of all, away
 * from the grid lines: which reading is right is then a colour per point.
 */
export function diagonalPoints(table: TableSpec): [number, number][] {
  const candidates = strokesOf(table, 'all');
  const xs = Array.from({ length: table.cols + 1 }, (_, k) => at(ORIGIN.x + k * CELL.w));
  const ys = Array.from({ length: table.rows + 1 }, (_, k) => at(ORIGIN.y + k * CELL.h));
  const seen = new Set<string>();
  const out: [number, number][] = [];
  for (const s of candidates) {
    for (let k = 1; k <= 9; k++) {
      const t = k / 10;
      const x = Math.floor(s.x0 + t * (s.x1 - s.x0));
      const y = Math.floor(s.down ? s.y0 + t * (s.y1 - s.y0) : s.y1 - t * (s.y1 - s.y0));
      const key = `${String(x)},${String(y)}`;
      if (seen.has(key)) continue;
      seen.add(key);
      if (
        [...xs.map((g) => Math.abs(x + 0.5 - g)), ...ys.map((g) => Math.abs(y + 0.5 - g))].some(
          (d) => d < GRID_CLEAR,
        )
      )
        continue;
      const distances = candidates.map((c) => ({ d: distanceTo(x, y, c), c }));
      const on = distances.filter(({ d }) => d <= ON_STROKE);
      if (on.length !== 1) continue;
      if (distances.some(({ d, c }) => d > ON_STROKE && d < c.half + OFF_STROKE)) continue;
      out.push([x, y]);
    }
  }
  return out;
}

/** The paint at pixel (x, y): the cell under it, by visual position. */
function paintUnder(o: Observed, x: number, y: number): string {
  const r = Math.floor((y - at(ORIGIN.y)) / at(CELL.h));
  const c = Math.floor((x - at(ORIGIN.x)) / at(CELL.w));
  return paintAt(o, r, c);
}

/** A slide's diagonal points whose pixel is a written diagonal's colour or its cell's paint. */
export function diagonalSamples(o: Observed): { x: number; y: number; seen: string }[] {
  const colours = new Set(strokesOf(o.table, 'all').map((s) => s.rgb));
  return diagonalPoints(o.table)
    .map(([x, y]) => {
      const seen = o.pixels.diagonal[`${String(x)},${String(y)}`];
      if (seen === undefined)
        throw new Error(`${o.deck.id}#${String(o.slide)}: no pixel at ${String(x)},${String(y)}`);
      return { x, y, seen };
    })
    .filter(({ x, y, seen }) => colours.has(seen) || seen === paintUnder(o, x, y));
}

/** Each slide with a written diagonal: the colour every reading puts at every clear point. */
export function diagonalRows(slides: readonly Observed[]): Row<DiagonalReading>[] {
  const rows: Row<DiagonalReading>[] = [];
  for (const o of slides) {
    const samples = diagonalSamples(o);
    if (samples.length === 0) continue;
    rows.push({
      key: `${o.deck.id}#${String(o.slide)}:diagonals`,
      observed: samples.map((p) => p.seen).join(' '),
      predict: (reading) => {
        const strokes = strokesOf(o.table, reading);
        return samples
          .map(
            ({ x, y }) =>
              strokes.findLast((s) => distanceTo(x, y, s) <= ON_STROKE)?.rgb ?? paintUnder(o, x, y),
          )
          .join(' ');
      },
    });
  }
  return rows;
}

/** A hard outer shadow: its offset in points and its colour. */
interface Shadow {
  readonly dx: number;
  readonly dy: number;
  readonly rgb: string;
}

function shadowsIn(xml: string): Shadow[] {
  return [
    ...xml.matchAll(
      /<a:outerShdw\b[^>]*?dist="(\d+)" dir="(\d+)"[^>]*>\s*<a:srgbClr val="([0-9A-F]{6})"/g,
    ),
  ].map((m) => {
    const dist = Number(m[1]) / 12700;
    const angle = (Number(m[2]) / 60000 / 180) * Math.PI;
    return { dx: dist * Math.cos(angle), dy: dist * Math.sin(angle), rgb: m[3] ?? '' };
  });
}

/** The shadows a style's `tblBg` effect gives in theme C: an `a:effectRef` into its effect styles. */
function backgroundShadows(style: TableStyle | null): Shadow[] {
  const effect = style?.background?.effect;
  if (effect === undefined) return [];
  if (effect.kind === 'value')
    throw new Error('a built-in tblBg with its own a:effect: not probed');
  const idx = effect.ref.idx;
  if (idx === 0) return [];
  const xml = THEMES.C.theme.formatScheme?.effects[idx - 1];
  if (xml === undefined) throw new Error(`theme C has no effect style ${String(idx)}`);
  return shadowsIn(xml);
}

export interface BackgroundReading {
  /** The style's `tblBg` effect. */
  readonly effect: 'drawn' | 'ignored';
  /** An `a:effectLst` in `a:tblPr`: in place of the style's, drawn over it, or ignored. */
  readonly tblPr: 'replaces' | 'adds' | 'ignored';
  /** What the background and its effect cover: the grid, or the frame as written. */
  readonly box: 'grid' | 'frame';
}

/** Where theme C's samples beyond the grid were taken, in points (`pixelsOf`). */
function samplePoints(
  table: TableSpec,
): Record<'right' | 'below' | 'beyondGrid', [number, number]> {
  const right = ORIGIN.x + table.cols * CELL.w;
  const y = ORIGIN.y + 2 * CELL.h + 30;
  return {
    right: [right + 3, y],
    below: [ORIGIN.x + 36, ORIGIN.y + table.rows * CELL.h + 3],
    beyondGrid: [right + 20, y],
  };
}

const SHADOW_COLOURS = new Set(['C9C0E1', 'C9C0E3', 'C9C003', 'C9C004']);

/** A sample as the question sees it: the page, a shadow's colour, or painted by the table. */
const sampleClass = (hex: string): string =>
  hex === PAGE ? 'page' : SHADOW_COLOURS.has(hex) ? hex : 'painted';

function sampleAt(o: Observed, name: string): string {
  const sample = o.pixels.samples?.[name];
  if (typeof sample !== 'string')
    throw new Error(`${o.deck.id}#${String(o.slide)} has no ${name} sample`);
  return sample;
}

/** Theme C's samples beyond the grid: which shadows fall there, and whether the background does. */
export function backgroundRows(slides: readonly Observed[]): Row<BackgroundReading>[] {
  const rows: Row<BackgroundReading>[] = [];
  for (const o of slides.filter((x) => x.deck.group === 'background')) {
    const style = o.table.style === null ? null : styleById(o.table.style);
    const own = o.table.tblPrEffects;
    const painted = o.table.tblPrFill !== undefined || style?.background?.fill !== undefined;
    for (const [name, [x, y]] of Object.entries(samplePoints(o.table))) {
      rows.push({
        key: `${o.deck.id}#${String(o.slide)}:${name}`,
        observed: sampleClass(sampleAt(o, name)),
        predict: (reading) => {
          const scale = reading.box === 'frame' ? (o.table.extScale ?? 1) : 1;
          const [w, h] = [o.table.cols * CELL.w * scale, o.table.rows * CELL.h * scale];
          const within = (dx: number, dy: number): boolean =>
            x >= ORIGIN.x + dx &&
            x < ORIGIN.x + dx + w &&
            y >= ORIGIN.y + dy &&
            y < ORIGIN.y + dy + h;
          if (within(0, 0)) return painted ? 'painted' : 'page';
          const shadows = [
            ...(reading.effect === 'drawn' && !(own !== undefined && reading.tblPr === 'replaces')
              ? backgroundShadows(style)
              : []),
            ...(own !== undefined && reading.tblPr !== 'ignored' ? shadowsIn(own) : []),
          ];
          return shadows.findLast((s) => within(s.dx, s.dy))?.rgb ?? 'page';
        },
      });
    }
  }
  return rows;
}

/** An `a:tcPr`'s lines and fill, as `tag:RRGGBB` in document order; empty for none. */
export function tcPrSummary(tcPr: string): string {
  return [
    ...tcPr.matchAll(
      /<a:(ln[LRTB]|lnTlToBr|lnBlToTr|solidFill)\b.*?<a:srgbClr val="([0-9A-F]{6})"/g,
    ),
  ]
    .map((m) => `${m[1] ?? ''}:${m[2] ?? ''}`)
    .join(' ');
}

/** Each covered cell that writes an `a:tcPr`, and what PowerPoint wrote there saving the deck. */
export function resaveRows(dir: string, slides: readonly Observed[]): Row<'kept' | 'emptied'>[] {
  const rows: Row<'kept' | 'emptied'>[] = [];
  for (const deck of new Set(slides.filter((o) => o.deck.resave === true).map((o) => o.deck))) {
    const path = join(dir, 'resaved', `${deck.id}.pptx`);
    if (!existsSync(path)) throw new Error(`no resaved/${deck.id}.pptx: run read.ps1 to the end`);
    const zip = readZip(readFileSync(path));
    deck.slides.forEach((table, k) => {
      const xml = new TextDecoder().decode(
        entry(zip, `ppt/slides/slide${String(k + 1)}.xml`) ?? new Uint8Array(),
      );
      const saved = [...xml.matchAll(/<a:tr\b.*?<\/a:tr>/gs)].map((tr) =>
        [...tr[0].matchAll(/<a:tc\b[^>]*?(?:\/>|>.*?<\/a:tc>)/gs)].map((tc) => tc[0]),
      );
      for (const [pos, spec] of Object.entries(table.cells ?? {})) {
        if (!/Merge=/.test(spec.attrs ?? '') || spec.tcPr === undefined) continue;
        const [r, c] = pos.split(',').map(Number) as [number, number];
        const tc = saved[r]?.[c];
        if (tc === undefined)
          throw new Error(`${deck.id}#${String(k)}: PowerPoint wrote no a:tc at ${pos}`);
        const tcPr = /<a:tcPr\b[^>]*\/>|<a:tcPr\b[^>]*>.*?<\/a:tcPr>/s.exec(tc)?.[0];
        if (tcPr === undefined) throw new Error(`${deck.id}#${String(k)}: no a:tcPr at ${pos}`);
        const summary = tcPrSummary(tcPr);
        rows.push({
          key: `${deck.id}#${String(k)}@${pos}:resaved`,
          observed: summary,
          predict: (how) => (how === 'kept' ? tcPrSummary(spec.tcPr ?? '') : ''),
        });
      }
    });
  }
  return rows;
}

/* -------------------------------------------------------------------------- */
/* the fixture                                                                */
/* -------------------------------------------------------------------------- */

const CORPUS_FILE_CAP = 512 * 1024;
const DIGITS = '0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ';

/**
 * A deck's observations as one code per unit, `width` digits each, every slide's in turn, deflated:
 * per cell its COM fill, COM text, COM sides (`full` decks only) and paint; then per grid edge, in
 * `edgesOf` order, `a|b|` and its descriptor. A code indexes `values`.
 */
function encodeDeck(deck: DeckSpec, observed: readonly Observed[]): Record<string, unknown> {
  const values: string[] = [];
  const index = new Map<string, number>();
  const code = (value: string): number => {
    let c = index.get(value);
    if (c === undefined) {
      c = values.length;
      index.set(value, c);
      values.push(value);
    }
    return c;
  };
  const raw = observed.map((o) => {
    const codes: number[] = [];
    for (let r = 0; r < o.table.rows; r++) {
      for (let c = 0; c < o.table.cols; c++) {
        const cell = o.com.cells[r]?.[c];
        if (cell === undefined) throw new Error('unreachable: guards checked every cell');
        const v = o.table.rtl === true ? o.table.cols - 1 - c : c;
        codes.push(code(cell.fill), code(cell.text));
        if (deck.read === 'full') codes.push(code(Object.values(cell.sides ?? {}).join(';')));
        codes.push(code(deck.group === 'background' ? '-' : paintAt(o, r, v)));
      }
    }
    for (const edge of edgesOf(o.table)) {
      const profile = profileAt(o, edge);
      codes.push(code(`${profile[0] ?? ''}|${profile.at(-1) ?? ''}|${describe(profile)}`));
    }
    const t = o.table;
    return {
      ...(t.name === undefined ? {} : { name: t.name }),
      rows: t.rows,
      cols: t.cols,
      mask: maskOf(t.flags),
      style: t.style,
      ...(t.rtl === true ? { rtl: true } : {}),
      ...(t.tblPrFill === undefined ? {} : { tblPrFill: t.tblPrFill }),
      ...(t.tblPrEffects === undefined ? {} : { tblPrEffects: t.tblPrEffects }),
      ...(t.extScale === undefined ? {} : { extScale: t.extScale }),
      ...(t.ph === undefined ? {} : { ph: t.ph }),
      ...(t.cells === undefined ? {} : { cells: t.cells }),
      ...(t.short === undefined ? {} : { short: t.short }),
      ...(o.com.background === null ? {} : { background: o.com.background }),
      codes,
    };
  });
  const width = values.length <= DIGITS.length ? 1 : 2;
  if (values.length > DIGITS.length ** 2)
    throw new Error(`${deck.id} has too many distinct observations`);
  const digit = (n: number): string =>
    width === 1
      ? (DIGITS[n] ?? '')
      : (DIGITS[Math.floor(n / DIGITS.length)] ?? '') + (DIGITS[n % DIGITS.length] ?? '');
  // A style every slide shares is the deck's; anything else a slide says is keyed by its index.
  const styles = new Set(raw.map((slide) => slide.style));
  const shared = styles.size === 1 ? [...styles][0] : undefined;
  const extras: Record<string, unknown> = {};
  raw.forEach(({ rows: _rows, cols: _cols, mask: _mask, style, codes: _codes, ...rest }, k) => {
    const said = { ...(shared === undefined && style !== null ? { style } : {}), ...rest };
    if (Object.keys(said).length > 0) extras[String(k)] = said;
  });
  const codes = raw.map((slide) => slide.codes.map(digit).join('')).join('');
  return {
    id: deck.id,
    group: deck.group,
    theme: deck.theme,
    read: deck.read,
    ...(deck.defaultTextStyle === undefined ? {} : { defaultTextStyle: deck.defaultTextStyle }),
    ...(deck.txStyles === undefined ? {} : { txStyles: deck.txStyles }),
    ...(deck.masterShapes === undefined ? {} : { masterShapes: deck.masterShapes }),
    ...(deck.layoutShapes === undefined ? {} : { layoutShapes: deck.layoutShapes }),
    ...(shared === undefined || shared === null ? {} : { style: shared }),
    width,
    values,
    shapes: raw.map((s) => `${String(s.rows)}x${String(s.cols)}/${String(s.mask)}`).join(' '),
    ...(Object.keys(extras).length === 0 ? {} : { extras }),
    codes: deflateRawSync(Buffer.from(codes, 'latin1'), { level: 9 }).toString('base64'),
  };
}

/** What PowerPoint wrote when it set a border, a fill or a flag itself: per `a:tc`, its lines and fill. */
function authored(dir: string): Record<string, unknown>[] {
  const logPath = join(dir, 'author-cascade-log.json');
  if (!existsSync(logPath)) throw new Error('no author-cascade-log.json: run author.ps1');
  const log = JSON.parse(readFileSync(logPath, 'utf8')) as {
    operations: { name: string; file: string }[];
  };
  return log.operations.map((op) => {
    const slide = new TextDecoder().decode(
      entry(readZip(readFileSync(join(dir, op.file))), 'ppt/slides/slide1.xml') ?? new Uint8Array(),
    );
    const tbl = /<a:tbl>.*<\/a:tbl>/s.exec(slide)?.[0];
    if (tbl === undefined) throw new Error(`${op.file} has no a:tbl`);
    const tblPr = /<a:tblPr[^>]*\/?>/.exec(tbl)?.[0] ?? '';
    const rows = [...tbl.matchAll(/<a:tr\b.*?<\/a:tr>/gs)].map((tr) =>
      [...tr[0].matchAll(/<a:tc\b([^>]*)>.*?<\/a:tc>/gs)].map((tc) => {
        const tcPr = /<a:tcPr\b[^>]*\/>|<a:tcPr\b[^>]*>.*?<\/a:tcPr>/s.exec(tc[0])?.[0] ?? '';
        const lines = [
          ...tcPr.matchAll(/<a:(ln[LRTB]|lnTlToBr|lnBlToTr)\b([^>]*?)(\/>|>(.*?)<\/a:\1>)/gs),
        ].map(
          (m) =>
            `${m[1] ?? ''}${/<a:noFill\/>/.test(m[4] ?? '') ? ':none' : `:${/w="(\d+)"/.exec(m[2] ?? '')?.[1] ?? '-'}/${/val="([0-9A-F]{6})"/.exec(m[4] ?? '')?.[1] ?? '-'}`}`,
        );
        const fill = /<a:tcPr[^>]*>\s*<a:solidFill><a:srgbClr val="([0-9A-F]{6})"/.exec(tcPr)?.[1];
        return {
          attrs: (tc[1] ?? '').trim(),
          lines,
          ...(fill === undefined ? {} : { fill }),
          tcPr,
        };
      }),
    );
    if (/border|merged|rtl/.test(op.name) && rows.flat().every((cell) => cell.lines.length === 0))
      throw new Error(`${op.file}: PowerPoint set a border and no a:tc carries a line`);
    return { name: op.name, tblPr, rows };
  });
}

function writeFixture(dir: string): void {
  const session = readings.get('control')?.session;
  if (session === undefined) throw new Error('no session');
  const byDeck = new Map<string, Observed[]>();
  for (const o of slides) byDeck.set(o.deck.id, [...(byDeck.get(o.deck.id) ?? []), o]);
  const encoded = (filter: (d: DeckSpec) => boolean): Record<string, unknown>[] =>
    [...byDeck.values()]
      .map((list) => list[0]?.deck)
      .filter((d): d is DeckSpec => d !== undefined && filter(d))
      .map((d) => encodeDeck(d, byDeck.get(d.id) ?? []));
  const sweep = encoded((d) => d.group === 'sweep');
  const ends = encoded((d) => d.group === 'ends');
  const sides = encoded((d) => d.group === 'sides');
  const third = Math.ceil(sweep.length / 3);
  // What a slide says beyond its shape, written once and named by the decks that share it.
  const templates: Record<string, unknown> = {};
  const named = new Map<string, string>();
  const small = encoded((d) => !['sweep', 'ends', 'sides'].includes(d.group)).map((deck) => {
    const extras = deck['extras'] as Record<string, unknown> | undefined;
    if (extras === undefined) return deck;
    const refs = Object.fromEntries(
      Object.entries(extras).map(([k, said]) => {
        const text = JSON.stringify(said);
        let key = named.get(text);
        if (key === undefined) {
          key = `s${String(named.size)}`;
          named.set(text, key);
          templates[key] = said;
        }
        return [k, key];
      }),
    );
    return { ...deck, extras: refs };
  });
  const measuredOn = {
    application: 'Microsoft PowerPoint',
    version: session.version,
    build: session.build,
    sessions: new Set(
      [...readings.values()].map((r) => `${String(r.session.pid)}@${r.session.started}`),
    ).size,
  };
  const header = (part: string): Record<string, unknown> => ({
    $comment: `Experiment C9 - the table-style cascade: ${part} Probe packages written by tools/ground-truth/model/tables/cascade/build-deck.ts, read by tools/ground-truth/model/tables/cascade/read.ps1 over COM and exported at 4 px per point, scored by tools/ground-truth/model/tables/cascade/analyse.ts.`,
    experiment: 'C9',
    subPhase: '4.3',
    generatedBy: 'tools/ground-truth/model/tables/cascade/analyse.ts',
    measuredOn,
  });
  const encoding = {
    digits: DIGITS,
    shapes:
      'one `rows`x`cols`/`mask` per slide, space-separated; `extras` holds anything else a slide says, by index, or in this file the key of it in `templates`; the style is the deck`s unless a slide names one',
    codes:
      'base64 of raw DEFLATE (RFC 1951) over every slide`s codes in order, `width` digits each, indexing `values`: per cell, row-major, its COM fill, COM text, COM sides in full decks and paint; then each grid edge, horizontal first, as a|b|descriptor',
    scale: SCALE,
    origin: ORIGIN,
    cell_pt: CELL,
    profileReach: PROFILE_REACH,
  };
  const files: [string, unknown][] = [
    [
      'table-cascade.json',
      {
        ...header(
          'which parts of a built-in style reach each cell and grid edge, which wins, how direct formatting, merges, rtl and the table background compose, and where the style text sits in the text cascade.',
        ),
        encoding,
        themes: Object.fromEntries(
          (['A', 'B', 'C'] as const).map((k) => [
            k,
            { xml: themeXml(THEMES[k].theme, 'Ground Truth 1'), clrMap: THEMES[k].clrMap },
          ]),
        ),
        floor: { defaultTextStyle: FLOOR_DEFAULT_TEXT_STYLE, txStyles: FLOOR_TX_STYLES },
        double: Object.fromEntries([...double].map(([w, c]) => [String(w), c])),
        levels: LEVELS,
        ladder: LADDER,
        findings: findings.map((f) => ({
          ...f,
          candidates: f.candidates.map((c) => ({
            name: c.name,
            fits: c.fits,
            of: c.of,
            wrong: c.wrong.slice(0, 1),
          })),
        })),
        indistinguishable: (
          inputs.separability as {
            dimension: string;
            pair: string[];
            unit: string;
            probe: string | null;
          }[]
        )
          .filter((v) => v.probe === null)
          .map((v) => `${v.dimension} ${v.pair.join(' / ')} (${v.unit})`),
        background: backgroundSamples(slides),
        dashes: dashRows(slides).map(({ key, runs, observed }) => ({ key, runs, observed })),
        diagonals: slides
          .map((o) => ({ o, samples: diagonalSamples(o) }))
          .filter(({ samples }) => samples.length > 0)
          .map(({ o, samples }) => {
            const colours = [...new Set(samples.map((p) => p.seen))];
            return {
              key: `${o.deck.id}#${String(o.slide)}`,
              colours,
              points: samples
                .map(
                  ({ x, y, seen }) => `${String(x)},${String(y)},${String(colours.indexOf(seen))}`,
                )
                .join(' '),
            };
          }),
        resaved: resaveRows(dir, slides).map((row) => ({
          key: row.key,
          written: row.predict('kept'),
          saved: row.observed,
        })),
        authored: authored(dir),
        templates,
        decks: small,
      },
    ],
    [
      'table-cascade-ends.json',
      {
        ...header('one, two and three rows or columns, every end flag, where the ends meet.'),
        encoding,
        decks: ends,
      },
    ],
    [
      'table-cascade-sides.json',
      {
        ...header('the second theme, eight flag sets, every side COM reports.'),
        encoding,
        decks: sides,
      },
    ],
    ...['first', 'second', 'last'].map((which, k): [string, unknown] => [
      `table-cascade-sweep-${String(k + 1)}.json`,
      {
        ...header(
          `the 5x5 sweep, every flag set, the ${which} third of the gallery${k === 2 ? ' and the default grid' : ''}.`,
        ),
        encoding,
        decks: sweep.slice(k * third, (k + 1) * third),
      },
    ]),
  ];
  const texts = files.map(([name, value]): [string, string] => [name, fixtureJson(value)]);
  for (const [name, text] of texts) {
    if (text.length > CORPUS_FILE_CAP)
      throw new Error(`${name} is ${String(text.length)} bytes, over the corpus cap`);
  }
  for (const [name, text] of texts) {
    writeFileSync(repoPath('corpus', 'ground-truth', name), text);
    console.log(`wrote corpus/ground-truth/${name}: ${String(text.length)} bytes`);
  }
}

/* -------------------------------------------------------------------------- */
/* main                                                                       */
/* -------------------------------------------------------------------------- */

const dir = process.argv[2];
if (dir === undefined)
  throw new Error('usage: tools/ground-truth/model/tables/cascade/analyse.ts <dir> [--fixture]');
const { inputs, slides, readings } = load(dir);
guards(dir, inputs, slides, readings);
console.log(`guards hold on ${String(slides.length)} slides`);
const double = measureDouble(slides);
console.log(
  `double lines: ${[...double].map(([w, c]) => `${String(w)} pt [${c.join(' ')}]`).join('; ') || 'none drawn'}`,
);
/** Where a merged cell's own lines draw, over the style lines the merged-cell reading gives. */
function mergedLineFindings(pooling: Merged): Finding[] {
  const rows = mergedLineRows(slides, double, pooling);
  const segments = ['whole', 'along', 'anchorPosition', 'position'] as const;
  const beneath = [
    'style',
    'afterIfCovered',
    'afterUnclaimed',
    'afterUnclaimedElseOwn',
    'afterLevelElseOwn',
  ] as const;
  const afters = ['whole', 'along', 'position'] as const;
  const all = segments.flatMap((segment) =>
    beneath.flatMap((b) => afters.map((after): MergedLine => ({ segment, beneath: b, after }))),
  );
  // The reading that fits every edge, found over every combination, then each axis around it.
  const best = all.find((m) => rows.every((row) => row.predict(m) === row.observed)) ?? all[0];
  if (best === undefined) throw new Error('no merged line reading');
  return [
    score(
      'merged cells: which segments its own line claims',
      segments.map((segment): [string, MergedLine] => [segment, { ...best, segment }]),
      rows,
    ),
    score(
      'merged cells: what an unclaimed segment draws',
      beneath.map((b): [string, MergedLine] => [b, { ...best, beneath: b }]),
      rows,
    ),
    score(
      'merged cells: how far the cell after reaches',
      afters.map((after): [string, MergedLine] => [after, { ...best, after }]),
      rows,
    ),
  ];
}

/** Which diagonals draw: over every combination, then each axis around the one that fits. */
function diagonalFindings(): Finding[] {
  const rows = diagonalRows(slides);
  const axes = {
    drawn: ['asWritten', 'swapped', 'none'],
    extent: ['merged', 'position'],
    covered: ['ignored', 'drawn'],
    rtl: ['mirrored', 'logical'],
  } as const;
  const all = axes.drawn.flatMap((drawn) =>
    axes.extent.flatMap((extent) =>
      axes.covered.flatMap((covered) =>
        axes.rtl.map((rtl): DiagonalReading => ({ drawn, extent, covered, rtl })),
      ),
    ),
  );
  const best = all.find((m) => rows.every((row) => row.predict(m) === row.observed)) ?? all[0];
  if (best === undefined) throw new Error('no diagonal reading');
  return [
    score(
      'diagonals: which',
      axes.drawn.map((drawn): [string, DiagonalReading] => [drawn, { ...best, drawn }]),
      rows,
    ),
    score(
      'diagonals: a merged cell’s',
      axes.extent.map((extent): [string, DiagonalReading] => [extent, { ...best, extent }]),
      rows,
    ),
    score(
      'diagonals: a covered cell’s own',
      axes.covered.map((covered): [string, DiagonalReading] => [covered, { ...best, covered }]),
      rows,
    ),
    score(
      'diagonals: in an rtl table',
      axes.rtl.map((rtl): [string, DiagonalReading] => [rtl, { ...best, rtl }]),
      rows,
    ),
  ];
}

/** The background's effect: over every combination, then each axis around the one that fits. */
function backgroundFindings(): Finding[] {
  const rows = backgroundRows(slides);
  const axes = {
    effect: ['drawn', 'ignored'],
    tblPr: ['replaces', 'adds', 'ignored'],
    box: ['grid', 'frame'],
  } as const;
  const all = axes.effect.flatMap((effect) =>
    axes.tblPr.flatMap((tblPr) =>
      axes.box.map((box): BackgroundReading => ({ effect, tblPr, box })),
    ),
  );
  const best = all.find((m) => rows.every((row) => row.predict(m) === row.observed)) ?? all[0];
  if (best === undefined) throw new Error('no background reading');
  return [
    score(
      'background: tblBg’s effect',
      axes.effect.map((effect): [string, BackgroundReading] => [effect, { ...best, effect }]),
      rows,
    ),
    score(
      'background: an effect in a:tblPr',
      axes.tblPr.map((tblPr): [string, BackgroundReading] => [tblPr, { ...best, tblPr }]),
      rows,
    ),
    score(
      'background: what it covers',
      axes.box.map((box): [string, BackgroundReading] => [box, { ...best, box }]),
      rows,
    ),
  ];
}

const questions: Finding[] = [
  ...cascadeFindings('fill', comRows(slides, 'fill'), ['applicability', 'order', 'merge']),
  ...cascadeFindings('text', comRows(slides, 'text'), [
    'applicability',
    'order',
    'merge',
    'source',
  ]),
  ...cascadeFindings('edge', edgeRows(slides, double), [
    'applicability',
    'order',
    'merge',
    'edges',
  ]),
  score('paint: translucent fills and tblBg', PAINTS, paintRows(slides)),
  score(
    'default grid: text',
    [
      ['layer', 'layer'],
      ['floor', 'floor'],
      ['endsBold', 'endsBold'],
    ],
    gridRows(slides),
  ),
  score(
    'default grid: lines',
    [
      ['black', 'black'],
      ['tx1', 'tx1'],
      ['none', 'none'],
    ],
    gridEdgeRows(slides),
  ),
  ...(() => {
    const rows = textRows(slides);
    // The reading that fits every rung, found over every combination, then each axis around it.
    const places = POSITIONS.flatMap((position) =>
      TERMINI.flatMap((terminus) =>
        REACHES.flatMap((reach) =>
          FLOORS.map((floor): TextPlace => ({ position, terminus, reach, floor })),
        ),
      ),
    );
    const best =
      places.find((p) => rows.every((row) => row.predict(p) === row.observed)) ?? places[0];
    if (best === undefined) throw new Error('no text place');
    return (['position', 'terminus', 'reach', 'floor'] as const).map((axis) => ({ axis, best }));
  })().map(({ axis, best }) => {
    const rows = textRows(slides);
    const values: readonly string[] = {
      position: POSITIONS,
      terminus: TERMINI,
      reach: REACHES,
      floor: FLOORS,
    }[axis];
    return score(
      `text cascade: ${axis}`,
      values.map((v): [string, TextPlace] => [v, { ...best, [axis]: v }]),
      rows,
    );
  }),
  score(
    'direct: fill',
    [
      ['tcPr', 'tcPr'],
      ['style', 'style'],
    ],
    directFillRows(slides),
  ),
  ...(() => {
    const rows = directRows(slides, double);
    const rules = ['owner', 'ownerThenOther', 'later', 'earlier', 'heavier', 'style'] as const;
    return [
      score(
        'direct: which line an edge draws',
        rules.map((rule): [string, DirectReading] => [rule, { rule, completion: 'replace' }]),
        rows,
      ),
      score(
        'direct: a partial line',
        (['replace', 'merge'] as const).map((completion): [string, DirectReading] => [
          completion,
          { rule: 'owner', completion },
        ]),
        rows,
      ),
    ];
  })(),
  ...(() => {
    const [paint, edge] = (['paint', 'edge'] as const).map((which) =>
      score(
        `merged cells: ${which}`,
        MERGED.map((m): [string, Merged] => [m, m]),
        mergeRows(slides, double, which),
      ),
    );
    if (paint === undefined || edge === undefined) throw new Error('no merged-cell findings');
    const pooling = MERGED.find((m) => edge.winner.includes(m) && BY_KIND.has(m));
    if (pooling === undefined) throw new Error('no merged reading offers edges by kind');
    return [paint, edge, ...mergedLineFindings(pooling)];
  })(),
  score(
    'rtl',
    [
      ['logical', 'logical'],
      ['visual', 'visual'],
    ],
    rtlRows(slides, double),
  ),
  score(
    'rtl: direct lines',
    [
      ['logical', 'logical'],
      ['visual', 'visual'],
    ],
    rtlDirectRows(slides, double),
  ),
  score(
    'tblPr fill',
    (['underCells', 'ignored', 'overCells'] as const).map((t): [string, TablePaint] => [t, t]),
    tblPrRows(slides),
  ),
  score(
    'direct: a dashed line',
    (['dashed', 'solid', 'nothing'] as const).map((d): [string, Dash] => [d, d]),
    dashRows(slides),
  ),
  ...diagonalFindings(),
  ...backgroundFindings(),
  score(
    'merged cells: a covered cell’s own a:tcPr, saved',
    (['kept', 'emptied'] as const).map((how): [string, 'kept' | 'emptied'] => [how, how]),
    resaveRows(dir, slides),
  ),
];
const findings = questions;
for (const f of findings) {
  console.log(
    `${f.question}: ${f.winner.join(' = ')} (${String(f.informative)} of ${String(f.rows)} rows informative)`,
  );
  for (const c of f.candidates)
    console.log(`   ${c.name.padEnd(18)} ${String(c.fits)}/${String(c.of)}`);
}
if (process.argv.includes('--fixture')) writeFixture(dir);
