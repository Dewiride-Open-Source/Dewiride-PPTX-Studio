import { GeometryError } from './errors.js';
import type {
  PresetAdjustHandle,
  PresetCommand,
  PresetConnectionSite,
  PresetGuide,
  PresetPath,
  PresetPathFill,
  PresetPoint,
  PresetShape,
  PresetTextRect,
} from './types.js';

/**
 * The reader for the generated preset buckets.
 *
 * ## Why the buckets are strings and not object literals
 *
 * The source is 539 KB of XML for 187 shapes, and a faithful object literal of
 * it comes out larger than the XML rather than smaller: `{ kind: 'lnTo', to: {
 * x: 'x2', y: 't' } }` spends 34 characters saying what `L x2 t` says in 6.
 * Written that way the six buckets would be most of a megabyte of source for
 * every consumer's bundler to parse, and would dominate the download of a
 * library whose whole job is drawing shapes.
 *
 * So each bucket is one string in the encoding below, and this module reads it.
 * `packages/xml/src/schema-order.gen.ts` made the same trade for the same
 * reason; the difference is that this data is an order of magnitude larger, so
 * the saving is the difference between shipping preset geometry and making it
 * something a consumer has to opt into.
 *
 * ## The encoding
 *
 * ```text
 * data    = shape (";" shape)*
 * shape   = name "|" av "|" gd "|" ah "|" cxn "|" rect "|" paths
 * av, gd  = [ guide ("," guide)* ]      guide  = name " " op (" " operand)*
 * ah      = [ handle ("," handle)* ]    handle = ("X"|"P") ref min max ref min max x y
 * cxn     = [ site ("," site)* ]        site   = ang " " x " " y
 * rect    = [ l " " t " " r " " b ]
 * paths   = [ path ("!" path)* ]        path   = w " " h " " fill " " stroke " " ext
 *                                                ("," command)*
 * command = "M" x y | "L" x y | "Q" x y x y | "C" x y x y x y | "A" wR hR stAng swAng | "Z"
 * ```
 *
 * Tokens within a record are separated by one space, records by a comma, shapes
 * by a semicolon, and a lone `-` reads as "absent, or the schema default".
 * None of those separators occurs in the source data - which the generator
 * checks against the actual file rather than assuming, because an operand there
 * is drawn from letters, digits, `+`, `*`, `/`, `?`, `:` and a leading minus,
 * and a single new character in a future POI release would otherwise corrupt a
 * bucket silently.
 *
 * ## Decoding is lazy, and per shape
 *
 * Splitting a bucket on `;` is cheap and happens once, in the constructor.
 * Decoding a shape happens when that shape is first asked for, and is cached
 * from then on. A slide holding three rounded rectangles decodes one preset,
 * not thirty - which is the point of splitting into six buckets at all: a page
 * pays for what it draws.
 */

/** A lone `-` means absent. Anything else is a real operand. */
function optional(token: string | undefined): string | null {
  return token === undefined || token === '-' ? null : token;
}

function fail(preset: string | null, message: string): never {
  throw new GeometryError('PRESET_DECODE', message, preset);
}

/** Split a record into tokens, refusing a count the grammar does not allow. */
function tokens(preset: string, record: string, expected: number, what: string): string[] {
  const parts = record.split(' ');
  if (parts.length !== expected) {
    fail(
      preset,
      `${what}: expected ${String(expected)} tokens, found ${String(parts.length)} in "${record}"`,
    );
  }
  return parts;
}

/** An empty field is no records at all, which `"".split(",")` would call one empty record. */
function records(field: string): string[] {
  return field === '' ? [] : field.split(',');
}

function point(preset: string, x: string | undefined, y: string | undefined): PresetPoint {
  if (x === undefined || y === undefined) fail(preset, 'a point is missing a coordinate');
  return { x, y };
}

function decodeGuides(preset: string, field: string): PresetGuide[] {
  return records(field).map((record) => {
    const parts = record.split(' ');
    const name = parts[0];
    if (name === undefined || parts.length < 2) fail(preset, `guide "${record}" has no formula`);
    return { name, fmla: parts.slice(1) };
  });
}

function decodeHandles(preset: string, field: string): PresetAdjustHandle[] {
  return records(field).map((record): PresetAdjustHandle => {
    const t = tokens(preset, record, 9, 'adjust handle');
    const pos = point(preset, t[7], t[8]);
    if (t[0] === 'X') {
      return {
        kind: 'xy',
        gdRefX: optional(t[1]),
        minX: optional(t[2]),
        maxX: optional(t[3]),
        gdRefY: optional(t[4]),
        minY: optional(t[5]),
        maxY: optional(t[6]),
        pos,
      };
    }
    if (t[0] === 'P') {
      return {
        kind: 'polar',
        gdRefR: optional(t[1]),
        minR: optional(t[2]),
        maxR: optional(t[3]),
        gdRefAng: optional(t[4]),
        minAng: optional(t[5]),
        maxAng: optional(t[6]),
        pos,
      };
    }
    return fail(preset, `adjust handle kind "${String(t[0])}" is neither X nor P`);
  });
}

function decodeSites(preset: string, field: string): PresetConnectionSite[] {
  return records(field).map((record) => {
    const t = tokens(preset, record, 3, 'connection site');
    return { ang: t[0] ?? '', pos: point(preset, t[1], t[2]) };
  });
}

function decodeRect(preset: string, field: string): PresetTextRect | null {
  if (field === '') return null;
  const t = tokens(preset, field, 4, 'text rectangle');
  return { l: t[0] ?? '', t: t[1] ?? '', r: t[2] ?? '', b: t[3] ?? '' };
}

const FILL_MODES = new Set<string>([
  'none',
  'norm',
  'lighten',
  'lightenLess',
  'darken',
  'darkenLess',
]);

function decodeCommand(preset: string, record: string): PresetCommand {
  const t = record.split(' ');
  switch (t[0]) {
    case 'M':
      return { kind: 'moveTo', to: point(preset, t[1], t[2]) };
    case 'L':
      return { kind: 'lnTo', to: point(preset, t[1], t[2]) };
    case 'Q':
      return { kind: 'quadBezTo', c1: point(preset, t[1], t[2]), to: point(preset, t[3], t[4]) };
    case 'C':
      return {
        kind: 'cubicBezTo',
        c1: point(preset, t[1], t[2]),
        c2: point(preset, t[3], t[4]),
        to: point(preset, t[5], t[6]),
      };
    case 'A': {
      const a = tokens(preset, record, 5, 'arcTo');
      return {
        kind: 'arcTo',
        wR: a[1] ?? '',
        hR: a[2] ?? '',
        stAng: a[3] ?? '',
        swAng: a[4] ?? '',
      };
    }
    case 'Z':
      return { kind: 'close' };
    default:
      return fail(preset, `unknown path command "${String(t[0])}"`);
  }
}

function decodePaths(preset: string, field: string): PresetPath[] {
  if (field === '') return [];
  return field.split('!').map((encoded) => {
    const [head, ...commands] = encoded.split(',');
    const t = tokens(preset, head ?? '', 5, 'path header');
    const fill = t[2] === '-' ? 'norm' : (t[2] ?? '');
    if (!FILL_MODES.has(fill)) fail(preset, `path fill "${fill}" is not an ST_PathFillMode`);
    return {
      w: t[0] === '-' ? 0 : Number(t[0]),
      h: t[1] === '-' ? 0 : Number(t[1]),
      fill: fill as PresetPathFill,
      stroke: t[3] !== '0',
      extrusionOk: t[4] !== '0',
      commands: commands.map((record) => decodeCommand(preset, record)),
    };
  });
}

/** Decode one shape's `|`-separated body. Exported so the generator can round-trip it. */
export function decodeShape(name: string, body: string): PresetShape {
  const f = body.split('|');
  if (f.length !== 6) fail(name, `expected 6 fields, found ${String(f.length)}`);
  return {
    name,
    avLst: decodeGuides(name, f[0] ?? ''),
    gdLst: decodeGuides(name, f[1] ?? ''),
    ahLst: decodeHandles(name, f[2] ?? ''),
    cxnLst: decodeSites(name, f[3] ?? ''),
    rect: decodeRect(name, f[4] ?? ''),
    pathLst: decodePaths(name, f[5] ?? ''),
  };
}

/**
 * One generated bucket, decoded on demand.
 *
 * Holds the encoded text and a cache and nothing else, so importing a bucket
 * you end up not drawing from costs a string and a map of substrings into it.
 */
export class PresetBucket {
  readonly #encoded = new Map<string, string>();
  readonly #decoded = new Map<string, PresetShape>();

  constructor(data: string) {
    for (const entry of data.split(';')) {
      const cut = entry.indexOf('|');
      if (cut < 0) fail(null, `bucket entry "${entry.slice(0, 40)}" has no name separator`);
      this.#encoded.set(entry.slice(0, cut), entry.slice(cut + 1));
    }
  }

  /** Every preset name in this bucket, in the order the source file lists them. */
  get names(): readonly string[] {
    return [...this.#encoded.keys()];
  }

  has(name: string): boolean {
    return this.#encoded.has(name);
  }

  /** `undefined` when the name lives in another bucket - callers chain over the six. */
  get(name: string): PresetShape | undefined {
    const cached = this.#decoded.get(name);
    if (cached !== undefined) return cached;
    const body = this.#encoded.get(name);
    if (body === undefined) return undefined;
    const shape = decodeShape(name, body);
    this.#decoded.set(name, shape);
    return shape;
  }
}
