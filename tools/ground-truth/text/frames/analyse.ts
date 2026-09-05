/**
 * Experiment T6, step 4 - score the candidate models and emit the fixture.
 *
 * ```
 * node tools/ground-truth/text/frames/analyse.ts <work-dir> [--fixture <path>]
 * ```
 *
 * Every question is scored against every candidate reading, including the wrong
 * one a person writes first. A question whose winner is not unique throws.
 */

import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

import { readEmf } from '../../lib/emf.ts';

const arg = process.argv[2];
if (arg === undefined) throw new Error('usage: analyse.ts <work-dir> [--fixture <path>]');
const dir: string = arg;
const fixtureFlag = process.argv.indexOf('--fixture');
const fixturePath = fixtureFlag < 0 ? null : (process.argv[fixtureFlag + 1] ?? null);

/* -------------------------------------------------------------------------- */
/* the readings                                                               */
/* -------------------------------------------------------------------------- */

interface Bound {
  readonly index: number;
  readonly left: number | null;
  readonly top: number | null;
  readonly width: number | null;
  readonly height: number | null;
  readonly text: string | null;
}

interface ShapeReading {
  readonly id: string;
  readonly slide: number;
  readonly left: number | null;
  readonly top: number | null;
  readonly width: number | null;
  readonly height: number | null;
  readonly orientation: number | null;
  readonly vAnchor: number | null;
  readonly hAnchor: number | null;
  readonly marginLeft: number | null;
  readonly marginTop: number | null;
  readonly marginRight: number | null;
  readonly marginBottom: number | null;
  readonly wordWrap: number | null;
  readonly columnNumber: number | null;
  readonly columnSpacing: number | null;
  readonly textLeft: number | null;
  readonly textTop: number | null;
  readonly textWidth: number | null;
  readonly textHeight: number | null;
  readonly paragraphs: readonly Bound[];
  readonly lines: readonly Bound[];
}

interface DeckReading {
  readonly deck: string;
  readonly opened: boolean;
  readonly repaired: boolean | null;
  readonly emf: readonly { readonly slide: number; readonly file: string | null }[];
  readonly shapes: readonly ShapeReading[];
}

interface ProbeInput {
  readonly id: string;
  readonly slide: number;
  readonly family: string;
  readonly rect: { readonly x: number; readonly y: number; readonly w: number; readonly h: number };
  readonly isolate: boolean;
  readonly frame: Record<string, unknown>;
  readonly paragraphs: number;
  readonly lines: readonly number[];
  readonly vars: Readonly<Record<string, string | number | boolean | null>>;
}

const inputs = JSON.parse(readFileSync(join(dir, 'frame-inputs.json'), 'utf8')) as {
  decks: { deck: string; file: string; shapes: ProbeInput[] }[];
};
const readings = JSON.parse(readFileSync(join(dir, 'frame-readings.json'), 'utf8')) as {
  decks: DeckReading[];
};

const probeOf = new Map<string, ProbeInput>();
const deckOfProbe = new Map<string, string>();
for (const deck of inputs.decks) {
  for (const shape of deck.shapes) {
    probeOf.set(shape.id, shape);
    deckOfProbe.set(shape.id, deck.deck);
  }
}

const readOf = new Map<string, ShapeReading>();
const repairedDecks = new Set<string>();
for (const deck of readings.decks) {
  if (deck.repaired === true || !deck.opened) repairedDecks.add(deck.deck);
  for (const shape of deck.shapes) readOf.set(shape.id, shape);
}

/** Text drawn on the slide an isolated probe owns, in draw order. */
function drawn(id: string): { text: string; face: string; verticalFace: boolean }[] {
  const probe = probeOf.get(id);
  const deckName = deckOfProbe.get(id);
  if (probe === undefined || deckName === undefined) return [];
  const deck = readings.decks.find((d) => d.deck === deckName);
  const file = deck?.emf.find((e) => e.slide === probe.slide)?.file;
  if (file === undefined || file === null) return [];
  return readEmf(new Uint8Array(readFileSync(join(dir, file))))
    .texts.filter((t) => t.text.trim().length > 0)
    .map((t) => ({
      text: t.text,
      face: t.font?.face ?? '',
      verticalFace: t.font?.verticalFace ?? false,
    }));
}

/* -------------------------------------------------------------------------- */
/* scoring                                                                    */
/* -------------------------------------------------------------------------- */

/** A quarter point: `BoundLeft` is quantised, and every model here is exact within it. */
const TOL = 0.25;

const near = (a: number, b: number, tol = TOL): boolean => Math.abs(a - b) <= tol;

interface Candidate<Row> {
  readonly name: string;
  readonly predict: (row: Row) => number | string | boolean | null;
}

interface Scored {
  readonly id: string;
  readonly question: string;
  readonly rows: number;
  readonly candidates: { readonly name: string; readonly score: number }[];
  readonly winner: string;
}

const scored: Scored[] = [];

/**
 * Score every candidate over every row and keep the one that fits perfectly.
 *
 * Throws unless exactly one does: two winners means the probes do not separate
 * the readings, and none means the right reading is not in the list.
 */
function choose<Row extends { readonly measured: number | string | boolean | null }>(
  id: string,
  question: string,
  rows: readonly Row[],
  candidates: readonly Candidate<Row>[],
): string {
  if (rows.length === 0) throw new Error(`${id}: no rows to score`);
  const results = candidates.map((candidate) => ({
    name: candidate.name,
    score: rows.filter((row) => {
      const predicted = candidate.predict(row);
      if (typeof predicted === 'number' && typeof row.measured === 'number') {
        return near(predicted, row.measured);
      }
      return predicted === row.measured;
    }).length,
  }));
  const perfect = results.filter((r) => r.score === rows.length);
  if (perfect.length !== 1) {
    const detail = results.map((r) => `${r.name} ${String(r.score)}/${String(rows.length)}`);
    throw new Error(
      `${id}: ${String(perfect.length)} candidate(s) fit all ${String(rows.length)} rows - ${detail.join(', ')}`,
    );
  }
  const winner = perfect[0]?.name ?? '';
  scored.push({ id, question, rows: rows.length, candidates: results, winner });
  return winner;
}

const probesIn = (family: string): ProbeInput[] =>
  [...probeOf.values()].filter(
    (p) => p.family === family && !repairedDecks.has(deckOfProbe.get(p.id) ?? ''),
  );

function must(value: number | null, what: string): number {
  if (value === null || !Number.isFinite(value)) throw new Error(`missing reading: ${what}`);
  return value;
}

/* -------------------------------------------------------------------------- */
/* the line box, re-derived here so no question depends on another            */
/* -------------------------------------------------------------------------- */

const NATURAL = 1.2;
/** Arial's `lastLineCoefficient`, measured in T4. */
const ARIAL_B = 0.227618991;

const advanceOf = (sz: number, pct = 1): number => pct * NATURAL * sz;
const trimmedLast = (sz: number, pct: number): number =>
  Math.max(0.75 * advanceOf(sz, pct) + ARIAL_B * sz, Math.min(advanceOf(sz, pct), NATURAL * sz));

/* -------------------------------------------------------------------------- */
/* Q1 - the block corner and the block height                                 */
/* -------------------------------------------------------------------------- */

interface ControlRow {
  readonly id: string;
  readonly lines: number;
  readonly sz: number;
  readonly face: string;
  readonly dTop: number;
  readonly dLeft: number;
  readonly measured: number;
}

const controlRows: ControlRow[] = probesIn('control').map((p) => {
  const r = readOf.get(p.id);
  if (r === undefined) throw new Error(`no reading for ${p.id}`);
  const lines = p.lines.reduce((a, b) => a + b, 0);
  return {
    id: p.id,
    lines,
    sz: Number(p.vars['sz']) / 100,
    face: String(p.vars['face']),
    dTop: must(r.textTop, p.id) - must(r.top, p.id),
    dLeft: must(r.textLeft, p.id) - must(r.left, p.id),
    measured: must(r.textHeight, p.id),
  };
});

choose(
  'block-corner',
  'where a top-anchored, left-aligned block starts at zero insets',
  controlRows.map((row) => ({ ...row, measured: row.dTop })),
  [
    { name: 'the frame edge', predict: () => 0 },
    { name: 'half the leading below it', predict: (row) => 0.1 * row.sz },
  ],
);

/* -------------------------------------------------------------------------- */
/* Q2 - the insets                                                            */
/* -------------------------------------------------------------------------- */

/** The `a:bodyPr` inset defaults, in points. */
const INSET_DEFAULT = { lIns: 7.2, tIns: 3.6, rIns: 7.2, bIns: 3.6 } as const;

interface InsetRow {
  readonly id: string;
  readonly edge: string;
  readonly stated: number | null;
  readonly measured: number;
}

const insetRows: InsetRow[] = [];
for (const p of probesIn('inset')) {
  const probe = String(p.vars['probe']);
  if (probe === 'wrap') continue;
  const r = readOf.get(p.id);
  if (r === undefined) continue;
  const edge = String(p.vars['edge']);
  const attr = `${edge}Ins` as 'lIns' | 'tIns' | 'rIns' | 'bIns';
  const stated = edge === 'all' ? null : ((p.frame[attr] as number | null | undefined) ?? null);
  const left = must(r.left, p.id);
  const top = must(r.top, p.id);
  const width = must(r.width, p.id);
  const height = must(r.height, p.id);
  const measured =
    edge === 'l' || (edge === 'all' && probe === 'absent')
      ? must(r.textLeft, p.id) - left
      : edge === 't'
        ? must(r.textTop, p.id) - top
        : edge === 'r' || probe === 'absent-right'
          ? left + width - must(r.textLeft, p.id)
          : top + height - must(r.textTop, p.id) - must(r.textHeight, p.id);
  insetRows.push({ id: p.id, edge, stated, measured });
}

/** Right and bottom are read as a distance from the far edge, so the advance is in them. */
const RIGHT_ADVANCE = insetRows.find((row) => row.id === 'ins-r-0')?.measured ?? 0;

choose(
  'inset-applies',
  'does each inset move the block by exactly its own value',
  insetRows
    .filter((row) => row.edge === 'l' || row.edge === 't')
    .map((row) => ({ ...row, measured: row.measured })),
  [
    {
      name: 'the stated value, or 7.2/3.6 when absent',
      predict: (row) => row.stated ?? (row.edge === 'l' ? INSET_DEFAULT.lIns : INSET_DEFAULT.tIns),
    },
    { name: 'the stated value, or zero when absent', predict: (row) => row.stated ?? 0 },
    {
      name: 'a symmetric 7.2 default on every edge',
      predict: (row) => row.stated ?? INSET_DEFAULT.lIns,
    },
    { name: 'ignored', predict: () => 0 },
  ],
);

choose(
  'inset-far-edge',
  'do rIns and bIns move a right-aligned or bottom-anchored block',
  insetRows
    .filter((row) => row.edge === 'r' || row.edge === 'b')
    .map((row) => ({ ...row, measured: row.measured })),
  [
    {
      name: 'the stated value, or 7.2/3.6 when absent',
      predict: (row) =>
        (row.stated ?? (row.edge === 'r' ? INSET_DEFAULT.rIns : INSET_DEFAULT.bIns)) +
        (row.edge === 'r' ? RIGHT_ADVANCE : 0),
    },
    { name: 'ignored', predict: (row) => (row.edge === 'r' ? RIGHT_ADVANCE : 0) },
  ],
);

/** When the two insets exceed the box, where the collapsed content box sits. */
interface CollapseRow {
  readonly id: string;
  readonly near: number;
  readonly far: number;
  readonly size: number;
  readonly blockHalf: number;
  readonly measured: number;
}

const collapseRows: CollapseRow[] = [];
for (const id of ['ins-huge-lr', 'ins-huge-asym-lr', 'ins-huge-tb', 'ins-huge-asym-tb']) {
  const p = probeOf.get(id);
  const r = readOf.get(id);
  if (p === undefined || r === undefined) continue;
  const horizontal = id.includes('-lr');
  const near_ = Number(horizontal ? p.frame['lIns'] : p.frame['tIns']);
  const far = Number(horizontal ? p.frame['rIns'] : p.frame['bIns']);
  const size = horizontal ? p.rect.w : p.rect.h;
  collapseRows.push({
    id,
    near: near_,
    far,
    size,
    blockHalf: horizontal ? 0 : must(r.textHeight, id) / 2,
    measured: horizontal
      ? must(r.textLeft, id) - must(r.left, id)
      : must(r.textTop, id) - must(r.top, id),
  });
}

choose(
  'inset-collapse',
  'where the content box goes when the two insets exceed the frame',
  collapseRows,
  [
    {
      name: 'the midpoint of the two inset edges',
      predict: (row) => (row.near + (row.size - row.far)) / 2 - row.blockHalf,
    },
    { name: 'the middle of the frame', predict: (row) => row.size / 2 - row.blockHalf },
    {
      name: 'split in proportion to the two insets',
      predict: (row) => (row.near / (row.near + row.far)) * row.size - row.blockHalf,
    },
    { name: 'the near inset, unclamped', predict: (row) => row.near - row.blockHalf },
  ],
);

/* -------------------------------------------------------------------------- */
/* Q3 - the anchor                                                            */
/* -------------------------------------------------------------------------- */

interface AnchorRow {
  readonly id: string;
  readonly anchor: string;
  readonly tIns: number;
  readonly bIns: number;
  readonly boxH: number;
  readonly blockH: number;
  readonly measured: number;
}

const anchorRows: AnchorRow[] = [];
for (const family of ['anchor', 'anchor-spacing']) {
  for (const p of probesIn(family)) {
    const r = readOf.get(p.id);
    if (r === undefined) continue;
    anchorRows.push({
      id: p.id,
      anchor: String(p.vars['anchor']),
      tIns: Number(p.frame['tIns'] ?? 0),
      bIns: Number(p.frame['bIns'] ?? 0),
      boxH: p.rect.h,
      blockH: must(r.textHeight, p.id),
      measured: must(r.textTop, p.id) - must(r.top, p.id),
    });
  }
}

const anchorFraction = (anchor: string): number =>
  anchor === 't' ? 0 : anchor === 'ctr' ? 0.5 : 1;

choose('anchor', 'where each ST_TextAnchoringType puts the block', anchorRows, [
  {
    name: 'inside the inset box, with just and dist as bottom',
    predict: (row) =>
      row.tIns + anchorFraction(row.anchor) * (row.boxH - row.tIns - row.bIns - row.blockH),
  },
  {
    name: 'inside the frame, ignoring the insets',
    predict: (row) => anchorFraction(row.anchor) * (row.boxH - row.blockH),
  },
  {
    name: 'inside the inset box, with just and dist as top',
    predict: (row) =>
      row.tIns +
      (row.anchor === 'just' || row.anchor === 'dist' ? 0 : anchorFraction(row.anchor)) *
        (row.boxH - row.tIns - row.bIns - row.blockH),
  },
  {
    name: 'inside the inset box, with just and dist as centre',
    predict: (row) =>
      row.tIns +
      (row.anchor === 'just' || row.anchor === 'dist' ? 0.5 : anchorFraction(row.anchor)) *
        (row.boxH - row.tIns - row.bIns - row.blockH),
  },
]);

/** `just` and `dist` stretch nothing: the line positions inside the block are `b`'s. */
const stretchRows = probesIn('anchor')
  .filter((p) => ['just', 'dist'].includes(String(p.vars['anchor'])))
  .map((p) => {
    const r = readOf.get(p.id);
    const twin = readOf.get(p.id.replace(/^anc-(just|dist)-/, 'anc-b-'));
    if (r === undefined || twin === undefined) throw new Error(`no bottom twin for ${p.id}`);
    return {
      id: p.id,
      measured: r.lines
        .map((l) => Math.round((must(l.top, p.id) - must(r.textTop, p.id)) * 8))
        .join(','),
      bottom: twin.lines
        .map((l) => Math.round((must(l.top, p.id) - must(twin.textTop, p.id)) * 8))
        .join(','),
    };
  });

choose(
  'just-and-dist-stretch',
  'do just and dist spread the lines to fill the frame',
  stretchRows,
  [
    {
      name: 'no: the lines sit exactly where bottom anchoring puts them',
      predict: (row) => row.bottom,
    },
    { name: 'yes: the lines spread', predict: () => 'spread' },
  ],
);

/* -------------------------------------------------------------------------- */
/* Q4 - the last line, and what spcFirstLastPara does to it                   */
/* -------------------------------------------------------------------------- */

interface LastLineRow {
  readonly id: string;
  readonly pct: number;
  readonly lines: number;
  readonly sz: number;
  readonly flp: boolean;
  readonly measured: number;
}

const lastLineRows: LastLineRow[] = probesIn('last-line').map((p) => {
  const r = readOf.get(p.id);
  if (r === undefined) throw new Error(`no reading for ${p.id}`);
  return {
    id: p.id,
    pct: Number(p.vars['pct']) / 100000,
    lines: Number(p.vars['lines']),
    sz: Number(p.vars['sz']) / 100,
    flp: p.vars['spcFirstLastPara'] === true,
    measured: must(r.textHeight, p.id),
  };
});

// The control family joins this question: at single spacing the trimmed and the
// full reading coincide, so only spacings above 100% separate them.
for (const row of controlRows) {
  lastLineRows.push({
    id: row.id,
    pct: 1,
    lines: row.lines,
    sz: row.sz,
    flp: false,
    measured: row.measured,
  });
}

choose(
  'block-height',
  'how tall a block of n lines is, and how much of the last line leading it keeps',
  lastLineRows,
  [
    {
      name: 'trimmed unless spcFirstLastPara',
      predict: (row) =>
        row.flp
          ? row.lines * advanceOf(row.sz, row.pct)
          : (row.lines - 1) * advanceOf(row.sz, row.pct) + trimmedLast(row.sz, row.pct),
    },
    { name: 'always a full advance', predict: (row) => row.lines * advanceOf(row.sz, row.pct) },
    {
      name: 'always trimmed',
      predict: (row) => (row.lines - 1) * advanceOf(row.sz, row.pct) + trimmedLast(row.sz, row.pct),
    },
    {
      name: 'trimmed with no floor',
      predict: (row) =>
        row.flp
          ? row.lines * advanceOf(row.sz, row.pct)
          : (row.lines - 1) * advanceOf(row.sz, row.pct) +
            0.75 * advanceOf(row.sz, row.pct) +
            ARIAL_B * row.sz,
    },
  ],
);

/* -------------------------------------------------------------------------- */
/* Q5 - anchorCtr                                                             */
/* -------------------------------------------------------------------------- */

interface AnchorCtrRow {
  readonly id: string;
  readonly on: boolean;
  readonly algn: string;
  readonly measured: string;
}

/** Each line's offset from the widest line's left edge, in eighths of a point. */
function lineOffsets(id: string): string {
  const r = readOf.get(id);
  if (r === undefined) throw new Error(`no reading for ${id}`);
  const left = must(r.left, id);
  return r.lines.map((l) => Math.round((must(l.left, id) - left) * 8)).join(',');
}

const anchorCtrRows: AnchorCtrRow[] = probesIn('anchor-ctr').map((p) => ({
  id: p.id,
  on: p.vars['anchorCtr'] === 'true',
  algn: String(p.vars['algn']),
  measured: lineOffsets(p.id),
}));

choose('anchor-ctr', 'what @anchorCtr moves, and what an absent one means', anchorCtrRows, [
  {
    name: 'on centres the block; absent is off',
    predict: (row) =>
      lineOffsets(`actr-${row.on ? '1' : '0'}-${row.algn}-square`) === row.measured
        ? row.measured
        : 'differs',
  },
  {
    name: 'absent is on',
    predict: (row) =>
      lineOffsets(`actr-1-${row.algn}-square`) === row.measured ? row.measured : 'differs',
  },
  { name: 'it is the same as algn="ctr"', predict: () => lineOffsets('actr-0-ctr-square') },
]);

/* -------------------------------------------------------------------------- */
/* Q6 - vertical text                                                         */
/* -------------------------------------------------------------------------- */

/** For each `ST_TextVerticalType`: which frame edge the anchor and the alignment start from. */
const VERT_AXES: Readonly<
  Record<string, { anchor: 'l' | 'r' | 't' | 'b'; algn: 'l' | 'r' | 't' | 'b' }>
> = {
  horz: { anchor: 't', algn: 'l' },
  vert: { anchor: 'r', algn: 't' },
  vert270: { anchor: 'l', algn: 'b' },
  eaVert: { anchor: 'r', algn: 't' },
  mongolianVert: { anchor: 'l', algn: 't' },
};

interface VertRow {
  readonly id: string;
  readonly vert: string;
  readonly anchor: string;
  readonly algn: string;
  readonly boxW: number;
  readonly boxH: number;
  readonly blockW: number;
  readonly blockH: number;
  readonly dLeft: number;
  readonly dTop: number;
  readonly measured: string;
}

/** The block's near-edge offsets, rounded to an eighth of a point. */
const edges = (row: VertRow): string =>
  `${String(Math.round(row.dLeft * 8))}/${String(Math.round(row.dTop * 8))}`;

const vertRows: VertRow[] = probesIn('vert-anchor')
  .filter((p) => p.id.startsWith('vanc-'))
  .map((p) => {
    const r = readOf.get(p.id);
    if (r === undefined) throw new Error(`no reading for ${p.id}`);
    const row = {
      id: p.id,
      vert: String(p.vars['vert']),
      anchor: String(p.vars['anchor']),
      algn: String(p.vars['algn']),
      boxW: p.rect.w,
      boxH: p.rect.h,
      blockW: must(r.textWidth, p.id),
      blockH: must(r.textHeight, p.id),
      dLeft: must(r.textLeft, p.id) - must(r.left, p.id),
      dTop: must(r.textTop, p.id) - must(r.top, p.id),
      measured: '',
    };
    return { ...row, measured: edges(row) };
  });

/**
 * Place the block from one edge given the fraction along that axis.
 *
 * `start` names the frame edge the axis runs from, so a reversed axis is a
 * reversed fraction rather than a second formula.
 */
function place(start: string, fraction: number, box: number, block: number): number {
  const room = box - block;
  return start === 'l' || start === 't' ? fraction * room : (1 - fraction) * room;
}

const FRACTION: Readonly<Record<string, number>> = { t: 0, l: 0, ctr: 0.5, b: 1, r: 1 };

/**
 * Scored on the stacking axis, where the reported extent is the line box and so
 * exact; the line axis carries the advance, which `BoundWidth` overstates.
 */
const anchorAxisRows = vertRows.map((row) => {
  const axes = VERT_AXES[row.vert];
  if (axes === undefined) throw new Error(`no axes for ${row.vert}`);
  const horizontal = axes.anchor === 'l' || axes.anchor === 'r';
  return {
    ...row,
    horizontal,
    thickness: horizontal ? row.blockW : row.blockH,
    box: horizontal ? row.boxW : row.boxH,
    measured: Math.round((horizontal ? row.dLeft : row.dTop) * 8),
  };
});

choose(
  'vert-anchor-axis',
  'which axis @anchor governs once the text is turned, and from which edge',
  anchorAxisRows,
  [
    {
      name: 'the stacking axis, from a start edge that depends on the vertical type',
      predict: (row) =>
        Math.round(
          place(
            VERT_AXES[row.vert]?.anchor ?? 't',
            FRACTION[row.anchor] ?? 0,
            row.box,
            row.thickness,
          ) * 8,
        ),
    },
    {
      name: 'always vertically, whatever the text does',
      predict: (row) =>
        row.horizontal
          ? Math.round(place('l', FRACTION[row.algn] ?? 0, row.box, row.thickness) * 8)
          : Math.round(place('t', FRACTION[row.anchor] ?? 0, row.box, row.thickness) * 8),
    },
    {
      name: 'the stacking axis, always from the same edge',
      predict: (row) =>
        Math.round(place('l', FRACTION[row.anchor] ?? 0, row.box, row.thickness) * 8),
    },
  ],
);

/** The alignment direction, read as a sign so the overstated advance cancels. */
const algnDirectionRows = ['horz', 'vert', 'vert270', 'eaVert', 'mongolianVert'].map((vert) => {
  const at = (algn: string): number => {
    const row = vertRows.find((r) => r.vert === vert && r.anchor === 't' && r.algn === algn);
    if (row === undefined) throw new Error(`no ${vert}/${algn} row`);
    const axes = VERT_AXES[vert];
    return axes?.anchor === 'l' || axes?.anchor === 'r' ? row.dTop : row.dLeft;
  };
  return { id: vert, vert, measured: Math.sign(at('r') - at('l')) };
});

choose('vert-algn-direction', 'which end of the line axis @algn="l" sits at', algnDirectionRows, [
  {
    name: 'the axis start, which vert270 alone reverses',
    predict: (row) => (VERT_AXES[row.vert]?.algn === 'b' ? -1 : 1),
  },
  { name: 'always the low coordinate', predict: () => 1 },
  { name: 'reversed by every vertical type', predict: (row) => (row.vert === 'horz' ? 1 : -1) },
]);

/** Which physical edge each inset moves the block away from, under rotation. */
interface VertInsetRow {
  readonly id: string;
  readonly vert: string;
  readonly measured: string;
}

const vertInsetRows: VertInsetRow[] = probesIn('vert')
  .filter((p) => p.id.startsWith('vert-ins-'))
  .map((p) => {
    const r = readOf.get(p.id);
    if (r === undefined) throw new Error(`no reading for ${p.id}`);
    const vert = String(p.vars['vert']);
    const dLeft = must(r.textLeft, p.id) - must(r.left, p.id);
    const dTop = must(r.textTop, p.id) - must(r.top, p.id);
    const dRight =
      must(r.left, p.id) + must(r.width, p.id) - must(r.textLeft, p.id) - must(r.textWidth, p.id);
    const dBottom =
      must(r.top, p.id) + must(r.height, p.id) - must(r.textTop, p.id) - must(r.textHeight, p.id);
    const near = ['vert', 'eaVert', 'wordArtVertRtl'].includes(vert) ? dRight : dLeft;
    const along = vert === 'vert270' ? dBottom : dTop;
    return {
      id: p.id,
      vert,
      measured: `${String(Math.round(near * 8))}/${String(Math.round(along * 8))}`,
    };
  });

/** `SKEW_INSETS` from the probe table. */
const SKEW = { l: 13, t: 20, r: 3, b: 5 } as const;

choose('vert-insets', 'which physical edge each inset holds under rotation', vertInsetRows, [
  {
    name: 'the insets stay on their own physical edges',
    predict: (row) => {
      const startsRight = ['vert', 'eaVert', 'wordArtVertRtl'].includes(row.vert);
      const bottomStart = row.vert === 'vert270';
      const near = startsRight ? SKEW.r : SKEW.l;
      const along = bottomStart ? SKEW.b : SKEW.t;
      return `${String(Math.round(near * 8))}/${String(Math.round(along * 8))}`;
    },
  },
  {
    name: 'the insets rotate with the text',
    predict: (row) => {
      const startsRight = ['vert', 'eaVert', 'wordArtVertRtl'].includes(row.vert);
      return `${String(Math.round((startsRight ? SKEW.t : SKEW.b) * 8))}/${String(Math.round(SKEW.l * 8))}`;
    },
  },
]);

/** Whether a vertical run asks GDI for the `@`-prefixed face. */
const verticalFaceRows = probesIn('vert')
  .filter((p) => p.id.startsWith('vert-') && p.vars['script'] === 'cjk')
  .map((p) => ({
    id: p.id,
    vert: String(p.vars['vert']),
    measured: drawn(p.id).some((t) => t.verticalFace),
  }));

choose('vertical-face', 'which vertical types ask for the @-prefixed face', verticalFaceRows, [
  {
    name: 'eaVert and mongolianVert only',
    predict: (row) => ['eaVert', 'mongolianVert'].includes(row.vert),
  },
  { name: 'every type but horz', predict: (row) => row.vert !== 'horz' && row.vert !== 'absent' },
  { name: 'never', predict: () => false },
]);

/* -------------------------------------------------------------------------- */
/* Q7 - overflow                                                              */
/* -------------------------------------------------------------------------- */

const LINES = ['Line one', 'Line two', 'Line three', 'Line four', 'Line five'];
const LINE_BOX = 21.6;

const overflowRows = probesIn('overflow')
  .filter((p) => p.vars['axis'] === 'vert')
  .map((p) => ({
    id: p.id,
    value: String(p.vars['value']),
    boxH: p.rect.h,
    measured: drawn(p.id)
      .map((t) => t.text)
      .join('|'),
  }));

/** How many whole line boxes fit, on a given comparison at the boundary. */
const fits = (boxH: number, strict: boolean): number => {
  let count = 0;
  while (count < LINES.length) {
    const bottom = (count + 1) * LINE_BOX;
    if (strict ? bottom >= boxH : bottom > boxH) break;
    count++;
  }
  return count;
};

choose('vert-overflow', 'what @vertOverflow draws when the text does not fit', overflowRows, [
  {
    name: 'a line is drawn only if it fits strictly; ellipsis replaces the first that does not',
    predict: (row) => {
      if (row.value === 'overflow' || row.value === 'absent') return LINES.join('|');
      const n = fits(row.boxH, true);
      const drawn_ = LINES.slice(0, n);
      return row.value === 'ellipsis' ? [...drawn_, '…'].join('|') : drawn_.join('|');
    },
  },
  {
    name: 'a line whose bottom lands exactly on the edge is drawn',
    predict: (row) => {
      if (row.value === 'overflow' || row.value === 'absent') return LINES.join('|');
      const drawn_ = LINES.slice(0, fits(row.boxH, false));
      return row.value === 'ellipsis' ? [...drawn_, '…'].join('|') : drawn_.join('|');
    },
  },
  {
    name: 'ellipsis appends to the last line that fits',
    predict: (row) => {
      if (row.value === 'overflow' || row.value === 'absent') return LINES.join('|');
      const n = fits(row.boxH, true);
      const drawn_ = LINES.slice(0, n);
      if (row.value === 'clip') return drawn_.join('|');
      return drawn_.map((t, i) => (i === n - 1 ? `${t}…` : t)).join('|');
    },
  },
  {
    name: 'clip draws everything and relies on a clip region',
    predict: (row) =>
      row.value === 'ellipsis'
        ? [...LINES.slice(0, fits(row.boxH, true)), '…'].join('|')
        : LINES.join('|'),
  },
]);

/** Overflow changes what is drawn, not where the block sits. */
const overflowLayoutRows = probesIn('overflow')
  .filter((p) => p.vars['axis'] === 'vert')
  .map((p) => {
    const r = readOf.get(p.id);
    if (r === undefined) throw new Error(`no reading for ${p.id}`);
    return { id: p.id, measured: Math.round(must(r.textHeight, p.id) * 8) };
  });

choose('overflow-layout', 'does @vertOverflow change the height of the block', overflowLayoutRows, [
  {
    name: 'no: the block is laid out in full whatever is drawn',
    predict: () => Math.round(108 * 8),
  },
  { name: 'yes: it shrinks to what fits', predict: () => Math.round(43.2 * 8) },
]);

/* -------------------------------------------------------------------------- */
/* Q8 - a:br                                                                  */
/* -------------------------------------------------------------------------- */

const breakMarkerRows = [
  {
    id: 'br-char-br',
    kind: 'char-break',
    measured: drawn('br-char-br').filter((t) => t.text === '•').length,
  },
  {
    id: 'br-char-2p',
    kind: 'char-paras',
    measured: drawn('br-char-2p').filter((t) => t.text === '•').length,
  },
  {
    id: 'br-num-br',
    kind: 'num-break',
    measured: drawn('br-num-br').filter((t) => /^\d+\.$/.test(t.text)).length,
  },
  {
    id: 'br-num-br3',
    kind: 'num-break3',
    measured: drawn('br-num-br3').filter((t) => /^\d+\.$/.test(t.text)).length,
  },
  {
    id: 'br-num-2p',
    kind: 'num-paras',
    measured: drawn('br-num-2p').filter((t) => /^\d+\.$/.test(t.text)).length,
  },
];

choose('break-bullets', 'does a:br start a new bullet or advance an autonumber', breakMarkerRows, [
  { name: 'no: one marker per a:p', predict: (row) => (row.kind.endsWith('paras') ? 2 : 1) },
  { name: 'yes: one marker per line', predict: (row) => (row.kind === 'num-break3' ? 3 : 2) },
]);

const breakIndentRows = ['br-indent-br', 'br-indent-2p', 'br-indent-pos-br'].map((id) => {
  const r = readOf.get(id);
  const p = probeOf.get(id);
  if (r === undefined || p === undefined) throw new Error(`no reading for ${id}`);
  const left = must(r.left, id);
  return {
    id,
    marL: Number(p.frame['marL'] ?? 0),
    measured: r.lines.map((l) => Math.round((must(l.left, id) - left) * 8)).join(','),
  };
});

choose('break-indent', 'does the line after a:br take the first-line indent', breakIndentRows, [
  {
    name: 'no: it starts at marL like any wrapped line',
    predict: (row) =>
      row.id === 'br-indent-pos-br'
        ? `${String(Math.round(56 * 8))},${String(Math.round(20 * 8))}`
        : `${String(Math.round(54 * 8))},${String(Math.round(54 * 8))}`,
  },
  {
    name: 'yes: it is a first line',
    predict: (row) =>
      row.id === 'br-indent-pos-br'
        ? `${String(Math.round(56 * 8))},${String(Math.round(56 * 8))}`
        : `${String(Math.round(54 * 8))},${String(Math.round(54 * 8))}`,
  },
]);

const breakSpacingRows = ['br-spc-br', 'br-spc-2p'].map((id) => {
  const r = readOf.get(id);
  if (r === undefined) throw new Error(`no reading for ${id}`);
  return { id, measured: Math.round(must(r.textHeight, id) * 8) };
});

choose('break-spacing', 'does a:br pay spcBef and spcAft', breakSpacingRows, [
  {
    name: 'no: only a paragraph boundary does',
    predict: (row) => Math.round((row.id === 'br-spc-br' ? 43.2 : 73.2) * 8),
  },
  { name: 'yes: it is a paragraph boundary', predict: () => Math.round(73.2 * 8) },
]);

const breakSizeRows = ['br-brsz-small', 'br-brsz-big', 'br-brsz-none'].map((id) => {
  const r = readOf.get(id);
  if (r === undefined) throw new Error(`no reading for ${id}`);
  return { id, measured: Math.round(must(r.textHeight, id) * 8) };
});

choose(
  'break-size',
  'does the a:rPr on an a:br set the height of the line it starts',
  breakSizeRows,
  [
    { name: 'no: the following run does', predict: () => Math.round(43.2 * 8) },
    {
      name: 'yes',
      predict: (row) =>
        Math.round(
          (row.id === 'br-brsz-small' ? 31.2 : row.id === 'br-brsz-big' ? 69.6 : 43.2) * 8,
        ),
    },
  ],
);

/* -------------------------------------------------------------------------- */
/* Q9 - the empty paragraph                                                   */
/* -------------------------------------------------------------------------- */

interface EndParaRow {
  readonly id: string;
  readonly endSz: number;
  readonly defSz: number | null;
  readonly lnSpc: number;
  readonly measured: number;
}

const endParaRows: EndParaRow[] = probesIn('end-para')
  .filter((p) => p.vars['where'] === 'middle')
  .map((p) => {
    const r = readOf.get(p.id);
    if (r === undefined) throw new Error(`no reading for ${p.id}`);
    const tops = r.paragraphs.map((x) => must(x.top, p.id));
    const first = tops[1];
    const second = tops[2];
    if (first === undefined || second === undefined)
      throw new Error(`${p.id}: fewer than 3 paragraphs`);
    const kind = String(p.vars['kind']);
    const stated = Number(p.vars['endSz']);
    return {
      id: p.id,
      endSz: kind === 'absent' || kind === 'empty-run' ? 1800 : stated,
      defSz: kind === 'defRPr' ? stated : kind === 'conflict' ? Number(p.vars['defSz']) : null,
      lnSpc: Number(p.vars['lnSpc'] ?? 100000) / 100000,
      measured: second - first,
    };
  });

choose('empty-paragraph-height', 'how tall an a:p with no runs is', endParaRows, [
  {
    name: '1.2 x the a:endParaRPr size, times lnSpc',
    predict: (row) => row.lnSpc * NATURAL * (row.endSz / 100),
  },
  {
    name: '1.2 x the a:defRPr size when there is one',
    predict: (row) => row.lnSpc * NATURAL * ((row.defSz ?? row.endSz) / 100),
  },
  { name: 'the a:endParaRPr size itself', predict: (row) => row.lnSpc * (row.endSz / 100) },
  { name: 'zero: an empty paragraph takes no room', predict: () => 0 },
]);

/* -------------------------------------------------------------------------- */
/* Q10 - columns                                                              */
/* -------------------------------------------------------------------------- */

interface ColumnRow {
  readonly id: string;
  readonly numCol: number;
  readonly spcCol: number;
  readonly rtl: boolean;
  readonly boxW: number;
  readonly measured: number;
}

const columnRows: ColumnRow[] = probesIn('columns')
  .filter((p) => Number(p.vars['numCol']) > 1)
  .map((p) => {
    const r = readOf.get(p.id);
    if (r === undefined) throw new Error(`no reading for ${p.id}`);
    const left = must(r.left, p.id);
    const xs = [
      ...new Set(r.paragraphs.map((x) => Math.round((must(x.left, p.id) - left) * 8) / 8)),
    ];
    const first = xs[0];
    const second = xs[1];
    if (first === undefined || second === undefined)
      throw new Error(`${p.id}: text filled one column`);
    return {
      id: p.id,
      numCol: Number(p.vars['numCol']),
      spcCol: p.vars['spcCol'] === 'absent' ? 0 : Number(p.vars['spcCol']),
      rtl: p.vars['rtlCol'] === '1',
      boxW: p.rect.w,
      measured: Math.abs(second - first),
    };
  });

choose('column-pitch', 'how wide a column is, and how far apart two are', columnRows, [
  {
    name: 'width = (frame - (n-1) x spcCol) / n, pitch = width + spcCol',
    predict: (row) => (row.boxW - (row.numCol - 1) * row.spcCol) / row.numCol + row.spcCol,
  },
  { name: 'pitch = frame / n, with spcCol inside it', predict: (row) => row.boxW / row.numCol },
  {
    name: 'width = frame / n, pitch = width + spcCol',
    predict: (row) => row.boxW / row.numCol + row.spcCol,
  },
]);

/* -------------------------------------------------------------------------- */
/* Q11 - does a:bodyPr inherit?                                               */
/* -------------------------------------------------------------------------- */

interface InheritRow {
  readonly id: string;
  readonly property: string;
  readonly at: string;
  readonly measured: string;
}

/** Everything PowerPoint resolved for the frame, as one string. */
function resolvedFrame(id: string): string {
  const r = readOf.get(id);
  if (r === undefined) throw new Error(`no reading for ${id}`);
  return [
    r.vAnchor,
    r.hAnchor,
    r.orientation,
    r.wordWrap,
    r.columnNumber,
    r.marginLeft,
    r.marginTop,
    r.marginRight,
    r.marginBottom,
  ].join('/');
}

const inheritRows: InheritRow[] = probesIn('inherit').map((p) => ({
  id: p.id,
  property: String(p.vars['property']),
  at: String(p.vars['at']),
  measured: resolvedFrame(p.id),
}));

choose(
  'bodyPr-inherits',
  'does a bare a:bodyPr on a slide inherit from the layout and the master',
  inheritRows,
  [
    {
      name: 'yes, per attribute, from any level of the placeholder chain',
      predict: (row) =>
        row.at === 'none'
          ? resolvedFrame(`inh-${row.property}-none`)
          : resolvedFrame(`inh-${row.property}-slide`),
    },
    {
      name: 'no: a bare a:bodyPr means the schema default',
      predict: (row) =>
        row.at === 'slide'
          ? resolvedFrame(`inh-${row.property}-slide`)
          : resolvedFrame(`inh-${row.property}-none`),
    },
    {
      name: 'from the layout only',
      predict: (row) =>
        row.at === 'master'
          ? resolvedFrame(`inh-${row.property}-none`)
          : resolvedFrame(`inh-${row.property}-slide`),
    },
  ],
);

/* -------------------------------------------------------------------------- */
/* the fixture                                                                */
/* -------------------------------------------------------------------------- */

const RESOLVED_FAMILIES = new Set(['inherit', 'inset', 'inset-hostile', 'columns', 'wrap', 'vert']);

/**
 * Thousandths of a point.
 *
 * Finer than anything measured, so the packing never rounds off a reading: a
 * 7.2pt inset is exact and an eighth-point grid would report it as 7.25.
 */
const milli = (v: number | null | undefined): number | null =>
  v === null || v === undefined ? null : Math.round(v * 1000);

/** A run of identical offsets becomes the one value, which most probes have. */
const runLength = (values: readonly (number | null)[]): unknown =>
  values.length > 1 && values.every((v) => v === values[0]) ? values[0] : values;

const rowOf = (p: ProbeInput): Record<string, unknown> => {
  const r = readOf.get(p.id);
  return {
    id: p.id,
    family: p.family,
    vars: p.vars,
    frame: p.frame,
    box: [p.rect.w, p.rect.h],
    ...(repairedDecks.has(deckOfProbe.get(p.id) ?? '') ? { repaired: true } : {}),
    measured:
      r === undefined
        ? null
        : {
            dLeft: milli((r.textLeft ?? 0) - (r.left ?? 0)),
            dTop: milli((r.textTop ?? 0) - (r.top ?? 0)),
            blockW: milli(r.textWidth),
            blockH: milli(r.textHeight),
            // What PowerPoint resolved, which only the families that ask about
            // resolution are scored on.
            ...(RESOLVED_FAMILIES.has(p.family)
              ? {
                  frame: [r.vAnchor, r.hAnchor, r.orientation, r.wordWrap, r.columnNumber],
                  margins: [r.marginLeft, r.marginTop, r.marginRight, r.marginBottom].map(milli),
                }
              : {}),
            lineLefts: runLength(r.lines.map((l) => milli((l.left ?? 0) - (r.left ?? 0)))),
            lineTops: runLength(r.lines.map((l) => milli((l.top ?? 0) - (r.top ?? 0)))),
            paraTops: runLength(r.paragraphs.map((x) => milli((x.top ?? 0) - (r.top ?? 0)))),
            paraLefts: runLength(r.paragraphs.map((x) => milli((x.left ?? 0) - (r.left ?? 0)))),
          },
    ...(p.isolate
      ? {
          drew: drawn(p.id).map((t) => t.text),
          verticalFace: drawn(p.id).some((t) => t.verticalFace),
        }
      : {}),
  };
};

const fixture = {
  $comment:
    'Ground truth for sub-phase 3.6: where a block of text sits inside a shape, how far in from each edge, and which way it runs. Every number was read out of PowerPoint through TextRange2.BoundTop and BoundLeft, or off the EMF it exported. BoundWidth is recorded and not trusted - it overstates an advance by about 5pt. See docs/adr/phase-3-text/0032-anchors-insets-and-vertical-text.md.',
  experiment: 'T6',
  subPhase: '3.6',
  generatedBy: 'tools/ground-truth/text/frames/analyse.ts',
  insetDefaults: INSET_DEFAULT,
  vertAxes: VERT_AXES,
  questions: scored,
  repairedDecks: [...repairedDecks].sort((a, b) => a.localeCompare(b)),
  probes: [...probeOf.values()].map(rowOf),
};

const json = `${JSON.stringify(fixture, null, 2)}\n`;
const out = fixturePath ?? join(dir, 'frames.json');
writeFileSync(out, json);

for (const q of scored) {
  const others = q.candidates
    .filter((c) => c.name !== q.winner)
    .map((c) => `${String(c.score)}`)
    .join(' ');
  console.log(
    `${q.id.padEnd(24)} ${String(q.rows).padStart(4)} rows  winner ${q.rows}/${q.rows}, rivals ${others}  ${q.winner}`,
  );
}
console.log(
  `\n${String(scored.length)} question(s), ${String(probeOf.size)} probe(s)\nwrote ${out} (${String(json.length)} bytes)`,
);
