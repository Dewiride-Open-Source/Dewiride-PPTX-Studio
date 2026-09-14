/**
 * Experiment C7, step 3 - score the candidate rules and emit the fixture.
 *
 * ```
 * node tools/ground-truth/model/tables/analyse.ts <dir> [--fixture <path>]
 * ```
 *
 * Every question is scored over the probes PowerPoint opened as written, against the rectangle
 * `Table.Cell(r, c).Shape` reported per position; it throws unless exactly one candidate fits
 * every row, then scores the authored decks under the winner. ADR 0056 has the findings.
 */

import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

import { fixtureJson } from '../../lib/json.ts';
import { entry, readZip } from '../../lib/zip.ts';

const args = process.argv.slice(2);
const dir = args[0];
if (dir === undefined)
  throw new Error('usage: tools/ground-truth/model/tables/analyse.ts <dir> [--fixture <path>]');
const fixtureAt = args.indexOf('--fixture');
const fixturePath = fixtureAt === -1 ? null : (args[fixtureAt + 1] ?? null);

const EMU_PER_POINT = 12700;
const TOL = 0.05;

/* -------------------------------------------------------------------------- */
/* inputs                                                                     */
/* -------------------------------------------------------------------------- */

interface CellInput {
  readonly gridSpan?: number | string;
  readonly rowSpan?: number | string;
  readonly hMerge?: boolean | string;
  readonly vMerge?: boolean | string;
  readonly text?: string;
  readonly lines?: number;
  readonly noBody?: boolean;
  readonly emptyBody?: boolean;
  readonly tcPr?: string;
}

interface DeckInput {
  readonly deck: string;
  readonly file: string;
  readonly family: string;
  readonly question: string;
  readonly hostile: boolean;
  readonly cols: readonly (number | string | null)[];
  readonly rows: readonly {
    readonly h: number | string | null;
    readonly cells: readonly CellInput[];
  }[];
  readonly frame: { readonly cx?: number; readonly cy?: number; readonly omit?: boolean } | null;
  readonly markup: string;
}

interface OmCell {
  readonly l: number | null;
  readonly t: number | null;
  readonly w: number | null;
  readonly h: number | null;
  readonly text: string | null;
  readonly margins: readonly number[] | null;
  readonly vAnchor: number | null;
  readonly orientation: number | null;
  readonly error?: string | null;
}

interface OmTable {
  readonly left: number;
  readonly top: number;
  readonly width: number;
  readonly height: number;
  readonly rows: number;
  readonly cols: number;
  readonly colWidths: readonly (number | null)[];
  readonly rowHeights: readonly (number | null)[];
  readonly cells: readonly (readonly OmCell[])[];
  readonly error?: string;
}

interface DeckReading {
  readonly deck: string;
  readonly opened: boolean;
  readonly repaired: boolean | null;
  readonly error: string | null;
  readonly table: OmTable | null;
  readonly resaved: string | null;
}

interface AuthoredEntry {
  readonly deck: string;
  readonly slide: number;
  readonly name: string;
  readonly before?: OmTable;
  readonly grid: OmTable;
}

const readJson = <T>(name: string): T =>
  JSON.parse(readFileSync(join(dir, name), 'utf8').replace(/^\uFEFF/, '')) as T;

const inputs = readJson<{ decks: readonly DeckInput[] }>('table-inputs.json');
const readings = readJson<{ decks: readonly DeckReading[] }>('table-readings.json');
const authored = readJson<{ log: readonly AuthoredEntry[] }>('author-tables-log.json');

function fail(message: string): never {
  throw new Error(`analyse-tables: ${message}`);
}

/* -------------------------------------------------------------------------- */
/* the markup, as attributes                                                  */
/* -------------------------------------------------------------------------- */

/** One `a:tc` as scanned: attribute values verbatim, its paragraphs' text joined by CR. */
interface ScannedCell {
  readonly attrs: Readonly<Record<string, string>>;
  readonly text: string;
}

interface ScannedTable {
  /** The frame's `a:ext` as `[cx, cy]`, or `null` without a `p:xfrm`. */
  readonly ext: readonly [string, string] | null;
  /** `a:tblPr`'s attributes and its `a:tableStyleId`, or `null` without one. */
  readonly tblPr: {
    readonly attrs: Readonly<Record<string, string>>;
    readonly styleId: string | null;
  } | null;
  readonly cols: readonly (string | null)[];
  readonly rows: readonly { readonly h: string | null; readonly cells: readonly ScannedCell[] }[];
}

function attributesOf(raw: string): Record<string, string> {
  const out: Record<string, string> = {};
  for (const m of raw.matchAll(/([\w:]+)="([^"]*)"/g)) out[m[1]!] = m[2]!;
  return out;
}

/** The `a:tbl` of a slide part, by tag scanning: enough for spans, sizes and cell text. */
function scanTable(xml: string): ScannedTable {
  const start = xml.indexOf('<a:tbl>');
  const end = xml.indexOf('</a:tbl>');
  if (start === -1 || end === -1) fail('no a:tbl in the part');
  const tbl = xml.slice(start, end + '</a:tbl>'.length);
  const ext = /<p:xfrm>.*?<a:ext cx="(\d+)" cy="(\d+)"\/>/.exec(xml);
  const tblPr = /<a:tblPr\b([^>]*?)(\/?)>/.exec(tbl);
  const styleId = /<a:tableStyleId>([^<]*)<\/a:tableStyleId>/.exec(tbl);
  const cols = [...tbl.matchAll(/<a:gridCol\b([^>]*?)\/?>/g)].map(
    (m) => attributesOf(m[1]!)['w'] ?? null,
  );
  const rows: { h: string | null; cells: ScannedCell[] }[] = [];
  const tags = /<(\/?)(a:tr|a:tc|a:p|a:t)\b([^>]*?)(\/?)>/g;
  let row: { h: string | null; cells: ScannedCell[] } | null = null;
  let cell: Record<string, string> | null = null;
  let paragraphs: string[] = [];
  let textStart = -1;
  let m: RegExpExecArray | null;
  while ((m = tags.exec(tbl)) !== null) {
    const [, close, name, raw, self] = m;
    if (name === 'a:tr' && close === '') {
      row = { h: attributesOf(raw!)['h'] ?? null, cells: [] };
      rows.push(row);
    } else if (name === 'a:tc' && close === '') {
      cell = attributesOf(raw!);
      paragraphs = [];
      if (self === '/') {
        row?.cells.push({ attrs: cell, text: '' });
        cell = null;
      }
    } else if (name === 'a:tc') {
      if (cell !== null) row?.cells.push({ attrs: cell, text: paragraphs.join('\r') });
      cell = null;
    } else if (name === 'a:p' && close === '') {
      paragraphs.push('');
    } else if (name === 'a:t' && close === '' && self === '') {
      textStart = tags.lastIndex;
    } else if (name === 'a:t' && close === '/') {
      paragraphs[paragraphs.length - 1] = (paragraphs.at(-1) ?? '') + tbl.slice(textStart, m.index);
    }
  }
  return {
    ext: ext === null ? null : [ext[1]!, ext[2]!],
    tblPr:
      tblPr === null ? null : { attrs: attributesOf(tblPr[1]!), styleId: styleId?.[1] ?? null },
    cols,
    rows,
  };
}

function slidePart(file: string, slide: number): string {
  const bytes = entry(readZip(readFileSync(file)), `ppt/slides/slide${String(slide)}.xml`);
  if (bytes === undefined) fail(`${file} has no slide ${String(slide)}`);
  return new TextDecoder().decode(bytes);
}

/* -------------------------------------------------------------------------- */
/* the candidate occupancy rules                                              */
/* -------------------------------------------------------------------------- */

/** The text a position reports: its anchor's cell, or nothing where no column was declared. */
function textAt(table: ScannedTable, anchor: Anchor): string {
  if (table.cols.length === 0) return '';
  return table.rows[anchor.r]?.cells[anchor.c]?.text ?? '';
}

/** A grid position's anchor and the anchor's clamped span, or `null` past the grid. */
interface Anchor {
  readonly r: number;
  readonly c: number;
  readonly rows: number;
  readonly cols: number;
}

type Occupancy = readonly (readonly Anchor[])[];

/** The span attribute as a count; `null` when it is not an integer. */
function spanValue(raw: string | undefined): number | null {
  if (raw === undefined) return 1;
  if (!/^[+-]?\d+$/.test(raw)) return null;
  return Number(raw);
}

const isTrue = (raw: string | undefined): boolean => raw === '1' || raw === 'true';

type SpanRule = (n: number) => number;
/** Zero is one and a negative runs to the edge. */
const spanEdge: SpanRule = (n) => (n < 0 ? Number.POSITIVE_INFINITY : Math.max(1, n));
/** Anything below one is one. */
const spanOne: SpanRule = (n) => Math.max(1, n);
/** The magnitude, as if the sign were a slip. */
const spanAbs: SpanRule = (n) => Math.max(1, Math.abs(n));

/** The dimensions PowerPoint reports: the grid's, and never fewer than one each. */
function dimensions(table: ScannedTable): { rows: number; cols: number } {
  return { rows: Math.max(1, table.rows.length), cols: Math.max(1, table.cols.length) };
}

/**
 * Spans on the anchor are the whole story: each `a:tc` takes the next column, a position
 * an earlier span claimed is covered whatever it says, and a span stops at the edge or at
 * the first claimed position.
 */
function bySpans(table: ScannedTable, rule: SpanRule): Occupancy {
  const { rows, cols } = dimensions(table);
  const grid: (Anchor | null)[][] = Array.from({ length: rows }, () =>
    Array.from({ length: cols }, () => null),
  );
  table.rows.forEach((row, r) => {
    row.cells.forEach((cell, c) => {
      if (c >= cols || grid[r]![c] !== null) return;
      const gs = rule(spanValue(cell.attrs['gridSpan']) ?? 1);
      const rs = rule(spanValue(cell.attrs['rowSpan']) ?? 1);
      let width = 1;
      while (c + width < cols && width < gs && grid[r]![c + width] === null) width += 1;
      const height = Math.min(rs, rows - r);
      const anchor = { r, c, rows: height, cols: width };
      for (let rr = r; rr < r + height; rr++)
        for (let cc = c; cc < c + width; cc++) grid[rr]![cc] = anchor;
    });
  });
  return grid.map((line, r) => line.map((a, c) => a ?? { r, c, rows: 1, cols: 1 }));
}

/** The flags are the whole story: hMerge joins the cell to the left, vMerge the one above. */
function byFlags(table: ScannedTable): Occupancy {
  const { rows, cols } = dimensions(table);
  const grid: Anchor[][] = [];
  for (let r = 0; r < rows; r++) {
    const line: Anchor[] = [];
    for (let c = 0; c < cols; c++) {
      const attrs = table.rows[r]?.cells[c]?.attrs ?? {};
      const h = isTrue(attrs['hMerge']) && c > 0;
      const v = isTrue(attrs['vMerge']) && r > 0;
      let anchor: Anchor = { r, c, rows: 1, cols: 1 };
      if (h && v) anchor = grid[r - 1]![c - 1]!;
      else if (h) anchor = line[c - 1]!;
      else if (v) anchor = grid[r - 1]![c]!;
      line.push(anchor);
    }
    grid.push(line);
  }
  return sized(grid);
}

/** Covered only where the span claims the position AND the cell there says so. */
function byBoth(table: ScannedTable): Occupancy {
  const spans = bySpans(table, spanEdge);
  const grid = spans.map((line, r) =>
    line.map((anchor, c) => {
      if (anchor.r === r && anchor.c === c) return anchor;
      const attrs = table.rows[r]?.cells[c]?.attrs ?? {};
      const wantsH = anchor.c < c;
      const wantsV = anchor.r < r;
      const ok = (!wantsH || isTrue(attrs['hMerge'])) && (!wantsV || isTrue(attrs['vMerge']));
      return ok ? anchor : { r, c, rows: 1, cols: 1 };
    }),
  );
  return sized(grid);
}

/** Covered where the span claims the position OR the cell there carries a flag. */
function byEither(table: ScannedTable): Occupancy {
  const spans = bySpans(table, spanEdge);
  const flags = byFlags(table);
  const grid = spans.map((line, r) =>
    line.map((anchor, c) => {
      if (anchor.r !== r || anchor.c !== c) return anchor;
      return flags[r]![c]!;
    }),
  );
  return sized(grid);
}

/** Recomputes every anchor's extent from the positions that point at it. */
function sized(grid: readonly (readonly Anchor[])[]): Occupancy {
  const extents = new Map<string, { r: number; c: number; maxR: number; maxC: number }>();
  grid.forEach((line, r) =>
    line.forEach((a, c) => {
      const key = `${String(a.r)},${String(a.c)}`;
      const e = extents.get(key) ?? { r: a.r, c: a.c, maxR: a.r, maxC: a.c };
      e.maxR = Math.max(e.maxR, r);
      e.maxC = Math.max(e.maxC, c);
      extents.set(key, e);
    }),
  );
  return grid.map((line) =>
    line.map((a) => {
      const e = extents.get(`${String(a.r)},${String(a.c)}`)!;
      return { r: e.r, c: e.c, rows: e.maxR - e.r + 1, cols: e.maxC - e.c + 1 };
    }),
  );
}

const OCCUPANCY: Readonly<Record<string, (table: ScannedTable) => Occupancy>> = {
  'spans, zero is one, negative to the edge': (t) => bySpans(t, spanEdge),
  'spans, below one is one': (t) => bySpans(t, spanOne),
  'spans, by magnitude': (t) => bySpans(t, spanAbs),
  'flags only': byFlags,
  'span and flag agree': byBoth,
  'span or flag': byEither,
};

/* -------------------------------------------------------------------------- */
/* scoring                                                                    */
/* -------------------------------------------------------------------------- */

const near = (a: number, b: number): boolean => Math.abs(a - b) <= TOL;
/** COM reports single-precision floats as doubles; three decimals is what it measured. */
const pt = (v: number | null): number | null => (v === null ? null : Math.round(v * 1000) / 1000);

/** The rectangle a position would report under an occupancy, from the sizes PowerPoint read. */
function predictedRect(
  om: OmTable,
  anchor: Anchor,
): { l: number; t: number; w: number; h: number } | null {
  const widths = om.colWidths.slice(anchor.c, anchor.c + anchor.cols);
  const heights = om.rowHeights.slice(anchor.r, anchor.r + anchor.rows);
  if (widths.some((w) => w === null) || heights.some((h) => h === null)) return null;
  const sum = (xs: readonly (number | null)[]): number =>
    xs.reduce<number>((a, b) => a + (b ?? 0), 0);
  return {
    l: om.left + sum(om.colWidths.slice(0, anchor.c)),
    t: om.top + sum(om.rowHeights.slice(0, anchor.r)),
    w: sum(widths),
    h: sum(heights),
  };
}

/** Every position's misfit under an occupancy, as `r,c: predicted / measured` lines. */
function misfits(om: OmTable, occupancy: Occupancy, table: ScannedTable): string[] {
  const out: string[] = [];
  if (occupancy.length !== om.rows || (occupancy[0]?.length ?? 0) !== om.cols) {
    return [
      `${String(occupancy.length)}x${String(occupancy[0]?.length ?? 0)} predicted, ${String(om.rows)}x${String(om.cols)} measured`,
    ];
  }
  om.cells.forEach((line, r) =>
    line.forEach((cell, c) => {
      if ((cell.error ?? null) !== null || cell.l === null) return;
      const rect = predictedRect(om, occupancy[r]![c]!);
      if (rect === null) return;
      const ok =
        near(rect.l, cell.l) &&
        near(rect.t, cell.t!) &&
        near(rect.w, cell.w!) &&
        near(rect.h, cell.h!);
      if (!ok)
        out.push(
          `${String(r + 1)},${String(c + 1)}: predicted ${String(rect.l)},${String(rect.t)} ${String(rect.w)}x${String(rect.h)} / measured ${String(cell.l)},${String(cell.t)} ${String(cell.w)}x${String(cell.h)}`,
        );
      const text = textAt(table, occupancy[r]![c]!);
      if (text !== (cell.text ?? ''))
        out.push(
          `${String(r + 1)},${String(c + 1)}: predicted "${text}" / measured "${cell.text ?? ''}"`,
        );
    }),
  );
  return out;
}

interface Scored {
  readonly name: string;
  readonly fits: number;
  readonly of: number;
  readonly wrong: readonly string[];
}

/** Exactly one perfect candidate, or the run stops here. */
function choose(question: string, scores: readonly Scored[]): string {
  const perfect = scores.filter((s) => s.fits === s.of);
  if (perfect.length !== 1) {
    fail(
      `${question}: ${String(perfect.length)} candidate(s) fit every row\n` +
        scores
          .map(
            (s) =>
              `  ${s.name}: ${String(s.fits)}/${String(s.of)}${s.wrong.length ? ` - ${s.wrong.slice(0, 4).join('; ')}` : ''}`,
          )
          .join('\n'),
    );
  }
  return perfect[0]!.name;
}

/* -------------------------------------------------------------------------- */
/* the probes                                                                 */
/* -------------------------------------------------------------------------- */

interface Probe {
  readonly input: DeckInput;
  readonly reading: DeckReading;
  readonly scanned: ScannedTable;
  readonly resaved: ScannedTable | null;
}

const probes: Probe[] = inputs.decks.map((input) => {
  const reading = readings.decks.find((d) => d.deck === input.deck);
  if (reading === undefined) fail(`${input.deck} was not read`);
  const scanned = scanTable(input.markup);
  const resaved =
    reading.resaved === null ? null : scanTable(slidePart(join(dir, reading.resaved), 1));
  return { input, reading, scanned, resaved };
});

const clean = probes.filter(
  (p) => p.reading.opened && p.reading.repaired === false && p.reading.table?.error === undefined,
);
const repaired = probes.filter((p) => p.reading.opened && p.reading.repaired === true);
const refused = probes.filter((p) => !p.reading.opened);
if (refused.length > 0) fail(`refused: ${refused.map((p) => p.input.deck).join(', ')}`);
for (const p of probes) {
  if (p.reading.opened && p.reading.table === null) fail(`${p.input.deck}: no table on the slide`);
}
const om = (p: Probe): OmTable => p.reading.table!;

// The occupancy, under every candidate, over every probe that opened as written.
const occupancyScores: Scored[] = Object.entries(OCCUPANCY).map(([name, model]) => {
  const wrong: string[] = [];
  for (const p of clean) {
    const bad = misfits(om(p), model(p.scanned), p.scanned);
    if (bad.length > 0) wrong.push(`${p.input.deck}: ${bad[0]!}`);
  }
  return { name, fits: clean.length - wrong.length, of: clean.length, wrong };
});
const occupancyWinner = choose('occupancy', occupancyScores);
const occupancy = OCCUPANCY[occupancyWinner]!;

// The frame's extent: the grid's sums, or the ext the file wrote.
const extentScores: Scored[] = (
  [
    [
      'the grid: the column widths by the row heights',
      (p: Probe) => ({ w: sum(om(p).colWidths), h: sum(om(p).rowHeights) }),
    ],
    [
      'the ext of p:xfrm',
      (p: Probe) => ({ w: extOf(p).cx / EMU_PER_POINT, h: extOf(p).cy / EMU_PER_POINT }),
    ],
  ] as const
).map(([name, predict]) => {
  const wrong: string[] = [];
  for (const p of clean) {
    const { w, h } = predict(p);
    if (!near(w, om(p).width) || !near(h, om(p).height))
      wrong.push(
        `${p.input.deck}: predicted ${String(w)}x${String(h)}, measured ${String(om(p).width)}x${String(om(p).height)}`,
      );
  }
  return { name, fits: clean.length - wrong.length, of: clean.length, wrong };
});
const extentWinner = choose('frame extent', extentScores);

function sum(xs: readonly (number | null)[]): number {
  return xs.reduce<number>((a, b) => a + (b ?? 0), 0);
}
function extOf(p: Probe): { cx: number; cy: number } {
  const colsEmu = p.input.cols.reduce<number>((a, w) => a + (typeof w === 'number' ? w : 0), 0);
  const rowsEmu = p.input.rows.reduce<number>((a, r) => a + (typeof r.h === 'number' ? r.h : 0), 0);
  return { cx: p.input.frame?.cx ?? colsEmu, cy: p.input.frame?.cy ?? rowsEmu };
}

// Row heights: 12 pt text at the 1.2 line height, an empty paragraph at the 18 pt default.
const LINE = 1.2;
function contentHeight(p: Probe, r: number): number {
  const row = p.input.rows[r];
  if (row === undefined) return 18 * LINE + 7.2;
  let tallest = 0;
  for (const cell of row.cells) {
    const empty = cell.noBody === true || cell.emptyBody === true || cell.text === '';
    const lines = empty ? 18 * LINE : (cell.lines ?? 1) * 12 * LINE;
    const tcPr = cell.tcPr ?? '';
    const marT = Number(/marT="(\d+)"/.exec(tcPr)?.[1] ?? 45720) / EMU_PER_POINT;
    const marB = Number(/marB="(\d+)"/.exec(tcPr)?.[1] ?? 45720) / EMU_PER_POINT;
    tallest = Math.max(tallest, lines + marT + marB);
  }
  return row.cells.length === 0 ? 18 * LINE + 7.2 : tallest;
}
const writtenHeight = (p: Probe, r: number): number => {
  const h = p.input.rows[r]?.h;
  return typeof h === 'number' ? h / EMU_PER_POINT : 0;
};
// A narrow column wraps its marker, which is line breaking's question and not the row's.
const heightProbes = clean.filter((p) => !p.input.deck.startsWith('col-'));
const heightScores: Scored[] = (
  [
    [
      'h is a minimum the content grows',
      (p: Probe, r: number) => Math.max(writtenHeight(p, r), contentHeight(p, r)),
    ],
    ['h is the height', writtenHeight],
    ['the content is the height', contentHeight],
  ] as const
).map(([name, predict]) => {
  const wrong: string[] = [];
  for (const p of heightProbes) {
    om(p).rowHeights.forEach((measured, r) => {
      if (measured === null) return;
      const predicted = predict(p, r);
      if (!near(predicted, measured))
        wrong.push(
          `${p.input.deck} row ${String(r + 1)}: predicted ${String(predicted)}, measured ${String(measured)}`,
        );
    });
  }
  return {
    name,
    fits: heightProbes.length - new Set(wrong.map((w) => w.split(' row ')[0])).size,
    of: heightProbes.length,
    wrong,
  };
});
const heightWinner = choose('row height', heightScores);

// Column widths: `w` as written, or never less than the side margins plus two points.
const MIN_GAP = 2;
function marginsOf(p: Probe, c: number): { l: number; r: number } {
  let l = 91440;
  let r = 91440;
  for (const row of p.input.rows) {
    const tcPr = row.cells[c]?.tcPr ?? '';
    const ml = /marL="(\d+)"/.exec(tcPr)?.[1];
    const mr = /marR="(\d+)"/.exec(tcPr)?.[1];
    if (ml !== undefined) l = Number(ml);
    if (mr !== undefined) r = Number(mr);
  }
  return { l: l / EMU_PER_POINT, r: r / EMU_PER_POINT };
}
const writtenWidth = (p: Probe, c: number): number => {
  const w = p.input.cols[c];
  if (typeof w === 'number') return w / EMU_PER_POINT;
  if (typeof w === 'string' && w.endsWith('pt')) return Number(w.slice(0, -2));
  return 0;
};
const widthProbes = clean.filter((p) => p.input.cols.length > 0);
const widthScores: Scored[] = (
  [
    [
      'w, and never less than the side margins plus 2 pt',
      (p: Probe, c: number) =>
        Math.max(writtenWidth(p, c), marginsOf(p, c).l + marginsOf(p, c).r + MIN_GAP),
    ],
    ['w is the width', writtenWidth],
    ['w, and never less than 16.4 pt', (p: Probe, c: number) => Math.max(writtenWidth(p, c), 16.4)],
  ] as const
).map(([name, predict]) => {
  const wrong: string[] = [];
  for (const p of widthProbes) {
    om(p).colWidths.forEach((measured, c) => {
      if (measured === null || c >= p.input.cols.length) return;
      const predicted = predict(p, c);
      if (!near(predicted, measured))
        wrong.push(
          `${p.input.deck} col ${String(c + 1)}: predicted ${String(predicted)}, measured ${String(measured)}`,
        );
    });
  }
  return {
    name,
    fits: widthProbes.length - new Set(wrong.map((w) => w.split(' col ')[0])).size,
    of: widthProbes.length,
    wrong,
  };
});
const widthWinner = choose('column width', widthScores);

// The cell properties PowerPoint read back, on the probes that wrote them.
const MSO_ANCHOR: Readonly<Record<number, string>> = { 1: 't', 3: 'ctr', 4: 'b' };
const MSO_ORIENTATION: Readonly<Record<number, string>> = { 1: 'horz', 2: 'vert270' };
function cellProps(p: Probe): {
  id: string;
  cells: {
    at: string;
    tcPr: string;
    margins: readonly (number | null)[] | null;
    anchor: string | null;
    vert: string | null;
  }[];
} {
  const cells = om(p).cells[0]!.map((cell, c) => ({
    at: `1,${String(c + 1)}`,
    tcPr: p.input.rows[0]?.cells[c]?.tcPr ?? '<a:tcPr/>',
    margins: cell.margins === null ? null : cell.margins.map(pt),
    anchor: cell.vAnchor === null ? null : (MSO_ANCHOR[cell.vAnchor] ?? String(cell.vAnchor)),
    vert:
      cell.orientation === null
        ? null
        : (MSO_ORIENTATION[cell.orientation] ?? String(cell.orientation)),
  }));
  return { id: p.input.deck, cells };
}
const propProbes = clean.filter((p) => p.input.deck.startsWith('tcpr-'));
for (const p of propProbes) {
  for (const cell of cellProps(p).cells) {
    const want = {
      marL: Number(/marL="(\d+)"/.exec(cell.tcPr)?.[1] ?? 91440) / EMU_PER_POINT,
      marT: Number(/marT="(\d+)"/.exec(cell.tcPr)?.[1] ?? 45720) / EMU_PER_POINT,
      marR: Number(/marR="(\d+)"/.exec(cell.tcPr)?.[1] ?? 91440) / EMU_PER_POINT,
      marB: Number(/marB="(\d+)"/.exec(cell.tcPr)?.[1] ?? 45720) / EMU_PER_POINT,
      anchor: /anchor="(\w+)"/.exec(cell.tcPr)?.[1] ?? 't',
      vert: /vert="(\w+)"/.exec(cell.tcPr)?.[1] ?? 'horz',
    };
    const got = cell.margins ?? [];
    const ok =
      near(got[0] ?? -1, want.marL) &&
      near(got[1] ?? -1, want.marT) &&
      near(got[2] ?? -1, want.marR) &&
      near(got[3] ?? -1, want.marB) &&
      cell.anchor === want.anchor &&
      cell.vert === want.vert;
    if (!ok) fail(`${p.input.deck} ${cell.at}: wrote ${cell.tcPr}, read ${JSON.stringify(cell)}`);
  }
}

/* -------------------------------------------------------------------------- */
/* the authored decks, under the winner                                       */
/* -------------------------------------------------------------------------- */

const authoredRows = authored.log.map((e) => {
  const scanned = scanTable(slidePart(join(dir, e.deck), e.slide));
  const bad = misfits(e.grid, occupancy(scanned), scanned);
  return { entry: e, scanned, bad };
});
const authoredWrong = authoredRows.filter((a) => a.bad.length > 0);
if (authoredWrong.length > 0) {
  fail(
    `the winner misreads PowerPoint's own markup:\n` +
      authoredWrong.map((a) => `  ${a.entry.name}: ${a.bad[0]!}`).join('\n'),
  );
}

/* -------------------------------------------------------------------------- */
/* the fixture                                                                */
/* -------------------------------------------------------------------------- */

const cellRecord = (cell: ScannedCell): Record<string, string> =>
  cell.text === '' ? cell.attrs : { ...cell.attrs, text: cell.text };

const summary = (t: ScannedTable) => ({
  ext: t.ext,
  tblPr: t.tblPr,
  cols: t.cols,
  rows: t.rows.map((r) => ({ h: r.h, cells: r.cells.map(cellRecord) })),
});

const omRecord = (t: OmTable) => ({
  frame: [pt(t.left), pt(t.top), pt(t.width), pt(t.height)],
  colWidths: t.colWidths.map(pt),
  rowHeights: t.rowHeights.map(pt),
  cells: t.cells.map((line) =>
    line.map((c) =>
      (c.error ?? null) !== null ? null : [pt(c.l), pt(c.t), pt(c.w), pt(c.h), c.text ?? ''],
    ),
  ),
});

const anchorsOf = (occ: Occupancy) => occ.map((line) => line.map((a) => [a.r + 1, a.c + 1]));

const fixture = {
  $comment:
    'Experiment C7 - which grid positions a merged cell occupies, what fixes a row height and a ' +
    'column width, and what PowerPoint writes back. Generated by ' +
    'tools/ground-truth/model/tables/analyse.ts from packages written by ' +
    'tools/ground-truth/model/tables/build-deck.ts, read through PowerPoint’s object model by ' +
    'tools/ground-truth/model/tables/read.ps1, and from decks PowerPoint authored in ' +
    'tools/ground-truth/model/tables/author.ps1. Positions are 1-based; sizes are points.',
  experiment: 'C7',
  subPhase: '4.1',
  generatedBy: 'tools/ground-truth/model/tables/analyse.ts',
  emuPerPoint: EMU_PER_POINT,
  findings: {
    occupancy: { winner: occupancyWinner, candidates: occupancyScores.map(scoreRecord) },
    extent: { winner: extentWinner, candidates: extentScores.map(scoreRecord) },
    rowHeight: {
      winner: heightWinner,
      lineHeight: LINE,
      candidates: heightScores.map(scoreRecord),
    },
    columnWidth: {
      winner: widthWinner,
      minimumGapPt: MIN_GAP,
      candidates: widthScores.map(scoreRecord),
    },
    cellProperties: propProbes.map(cellProps),
    repaired: repaired.map((p) => p.input.deck),
    canonical: authoredRows
      .filter((a) => a.entry.deck === 'pp-merges.pptx')
      .map((a) => ({
        name: a.entry.name,
        rows: a.scanned.rows.map((r) => r.cells.map((c) => c.attrs)),
      })),
  },
  probes: probes.map((p) => ({
    id: p.input.deck,
    family: p.input.family,
    question: p.input.question,
    hostile: p.input.hostile,
    opened: p.reading.opened,
    repaired: p.reading.repaired,
    markup: p.input.markup,
    frame: p.input.frame,
    om: p.reading.table === null ? null : omRecord(p.reading.table),
    anchors: p.reading.repaired === false ? anchorsOf(occupancy(p.scanned)) : null,
    resaved: p.resaved === null ? null : summary(p.resaved),
  })),
  authored: authoredRows.map((a) => ({
    deck: a.entry.deck,
    slide: a.entry.slide,
    name: a.entry.name,
    before: a.entry.before === undefined ? null : omRecord(a.entry.before),
    om: omRecord(a.entry.grid),
    markup: summary(a.scanned),
    anchors: anchorsOf(occupancy(a.scanned)),
  })),
};

function scoreRecord(s: Scored): {
  name: string;
  fits: number;
  of: number;
  wrong: readonly string[];
} {
  return { name: s.name, fits: s.fits, of: s.of, wrong: s.wrong };
}

const text = fixtureJson(fixture);
if (fixturePath === null) {
  console.log(text.slice(0, 4000));
} else {
  if (existsSync(fixturePath)) {
    const before = readFileSync(fixturePath, 'utf8');
    console.log(before === text ? 'fixture unchanged' : 'fixture rewritten');
  }
  writeFileSync(fixturePath, text);
}

console.log(
  `${String(probes.length)} probes: ${String(clean.length)} read as written, ${String(repaired.length)} repaired`,
);
for (const [label, scores] of [
  ['occupancy', occupancyScores],
  ['extent', extentScores],
  ['row height', heightScores],
  ['column width', widthScores],
] as const) {
  console.log(`${label}:`);
  for (const s of scores)
    console.log(`  ${s.fits === s.of ? '*' : ' '} ${s.name}: ${String(s.fits)}/${String(s.of)}`);
}
console.log(`authored: ${String(authoredRows.length)} slides, all fit under the winner`);
