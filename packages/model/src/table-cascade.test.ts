/**
 * The table-style cascade, against every cell and grid edge PowerPoint reported and drew in C9: each
 * probe's table is rebuilt from the fixture, resolved here, and held to COM's answer and the pixels.
 */

import {
  resolveColor,
  toHexColor,
  type ClrMap,
  type ColorContext,
  type Fill,
  type Line,
  type Rgba,
} from '@pptx-studio/paint';
import { parseXmlString } from '@pptx-studio/xml';
import { beforeAll, describe, expect, it } from 'vitest';

import main from '../../../corpus/ground-truth/table-cascade.json' with { type: 'json' };
import ends from '../../../corpus/ground-truth/table-cascade-ends.json' with { type: 'json' };
import sides from '../../../corpus/ground-truth/table-cascade-sides.json' with { type: 'json' };
import sweepOne from '../../../corpus/ground-truth/table-cascade-sweep-1.json' with { type: 'json' };
import sweepTwo from '../../../corpus/ground-truth/table-cascade-sweep-2.json' with { type: 'json' };
import sweepThree from '../../../corpus/ground-truth/table-cascade-sweep-3.json' with { type: 'json' };

import { isModelError } from './errors.js';
import { parseSheet, parseTheme } from './parse/sheet.js';
import { parseDefaultTextStyle } from './parse/text.js';
import { colorContextOf } from './resolve/resolve.js';
import {
  tableBackground,
  tableCellDiagonals,
  tableCellFill,
  TABLE_PART_ORDER,
  tableEdgeLine,
  tablePartsAt,
  tableStyleOf,
  tableTextLayer,
  themedEffects,
  themedFill,
  themedLine,
} from './resolve/table.js';
import {
  resolveBulletKind,
  resolveLatinTypeface,
  resolveRun,
  resolveSize,
  type TextContext,
} from './resolve/text.js';
import {
  TABLE_URI,
  tableGrid,
  type Table,
  type TableEdge,
  type TableSourced,
  type TableStylePartName,
  type Themeable,
} from './table.js';
import type { ListStyle, Paragraph, TextContent } from './text.js';
import type { Sheet, Theme } from './types.js';

/* -------------------------------------------------------------------------- */
/* the fixture                                                                */
/* -------------------------------------------------------------------------- */

interface CellRecord {
  readonly attrs?: string;
  readonly tcPr?: string;
  readonly body?: string;
}

interface SlideRecord {
  readonly name?: string;
  readonly rows: number;
  readonly cols: number;
  readonly mask: number;
  readonly style?: string;
  readonly rtl?: boolean;
  readonly tblPrFill?: string;
  readonly tblPrEffects?: string;
  readonly extScale?: number;
  readonly ph?: string;
  readonly cells?: Readonly<Record<string, CellRecord>>;
  /** Rows written with fewer `a:tc` than the grid has columns, keyed by row. */
  readonly short?: Readonly<Record<string, number>>;
  /** The slide's observations, one value per unit, in the fixture's order. */
  readonly units: readonly string[];
}

interface DeckRecord {
  readonly id: string;
  readonly group: string;
  readonly theme: 'A' | 'B' | 'C';
  readonly read: 'sweep' | 'full';
  readonly values: readonly string[];
  readonly slides: readonly SlideRecord[];
}

interface DeckFile {
  readonly id: string;
  readonly group: string;
  readonly theme: 'A' | 'B' | 'C';
  readonly read: 'sweep' | 'full';
  readonly style?: string;
  readonly width: number;
  readonly values: readonly string[];
  /** One `rows`x`cols`/`mask` per slide. */
  readonly shapes: string;
  /** What else each slide says, by index: inline, or in the main file a key of its `templates`. */
  readonly extras?: Readonly<Record<string, SlideExtras | string>>;
  /** Base64 of raw DEFLATE over every slide's codes in turn. */
  readonly codes: string;
}

type SlideExtras = Omit<SlideRecord, 'units' | 'rows' | 'cols' | 'mask'>;

const { digits: DIGITS, scale: SCALE, profileReach: REACH } = main.encoding;
const TEMPLATES = main.templates as unknown as Readonly<Record<string, SlideExtras>>;

/** A slide's extras, looked up in `templates` when the deck names them by key. */
function said(extras: SlideExtras | string | undefined): SlideExtras | undefined {
  if (typeof extras !== 'string') return extras;
  const found = TEMPLATES[extras];
  if (found === undefined) throw new Error(`no template ${extras}`);
  return found;
}

async function inflate(base64: string): Promise<string> {
  const bytes = Uint8Array.from(atob(base64), (ch) => ch.charCodeAt(0));
  const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
  return new Response(stream).text();
}

/** A deck as the fixture packs it: each observation `width` digits of `values`, slide after slide. */
async function unpack(deck: DeckFile): Promise<TextDeck> {
  const codes = await inflate(deck.codes);
  const valueAt = (code: string): string => {
    const n = [...code].reduce((acc, ch) => acc * DIGITS.length + DIGITS.indexOf(ch), 0);
    const value = deck.values[n];
    if (value === undefined) throw new Error(`${deck.id}: code ${code} names no value`);
    return value;
  };
  let at = 0;
  const slides = deck.shapes.split(' ').map((shape, k): SlideRecord => {
    const [, rows = 0, cols = 0, mask = 0] = (/^(\d+)x(\d+)\/(\d+)$/.exec(shape) ?? []).map(Number);
    const perCell = deck.read === 'full' ? 4 : 3;
    const count = rows * cols * perCell + (rows + 1) * cols + rows * (cols + 1);
    const units = Array.from({ length: count }, (_, u) =>
      valueAt(codes.slice(at + u * deck.width, at + (u + 1) * deck.width)),
    );
    at += count * deck.width;
    return {
      ...(deck.style === undefined ? {} : { style: deck.style }),
      ...said(deck.extras?.[String(k)]),
      rows,
      cols,
      mask,
      units,
    };
  });
  if (at !== codes.length) throw new Error(`${deck.id}: ${String(codes.length - at)} codes over`);
  return { ...deck, slides };
}

let DECKS: readonly TextDeck[] = [];
beforeAll(async () => {
  DECKS = await Promise.all(
    [main, ends, sides, sweepOne, sweepTwo, sweepThree].flatMap((file) =>
      (file.decks as unknown as readonly DeckFile[]).map(unpack),
    ),
  );
});
/** Per width in points, the coverage C9 measured across a double line's pixels. */
const DOUBLE = new Map<number, readonly number[]>(
  Object.entries(main.double as Readonly<Record<string, readonly number[]>>).map(([w, c]) => [
    Number(w),
    c,
  ]),
);

interface Observed {
  readonly fill: string;
  readonly text: string;
  readonly sides: string | null;
  readonly paint: string;
}

/** One slide's observations: per cell COM's fill, text and sides and the paint; per grid edge its pixels. */
function decode(deck: DeckRecord, slide: SlideRecord): { cells: Observed[][]; edges: string[] } {
  const next = (() => {
    let at = 0;
    return (): string => {
      const value = slide.units[at];
      at += 1;
      if (value === undefined) throw new Error(`${deck.id}: unit ${String(at)} is missing`);
      return value;
    };
  })();
  const cells = Array.from({ length: slide.rows }, () =>
    Array.from({ length: slide.cols }, () => ({
      fill: next(),
      text: next(),
      sides: deck.read === 'full' ? next() : null,
      paint: next(),
    })),
  );
  const edges = edgesOf(slide.rows, slide.cols).map(() => next());
  return { cells, edges };
}

/** Every grid edge, horizontal first: the order the fixture lists them in. */
function edgesOf(rows: number, cols: number): TableEdge[] {
  const out: TableEdge[] = [];
  for (let row = 0; row <= rows; row++)
    for (let col = 0; col < cols; col++) out.push({ axis: 'h', row, col });
  for (let row = 0; row < rows; row++)
    for (let col = 0; col <= cols; col++) out.push({ axis: 'v', row, col });
  return out;
}

/* -------------------------------------------------------------------------- */
/* rebuilding each probe                                                      */
/* -------------------------------------------------------------------------- */

const FLAGS = ['firstRow', 'lastRow', 'firstCol', 'lastCol', 'bandRow', 'bandCol'] as const;
const EMU = 12700;
const TEXT_BODY =
  '<a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="1200"/><a:t>X</a:t></a:r></a:p></a:txBody>';
const NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' +
  ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
  ' xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';

/** The probe's slide as `build-deck.ts` wrote it, parsed by this package. */
function slideOf(slide: SlideRecord): Omit<Sheet, 'parent' | 'theme'> {
  const attrs =
    (slide.rtl === true ? ' rtl="1"' : '') +
    FLAGS.filter((_, k) => (slide.mask & (1 << k)) !== 0)
      .map((f) => ` ${f}="1"`)
      .join('');
  const inner =
    (slide.tblPrFill ?? '') +
    (slide.tblPrEffects ?? '') +
    (slide.style === undefined ? '' : `<a:tableStyleId>${slide.style}</a:tableStyleId>`);
  const tblPr = inner === '' ? `<a:tblPr${attrs}/>` : `<a:tblPr${attrs}>${inner}</a:tblPr>`;
  const rows = Array.from({ length: slide.rows }, (_, r) => {
    const cells = Array.from({ length: slide.short?.[String(r)] ?? slide.cols }, (_, c) => {
      const cell = slide.cells?.[`${String(r)},${String(c)}`] ?? {};
      return `<a:tc${cell.attrs ?? ''}>${cell.body ?? TEXT_BODY}${cell.tcPr ?? '<a:tcPr/>'}</a:tc>`;
    });
    return `<a:tr h="${String(36 * EMU)}">${cells.join('')}</a:tr>`;
  });
  const tbl =
    `<a:tbl>${tblPr}<a:tblGrid>${`<a:gridCol w="${String(72 * EMU)}"/>`.repeat(slide.cols)}</a:tblGrid>` +
    `${rows.join('')}</a:tbl>`;
  const xml =
    `<p:sld ${NS}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/>` +
    '</p:nvGrpSpPr><p:grpSpPr/><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="t"/>' +
    `<p:cNvGraphicFramePr/>${slide.ph === undefined ? '<p:nvPr/>' : `<p:nvPr><p:ph ${slide.ph}/></p:nvPr>`}</p:nvGraphicFramePr>` +
    `<p:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></p:xfrm><a:graphic><a:graphicData uri="${TABLE_URI}">` +
    `${tbl}</a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>`;
  return parseSheet(parseXmlString(xml).root, '/ppt/slides/slide1.xml');
}

function tableOf(slide: SlideRecord): Table {
  const table = slideOf(slide).shapes[0]?.table;
  if (table === undefined) throw new Error('the probe rebuilt no table');
  return table;
}

const STYLES = new Map<string, ReturnType<typeof tableStyleOf>>();

/** `tableStyleOf`, parsed once per GUID: nothing in these tests changes a style it is given. */
function styleOf(table: Table): ReturnType<typeof tableStyleOf> {
  const key = table.props?.style?.id ?? '';
  if (!STYLES.has(key)) STYLES.set(key, tableStyleOf(table));
  return STYLES.get(key) ?? null;
}

interface Palette {
  readonly theme: Theme;
  readonly colors: ColorContext;
}

const PALETTES = Object.fromEntries(
  (['A', 'B', 'C'] as const).map((key) => {
    const record = main.themes[key];
    const theme = parseTheme(parseXmlString(record.xml).root, '/ppt/theme/theme1.xml');
    return [key, { theme, colors: { scheme: theme.scheme, map: record.clrMap as ClrMap } }];
  }),
) as Record<'A' | 'B' | 'C', Palette>;

/* -------------------------------------------------------------------------- */
/* what the model predicts, in the fixture's terms                            */
/* -------------------------------------------------------------------------- */

const hexOf = (rgba: Rgba): string => toHexColor(rgba).toUpperCase();

/** A fill as COM reports it: `RRGGBB/transparency`, or `none`. */
function fillText(sourced: TableSourced<Themeable<Fill>> | null, palette: Palette): string {
  if (sourced === null) return 'none';
  const { fill, phClr } = themedFill(sourced.value, palette.theme, palette.colors);
  if (fill === null || fill.type === 'none') return 'none';
  if (fill.type !== 'solid') throw new Error(`a ${fill.type} fill is not a COM colour`);
  const rgba = resolveColor(
    fill.color,
    phClr === null ? palette.colors : { ...palette.colors, phClr },
  );
  return `${hexOf(rgba)}/${String(Math.round((1 - rgba.a) * 1000) / 1000)}`;
}

interface Drawn {
  readonly rgb: string;
  readonly weight: number;
  readonly cmpd: string;
}

/** A grid edge's line as it draws: colour, width in points, compound; `null` for none. */
function drawn(sourced: TableSourced<Themeable<Line>> | null, palette: Palette): Drawn | null {
  if (sourced === null) return null;
  const { line, phClr } = themedLine(sourced.value, palette.theme, palette.colors);
  if (line?.fill === null || line?.fill === undefined || line.fill.type === 'none') return null;
  if (line.fill.type !== 'solid') throw new Error(`a ${line.fill.type} line is not drawn here`);
  const rgba = resolveColor(
    line.fill.color,
    phClr === null ? palette.colors : { ...palette.colors, phClr },
  );
  return { rgb: hexOf(rgba), weight: (line.w ?? 9525) / EMU, cmpd: line.cmpd ?? 'sng' };
}

type Rgb = readonly [number, number, number];
const rgbOf = (h: string): Rgb =>
  [0, 2, 4].map((i) => Number.parseInt(h.slice(i, i + 2), 16)) as unknown as Rgb;
const hexRgb = (c: Rgb): string =>
  c.map((v) => v.toString(16).toUpperCase().padStart(2, '0')).join('');
const roundDown = (n: number): number =>
  n % 1000 > 500 ? Math.floor(n / 1000) + 1 : Math.floor(n / 1000);

/** A translucent paint over another, each term rounded on its own (C9), ties down (ADR 0021). */
function over(under: Rgb, fill: string): Rgb {
  if (fill === 'none') return under;
  const [rgb = '', transparency = '0'] = fill.split('/');
  const a = Math.round((1 - Number(transparency)) * 1000);
  const top = rgbOf(rgb);
  return under.map(
    (u, i) => roundDown(a * (top[i] ?? 0)) + roundDown((1000 - a) * u),
  ) as unknown as Rgb;
}

/** An antialiased pixel, `c * line + (1 - c) * under`, the sum rounded with ties down. */
function blend(line: string, under: string, c: number): string {
  const l = rgbOf(line);
  const k = Math.round(c * 1000);
  return hexRgb(
    rgbOf(under).map((u, i) => roundDown(k * (l[i] ?? 0) + (1000 - k) * u)) as unknown as Rgb,
  );
}

/** The pixels a line paints across a grid line between paints `a` and `b`, and their descriptor. */
function describeLine(line: Drawn | null, a: string, b: string, observed: string): string {
  const out = Array.from({ length: 2 * REACH }, (_, k) => (k < REACH ? a : b));
  const drawnPixels = profileFrom(observed, a, b);
  if (line !== null) {
    const half = (line.weight * SCALE) / 2;
    const coverage = line.cmpd === 'dbl' ? DOUBLE.get(line.weight) : undefined;
    for (let k = REACH - Math.ceil(half); k < REACH + Math.ceil(half); k++) {
      if (line.cmpd !== 'sng' && coverage === undefined) {
        out[k] = '?';
        continue;
      }
      // A single line covers each pixel by its overlap with [-half, half) about the grid line.
      const overlap = Math.min(k + 1, REACH + half) - Math.max(k, REACH - half);
      const c = coverage === undefined ? overlap : (coverage[k - (REACH - Math.ceil(half))] ?? 0);
      if (c >= 1) out[k] = line.rgb;
      else if (c > 0) {
        // How much a partly covered pixel is covered is the rasteriser's (4.4): any blend of the two.
        const seen = drawnPixels[k] ?? '';
        out[k] = blendOf(line.rgb, out[k] ?? '', seen) ? seen : blend(line.rgb, out[k] ?? '', c);
      }
    }
  }
  let s = 0;
  while (s < out.length && out[s] === a) s++;
  let e = out.length;
  while (e > s && out[e - 1] === b) e--;
  const runs: string[] = [];
  for (let k = s; k < e;) {
    let n = 1;
    while (k + n < e && out[k + n] === out[k]) n++;
    runs.push(`${out[k] ?? ''}*${String(n)}`);
    k += n;
  }
  return `${String(s - REACH)}|${runs.join(',')}`;
}

/** The pixels a descriptor stands for over paints `a` and `b`. */
function profileFrom(descriptor: string, a: string, b: string): string[] {
  const [offset = '0', runs = ''] = descriptor.split('|');
  const out: string[] = Array.from({ length: Number(offset) + REACH }, () => a);
  for (const run of runs === '' ? [] : runs.split(',')) {
    const [colour = '', n = '0'] = run.split('*');
    for (let k = 0; k < Number(n); k++) out.push(colour);
  }
  while (out.length < 2 * REACH) out.push(b);
  return out;
}

/** Whether `x` is `line` over `under` at a coverage strictly between none and all, within a unit. */
function blendOf(line: string, under: string, x: string): boolean {
  if (![line, under, x].every((h) => /^[0-9A-F]{6}$/.test(h))) return false;
  const [l, u, o] = [rgbOf(line), rgbOf(under), rgbOf(x)];
  const near = (p: Rgb): boolean => p.every((v, j) => Math.abs(v - (o[j] ?? -9)) <= 1);
  const spread = l.map((v, i) => Math.abs(v - (u[i] ?? 0)));
  const i = spread.indexOf(Math.max(...spread));
  const span = (l[i] ?? 0) - (u[i] ?? 0);
  if (Math.abs(span) <= 2) return near(l) || near(u);
  if (x === line || x === under) return false;
  const c = ((o[i] ?? 0) - (u[i] ?? 0)) / span;
  return (
    c > 0 && c < 1 && l.every((v, j) => Math.abs(c * v + (1 - c) * (u[j] ?? 0) - (o[j] ?? 0)) <= 1)
  );
}

/** Logical grid edge `edge` of an rtl table, as the pixels see it: mirrored. */
function visual(edge: TableEdge, cols: number): TableEdge {
  return edge.axis === 'h'
    ? { ...edge, col: cols - 1 - edge.col }
    : { ...edge, col: cols - edge.col };
}

/* -------------------------------------------------------------------------- */
/* the questions                                                              */
/* -------------------------------------------------------------------------- */

const STYLED = new Set(['control', 'sweep', 'ends', 'sides', 'direct', 'merge', 'rtl', 'tblpr']);
/** Every cell or every grid edge of the slides in `groups`, counted from their shapes alone. */
function countOf(groups: ReadonlySet<string>, unit: 'cells' | 'edges'): number {
  return DECKS.filter((deck) => groups.has(deck.group))
    .flatMap((deck) => deck.slides)
    .reduce(
      (n, s) =>
        n + (unit === 'cells' ? s.rows * s.cols : (s.rows + 1) * s.cols + s.rows * (s.cols + 1)),
      0,
    );
}

function forEachSlide(
  groups: ReadonlySet<string>,
  visit: (deck: DeckRecord, slide: SlideRecord, index: number) => void,
): number {
  let seen = 0;
  for (const deck of DECKS) {
    if (!groups.has(deck.group)) continue;
    deck.slides.forEach((slide, index) => {
      visit(deck, slide, index);
      seen += 1;
    });
  }
  return seen;
}

describe('the table-style cascade, against every table C9 asked PowerPoint about', () => {
  it('scores every group of the fixture: dashes and theme C’s samples below, by their own tests', () => {
    expect(new Set(DECKS.map((deck) => deck.group))).toEqual(
      new Set([...STYLED, 'text', 'background', 'dash']),
    );
  });

  it('gives every cell the fill COM reported', () => {
    let cells = 0;
    const wrong: string[] = [];
    forEachSlide(STYLED, (deck, slide, index) => {
      const table = tableOf(slide);
      const style = styleOf(table);
      const grid = tableGrid(table);
      const { cells: observed } = decode(deck, slide);
      observed.forEach((row, r) =>
        row.forEach((cell, c) => {
          cells += 1;
          const predicted = fillText(tableCellFill(table, style, grid, r, c), PALETTES[deck.theme]);
          if (predicted !== cell.fill)
            wrong.push(
              `${deck.id}#${String(index)}@${String(r)},${String(c)}: ${predicted} for ${cell.fill}`,
            );
        }),
      );
    });
    expect(wrong.slice(0, 5)).toEqual([]);
    expect(cells).toBe(countOf(STYLED, 'cells'));
  });

  it('draws every grid edge PowerPoint drew, pixel for pixel across it', () => {
    let edges = 0;
    let crossed = 0;
    const wrong: string[] = [];
    forEachSlide(STYLED, (deck, slide, index) => {
      const table = tableOf(slide);
      const style = styleOf(table);
      const grid = tableGrid(table);
      const { edges: observed } = decode(deck, slide);
      const strokes = diagonalStrokes(slide, table, PALETTES[deck.theme]);
      edgesOf(slide.rows, slide.cols).forEach((edge, k) => {
        // The fixture lists edges as drawn; an rtl table is drawn mirrored (C9).
        const logical = slide.rtl === true ? visual(edge, slide.cols) : edge;
        const [a = '', b = '', ...rest] = (observed[k] ?? '').split('|');
        const described = rest.join('|');
        edges += 1;
        if (crossesProfile(strokes, edge)) {
          crossed += 1;
          return;
        }
        const predicted = describeLine(
          drawn(tableEdgeLine(table, style, grid, logical), PALETTES[deck.theme]),
          a,
          b,
          described,
        );
        if (predicted !== described) {
          wrong.push(
            `${deck.id}#${String(index)} ${edge.axis}${String(edge.row)},${String(edge.col)}: ${predicted} for ${described}`,
          );
        }
      });
    });
    expect(wrong.slice(0, 5)).toEqual([]);
    expect(edges).toBe(countOf(STYLED, 'edges'));
    // In each of the two tall merged cells, the diagonal runs through the profile of the line it
    // covers and of the line on its left.
    expect(crossed).toBe(4);
  });

  it('paints every cell over the page, the table background under it, as PowerPoint composited them', () => {
    let cells = 0;
    const wrong: string[] = [];
    forEachSlide(STYLED, (deck, slide, index) => {
      const table = tableOf(slide);
      const style = styleOf(table);
      const grid = tableGrid(table);
      const palette = PALETTES[deck.theme];
      const background = fillText(tableBackground(table, style).fill, palette);
      const { cells: observed } = decode(deck, slide);
      observed.forEach((row, r) =>
        row.forEach((cell, c) => {
          cells += 1;
          const fill = fillText(tableCellFill(table, style, grid, r, c), palette);
          const predicted = hexRgb(over(over([255, 255, 255], background), fill));
          if (predicted !== cell.paint)
            wrong.push(
              `${deck.id}#${String(index)}@${String(r)},${String(c)}: ${predicted} for ${cell.paint}`,
            );
        }),
      );
    });
    expect(wrong.slice(0, 5)).toEqual([]);
    expect(cells).toBe(countOf(STYLED, 'cells'));
  });
});

/* -------------------------------------------------------------------------- */
/* text, through the text cascade                                             */
/* -------------------------------------------------------------------------- */

interface TextDeck extends DeckRecord {
  readonly defaultTextStyle?: string;
  readonly txStyles?: string;
  readonly masterShapes?: readonly string[];
  readonly layoutShapes?: readonly string[];
}

/** The deck's master under its theme, colour map and text styles, and the probe slide under it. */
function contextOf(
  deck: TextDeck,
  slide: SlideRecord,
): { sheet: Sheet; defaultTextStyle: ListStyle | undefined } {
  const map = main.themes[deck.theme].clrMap as Readonly<Record<string, string>>;
  const clrMap = `<p:clrMap ${Object.entries(map)
    .map(([k, v]) => `${k}="${v}"`)
    .join(' ')}/>`;
  const tree = (shapes: readonly string[] | undefined): string =>
    '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
    `<p:grpSpPr/>${(shapes ?? []).join('')}</p:spTree>`;
  const masterXml =
    `<p:sldMaster ${NS}><p:cSld>${tree(deck.masterShapes)}</p:cSld>${clrMap}` +
    `${deck.txStyles ?? ''}</p:sldMaster>`;
  const master: Sheet = {
    ...parseSheet(parseXmlString(masterXml).root, '/ppt/slideMasters/slideMaster1.xml'),
    parent: null,
    theme: PALETTES[deck.theme].theme,
  };
  // The layout build-deck.ts writes: `obj` with the deck's layout shapes, or a blank one.
  const layoutXml =
    `<p:sldLayout ${NS} type="${deck.layoutShapes === undefined ? 'blank' : 'obj'}">` +
    `<p:cSld>${tree(deck.layoutShapes)}</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>` +
    '</p:sldLayout>';
  const layout: Sheet = {
    ...parseSheet(parseXmlString(layoutXml).root, '/ppt/slideLayouts/slideLayout1.xml'),
    parent: master,
    theme: null,
  };
  const sheet: Sheet = { ...slideOf(slide), parent: layout, theme: null };
  const presentation = parseXmlString(
    `<p:presentation ${NS}>${deck.defaultTextStyle ?? ''}</p:presentation>`,
  );
  return {
    sheet,
    defaultTextStyle: parseDefaultTextStyle(presentation.root, '/ppt/presentation.xml'),
  };
}

/** A cell's first character as the text cascade resolves it: colour, bold, face, and size in points. */
/** The text context of cell (r, c)'s first run, as a renderer would build it. */
function cellText(
  table: Table,
  style: ReturnType<typeof tableStyleOf>,
  sheet: Sheet,
  defaultTextStyle: ListStyle | undefined,
  r: number,
  c: number,
): { context: TextContext; paragraph: Paragraph; run: TextContent | undefined } {
  const grid = tableGrid(table);
  const owner = grid.positions[r]?.[c];
  const body = owner?.cell?.text;
  const shape = sheet.shapes[0];
  if (body === undefined || shape === undefined)
    throw new Error(`no text in cell ${String(r)},${String(c)}`);
  const paragraph = body.paragraphs[0];
  if (paragraph === undefined) throw new Error('an empty cell body');
  const context: TextContext = {
    sheet,
    shape,
    defaultTextStyle,
    cell: { lstStyle: body.lstStyle, layer: tableTextLayer(table, style, grid, r, c) },
  };
  return { context, paragraph, run: paragraph.content[0] };
}

function textOf(
  table: Table,
  style: ReturnType<typeof tableStyleOf>,
  sheet: Sheet,
  defaultTextStyle: ListStyle | undefined,
  r: number,
  c: number,
): { rgb: string; bold: boolean; face: string; size: number; bullet: boolean } {
  const { context, paragraph, run } = cellText(table, style, sheet, defaultTextStyle, r, c);
  const fill = resolveRun(context, paragraph, run, (p) => p.fill)?.value;
  if (fill?.type !== 'solid')
    throw new Error(`cell ${String(r)},${String(c)} resolves no solid text colour`);
  return {
    rgb: hexOf(resolveColor(fill.color, colorContextOf(sheet))),
    bold: resolveRun(context, paragraph, run, (p) => p.b)?.value ?? false,
    face: resolveLatinTypeface(context, paragraph, run).value,
    size: resolveSize(context, paragraph, run).value / 100,
    bullet: resolveBulletKind(context, paragraph)?.value === 'char',
  };
}

describe('a table cell’s text, through the text cascade', () => {
  it('has the colour, weight and face COM reported in every cell of every styled table', () => {
    let cells = 0;
    const wrong: string[] = [];
    forEachSlide(STYLED, (deck, slide, index) => {
      const table = tableOf(slide);
      const style = styleOf(table);
      const { sheet, defaultTextStyle } = contextOf(deck, slide);
      const grid = tableGrid(table);
      decode(deck, slide).cells.forEach((row, r) =>
        row.forEach((cell, c) => {
          cells += 1;
          // A position no `a:tc` reached is padded empty (ADR 0056): COM finds no text in it.
          if (grid.positions[r]?.[c]?.cell === null) {
            if (cell.text !== 'empty')
              wrong.push(
                `${deck.id}#${String(index)}@${String(r)},${String(c)}: text ${cell.text}`,
              );
            return;
          }
          const got = textOf(table, style, sheet, defaultTextStyle, r, c);
          const predicted = `${got.rgb}/${got.bold ? 'b' : '-'}/${got.face}`;
          const observed = cell.text.split('/').slice(0, 3).join('/');
          if (predicted !== observed)
            wrong.push(
              `${deck.id}#${String(index)}@${String(r)},${String(c)}: ${predicted} for ${observed}`,
            );
        }),
      );
    });
    expect(wrong.slice(0, 5)).toEqual([]);
    expect(cells).toBe(countOf(STYLED, 'cells'));
  });

  it('sits between the cell’s own list style and the package’s, on every rung of the ladder', () => {
    let cells = 0;
    const wrong: string[] = [];
    forEachSlide(new Set(['text']), (deck, slide, index) => {
      const table = tableOf(slide);
      const style = styleOf(table);
      const { sheet, defaultTextStyle } = contextOf(deck, slide);
      decode(deck, slide).cells.forEach((row, r) =>
        row.forEach((cell, c) => {
          cells += 1;
          const got = textOf(table, style, sheet, defaultTextStyle, r, c);
          const [rgb, bold, face, , size, bullet] = cell.text.split('/');
          const predicted = `${got.rgb}/${got.bold ? 'b' : '-'}/${got.face}/${String(got.size)}/${got.bullet ? 'bullet' : '-'}`;
          const observed = `${String(rgb)}/${String(bold)}/${String(face)}/${String(size)}/${String(bullet)}`;
          if (predicted !== observed)
            wrong.push(
              `${deck.id}#${String(index)}@${String(r)},${String(c)}: ${predicted} for ${observed}`,
            );
        }),
      );
    });
    expect(wrong).toEqual([]);
    expect(cells).toBe(countOf(new Set(['text']), 'cells'));
  });

  it('calls what the cell’s own list style says explicit, and what the style gives inherited', () => {
    const listOnly = main.ladder.flatMap((rung, c) =>
      rung.levels.length === 1 && rung.levels[0] === 'list' ? [c] : [],
    );
    const bare = main.ladder.findIndex((rung) => rung.levels.length === 0 && rung.lvl === 0);
    expect(listOnly).toHaveLength(2);
    let checked = 0;
    forEachSlide(new Set(['text']), (deck, slide, index) => {
      if (index !== 0) return;
      const table = tableOf(slide);
      const style = styleOf(table);
      const { sheet, defaultTextStyle } = contextOf(deck, slide);
      const at = (c: number): ReturnType<typeof cellText> =>
        cellText(table, style, sheet, defaultTextStyle, 1, c);
      for (const c of listOnly) {
        const { context, paragraph, run } = at(c);
        const size = resolveSize(context, paragraph, run);
        expect([size.value, size.origin, size.explicit], deck.id).toEqual([
          main.levels.list.sz,
          'cell',
          true,
        ]);
        checked += 1;
      }
      const { context, paragraph, run } = at(bare);
      const fill = resolveRun(context, paragraph, run, (p) => p.fill);
      expect([fill?.origin, fill?.explicit], deck.id).toEqual(['tableStyle', false]);
    });
    expect(checked).toBe(2 * DECKS.filter((deck) => deck.group === 'text').length);
  });
});

/* -------------------------------------------------------------------------- */
/* the rivals                                                                 */
/* -------------------------------------------------------------------------- */

interface Finding {
  readonly question: string;
  readonly winner: readonly string[];
  readonly candidates: readonly {
    readonly name: string;
    readonly fits: number;
    readonly of: number;
  }[];
}

/* -------------------------------------------------------------------------- */
/* dashes, diagonals and the background's effect                              */
/* -------------------------------------------------------------------------- */

const { origin: ORIGIN, cell_pt: CELL } = main.encoding;
const PAGE = 'FFFFFF';
const px = (pt: number): number => Math.round(pt * SCALE);

/** A fixture key's deck and slide: `deck#slide`. */
function slideAt(key: string): { deck: TextDeck; slide: SlideRecord } {
  const [id = '', index = ''] = key.split('#');
  const deck = DECKS.find((d) => d.id === id);
  const slide = deck?.slides[Number(index)];
  if (deck === undefined || slide === undefined) throw new Error(`no slide ${key}`);
  return { deck, slide };
}

/** Pixels from their run-length form, `RRGGBB*n`. */
const unrun = (runs: string): string[] =>
  runs.split(',').flatMap((run) => {
    const [hex = '', n = '0'] = run.split('*');
    return Array.from({ length: Number(n) }, () => hex);
  });

/** What an edge's centre line shows: `line` and the paints in turn, `line` alone, or no line. */
function dashClass(
  pixels: readonly string[],
  line: string | null,
  paints: readonly string[],
): string {
  const lit = pixels.filter((p) => p === line).length;
  const bare = pixels.filter((p) => paints.includes(p)).length;
  if (lit + bare < pixels.length * 0.8) return 'mixed';
  if (lit === 0) return 'nothing';
  return bare === 0 ? 'solid' : 'dashed';
}

interface Stroke {
  /** Left, top, right, bottom, in pixels. */
  readonly rect: readonly [number, number, number, number];
  readonly down: boolean;
  readonly rgb: string;
  /** Half the line's width, in pixels. */
  readonly half: number;
}

/**
 * The diagonals the model gives every grid position, in the export's pixels, each drawn across its
 * owner's span; mirrored with an rtl table.
 */
function diagonalStrokes(slide: SlideRecord, table: Table, palette: Palette): Stroke[] {
  const grid = tableGrid(table);
  const seen = new Set<string>();
  return grid.positions.flat().flatMap((anchor, k): Stroke[] => {
    const [row, col] = [Math.floor(k / grid.cols), k % grid.cols];
    const { down, up } = tableCellDiagonals(table, grid, row, col);
    const left = slide.rtl === true ? slide.cols - anchor.col - anchor.cols : anchor.col;
    const rect = [
      px(ORIGIN.x + left * CELL.w),
      px(ORIGIN.y + anchor.row * CELL.h),
      px(ORIGIN.x + (left + anchor.cols) * CELL.w),
      px(ORIGIN.y + (anchor.row + anchor.rows) * CELL.h),
    ] as const;
    const lines: readonly (readonly [Line | null, boolean])[] = [
      [down, true],
      [up, false],
    ];
    return lines.flatMap(([line, isDown]) => {
      const look =
        line === null
          ? null
          : drawn({ value: { kind: 'value', value: line }, source: 'tcPr' }, palette);
      const key = `${rect.join(',')} ${String(isDown)} ${look?.rgb ?? ''}`;
      if (look === null || seen.has(key)) return [];
      seen.add(key);
      return [{ rect, down: isDown, rgb: look.rgb, half: (look.weight * SCALE) / 2 }];
    });
  });
}

/** Whether a diagonal crosses the pixels a grid edge's profile reads, which then read the diagonal. */
function crossesProfile(strokes: readonly Stroke[], edge: TableEdge): boolean {
  const [x, y] =
    edge.axis === 'h'
      ? [px(ORIGIN.x + edge.col * CELL.w + 36), px(ORIGIN.y + edge.row * CELL.h)]
      : [px(ORIGIN.x + edge.col * CELL.w), px(ORIGIN.y + edge.row * CELL.h + 30)];
  return Array.from({ length: 2 * REACH }, (_, k) => k - REACH).some((k) =>
    strokes.some((s) =>
      edge.axis === 'h'
        ? distance(x, y + k, s) < s.half + 2.5
        : distance(x + k, y, s) < s.half + 2.5,
    ),
  );
}

/** How far a pixel's centre lies from a diagonal's centre line. */
function distance(x: number, y: number, s: Stroke): number {
  const [x0, y0, x1, y1] = s.rect;
  const [ax, ay, bx, by] = s.down ? [x0, y0, x1, y1] : [x0, y1, x1, y0];
  const [dx, dy] = [bx - ax, by - ay];
  const t = Math.max(
    0,
    Math.min(1, ((x + 0.5 - ax) * dx + (y + 0.5 - ay) * dy) / (dx * dx + dy * dy)),
  );
  return Math.hypot(x + 0.5 - (ax + t * dx), y + 0.5 - (ay + t * dy));
}

describe('dashes, diagonals and the table background’s effect', () => {
  it('dashes the dashed line an edge’s owner writes, and draws none it writes without a fill', () => {
    const rows = main.dashes as readonly { key: string; runs: string; observed: string }[];
    const seen = new Set<string>();
    for (const row of rows) {
      const [where = '', at = ''] = row.key.split('@');
      const { deck, slide } = slideAt(where);
      const palette = PALETTES[deck.theme];
      const table = tableOf(slide);
      const [r = 0, c = 0] = at.slice(1).split(',').map(Number);
      const edge: TableEdge = { axis: at.startsWith('h') ? 'h' : 'v', row: r, col: c };
      const sourced = tableEdgeLine(table, styleOf(table), tableGrid(table), edge);
      const colour = drawn(sourced, palette)?.rgb ?? null;
      const dash =
        sourced === null
          ? null
          : themedLine(sourced.value, palette.theme, palette.colors).line?.dash;
      let predicted = 'dashed';
      if (colour === null) predicted = 'nothing';
      else if (
        dash === null ||
        dash === undefined ||
        (dash.kind === 'preset' && dash.val === 'solid')
      )
        predicted = 'solid';
      const k = edgesOf(slide.rows, slide.cols).findIndex(
        (e) => e.axis === edge.axis && e.row === edge.row && e.col === edge.col,
      );
      const [a = '', b = ''] = (decode(deck, slide).edges[k] ?? '').split('|');
      expect([dashClass(unrun(row.runs), colour, [a, b]), predicted], row.key).toEqual([
        row.observed,
        row.observed,
      ]);
      seen.add(row.observed);
    }
    expect([...seen].sort()).toEqual(['dashed', 'nothing']);
  });

  it('draws every diagonal PowerPoint drew, at every point clear of the rest', () => {
    const records = main.diagonals as readonly {
      key: string;
      colours: readonly string[];
      points: string;
    }[];
    expect(new Set(records.map((r) => slideAt(r.key).deck.group))).toEqual(
      new Set(['direct', 'merge', 'rtl']),
    );
    let points = 0;
    const wrong: string[] = [];
    for (const record of records) {
      const { deck, slide } = slideAt(record.key);
      const palette = PALETTES[deck.theme];
      const table = tableOf(slide);
      const { cells } = decode(deck, slide);
      const strokes = diagonalStrokes(slide, table, palette);
      for (const point of record.points.split(' ')) {
        const [x = 0, y = 0, k = 0] = point.split(',').map(Number);
        const seen = record.colours[k];
        const r = Math.floor((y - px(ORIGIN.y)) / px(CELL.h));
        const v = Math.floor((x - px(ORIGIN.x)) / px(CELL.w));
        const c = slide.rtl === true ? slide.cols - 1 - v : v;
        const on = strokes.findLast((s) => distance(x, y, s) <= 1.5);
        const predicted = on?.rgb ?? cells[r]?.[c]?.paint;
        points += 1;
        if (predicted !== seen)
          wrong.push(`${record.key} ${point}: ${String(predicted)} for ${String(seen)}`);
      }
    }
    expect(wrong.slice(0, 5)).toEqual([]);
    expect(points).toBeGreaterThan(300);
  });

  it('casts the shadow of the effect a:tblPr or else tblBg gives, from the grid and not the frame', () => {
    const samples = main.background as readonly {
      deck: string;
      slide: number;
      right: string;
      below: string;
      beyondGrid: string;
    }[];
    for (const sample of samples) {
      const { deck, slide } = slideAt(`${sample.deck}#${String(sample.slide)}`);
      const palette = PALETTES[deck.theme];
      const table = tableOf(slide);
      const { effects } = tableBackground(table, styleOf(table));
      const shadows =
        effects === null
          ? []
          : themedEffects(effects.value, palette.theme, palette.colors).effects.filter(
              (e) => e.kind === 'outerShdw',
            );
      const [w, h] = [slide.cols * CELL.w, slide.rows * CELL.h];
      const middle = ORIGIN.y + 2 * CELL.h + 30;
      const places: readonly (readonly ['right' | 'below' | 'beyondGrid', number, number])[] = [
        ['right', ORIGIN.x + w + 3, middle],
        ['below', ORIGIN.x + 36, ORIGIN.y + h + 3],
        ['beyondGrid', ORIGIN.x + w + 20, middle],
      ];
      for (const [name, x, y] of places) {
        const hit = shadows.findLast((s) => {
          const angle = (s.dir / 60000 / 180) * Math.PI;
          const [dx, dy] = [(s.dist / EMU) * Math.cos(angle), (s.dist / EMU) * Math.sin(angle)];
          return (
            x >= ORIGIN.x + dx &&
            x < ORIGIN.x + dx + w &&
            y >= ORIGIN.y + dy &&
            y < ORIGIN.y + dy + h
          );
        });
        const predicted = hit === undefined ? PAGE : hexOf(resolveColor(hit.color, palette.colors));
        expect(predicted, `${sample.deck}#${String(sample.slide)} ${name}`).toBe(sample[name]);
      }
    }
    expect(samples.some((s) => s.right !== PAGE || s.below !== PAGE)).toBe(true);
  });
});

describe('a position the table does not have', () => {
  const codeOf = (resolve: () => unknown): string => {
    try {
      resolve();
    } catch (error) {
      return isModelError(error) ? error.code : 'not a ModelError';
    }
    return 'nothing thrown';
  };

  it('throws MODEL_TABLE_POSITION from every resolver, styled or on the default grid', () => {
    const slide = DECKS.find((deck) => deck.id === 'control')?.slides[0];
    if (slide === undefined) throw new Error('no control slide');
    const table = tableOf(slide);
    const style = styleOf(table);
    const grid = tableGrid(table);
    const { rows, cols } = grid;
    for (const [row, col] of [
      [rows, 0],
      [0, cols],
      [-1, 0],
      [0.5, 0],
    ] as const) {
      const where = `${String(row)},${String(col)}`;
      expect(
        codeOf(() => tablePartsAt(table.props, rows, cols, row, col)),
        where,
      ).toBe('MODEL_TABLE_POSITION');
      expect(
        codeOf(() => tableCellFill(table, style, grid, row, col)),
        where,
      ).toBe('MODEL_TABLE_POSITION');
      for (const layer of [style, null]) {
        expect(
          codeOf(() => tableTextLayer(table, layer, grid, row, col)),
          where,
        ).toBe('MODEL_TABLE_POSITION');
      }
    }
    expect(
      codeOf(() => tableEdgeLine(table, style, grid, { axis: 'h', row: rows + 1, col: 0 })),
    ).toBe('MODEL_TABLE_POSITION');
    expect(
      codeOf(() => tableEdgeLine(table, style, grid, { axis: 'v', row: 0, col: cols + 1 })),
    ).toBe('MODEL_TABLE_POSITION');
  });
});

describe('the readings PowerPoint refuted', () => {
  it('scored every rival below its rows, and every winner at all of them', () => {
    const findings = main.findings as unknown as readonly Finding[];
    expect(findings.length).toBeGreaterThan(20);
    for (const f of findings) {
      for (const c of f.candidates) {
        if (f.winner.includes(c.name)) expect(c.fits, `${f.question}: ${c.name}`).toBe(c.of);
        else expect(c.fits, `${f.question}: ${c.name}`).toBeLessThan(c.of);
      }
    }
  });

  /** How many grid edges of the styled sweep a rival line rule draws differently from PowerPoint. */
  function edgeMisses(
    rival: (
      table: Table,
      style: NonNullable<ReturnType<typeof tableStyleOf>>,
      grid: ReturnType<typeof tableGrid>,
      edge: TableEdge,
    ) => ReturnType<typeof tableEdgeLine>,
  ): number {
    let misses = 0;
    forEachSlide(new Set(['sweep']), (deck, slide) => {
      const table = tableOf(slide);
      const style = styleOf(table);
      if (style === null) return;
      const grid = tableGrid(table);
      const { edges } = decode(deck, slide);
      edgesOf(slide.rows, slide.cols).forEach((edge, k) => {
        const [a = '', b = '', ...rest] = (edges[k] ?? '').split('|');
        if (
          describeLine(
            drawn(rival(table, style, grid, edge), PALETTES[deck.theme]),
            a,
            b,
            rest.join('|'),
          ) !== rest.join('|')
        )
          misses += 1;
      });
    });
    return misses;
  }

  it('refutes resolving an edge from one cell’s own parts, the reading a per-cell renderer makes', () => {
    const lowerCellOnly = (
      table: Table,
      style: NonNullable<ReturnType<typeof tableStyleOf>>,
      grid: ReturnType<typeof tableGrid>,
      edge: TableEdge,
    ): ReturnType<typeof tableEdgeLine> => {
      const inside =
        edge.axis === 'h'
          ? edge.row > 0 && edge.row < grid.rows
          : edge.col > 0 && edge.col < grid.cols;
      if (!inside) return tableEdgeLine(table, style, grid, edge);
      const own = tablePartsAt(table.props, grid.rows, grid.cols, edge.row, edge.col);
      const other =
        edge.axis === 'h'
          ? tablePartsAt(table.props, grid.rows, grid.cols, edge.row - 1, edge.col)
          : tablePartsAt(table.props, grid.rows, grid.cols, edge.row, edge.col - 1);
      for (const part of own) {
        const role = other.includes(part)
          ? edge.axis === 'h'
            ? 'insideH'
            : 'insideV'
          : edge.axis === 'h'
            ? 'top'
            : 'left';
        const stated = style.parts[part]?.cell?.borders?.[role];
        if (stated !== undefined) return { value: stated, source: part };
      }
      return null;
    };
    // The model misses none of these edges: 'draws every grid edge PowerPoint drew' holds it to all.
    expect(edgeMisses(lowerCellOnly)).toBeGreaterThan(1000);
  });

  it('refutes composing the parts in any order but the one measured', () => {
    const reversed = [...TABLE_PART_ORDER].reverse();
    let misses = 0;
    forEachSlide(new Set(['sweep']), (deck, slide) => {
      const table = tableOf(slide);
      const style = styleOf(table);
      if (style === null) return;
      const grid = tableGrid(table);
      decode(deck, slide).cells.forEach((row, r) =>
        row.forEach((cell, c) => {
          const parts = [...tablePartsAt(table.props, grid.rows, grid.cols, r, c)].sort(
            // Highest first under the reversed order: whole table first, corners last.
            (p, q) => reversed.indexOf(q) - reversed.indexOf(p),
          );
          const found = parts.map((p) => style.parts[p]?.cell?.fill).find((f) => f !== undefined);
          const got = fillText(
            found === undefined ? null : { value: found, source: 'wholeTbl' },
            PALETTES[deck.theme],
          );
          if (got !== cell.fill) misses += 1;
        }),
      );
    });
    expect(misses).toBeGreaterThan(1000);
  });

  it('misses every fill the fixture says each rival order misses, and no other', () => {
    // The three orders C9 scored against the schema's, lowest first.
    const rivals: Readonly<Record<string, readonly TableStylePartName[]>> = {
      spreadsheetml: [
        'wholeTbl',
        'band1V',
        'band2V',
        'band1H',
        'band2H',
        'lastCol',
        'firstCol',
        'firstRow',
        'lastRow',
        'nwCell',
        'neCell',
        'swCell',
        'seCell',
      ],
      colsOverRows: [
        'wholeTbl',
        'band1H',
        'band2H',
        'band1V',
        'band2V',
        'firstRow',
        'lastRow',
        'firstCol',
        'lastCol',
        'nwCell',
        'neCell',
        'swCell',
        'seCell',
      ],
      bandsOverEnds: [
        'wholeTbl',
        'lastCol',
        'firstCol',
        'lastRow',
        'firstRow',
        'band1H',
        'band2H',
        'band1V',
        'band2V',
        'nwCell',
        'neCell',
        'swCell',
        'seCell',
      ],
    };
    const finding = (main.findings as unknown as readonly Finding[]).find(
      (f) => f.question === 'fill: order',
    );
    expect(finding?.winner).toEqual(['schema']);
    for (const [name, order] of Object.entries(rivals)) {
      let misses = 0;
      forEachSlide(new Set(['control', 'sweep', 'ends', 'sides']), (deck, slide) => {
        const table = tableOf(slide);
        const style = styleOf(table);
        if (style === null) return;
        const grid = tableGrid(table);
        decode(deck, slide).cells.forEach((row, r) =>
          row.forEach((cell, c) => {
            const parts = [...tablePartsAt(table.props, grid.rows, grid.cols, r, c)].sort(
              (p, q) => order.indexOf(q) - order.indexOf(p),
            );
            const part = parts.find((p) => style.parts[p]?.cell?.fill !== undefined);
            const fill = part === undefined ? undefined : style.parts[part]?.cell?.fill;
            const got = fillText(
              part === undefined || fill === undefined ? null : { value: fill, source: part },
              PALETTES[deck.theme],
            );
            if (got !== cell.fill) misses += 1;
          }),
        );
      });
      const recorded = finding?.candidates.find((candidate) => candidate.name === name);
      expect(misses, name).toBe((recorded?.of ?? -1) - (recorded?.fits ?? 0));
      expect(misses, name).toBeGreaterThan(0);
    }
  });

  it('refutes painting every translucent part over the last, where two bands cross', () => {
    let misses = 0;
    forEachSlide(new Set(['sweep']), (deck, slide) => {
      const table = tableOf(slide);
      const style = styleOf(table);
      if (style === null) return;
      const grid = tableGrid(table);
      const palette = PALETTES[deck.theme];
      const background = fillText(tableBackground(table, style).fill, palette);
      decode(deck, slide).cells.forEach((row, r) =>
        row.forEach((cell, c) => {
          const stack = [...tablePartsAt(table.props, grid.rows, grid.cols, r, c)]
            .reverse()
            .map((p) => style.parts[p]?.cell?.fill)
            .filter((f) => f !== undefined)
            .map((f) => fillText({ value: f, source: 'wholeTbl' }, palette));
          const stacked = stack.reduce(
            (under, fill) => over(under, fill),
            over([255, 255, 255], background),
          );
          if (hexRgb(stacked) !== cell.paint) misses += 1;
        }),
      );
    });
    expect(misses).toBeGreaterThan(100);
  });

  it('refutes drawing a lower or right cell’s own top or left border over its owner’s line', () => {
    let misses = 0;
    forEachSlide(new Set(['direct']), (deck, slide) => {
      const table = tableOf(slide);
      const style = styleOf(table);
      const grid = tableGrid(table);
      const { edges } = decode(deck, slide);
      edgesOf(slide.rows, slide.cols).forEach((edge, k) => {
        const after =
          edge.row < grid.rows && edge.col < grid.cols
            ? grid.positions[edge.row]?.[edge.col]
            : undefined;
        const own = after?.cell?.props?.borders[edge.axis === 'h' ? 'top' : 'left'];
        if (own === undefined || (edge.axis === 'h' ? edge.row : edge.col) === 0) return;
        const [a = '', b = '', ...rest] = (edges[k] ?? '').split('|');
        const later = describeLine(
          drawn({ value: { kind: 'value', value: own }, source: 'tcPr' }, PALETTES[deck.theme]),
          a,
          b,
          rest.join('|'),
        );
        const owner = describeLine(
          drawn(tableEdgeLine(table, style, grid, edge), PALETTES[deck.theme]),
          a,
          b,
          rest.join('|'),
        );
        expect(owner).toBe(rest.join('|'));
        if (later !== rest.join('|')) misses += 1;
      });
    });
    expect(misses).toBeGreaterThan(5);
  });
});
