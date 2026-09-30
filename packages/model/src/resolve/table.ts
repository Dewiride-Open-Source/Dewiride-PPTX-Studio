/**
 * Which style PowerPoint draws a table with (C8, ADR 0063), and what that style, the cells' own
 * `a:tcPr` and the table's `a:tblPr` give each cell and grid edge (C9, ADR 0064).
 */

import {
  resolveColor,
  type ColorContext,
  type Fill,
  type Line,
  type Rgba,
} from '@pptx-studio/paint';
import { firstChild, parseXmlString } from '@pptx-studio/xml';

import { BUILTIN_TABLE_STYLES } from '../builtin/table-styles.js';
import { ModelError } from '../errors.js';
import { parseTableStyle } from '../parse/table.js';
import { styleMatrixFill, styleMatrixLine } from '../style.js';
import type {
  GridCell,
  Table,
  TableEdge,
  TableGrid,
  TableProps,
  TableSourced,
  TableStyle,
  TableStyleBorders,
  TableStylePartName,
  Themeable,
} from '../table.js';
import type { RunProps, Typeface } from '../text.js';
import type { Theme } from '../types.js';

const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
/** The part name a built-in's errors carry: it comes from no package. */
const BUILTIN_PART = 'builtin table styles';
const BY_ID = new Map(BUILTIN_TABLE_STYLES.map((style) => [style.id, style]));

/** One of PowerPoint's 74 built-ins by GUID, in any case, parsed afresh; `undefined` for any other. */
export function builtinTableStyle(id: string): TableStyle | undefined {
  const style = BY_ID.get(id.toUpperCase());
  if (style === undefined) return undefined;
  const list = parseXmlString(`<a:tblStyleLst xmlns:a="${NS_A}">${style.xml}</a:tblStyleLst>`);
  const element = firstChild(list.root, 'a:tblStyle');
  if (element === undefined) {
    throw new ModelError('MODEL_TABLE_STYLE', `${style.name} is no a:tblStyle`, BUILTIN_PART);
  }
  return parseTableStyle(element, BUILTIN_PART);
}

/**
 * The built-in style a table is drawn with, or `null` for PowerPoint's default: a 1-pt black
 * grid and no fill, whatever the theme and whatever the package defines.
 */
export function tableStyleOf(table: Table): TableStyle | null {
  const ref = table.props?.style;
  return ref === undefined ? null : (builtinTableStyle(ref.id) ?? null);
}

/* -------------------------------------------------------------------------- */
/* which parts reach a position                                               */
/* -------------------------------------------------------------------------- */

/** The order the parts compose in, lowest first: C9 on every drawable table (ADR 0064). */
export const TABLE_PART_ORDER = [
  'wholeTbl',
  'band1H',
  'band2H',
  'band1V',
  'band2V',
  'lastCol',
  'firstCol',
  'lastRow',
  'seCell',
  'swCell',
  'firstRow',
  'neCell',
  'nwCell',
] as const satisfies readonly TableStylePartName[];

const RANK: ReadonlyMap<TableStylePartName, number> = new Map(
  TABLE_PART_ORDER.map((p, i) => [p, i]),
);

/**
 * The parts reaching grid position (row, col), highest first: ECMA-376's reading, which C9 found
 * alone fits (ADR 0064). Bands skip the rows and columns their end flags take, and a corner part
 * needs both of its flags.
 */
export function tablePartsAt(
  props: TableProps | undefined,
  rows: number,
  cols: number,
  row: number,
  col: number,
): readonly TableStylePartName[] {
  const within = (n: number, size: number): boolean => Number.isInteger(n) && n >= 0 && n < size;
  if (!within(row, rows) || !within(col, cols)) {
    throw new ModelError(
      'MODEL_TABLE_POSITION',
      `no grid position ${String(row)},${String(col)} in ${String(rows)}x${String(cols)}`,
    );
  }
  const parts: TableStylePartName[] = ['wholeTbl'];
  const firstRow = props?.firstRow === true && row === 0;
  const lastRow = props?.lastRow === true && row === rows - 1;
  const firstCol = props?.firstCol === true && col === 0;
  const lastCol = props?.lastCol === true && col === cols - 1;
  if (props?.bandRow === true && !firstRow && !lastRow) {
    parts.push((row - (props.firstRow ? 1 : 0)) % 2 === 0 ? 'band1H' : 'band2H');
  }
  if (props?.bandCol === true && !firstCol && !lastCol) {
    parts.push((col - (props.firstCol ? 1 : 0)) % 2 === 0 ? 'band1V' : 'band2V');
  }
  if (firstRow) parts.push('firstRow');
  if (lastRow) parts.push('lastRow');
  if (firstCol) parts.push('firstCol');
  if (lastCol) parts.push('lastCol');
  if (firstRow && firstCol) parts.push('nwCell');
  if (firstRow && lastCol) parts.push('neCell');
  if (lastRow && firstCol) parts.push('swCell');
  if (lastRow && lastCol) parts.push('seCell');
  return parts.sort((a, b) => (RANK.get(b) ?? 0) - (RANK.get(a) ?? 0));
}

/** The grid cell at (row, col), or `null` beyond the table. */
function cellAt(grid: TableGrid, row: number, col: number): GridCell | null {
  return grid.positions[row]?.[col] ?? null;
}

/** The grid cell at (row, col); a position beyond the table throws. */
function cellOf(grid: TableGrid, row: number, col: number): GridCell {
  const cell = cellAt(grid, row, col);
  if (cell === null) {
    throw new ModelError('MODEL_TABLE_POSITION', `no grid position ${String(row)},${String(col)}`);
  }
  return cell;
}

/** A merged cell's parts, highest first: its anchor's, and every covered position's ends and corners (C9). */
function partsOf(table: Table, grid: TableGrid, cell: GridCell): readonly TableStylePartName[] {
  const pooled = new Set(tablePartsAt(table.props, grid.rows, grid.cols, cell.row, cell.col));
  for (let row = cell.row; row < cell.row + cell.rows; row++) {
    for (let col = cell.col; col < cell.col + cell.cols; col++) {
      for (const part of tablePartsAt(table.props, grid.rows, grid.cols, row, col))
        if (!BANDS.has(part)) pooled.add(part);
    }
  }
  return [...pooled].sort((a, b) => (RANK.get(b) ?? 0) - (RANK.get(a) ?? 0));
}

const BANDS: ReadonlySet<TableStylePartName> = new Set(['band1H', 'band2H', 'band1V', 'band2V']);

/** The edge an end part offers any cell outside it, whichever way the edge runs (C9, merged cells). */
const KIND_EDGE: Readonly<Partial<Record<TableStylePartName, Edge>>> = {
  firstRow: 'bottom',
  lastRow: 'top',
  firstCol: 'right',
  lastCol: 'left',
};

/* -------------------------------------------------------------------------- */
/* fills                                                                      */
/* -------------------------------------------------------------------------- */

/**
 * The fill of the cell at grid position (row, col): its own `a:tcPr` fill, else the highest part
 * that states one. Parts compose per property (C9); `null` paints nothing, and one translucent fill
 * is painted alone, never over another part's.
 */
export function tableCellFill(
  table: Table,
  style: TableStyle | null,
  grid: TableGrid,
  row: number,
  col: number,
): TableSourced<Themeable<Fill>> | null {
  const cell = cellOf(grid, row, col);
  const own = cell.cell?.props?.fill;
  if (own !== undefined) return { value: { kind: 'value', value: own }, source: 'tcPr' };
  if (style === null) return null;
  for (const part of partsOf(table, grid, cell)) {
    const fill = style.parts[part]?.cell?.fill;
    if (fill !== undefined) return { value: fill, source: part };
  }
  return null;
}

/** What paints under the cells: `a:tblPr`'s own fill in place of the style's `tblBg` (C9). */
export function tableBackground(
  table: Table,
  style: TableStyle | null,
): TableSourced<Themeable<Fill>> | null {
  const own = table.props?.fill;
  if (own !== undefined) return { value: { kind: 'value', value: own }, source: 'tblPr' };
  const background = style?.background?.fill;
  return background === undefined ? null : { value: background, source: 'tblBg' };
}

/* -------------------------------------------------------------------------- */
/* grid edges                                                                 */
/* -------------------------------------------------------------------------- */

/** PowerPoint's default grid: a 1-pt black line on every edge of a table that names no built-in (C8). */
export const DEFAULT_GRID_LINE: Line = {
  w: 12700,
  cap: null,
  cmpd: null,
  algn: null,
  fill: { type: 'solid', color: { space: 'srgb', hex: '000000', transforms: [] } },
  dash: null,
  join: null,
  headEnd: null,
  tailEnd: null,
};

type Edge = keyof Omit<TableStyleBorders, 'tlToBr' | 'blToTr'>;

/** The side of the cell before (above, left of) and after an edge that faces it. */
const FACING = {
  h: { before: 'bottom', after: 'top', inside: 'insideH' },
  v: { before: 'right', after: 'left', inside: 'insideV' },
} as const;

/**
 * The line a grid edge draws: its owner's `a:tcPr` line, else the highest part of either cell that
 * states it, a part reaching both offering its inside edge (C9).
 * The owner is the cell before (above, left) where its anchor is level with the edge, else the cell
 * after where its anchor is, else the cell before; `null` inside a merged cell.
 */
export function tableEdgeLine(
  table: Table,
  style: TableStyle | null,
  grid: TableGrid,
  edge: TableEdge,
): TableSourced<Themeable<Line>> | null {
  const facing = FACING[edge.axis];
  const before =
    edge.axis === 'h' ? cellAt(grid, edge.row - 1, edge.col) : cellAt(grid, edge.row, edge.col - 1);
  const after =
    edge.row < grid.rows && edge.col < grid.cols ? cellAt(grid, edge.row, edge.col) : null;
  if (before === null && after === null) {
    throw new ModelError(
      'MODEL_TABLE_POSITION',
      `no grid edge ${edge.axis}${String(edge.row)},${String(edge.col)}`,
    );
  }
  if (before !== null && before === after) return null;
  const level = (cell: GridCell | null): boolean =>
    cell !== null && (edge.axis === 'h' ? edge.col === cell.col : edge.row === cell.row);
  const beforeOwns = before !== null && (level(before) || !level(after));
  const direct = beforeOwns
    ? before.cell?.props?.borders[facing.before]
    : after?.cell?.props?.borders[facing.after];
  if (direct !== undefined) return { value: { kind: 'value', value: direct }, source: 'tcPr' };
  if (style === null) return { value: { kind: 'value', value: DEFAULT_GRID_LINE }, source: 'grid' };

  const partsBefore = before === null ? [] : partsOf(table, grid, before);
  const partsAfter = after === null ? [] : partsOf(table, grid, after);
  const all = [...new Set([...partsBefore, ...partsAfter])].sort(
    (a, b) => (RANK.get(b) ?? 0) - (RANK.get(a) ?? 0),
  );
  for (const part of all) {
    const inBefore = partsBefore.includes(part);
    const inAfter = partsAfter.includes(part);
    const geometric: Edge =
      inBefore && inAfter ? facing.inside : inBefore ? facing.before : facing.after;
    const kind =
      before !== null && after !== null && inBefore !== inAfter ? KIND_EDGE[part] : undefined;
    const role = kind ?? geometric;
    const stated = style.parts[part]?.cell?.borders?.[role];
    if (stated !== undefined) return { value: stated, source: part };
  }
  return null;
}

/* -------------------------------------------------------------------------- */
/* text                                                                       */
/* -------------------------------------------------------------------------- */

const themeFace = (typeface: string): Typeface => ({
  typeface,
  panose: undefined,
  pitchFamily: undefined,
  charset: undefined,
});

function runProps(node: Table['node'], fields: Partial<RunProps>): RunProps {
  return {
    sz: undefined,
    b: undefined,
    i: undefined,
    u: undefined,
    strike: undefined,
    cap: undefined,
    spc: undefined,
    kern: undefined,
    baseline: undefined,
    noProof: undefined,
    lang: undefined,
    altLang: undefined,
    latin: undefined,
    ea: undefined,
    cs: undefined,
    sym: undefined,
    fill: undefined,
    line: undefined,
    effects: undefined,
    highlight: undefined,
    node,
    ...fields,
  };
}

/**
 * What the style gives the text of the cell at (row, col): bold, italic, face and colour, each from
 * the highest part that states it (C9), as run properties for the text cascade.
 * The default grid gives `tx1` in the minor face whatever the flags; `node` is the style's element,
 * or the table's for the grid.
 */
export function tableTextLayer(
  table: Table,
  style: TableStyle | null,
  grid: TableGrid,
  row: number,
  col: number,
): RunProps {
  const cell = cellOf(grid, row, col);
  const minor = { latin: themeFace('+mn-lt'), ea: themeFace('+mn-ea'), cs: themeFace('+mn-cs') };
  if (style === null) {
    return runProps(table.node, {
      ...minor,
      fill: { type: 'solid', color: { space: 'scheme', name: 'tx1', transforms: [] } },
    });
  }
  const texts = partsOf(table, grid, cell)
    .map((part) => style.parts[part]?.text)
    .filter((t) => t !== undefined);
  const b = texts.find((t) => t.b !== undefined)?.b;
  const i = texts.find((t) => t.i !== undefined)?.i;
  const color = texts.find((t) => t.color !== undefined)?.color;
  const font = texts.find((t) => t.font !== undefined)?.font;
  const onOff = (v: typeof b): boolean | undefined =>
    v === 'on' ? true : v === 'off' ? false : undefined;
  let faces: Partial<Pick<RunProps, 'latin' | 'ea' | 'cs'>> = {};
  if (font?.kind === 'ref' && font.ref.idx !== 'none') {
    const key = font.ref.idx === 'major' ? 'mj' : 'mn';
    faces = {
      latin: themeFace(`+${key}-lt`),
      ea: themeFace(`+${key}-ea`),
      cs: themeFace(`+${key}-cs`),
    };
  } else if (font?.kind === 'value') {
    const face = (name: string | null): Typeface | undefined =>
      name === null ? undefined : themeFace(name);
    faces = { latin: face(font.value.latin), ea: face(font.value.ea), cs: face(font.value.cs) };
  }
  return runProps(style.node, {
    b: onOff(b),
    i: onOff(i),
    ...faces,
    ...(color === undefined ? {} : { fill: { type: 'solid' as const, color } }),
  });
}

/* -------------------------------------------------------------------------- */
/* the theme                                                                  */
/* -------------------------------------------------------------------------- */

/** A style fill with a `fillRef` followed into the theme, and the colour `phClr` takes in it. */
export function themedFill(
  fill: Themeable<Fill>,
  theme: Theme,
  colors: ColorContext,
): { readonly fill: Fill | null; readonly phClr: Rgba | null } {
  if (fill.kind === 'value') return { fill: fill.value, phClr: null };
  const found = styleMatrixFill(fill.ref, theme);
  return {
    fill: found,
    phClr: fill.ref.color === null ? null : resolveColor(fill.ref.color, colors),
  };
}

/** A style line with an `lnRef` followed into the theme, and the colour `phClr` takes in it. */
export function themedLine(
  line: Themeable<Line>,
  theme: Theme,
  colors: ColorContext,
): { readonly line: Line | null; readonly phClr: Rgba | null } {
  if (line.kind === 'value') return { line: line.value, phClr: null };
  const found = styleMatrixLine(line.ref, theme);
  return {
    line: found,
    phClr: line.ref.color === null ? null : resolveColor(line.ref.color, colors),
  };
}
