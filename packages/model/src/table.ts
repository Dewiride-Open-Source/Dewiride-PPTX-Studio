/**
 * A `p:graphicFrame`'s `a:tbl` as its part wrote it, the grid PowerPoint reads from it, and the
 * shape of a table style. Measured in C7 and C8; ADRs 0056 and 0063 have the evidence.
 */

import type { Color, Effect, Fill, Line } from '@pptx-studio/paint';
import type { XElement } from '@pptx-studio/xml';

import type { HorzOverflow, TextAnchor, TextBody, VerticalText } from './text.js';
import type { FontCollection, FontRef, StyleRef } from './types.js';

/** `a:graphicData/@uri` of a table. */
export const TABLE_URI = 'http://schemas.openxmlformats.org/drawingml/2006/table';

/** The column PowerPoint invents for a grid that declares none, empty: two margins and 2 pt. */
export const INVENTED_COLUMN_WIDTH = 208280;

/** `a:tblPr`. The seven flags are the file's own and inherit from nothing, so absent is false. */
export interface TableProps {
  readonly rtl: boolean;
  readonly firstRow: boolean;
  readonly firstCol: boolean;
  readonly lastRow: boolean;
  readonly lastCol: boolean;
  readonly bandRow: boolean;
  readonly bandCol: boolean;
  readonly fill: Fill | undefined;
  readonly effects: readonly Effect[] | undefined;
  readonly style: TableStyleRef | undefined;
  readonly node: XElement;
}

/** `a:tableStyleId`, or an inline `a:tableStyle`, which PowerPoint resolves by its `@styleId` alone. */
export type TableStyleRef =
  | { readonly kind: 'id'; readonly id: string }
  | { readonly kind: 'inline'; readonly id: string; readonly node: XElement };

/** `a:gridCol`. */
export interface TableColumn {
  /** `@w` in EMU. */
  readonly w: number;
  readonly node: XElement;
}

/** `a:tr`. */
export interface TableRow {
  /** `@h` in EMU: the least the row is drawn at, never the most. */
  readonly h: number;
  readonly cells: readonly TableCell[];
  readonly node: XElement;
}

/** `a:tc`: an absent span reads as 1 and an absent flag as false; `tableGrid` says what they mean. */
export interface TableCell {
  readonly gridSpan: number;
  readonly rowSpan: number;
  readonly hMerge: boolean;
  readonly vMerge: boolean;
  /** `@id`, which `a:headers` in other cells name. */
  readonly id: string | null;
  readonly text: TextBody | undefined;
  readonly props: TableCellProps | undefined;
  readonly node: XElement;
}

/** The six `a:tcPr` borders, each an `a:ln` under another name. */
export interface TableCellBorders {
  readonly left: Line | undefined;
  readonly right: Line | undefined;
  readonly top: Line | undefined;
  readonly bottom: Line | undefined;
  readonly tlToBr: Line | undefined;
  readonly blToTr: Line | undefined;
}

/** `a:tcPr`. Margins are EMU as written; the schema's defaults are 91440 and 45720. */
export interface TableCellProps {
  readonly marL: number | undefined;
  readonly marR: number | undefined;
  readonly marT: number | undefined;
  readonly marB: number | undefined;
  readonly vert: VerticalText | undefined;
  readonly anchor: TextAnchor | undefined;
  readonly anchorCtr: boolean | undefined;
  readonly horzOverflow: HorzOverflow | undefined;
  readonly borders: TableCellBorders;
  readonly fill: Fill | undefined;
  /** `a:headers/a:header`, the ids of the cells that head this one. */
  readonly headers: readonly string[];
  readonly node: XElement;
}

export interface Table {
  readonly props: TableProps | undefined;
  readonly columns: readonly TableColumn[];
  readonly rows: readonly TableRow[];
  readonly node: XElement;
}

/* -------------------------------------------------------------------------- */
/* the grid                                                                   */
/* -------------------------------------------------------------------------- */

/** A grid position's owner: the anchor cell, its clamped span, and the `a:tc` there. */
export interface GridCell {
  readonly row: number;
  readonly col: number;
  readonly rows: number;
  readonly cols: number;
  /** `null` for a position no `a:tc` reached, which PowerPoint draws empty. */
  readonly cell: TableCell | null;
}

/** The grid as drawn: `positions[r][c]` is the owner of (r, c); `anchors` lists each owner once. */
export interface TableGrid {
  readonly rows: number;
  readonly cols: number;
  /** Column widths in EMU. */
  readonly widths: readonly number[];
  /** Row heights in EMU, each the least its row is drawn at. */
  readonly heights: readonly number[];
  readonly positions: readonly (readonly GridCell[])[];
  readonly anchors: readonly GridCell[];
}

/** A span as PowerPoint counts it: zero is one, a negative runs to the edge. */
function spanOf(written: number): number {
  return written < 0 ? Number.POSITIVE_INFINITY : Math.max(1, written);
}

/**
 * The occupancy C7 measured, 76/76 against 52/76 for the flags: each `a:tc` takes the next
 * column, a claimed position is covered whatever it says, a span stops at the edge or a claim.
 */
export function tableGrid(table: Table): TableGrid {
  const widths =
    table.columns.length === 0 ? [INVENTED_COLUMN_WIDTH] : table.columns.map((c) => c.w);
  const heights = table.rows.length === 0 ? [0] : table.rows.map((r) => r.h);
  const rows = heights.length;
  const cols = widths.length;
  const grid: (GridCell | null)[][] = Array.from({ length: rows }, () =>
    Array.from<GridCell | null>({ length: cols }).fill(null),
  );
  const anchors: GridCell[] = [];
  const place = (anchor: GridCell): void => {
    anchors.push(anchor);
    for (let r = anchor.row; r < anchor.row + anchor.rows; r++) {
      for (let c = anchor.col; c < anchor.col + anchor.cols; c++) grid[r]![c] = anchor;
    }
  };
  // A grid that declares no column keeps none of its cells: PowerPoint draws the invented one empty.
  const placed = table.columns.length === 0 ? [] : table.rows;
  placed.forEach((row, r) => {
    row.cells.forEach((cell, c) => {
      if (c >= cols || grid[r]![c] !== null) return;
      const wide = spanOf(cell.gridSpan);
      const tall = spanOf(cell.rowSpan);
      let width = 1;
      while (c + width < cols && width < wide && grid[r]![c + width] === null) width += 1;
      // Below the anchor nothing can be claimed yet: an earlier anchor covering those rows
      // would have covered this row too, where the width just stopped.
      const height = Math.min(tall, rows - r);
      place({ row: r, col: c, rows: height, cols: width, cell });
    });
  });
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      if (grid[r]![c] === null) place({ row: r, col: c, rows: 1, cols: 1, cell: null });
    }
  }
  anchors.sort((a, b) => a.row - b.row || a.col - b.col);
  return {
    rows,
    cols,
    widths,
    heights,
    positions: grid.map((line) => line.map((owner) => owner!)),
    anchors,
  };
}

/* -------------------------------------------------------------------------- */
/* table styles                                                               */
/* -------------------------------------------------------------------------- */

/** `ST_OnOffStyleType`. */
export type OnOffStyle = 'on' | 'off' | 'def';

/** A value written out, or a reference into the theme: the `EG_Themeable*` choices. */
export type Themeable<T, Ref = StyleRef> =
  { readonly kind: 'value'; readonly value: T } | { readonly kind: 'ref'; readonly ref: Ref };

/** `a:tcTxStyle`. */
export interface TableStyleText {
  readonly b: OnOffStyle | undefined;
  readonly i: OnOffStyle | undefined;
  readonly font: Themeable<FontCollection, FontRef> | undefined;
  readonly color: Color | undefined;
}

/** `a:tcBdr`; `tl2br` and `tr2bl` take the names `TableCellBorders` gives the diagonals. */
export interface TableStyleBorders {
  readonly left: Themeable<Line> | undefined;
  readonly right: Themeable<Line> | undefined;
  readonly top: Themeable<Line> | undefined;
  readonly bottom: Themeable<Line> | undefined;
  readonly insideH: Themeable<Line> | undefined;
  readonly insideV: Themeable<Line> | undefined;
  readonly tlToBr: Themeable<Line> | undefined;
  readonly blToTr: Themeable<Line> | undefined;
}

/** `a:tcStyle`. */
export interface TableStyleCell {
  readonly borders: TableStyleBorders | undefined;
  readonly fill: Themeable<Fill> | undefined;
  /** `a:cell3D`, kept as written. */
  readonly cell3D: XElement | undefined;
}

/** One of a style's thirteen parts. */
export interface TableStylePart {
  readonly text: TableStyleText | undefined;
  readonly cell: TableStyleCell | undefined;
}

/** `a:tblBg`. An `a:effect` holding an `a:effectDag` reads as `undefined`, as `parseEffects` does. */
export interface TableBackground {
  readonly fill: Themeable<Fill> | undefined;
  readonly effect: Themeable<readonly Effect[] | undefined> | undefined;
}

/** The thirteen parts in schema order, which is the order they compose in, lowest first (C9). */
export type TableStylePartName =
  | 'wholeTbl'
  | 'band1H'
  | 'band2H'
  | 'band1V'
  | 'band2V'
  | 'lastCol'
  | 'firstCol'
  | 'lastRow'
  | 'seCell'
  | 'swCell'
  | 'firstRow'
  | 'neCell'
  | 'nwCell';

/** `CT_TableStyle`: an `a:tblStyle`, or an inline `a:tableStyle`. Nothing absent is defaulted. */
export interface TableStyle {
  readonly id: string;
  readonly name: string;
  readonly background: TableBackground | undefined;
  readonly parts: Readonly<Partial<Record<TableStylePartName, TableStylePart>>>;
  readonly node: XElement;
}

/* -------------------------------------------------------------------------- */
/* the cascade                                                                */
/* -------------------------------------------------------------------------- */

/** A grid edge: `h` is the top of row `row` at column `col` (`row = rows` is the bottom); `v` the left of column `col`. */
export interface TableEdge {
  readonly axis: 'h' | 'v';
  readonly row: number;
  readonly col: number;
}

/** Where a table's format came from: a style part, `tblBg`, a cell's `a:tcPr`, `a:tblPr`, or the default grid. */
export type TableSource = TableStylePartName | 'tblBg' | 'tcPr' | 'tblPr' | 'grid';

export interface TableSourced<T> {
  readonly value: T;
  readonly source: TableSource;
}
