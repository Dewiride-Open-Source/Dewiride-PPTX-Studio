/**
 * Experiment C7 - the probes: which grid positions a merged cell occupies, and what
 * fixes a row's height. One table per package, so a repair names its cause; ADR 0056.
 */

import { escapeText } from '../../lib/pptx.ts';

/* -------------------------------------------------------------------------- */
/* the grid                                                                   */
/* -------------------------------------------------------------------------- */

export const EMU_PER_POINT = 12700;
/** Where every probe's frame sits: one inch in from the top-left corner. */
export const FRAME = { x: 914400, y: 914400 } as const;
/** Four columns of 108 pt and three rows of 36 pt, unless a probe says otherwise. */
export const COL = 1371600;
export const ROW = 457200;
const COLS = [COL, COL, COL, COL];
const SZ = 1200;

export const TABLE_URI = 'http://schemas.openxmlformats.org/drawingml/2006/table';

export type Family =
  'control' | 'disagree' | 'conflict' | 'edge' | 'count' | 'lexical' | 'geometry' | 'style';

/** One `a:tc`. A string where a number or boolean is expected is written verbatim. */
export interface CellSpec {
  readonly gridSpan?: number | string;
  readonly rowSpan?: number | string;
  readonly hMerge?: boolean | string;
  readonly vMerge?: boolean | string;
  /** Replaces the `r{row}c{col}` marker. */
  readonly text?: string;
  /** This many one-line paragraphs, for a row that has to grow. */
  readonly lines?: number;
  readonly noBody?: boolean;
  readonly noProps?: boolean;
  /** An `a:txBody` with no `a:p` at all. */
  readonly emptyBody?: boolean;
  readonly tcPr?: string;
}

export interface RowSpec {
  /** `@h` in EMU; a string is written verbatim; `null` omits the attribute. */
  readonly h: number | string | null;
  readonly cells: readonly CellSpec[];
}

export interface Probe {
  readonly id: string;
  readonly family: Family;
  /** One line, in the present tense, saying what is being asked. */
  readonly question: string;
  readonly cols: readonly (number | string | null)[];
  readonly rows: readonly RowSpec[];
  readonly frame?: { readonly cx?: number; readonly cy?: number; readonly omit?: boolean };
  readonly tblPr?: string;
  /** Expected to be refused or repaired; a repaired deck measures the repair, not the probe. */
  readonly hostile?: true;
}

/* -------------------------------------------------------------------------- */
/* markup                                                                     */
/* -------------------------------------------------------------------------- */

function attr(name: string, value: number | string | boolean | undefined): string {
  if (value === undefined) return '';
  if (typeof value === 'boolean') return value ? ` ${name}="1"` : '';
  return ` ${name}="${String(value)}"`;
}

function paragraph(text: string): string {
  return `<a:p><a:r><a:rPr lang="en-US" sz="${String(SZ)}"/><a:t>${escapeText(text)}</a:t></a:r></a:p>`;
}

function cellXml(cell: CellSpec, row: number, col: number): string {
  const attrs =
    attr('rowSpan', cell.rowSpan) +
    attr('gridSpan', cell.gridSpan) +
    attr('hMerge', cell.hMerge) +
    attr('vMerge', cell.vMerge);
  if (cell.noBody === true && cell.noProps === true) return `<a:tc${attrs}/>`;
  const marker = cell.text ?? `r${String(row)}c${String(col)}`;
  const lines = cell.lines ?? 1;
  const paras = Array.from({ length: lines }, (_, i) =>
    paragraph(lines === 1 ? marker : `${marker} line ${String(i + 1)}`),
  ).join('');
  const body =
    cell.noBody === true
      ? ''
      : `<a:txBody><a:bodyPr/><a:lstStyle/>${cell.emptyBody === true ? '' : paras}</a:txBody>`;
  const props = cell.noProps === true ? '' : (cell.tcPr ?? '<a:tcPr/>');
  return `<a:tc${attrs}>${body}${props}</a:tc>`;
}

export function tableXml(probe: Probe): string {
  const grid = probe.cols
    .map((w) => (w === null ? '<a:gridCol/>' : `<a:gridCol w="${String(w)}"/>`))
    .join('');
  const rows = probe.rows
    .map(
      (row, r) =>
        `<a:tr${row.h === null ? '' : ` h="${String(row.h)}"`}>` +
        row.cells.map((cell, c) => cellXml(cell, r + 1, c + 1)).join('') +
        '</a:tr>',
    )
    .join('');
  return `<a:tbl>${probe.tblPr ?? '<a:tblPr/>'}<a:tblGrid>${grid}</a:tblGrid>${rows}</a:tbl>`;
}

const sum = (values: readonly (number | string | null)[]): number =>
  values.reduce<number>((total, v) => total + (typeof v === 'number' ? v : 0), 0);

/** The whole `p:graphicFrame`, its extent the grid's sum unless the probe overrides it. */
export function frameXml(probe: Probe): string {
  const cx = probe.frame?.cx ?? sum(probe.cols);
  const cy = probe.frame?.cy ?? sum(probe.rows.map((row) => row.h));
  const xfrm =
    probe.frame?.omit === true
      ? ''
      : `<p:xfrm><a:off x="${String(FRAME.x)}" y="${String(FRAME.y)}"/><a:ext cx="${String(cx)}" cy="${String(cy)}"/></p:xfrm>`;
  return (
    '<p:graphicFrame><p:nvGraphicFramePr>' +
    `<p:cNvPr id="2" name="${probe.id}"/>` +
    '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/>' +
    '</p:nvGraphicFramePr>' +
    xfrm +
    `<a:graphic><a:graphicData uri="${TABLE_URI}">` +
    tableXml(probe) +
    '</a:graphicData></a:graphic></p:graphicFrame>'
  );
}

/* -------------------------------------------------------------------------- */
/* probes                                                                     */
/* -------------------------------------------------------------------------- */

const plain = (): CellSpec => ({});
const row = (cells: readonly CellSpec[], h: number | string | null = ROW): RowSpec => ({
  h,
  cells,
});
/** Three rows of four plain cells, with the named cells replaced. */
function grid(overrides: Readonly<Record<string, CellSpec>> = {}): RowSpec[] {
  return [1, 2, 3].map((r) =>
    row(
      [1, 2, 3, 4].map((c) => {
        const at = overrides[`${String(r)},${String(c)}`];
        return at ?? plain();
      }),
    ),
  );
}

function probe(
  id: string,
  family: Family,
  question: string,
  rows: readonly RowSpec[],
  extra: Partial<Pick<Probe, 'cols' | 'frame' | 'tblPr' | 'hostile'>> = {},
): Probe {
  return { id, family, question, rows, cols: extra.cols ?? COLS, ...extra };
}

const CONTROLS: Probe[] = [
  probe(
    'h2',
    'control',
    'gridSpan="2" with an hMerge cell covers two columns',
    grid({
      '1,1': { gridSpan: 2 },
      '1,2': { hMerge: true },
    }),
  ),
  probe(
    'h3',
    'control',
    'gridSpan="3" with two hMerge cells covers three columns',
    grid({
      '1,1': { gridSpan: 3 },
      '1,2': { hMerge: true },
      '1,3': { hMerge: true },
    }),
  ),
  probe(
    'v2',
    'control',
    'rowSpan="2" with a vMerge cell covers two rows',
    grid({
      '1,1': { rowSpan: 2 },
      '2,1': { vMerge: true },
    }),
  ),
  probe(
    'v3',
    'control',
    'rowSpan="3" with two vMerge cells covers three rows',
    grid({
      '1,1': { rowSpan: 3 },
      '2,1': { vMerge: true },
      '3,1': { vMerge: true },
    }),
  ),
  probe(
    'block22',
    'control',
    'a 2x2 block: both spans on the anchor, both flags on the corner',
    grid({
      '1,2': { gridSpan: 2, rowSpan: 2 },
      '1,3': { hMerge: true },
      '2,2': { vMerge: true },
      '2,3': { hMerge: true, vMerge: true },
    }),
  ),
  probe(
    'block23',
    'control',
    'a 2x3 block, three columns wide and two rows tall',
    grid({
      '1,2': { gridSpan: 3, rowSpan: 2 },
      '1,3': { hMerge: true },
      '1,4': { hMerge: true },
      '2,2': { vMerge: true },
      '2,3': { hMerge: true, vMerge: true },
      '2,4': { hMerge: true, vMerge: true },
    }),
  ),
  probe(
    'h2-last',
    'control',
    'a horizontal merge ending in the last column',
    grid({
      '2,3': { gridSpan: 2 },
      '2,4': { hMerge: true },
    }),
  ),
  probe(
    'v2-last',
    'control',
    'a vertical merge ending in the last row',
    grid({
      '2,4': { rowSpan: 2 },
      '3,4': { vMerge: true },
    }),
  ),
  probe(
    'two-in-a-row',
    'control',
    'two horizontal merges in one row keep their columns',
    grid({
      '1,1': { gridSpan: 2 },
      '1,2': { hMerge: true },
      '1,3': { gridSpan: 2 },
      '1,4': { hMerge: true },
    }),
  ),
  probe(
    'h-over-v',
    'control',
    'a horizontal merge above a vertical one in the same column',
    grid({
      '1,1': { gridSpan: 2 },
      '1,2': { hMerge: true },
      '2,1': { rowSpan: 2 },
      '3,1': { vMerge: true },
    }),
  ),
  probe(
    'covered-text',
    'control',
    'text in a covered cell: is it shown, and is it kept',
    grid({
      '1,1': { gridSpan: 2 },
      '1,2': { hMerge: true, text: 'HIDDEN' },
    }),
  ),
  probe(
    'covered-text-v',
    'control',
    'text in a vertically covered cell',
    grid({
      '1,1': { rowSpan: 2 },
      '2,1': { vMerge: true, text: 'HIDDEN' },
    }),
  ),
];

const DISAGREE: Probe[] = [
  probe(
    'span-no-flag-h',
    'disagree',
    'gridSpan="2" but the next cell carries no hMerge',
    grid({
      '1,1': { gridSpan: 2 },
    }),
  ),
  probe(
    'flag-no-span-h',
    'disagree',
    'hMerge on a cell whose left neighbour has no gridSpan',
    grid({
      '1,2': { hMerge: true },
    }),
  ),
  probe(
    'span-no-flag-v',
    'disagree',
    'rowSpan="2" but the cell below carries no vMerge',
    grid({
      '1,1': { rowSpan: 2 },
    }),
  ),
  probe(
    'flag-no-span-v',
    'disagree',
    'vMerge on a cell whose upper neighbour has no rowSpan',
    grid({
      '2,1': { vMerge: true },
    }),
  ),
  probe(
    'flag-first-col',
    'disagree',
    'hMerge in the first column, with nothing to its left',
    grid({
      '1,1': { hMerge: true },
    }),
  ),
  probe(
    'flag-first-row',
    'disagree',
    'vMerge in the first row, with nothing above it',
    grid({
      '1,1': { vMerge: true },
    }),
  ),
  probe(
    'span-one-flag',
    'disagree',
    'an explicit gridSpan="1" beside an hMerge cell',
    grid({
      '1,1': { gridSpan: 1 },
      '1,2': { hMerge: true },
    }),
  ),
  probe(
    'flags-past-span-h',
    'disagree',
    'gridSpan="2" followed by two hMerge cells',
    grid({
      '1,1': { gridSpan: 2 },
      '1,2': { hMerge: true },
      '1,3': { hMerge: true },
    }),
  ),
  probe(
    'flags-past-span-v',
    'disagree',
    'rowSpan="2" followed by two vMerge cells',
    grid({
      '1,1': { rowSpan: 2 },
      '2,1': { vMerge: true },
      '3,1': { vMerge: true },
    }),
  ),
  probe(
    'span-short-of-flags-h',
    'disagree',
    'gridSpan="3" but only the first covered cell says hMerge',
    grid({
      '1,1': { gridSpan: 3 },
      '1,2': { hMerge: true },
    }),
  ),
  probe(
    'block-flags-only',
    'disagree',
    'a 2x2 block written with flags and no spans',
    grid({
      '1,3': { hMerge: true },
      '2,2': { vMerge: true },
      '2,3': { hMerge: true, vMerge: true },
    }),
  ),
  probe(
    'block-spans-only',
    'disagree',
    'a 2x2 block written with spans and no flags',
    grid({
      '1,2': { gridSpan: 2, rowSpan: 2 },
    }),
  ),
  probe(
    'block-corner-bare',
    'disagree',
    'a 2x2 block whose corner cell carries no flag',
    grid({
      '1,2': { gridSpan: 2, rowSpan: 2 },
      '1,3': { hMerge: true },
      '2,2': { vMerge: true },
    }),
  ),
  probe(
    'block-corner-h-only',
    'disagree',
    'a 2x2 block whose corner says hMerge but not vMerge',
    grid({
      '1,2': { gridSpan: 2, rowSpan: 2 },
      '1,3': { hMerge: true },
      '2,2': { vMerge: true },
      '2,3': { hMerge: true },
    }),
  ),
  probe(
    'both-flags-no-span',
    'disagree',
    'hMerge and vMerge together with no span anywhere',
    grid({
      '2,2': { hMerge: true, vMerge: true },
    }),
  ),
];

const CONFLICT: Probe[] = [
  probe(
    'covered-anchor-v',
    'conflict',
    'a horizontally covered cell that itself claims rowSpan="2"',
    grid({
      '1,1': { gridSpan: 2 },
      '1,2': { hMerge: true, rowSpan: 2 },
      '2,2': { vMerge: true },
    }),
  ),
  probe(
    'covered-anchor-h',
    'conflict',
    'a vertically covered cell that itself claims gridSpan="2"',
    grid({
      '1,1': { rowSpan: 2 },
      '2,1': { vMerge: true, gridSpan: 2 },
      '2,2': { hMerge: true },
    }),
  ),
  probe(
    'cross',
    'conflict',
    'one position claimed by a rowSpan from above and a gridSpan from the left',
    grid({
      '1,2': { rowSpan: 2 },
      '2,1': { gridSpan: 2 },
      '2,2': { hMerge: true, vMerge: true },
    }),
  ),
  probe(
    'overlap-h',
    'conflict',
    'two gridSpans in one row whose ranges overlap',
    grid({
      '1,1': { gridSpan: 3 },
      '1,2': { hMerge: true, gridSpan: 2 },
      '1,3': { hMerge: true },
    }),
  ),
  probe(
    'cross-partial',
    'conflict',
    'a gridSpan="4" that meets a rowSpan from above at its third column',
    grid({
      '1,3': { rowSpan: 2 },
      '2,1': { gridSpan: 4 },
      '2,2': { hMerge: true },
      '2,3': { hMerge: true, vMerge: true },
      '2,4': { hMerge: true },
    }),
  ),
];

const EDGE: Probe[] = [
  probe(
    'span-over-edge-h',
    'edge',
    'gridSpan="3" with two columns left',
    grid({
      '1,3': { gridSpan: 3 },
      '1,4': { hMerge: true },
    }),
  ),
  probe(
    'span-over-edge-v',
    'edge',
    'rowSpan="3" with two rows left',
    grid({
      '2,1': { rowSpan: 3 },
      '3,1': { vMerge: true },
    }),
  ),
  probe('span-zero', 'edge', 'gridSpan="0"', grid({ '1,1': { gridSpan: 0 } })),
  probe('span-negative', 'edge', 'gridSpan="-1"', grid({ '1,1': { gridSpan: -1 } })),
  probe('rowspan-zero', 'edge', 'rowSpan="0"', grid({ '1,1': { rowSpan: 0 } })),
  probe(
    'rowspan-one',
    'edge',
    'an explicit rowSpan="1" changes nothing',
    grid({ '1,1': { rowSpan: 1 } }),
  ),
  probe('rowspan-negative', 'edge', 'rowSpan="-1"', grid({ '1,2': { rowSpan: -1 } })),
  probe('span-negative-2', 'edge', 'gridSpan="-2"', grid({ '2,2': { gridSpan: -2 } })),
  probe('span-huge', 'edge', 'gridSpan="99"', grid({ '1,2': { gridSpan: 99 } })),
  probe('rowspan-huge', 'edge', 'rowSpan="99"', grid({ '1,2': { rowSpan: 99 } })),
  probe(
    'span-at-edge',
    'edge',
    'gridSpan="2" in the last column',
    grid({ '1,4': { gridSpan: 2 } }),
  ),
  probe('rowspan-at-edge', 'edge', 'rowSpan="2" in the last row', grid({ '3,1': { rowSpan: 2 } })),
];

const COUNT: Probe[] = [
  probe('row-short', 'count', 'a row with three cells under a four-column grid', [
    row([plain(), plain(), plain()]),
    ...grid().slice(1),
  ]),
  probe('row-long', 'count', 'a row with five cells under a four-column grid', [
    row([plain(), plain(), plain(), plain(), { text: 'FIFTH' }]),
    ...grid().slice(1),
  ]),
  probe(
    'row-short-span',
    'count',
    'gridSpan="2" and no covered cell: three cells summing to four columns',
    [row([{ gridSpan: 2 }, plain(), plain()]), ...grid().slice(1)],
  ),
  probe('row-short-span-v', 'count', 'rowSpan="2" and no covered cell in the row below', [
    row([{ rowSpan: 2 }, plain(), plain(), plain()]),
    row([plain(), plain(), plain()]),
    row([plain(), plain(), plain(), plain()]),
  ]),
  probe('row-empty', 'count', 'a row with no cells at all', [row([]), ...grid().slice(1)], {
    hostile: true,
  }),
  probe('no-rows', 'count', 'a grid and no rows', [], { hostile: true }),
  probe('no-cols', 'count', 'rows of four cells over an empty grid', grid(), {
    cols: [],
    hostile: true,
  }),
  probe('cols-more', 'count', 'five grid columns over rows of four cells', grid(), {
    cols: [COL, COL, COL, COL, COL],
  }),
  probe('gridcol-no-w', 'count', 'a gridCol without @w', grid(), {
    cols: [COL, null, COL, COL],
    hostile: true,
  }),
  probe(
    'tr-no-h',
    'count',
    'a row without @h',
    [row([plain(), plain(), plain(), plain()], null), ...grid().slice(1)],
    { hostile: true },
  ),
  probe('tc-no-body', 'count', 'a cell with no a:txBody', grid({ '1,1': { noBody: true } })),
  probe('tc-no-props', 'count', 'a cell with no a:tcPr', grid({ '1,1': { noProps: true } })),
  probe('tc-bare', 'count', 'an empty a:tc', grid({ '1,1': { noBody: true, noProps: true } })),
  probe(
    'tc-body-no-p',
    'count',
    'an a:txBody with no paragraph',
    grid({ '1,1': { emptyBody: true } }),
    {
      hostile: true,
    },
  ),
];

const LEXICAL: Probe[] = [
  probe(
    'merge-true',
    'lexical',
    'hMerge="true" and vMerge="true" are the flags',
    grid({
      '1,1': { gridSpan: 2, rowSpan: 2 },
      '1,2': { hMerge: 'true' },
      '2,1': { vMerge: 'true' },
      '2,2': { hMerge: 'true', vMerge: 'true' },
    }),
  ),
  probe(
    'merge-false',
    'lexical',
    'hMerge="false" is no flag, even beside a gridSpan',
    grid({
      '1,1': { gridSpan: 2 },
      '1,2': { hMerge: 'false' },
      '2,2': { hMerge: 'false' },
    }),
  ),
  probe(
    'merge-on',
    'lexical',
    'hMerge="on" is not an xsd:boolean',
    grid({
      '1,1': { gridSpan: 2 },
      '1,2': { hMerge: 'on' },
    }),
    { hostile: true },
  ),
  probe(
    'span-plus',
    'lexical',
    'gridSpan="+2" is a lexically legal xsd:int',
    grid({
      '1,1': { gridSpan: '+2' },
      '1,2': { hMerge: true },
    }),
  ),
  probe(
    'span-float',
    'lexical',
    'gridSpan="2.0" is not an xsd:int',
    grid({
      '1,1': { gridSpan: '2.0' },
      '1,2': { hMerge: true },
    }),
    { hostile: true },
  ),
  probe('w-points', 'lexical', 'gridCol w="108pt", the universal-measure spelling', grid(), {
    cols: ['108pt', COL, COL, COL],
    hostile: true,
  }),
];

const GEOMETRY: Probe[] = [
  probe('frame-narrow', 'geometry', 'a frame half as wide as its grid', grid(), {
    frame: { cx: 2 * COL },
  }),
  probe('frame-wide', 'geometry', 'a frame twice as wide as its grid', grid(), {
    frame: { cx: 8 * COL },
  }),
  probe('frame-short', 'geometry', 'a frame half as tall as its rows', grid(), {
    frame: { cy: Math.round(1.5 * ROW) },
  }),
  probe('frame-tall', 'geometry', 'a frame twice as tall as its rows', grid(), {
    frame: { cy: 6 * ROW },
  }),
  probe('frame-missing', 'geometry', 'a graphicFrame with no p:xfrm', grid(), {
    frame: { omit: true },
    hostile: true,
  }),
  probe(
    'row-grows',
    'geometry',
    'three lines in a 36 pt row: does the row grow',
    grid({
      '2,2': { lines: 3 },
    }),
  ),
  probe('row-generous', 'geometry', 'one line in a 108 pt row: does the row shrink', [
    grid()[0]!,
    row(grid()[1]!.cells, 3 * ROW),
    grid()[2]!,
  ]),
  probe('row-zero', 'geometry', 'h="0": the least a row of one line can be', [
    grid()[0]!,
    row(grid()[1]!.cells, 0),
    grid()[2]!,
  ]),
  probe('row-zero-empty', 'geometry', 'h="0" on a row of empty cells', [
    grid()[0]!,
    row([{ text: '' }, { text: '' }, { text: '' }, { text: '' }], 0),
    grid()[2]!,
  ]),
  probe('col-zero', 'geometry', 'a column of zero width', grid(), {
    cols: [COL, 0, COL, COL],
  }),
  probe('col-narrow', 'geometry', 'a column narrower than its cells’ margins', grid(), {
    cols: [COL, 100000, COL, COL],
  }),
  probe(
    'col-zero-no-margin',
    'geometry',
    'a zero column whose cells have no side margins',
    grid({
      '1,2': { tcPr: '<a:tcPr marL="0" marR="0"/>' },
      '2,2': { tcPr: '<a:tcPr marL="0" marR="0"/>' },
      '3,2': { tcPr: '<a:tcPr marL="0" marR="0"/>' },
    }),
    { cols: [COL, 0, COL, COL] },
  ),
  probe(
    'col-zero-wide-margin',
    'geometry',
    'a zero column whose cells have 0.2 in side margins',
    grid({
      '1,2': { tcPr: '<a:tcPr marL="182880" marR="182880"/>' },
      '2,2': { tcPr: '<a:tcPr marL="182880" marR="182880"/>' },
      '3,2': { tcPr: '<a:tcPr marL="182880" marR="182880"/>' },
    }),
    { cols: [COL, 0, COL, COL] },
  ),
  probe(
    'tcpr-margins',
    'geometry',
    'the four margins as written, beside the defaults',
    grid({
      '1,1': { tcPr: '<a:tcPr marL="0" marR="182880" marT="0" marB="91440"/>' },
      '1,2': { tcPr: '<a:tcPr marL="0" marR="0" marT="0" marB="0"/>' },
    }),
  ),
  probe(
    'tcpr-anchor',
    'geometry',
    'anchor="b", anchor="ctr" and vert="vert270" as written',
    grid({
      '1,1': { tcPr: '<a:tcPr anchor="b"/>' },
      '1,2': { tcPr: '<a:tcPr anchor="ctr"/>' },
      '1,3': { tcPr: '<a:tcPr vert="vert270"/>' },
      '1,4': { tcPr: '<a:tcPr anchor="ctr" anchorCtr="1"/>' },
    }),
  ),
];

const STYLE: Probe[] = [
  probe('no-tblpr', 'style', 'a table with no a:tblPr', grid(), { tblPr: '' }),
  probe('style-unknown', 'style', 'a tableStyleId nobody defines', grid(), {
    tblPr:
      '<a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>{00000000-0000-0000-0000-00000000C700}</a:tableStyleId></a:tblPr>',
  }),
  probe(
    'style-builtin',
    'style',
    'the id of a built-in style the package does not define',
    grid(),
    {
      tblPr:
        '<a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId></a:tblPr>',
    },
  ),
];

export function allProbes(): readonly Probe[] {
  const probes = [
    ...CONTROLS,
    ...DISAGREE,
    ...CONFLICT,
    ...EDGE,
    ...COUNT,
    ...LEXICAL,
    ...GEOMETRY,
    ...STYLE,
  ];
  const seen = new Set<string>();
  for (const p of probes) {
    if (seen.has(p.id)) throw new Error(`duplicate probe id: ${p.id}`);
    seen.add(p.id);
  }
  return probes;
}
