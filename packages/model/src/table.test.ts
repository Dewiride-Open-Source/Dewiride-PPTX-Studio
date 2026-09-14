/**
 * Tables, against what PowerPoint actually did: every probe's `a:tbl` is parsed by this package
 * and the rectangle each position would report under `tableGrid` has to be the one PowerPoint
 * reported, re-derived from the widths and heights it read rather than from a summary.
 */

import { parseXmlString } from '@pptx-studio/xml';
import { describe, expect, it } from 'vitest';

import fixture from '../../../corpus/ground-truth/tables.json' with { type: 'json' };

import { ModelError } from './errors.js';
import { parseSheet } from './parse/sheet.js';
import { INVENTED_COLUMN_WIDTH, TABLE_URI, tableGrid, type GridCell, type Table } from './table.js';

/* -------------------------------------------------------------------------- */
/* building slides by hand                                                    */
/* -------------------------------------------------------------------------- */

const NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' +
  ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
  ' xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';

const SP_TREE_HEAD =
  '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  '<p:grpSpPr/>';

function frameXml(inner: string, uri = TABLE_URI, xfrm = true): string {
  return (
    '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="table"/>' +
    '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/>' +
    '</p:nvGraphicFramePr>' +
    (xfrm
      ? '<p:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="1371600"/></p:xfrm>'
      : '') +
    `<a:graphic><a:graphicData uri="${uri}">${inner}</a:graphicData></a:graphic></p:graphicFrame>`
  );
}

function tableOf(shapes: string): Table | undefined {
  const xml =
    `<p:sld ${NS}><p:cSld name="slide">${SP_TREE_HEAD}${shapes}</p:spTree></p:cSld>` +
    '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>';
  const sheet = parseSheet(parseXmlString(xml).root, '/ppt/slides/slide1.xml');
  return sheet.shapes[0]?.table;
}

function tableFrom(tbl: string): Table {
  const table = tableOf(frameXml(tbl));
  if (table === undefined) throw new Error('expected a table');
  return table;
}

const cell = (text: string, attrs = ''): string =>
  `<a:tc${attrs}><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>${text}</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>`;
const row = (cells: string, h = 457200): string => `<a:tr h="${String(h)}">${cells}</a:tr>`;
const grid = (n: number, w = 1371600): string =>
  `<a:tblGrid>${`<a:gridCol w="${String(w)}"/>`.repeat(n)}</a:tblGrid>`;

/* -------------------------------------------------------------------------- */
/* the fixture                                                                */
/* -------------------------------------------------------------------------- */

type FixtureProbe = (typeof fixture.probes)[number];
type FixtureOm = NonNullable<FixtureProbe['om']>;

/** The rectangle a position reports under an occupancy, from the sizes PowerPoint read. */
function predicted(om: FixtureOm, owner: GridCell): readonly number[] {
  const sum = (xs: readonly (number | null)[]): number =>
    xs.reduce<number>((a, b) => a + (b ?? 0), 0);
  return [
    om.frame[0]! + sum(om.colWidths.slice(0, owner.col)),
    om.frame[1]! + sum(om.rowHeights.slice(0, owner.row)),
    sum(om.colWidths.slice(owner.col, owner.col + owner.cols)),
    sum(om.rowHeights.slice(owner.row, owner.row + owner.rows)),
  ];
}

/** The text a position reports: its owner's paragraphs, as the object model joins them. */
function textOf(owner: GridCell): string {
  return (owner.cell?.text?.paragraphs ?? [])
    .map((p) => p.content.map((c) => (c.kind === 'run' ? c.text : '')).join(''))
    .join('\r');
}

/** Every position whose predicted rectangle or text is not the one PowerPoint reported. */
function misfits(om: FixtureOm, positions: readonly (readonly GridCell[])[]): string[] {
  const out: string[] = [];
  if (positions.length !== om.cells.length || positions[0]!.length !== om.cells[0]!.length) {
    return [
      `${String(positions.length)}x${String(positions[0]!.length)} against ${String(om.cells.length)}x${String(om.cells[0]!.length)}`,
    ];
  }
  om.cells.forEach((line, r) =>
    line.forEach((measured, c) => {
      if (measured === null) return;
      const owner = positions[r]![c]!;
      const rect = predicted(om, owner);
      const same = rect.every((v, i) => Math.abs(v - (measured[i] as number)) <= 0.05);
      if (!same)
        out.push(
          `${String(r + 1)},${String(c + 1)}: ${rect.join('/')} against ${measured.slice(0, 4).join('/')}`,
        );
      if (textOf(owner) !== measured[4])
        out.push(
          `${String(r + 1)},${String(c + 1)}: "${textOf(owner)}" against "${String(measured[4])}"`,
        );
    }),
  );
  return out;
}

const asWritten = fixture.probes.filter((p) => p.repaired === false && p.om !== null);
const repaired = fixture.probes.filter((p) => p.repaired === true);

describe('the grid, re-derived from every rectangle PowerPoint reported', () => {
  it('is the measurement it says it is', () => {
    expect(fixture.probes).toHaveLength(82);
    expect(asWritten).toHaveLength(76);
    expect(repaired.map((p) => p.id)).toEqual(fixture.findings.repaired);
    expect(fixture.findings.occupancy.winner).toBe('spans, zero is one, negative to the edge');
  });

  it.each(asWritten.map((p) => [p.id, p] as const))('reproduces %s', (_, probe) => {
    const positions = tableGrid(tableFrom(probe.markup)).positions;
    expect(misfits(probe.om, positions), probe.question).toEqual([]);
  });

  it('refutes the flags a person reads first, scored over the same cases', () => {
    const fits = asWritten.filter(
      (probe) => misfits(probe.om, byFlags(tableFrom(probe.markup))).length === 0,
    );
    const recorded = fixture.findings.occupancy.candidates.find((c) => c.name === 'flags only');
    expect(fits.length).toBe(recorded?.fits);
    expect(fits.length).toBeLessThan(asWritten.length);
  });

  it.each(
    fixture.findings.cellProperties.flatMap((p) =>
      p.cells.map((c) => [`${p.id} ${c.at}`, c] as const),
    ),
  )('reads a:tcPr as PowerPoint reads %s', (_, cell) => {
    const props = tableFrom(
      `<a:tbl>${grid(1)}<a:tr h="457200"><a:tc>${cell.tcPr}</a:tc></a:tr></a:tbl>`,
    ).rows[0]!.cells[0]!.props;
    const emu = (v: number | undefined, fallback: number): number => (v ?? fallback) / 12700;
    expect([
      emu(props?.marL, 91440),
      emu(props?.marT, 45720),
      emu(props?.marR, 91440),
      emu(props?.marB, 45720),
    ]).toEqual(cell.margins);
    expect(props?.anchor ?? 't').toBe(cell.anchor);
    expect(props?.vert ?? 'horz').toBe(cell.vert);
  });

  it('reads PowerPoint’s own merges and splits the same way', () => {
    for (const slide of fixture.authored) {
      const tbl =
        '<a:tbl><a:tblPr/>' +
        `<a:tblGrid>${slide.markup.cols.map((w) => `<a:gridCol w="${String(w)}"/>`).join('')}</a:tblGrid>` +
        slide.markup.rows
          .map(
            (r) =>
              `<a:tr h="${String(r.h)}">${r.cells
                .map((c) => {
                  const { text, ...rest } = c as Record<string, string> & { text?: string };
                  const attrs = Object.entries(rest)
                    .map(([k, v]) => ` ${k}="${v}"`)
                    .join('');
                  const paragraphs = (text ?? '')
                    .split('\r')
                    .map((t) =>
                      t === ''
                        ? '<a:p/>'
                        : `<a:p><a:r><a:rPr lang="en-US"/><a:t>${t}</a:t></a:r></a:p>`,
                    )
                    .join('');
                  return `<a:tc${attrs}><a:txBody><a:bodyPr/><a:lstStyle/>${paragraphs}</a:txBody><a:tcPr/></a:tc>`;
                })
                .join('')}</a:tr>`,
          )
          .join('') +
        '</a:tbl>';
      expect(
        slide.om.cells.flat().every((cell) => cell !== null),
        slide.name,
      ).toBe(true);
      expect(misfits(slide.om, tableGrid(tableFrom(tbl)).positions), slide.name).toEqual([]);
    }
    expect(fixture.authored).toHaveLength(21);
  });
});

/** The wrong rule: `hMerge` joins a cell to its left neighbour, `vMerge` to the one above. */
function byFlags(table: Table): GridCell[][] {
  const rows = Math.max(1, table.rows.length);
  const cols = Math.max(1, table.columns.length);
  const grid: GridCell[][] = [];
  for (let r = 0; r < rows; r++) {
    const line: GridCell[] = [];
    for (let c = 0; c < cols; c++) {
      const tc = table.columns.length === 0 ? undefined : table.rows[r]?.cells[c];
      const h = tc?.hMerge === true && c > 0;
      const v = tc?.vMerge === true && r > 0;
      const own: GridCell = { row: r, col: c, rows: 1, cols: 1, cell: tc ?? null };
      line.push(h && v ? grid[r - 1]![c - 1]! : h ? line[c - 1]! : v ? grid[r - 1]![c]! : own);
    }
    grid.push(line);
  }
  const extent = new Map<GridCell, { rows: number; cols: number }>();
  grid.forEach((line, r) =>
    line.forEach((a, c) => {
      const e = extent.get(a) ?? { rows: 1, cols: 1 };
      extent.set(a, {
        rows: Math.max(e.rows, r - a.row + 1),
        cols: Math.max(e.cols, c - a.col + 1),
      });
    }),
  );
  return grid.map((line) => line.map((a) => ({ ...a, ...extent.get(a)! })));
}

/* -------------------------------------------------------------------------- */
/* what PowerPoint repairs                                                    */
/* -------------------------------------------------------------------------- */

describe('what PowerPoint repairs', () => {
  const markup = (id: string): string => {
    const probe = fixture.probes.find((p) => p.id === id);
    if (probe === undefined) throw new Error(`no probe ${id}`);
    return probe.markup;
  };

  it.each(['gridcol-no-w', 'tr-no-h', 'merge-on', 'span-float'])(
    'refuses %s, as a typed error',
    (id) => {
      expect(() => tableFrom(markup(id))).toThrow(ModelError);
      try {
        tableFrom(markup(id));
      } catch (error) {
        expect((error as ModelError).code).toBe('MODEL_TABLE_ATTR');
        expect((error as ModelError).partName).toBe('/ppt/slides/slide1.xml');
      }
    },
  );

  it('reads a cell body with no paragraph as an empty body', () => {
    const table = tableFrom(markup('tc-body-no-p'));
    expect(table.rows[0]!.cells[0]!.text?.paragraphs).toEqual([]);
  });

  it('reads a frame with no p:xfrm, and keeps the absence', () => {
    const xml =
      `<p:sld ${NS}><p:cSld name="slide">${SP_TREE_HEAD}${frameXml(markup('frame-missing'), TABLE_URI, false)}</p:spTree></p:cSld>` +
      '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>';
    const shape = parseSheet(parseXmlString(xml).root, '/ppt/slides/slide1.xml').shapes[0]!;
    expect(shape.xfrm).toBeUndefined();
    expect(shape.table?.rows).toHaveLength(3);
  });
});

/* -------------------------------------------------------------------------- */
/* parsing                                                                    */
/* -------------------------------------------------------------------------- */

describe('parsing', () => {
  it('is undefined for a frame that holds no table, and for a shape', () => {
    expect(
      tableOf(frameXml('<c:chart/>', 'http://schemas.openxmlformats.org/drawingml/2006/chart')),
    ).toBeUndefined();
    expect(
      tableOf(
        '<p:sp><p:nvSpPr><p:cNvPr id="3" name="x"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp>',
      ),
    ).toBeUndefined();
  });

  it('reads the grid, the rows and every cell’s spans as written', () => {
    const table = tableFrom(
      `<a:tbl><a:tblPr/>${grid(3, 1000000)}` +
        row(cell('a', ' gridSpan="2"') + cell('', ' hMerge="1"') + cell('c'), 300000) +
        row(cell('d', ' rowSpan="-1"') + cell('e') + cell('f', ' rowSpan="1"')) +
        '</a:tbl>',
    );
    expect(table.columns.map((c) => c.w)).toEqual([1000000, 1000000, 1000000]);
    expect(table.rows.map((r) => r.h)).toEqual([300000, 457200]);
    const [first, second] = table.rows;
    expect(first!.cells.map((c) => [c.gridSpan, c.rowSpan, c.hMerge, c.vMerge])).toEqual([
      [2, 1, false, false],
      [1, 1, true, false],
      [1, 1, false, false],
    ]);
    expect(second!.cells[0]!.rowSpan).toBe(-1);
    const flags = tableFrom(
      `<a:tbl>${grid(2)}<a:tr h="1"><a:tc hMerge="false"/><a:tc vMerge="0" hMerge="true"/></a:tr></a:tbl>`,
    );
    expect(flags.rows[0]!.cells.map((c) => [c.hMerge, c.vMerge])).toEqual([
      [false, false],
      [true, false],
    ]);
    expect(first!.cells[0]!.text?.paragraphs[0]?.content[0]).toMatchObject({
      kind: 'run',
      text: 'a',
    });
  });

  it('honours a universal measure, which PowerPoint reads as the same width', () => {
    const table = tableFrom(`<a:tbl>${grid(1)}</a:tbl>`.replace('w="1371600"', 'w="108pt"'));
    expect(table.columns[0]!.w).toBe(1371600);
    expect(
      tableFrom(
        `<a:tbl><a:tblGrid><a:gridCol w="1in"/><a:gridCol w="2.5cm"/></a:tblGrid></a:tbl>`,
      ).columns.map((c) => c.w),
    ).toEqual([914400, 900000]);
  });

  it('reads a:tblPr: the seven flags, a fill, and either style source', () => {
    const byId = tableFrom(
      '<a:tbl><a:tblPr firstRow="1" bandRow="true" rtl="0"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>' +
        '<a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId></a:tblPr>' +
        `${grid(1)}</a:tbl>`,
    );
    expect(byId.props).toMatchObject({
      firstRow: true,
      bandRow: true,
      rtl: false,
      firstCol: false,
      lastRow: false,
      lastCol: false,
      bandCol: false,
      fill: { type: 'solid' },
      style: { kind: 'id', id: '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}' },
    });
    const inline = tableFrom(
      `<a:tbl><a:tblPr><a:tableStyle styleId="{1B1D64F0-0000-4000-8000-0000000A200A}" styleName="x"/></a:tblPr>${grid(1)}</a:tbl>`,
    );
    expect(inline.props?.style?.kind).toBe('inline');
    expect(tableFrom(`<a:tbl>${grid(1)}</a:tbl>`).props).toBeUndefined();
  });

  it('reads a:tcPr: margins as written, the anchors, the six borders, a fill and the headers', () => {
    const table = tableFrom(
      `<a:tbl>${grid(1)}<a:tr h="1"><a:tc id="h1"><a:txBody><a:bodyPr/><a:p/></a:txBody>` +
        '<a:tcPr marL="0" marR="182880" vert="vert270" anchor="ctr" anchorCtr="1" horzOverflow="clip">' +
        '<a:lnL w="12700"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:lnL>' +
        '<a:lnB w="25400"><a:noFill/></a:lnB>' +
        '<a:lnTlToBr w="6350"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill><a:prstDash val="dash"/></a:lnTlToBr>' +
        '<a:solidFill><a:srgbClr val="FFFF00"/></a:solidFill>' +
        '<a:headers><a:header>h0</a:header><a:header>h9</a:header></a:headers>' +
        '</a:tcPr></a:tc></a:tr></a:tbl>',
    );
    const tc = table.rows[0]!.cells[0]!;
    expect(tc.id).toBe('h1');
    expect(tc.props).toMatchObject({
      marL: 0,
      marR: 182880,
      marT: undefined,
      marB: undefined,
      vert: 'vert270',
      anchor: 'ctr',
      anchorCtr: true,
      horzOverflow: 'clip',
      fill: { type: 'solid' },
      headers: ['h0', 'h9'],
    });
    expect(tc.props?.borders.left).toMatchObject({ w: 12700, fill: { type: 'solid' } });
    expect(tc.props?.borders.bottom).toMatchObject({ w: 25400, fill: { type: 'none' } });
    expect(tc.props?.borders.tlToBr).toMatchObject({
      w: 6350,
      dash: { kind: 'preset', val: 'dash' },
    });
    expect(tc.props?.borders.right).toBeUndefined();
    expect(tc.props?.borders.top).toBeUndefined();
    expect(tc.props?.borders.blToTr).toBeUndefined();
  });

  it('keeps an absent a:tcPr and an absent a:txBody absent', () => {
    const table = tableFrom(
      `<a:tbl>${grid(2)}<a:tr h="1"><a:tc/><a:tc><a:tcPr/></a:tc></a:tr></a:tbl>`,
    );
    expect(table.rows[0]!.cells[0]).toMatchObject({ text: undefined, props: undefined });
    expect(table.rows[0]!.cells[1]!.props?.marL).toBeUndefined();
  });

  it('refuses a span that is not an int and a flag that is not a boolean', () => {
    expect(() =>
      tableFrom(`<a:tbl>${grid(1)}<a:tr h="1"><a:tc gridSpan="two"/></a:tr></a:tbl>`),
    ).toThrow(ModelError);
    expect(() =>
      tableFrom(`<a:tbl>${grid(1)}<a:tr h="1"><a:tc vMerge="yes"/></a:tr></a:tbl>`),
    ).toThrow(ModelError);
    expect(() => tableFrom(`<a:tbl><a:tblGrid><a:gridCol w="wide"/></a:tblGrid></a:tbl>`)).toThrow(
      ModelError,
    );
    expect(
      tableFrom(`<a:tbl>${grid(1)}<a:tr h="1"><a:tc gridSpan="+2"/></a:tr></a:tbl>`).rows[0]!
        .cells[0]!.gridSpan,
    ).toBe(2);
  });
});

/* -------------------------------------------------------------------------- */
/* the grid                                                                   */
/* -------------------------------------------------------------------------- */

describe('the grid', () => {
  it('invents one column for a grid with none and one row for a table with none, as PowerPoint does', () => {
    const noCols = tableGrid(tableFrom(`<a:tbl><a:tblGrid/>${row(cell('a') + cell('b'))}</a:tbl>`));
    expect([noCols.rows, noCols.cols, noCols.widths]).toEqual([1, 1, [INVENTED_COLUMN_WIDTH]]);
    expect(noCols.anchors.every((a) => a.cell === null)).toBe(true);
    const noRows = tableGrid(tableFrom(`<a:tbl>${grid(4)}</a:tbl>`));
    expect([noRows.rows, noRows.cols, noRows.heights]).toEqual([1, 4, [0]]);
    expect(noRows.anchors.every((a) => a.cell === null)).toBe(true);
  });

  it('pads a short row with empty positions and drops a cell past the last column', () => {
    const table = tableFrom(
      `<a:tbl>${grid(3)}${row(cell('a') + cell('b'))}${row(cell('c') + cell('d') + cell('e') + cell('f'))}</a:tbl>`,
    );
    const g = tableGrid(table);
    expect(g.positions[0]![2]!.cell).toBeNull();
    expect(g.positions[1]!.map((p) => p.cell?.text?.paragraphs[0]?.content[0])).toMatchObject([
      { text: 'c' },
      { text: 'd' },
      { text: 'e' },
    ]);
    expect(g.anchors).toHaveLength(6);
  });

  it('lists each anchor once, in reading order, with its clamped span', () => {
    const table = tableFrom(
      `<a:tbl>${grid(4)}` +
        row(
          cell('a') + cell('b', ' gridSpan="2" rowSpan="2"') + cell('', ' hMerge="1"') + cell('d'),
        ) +
        row(
          cell('e') +
            cell('', ' vMerge="1"') +
            cell('', ' hMerge="1" vMerge="1"') +
            cell('h', ' gridSpan="99"'),
        ) +
        '</a:tbl>',
    );
    expect(tableGrid(table).anchors.map((a) => [a.row, a.col, a.rows, a.cols])).toEqual([
      [0, 0, 1, 1],
      [0, 1, 2, 2],
      [0, 3, 1, 1],
      [1, 0, 1, 1],
      [1, 3, 1, 1],
    ]);
  });
});
