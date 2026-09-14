/**
 * `a:tbl` into a `Table`: the grid, the rows, every cell's spans, body and properties.
 * Nothing here resolves the merges; `tableGrid` does, by the rule C7 measured.
 */

import { attributeValue, childElements, firstChild, textContent } from '@pptx-studio/xml';
import type { XElement } from '@pptx-studio/xml';

import { ModelError } from '../errors.js';
import { TABLE_URI } from '../table.js';
import type {
  Table,
  TableCell,
  TableCellBorders,
  TableCellProps,
  TableColumn,
  TableProps,
  TableRow,
  TableStyleRef,
} from '../table.js';
import type { HorzOverflow, TextAnchor, VerticalText } from '../text.js';
import { parseEffects, parseFill, parseLineElement } from './paint.js';
import { parseTextBodyChild } from './text.js';

const ANCHORS: readonly TextAnchor[] = ['t', 'ctr', 'b', 'just', 'dist'];
const VERTICALS: readonly VerticalText[] = [
  'horz',
  'vert',
  'vert270',
  'wordArtVert',
  'eaVert',
  'mongolianVert',
  'wordArtVertRtl',
];
const HORZ_OVERFLOWS: readonly HorzOverflow[] = ['overflow', 'clip'];

/** EMU per unit of `ST_UniversalMeasure`. */
const UNITS: Readonly<Record<string, number>> = {
  mm: 36000,
  cm: 360000,
  in: 914400,
  pt: 12700,
  pc: 152400,
  pi: 152400,
};

function attrError(element: XElement, name: string, raw: string, why: string, part: string): never {
  throw new ModelError(
    'MODEL_TABLE_ATTR',
    `${element.qname}/@${name} is "${raw}", which is not ${why}`,
    part,
    name,
  );
}

function boolOf(element: XElement, name: string, part: string): boolean | undefined {
  const raw = attributeValue(element, name);
  if (raw === undefined) return undefined;
  if (raw === '1' || raw === 'true') return true;
  if (raw === '0' || raw === 'false') return false;
  return attrError(element, name, raw, 'an xsd:boolean', part);
}

function intOf(element: XElement, name: string, part: string): number | undefined {
  const raw = attributeValue(element, name);
  if (raw === undefined) return undefined;
  if (!/^[+-]?\d+$/.test(raw)) return attrError(element, name, raw, 'an xsd:int', part);
  return Number(raw);
}

/** `ST_Coordinate`: EMU, or a universal measure such as `108pt`, which PowerPoint honours. */
function coordinateOf(element: XElement, name: string, part: string): number | undefined {
  const raw = attributeValue(element, name);
  if (raw === undefined) return undefined;
  if (/^[+-]?\d+$/.test(raw)) return Number(raw);
  const measure = /^(-?\d+(?:\.\d+)?)(mm|cm|in|pt|pc|pi)$/.exec(raw);
  if (measure === null) return attrError(element, name, raw, 'a ST_Coordinate', part);
  return Math.round(Number(measure[1]) * UNITS[measure[2]!]!);
}

function required(
  element: XElement,
  name: string,
  value: number | undefined,
  part: string,
): number {
  if (value === undefined) {
    throw new ModelError(
      'MODEL_TABLE_ATTR',
      `${element.qname} has no @${name}, which the schema requires`,
      part,
      name,
    );
  }
  return value;
}

function enumOf<T extends string>(
  element: XElement,
  name: string,
  allowed: readonly T[],
  part: string,
): T | undefined {
  const raw = attributeValue(element, name);
  if (raw === undefined) return undefined;
  if (!(allowed as readonly string[]).includes(raw)) {
    return attrError(element, name, raw, `one of ${allowed.join(', ')}`, part);
  }
  return raw as T;
}

function lineChild(parent: XElement, qname: string, part: string) {
  const element = firstChild(parent, qname);
  return element === undefined ? undefined : parseLineElement(element, part);
}

function parseCellProps(element: XElement, part: string): TableCellProps {
  const borders: TableCellBorders = {
    left: lineChild(element, 'a:lnL', part),
    right: lineChild(element, 'a:lnR', part),
    top: lineChild(element, 'a:lnT', part),
    bottom: lineChild(element, 'a:lnB', part),
    tlToBr: lineChild(element, 'a:lnTlToBr', part),
    blToTr: lineChild(element, 'a:lnBlToTr', part),
  };
  const headers = firstChild(element, 'a:headers');
  return {
    marL: coordinateOf(element, 'marL', part),
    marR: coordinateOf(element, 'marR', part),
    marT: coordinateOf(element, 'marT', part),
    marB: coordinateOf(element, 'marB', part),
    vert: enumOf(element, 'vert', VERTICALS, part),
    anchor: enumOf(element, 'anchor', ANCHORS, part),
    anchorCtr: boolOf(element, 'anchorCtr', part),
    horzOverflow: enumOf(element, 'horzOverflow', HORZ_OVERFLOWS, part),
    borders,
    fill: parseFill(element, part),
    headers:
      headers === undefined
        ? []
        : childElements(headers)
            .filter((child) => child.qname === 'a:header')
            .map((child) => textContent(child)),
    node: element,
  };
}

function parseCell(element: XElement, part: string): TableCell {
  const props = firstChild(element, 'a:tcPr');
  return {
    gridSpan: intOf(element, 'gridSpan', part) ?? 1,
    rowSpan: intOf(element, 'rowSpan', part) ?? 1,
    hMerge: boolOf(element, 'hMerge', part) ?? false,
    vMerge: boolOf(element, 'vMerge', part) ?? false,
    id: attributeValue(element, 'id') ?? null,
    text: parseTextBodyChild(element, part),
    props: props === undefined ? undefined : parseCellProps(props, part),
    node: element,
  };
}

function parseRow(element: XElement, part: string): TableRow {
  return {
    h: required(element, 'h', coordinateOf(element, 'h', part), part),
    cells: childElements(element)
      .filter((child) => child.qname === 'a:tc')
      .map((child) => parseCell(child, part)),
    node: element,
  };
}

function parseColumn(element: XElement, part: string): TableColumn {
  return { w: required(element, 'w', coordinateOf(element, 'w', part), part), node: element };
}

function parseStyleRef(element: XElement): TableStyleRef | undefined {
  const id = firstChild(element, 'a:tableStyleId');
  if (id !== undefined) return { kind: 'id', id: textContent(id).trim() };
  const inline = firstChild(element, 'a:tableStyle');
  return inline === undefined ? undefined : { kind: 'inline', node: inline };
}

function parseTableProps(element: XElement, part: string): TableProps {
  const flag = (name: string): boolean => boolOf(element, name, part) ?? false;
  return {
    rtl: flag('rtl'),
    firstRow: flag('firstRow'),
    firstCol: flag('firstCol'),
    lastRow: flag('lastRow'),
    lastCol: flag('lastCol'),
    bandRow: flag('bandRow'),
    bandCol: flag('bandCol'),
    fill: parseFill(element, part),
    effects: parseEffects(element),
    style: parseStyleRef(element),
    node: element,
  };
}

/** An `a:tbl`. */
export function parseTable(element: XElement, part: string): Table {
  const props = firstChild(element, 'a:tblPr');
  const grid = firstChild(element, 'a:tblGrid');
  return {
    props: props === undefined ? undefined : parseTableProps(props, part),
    columns:
      grid === undefined
        ? []
        : childElements(grid)
            .filter((child) => child.qname === 'a:gridCol')
            .map((child) => parseColumn(child, part)),
    rows: childElements(element)
      .filter((child) => child.qname === 'a:tr')
      .map((child) => parseRow(child, part)),
    node: element,
  };
}

/** The table of a `p:graphicFrame`, when its `a:graphicData` carries one. */
export function parseTableChild(frame: XElement, part: string): Table | undefined {
  const graphic = firstChild(frame, 'a:graphic');
  const data = graphic === undefined ? undefined : firstChild(graphic, 'a:graphicData');
  if (data === undefined || attributeValue(data, 'uri') !== TABLE_URI) return undefined;
  const table = firstChild(data, 'a:tbl');
  return table === undefined ? undefined : parseTable(table, part);
}
