/**
 * `a:tbl` into a `Table`, and `CT_TableStyle` into a `TableStyle`. Nothing here resolves the
 * merges or the style; `tableGrid` and `tableStyleOf` do, by the rules C7 and C8 measured.
 */

import { attributeValue, childElements, firstChild, textContent } from '@pptx-studio/xml';
import type { Fill, Line } from '@pptx-studio/paint';
import type { XElement } from '@pptx-studio/xml';

import { ModelError } from '../errors.js';
import { TABLE_URI } from '../table.js';
import type {
  OnOffStyle,
  Table,
  TableBackground,
  TableCell,
  TableCellBorders,
  TableCellProps,
  TableColumn,
  TableProps,
  TableRow,
  TableStyle,
  TableStyleBorders,
  TableStyleCell,
  TableStylePart,
  TableStylePartName,
  TableStyleRef,
  TableStyleText,
  Themeable,
} from '../table.js';
import type { HorzOverflow, TextAnchor, VerticalText } from '../text.js';
import type { FontCollection, FontRef, StyleRef } from '../types.js';
import { parseColorChild, parseEffects, parseFill, parseLineElement } from './paint.js';
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

/** `ST_Guid` in any case: C8 measured lower case matching, and braces or padding as a repair. */
const GUID = /^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$/i;

function guidOf(element: XElement, raw: string, where: string, part: string): string {
  if (!GUID.test(raw)) {
    throw new ModelError(
      'MODEL_TABLE_ATTR',
      `${where} is "${raw}", which is not an ST_Guid`,
      part,
      element.qname,
    );
  }
  return raw;
}

function requiredText(element: XElement, name: string, part: string): string {
  const raw = attributeValue(element, name);
  if (raw === undefined) {
    throw new ModelError(
      'MODEL_TABLE_ATTR',
      `${element.qname} has no @${name}, which the schema requires`,
      part,
      name,
    );
  }
  return raw;
}

function parseStyleRef(element: XElement, part: string): TableStyleRef | undefined {
  const id = firstChild(element, 'a:tableStyleId');
  if (id !== undefined)
    return { kind: 'id', id: guidOf(id, textContent(id), 'a:tableStyleId', part) };
  const inline = firstChild(element, 'a:tableStyle');
  if (inline === undefined) return undefined;
  const styleId = requiredText(inline, 'styleId', part);
  return {
    kind: 'inline',
    id: guidOf(inline, styleId, 'a:tableStyle/@styleId', part),
    node: inline,
  };
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
    style: parseStyleRef(element, part),
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

/* -------------------------------------------------------------------------- */
/* table styles                                                               */
/* -------------------------------------------------------------------------- */

const PART_NAMES: readonly TableStylePartName[] = [
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
];
const ON_OFF: readonly OnOffStyle[] = ['on', 'off', 'def'];
const FONT_COLLECTIONS: readonly FontRef['idx'][] = ['major', 'minor', 'none'];

function styleError(element: XElement, why: string, part: string): never {
  throw new ModelError('MODEL_TABLE_STYLE', `${element.qname} ${why}`, part, element.qname);
}

/** `CT_StyleMatrixReference`, whose `@idx` the schema requires. */
function matrixRefOf(element: XElement, part: string): StyleRef {
  const raw = requiredText(element, 'idx', part);
  if (!/^\d+$/.test(raw)) {
    throw new ModelError('MODEL_STYLE_IDX', `${element.qname}/@idx is "${raw}"`, part, raw);
  }
  return { idx: Number(raw), color: parseColorChild(element) };
}

function fontRefOf(element: XElement, part: string): FontRef {
  const raw = requiredText(element, 'idx', part);
  if (!(FONT_COLLECTIONS as readonly string[]).includes(raw)) {
    throw new ModelError(
      'MODEL_FONT_COLLECTION',
      `a:fontRef/@idx is "${raw}", not major, minor or none`,
      part,
      raw,
    );
  }
  return { idx: raw as FontRef['idx'], color: parseColorChild(element) };
}

function fontCollectionOf(element: XElement): FontCollection {
  const face = (script: string): string | null => {
    const child = firstChild(element, script);
    return child === undefined ? null : (attributeValue(child, 'typeface') ?? null);
  };
  return { latin: face('a:latin'), ea: face('a:ea'), cs: face('a:cs') };
}

/** `a:fill` or `a:fillRef`; an `a:fill` holding no fill is a repair (C8). */
function themeableFill(parent: XElement, part: string): Themeable<Fill> | undefined {
  const fill = firstChild(parent, 'a:fill');
  if (fill !== undefined) {
    const value = parseFill(fill, part);
    return value === undefined ? styleError(fill, 'holds no fill', part) : { kind: 'value', value };
  }
  const ref = firstChild(parent, 'a:fillRef');
  return ref === undefined ? undefined : { kind: 'ref', ref: matrixRefOf(ref, part) };
}

/** One edge of `a:tcBdr`; an edge with neither `a:ln` nor `a:lnRef` is a repair (C8). */
function edgeOf(borders: XElement, qname: string, part: string): Themeable<Line> | undefined {
  const edge = firstChild(borders, qname);
  if (edge === undefined) return undefined;
  const ln = firstChild(edge, 'a:ln');
  if (ln !== undefined) return { kind: 'value', value: parseLineElement(ln, part) };
  const ref = firstChild(edge, 'a:lnRef');
  if (ref !== undefined) return { kind: 'ref', ref: matrixRefOf(ref, part) };
  return styleError(edge, 'has neither a:ln nor a:lnRef', part);
}

function parseTextStyle(element: XElement, part: string): TableStyleText {
  const fontRef = firstChild(element, 'a:fontRef');
  const font = firstChild(element, 'a:font');
  return {
    b: enumOf(element, 'b', ON_OFF, part),
    i: enumOf(element, 'i', ON_OFF, part),
    font:
      fontRef !== undefined
        ? { kind: 'ref', ref: fontRefOf(fontRef, part) }
        : font === undefined
          ? undefined
          : { kind: 'value', value: fontCollectionOf(font) },
    color: parseColorChild(element) ?? undefined,
  };
}

function parseCellStyle(element: XElement, part: string): TableStyleCell {
  const borders = firstChild(element, 'a:tcBdr');
  const edges: TableStyleBorders | undefined =
    borders === undefined
      ? undefined
      : {
          left: edgeOf(borders, 'a:left', part),
          right: edgeOf(borders, 'a:right', part),
          top: edgeOf(borders, 'a:top', part),
          bottom: edgeOf(borders, 'a:bottom', part),
          insideH: edgeOf(borders, 'a:insideH', part),
          insideV: edgeOf(borders, 'a:insideV', part),
          tlToBr: edgeOf(borders, 'a:tl2br', part),
          blToTr: edgeOf(borders, 'a:tr2bl', part),
        };
  return {
    borders: edges,
    fill: themeableFill(element, part),
    cell3D: firstChild(element, 'a:cell3D'),
  };
}

function parseBackground(element: XElement, part: string): TableBackground {
  const effect = firstChild(element, 'a:effect');
  const effectRef = firstChild(element, 'a:effectRef');
  return {
    fill: themeableFill(element, part),
    effect:
      effect !== undefined
        ? { kind: 'value', value: parseEffects(effect) }
        : effectRef === undefined
          ? undefined
          : { kind: 'ref', ref: matrixRefOf(effectRef, part) },
  };
}

/** An `a:tblStyle` or an inline `a:tableStyle`: `CT_TableStyle`. */
export function parseTableStyle(element: XElement, part: string): TableStyle {
  const id = guidOf(
    element,
    requiredText(element, 'styleId', part),
    `${element.qname}/@styleId`,
    part,
  );
  const background = firstChild(element, 'a:tblBg');
  const parts: Partial<Record<TableStylePartName, TableStylePart>> = {};
  for (const name of PART_NAMES) {
    const child = firstChild(element, `a:${name}`);
    if (child === undefined) continue;
    const text = firstChild(child, 'a:tcTxStyle');
    const cell = firstChild(child, 'a:tcStyle');
    parts[name] = {
      text: text === undefined ? undefined : parseTextStyle(text, part),
      cell: cell === undefined ? undefined : parseCellStyle(cell, part),
    };
  }
  return {
    id,
    name: requiredText(element, 'styleName', part),
    background: background === undefined ? undefined : parseBackground(background, part),
    parts,
    node: element,
  };
}
