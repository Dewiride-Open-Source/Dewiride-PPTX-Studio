import { writeZip } from '../../../../ground-truth/lib/zip.ts';
import { DECLARATION } from '../markup/chassis.ts';

/**
 * A minimal SpreadsheetML workbook, for the OLE probe's payload.
 *
 * Experiment E6 measured what PowerPoint 16.0.20326 actually embeds for an OLE
 * worksheet: **`ppt/embeddings/Microsoft_Excel_Worksheet.xlsx`**, an ordinary
 * OPC package, through a `…/relationships/package` relationship, with a
 * `Default Extension="xlsx"`. Not a CFB compound file - that was route (a) and
 * it turned out not to be what a modern build writes.
 *
 * Which is convenient, because a package is something this repository can
 * author byte for byte and a CFB is not. Five parts, all of them XML, all of
 * them ours:
 *
 * ```
 * [Content_Types].xml
 * _rels/.rels                     -> xl/workbook.xml
 * xl/workbook.xml                 one sheet
 * xl/_rels/workbook.xml.rels      -> xl/worksheets/sheet1.xml
 * xl/worksheets/sheet1.xml        two rows of inline strings
 * ```
 *
 * The cells are `t="inlineStr"` rather than shared strings, so there is no
 * `sharedStrings.xml` and no index to keep consistent - the whole workbook is
 * readable in one pass. Excel opens it.
 *
 * Entries are **stored**, like every other Tier A byte, so the deck's hash does
 * not depend on which zlib built it. See `package.ts`'s file comment.
 */

const NS_SS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';

const CT_WORKBOOK = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml';
const CT_WORKSHEET = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';

const encoder = new TextEncoder();

/** One cell holding a literal string. `@r` is the A1 reference, and required. */
const cell = (reference: string, text: string): string =>
  `<c r="${reference}" t="inlineStr"><is><t>${text}</t></is></c>`;

const row = (number: number, cells: readonly string[]): string =>
  `<row r="${String(number)}">${cells.join('')}</row>`;

/**
 * The workbook `a26-ole` embeds.
 *
 * `rows` is a list of rows of cell text, starting at A1. Everything else is
 * fixed, because the payload's job is to be a real package, not to be varied.
 */
export function probeXlsx(rows: ReadonlyArray<readonly string[]>): Uint8Array {
  const contentTypes =
    DECLARATION +
    `<Types xmlns="${NS_CT}">` +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    `<Override PartName="/xl/workbook.xml" ContentType="${CT_WORKBOOK}"/>` +
    `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="${CT_WORKSHEET}"/>` +
    '</Types>';

  const packageRels =
    DECLARATION +
    `<Relationships xmlns="${NS_REL}">` +
    `<Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="xl/workbook.xml"/>` +
    '</Relationships>';

  const workbook =
    DECLARATION +
    `<workbook xmlns="${NS_SS}" xmlns:r="${NS_R}">` +
    '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>' +
    '</workbook>';

  const workbookRels =
    DECLARATION +
    `<Relationships xmlns="${NS_REL}">` +
    `<Relationship Id="rId1" Type="${NS_R}/worksheet" Target="worksheets/sheet1.xml"/>` +
    '</Relationships>';

  const sheet =
    DECLARATION +
    `<worksheet xmlns="${NS_SS}" xmlns:r="${NS_R}"><sheetData>` +
    rows
      .map((cells, index) =>
        row(
          index + 1,
          cells.map((text, column) =>
            cell(String.fromCharCode(65 + column) + String(index + 1), text),
          ),
        ),
      )
      .join('') +
    '</sheetData></worksheet>';

  return writeZip([
    { name: '[Content_Types].xml', bytes: encoder.encode(contentTypes), store: true },
    { name: '_rels/.rels', bytes: encoder.encode(packageRels), store: true },
    { name: 'xl/workbook.xml', bytes: encoder.encode(workbook), store: true },
    { name: 'xl/_rels/workbook.xml.rels', bytes: encoder.encode(workbookRels), store: true },
    { name: 'xl/worksheets/sheet1.xml', bytes: encoder.encode(sheet), store: true },
  ]);
}
