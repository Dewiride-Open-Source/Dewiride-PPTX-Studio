import { CONTENT_TYPE, isRelationshipPartName } from '@pptx-studio/opc';
import {
  attribute,
  childElements,
  descendantElements,
  namespaceOf,
  NS,
  type XElement,
} from '@pptx-studio/xml';
import type { Context } from '../context.js';
import { elementLocation } from '../report/location.js';

/**
 * `V013` … `V017`, `V030` and `V031`: children and attributes that are not optional.
 *
 * Every one of these is a `minOccurs` the schema states plainly, and every one
 * is broken the same way: by an editor deleting the last of something. The last
 * paragraph goes and a `p:txBody` is left with no `a:p`; a shape is dragged out
 * of a group and the group's `p:grpSpPr` is dropped with it. That is why these
 * are separate rules from the ordering ones even though a missing child and a
 * misplaced child are both "the sequence is wrong": the messages have to say
 * *add this*, not *move this*, and they have to fire on the parent rather than
 * on a child that is not there to point at.
 */

/** The twelve attributes of `CT_ColorMapping`. All required, no defaults. */
const CLR_MAP_ATTRIBUTES = [
  'bg1',
  'tx1',
  'bg2',
  'tx2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hlink',
  'folHlink',
] as const;

/** `p:presentation` carries `p:notesSz`. */
export function v013NotesSize(ctx: Context): void {
  for (const part of ctx.parts()) {
    const document = ctx.document(part);
    if (document === null) continue;
    const root = document.root;
    if (root.local !== 'presentation' || namespaceOf(root) !== NS.p) continue;

    if (childElements(root).some((child) => child.local === 'notesSz')) continue;
    ctx.add(
      'V013',
      elementLocation(part, root),
      'no <p:notesSz>. In CT_Presentation p:sldSz is [0..1] and p:notesSz is [1..1] - the ' +
        'asymmetry is the whole of this rule, and a generator that treats the two the same way ' +
        'writes a presentation part that is missing a required child.',
    );
  }
}

/** `p:clrMap` carries all twelve attributes. */
export function v014ColorMap(ctx: Context): void {
  forEachElement(ctx, (part, element) => {
    if (element.local !== 'clrMap' || namespaceOf(element) !== NS.p) return;
    const missing = CLR_MAP_ATTRIBUTES.filter((name) => attribute(element, name) === undefined);
    if (missing.length === 0) return;
    ctx.add(
      'V014',
      elementLocation(part, element),
      'the colour map is missing ' +
        missing.join(', ') +
        '. All twelve attributes of CT_ColorMapping are required and none has a default; a ' +
        'partial map is not a map with a gap, it is an invalid element - and it is the element ' +
        'that decides what bg1 and tx1 resolve to everywhere in the deck.',
    );
  });
}

/** `p:spTree` begins with `p:nvGrpSpPr` then `p:grpSpPr`. */
export function v015ShapeTreePrologue(ctx: Context): void {
  forEachElement(ctx, (part, element) => {
    if (element.local !== 'spTree' || namespaceOf(element) !== NS.p) return;
    const children = childElements(element);
    if (children[0]?.local !== 'nvGrpSpPr' || children[1]?.local !== 'grpSpPr') {
      ctx.add(
        'V015',
        elementLocation(part, element),
        'a shape tree begins with <nvGrpSpPr> then <grpSpPr>, in that order, before any shape. ' +
          'This one begins ' +
          (children.length === 0
            ? 'with nothing'
            : children
                .slice(0, 2)
                .map((child) => '<' + child.qname + '>')
                .join(' then ')) +
          '. A shape tree is a group shape, and CT_GroupShape requires both even when the slide ' +
          'is empty.',
      );
    }
  });
}

/** Every text body has `a:bodyPr` and at least one `a:p`. */
export function v016TextBody(ctx: Context): void {
  forEachElement(ctx, (part, element) => {
    if (element.local !== 'txBody') return;
    const children = childElements(element);
    if (!children.some((child) => child.local === 'bodyPr')) {
      ctx.add(
        'V016',
        elementLocation(part, element),
        '<' + element.qname + '> has no <a:bodyPr>. CT_TextBody requires it first.',
      );
    }
    if (!children.some((child) => child.local === 'p')) {
      ctx.add(
        'V016',
        elementLocation(part, element),
        '<' +
          element.qname +
          '> has no <a:p>. CT_TextBody is [1..unbounded] paragraphs, and an empty paragraph is ' +
          'not the same thing as no paragraph - a:endParaRPr is where an empty one gets its ' +
          'height. This is what deleting the last paragraph in an editor produces.',
      );
    }
  });
}

/** `p:graphicFrame` carries `p:xfrm` and `a:graphic`. */
export function v017GraphicFrame(ctx: Context): void {
  forEachElement(ctx, (part, element) => {
    if (element.local !== 'graphicFrame' || namespaceOf(element) !== NS.p) return;

    const xfrm = childElements(element).find((child) => child.local === 'xfrm');
    if (xfrm === undefined) {
      ctx.add(
        'V017',
        elementLocation(part, element),
        'no <p:xfrm>. Note the namespace: a graphic frame carries the PresentationML transform, ' +
          'not the a:xfrm every other shape uses. It has no placeholder-inherited geometry path ' +
          'either, which is why a rebound chart or table gets a copy of its layout transform ' +
          'rather than having its own deleted.',
      );
    } else if (namespaceOf(xfrm) !== NS.p) {
      ctx.add(
        'V017',
        elementLocation(part, xfrm),
        'the transform here is <' +
          xfrm.qname +
          '>, in the DrawingML namespace. CT_GraphicalObjectFrame takes p:xfrm; a:xfrm in its ' +
          'place is the same element name bound to the wrong schema.',
      );
    }

    // By local name and namespace, never by qname: `a:` is the conventional
    // prefix for DrawingML and not a guaranteed one, and a part that binds it
    // differently would silently stop being checked.
    const graphic = childElements(element).find(
      (child) => child.local === 'graphic' && namespaceOf(child) === NS.a,
    );
    if (graphic === undefined) {
      ctx.add(
        'V017',
        elementLocation(part, element),
        'no <a:graphic>. The frame is the wrapper; the graphic is the thing in it, and a frame ' +
          'without one names no table, chart, diagram or OLE object at all.',
      );
    }
  });
}

/** The lexical forms of `xsd:boolean`. */
const BOOLEANS = new Set(['1', '0', 'true', 'false']);
const INT = /^[+-]?\d+$/;

/** `a:gridCol` carries `@w`, `a:tr` carries `@h`, and a cell's four merge attributes are typed. */
export function v031TableAttributes(ctx: Context): void {
  forEachElement(ctx, (part, element) => {
    if (namespaceOf(element) !== NS.a) return;
    const missing = (name: string): void => {
      ctx.add(
        'V031',
        elementLocation(part, element),
        '<' +
          element.qname +
          '> has no @' +
          name +
          '. CT_TableCol and CT_TableRow require it, and C7 measured the repair prompt: ' +
          'PowerPoint opens the deck only after rewriting the part.',
      );
    };
    if (element.local === 'gridCol' && attribute(element, 'w') === undefined) missing('w');
    if (element.local === 'tr' && attribute(element, 'h') === undefined) missing('h');
    if (element.local !== 'tc') return;
    for (const [name, ok, type] of [
      ['gridSpan', INT, 'xsd:int'],
      ['rowSpan', INT, 'xsd:int'],
      ['hMerge', BOOLEANS, 'xsd:boolean'],
      ['vMerge', BOOLEANS, 'xsd:boolean'],
    ] as const) {
      const raw = attribute(element, name)?.value;
      if (raw === undefined || (ok instanceof RegExp ? ok.test(raw) : ok.has(raw))) continue;
      ctx.add(
        'V031',
        elementLocation(part, element),
        '<a:tc>/@' +
          name +
          ' is "' +
          raw +
          '", which is not an ' +
          type +
          '. C7 measured gridSpan="2.0" and hMerge="on": a repair prompt each; the span is ' +
          'dropped with the merge it named and the flag comes back as "1" under its span.',
      );
    }
  });
}

/** A grid position's owner under the rule C7 measured. */
interface Owner {
  readonly row: number;
  readonly col: number;
  rows: number;
  cols: number;
}

/** Zero is one and a negative runs to the edge: C7's span-zero and span-negative. */
function spanOf(raw: string | undefined): number {
  if (raw === undefined) return 1;
  const n = Number(raw);
  return n < 0 ? Number.POSITIVE_INFINITY : Math.max(1, n);
}

function flagged(cell: XElement, name: string): boolean {
  const raw = attribute(cell, name)?.value;
  return raw === '1' || raw === 'true';
}

/**
 * The occupancy `@pptx-studio/model` ships, over attributes: each cell takes the next column, a
 * claimed position is covered whatever it says, a span stops at the edge or at a claimed position.
 */
function owners(cells: readonly (readonly XElement[])[], cols: number): (Owner | null)[][] {
  const grid: (Owner | null)[][] = cells.map(() =>
    Array.from<Owner | null>({ length: cols }).fill(null),
  );
  cells.forEach((line, r) => {
    line.forEach((cell, c) => {
      if (c >= cols || grid[r]![c] !== null) return;
      const wide = spanOf(attribute(cell, 'gridSpan')?.value);
      const tall = spanOf(attribute(cell, 'rowSpan')?.value);
      const owner: Owner = { row: r, col: c, rows: 1, cols: 1 };
      while (c + owner.cols < cols && owner.cols < wide && grid[r]![c + owner.cols] === null) {
        owner.cols += 1;
      }
      owner.rows = Math.min(tall, cells.length - r);
      for (let rr = r; rr < r + owner.rows; rr++) {
        for (let cc = c; cc < c + owner.cols; cc++) grid[rr]![cc] = owner;
      }
    });
  });
  return grid;
}

/** The table PowerPoint reads: one cell per column in every row, a flag exactly where a span covers. */
export function v030TableGrid(ctx: Context): void {
  forEachElement(ctx, (part, element) => {
    if (element.local !== 'tbl' || namespaceOf(element) !== NS.a) return;
    const add = (where: XElement, message: string): void => {
      ctx.add('V030', elementLocation(part, where), message);
    };
    const grid = childElements(element).find((child) => child.local === 'tblGrid');
    const cols =
      grid === undefined ? 0 : childElements(grid).filter((c) => c.local === 'gridCol').length;
    const rows = childElements(element).filter((child) => child.local === 'tr');
    if (cols === 0 || rows.length === 0) {
      add(
        element,
        'the table has ' +
          String(cols) +
          ' column(s) and ' +
          String(rows.length) +
          ' row(s). PowerPoint invents one of whichever is missing, empty, and writes it back - ' +
          'measured in C7.',
      );
      return;
    }
    const cells = rows.map((row) => childElements(row).filter((child) => child.local === 'tc'));
    cells.forEach((line, r) => {
      if (line.length === cols) return;
      add(
        rows[r]!,
        'row ' +
          String(r + 1) +
          ' holds ' +
          String(line.length) +
          ' cell(s) for ' +
          String(cols) +
          ' column(s). PowerPoint pads a short row with an empty cell, drops the cells past the ' +
          'last column, and writes the result back - measured in C7.',
      );
    });
    const grid2 = owners(cells, cols);
    cells.forEach((line, r) => {
      line.forEach((cell, c) => {
        const owner = grid2[r]?.[c];
        if (owner === null || owner === undefined) return;
        const at = 'cell (' + String(r + 1) + ', ' + String(c + 1) + ')';
        if (owner.row === r && owner.col === c) {
          for (const [name, span] of [
            ['gridSpan', owner.cols],
            ['rowSpan', owner.rows],
          ] as const) {
            const written = attribute(cell, name)?.value;
            if (written === undefined || Number(written) === span) continue;
            add(
              cell,
              at +
                ' says ' +
                name +
                '="' +
                written +
                '" and covers ' +
                String(span) +
                '. A span stops at the grid edge or at a position an earlier span claimed, zero ' +
                'is one and a negative runs to the edge; PowerPoint writes the span it drew - ' +
                'measured in C7.',
            );
          }
          for (const name of ['hMerge', 'vMerge'] as const) {
            if (!flagged(cell, name)) continue;
            add(
              cell,
              at +
                ' says ' +
                name +
                '="1" and no span covers it. The flag changes nothing: PowerPoint reads the ' +
                'spans alone, 76 of 76 in C7, and drops the flag on save.',
            );
          }
          return;
        }
        for (const [name, covered] of [
          ['hMerge', owner.col < c],
          ['vMerge', owner.row < r],
        ] as const) {
          if (covered === flagged(cell, name)) continue;
          if (!covered) {
            add(
              cell,
              at +
                ' says ' +
                name +
                '="1" and the span of cell (' +
                String(owner.row + 1) +
                ', ' +
                String(owner.col + 1) +
                ') covers it the other way. PowerPoint drops the flag on save - measured in C7.',
            );
            continue;
          }
          add(
            cell,
            at +
              ' is covered by the span of cell (' +
              String(owner.row + 1) +
              ', ' +
              String(owner.col + 1) +
              ') and does not say ' +
              name +
              '="1". PowerPoint reads it as covered anyway and adds the flag on save; a reader ' +
              'that trusts the flag draws a cell PowerPoint does not - measured in C7.',
          );
        }
        // PowerPoint's own form repeats the anchor's rowSpan along its row and its gridSpan
        // down its column; any other span on a covered cell is deleted on save.
        for (const [name, allowed] of [
          ['gridSpan', c === owner.col ? owner.cols : null],
          ['rowSpan', r === owner.row ? owner.rows : null],
        ] as const) {
          const written = attribute(cell, name)?.value;
          if (written === undefined || Number(written) === allowed) continue;
          add(
            cell,
            at +
              ' is covered and says ' +
              name +
              '="' +
              written +
              '". A covered cell’s own span is ignored and PowerPoint deletes it on save unless ' +
              'it repeats the anchor’s other span, as PowerPoint’s own form does - measured in C7.',
          );
        }
      });
    });
  });
}

/**
 * Presentation-family parts only.
 *
 * A `.pptx` carries whole foreign documents inside it - an embedded workbook
 * behind every chart, a Word document behind some OLE objects - and their parts
 * are XML with elements called `txBody` and `graphicFrame` in namespaces these
 * rules do not govern. Checking them would be applying PresentationML's
 * cardinalities to SpreadsheetML's schema, which is why every rule above tests
 * the namespace and not only the local name, and why the walk itself stops at
 * parts whose content type belongs to another format.
 */
const FOREIGN_CONTENT_TYPES = new Set<string>([
  CONTENT_TYPE.spreadsheet,
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
]);

function forEachElement(ctx: Context, visit: (part: string, element: XElement) => void): void {
  for (const part of ctx.parts()) {
    if (isRelationshipPartName(part)) continue;
    const contentType = ctx.contentType(part);
    if (contentType !== undefined && FOREIGN_CONTENT_TYPES.has(contentType)) continue;
    const document = ctx.document(part);
    if (document === null) continue;
    for (const element of descendantElements(document.root)) visit(part, element);
  }
}

/** Exported for `refusal-rules.ts`, which walks the same parts for the same reason. */
export { forEachElement, FOREIGN_CONTENT_TYPES };
