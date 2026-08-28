import { placeholderXml, TITLE_BOX, type ProbeLayout } from '../package.ts';
import { grid, scheme, shape, solidFill } from '../shapes.ts';
import { field, para, run, textLine, txBody } from '../text.ts';
import type { ProbeDeck } from '../types.ts';

/**
 * `a:fld`: every reserved type, its GUID, and its cached text.
 *
 * ## The count is fifteen, not fourteen
 *
 * The approved plan says "the 14 reserved `a:fld` types". ECMA-376 §21.1.2.2.4
 * lists **fifteen**: `slidenum`, `datetime`, and `datetime1` through
 * `datetime13`. The plan is off by one and this deck carries the real set, plus
 * one value that is not reserved at all.
 *
 * That extra one is `datetimeFigureOut`, which PowerPoint itself writes when a
 * date placeholder is set to update automatically in the locale's own format.
 * `@type` is `xsd:string` with no enumeration behind it, so an unreserved value
 * is not an error - and a renderer that switches on a closed set of fifteen
 * shows nothing where PowerPoint shows today's date. The generic rule is to
 * fall back to the cached `a:t`, and having a real-world unreserved value in
 * the corpus is what makes that rule testable rather than aspirational.
 *
 * ## `@id` is not decorative
 *
 * It is an `ST_Guid`, required, generated once when the field is created and
 * persisting unchanged for the life of the file. Regenerating one on export is
 * a documented route to a repair prompt, so every GUID here is a literal in
 * this module and `cli roundtrip` can assert they survived.
 *
 * ## `a:t` is the cache, and it is what renders when nothing else can
 *
 * A field's text child holds what the authoring application last drew. A viewer
 * that cannot compute the field - an unknown type, a locale it has no data for,
 * a date format it has not implemented - falls back to it. So every field here
 * carries plausible cached text, and a field with an empty cache would render
 * as nothing at all, which is the failure this deck exists to make visible.
 *
 * ## Where fields actually live
 *
 * Slide 3 puts them where PowerPoint puts them: inside the `dt`, `ftr` and
 * `sldNum` placeholders, bound through `p:hf`. A field in an ordinary text box
 * is legal and slides 1 and 2 use that form for density, but the placeholder
 * case is the one every real deck has.
 */

const BOX_LINE = '<a:ln w="9525"><a:solidFill>' + scheme('tx1') + '</a:solidFill></a:ln>';
const BOX_BODY_PR = '<a:bodyPr wrap="square"><a:normAutofit/></a:bodyPr>';

interface FieldRow {
  readonly type: string;
  /** The format ECMA documents for this type, shown beside it on the slide. */
  readonly format: string;
  /** What the authoring application last drew - and the only fallback there is. */
  readonly cached: string;
  readonly guid: string;
}

/**
 * The fifteen reserved types, their documented format, and cached text in that
 * format for the same instant: Friday 12 October 2007, 16:28:34.
 *
 * The instant is the one ECMA's own examples use, which makes the table
 * checkable against the specification line by line rather than against taste.
 */
const RESERVED_FIELDS: readonly FieldRow[] = [
  {
    type: 'slidenum',
    format: 'presentation slide number',
    cached: '1',
    guid: '{424CEEAC-8F67-4238-9622-1B74DC6E8318}',
  },
  {
    type: 'datetime',
    format: "the rendering application's default",
    cached: '12/10/2007',
    guid: '{5A2B8C61-0F4D-4E7A-9B3C-2D6E8F1A4B70}',
  },
  {
    type: 'datetime1',
    format: 'MM/DD/YYYY',
    cached: '10/12/2007',
    guid: '{7C3D9E12-5A6B-4C8D-9E0F-1A2B3C4D5E61}',
  },
  {
    type: 'datetime2',
    format: 'Day, Month DD, YYYY',
    cached: 'Friday, October 12, 2007',
    guid: '{9E5F1A23-6B7C-4D9E-8F01-2B3C4D5E6F72}',
  },
  {
    type: 'datetime3',
    format: 'DD Month YYYY',
    cached: '12 October 2007',
    guid: '{1A6B2C34-7C8D-4E0F-9012-3C4D5E6F7083}',
  },
  {
    type: 'datetime4',
    format: 'Month DD, YYYY',
    cached: 'October 12, 2007',
    guid: '{2B7C3D45-8D9E-4F10-A123-4D5E6F708194}',
  },
  {
    type: 'datetime5',
    format: 'DD-Mon-YY',
    cached: '12-Oct-07',
    guid: '{3C8D4E56-9E0F-4021-B234-5E6F708192A5}',
  },
  {
    type: 'datetime6',
    format: 'Month YY',
    cached: 'October 07',
    guid: '{4D9E5F67-0F10-4132-C345-6F70819203B6}',
  },
  {
    type: 'datetime7',
    format: 'Mon-YY',
    cached: 'Oct-07',
    guid: '{5E0F6078-1021-4243-D456-7081920314C7}',
  },
  {
    type: 'datetime8',
    format: 'MM/DD/YYYY hh:mm AM/PM',
    cached: '10/12/2007 4:28 PM',
    guid: '{6F107189-2132-4354-E567-8192031425D8}',
  },
  {
    type: 'datetime9',
    format: 'MM/DD/YYYY hh:mm:ss AM/PM',
    cached: '10/12/2007 4:28:34 PM',
    guid: '{7021829A-3243-4465-F678-92031425364E}',
  },
  {
    type: 'datetime10',
    format: 'hh:mm',
    cached: '16:28',
    guid: '{813293AB-4354-4576-0789-031425364759}',
  },
  {
    type: 'datetime11',
    format: 'hh:mm:ss',
    cached: '16:28:34',
    guid: '{9243A4BC-5465-4687-189A-14253647586A}',
  },
  {
    type: 'datetime12',
    format: 'hh:mm AM/PM',
    cached: '4:28 PM',
    guid: '{A354B5CD-6576-4798-29AB-25364758697B}',
  },
  {
    type: 'datetime13',
    format: 'hh:mm:ss AM/PM',
    cached: '4:28:34 PM',
    guid: '{B465C6DE-7687-48A9-3ABC-36475869798C}',
  },
];

/** Not reserved, and written by PowerPoint anyway. */
const FIGURE_OUT: FieldRow = {
  type: 'datetimeFigureOut',
  format: 'not in ECMA-376; PowerPoint writes it for "update automatically"',
  cached: '12/10/2007',
  guid: '{C576D7EF-8798-49BA-4BCD-475869798A9D}',
};

function fieldRow(entry: FieldRow): string {
  return para({
    content:
      run(entry.type + ' — ', { lang: 'en-GB', sz: 1000, fill: solidFill(scheme('tx1')) }) +
      field({
        id: entry.guid,
        type: entry.type,
        text: entry.cached,
        props: { lang: 'en-GB', sz: 1000, b: true, fill: solidFill(scheme('accent1')) },
      }) +
      run(' · ' + entry.format, {
        lang: 'en-GB',
        sz: 900,
        i: true,
        fill: solidFill(scheme('accent3')),
      }),
  });
}

function fieldColumns(entries: readonly FieldRow[], columns: number, firstId: number): string {
  const cell = grid(columns, 1);
  const perColumn = Math.ceil(entries.length / columns);
  const boxes: string[] = [];
  for (let column = 0; column < columns; column++) {
    const slice = entries.slice(column * perColumn, (column + 1) * perColumn);
    if (slice.length === 0) continue;
    boxes.push(
      shape({
        id: firstId + column,
        name: 'Fields ' + String(column + 1),
        ...cell(column),
        line: BOX_LINE,
        textBody: txBody({ bodyPr: BOX_BODY_PR, paras: slice.map(fieldRow).join('') }),
      }),
    );
  }
  return boxes.join('');
}

// ---------------------------------------------------------------- the layout

const HF_ROW_Y = 6356350;
const HF_ROW_HEIGHT = 365125;

const LAYOUT_HF_PLACEHOLDERS =
  placeholderXml({
    id: 3,
    name: 'Date Placeholder 2',
    type: 'dt',
    idx: 10,
    x: 457200,
    y: HF_ROW_Y,
    cx: 2743200,
    cy: HF_ROW_HEIGHT,
  }) +
  placeholderXml({
    id: 4,
    name: 'Footer Placeholder 3',
    type: 'ftr',
    idx: 11,
    x: 4038600,
    y: HF_ROW_Y,
    cx: 4114800,
    cy: HF_ROW_HEIGHT,
  }) +
  placeholderXml({
    id: 5,
    name: 'Slide Number Placeholder 4',
    type: 'sldNum',
    idx: 12,
    x: 8610600,
    y: HF_ROW_Y,
    cx: 2743200,
    cy: HF_ROW_HEIGHT,
  });

const LAYOUTS: readonly ProbeLayout[] = [
  {
    type: 'cust',
    name: 'Fields',
    hasTitle: true,
    shapes:
      placeholderXml({ id: 2, name: 'Title 1', type: 'title', ...TITLE_BOX }) +
      LAYOUT_HF_PLACEHOLDERS,
    hf: '<p:hf hdr="0"/>',
  },
];

export const a09Fields: ProbeDeck = {
  id: 'a09-fields',
  title: 'PPTX Studio corpus: a09 fields',
  description:
    'All fifteen reserved a:fld types - slidenum, datetime and datetime1 through datetime13, ' +
    'which is one more than the plan says - each with a literal ST_Guid and cached text for one ' +
    'fixed instant, plus the unreserved datetimeFigureOut PowerPoint itself writes. Slide 3 puts ' +
    'the date, footer and slide-number fields where real decks keep them: inside p:hf placeholders.',
  features: {
    shape: 17,
    placeholder: 12,
    presetGeom: 5,
    field: 18,
    gradientFill: 2,
  },
  build: () => ({
    title: 'PPTX Studio corpus: a09 fields',
    layouts: LAYOUTS,
    slides: [
      {
        title: 'a09 — slidenum and datetime1 to datetime7',
        layout: 0,
        body: fieldColumns(RESERVED_FIELDS.slice(0, 8), 2, 10),
      },
      {
        title: 'a09 — datetime8 to datetime13, and one that is not reserved',
        layout: 0,
        body: fieldColumns([...RESERVED_FIELDS.slice(8), FIGURE_OUT], 2, 10),
      },
      {
        title: 'a09 — where fields actually live',
        layout: 0,
        body:
          shape({
            id: 10,
            name: 'What is below',
            ...grid(1, 3)(0),
            line: BOX_LINE,
            textBody: txBody({
              bodyPr: BOX_BODY_PR,
              paras:
                textLine(
                  'The three placeholders along the bottom edge are bound through p:hf, and each ' +
                    'holds a field rather than literal text.',
                  { lang: 'en-GB', sz: 1400, fill: solidFill(scheme('tx1')) },
                ) +
                textLine(
                  'A viewer that cannot compute a field falls back to its cached a:t, which is ' +
                    'why an empty cache renders as nothing at all.',
                  { lang: 'en-GB', sz: 1400, fill: solidFill(scheme('tx1')) },
                ),
            }),
          }) +
          placeholderXml({
            id: 11,
            name: 'Date Placeholder 2',
            type: 'dt',
            idx: 10,
            body: para({
              content: field({
                id: '{D687E8F0-98A9-40CB-5CDE-58697A8B9BAE}',
                type: 'datetimeFigureOut',
                text: '27/08/2026',
                props: { lang: 'en-GB', smtClean: false },
              }),
              endProps: { lang: 'en-GB' },
            }),
          }) +
          placeholderXml({
            id: 12,
            name: 'Footer Placeholder 3',
            type: 'ftr',
            idx: 11,
            body: textLine('PPTX Studio corpus — a09', { lang: 'en-GB' }),
          }) +
          placeholderXml({
            id: 13,
            name: 'Slide Number Placeholder 4',
            type: 'sldNum',
            idx: 12,
            body: para({
              content: field({
                id: '{E798F901-A9BA-41DC-6DEF-697A8B9CACBF}',
                type: 'slidenum',
                text: '3',
                props: { lang: 'en-GB', smtClean: false },
              }),
              endProps: { lang: 'en-GB' },
            }),
          }),
      },
    ],
  }),
};
