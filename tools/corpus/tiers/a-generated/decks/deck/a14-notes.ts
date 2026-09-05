import { placeholderXml, type ProbeSheetMaster } from '../package.ts';
import { shape } from '../shapes.ts';
import { field, para, textLine } from '../text.ts';
import type { ProbeDeck } from '../types.ts';

/**
 * The notes family: a notes master, three notes slides, and a handout master.
 *
 * ## Where the geometry comes from
 *
 * Not from the schema, and not from memory. On 2026-08-27 a deck was authored
 * through PowerPoint COM - one blank slide, one line of speaker notes, and the
 * handout master's header switched on - and saved. Every placeholder below is
 * the type, `@sz`, `@idx` and box PowerPoint 16.0.20326 wrote, in the order it
 * wrote them. Two facts fell out that no reading of the specification produces:
 *
 * **`p:hf` is absent by default.** PowerPoint writes no `p:hf` anywhere until a
 * user turns something on. This deck writes one on each sheet on purpose, so
 * there is a fixture where the element exists.
 *
 * **There is no handout master until somebody edits one.** A deck with notes
 * gets `notesMaster1.xml` and `theme2.xml`; the handout master appears only
 * once the handout master view has been changed. So most real decks have a
 * notes master and no handout master, and a reader that assumes the pair is
 * either both-or-neither is wrong about the common case.
 *
 * ## `hdr` and `sldImg`, finally somewhere legal
 *
 * `a02-placeholders` covers fourteen of the sixteen `ST_PlaceholderType`
 * values. The other two - `hdr` and `sldImg` - were bisected out of it because
 * either one alone, on a slide layout or on a slide, is a **whole-package
 * refusal** with no part name and no element name in the message. Nothing in
 * the schema says so: `CT_Placeholder` is one complex type shared by masters,
 * layouts, slides, notes slides and handout masters, and `ST_PlaceholderType`
 * is one enumeration holding all sixteen.
 *
 * Here they are legal and PowerPoint writes them itself. The notes master has
 * both; the handout master has `hdr`; notes slide 2 has both again. That closes
 * the enumeration at sixteen of sixteen across `a02` and this deck, and it
 * pins down where the line actually falls - by sheet family, not by value.
 *
 * ## The idx that does not match, and why nothing is wrong
 *
 * On the notes **master** the notes body is `idx="3"`. On the notes **slide**
 * PowerPoint writes `idx="1"`, and the slide image carries no `@idx` at all
 * against a master that says `idx="2"`. Both were measured from the same file.
 *
 * So the notes-slide-to-notes-master hop does **not** match on `(type, idx)`.
 * It is the same shape as sub-phase 7.1's rule that a layout binds to a master
 * on type alone, and it is the fourth and third tiers of the matcher doing the
 * work: the sole `body` wins by being the sole body, and `sldImg` wins on raw
 * type. A matcher that requires `idx` to agree orphans the notes text of every
 * deck PowerPoint has ever written.
 *
 * ## What a notes slide is not
 *
 * `CT_NotesSlide` is `cSld, clrMapOvr, extLst`. No `p:transition`, no
 * `p:timing`, no `p:hf`. And it relates **both** ways - forward to the notes
 * master and back to the slide it annotates - so the part number is not what
 * binds a notes page to its slide.
 */

const NOTES_WIDTH = 6858000;
const HALF = 2971800;
const BAR_HEIGHT = 458788;
const RIGHT_COLUMN = 3884613;
const BOTTOM = 8685213;

/** PowerPoint's own notes-master geometry, four boxes at the page corners. */
function cornerPlaceholders(spec: {
  readonly dtIdx: number;
  readonly ftrIdx: number;
  readonly sldNumIdx: number;
  readonly firstId: number;
  /** `@sz` on the date placeholder. The notes master omits it; the handout does not. */
  readonly dateSize?: string;
  readonly dateGuid: string;
  readonly slideNumberGuid: string;
}): {
  readonly header: string;
  readonly date: string;
  readonly footer: string;
  readonly slideNumber: string;
} {
  const smallText =
    '<a:lstStyle><a:lvl1pPr algn="l"><a:defRPr sz="1200"/></a:lvl1pPr></a:lstStyle>';
  const rightText =
    '<a:lstStyle><a:lvl1pPr algn="r"><a:defRPr sz="1200"/></a:lvl1pPr></a:lstStyle>';
  return {
    header: placeholderXml({
      id: spec.firstId,
      name: 'Header Placeholder 1',
      type: 'hdr',
      size: 'quarter',
      x: 0,
      y: 0,
      cx: HALF,
      cy: BAR_HEIGHT,
      lstStyle: smallText,
    }),
    date: placeholderXml({
      id: spec.firstId + 1,
      name: 'Date Placeholder 2',
      type: 'dt',
      ...(spec.dateSize === undefined ? {} : { size: spec.dateSize }),
      idx: spec.dtIdx,
      x: RIGHT_COLUMN,
      y: 0,
      cx: HALF,
      cy: BAR_HEIGHT,
      lstStyle: rightText,
      body: para({
        content: field({
          id: spec.dateGuid,
          // Not one of the fifteen reserved types. PowerPoint writes it anyway,
          // on every date placeholder it creates - see `a09-fields`.
          type: 'datetimeFigureOut',
          text: '27/08/2026',
          props: { lang: 'en-GB', smtClean: false },
        }),
        endProps: { lang: 'en-GB' },
      }),
    }),
    footer: placeholderXml({
      id: spec.firstId + 2,
      name: 'Footer Placeholder 3',
      type: 'ftr',
      size: 'quarter',
      idx: spec.ftrIdx,
      x: 0,
      y: BOTTOM,
      cx: HALF,
      cy: BAR_HEIGHT - 1,
      lstStyle: smallText,
    }),
    slideNumber: placeholderXml({
      id: spec.firstId + 3,
      name: 'Slide Number Placeholder 4',
      type: 'sldNum',
      size: 'quarter',
      idx: spec.sldNumIdx,
      x: RIGHT_COLUMN,
      y: BOTTOM,
      cx: HALF,
      cy: BAR_HEIGHT - 1,
      lstStyle: rightText,
      body: para({
        content: field({
          id: spec.slideNumberGuid,
          type: 'slidenum',
          // What PowerPoint caches in a slide-number placeholder on a master:
          // the single-character placeholder glyph, not a number.
          text: '‹#›',
          props: { lang: 'en-GB', smtClean: false },
        }),
        endProps: { lang: 'en-GB' },
      }),
    }),
  };
}

const NOTES_CORNERS = cornerPlaceholders({
  dtIdx: 1,
  ftrIdx: 4,
  sldNumIdx: 5,
  firstId: 2,
  dateGuid: '{3A1F7C08-5E62-4D91-B4A7-0C93E518D264}',
  slideNumberGuid: '{6B4E29D1-7F30-4A85-9C12-58D0A7E36F49}',
});

const NOTES_MASTER: ProbeSheetMaster = {
  hf: '<p:hf hdr="1" ftr="1" dt="1" sldNum="1"/>',
  shapes:
    NOTES_CORNERS.header +
    NOTES_CORNERS.date +
    // `noRot` and `noChangeAspect` as well as `noGrp`: the slide image is a
    // rendering of another sheet and PowerPoint will not let it be reshaped.
    placeholderXml({
      id: 6,
      name: 'Slide Image Placeholder 5',
      type: 'sldImg',
      idx: 2,
      x: 685800,
      y: 1143000,
      cx: 5486400,
      cy: 3086100,
      locks: 'noGrp="1" noRot="1" noChangeAspect="1"',
    }) +
    placeholderXml({
      id: 7,
      name: 'Notes Placeholder 6',
      type: 'body',
      size: 'quarter',
      idx: 3,
      x: 685800,
      y: 4400550,
      cx: 5486400,
      cy: 3600450,
      body:
        textLine('Click to edit Master text styles', { lang: 'en-GB' }) +
        textLine('Second level', { lang: 'en-GB' }, { lvl: 1 }),
    }) +
    NOTES_CORNERS.footer +
    NOTES_CORNERS.slideNumber,
};

const HANDOUT_CORNERS = cornerPlaceholders({
  dtIdx: 1,
  ftrIdx: 2,
  sldNumIdx: 3,
  firstId: 2,
  dateSize: 'quarter',
  dateGuid: '{9D27F4B6-1A83-4E70-8F55-2C6B0E419A73}',
  slideNumberGuid: '{C8E15A30-46D9-4B27-A0F3-71E5D2860C14}',
});

/**
 * The handout master: the same four corner boxes, and no slide image and no
 * body. A handout shows several slides at once, so there is nothing for a
 * single `sldImg` to be, and `CT_HandoutMaster` has no text styles at all.
 */
const HANDOUT_MASTER: ProbeSheetMaster = {
  hf: '<p:hf hdr="1" ftr="0" dt="1" sldNum="1"/>',
  shapes:
    HANDOUT_CORNERS.header +
    HANDOUT_CORNERS.date +
    HANDOUT_CORNERS.footer +
    HANDOUT_CORNERS.slideNumber,
};

/** A notes-slide placeholder: no `a:xfrm`, because it inherits the master's. */
const notesPlaceholder = (spec: {
  id: number;
  name: string;
  type: string;
  idx?: number;
  size?: string;
  body?: string;
  locks?: string;
  noTextBody?: boolean;
}): string => placeholderXml(spec);

/** The shape PowerPoint writes for a notes page: image, body, slide number. */
const POWERPOINT_NOTES =
  notesPlaceholder({
    id: 2,
    name: 'Slide Image Placeholder 1',
    type: 'sldImg',
    locks: 'noGrp="1" noRot="1" noChangeAspect="1"',
    // Measured: PowerPoint writes `<p:spPr/>` and no `p:txBody` at all here.
    noTextBody: true,
  }) +
  notesPlaceholder({
    id: 3,
    name: 'Notes Placeholder 2',
    type: 'body',
    // One, against a master that says three. See the header comment.
    idx: 1,
    body: textLine('Speaker notes for slide 1, on the placeholder PowerPoint writes.', {
      lang: 'en-GB',
    }),
  }) +
  notesPlaceholder({
    id: 4,
    name: 'Slide Number Placeholder 3',
    type: 'sldNum',
    size: 'quarter',
    idx: 5,
    body: para({
      content: field({
        id: '{1F60C83A-9B47-42D5-8E01-3A7C9146B5E8}',
        type: 'slidenum',
        text: '1',
        props: { lang: 'en-GB', smtClean: false },
      }),
      endProps: { lang: 'en-GB' },
    }),
  });

/** All six notes-family types on one notes slide, including the refused pair. */
const ALL_SIX_NOTES =
  notesPlaceholder({
    id: 2,
    name: 'Header Placeholder 1',
    type: 'hdr',
    size: 'quarter',
    body: textLine('hdr — refused on a slide layout, legal here', { lang: 'en-GB', sz: 1200 }),
  }) +
  notesPlaceholder({
    id: 3,
    name: 'Date Placeholder 2',
    type: 'dt',
    idx: 1,
    body: para({
      content: field({
        id: '{4C7B0E52-D318-4A96-B72F-05E8A3D14C60}',
        type: 'datetime1',
        text: '27/08/2026',
        props: { lang: 'en-GB', smtClean: false },
      }),
      endProps: { lang: 'en-GB' },
    }),
  }) +
  notesPlaceholder({
    id: 4,
    name: 'Slide Image Placeholder 3',
    type: 'sldImg',
    idx: 2,
    locks: 'noGrp="1" noRot="1" noChangeAspect="1"',
    noTextBody: true,
  }) +
  notesPlaceholder({
    id: 5,
    name: 'Notes Placeholder 4',
    type: 'body',
    size: 'quarter',
    // Three this time, matching the master exactly - so the two notes slides
    // between them exercise both the exact match and the fallback.
    idx: 3,
    body: textLine('sldImg — the other refused type, also legal here.', { lang: 'en-GB' }),
  }) +
  notesPlaceholder({
    id: 6,
    name: 'Footer Placeholder 5',
    type: 'ftr',
    size: 'quarter',
    idx: 4,
    body: textLine('a14 footer', { lang: 'en-GB', sz: 1200 }),
  }) +
  notesPlaceholder({
    id: 7,
    name: 'Slide Number Placeholder 6',
    type: 'sldNum',
    size: 'quarter',
    idx: 5,
    body: para({
      content: field({
        id: '{8E3D5170-2A64-4B08-9F5C-6D01B472E3A9}',
        type: 'slidenum',
        text: '2',
        props: { lang: 'en-GB', smtClean: false },
      }),
      endProps: { lang: 'en-GB' },
    }),
  });

/** A notes page is a sheet like any other: it may hold ordinary shapes too. */
const NOTES_WITH_A_SHAPE =
  notesPlaceholder({
    id: 2,
    name: 'Notes Placeholder 1',
    type: 'body',
    idx: 1,
    body: textLine('A notes page is a sheet, so it can hold shapes as well.', { lang: 'en-GB' }),
  }) +
  shape({
    id: 3,
    name: 'Ordinary shape on a notes page',
    x: 685800,
    y: 1143000,
    cx: NOTES_WIDTH - 2 * 685800,
    cy: 914400,
  });

export const a14Notes: ProbeDeck = {
  id: 'a14-notes',
  title: 'PPTX Studio corpus: a14 notes',
  description:
    'A notes master with all six of its placeholder types, a handout master with four, and three ' +
    'notes slides: the exact shape PowerPoint writes, one carrying every notes-family type, and ' +
    'one holding an ordinary shape beside its body. The geometry and the idx values were measured ' +
    'from a deck PowerPoint 16.0.20326 saved, which is where the finding that a notes slide binds ' +
    'to its master on type rather than on idx comes from. This is also where hdr and sldImg are ' +
    'legal - a02 covers the other fourteen types because those two are refused on a slide.',
  features: {
    shape: 27,
    placeholder: 26,
    presetGeom: 1,
    field: 7,
    gradientFill: 6,
    notesSlide: 3,
  },
  build: () => ({
    title: 'PPTX Studio corpus: a14 notes',
    notesMaster: NOTES_MASTER,
    handoutMaster: HANDOUT_MASTER,
    slides: [
      {
        title: 'a14 — the notes page PowerPoint writes',
        layout: 0,
        body: '',
        notes: POWERPOINT_NOTES,
      },
      {
        title: 'a14 — every notes-family placeholder type',
        layout: 0,
        body: '',
        notes: ALL_SIX_NOTES,
      },
      {
        title: 'a14 — a notes page with a shape on it',
        layout: 0,
        body: '',
        notes: NOTES_WITH_A_SHAPE,
      },
    ],
  }),
};
