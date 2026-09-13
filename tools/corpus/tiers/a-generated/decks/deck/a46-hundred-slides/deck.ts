import {
  placeholderXml,
  TITLE_BOX,
  type ProbeLayout,
  type ProbeSlide,
} from '../../../markup/chassis.ts';
import { scheme, solidFill } from '../../../markup/shapes.ts';
import { buChar, buFont, pPr, rPr } from '../../../markup/text.ts';
import type { ProbeDeck } from '../../../markup/types.ts';

import { hundredSlides } from './kinds.ts';
import { MEDIA_PARTS } from './media.ts';
import { SLIDE_COUNT } from './schedule.ts';

/**
 * A hundred slides of ten kinds from one deterministic schedule: Gate 3's deck.
 *
 * Every other Tier A deck says one thing in three slides. This one is the size a deck has when a
 * person has been writing it for a quarter, and the only slide count that asks the viewer, the
 * harness and the oracle whether they scale. `deflate: true` for scale, not compression: stored,
 * it would sit over the corpus's per-file cap. ADR 0054.
 */

/** The body placeholder every bullets slide inherits, at the content area. */
const BODY_BOX = { x: 457200, y: 1188720, cx: 11277600, cy: 5181600 } as const;

const LAYOUTS: readonly ProbeLayout[] = [
  {
    type: 'titleOnly',
    name: 'Title Only',
    hasTitle: true,
    shapes: placeholderXml({ id: 2, name: 'Title 1', type: 'title', ...TITLE_BOX }),
  },
  { type: 'blank', name: 'Blank' },
  {
    type: 'obj',
    name: 'Title and Content',
    hasTitle: true,
    shapes:
      placeholderXml({ id: 2, name: 'Title 1', type: 'title', ...TITLE_BOX }) +
      placeholderXml({ id: 3, name: 'Content Placeholder 2', idx: 1, ...BODY_BOX }),
  },
  {
    // A section header's own gradient, for the chapters that state no background of their own.
    type: 'secHead',
    name: 'Section Header',
    background:
      '<p:bg><p:bgPr><a:gradFill rotWithShape="1"><a:gsLst>' +
      `<a:gs pos="0">${scheme('lt1')}</a:gs>` +
      `<a:gs pos="100000">${scheme('accent1', '<a:tint val="25000"/>')}</a:gs>` +
      '</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:effectLst/></p:bgPr></p:bg>',
  },
];

/** Bullets on the body levels, the way a PowerPoint master carries them. */
const TEXT_STYLES =
  '<p:txStyles><p:titleStyle>' +
  pPr('lvl1pPr', {
    algn: 'l',
    defRPr: rPr('defRPr', { sz: 3200, b: true, fill: solidFill(scheme('tx1')), latin: '+mj-lt' }),
  }) +
  '</p:titleStyle><p:bodyStyle>' +
  pPr('lvl1pPr', {
    marL: 342900,
    indent: -342900,
    bullet: buChar('•'),
    buFont: buFont('Arial'),
    defRPr: rPr('defRPr', { sz: 2000, fill: solidFill(scheme('tx1')), latin: '+mn-lt' }),
  }) +
  pPr('lvl2pPr', {
    marL: 742950,
    indent: -285750,
    bullet: buChar('–'),
    buFont: buFont('Arial'),
    defRPr: rPr('defRPr', { sz: 1800, fill: solidFill(scheme('tx1')), latin: '+mn-lt' }),
  }) +
  pPr('lvl3pPr', {
    marL: 1143000,
    indent: -228600,
    bullet: buChar('•'),
    buFont: buFont('Arial'),
    defRPr: rPr('defRPr', { sz: 1600, fill: solidFill(scheme('tx1')), latin: '+mn-lt' }),
  }) +
  '</p:bodyStyle><p:otherStyle>' +
  pPr('lvl1pPr', {
    defRPr: rPr('defRPr', { sz: 1800, fill: solidFill(scheme('tx1')), latin: '+mn-lt' }),
  }) +
  '</p:otherStyle></p:txStyles>';

const SLIDES: readonly ProbeSlide[] = hundredSlides();

export const a46HundredSlides: ProbeDeck = {
  id: 'a46-hundred-slides',
  title: 'PPTX Studio corpus: a46 hundred slides',
  description:
    'A hundred slides of ten kinds - cover, agenda, eight section headers, bullets, two-column ' +
    'text, picture and caption, shape grids, quotes, bar diagrams and a closing - from one fixed ' +
    'schedule and a seeded generator, the size a deck has after a quarter of writing. Gate 3 asks ' +
    'whether the viewer, the harness and the oracle scale to it and whether every slide is faithful ' +
    'at every zoom; no three-slide probe can ask either. The first Tier A deck past ten slides and ' +
    'the first deflated for scale rather than to probe compression.',
  addedIn: '3.11',
  // Every count is the chassis plus the kind counts times what each kind draws, and
  // `schedule.test.ts` holds that arithmetic; the census locks the literal to the bytes.
  features: {
    shape: 803,
    placeholder: 108,
    presetGeom: 746,
    gradientFill: 182,
    patternFill: 35,
    picture: 12,
    group: 13,
    connector: 39,
    field: 90,
    shadow: 13,
    glow: 13,
  },
  build: () => ({
    title: 'PPTX Studio corpus: a46 hundred slides',
    deflate: true,
    layouts: LAYOUTS,
    textStyles: TEXT_STYLES,
    parts: MEDIA_PARTS,
    slides: SLIDES,
  }),
};

if (SLIDES.length !== SLIDE_COUNT) {
  throw new Error(`a46 built ${String(SLIDES.length)} slides, not ${String(SLIDE_COUNT)}`);
}
