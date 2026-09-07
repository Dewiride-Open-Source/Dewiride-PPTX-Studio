import {
  placeholderXml,
  TITLE_BOX,
  type ProbeLayout,
  type ProbeSlide,
} from '../../markup/chassis.ts';
import { grid, line, scheme, shape, solidFill, srgb } from '../../markup/shapes.ts';
import type { ProbeDeck } from '../../markup/types.ts';

/**
 * Where a slide's background comes from, and what `p:bgRef/@idx` indexes into.
 *
 * Every other deck in the corpus carries the same one background - the master's
 * `<p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef>` - and no slide or
 * layout overrides it. So the 1000-offset, the `phClr` substitution and the
 * inheritance walk had no corpus slide to be wrong on, and Gate 2 names themed
 * backgrounds. ADR 0038.
 *
 * The offset is what the first two slides separate, and they are chosen so a
 * renderer that ignores it cannot pass by luck: `idx="2"` is
 * `fillStyleLst[1]`, a **gradient**, while `idx="1002"` is
 * `bgFillStyleLst[1]`, a near-white **solid**. Reading either list for the
 * other swaps a gradient for a flat tint across the whole slide.
 */

const STROKE = line({ width: 12700, fill: solidFill(scheme('tx1')) });

/** The colour every `a:schemeClr val="phClr"` in the theme's fill lists becomes. */
const PH = '<a:schemeClr val="accent2"/>';

function bgRef(idx: number): string {
  return `<p:bg><p:bgRef idx="${String(idx)}">${PH}</p:bgRef></p:bg>`;
}

function bgPr(fill: string): string {
  return `<p:bg><p:bgPr>${fill}<a:effectLst/></p:bgPr></p:bg>`;
}

/** Three swatches across the slide: one transparent, so the background shows. */
function swatches(startId: number): string {
  const at = grid(3, 2);
  const fills = ['<a:noFill/>', solidFill(scheme('bg1')), solidFill(srgb('1F1F1F'))];
  const names = ['noFill - the background shows through', 'bg1', 'near black'];
  return fills
    .map((fill, n) =>
      shape({
        id: startId + n,
        name: names[n] ?? '',
        ...at(n),
        fill,
        line: STROKE,
      }),
    )
    .join('');
}

const LAYOUTS: readonly ProbeLayout[] = [
  {
    type: 'titleOnly',
    name: 'Title Only',
    hasTitle: true,
    shapes: placeholderXml({ id: 2, name: 'Title 1', type: 'title', ...TITLE_BOX }),
  },
  {
    // The layout a slide with no `p:bg` of its own inherits from, which is the
    // hop between the slide and the master that nothing else in the corpus has.
    type: 'titleOnly',
    name: 'Title Only, layout background',
    hasTitle: true,
    background: bgPr(solidFill(scheme('accent6', '<a:tint val="40000"/>'))),
    shapes: placeholderXml({ id: 2, name: 'Title 1', type: 'title', ...TITLE_BOX }),
  },
];

const GRADIENT =
  '<a:gradFill rotWithShape="0"><a:gsLst>' +
  `<a:gs pos="0">${srgb('C0392B')}</a:gs>` +
  `<a:gs pos="50000">${srgb('F1C40F')}</a:gs>` +
  `<a:gs pos="100000">${srgb('27AE60')}</a:gs>` +
  '</a:gsLst><a:lin ang="2700000" scaled="0"/></a:gradFill>';

const SLIDES: readonly ProbeSlide[] = [
  {
    title: 'a45 - bgRef idx=2, which is fillStyleLst[1]: a gradient',
    layout: 0,
    background: bgRef(2),
    body: swatches(10),
  },
  {
    title: 'a45 - bgRef idx=1002, which is bgFillStyleLst[1]: a solid tint',
    layout: 0,
    background: bgRef(1002),
    body: swatches(20),
  },
  {
    title: 'a45 - bgPr, a background that owes the theme nothing',
    layout: 0,
    background: bgPr(GRADIENT),
    body: swatches(30),
  },
  {
    title: 'a45 - no p:bg at all: the layout supplies one',
    layout: 1,
    body: swatches(40),
  },
];

export const a45Backgrounds: ProbeDeck = {
  id: 'a45-backgrounds',
  title: 'PPTX Studio corpus: a45 backgrounds',
  description:
    'Slide backgrounds through all three routes: p:bgRef into fillStyleLst and into ' +
    'bgFillStyleLst across the 1000 offset, an explicit p:bgPr gradient, and a slide with no p:bg ' +
    'that inherits its layout instead of the master. First deck in the corpus to write a p:bgPr, ' +
    'a bgRef anywhere but the master, or a background on a layout.',
  addedIn: '2.14',
  features: {
    // Chassis 4 + 4 slide titles + 4 slides of 3 swatches.
    shape: 20,
    placeholder: 8,
    presetGeom: 12,
    gradientFill: 3,
  },
  build: () => ({
    title: 'PPTX Studio corpus: a45 backgrounds',
    layouts: LAYOUTS,
    slides: SLIDES,
  }),
};
