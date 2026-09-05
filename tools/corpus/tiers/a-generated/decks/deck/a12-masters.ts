import { IDENTITY_CLR_MAP, placeholderXml, TITLE_BOX, type ProbeMaster } from '../package.ts';
import { grid, scheme, shape, solidFill } from '../shapes.ts';
import type { ProbeDeck } from '../types.ts';

/**
 * Three masters, three themes, three colour maps.
 *
 * Every deck before this one has exactly one master, and a renderer written
 * against those decks can hold one theme in a variable and be right every time.
 * It stops being right here, and the failure is silent: a shape that says
 * `a:schemeClr val="accent1"` renders in whichever accent the reader happened
 * to load, no markup looks wrong, and nothing reports an error.
 *
 * ## The two dials, which are not the same dial
 *
 * A slide names no theme and no master. It names a **layout**, in its own
 * `.rels`; the layout names a master, in its own `.rels`; the master names a
 * theme, in its own `.rels`. So resolving `accent1` on a slide is a three-hop
 * walk through three relationship parts, and the shortcut - read
 * `ppt/theme/theme1.xml` once - is right only for a single-master deck.
 *
 * On top of that sits `p:clrMap`, which is a different dial entirely:
 *
 * | dial          | lives on   | says                                       |
 * | ------------- | ---------- | ------------------------------------------ |
 * | `a:clrScheme` | the theme  | what `dk1`, `lt1`, `accent1` … *are*       |
 * | `p:clrMap`    | the master | which of those `bg1`, `tx1`, `bg2` … *mean* |
 *
 * Master 2 here holds the **same** theme colours as master 1 and only swaps
 * four attributes of its map - `bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2"` - so
 * the slides bound to it come out inverted. That is how every real dark master
 * is built, and it is why `dk1`/`lt1`/`dk2`/`lt2` **bypass the map** while
 * `bg1`/`tx1`/`bg2`/`tx2` go through it: a swatch asking for `dk1` is asking
 * for the colour, and a swatch asking for `tx1` is asking whatever this master
 * says text is.
 *
 * Master 3 goes the other way: the identity map again, and a genuinely
 * different `a:clrScheme` in a theme part of its own.
 *
 * ## `a:overrideClrMapping`, on slide 3
 *
 * `p:clrMapOvr` is one of two branches. Every deck so far writes
 * `<a:masterClrMapping/>`, which means "whatever the master said". The other is
 * `a:overrideClrMapping`, which carries all twelve attributes again and
 * replaces the master's map **for that slide alone**. Slide 3 re-inverts a
 * master that was not inverted, so a reader that stops at the master paints
 * every one of its eight swatches wrong.
 *
 * ## What the eight swatches are for
 *
 * Six accents, then `bg1` and `tx1` - the two that go through the map. Reading
 * the three slides side by side, the accents move when the *theme* changes and
 * the last two move when the *map* changes, which is the distinction the whole
 * deck exists to make visible.
 */

const SWATCHES = [
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'bg1',
  'tx1',
] as const;

/** One row of eight, each filled with the scheme colour it is labelled with. */
function swatchRow(firstId: number): string {
  const cell = grid(4, 2);
  return SWATCHES.map((name, index) =>
    shape({
      id: firstId + index,
      name: name,
      ...cell(index),
      fill: solidFill(scheme(name)),
    }),
  ).join('');
}

/** A twelve-attribute map with the four mapped slots swapped light for dark. */
const INVERTED_CLR_MAP =
  'bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent1" accent2="accent2"' +
  ' accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6"' +
  ' hlink="hlink" folHlink="folHlink"';

/** A second colour scheme, far enough from the first that a swap is obvious. */
const WARM_SCHEME =
  '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>' +
  '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>' +
  '<a:dk2><a:srgbClr val="5B3A29"/></a:dk2>' +
  '<a:lt2><a:srgbClr val="FBF3E4"/></a:lt2>' +
  '<a:accent1><a:srgbClr val="C0392B"/></a:accent1>' +
  '<a:accent2><a:srgbClr val="E67E22"/></a:accent2>' +
  '<a:accent3><a:srgbClr val="F1C40F"/></a:accent3>' +
  '<a:accent4><a:srgbClr val="8E6E53"/></a:accent4>' +
  '<a:accent5><a:srgbClr val="A93226"/></a:accent5>' +
  '<a:accent6><a:srgbClr val="D68910"/></a:accent6>' +
  '<a:hlink><a:srgbClr val="B9770E"/></a:hlink>' +
  '<a:folHlink><a:srgbClr val="7E5109"/></a:folHlink>';

const titlePlaceholder = (id: number, name: string): string =>
  placeholderXml({ id, name, type: 'title', ...TITLE_BOX });

const MASTERS: readonly ProbeMaster[] = [
  {
    // Master 1: the chassis default, spelled out so the three read alike.
    clrMap: IDENTITY_CLR_MAP,
    themeName: 'PPTX Studio Corpus',
    layouts: [
      {
        type: 'titleOnly',
        name: 'Cool Title Only',
        hasTitle: true,
        shapes: titlePlaceholder(2, 'Title 1'),
      },
      { type: 'blank', name: 'Cool Blank' },
    ],
  },
  {
    // Master 2: the same theme, four attributes of the map swapped.
    clrMap: INVERTED_CLR_MAP,
    themeName: 'PPTX Studio Corpus',
    layouts: [
      {
        type: 'titleOnly',
        name: 'Inverted Title Only',
        hasTitle: true,
        shapes: titlePlaceholder(2, 'Title 1'),
      },
    ],
  },
  {
    // Master 3: the identity map, and a theme part of its own.
    clrMap: IDENTITY_CLR_MAP,
    themeName: 'PPTX Studio Corpus Warm',
    colours: WARM_SCHEME,
    layouts: [
      {
        type: 'titleOnly',
        name: 'Warm Title Only',
        hasTitle: true,
        shapes: titlePlaceholder(2, 'Title 1'),
      },
    ],
  },
];

export const a12Masters: ProbeDeck = {
  id: 'a12-masters',
  title: 'PPTX Studio corpus: a12 masters',
  description:
    'Three slide masters, each with its own theme part and its own p:clrMap, and one slide per ' +
    'master painting the same eight scheme colours. The second master differs from the first only ' +
    'in four p:clrMap attributes, so bg1 and tx1 invert while the accents do not; the third has a ' +
    'different a:clrScheme entirely. Slide 3 carries a:overrideClrMapping, which replaces its ' +
    "master's map for that slide alone.",
  features: {
    shape: 36,
    placeholder: 12,
    presetGeom: 24,
    gradientFill: 6,
  },
  build: () => ({
    title: 'PPTX Studio corpus: a12 masters',
    masters: MASTERS,
    slides: [
      {
        title: 'a12 — master 1, identity map',
        layout: 0,
        body: swatchRow(10),
      },
      {
        title: 'a12 — master 2, inverted map, same theme',
        layout: 2,
        body: swatchRow(10),
      },
      {
        title: 'a12 — master 3, warm theme, and an override on the slide',
        layout: 3,
        // The master this slide is bound to maps identity. This puts the
        // inverted map back on top of it, for this slide only.
        clrMapOvr: `<a:overrideClrMapping ${INVERTED_CLR_MAP}/>`,
        body: swatchRow(10),
      },
    ],
  }),
};
