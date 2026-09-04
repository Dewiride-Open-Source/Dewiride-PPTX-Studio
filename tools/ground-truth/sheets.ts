/**
 * Experiment C5 - the probes. What does a shape inherit, and from where?
 *
 * ## The measurement is a position, and PowerPoint reports it directly
 *
 * C3 and C4 sampled bitmaps because a fill and a stroke are pictures. An
 * inheritance is not: a placeholder with no `a:xfrm` of its own lands *exactly*
 * where the sheet it inherited from put it, and `Shape.Left/Top/Width/Height`
 * over COM reports that in points with no antialiasing to interpolate through.
 * So every candidate parent gets a rectangle no other candidate shares, and the
 * rectangle a probe reports back names the parent it matched. One number, no
 * fitting, no error bars.
 *
 * That is the cheapest measurement in the project so far and it is also the
 * most direct: the object model's answer is the resolver's answer, read out of
 * the resolver rather than reconstructed from what it painted.
 *
 * ## The grid
 *
 * Five bands of seven boxes. `M0..M6` on the master, `A0..A6` on the first
 * layout, `B`, `C`, `D` on the others. Any two are at least 80 points apart in
 * y and 130 in x, which is far more than any plausible rounding.
 */

import {
  SCHEME_ONE,
  SCHEME_TWO,
  shape,
  type LayoutSpec,
  type MasterSpec,
  type Rect,
  type SheetPackage,
  type SlideSpec,
} from './sheet-pptx.ts';

/* -------------------------------------------------------------------------- */
/* the grid                                                                   */
/* -------------------------------------------------------------------------- */

const BANDS = { A: 20, B: 100, C: 180, D: 260, M: 340 } as const;
type Band = keyof typeof BANDS;

/** `A0` .. `M6`. Every box is unique in both axes. */
export function box(name: string): Rect {
  const band = name[0] as Band;
  const col = Number(name.slice(1));
  const y = BANDS[band];
  if (y === undefined || !Number.isInteger(col) || col < 0 || col > 6) {
    throw new Error(`no such box: ${name}`);
  }
  return { x: 20 + col * 130, y, w: 110, h: 60 };
}

export const ALL_BOXES: Readonly<Record<string, Rect>> = Object.fromEntries(
  (['A', 'B', 'C', 'D', 'M'] as const).flatMap((band) =>
    [0, 1, 2, 3, 4, 5, 6].map((col) => [`${band}${String(col)}`, box(`${band}${String(col)}`)]),
  ),
);

/** Where a probe puts a shape that states its own geometry: clear of every band. */
const OWN_RECT: Rect = { x: 400, y: 460, w: 120, h: 50 };

const MAGENTA = '<a:solidFill><a:srgbClr val="FF00FF"/></a:solidFill>';
const NO_LINE = '<a:ln><a:noFill/></a:ln>';

/* -------------------------------------------------------------------------- */
/* probes and decks                                                           */
/* -------------------------------------------------------------------------- */

export type ProbeKind = 'match' | 'style' | 'background' | 'clrmap' | 'hostile';

export interface Probe {
  /** Unique across the experiment, and the shape's `@name` when it has one. */
  readonly id: string;
  readonly deck: string;
  readonly kind: ProbeKind;
  /** 1-based index of the slide the probe lives on. */
  readonly slide: number;
  /** One line, in the present tense, saying what is being asked. */
  readonly question: string;
  /**
   * The box this probe is expected to land in, when the expectation is
   * falsifiable in advance. `null` means "nobody knows, that is the point".
   */
  readonly expectBox?: string | null | undefined;
  /** Expected resolved fill, as six hex digits, when that is the measurement. */
  readonly expectFill?: string | null | undefined;
}

export interface SheetDeck {
  readonly deck: string;
  readonly pkg: SheetPackage;
  readonly probes: readonly Probe[];
  /** Alone in its own package, so a refusal names its own cause. */
  readonly hostile: boolean;
}

let nextId = 100;
function shapeId(): number {
  return nextId++;
}

/** A placeholder that states no geometry: the subject of every match probe. */
function phProbe(name: string, ph: string): string {
  return shape({ id: shapeId(), name, ph, fill: MAGENTA, line: NO_LINE });
}

/** A placeholder that does state geometry: a candidate parent. */
function phBox(name: string, ph: string, boxName: string): string {
  return shape({ id: shapeId(), name, ph, rect: box(boxName), fill: MAGENTA, line: NO_LINE });
}

/** A placeholder with neither geometry nor a parent that has any. */
function phBare(name: string, ph: string): string {
  return shape({ id: shapeId(), name, ph, fill: MAGENTA, line: NO_LINE });
}

/**
 * A placeholder that states nothing whatsoever - no geometry, no fill, no line.
 *
 * `phBare` carries a magenta fill so a position probe is visible in the
 * exported bitmap, and that fill is inherited like everything else: the first
 * run of the two-hop *style* probes measured the layout's magenta rather than
 * the master's style, and reported a perfectly plausible colour for it. A
 * probe that asks what descends must not itself declare the thing it is asking
 * about.
 */
function phNaked(name: string, ph: string): string {
  return shape({ id: shapeId(), name, ph });
}

/* -------------------------------------------------------------------------- */
/* the master every match deck shares                                         */
/* -------------------------------------------------------------------------- */

/**
 * Five placeholders on the master, at the `M` band, and five is the maximum.
 *
 * A slide master's placeholder vocabulary is **not** `ST_PlaceholderType`. It is
 * `title`, `body`, `dt`, `ftr`, `sldNum` and `hdr`, and nothing else: a master
 * carrying `ctrTitle`, `subTitle`, `obj` or `pic` is repaired on open. That was
 * found the expensive way - four otherwise clean probe decks came back REPAIRED
 * and had to be bisected down to one attribute - and it is recorded as a hostile
 * probe below so nobody has to find it twice.
 *
 * The `@idx` values are PowerPoint's own: a stock master numbers its date,
 * footer and slide-number placeholders 2, 3 and 4 while every stock *layout*
 * numbers the same three 10, 11 and 12. That mismatch is not a quirk of one
 * template - it is in every deck Office has ever written - and it is why
 * layout-to-master matching cannot be the same function as slide-to-layout
 * matching.
 */
function matchMaster(): MasterSpec {
  return {
    theme: 0,
    shapes: [
      phBox('m-title', 'type="title"', 'M0'),
      phBox('m-body', 'type="body" idx="1"', 'M1'),
      phBox('m-dt', 'type="dt" sz="half" idx="2"', 'M2'),
      phBox('m-ftr', 'type="ftr" sz="quarter" idx="3"', 'M3'),
      phBox('m-sldNum', 'type="sldNum" sz="quarter" idx="4"', 'M4'),
    ],
  };
}

/** One slide carrying one probe, so no probe can be confused with another. */
function probeSlide(layout: number, name: string, ph: string): SlideSpec {
  return { layout, name, shapes: [phProbe(name, ph)] };
}

/* -------------------------------------------------------------------------- */
/* deck: match - the five tiers against a layout that has everything          */
/* -------------------------------------------------------------------------- */

function matchDeck(): SheetDeck {
  const layout: LayoutSpec = {
    master: 0,
    type: 'obj',
    name: 'Full',
    shapes: [
      phBox('l-title', 'type="title"', 'A0'),
      phBox('l-body1', 'type="body" idx="1"', 'A1'),
      phBox('l-body2', 'type="body" idx="2"', 'A2'),
      phBox('l-body3', 'type="body" idx="3"', 'A3'),
      phBox('l-dt', 'type="dt" sz="half" idx="10"', 'A4'),
      phBox('l-ftr', 'type="ftr" sz="quarter" idx="11"', 'A5'),
      phBox('l-sldNum', 'type="sldNum" sz="quarter" idx="12"', 'A6'),
    ],
  };

  const cases: readonly {
    id: string;
    ph: string;
    q: string;
    expect: string | null;
  }[] = [
    {
      id: 'exact-title',
      ph: 'type="title"',
      q: 'a title with no idx finds the layout title',
      expect: 'A0',
    },
    {
      id: 'exact-body2',
      ph: 'type="body" idx="2"',
      q: 'an exact (type, idx) pair wins',
      expect: 'A2',
    },
    {
      id: 'ctrTitle',
      ph: 'type="ctrTitle"',
      q: 'ctrTitle finds a layout that only has title',
      expect: 'A0',
    },
    {
      id: 'title-idx9',
      ph: 'type="title" idx="9"',
      q: 'a title matches regardless of idx',
      expect: 'A0',
    },
    {
      id: 'sldNum-idx9',
      ph: 'type="sldNum" idx="9"',
      q: 'sldNum matches on raw type, ignoring idx',
      expect: 'A6',
    },
    { id: 'dt-idx9', ph: 'type="dt" idx="9"', q: 'dt matches on raw type', expect: 'A4' },
    { id: 'ftr-idx9', ph: 'type="ftr" idx="9"', q: 'ftr matches on raw type', expect: 'A5' },
    {
      id: 'bare',
      ph: '',
      q: 'a p:ph with no attributes at all: which default does it take',
      expect: null,
    },
    {
      id: 'idx-only-2',
      ph: 'idx="2"',
      q: 'an idx with no type against a typed layout placeholder',
      expect: null,
    },
    {
      id: 'obj-idx2',
      ph: 'type="obj" idx="2"',
      q: 'does obj match a body at the same idx',
      expect: null,
    },
    {
      id: 'subTitle-idx1',
      ph: 'type="subTitle" idx="1"',
      q: 'does subTitle match a body at the same idx',
      expect: null,
    },
    {
      id: 'pic-idx1',
      ph: 'type="pic" idx="1"',
      q: 'does pic match a body at the same idx',
      expect: null,
    },
  ];

  const slides: SlideSpec[] = cases.map((c) => probeSlide(0, `p-${c.id}`, c.ph));
  // The control: a placeholder that states its own geometry must keep it, or
  // every measurement above is measuring the wrong thing.
  slides.push({
    layout: 0,
    name: 'p-explicit',
    shapes: [
      shape({
        id: shapeId(),
        name: 'p-explicit',
        ph: 'type="body" idx="1"',
        rect: OWN_RECT,
        fill: MAGENTA,
        line: NO_LINE,
      }),
    ],
  });

  const probes: Probe[] = cases.map((c, i) => ({
    id: `p-${c.id}`,
    deck: 'match',
    kind: 'match' as const,
    slide: i + 1,
    question: c.q,
    expectBox: c.expect,
  }));
  probes.push({
    id: 'p-explicit',
    deck: 'match',
    kind: 'match',
    slide: cases.length + 1,
    question: 'an explicit a:xfrm on the slide overrides the layout',
    expectBox: 'own',
  });

  return {
    deck: 'match',
    hostile: false,
    pkg: { themes: [SCHEME_ONE], masters: [matchMaster()], layouts: [layout], slides },
    probes,
  };
}

/* -------------------------------------------------------------------------- */
/* deck: tier4 - a layout with exactly one body                               */
/* -------------------------------------------------------------------------- */

function tier4Deck(): SheetDeck {
  const layout: LayoutSpec = {
    master: 0,
    type: 'obj',
    name: 'One body',
    shapes: [phBox('l-title', 'type="title"', 'B0'), phBox('l-body7', 'type="body" idx="7"', 'B1')],
  };

  const cases: readonly { id: string; ph: string; q: string }[] = [
    {
      id: 't4-body0',
      ph: 'type="body" idx="0"',
      q: 'one body in the layout, wrong idx on the slide',
    },
    { id: 't4-body-noidx', ph: 'type="body"', q: 'one body in the layout, no idx on the slide' },
    { id: 't4-obj', ph: 'type="obj"', q: 'obj against a layout whose only content is a body' },
    { id: 't4-bare', ph: '', q: 'a bare p:ph against a layout with one body at idx 7' },
    { id: 't4-obj7', ph: 'type="obj" idx="7"', q: 'obj at the body idx' },
    {
      id: 't4-title9',
      ph: 'type="title" idx="9"',
      q: 'title with a stray idx, layout title has none',
    },
  ];

  return {
    deck: 'tier4',
    hostile: false,
    pkg: {
      themes: [SCHEME_ONE],
      masters: [matchMaster()],
      layouts: [layout],
      slides: cases.map((c) => probeSlide(0, c.id, c.ph)),
    },
    probes: cases.map((c, i) => ({
      id: c.id,
      deck: 'tier4',
      kind: 'match' as const,
      slide: i + 1,
      question: c.q,
      expectBox: null,
    })),
  };
}

/* -------------------------------------------------------------------------- */
/* deck: family - the title family and the body family                        */
/* -------------------------------------------------------------------------- */

function familyDeck(): SheetDeck {
  const layout: LayoutSpec = {
    master: 0,
    type: 'title',
    name: 'Centre title',
    shapes: [
      phBox('l-ctrTitle', 'type="ctrTitle"', 'C0'),
      phBox('l-subTitle', 'type="subTitle" idx="1"', 'C1'),
      phBox('l-obj2', 'type="obj" idx="2"', 'C2'),
      phBox('l-clipArt', 'type="clipArt" idx="3"', 'C3'),
      phBox('l-chart', 'type="chart" idx="4"', 'C4'),
      phBox('l-tbl', 'type="tbl" idx="5"', 'C5'),
      phBox('l-media', 'type="media" idx="6"', 'C6'),
    ],
  };

  const cases: readonly { id: string; ph: string; q: string }[] = [
    { id: 'f-title', ph: 'type="title"', q: 'title against a layout that only has ctrTitle' },
    { id: 'f-body1', ph: 'type="body" idx="1"', q: 'body against a subTitle at the same idx' },
    { id: 'f-body2', ph: 'type="body" idx="2"', q: 'body against an obj at the same idx' },
    { id: 'f-obj3', ph: 'type="obj" idx="3"', q: 'obj against a clipArt at the same idx' },
    { id: 'f-body4', ph: 'type="body" idx="4"', q: 'body against a chart at the same idx' },
    { id: 'f-tbl5', ph: 'type="tbl" idx="5"', q: 'an exact match on a content type' },
    {
      id: 'f-media9',
      ph: 'type="media" idx="9"',
      q: 'media with the wrong idx, one media present',
    },
  ];

  return {
    deck: 'family',
    hostile: false,
    pkg: {
      themes: [SCHEME_ONE],
      masters: [matchMaster()],
      layouts: [layout],
      slides: cases.map((c) => probeSlide(0, c.id, c.ph)),
    },
    probes: cases.map((c, i) => ({
      id: c.id,
      deck: 'family',
      kind: 'match' as const,
      slide: i + 1,
      question: c.q,
      expectBox: null,
    })),
  };
}

/* -------------------------------------------------------------------------- */
/* deck: twohop - slide to layout to master                                   */
/* -------------------------------------------------------------------------- */

/**
 * The layout's placeholders carry no geometry, so a slide probe can only land
 * anywhere at all by going through the layout to the master. And the layout's
 * `@idx` values deliberately disagree with the master's on all six, which is
 * the stock template's own configuration.
 */
function twoHopDeck(): SheetDeck {
  const layout: LayoutSpec = {
    master: 0,
    type: 'obj',
    name: 'Inherits everything',
    shapes: [
      phBare('l-title', 'type="title"'),
      phBare('l-body7', 'type="body" idx="7"'),
      phBare('l-dt', 'type="dt" sz="half" idx="10"'),
      phBare('l-ftr', 'type="ftr" sz="quarter" idx="11"'),
      phBare('l-sldNum', 'type="sldNum" sz="quarter" idx="12"'),
      phBare('l-pic', 'type="pic" idx="9"'),
      phBare('l-obj', 'type="obj" idx="8"'),
    ],
  };

  const cases: readonly { id: string; ph: string; q: string; expect: string | null }[] = [
    {
      id: 'h2-title',
      ph: 'type="title"',
      q: 'title through an empty layout to the master',
      expect: 'M0',
    },
    {
      id: 'h2-body',
      ph: 'type="body" idx="7"',
      q: 'body idx 7 on the layout reaches the master body at idx 1',
      expect: 'M1',
    },
    {
      id: 'h2-dt',
      ph: 'type="dt" idx="10"',
      q: 'the layout dt at idx 10 reaches the master dt at idx 2',
      expect: 'M2',
    },
    {
      id: 'h2-ftr',
      ph: 'type="ftr" idx="11"',
      q: 'the layout ftr at idx 11 reaches the master ftr at idx 3',
      expect: 'M3',
    },
    {
      id: 'h2-sldNum',
      ph: 'type="sldNum" idx="12"',
      q: 'the layout sldNum at idx 12 reaches the master sldNum at idx 4',
      expect: 'M4',
    },
    // A master may not carry a `pic` or an `obj` at all, so if either of these
    // lands on `M1` then a content placeholder normalises to `body` on the way
    // up - which is the only reading under which the stock "Title and Content"
    // layout can inherit anything from the stock master.
    {
      id: 'h2-pic',
      ph: 'type="pic" idx="9"',
      q: 'a layout pic, which no master may hold, reaches the master body',
      expect: null,
    },
    {
      id: 'h2-obj',
      ph: 'type="obj" idx="8"',
      q: 'a layout obj reaches the master body',
      expect: null,
    },
  ];

  return {
    deck: 'twohop',
    hostile: false,
    pkg: {
      themes: [SCHEME_ONE],
      masters: [matchMaster()],
      layouts: [layout],
      slides: cases.map((c) => probeSlide(0, c.id, c.ph)),
    },
    probes: cases.map((c, i) => ({
      id: c.id,
      deck: 'twohop',
      kind: 'match' as const,
      slide: i + 1,
      question: c.q,
      expectBox: c.expect,
    })),
  };
}

/* -------------------------------------------------------------------------- */
/* deck: orphan - placeholders with nowhere to inherit from                   */
/* -------------------------------------------------------------------------- */

function orphanDeck(): SheetDeck {
  const layouts: LayoutSpec[] = [
    {
      master: 0,
      name: 'No title',
      shapes: [
        phBox('l-body1', 'type="body" idx="1"', 'A1'),
        phBox('l-body2', 'type="body" idx="2"', 'A2'),
      ],
    },
    { master: 0, name: 'Nothing at all', type: 'blank', shapes: [] },
  ];

  const cases: readonly { id: string; layout: number; ph: string; q: string }[] = [
    {
      id: 'o-title-notitle',
      layout: 0,
      ph: 'type="title"',
      q: 'a title where the layout has none',
    },
    {
      id: 'o-body9',
      layout: 0,
      ph: 'type="body" idx="9"',
      q: 'a body idx nobody has, with two bodies present',
    },
    {
      id: 'o-title-empty',
      layout: 1,
      ph: 'type="title"',
      q: 'a title where the layout has no placeholders at all',
    },
    {
      id: 'o-body-empty',
      layout: 1,
      ph: 'type="body" idx="1"',
      q: 'a body where the layout has nothing',
    },
  ];

  return {
    deck: 'orphan',
    hostile: false,
    pkg: {
      themes: [SCHEME_ONE],
      masters: [{ theme: 0, shapes: [] }],
      layouts,
      slides: cases.map((c) => probeSlide(c.layout, c.id, c.ph)),
    },
    probes: cases.map((c, i) => ({
      id: c.id,
      deck: 'orphan',
      kind: 'match' as const,
      slide: i + 1,
      question: c.q,
      expectBox: null,
    })),
  };
}

/* -------------------------------------------------------------------------- */
/* deck: match2 - is any placeholder type privileged at the first hop         */
/* -------------------------------------------------------------------------- */

/**
 * The first round said the slide-to-layout hop matches on `@idx` and ignores
 * `@type` entirely, on all thirty-seven probes. That is a strong claim, and the
 * way to test a strong claim is to look for the case that would break it: a
 * type that is *privileged*. A title is the obvious candidate, and the
 * header-and-footer trio is the other, because both get special handling in
 * every implementation that has ever written a matcher.
 *
 * So this layout deliberately puts the wrong type at every index.
 */
function match2Deck(): SheetDeck {
  const layout: LayoutSpec = {
    master: 0,
    name: 'Types shuffled',
    shapes: [
      phBox('l-body0', 'type="body" idx="0"', 'A0'),
      phBox('l-title1', 'type="title" idx="1"', 'A1'),
      phBox('l-sldNum2', 'type="sldNum" idx="2"', 'A2'),
      phBox('l-dt3', 'type="dt" idx="3"', 'A3'),
      phBox('l-ftr4', 'type="ftr" idx="4"', 'A4'),
    ],
  };

  const cases: readonly { id: string; ph: string; q: string }[] = [
    { id: 'm2-title0', ph: 'type="title"', q: 'a title at an index the layout gives to a body' },
    {
      id: 'm2-body1',
      ph: 'type="body" idx="1"',
      q: 'a body at an index the layout gives to a title',
    },
    {
      id: 'm2-body2',
      ph: 'type="body" idx="2"',
      q: 'a body at an index the layout gives to sldNum',
    },
    {
      id: 'm2-sldNum3',
      ph: 'type="sldNum" idx="3"',
      q: 'sldNum at an index the layout gives to dt',
    },
    { id: 'm2-ftr0', ph: 'type="ftr" idx="0"', q: 'ftr at an index the layout gives to a body' },
  ];

  return {
    deck: 'match2',
    hostile: false,
    pkg: {
      themes: [SCHEME_ONE],
      masters: [matchMaster()],
      layouts: [layout],
      slides: cases.map((c) => probeSlide(0, c.id, c.ph)),
    },
    probes: cases.map((c, i) => ({
      id: c.id,
      deck: 'match2',
      kind: 'match' as const,
      slide: i + 1,
      question: c.q,
      expectBox: null,
    })),
  };
}

/* -------------------------------------------------------------------------- */
/* deck: twohop2 - type-only, but which one when there are two               */
/* -------------------------------------------------------------------------- */

/**
 * "Type only" leaves one thing unsaid: what happens when the master has *two*
 * placeholders of the type asked for. If the answer is the one whose `@idx`
 * agrees, then the rule is really "type first, idx as a tiebreak" and a matcher
 * that returns the first is wrong on any master with two body placeholders -
 * which is every master with a two-column layout family behind it.
 */
function twoHop2Deck(): SheetDeck {
  const master: MasterSpec = {
    theme: 0,
    shapes: [
      phBox('m-title', 'type="title"', 'M0'),
      phBox('m-body1', 'type="body" idx="1"', 'M1'),
      phBox('m-body5', 'type="body" idx="5"', 'M2'),
      phBox('m-dt', 'type="dt" sz="half" idx="2"', 'M3'),
    ],
  };
  const layout: LayoutSpec = {
    master: 0,
    name: 'Inherits by type',
    shapes: [
      phBare('l-ctrTitle', 'type="ctrTitle"'),
      phBare('l-body5', 'type="body" idx="5"'),
      phBare('l-sub6', 'type="subTitle" idx="6"'),
      phBare('l-dt10', 'type="dt" sz="half" idx="10"'),
    ],
  };

  const cases: readonly { id: string; ph: string; q: string }[] = [
    { id: 'h3-ctrTitle', ph: 'idx="0"', q: 'a layout ctrTitle reaches the master title' },
    {
      id: 'h3-body5',
      ph: 'idx="5"',
      q: 'a layout body at idx 5, with master bodies at idx 1 and idx 5: which one',
    },
    { id: 'h3-sub6', ph: 'idx="6"', q: 'a layout subTitle reaches a master body' },
    { id: 'h3-dt10', ph: 'idx="10"', q: 'the control: a layout dt reaches the master dt' },
  ];

  return {
    deck: 'twohop2',
    hostile: false,
    pkg: {
      themes: [SCHEME_ONE],
      masters: [master],
      layouts: [layout],
      slides: cases.map((c) => probeSlide(0, c.id, c.ph)),
    },
    probes: cases.map((c, i) => ({
      id: c.id,
      deck: 'twohop2',
      kind: 'match' as const,
      slide: i + 1,
      question: c.q,
      expectBox: null,
    })),
  };
}

/* -------------------------------------------------------------------------- */
/* deck: stylehop - do properties travel the same chain as geometry           */
/* -------------------------------------------------------------------------- */

function styleHopDeck(): SheetDeck {
  const master: MasterSpec = {
    theme: 0,
    shapes: [
      shape({
        id: shapeId(),
        name: 'm-styled',
        ph: 'type="body" idx="1"',
        rect: box('M1'),
        style: STYLE(2, 2, 'accent2'),
      }),
      shape({
        id: shapeId(),
        name: 'm-filled',
        ph: 'type="title"',
        rect: box('M0'),
        fill: '<a:solidFill><a:srgbClr val="00A000"/></a:solidFill>',
      }),
    ],
  };
  const layout: LayoutSpec = {
    master: 0,
    name: 'Passes through',
    shapes: [phNaked('l-body1', 'type="body" idx="1"'), phNaked('l-title', 'type="title"')],
  };

  return {
    deck: 'stylehop',
    hostile: false,
    pkg: {
      themes: [SCHEME_ONE],
      masters: [master],
      layouts: [layout],
      slides: [
        {
          layout: 0,
          name: 'sh-style',
          shapes: [shape({ id: shapeId(), name: 'sh-style', ph: 'idx="1"' })],
        },
        {
          layout: 0,
          name: 'sh-fill',
          shapes: [shape({ id: shapeId(), name: 'sh-fill', ph: 'idx="0"' })],
        },
      ],
    },
    probes: [
      {
        id: 'sh-style',
        deck: 'stylehop',
        kind: 'style',
        slide: 1,
        question: 'does a p:style reach a slide through a layout placeholder that has none',
        expectFill: null,
      },
      {
        id: 'sh-fill',
        deck: 'stylehop',
        kind: 'style',
        slide: 2,
        question: 'does an spPr fill reach a slide two hops up',
        expectFill: null,
      },
    ],
  };
}

/* -------------------------------------------------------------------------- */
/* deck: inherit - what else comes down the chain besides geometry            */
/* -------------------------------------------------------------------------- */

const STYLE = (fill: number, ln: number, clr: string): string =>
  '<p:style>' +
  `<a:lnRef idx="${String(ln)}"><a:schemeClr val="${clr}"/></a:lnRef>` +
  `<a:fillRef idx="${String(fill)}"><a:schemeClr val="${clr}"/></a:fillRef>` +
  `<a:effectRef idx="0"><a:schemeClr val="${clr}"/></a:effectRef>` +
  '<a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef>' +
  '</p:style>';

/**
 * Position is not the only thing that descends. Whether a *fill* does is a
 * separate question with a separate answer, and a renderer that assumes one
 * from the other paints half a deck wrong.
 */
function inheritDeck(): SheetDeck {
  const layout: LayoutSpec = {
    master: 0,
    name: 'Fills and styles',
    shapes: [
      shape({
        id: shapeId(),
        name: 'l-fill',
        ph: 'type="body" idx="1"',
        rect: box('A1'),
        fill: '<a:solidFill><a:srgbClr val="00A000"/></a:solidFill>',
        line: NO_LINE,
      }),
      shape({
        id: shapeId(),
        name: 'l-style',
        ph: 'type="body" idx="2"',
        rect: box('A2'),
        style: STYLE(1, 2, 'accent2'),
      }),
      shape({
        id: shapeId(),
        name: 'l-line',
        ph: 'type="body" idx="3"',
        rect: box('A3'),
        fill: '<a:noFill/>',
        line: '<a:ln w="57150"><a:solidFill><a:srgbClr val="0000A0"/></a:solidFill></a:ln>',
      }),
    ],
  };

  const slides: SlideSpec[] = [
    {
      layout: 0,
      name: 'i-fill',
      shapes: [shape({ id: shapeId(), name: 'i-fill', ph: 'type="body" idx="1"' })],
    },
    {
      layout: 0,
      name: 'i-fill-override',
      shapes: [
        shape({
          id: shapeId(),
          name: 'i-fill-override',
          ph: 'type="body" idx="1"',
          fill: '<a:solidFill><a:srgbClr val="A00000"/></a:solidFill>',
        }),
      ],
    },
    {
      layout: 0,
      name: 'i-style',
      shapes: [shape({ id: shapeId(), name: 'i-style', ph: 'type="body" idx="2"' })],
    },
    {
      layout: 0,
      name: 'i-line',
      shapes: [shape({ id: shapeId(), name: 'i-line', ph: 'type="body" idx="3"' })],
    },
    {
      layout: 0,
      name: 'i-nonph',
      shapes: [
        shape({ id: shapeId(), name: 'i-nonph', rect: box('D0') }),
        shape({
          id: shapeId(),
          name: 'i-nonph-style',
          rect: box('D1'),
          style: STYLE(1, 2, 'accent3'),
        }),
      ],
    },
  ];

  return {
    deck: 'inherit',
    hostile: false,
    pkg: { themes: [SCHEME_ONE], masters: [{ theme: 0, shapes: [] }], layouts: [layout], slides },
    probes: [
      {
        id: 'i-fill',
        deck: 'inherit',
        kind: 'style',
        slide: 1,
        question: 'does a slide placeholder inherit the layout placeholder solid fill',
        expectFill: null,
      },
      {
        id: 'i-fill-override',
        deck: 'inherit',
        kind: 'style',
        slide: 2,
        question: 'an explicit fill on the slide beats the layout',
        expectFill: 'A00000',
      },
      {
        id: 'i-style',
        deck: 'inherit',
        kind: 'style',
        slide: 3,
        question: 'does a slide placeholder inherit the layout p:style',
        expectFill: null,
      },
      {
        id: 'i-line',
        deck: 'inherit',
        kind: 'style',
        slide: 4,
        question: 'does a slide placeholder inherit the layout a:ln',
        expectFill: null,
      },
      {
        id: 'i-nonph',
        deck: 'inherit',
        kind: 'style',
        slide: 5,
        question: 'what a shape with no fill, no style and no placeholder paints',
        expectFill: null,
      },
      {
        id: 'i-nonph-style',
        deck: 'inherit',
        kind: 'style',
        slide: 5,
        question: 'a non-placeholder with only a p:style',
        expectFill: null,
      },
    ],
  };
}

/* -------------------------------------------------------------------------- */
/* deck: style - the style matrix and phClr                                   */
/* -------------------------------------------------------------------------- */

const STYLE_CASES: readonly {
  id: string;
  fill: number;
  ln: number;
  colour: string;
  q: string;
  fillXml?: string;
}[] = [
  { id: 's-fill1-a1', fill: 1, ln: 1, colour: 'accent1', q: 'fillRef 1 is fillStyleLst entry one' },
  { id: 's-fill2-a1', fill: 2, ln: 2, colour: 'accent1', q: 'fillRef 2 is fillStyleLst entry two' },
  {
    id: 's-fill3-a1',
    fill: 3,
    ln: 3,
    colour: 'accent1',
    q: 'fillRef 3 is fillStyleLst entry three',
  },
  {
    id: 's-fill1001-a1',
    fill: 1001,
    ln: 1,
    colour: 'accent1',
    q: 'fillRef 1001 is bgFillStyleLst entry one',
  },
  {
    id: 's-fill1002-a1',
    fill: 1002,
    ln: 1,
    colour: 'accent1',
    q: 'fillRef 1002 is bgFillStyleLst two',
  },
  {
    id: 's-fill1003-a1',
    fill: 1003,
    ln: 1,
    colour: 'accent1',
    q: 'fillRef 1003 is bgFillStyleLst three',
  },
  { id: 's-fill1-a2', fill: 1, ln: 1, colour: 'accent2', q: 'phClr is the colour the ref carries' },
  {
    id: 's-fill2-a2',
    fill: 2,
    ln: 1,
    colour: 'accent2',
    q: 'the entry transform applies to that colour',
  },
  { id: 's-fill2-a6', fill: 2, ln: 1, colour: 'accent6', q: 'and again on a third colour' },
  { id: 's-fill0', fill: 0, ln: 1, colour: 'accent1', q: 'fillRef 0 is no fill at all' },
  { id: 's-ln0', fill: 1, ln: 0, colour: 'accent1', q: 'lnRef 0 is no line at all' },
];

function styleDeck(): SheetDeck {
  const shapes: string[] = [];
  STYLE_CASES.forEach((c, i) => {
    const col = i % 5;
    const row = Math.floor(i / 5);
    shapes.push(
      shape({
        id: shapeId(),
        name: c.id,
        rect: { x: 20 + col * 180, y: 20 + row * 120, w: 150, h: 100 },
        style: STYLE(c.fill, c.ln, c.colour),
      }),
    );
  });
  // The two shapes that ask what a ref's *child* may be, rather than which
  // entry it names: a literal colour, and a transformed scheme colour of the
  // kind PowerPoint's own gallery writes.
  shapes.push(
    shape({
      id: shapeId(),
      name: 's-fill1-srgb',
      rect: { x: 20, y: 380, w: 150, h: 100 },
      style:
        '<p:style><a:lnRef idx="1"><a:srgbClr val="FF0000"/></a:lnRef>' +
        '<a:fillRef idx="1"><a:srgbClr val="FF0000"/></a:fillRef>' +
        '<a:effectRef idx="0"><a:srgbClr val="FF0000"/></a:effectRef>' +
        '<a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></p:style>',
    }),
    shape({
      id: shapeId(),
      name: 's-fill1-shade',
      rect: { x: 200, y: 380, w: 150, h: 100 },
      style:
        '<p:style><a:lnRef idx="1"><a:schemeClr val="accent1"><a:shade val="15000"/></a:schemeClr></a:lnRef>' +
        '<a:fillRef idx="1"><a:schemeClr val="accent1"><a:shade val="15000"/></a:schemeClr></a:fillRef>' +
        '<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>' +
        '<a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></p:style>',
    }),
    shape({
      id: shapeId(),
      name: 's-override',
      rect: { x: 380, y: 380, w: 150, h: 100 },
      fill: '<a:solidFill><a:srgbClr val="123456"/></a:solidFill>',
      style: STYLE(1, 1, 'accent1'),
    }),
    shape({
      id: shapeId(),
      name: 's-nofill-override',
      rect: { x: 560, y: 380, w: 150, h: 100 },
      fill: '<a:noFill/>',
      style: STYLE(1, 1, 'accent1'),
    }),
    shape({ id: shapeId(), name: 's-nostyle', rect: { x: 740, y: 380, w: 150, h: 100 } }),
  );

  const extra: readonly { id: string; q: string; expect: string | null }[] = [
    { id: 's-fill1-srgb', q: 'phClr may be a literal srgbClr', expect: 'FF0000' },
    {
      id: 's-fill1-shade',
      q: 'phClr carries its own transforms, as PowerPoint writes',
      expect: null,
    },
    { id: 's-override', q: 'an spPr fill beats the fillRef', expect: '123456' },
    { id: 's-nofill-override', q: 'an spPr noFill beats the fillRef', expect: null },
    { id: 's-nostyle', q: 'no fill and no style at all', expect: null },
  ];

  return {
    deck: 'style',
    hostile: false,
    pkg: {
      themes: [SCHEME_ONE],
      masters: [{ theme: 0, shapes: [] }],
      layouts: [{ master: 0, type: 'blank', name: 'Blank', shapes: [] }],
      slides: [{ layout: 0, name: 'styles', shapes }],
    },
    probes: [
      ...STYLE_CASES.map((c) => ({
        id: c.id,
        deck: 'style',
        kind: 'style' as const,
        slide: 1,
        question: c.q,
        expectFill: null,
      })),
      ...extra.map((c) => ({
        id: c.id,
        deck: 'style',
        kind: 'style' as const,
        slide: 1,
        question: c.q,
        expectFill: c.expect,
      })),
    ],
  };
}

/* -------------------------------------------------------------------------- */
/* deck: bgidx - what p:bgRef/@idx indexes                                    */
/* -------------------------------------------------------------------------- */

const BG_IDX = [1, 2, 3, 1001, 1002, 1003] as const;

/**
 * Six masters, one per index, because a master has exactly one background.
 *
 * ECMA describes a 1000-offset in prose that no two implementations read the
 * same way. Six numbers and six painted colours settle it, and the theme is
 * built so the colour names the list *and* the entry.
 *
 * Six identical themes rather than one shared one, because **a slide master
 * must own its theme part**: two masters pointing at the same `theme1.xml` is a
 * repair, six masters with six themes is not. Another rule found by bisection
 * rather than read anywhere, and recorded as a hostile probe below.
 */
function bgIdxDeck(): SheetDeck {
  const masters: MasterSpec[] = BG_IDX.map((idx, m) => ({
    theme: m,
    bg: `<p:bg><p:bgRef idx="${String(idx)}"><a:schemeClr val="accent1"/></p:bgRef></p:bg>`,
    shapes: [],
  }));
  const layouts: LayoutSpec[] = BG_IDX.map((idx, m) => ({
    master: m,
    type: 'blank',
    name: `bg${String(idx)}`,
    shapes: [],
  }));
  const slides: SlideSpec[] = BG_IDX.map((idx, m) => ({
    layout: m,
    name: `bg-${String(idx)}`,
    shapes: [],
  }));

  return {
    deck: 'bgidx',
    hostile: false,
    pkg: { themes: BG_IDX.map(() => SCHEME_ONE), masters, layouts, slides },
    probes: BG_IDX.map((idx, i) => ({
      id: `bg-${String(idx)}`,
      deck: 'bgidx',
      kind: 'background' as const,
      slide: i + 1,
      question: `p:bgRef idx=${String(idx)} against a theme whose six entries are distinguishable`,
      expectFill: null,
    })),
  };
}

/* -------------------------------------------------------------------------- */
/* deck: bgchain - which sheet's background wins                              */
/* -------------------------------------------------------------------------- */

function bgChainDeck(): SheetDeck {
  const solid = (hex: string): string =>
    `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${hex}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>`;

  return {
    deck: 'bgchain',
    hostile: false,
    pkg: {
      themes: [SCHEME_ONE],
      masters: [{ theme: 0, bg: solid('FF0000'), shapes: [] }],
      layouts: [
        { master: 0, type: 'blank', name: 'Layout with bg', bg: solid('00FF00'), shapes: [] },
        { master: 0, type: 'blank', name: 'Layout without bg', shapes: [] },
      ],
      slides: [
        { layout: 0, name: 'bc-layout' },
        { layout: 1, name: 'bc-master' },
        { layout: 0, name: 'bc-slide', bg: solid('0000FF') },
        {
          layout: 0,
          name: 'bc-nofill',
          bg: '<p:bg><p:bgPr><a:noFill/><a:effectLst/></p:bgPr></p:bg>',
        },
      ],
    },
    probes: [
      {
        id: 'bc-layout',
        deck: 'bgchain',
        kind: 'background',
        slide: 1,
        question: "the layout's background beats the master's",
        expectFill: '00FF00',
      },
      {
        id: 'bc-master',
        deck: 'bgchain',
        kind: 'background',
        slide: 2,
        question: 'a layout with no background falls through to the master',
        expectFill: 'FF0000',
      },
      {
        id: 'bc-slide',
        deck: 'bgchain',
        kind: 'background',
        slide: 3,
        question: "the slide's own background beats both",
        expectFill: '0000FF',
      },
      {
        id: 'bc-nofill',
        deck: 'bgchain',
        kind: 'background',
        slide: 4,
        question: 'a slide that declares noFill: does it inherit or paint nothing',
        expectFill: null,
      },
    ],
  };
}

/* -------------------------------------------------------------------------- */
/* deck: masters - two masters, two themes                                    */
/* -------------------------------------------------------------------------- */

function mastersDeck(): SheetDeck {
  const probeShapes = (tag: string): string[] => [
    shape({
      id: shapeId(),
      name: `${tag}-scheme`,
      rect: box('A0'),
      fill: '<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>',
      line: NO_LINE,
    }),
    shape({ id: shapeId(), name: `${tag}-style`, rect: box('A1'), style: STYLE(1, 1, 'accent1') }),
  ];

  return {
    deck: 'masters',
    hostile: false,
    pkg: {
      themes: [SCHEME_ONE, SCHEME_TWO],
      masters: [
        {
          theme: 0,
          bg: '<p:bg><p:bgRef idx="1001"><a:schemeClr val="accent1"/></p:bgRef></p:bg>',
          shapes: [],
        },
        {
          theme: 1,
          bg: '<p:bg><p:bgRef idx="1001"><a:schemeClr val="accent1"/></p:bgRef></p:bg>',
          shapes: [],
        },
      ],
      layouts: [
        { master: 0, type: 'blank', name: 'On one', shapes: [] },
        { master: 1, type: 'blank', name: 'On two', shapes: [] },
      ],
      slides: [
        { layout: 0, name: 'one', shapes: probeShapes('one') },
        { layout: 1, name: 'two', shapes: probeShapes('two') },
      ],
    },
    probes: [
      {
        id: 'one-scheme',
        deck: 'masters',
        kind: 'style',
        slide: 1,
        question: 'accent1 on the first master',
        expectFill: SCHEME_ONE.accent1,
      },
      {
        id: 'one-style',
        deck: 'masters',
        kind: 'style',
        slide: 1,
        question: 'fillRef on the first master',
        expectFill: SCHEME_ONE.accent1,
      },
      {
        id: 'two-scheme',
        deck: 'masters',
        kind: 'style',
        slide: 2,
        question:
          "accent1 on the second master, whose theme the presentation's own rel does not name",
        expectFill: SCHEME_TWO.accent1,
      },
      {
        id: 'two-style',
        deck: 'masters',
        kind: 'style',
        slide: 2,
        question: 'fillRef resolves against the second master style matrix',
        expectFill: SCHEME_TWO.accent1,
      },
      {
        id: 'two-bg',
        deck: 'masters',
        kind: 'background',
        slide: 2,
        question: "the second master's background uses the second theme",
        expectFill: null,
      },
      {
        id: 'one-bg',
        deck: 'masters',
        kind: 'background',
        slide: 1,
        question: "the first master's background",
        expectFill: null,
      },
    ],
  };
}

/* -------------------------------------------------------------------------- */
/* deck: clrmapovr - the colour map override                                  */
/* -------------------------------------------------------------------------- */

const OVERRIDE_MAP =
  '<a:overrideClrMapping bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent6"' +
  ' accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent1"' +
  ' hlink="hlink" folHlink="folHlink"/>';

/** A second override that disagrees with the first, which is the only way to
 *  find out which of the two a slide obeys. */
const OVERRIDE_MAP_2 =
  '<a:overrideClrMapping bg1="lt2" tx1="dk1" bg2="lt1" tx2="dk2" accent1="accent4"' +
  ' accent2="accent2" accent3="accent3" accent4="accent1" accent5="accent5" accent6="accent6"' +
  ' hlink="hlink" folHlink="folHlink"/>';

function clrMapOvrDeck(): SheetDeck {
  const mapped = (tag: string): string[] => [
    shape({
      id: shapeId(),
      name: `${tag}-accent1`,
      rect: box('A0'),
      fill: '<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>',
      line: NO_LINE,
    }),
    shape({
      id: shapeId(),
      name: `${tag}-bg1`,
      rect: box('A1'),
      fill: '<a:solidFill><a:schemeClr val="bg1"/></a:solidFill>',
      line: NO_LINE,
    }),
    shape({
      id: shapeId(),
      name: `${tag}-dk1`,
      rect: box('A2'),
      fill: '<a:solidFill><a:schemeClr val="dk1"/></a:solidFill>',
      line: NO_LINE,
    }),
  ];

  return {
    deck: 'clrmapovr',
    hostile: false,
    pkg: {
      themes: [SCHEME_ONE],
      masters: [{ theme: 0, shapes: [] }],
      layouts: [
        {
          master: 0,
          type: 'blank',
          name: 'Overriding layout',
          clrMapOvr: OVERRIDE_MAP,
          shapes: [],
        },
        { master: 0, type: 'blank', name: 'Plain layout', shapes: [] },
      ],
      slides: [
        { layout: 0, name: 'ov-inherit', shapes: mapped('ov') },
        // The slide's own override disagrees with the layout's on accent1:
        // the layout says accent6, this says accent4. Two identical overrides
        // cannot tell you which one applied.
        { layout: 0, name: 'ov-own', clrMapOvr: OVERRIDE_MAP_2, shapes: mapped('own') },
        { layout: 1, name: 'ov-plain', shapes: mapped('plain') },
      ],
    },
    probes: [
      {
        id: 'ov-accent1',
        deck: 'clrmapovr',
        kind: 'clrmap',
        slide: 1,
        question: "does a layout's overrideClrMapping reach the slides bound to it",
        expectFill: null,
      },
      {
        id: 'own-accent1',
        deck: 'clrmapovr',
        kind: 'clrmap',
        slide: 2,
        question:
          "the slide's override says accent4 and the layout's says accent6: which one applies",
        expectFill: null,
      },
      {
        id: 'own-bg1',
        deck: 'clrmapovr',
        kind: 'clrmap',
        slide: 2,
        question: "the slide's override maps bg1 to lt2 where the layout's maps it to dk1",
        expectFill: null,
      },
      {
        id: 'own-dk1',
        deck: 'clrmapovr',
        kind: 'clrmap',
        slide: 2,
        question: 'dk1 bypasses the override, as it bypasses the map',
        expectFill: SCHEME_ONE.dk1,
      },
      {
        id: 'plain-accent1',
        deck: 'clrmapovr',
        kind: 'clrmap',
        slide: 3,
        question: 'the control: no override anywhere',
        expectFill: SCHEME_ONE.accent1,
      },
    ],
  };
}

/* -------------------------------------------------------------------------- */
/* hostile probes - one package each                                          */
/* -------------------------------------------------------------------------- */

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function hostile(id: string, question: string, pkg: SheetPackage): SheetDeck {
  return {
    deck: id,
    hostile: true,
    pkg,
    probes: [{ id, deck: id, kind: 'hostile', slide: 1, question }],
  };
}

function plainMaster(): MasterSpec {
  return { theme: 0, shapes: [] };
}

function hostileDecks(): SheetDeck[] {
  const themes = [SCHEME_ONE];
  const layout: LayoutSpec = { master: 0, type: 'blank', name: 'Blank', shapes: [] };
  const one = (shapes: readonly string[] = []): SlideSpec => ({ layout: 0, name: 'h', shapes });

  const swatch = shape({
    id: 900,
    name: 'h-swatch',
    rect: box('A0'),
    fill: '<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>',
    line: NO_LINE,
  });

  return [
    hostile('h-no-layout-rel', 'a slide whose rels name no layout at all', {
      themes,
      masters: [plainMaster()],
      layouts: [layout],
      slides: [{ layout: 0, name: 'h', rels: [], shapes: [swatch] }],
    }),
    hostile('h-dangling-layout', 'a slide bound to a layout part that does not exist', {
      themes,
      masters: [plainMaster()],
      layouts: [layout],
      slides: [
        {
          layout: 0,
          name: 'h',
          rels: [
            { id: 'rId1', type: `${REL}/slideLayout`, target: '../slideLayouts/slideLayout9.xml' },
          ],
          shapes: [swatch],
        },
      ],
    }),
    hostile('h-two-layouts', 'a slide with two slideLayout relationships', {
      themes,
      masters: [plainMaster()],
      layouts: [layout, { master: 0, type: 'blank', name: 'Second', shapes: [] }],
      slides: [
        {
          layout: 0,
          name: 'h',
          rels: [
            { id: 'rId1', type: `${REL}/slideLayout`, target: '../slideLayouts/slideLayout1.xml' },
            { id: 'rId2', type: `${REL}/slideLayout`, target: '../slideLayouts/slideLayout2.xml' },
          ],
          shapes: [swatch],
        },
      ],
    }),
    hostile('h-layout-no-master', 'a layout whose rels name no master', {
      themes,
      masters: [plainMaster()],
      layouts: [{ master: 0, type: 'blank', name: 'Blank', rels: [], shapes: [] }],
      slides: [one([swatch])],
    }),
    hostile('h-master-no-theme', 'a master with no theme relationship', {
      themes,
      masters: [
        {
          theme: 0,
          rels: [
            { id: 'rId1', type: `${REL}/slideLayout`, target: '../slideLayouts/slideLayout1.xml' },
          ],
          shapes: [],
        },
      ],
      layouts: [layout],
      slides: [one([swatch])],
    }),
    hostile('h-unlisted-layout', 'a layout its master does not list in sldLayoutIdLst', {
      themes,
      masters: [plainMaster()],
      layouts: [{ ...layout, unlisted: true }],
      slides: [one([swatch])],
    }),
    hostile('h-ph-unknown-type', 'a p:ph with a type that is not in ST_PlaceholderType', {
      themes,
      masters: [plainMaster()],
      layouts: [layout],
      slides: [one([phProbe('h-ph-unknown-type', 'type="wibble" idx="1"')])],
    }),
    hostile('h-ph-idx-huge', 'a p:ph idx past the unsigned 32-bit range', {
      themes,
      masters: [plainMaster()],
      layouts: [layout],
      slides: [one([phProbe('h-ph-idx-huge', 'type="body" idx="4294967296"')])],
    }),
    hostile('h-ph-idx-negative', 'a negative p:ph idx', {
      themes,
      masters: [plainMaster()],
      layouts: [layout],
      slides: [one([phProbe('h-ph-idx-negative', 'type="body" idx="-1"')])],
    }),
    hostile('h-bgref-9999', 'a p:bgRef idx nothing can index', {
      themes,
      masters: [
        {
          theme: 0,
          bg: '<p:bg><p:bgRef idx="9999"><a:schemeClr val="accent1"/></p:bgRef></p:bg>',
          shapes: [],
        },
      ],
      layouts: [layout],
      slides: [one([swatch])],
    }),
    hostile('h-bgref-0', 'a p:bgRef idx of zero', {
      themes,
      masters: [
        {
          theme: 0,
          bg: '<p:bg><p:bgRef idx="0"><a:schemeClr val="accent1"/></p:bgRef></p:bg>',
          shapes: [],
        },
      ],
      layouts: [layout],
      slides: [one([swatch])],
    }),
    hostile('h-bgref-1000', 'a p:bgRef idx of exactly one thousand', {
      themes,
      masters: [
        {
          theme: 0,
          bg: '<p:bg><p:bgRef idx="1000"><a:schemeClr val="accent1"/></p:bgRef></p:bg>',
          shapes: [],
        },
      ],
      layouts: [layout],
      slides: [one([swatch])],
    }),
    hostile('h-fillref-4', 'a fillRef idx one past the three the theme declares', {
      themes,
      masters: [plainMaster()],
      layouts: [layout],
      slides: [
        one([
          shape({
            id: 901,
            name: 'h-fillref-4',
            rect: box('A0'),
            style: STYLE(4, 1, 'accent1'),
          }),
        ]),
      ],
    }),
    hostile('h-fontref-numeric', 'a fontRef idx spelled as a number rather than major or minor', {
      themes,
      masters: [plainMaster()],
      layouts: [layout],
      slides: [
        one([
          shape({
            id: 902,
            name: 'h-fontref-numeric',
            rect: box('A0'),
            style:
              '<p:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef>' +
              '<a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>' +
              '<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>' +
              '<a:fontRef idx="1"><a:schemeClr val="tx1"/></a:fontRef></p:style>',
          }),
        ]),
      ],
    }),
    hostile('h-dup-idx', 'a layout with two body placeholders at the same idx', {
      themes,
      masters: [plainMaster()],
      layouts: [
        {
          master: 0,
          name: 'Duplicate',
          shapes: [
            phBox('l-first', 'type="body" idx="1"', 'A1'),
            phBox('l-second', 'type="body" idx="1"', 'B1'),
          ],
        },
      ],
      slides: [one([phProbe('h-dup-idx', 'type="body" idx="1"')])],
    }),
    hostile('h-clrmap-crossed', 'a master colour map that crosses bg1 and tx1', {
      themes,
      masters: [
        {
          theme: 0,
          clrMap: {
            bg1: 'dk1',
            tx1: 'lt1',
            bg2: 'dk2',
            tx2: 'lt2',
            accent1: 'accent2',
            accent2: 'accent1',
            accent3: 'accent3',
            accent4: 'accent4',
            accent5: 'accent5',
            accent6: 'accent6',
            hlink: 'hlink',
            folHlink: 'folHlink',
          },
          shapes: [],
        },
      ],
      layouts: [layout],
      slides: [one([swatch])],
    }),
    // The four types a slide master may not carry. Found by bisecting four
    // clean-looking decks that came back REPAIRED; recorded here so the rule is
    // in the fixture rather than in somebody's memory of an afternoon.
    ...(['ctrTitle', 'subTitle', 'obj', 'pic'] as const).map((type) =>
      hostile(`h-master-${type}`, `a slide master carrying a ${type} placeholder`, {
        themes,
        masters: [
          {
            theme: 0,
            shapes: [
              phBox('m-title', 'type="title"', 'M0'),
              phBox('m-body', 'type="body" idx="1"', 'M1'),
              phBox(`m-${type}`, `type="${type}" idx="5"`, 'M2'),
            ],
          },
        ],
        layouts: [layout],
        slides: [one([swatch])],
      }),
    ),
    hostile('h-shared-theme', 'two slide masters pointing at the same theme part', {
      themes,
      masters: [plainMaster(), plainMaster()],
      layouts: [layout, { master: 1, type: 'blank', name: 'Second', shapes: [] }],
      slides: [one([swatch]), { layout: 1, name: 'h2', shapes: [] }],
    }),
    hostile('h-master-hdr', 'a slide master carrying an hdr placeholder', {
      themes,
      masters: [
        {
          theme: 0,
          shapes: [
            phBox('m-title', 'type="title"', 'M0'),
            phBox('m-hdr', 'type="hdr" idx="5"', 'M2'),
          ],
        },
      ],
      layouts: [layout],
      slides: [one([swatch])],
    }),
    hostile('h-slide-cycle', 'a layout bound to a master that lists a different layout', {
      themes,
      masters: [plainMaster()],
      layouts: [
        layout,
        {
          master: 0,
          type: 'blank',
          name: 'Points at itself',
          unlisted: true,
          rels: [
            {
              id: 'rId1',
              type: `${REL}/slideMaster`,
              target: '../slideLayouts/slideLayout2.xml',
            },
          ],
          shapes: [],
        },
      ],
      slides: [{ layout: 1, name: 'h', shapes: [swatch] }],
    }),
  ];
}

/* -------------------------------------------------------------------------- */
/* the experiment                                                             */
/* -------------------------------------------------------------------------- */

export function sheetDecks(): SheetDeck[] {
  nextId = 100;
  return [
    matchDeck(),
    match2Deck(),
    tier4Deck(),
    familyDeck(),
    twoHopDeck(),
    twoHop2Deck(),
    orphanDeck(),
    inheritDeck(),
    styleHopDeck(),
    styleDeck(),
    bgIdxDeck(),
    bgChainDeck(),
    mastersDeck(),
    clrMapOvrDeck(),
    ...hostileDecks(),
  ];
}
