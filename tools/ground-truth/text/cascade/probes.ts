/**
 * Experiment T1 - what PowerPoint does with text that states nothing.
 *
 * Sub-phase 3.1's whole subject is a run with no `a:rPr` and a paragraph with no
 * `a:pPr`: nine places could have declared the size, and the file names none of
 * them. Every implementation this project has read hard-codes an order for those
 * nine, and no two of the orders agree. So this asks.
 *
 * ## Why a size, and why through the object model
 *
 * A resolved font size is a number PowerPoint will tell you - `Font.Size` on a
 * run whose `a:rPr` is empty is the answer to the whole cascade, reported to a
 * quarter of a point, with no bitmap and no antialiasing in the way. Give each
 * of the nine levels a size no other level has and the reported number *names
 * the level that won*. That is the entire measurement, and the reason this
 * experiment is a table of sizes rather than a rendering comparison.
 *
 * Sizes are chosen four points apart, none of them 18, because 18 is the size a
 * `p:txBody` gets when nothing anywhere declares one and a probe that landed on
 * it by coincidence would be unreadable.
 *
 * ## The shape of the ladder
 *
 * Probe `order-k` declares levels k through 9 and omits everything above.  The
 * reported size names the winner among those present, so the nine probes
 * together pin the total order rather than merely confirming one guess.  Two of
 * the levels - `p:defaultTextStyle` and the theme's `a:objectDefaults` - live
 * outside any sheet, so those rungs need their own package rather than their own
 * slide, which is why this is nine decks and not nine slides.
 *
 * Every ladder slide carries a second shape that is **not** a placeholder, so
 * each package answers the placeholder question and the non-placeholder question
 * at once - the plan claims `p:defaultTextStyle` reaches only the second, and
 * that claim is worth one shape rather than one more deck.
 */

import {
  IDENTITY_CLR_MAP,
  SCHEME_ONE,
  shape,
  type LayoutSpec,
  type MasterSpec,
  type Rect,
  type SheetPackage,
  type SlideSpec,
} from './sheet-pptx.ts';

/* -------------------------------------------------------------------------- */
/* the ladder                                                                 */
/* -------------------------------------------------------------------------- */

/**
 * The nine sources, nearest first, with the size each one declares.
 *
 * The names are the plan's. The order below is the plan's *claim*, and the
 * experiment exists to refute it: nothing in the deck generator assumes the
 * order is right, and `analyse-text.ts` scores this order against every
 * transposition of it.
 */
export const SOURCES = [
  'run',
  'para',
  'shapeLst',
  'layoutPh',
  'masterPh',
  'txStyles',
  'defaultText',
  'objDefaults',
] as const;

export type Source = (typeof SOURCES)[number];

/** Points. Four apart, and none of them 18. */
export const LADDER_PT: Readonly<Record<Source, number>> = {
  run: 44,
  para: 40,
  shapeLst: 36,
  layoutPh: 32,
  masterPh: 28,
  txStyles: 24,
  defaultText: 20,
  objDefaults: 16,
};

/**
 * The three `p:txStyles` buckets, at three sizes.
 *
 * `txStyles` in the ladder above is the body bucket; a shape that reaches the
 * title or the other bucket instead reports 54 or 12 and says so.
 */
export const BUCKET_PT = { title: 54, body: 24, other: 12 } as const;

/** The theme's two font faces. Both ship with Windows, so neither substitutes. */
export const THEME_MAJOR = 'Georgia';
export const THEME_MINOR = 'Verdana';

/* -------------------------------------------------------------------------- */
/* XML helpers                                                                */
/* -------------------------------------------------------------------------- */

const hundredths = (pt: number): string => String(Math.round(pt * 100));

/** `a:defRPr` with the given attributes and children. */
export function defRPr(attrs: string, children = ''): string {
  const open = attrs === '' ? '<a:defRPr' : `<a:defRPr ${attrs}`;
  return children === '' ? `${open}/>` : `${open}>${children}</a:defRPr>`;
}

/** One `a:lvlNpPr`: its level, its own attributes, and its children. */
export interface LvlEntry {
  /** 1-based, as the element names are. */
  readonly level: number;
  /** Attributes on the `a:lvlNpPr` itself - `marL`, `indent`, `algn`. */
  readonly attrs?: string | undefined;
  /** Children, in `CT_TextParagraphProperties` order. */
  readonly inner?: string | undefined;
}

/** A `a:lvlNpPr`, 1-based as the element names are. */
export function lvlPr(entry: LvlEntry): string {
  if (entry.level < 1 || entry.level > 9) {
    throw new Error(`a:lvl${String(entry.level)}pPr does not exist`);
  }
  const tag = `a:lvl${String(entry.level)}pPr`;
  const open = `<${tag}${entry.attrs === undefined || entry.attrs === '' ? '' : ` ${entry.attrs}`}`;
  const inner = entry.inner ?? '';
  return inner === '' ? `${open}/>` : `${open}>${inner}</${tag}>`;
}

/**
 * The `a:lvlNpPr` sequence on its own, in level order.
 *
 * A `p:txStyles` bucket holds this sequence directly rather than an
 * `a:lstStyle` around it - `CT_TextListStyle` is the type of both, but only one
 * of the two is wrapped.
 */
export function lvlList(entries: readonly LvlEntry[]): string {
  return [...entries]
    .sort((a, b) => a.level - b.level)
    .map((entry) => lvlPr(entry))
    .join('');
}

/** An `a:lstStyle` holding one `a:lvlNpPr` per entry, in level order. */
export function lstStyle(entries: readonly LvlEntry[]): string {
  return `<a:lstStyle>${lvlList(entries)}</a:lstStyle>`;
}

/** The commonest case: one level, declaring only a size. */
export function sizeAt(level: number, pt: number): string {
  return lstStyle([{ level, inner: defRPr(`sz="${hundredths(pt)}"`) }]);
}

/**
 * `p:txStyles`, all three buckets.
 *
 * All three are written every time even when a probe only asks about one,
 * because a bucket that is absent and a bucket that is present but silent are
 * different states and the ladder needs the second.
 */
export function txStyles(buckets: {
  readonly title: string;
  readonly body: string;
  readonly other: string;
}): string {
  return (
    '<p:txStyles>' +
    `<p:titleStyle>${buckets.title}</p:titleStyle>` +
    `<p:bodyStyle>${buckets.body}</p:bodyStyle>` +
    `<p:otherStyle>${buckets.other}</p:otherStyle>` +
    '</p:txStyles>'
  );
}

/** One bucket declaring nothing but a size at level one. */
export function bucketSize(pt: number): string {
  return lvlList([{ level: 1, inner: defRPr(`sz="${hundredths(pt)}"`) }]);
}

/** The three buckets at the three sizes, one level each. */
export function bucketTxStyles(): string {
  return txStyles({
    title: bucketSize(BUCKET_PT.title),
    body: bucketSize(BUCKET_PT.body),
    other: bucketSize(BUCKET_PT.other),
  });
}

/** `p:defaultTextStyle`, one level. */
export function defaultTextStyle(pt: number): string {
  return `<p:defaultTextStyle>${bucketSize(pt)}</p:defaultTextStyle>`;
}

/**
 * `a:objectDefaults/a:spDef`, one level.
 *
 * `CT_DefaultShapeDefinition` sequences `a:spPr`, `a:bodyPr`, `a:lstStyle`, and
 * the first two are required, so an `a:lstStyle` on its own is a repair rather
 * than an experiment.
 */
export function objectDefaults(pt: number): string {
  return (
    '<a:objectDefaults><a:spDef><a:spPr/><a:bodyPr/>' +
    sizeAt(1, pt) +
    '</a:spDef></a:objectDefaults>'
  );
}

/** One `a:r` with text, and whatever `a:rPr` the probe wants. */
export function run(rPr: string, text = 'Ag'): string {
  return `<a:r>${rPr}<a:t>${text}</a:t></a:r>`;
}

/** An `a:p`. Children go in schema order: `a:pPr`, runs, `a:endParaRPr`. */
export function paragraph(opts: {
  readonly pPr?: string | undefined;
  readonly runs?: readonly string[] | undefined;
  readonly endParaRPr?: string | undefined;
}): string {
  return (
    '<a:p>' +
    (opts.pPr ?? '') +
    (opts.runs ?? [run('<a:rPr lang="en-US"/>')]).join('') +
    (opts.endParaRPr ?? '') +
    '</a:p>'
  );
}

/**
 * An `a:pPr`.
 *
 * `a:defRPr` sits after the bullet group and before `a:extLst` in
 * `CT_TextParagraphProperties`, and out of order it is a repair prompt rather
 * than a measurement.
 */
export function pPr(opts: {
  readonly attrs?: string | undefined;
  readonly lnSpc?: string | undefined;
  readonly spcBef?: string | undefined;
  readonly spcAft?: string | undefined;
  readonly bullet?: string | undefined;
  readonly defRPr?: string | undefined;
}): string {
  const inner =
    (opts.lnSpc ?? '') +
    (opts.spcBef ?? '') +
    (opts.spcAft ?? '') +
    (opts.bullet ?? '') +
    (opts.defRPr ?? '');
  const open = `<a:pPr${opts.attrs === undefined || opts.attrs === '' ? '' : ` ${opts.attrs}`}`;
  return inner === '' ? `${open}/>` : `${open}>${inner}</a:pPr>`;
}

/* -------------------------------------------------------------------------- */
/* the probe record                                                           */
/* -------------------------------------------------------------------------- */

/** The three `p:txStyles` buckets a package declares, in hundredths of a point. */
export interface BucketSizes {
  readonly title: number;
  readonly body: number;
  readonly other: number;
}

/**
 * Enough of a probe's package for the analysis to re-derive its answer.
 *
 * Only the two contested claims carry one. Scoring a candidate order against
 * the ladder means predicting what each rung would report *under that order*,
 * and that needs the declarations rather than the deck; the same is true of
 * scoring a candidate bucket rule against the crossed placeholders. The rest of
 * the probes assert one measured number each and need no simulation.
 */
export type Situation =
  | {
      readonly kind: 'ladder';
      /** Whether the probed shape is a placeholder. */
      readonly shape: 'ph' | 'plain';
      /** Hundredths of a point, per source, for the sources this package declares. */
      readonly declares: Readonly<Partial<Record<Source, number>>>;
      readonly txStyles: BucketSizes | null;
      readonly defaultTextStyle: number | null;
      readonly objectDefaults: number | null;
    }
  | {
      readonly kind: 'bucket';
      /** `@type` on the slide's own `p:ph`, or `null` when it has none. */
      readonly slideType: string | null;
      /** `@type` of the layout placeholder it matched, or `null`. */
      readonly layoutType: string | null;
      /** `@type` of the master placeholder that layout reached, or `null`. */
      readonly masterType: string | null;
      readonly txStyles: BucketSizes | null;
      readonly defaultTextStyle: number | null;
    };

export interface TextProbe {
  readonly id: string;
  readonly deck: string;
  readonly kind: string;
  /** 1-based, as the object model counts. */
  readonly slide: number;
  /** `p:cNvPr/@name`, which is how the reading is keyed back to the probe. */
  readonly shape: string;
  /** 1-based paragraph within the shape. */
  readonly paragraph: number;
  readonly question: string;
  /** The size the plan's claimed order predicts, in points. */
  readonly expectPt?: number | undefined;
  /** The typeface the plan's claimed order predicts. */
  readonly expectFont?: string | undefined;
  /** Anything else the probe asserts, as object-model field to value. */
  readonly expect?: Readonly<Record<string, string | number | boolean>> | undefined;
  /** For the two probes the analysis simulates rather than merely records. */
  readonly situation?: Situation | undefined;
}

export interface TextDeck {
  readonly deck: string;
  readonly pkg: SheetPackage;
  readonly probes: readonly TextProbe[];
}

/* -------------------------------------------------------------------------- */
/* geometry the probes share                                                  */
/* -------------------------------------------------------------------------- */

const PH_RECT: Rect = { x: 60, y: 60, w: 400, h: 120 };
const PLAIN_RECT: Rect = { x: 60, y: 260, w: 400, h: 120 };
const MASTER_RECT: Rect = { x: 60, y: 60, w: 840, h: 120 };

/** The six placeholder types a master may carry, with the rects they occupy. */
function ladderMaster(opts: {
  readonly phLstStyle?: string | undefined;
  readonly txStyles?: string | undefined;
}): MasterSpec {
  return {
    theme: 0,
    clrMap: IDENTITY_CLR_MAP,
    shapes: [
      shape({
        id: 2,
        name: 'm-title',
        rect: MASTER_RECT,
        ph: 'type="title"',
      }),
      shape({
        id: 3,
        name: 'm-body',
        rect: MASTER_RECT,
        ph: 'type="body" idx="1"',
        lstStyle: opts.phLstStyle,
      }),
    ],
    txStyles: opts.txStyles,
  };
}

function ladderLayout(phLstStyle: string | undefined): LayoutSpec {
  return {
    master: 0,
    name: 'Ladder',
    shapes: [
      shape({
        id: 2,
        name: 'l-body',
        rect: PH_RECT,
        ph: 'type="body" idx="1"',
        lstStyle: phLstStyle,
      }),
    ],
  };
}

/* -------------------------------------------------------------------------- */
/* deck: the ladder                                                           */
/* -------------------------------------------------------------------------- */

/**
 * One package per rung, each declaring the rung and everything below it.
 *
 * `from` is the index into `SOURCES` of the highest level still present, so
 * `from = 0` is the full ladder and `from = SOURCES.length` is a package that
 * declares no size anywhere - which is the ninth rung, and the only one whose
 * answer nobody can predict from the markup.
 */
function ladderDeck(from: number): TextDeck {
  const has = (source: Source): boolean => SOURCES.indexOf(source) >= from;
  const size = (source: Source): string | undefined =>
    has(source) ? sizeAt(1, LADDER_PT[source]) : undefined;

  const runPr = has('run')
    ? `<a:rPr lang="en-US" sz="${hundredths(LADDER_PT.run)}"/>`
    : '<a:rPr lang="en-US"/>';
  const paraPr = has('para')
    ? pPr({ defRPr: defRPr(`sz="${hundredths(LADDER_PT.para)}"`) })
    : undefined;
  const body = paragraph({ pPr: paraPr, runs: [run(runPr)] });

  const slide: SlideSpec = {
    layout: 0,
    name: `Ladder ${String(from)}`,
    shapes: [
      shape({
        id: 2,
        name: 's-ph',
        ph: 'type="body" idx="1"',
        lstStyle: size('shapeLst'),
        paragraphs: [body],
      }),
      shape({
        id: 3,
        name: 's-plain',
        rect: PLAIN_RECT,
        lstStyle: size('shapeLst'),
        paragraphs: [body],
      }),
    ],
  };

  const deck =
    from === SOURCES.length ? 'order-9-none' : `order-${String(from + 1)}-${SOURCES[from] ?? ''}`;

  // What the plan's order predicts for each of the two shapes. The placeholder
  // walks all nine; the plain shape has no layout or master placeholder to walk,
  // and the plan says it reaches `p:defaultTextStyle` where the placeholder does
  // not - so the two predictions diverge from rung four onwards, and that
  // divergence is the point of carrying both shapes.
  const phPrediction = ((): number | undefined => {
    for (const source of SOURCES.slice(from)) {
      if (source === 'defaultText') continue; // the plan says a placeholder skips it
      if (source === 'txStyles') return BUCKET_PT.body;
      return LADDER_PT[source];
    }
    return undefined;
  })();
  const plainPrediction = ((): number | undefined => {
    for (const source of SOURCES.slice(from)) {
      if (source === 'layoutPh' || source === 'masterPh') continue; // it has neither
      if (source === 'txStyles') return BUCKET_PT.other;
      return LADDER_PT[source];
    }
    return undefined;
  })();

  const declares: Partial<Record<Source, number>> = {};
  for (const source of SOURCES.slice(from)) {
    if (source === 'txStyles' || source === 'defaultText' || source === 'objDefaults') continue;
    declares[source] = Math.round(LADDER_PT[source] * 100);
  }
  const situationOf = (kind: 'ph' | 'plain'): Situation => ({
    kind: 'ladder',
    shape: kind,
    declares,
    txStyles: has('txStyles')
      ? {
          title: Math.round(BUCKET_PT.title * 100),
          body: Math.round(BUCKET_PT.body * 100),
          other: Math.round(BUCKET_PT.other * 100),
        }
      : null,
    defaultTextStyle: has('defaultText') ? Math.round(LADDER_PT.defaultText * 100) : null,
    objectDefaults: has('objDefaults') ? Math.round(LADDER_PT.objDefaults * 100) : null,
  });

  return {
    deck,
    pkg: {
      themes: [
        {
          scheme: SCHEME_ONE,
          majorLatin: THEME_MAJOR,
          minorLatin: THEME_MINOR,
          ...(has('objDefaults') ? { objectDefaults: objectDefaults(LADDER_PT.objDefaults) } : {}),
        },
      ],
      masters: [
        ladderMaster({
          phLstStyle: size('masterPh'),
          ...(has('txStyles') ? { txStyles: bucketTxStyles() } : {}),
        }),
      ],
      layouts: [ladderLayout(size('layoutPh'))],
      slides: [slide],
      ...(has('defaultText') ? { defaultTextStyle: defaultTextStyle(LADDER_PT.defaultText) } : {}),
    },
    probes: [
      {
        id: `${deck}-ph`,
        deck,
        kind: 'order',
        slide: 1,
        shape: 's-ph',
        paragraph: 1,
        question: `a body placeholder, with the cascade declared from "${SOURCES[from] ?? 'nowhere'}" down`,
        ...(phPrediction === undefined ? {} : { expectPt: phPrediction }),
        situation: situationOf('ph'),
      },
      {
        id: `${deck}-plain`,
        deck,
        kind: 'order',
        slide: 1,
        shape: 's-plain',
        paragraph: 1,
        question: 'a shape that is not a placeholder, same package',
        ...(plainPrediction === undefined ? {} : { expectPt: plainPrediction }),
        situation: situationOf('plain'),
      },
    ],
  };
}

/* -------------------------------------------------------------------------- */
/* deck: which bucket                                                         */
/* -------------------------------------------------------------------------- */

/**
 * Every `ST_PlaceholderType`, one per slide, against the three buckets.
 *
 * The layout carries the same type at the same index, so the first hop matches
 * and nothing between the slide and the master declares a size. The size that
 * comes back names the bucket, and the plan's rule - title and ctrTitle to the
 * title bucket, the header-and-footer trio and a non-placeholder to the other
 * bucket, everything else to the body bucket - is scored against all sixteen.
 */
const ALL_PH_TYPES = [
  'title',
  'ctrTitle',
  'subTitle',
  'body',
  'obj',
  'chart',
  'tbl',
  'clipArt',
  'dgm',
  'media',
  'sldImg',
  'pic',
  'sldNum',
  'hdr',
  'ftr',
  'dt',
] as const;

/** The plan's claimed bucket for a placeholder type. */
function claimedBucket(type: string): keyof typeof BUCKET_PT {
  if (type === 'title' || type === 'ctrTitle') return 'title';
  if (type === 'dt' || type === 'ftr' || type === 'sldNum') return 'other';
  return 'body';
}

/**
 * A master carrying all six types a master may hold, so the second hop lands.
 *
 * `ladderMaster` carries only a title and a body, which is enough for the rungs
 * and not enough here: a layout `ftr` whose master has no footer ends its chain
 * at the layout, and a bucket measured on that chain would be answering a
 * different question from the one asked.
 */
function bucketMaster(): MasterSpec {
  return {
    theme: 0,
    clrMap: IDENTITY_CLR_MAP,
    shapes: MASTER_PH_TYPES.map((type, i) =>
      shape({
        id: i + 2,
        name: `m-${type}`,
        rect: MASTER_RECT,
        ph: type === 'title' ? 'type="title"' : `type="${type}" idx="${String(i)}"`,
      }),
    ),
    txStyles: bucketTxStyles(),
  };
}

const MASTER_PH_TYPES = ['title', 'body', 'dt', 'ftr', 'sldNum', 'hdr'] as const;

/**
 * The type a placeholder presents to a master, measured in 2.9.
 *
 * Repeated here rather than imported because the analysis must be able to say
 * what the *measurement* found without the model under test in the loop.
 */
function folded(type: string): string {
  if (type === 'title' || type === 'ctrTitle') return 'title';
  return (MASTER_PH_TYPES as readonly string[]).includes(type) ? type : 'body';
}

const DEFAULT_TEXT_HUNDREDTHS = Math.round(LADDER_PT.defaultText * 100);

/** The bucket sizes every bucket deck declares, in hundredths. */
const BUCKET_SIZES: BucketSizes = {
  title: Math.round(BUCKET_PT.title * 100),
  body: Math.round(BUCKET_PT.body * 100),
  other: Math.round(BUCKET_PT.other * 100),
};

/**
 * One deck per placeholder type.
 *
 * Sixteen types on one slide is the arrangement PowerPoint repaired on the first
 * run, and a repaired package answers a question about the file PowerPoint wrote
 * rather than the one we did. One question per package, so a refusal names its
 * own cause.
 */
function bucketDecks(): readonly TextDeck[] {
  return ALL_PH_TYPES.map((type) => {
    const deck = `bucket-${type}`;
    return {
      deck,
      pkg: {
        themes: [{ scheme: SCHEME_ONE, majorLatin: THEME_MAJOR, minorLatin: THEME_MINOR }],
        masters: [bucketMaster()],
        layouts: [
          {
            master: 0,
            name: `Layout ${type}`,
            shapes: [
              shape({ id: 2, name: `l-${type}`, rect: PH_RECT, ph: `type="${type}" idx="0"` }),
            ],
          },
        ],
        slides: [
          {
            layout: 0,
            name: `Bucket ${type}`,
            shapes: [
              shape({
                id: 2,
                name: 's-ph',
                ph: `type="${type}" idx="0"`,
                paragraphs: [paragraph({})],
              }),
              shape({ id: 3, name: 's-plain', rect: PLAIN_RECT, paragraphs: [paragraph({})] }),
            ],
          },
        ],
      },
      probes: [
        {
          id: `bucket-${type}`,
          deck,
          kind: 'bucket',
          slide: 1,
          shape: 's-ph',
          paragraph: 1,
          question: `which p:txStyles bucket a "${type}" placeholder reads`,
          expectPt: BUCKET_PT[claimedBucket(type)],
          situation: {
            kind: 'bucket',
            slideType: type,
            layoutType: type,
            masterType: folded(type),
            txStyles: BUCKET_SIZES,
            defaultTextStyle: null,
          },
        },
        {
          id: `bucket-${type}-plain`,
          deck,
          kind: 'bucket',
          slide: 1,
          shape: 's-plain',
          paragraph: 1,
          question:
            'a shape that is not a placeholder, in a package that declares p:otherStyle and no p:defaultTextStyle',
          expectPt: BUCKET_PT.other,
          situation: {
            kind: 'bucket',
            slideType: null,
            layoutType: null,
            masterType: null,
            txStyles: BUCKET_SIZES,
            defaultTextStyle: null,
          },
        },
      ],
    };
  });
}

/**
 * Whose `@type` picks the bucket - the slide's, or the one it inherited from.
 *
 * The first hop matches on `@idx` alone, measured in 2.9, so a slide `title` can
 * and does land on a layout `body`. If the bucket follows the slide's own type
 * the two probes below report 54 and 24; if it follows the sheet the value was
 * found on, they report 24 and 54; and no other reading produces either pair.
 */
const CROSSED = [
  { slide: 'title', layout: 'body', bucket: 'body' },
  { slide: 'body', layout: 'title', bucket: 'title' },
  { slide: 'title', layout: 'sldNum', bucket: 'other' },
  { slide: 'sldNum', layout: 'body', bucket: 'body' },
] as const;

function crossedBucketDeck(): TextDeck {
  return {
    deck: 'bucket-crossed',
    pkg: {
      themes: [{ scheme: SCHEME_ONE, majorLatin: THEME_MAJOR, minorLatin: THEME_MINOR }],
      masters: [bucketMaster()],
      layouts: [
        {
          master: 0,
          name: 'Crossed',
          shapes: CROSSED.map((cross, i) =>
            shape({
              id: i + 2,
              name: `l-${cross.layout}${String(i)}`,
              rect: { x: 60 + i * 220, y: 60, w: 200, h: 100 },
              ph: `type="${cross.layout}" idx="${String(i)}"`,
            }),
          ),
        },
      ],
      slides: [
        {
          layout: 0,
          name: 'Crossed',
          shapes: CROSSED.map((cross, i) =>
            shape({
              id: i + 2,
              name: `s-${cross.slide}-on-${cross.layout}`,
              ph: `type="${cross.slide}" idx="${String(i)}"`,
              paragraphs: [paragraph({})],
            }),
          ),
        },
      ],
    },
    probes: CROSSED.map((cross, i) => ({
      id: `bucket-crossed-${cross.slide}-on-${cross.layout}`,
      deck: 'bucket-crossed',
      kind: 'bucket',
      slide: 1,
      shape: `s-${cross.slide}-on-${cross.layout}`,
      paragraph: 1,
      question:
        `a slide ${cross.slide} at idx ${String(i)} matching a layout ${cross.layout}: ` +
        'whose @type picks the bucket, the shape asking or the placeholder it inherited from',
      expectPt: BUCKET_PT[claimedBucket(cross.slide)],
      situation: {
        kind: 'bucket',
        slideType: cross.slide,
        layoutType: cross.layout,
        masterType: folded(cross.layout),
        txStyles: BUCKET_SIZES,
        defaultTextStyle: null,
      },
    })),
  };
}

/* -------------------------------------------------------------------------- */
/* deck: levels                                                               */
/* -------------------------------------------------------------------------- */

/** `a:lvl1pPr`..`a:lvl9pPr` at nine sizes, read back through `@lvl` 0..8. */
const LEVEL_PT = [8, 12, 16, 20, 24, 28, 32, 36, 40] as const;
const OTHER_LEVEL_PT = [9, 13, 17, 21, 25, 29, 33, 37, 41] as const;

function levelEntries(sizes: readonly number[]): readonly LvlEntry[] {
  return sizes.map((pt, i) => ({ level: i + 1, inner: defRPr(`sz="${hundredths(pt)}"`) }));
}

function levelDeck(): TextDeck {
  const nine = (): readonly string[] =>
    LEVEL_PT.map((_, i) => paragraph({ pPr: pPr({ attrs: `lvl="${String(i)}"` }) }));

  // A shape `a:lstStyle` that declares only level three. A paragraph at `lvl="2"`
  // takes it; a paragraph at `lvl="0"` must not, and a cascade that quietly falls
  // back to `a:lvl1pPr` when the asked-for level is absent would give both 64.
  const sparse = lstStyle([{ level: 3, inner: defRPr('sz="6400"') }]);

  return {
    deck: 'level',
    pkg: {
      themes: [{ scheme: SCHEME_ONE, majorLatin: THEME_MAJOR, minorLatin: THEME_MINOR }],
      masters: [
        {
          theme: 0,
          clrMap: IDENTITY_CLR_MAP,
          shapes: [
            shape({ id: 2, name: 'm-title', rect: MASTER_RECT, ph: 'type="title"' }),
            shape({ id: 3, name: 'm-body', rect: MASTER_RECT, ph: 'type="body" idx="1"' }),
          ],
          txStyles: txStyles({
            title: bucketSize(BUCKET_PT.title),
            body: lvlList(levelEntries(LEVEL_PT)),
            other: lvlList(levelEntries(OTHER_LEVEL_PT)),
          }),
        },
      ],
      layouts: [ladderLayout(undefined)],
      slides: [
        {
          layout: 0,
          name: 'Levels',
          shapes: [
            shape({ id: 2, name: 's-ph', ph: 'type="body" idx="1"', paragraphs: nine() }),
            shape({ id: 3, name: 's-plain', rect: PLAIN_RECT, paragraphs: nine() }),
            shape({
              id: 4,
              name: 's-sparse',
              ph: 'type="body" idx="1"',
              lstStyle: sparse,
              paragraphs: [
                paragraph({ pPr: pPr({ attrs: 'lvl="0"' }) }),
                paragraph({ pPr: pPr({ attrs: 'lvl="2"' }) }),
              ],
            }),
          ],
        },
      ],
    },
    probes: [
      ...LEVEL_PT.map((pt, i) => ({
        id: `level-body-${String(i)}`,
        deck: 'level',
        kind: 'level',
        slide: 1,
        shape: 's-ph',
        paragraph: i + 1,
        question: `a paragraph at lvl="${String(i)}" against p:bodyStyle/a:lvl${String(i + 1)}pPr`,
        expectPt: pt,
      })),
      ...OTHER_LEVEL_PT.map((pt, i) => ({
        id: `level-other-${String(i)}`,
        deck: 'level',
        kind: 'level',
        slide: 1,
        shape: 's-plain',
        paragraph: i + 1,
        question: `a non-placeholder paragraph at lvl="${String(i)}" against p:otherStyle`,
        expectPt: pt,
      })),
      {
        id: 'level-sparse-0',
        deck: 'level',
        kind: 'level',
        slide: 1,
        shape: 's-sparse',
        paragraph: 1,
        question: 'lvl="0" where the shape lstStyle declares only a:lvl3pPr',
        expectPt: LEVEL_PT[0],
      },
      {
        id: 'level-sparse-2',
        deck: 'level',
        kind: 'level',
        slide: 1,
        shape: 's-sparse',
        paragraph: 2,
        question: 'lvl="2" where the shape lstStyle declares only a:lvl3pPr',
        expectPt: 64,
      },
    ],
  };
}

/**
 * Does a level inherit level by level, or does the nearest sheet take all nine?
 *
 * The layout placeholder declares `a:lvl2pPr` and nothing else. If levels are
 * independent, a paragraph at `lvl="1"` takes the layout's 50 while its
 * neighbours at `lvl="0"` and `lvl="2"` still take the master's 8 and 16. If the
 * whole `a:lstStyle` is taken from the first sheet that has one - which is what
 * a `??` between two list styles produces, and it is an easy thing to write -
 * the other two paragraphs lose their sizes entirely.
 */
function levelCrossDeck(): TextDeck {
  return {
    deck: 'level-cross',
    pkg: {
      themes: [{ scheme: SCHEME_ONE, majorLatin: THEME_MAJOR, minorLatin: THEME_MINOR }],
      masters: [
        {
          theme: 0,
          clrMap: IDENTITY_CLR_MAP,
          shapes: [
            shape({ id: 2, name: 'm-title', rect: MASTER_RECT, ph: 'type="title"' }),
            shape({ id: 3, name: 'm-body', rect: MASTER_RECT, ph: 'type="body" idx="1"' }),
          ],
          txStyles: txStyles({
            title: bucketSize(BUCKET_PT.title),
            body: lvlList(levelEntries([8, 12, 16])),
            other: bucketSize(BUCKET_PT.other),
          }),
        },
      ],
      layouts: [ladderLayout(lstStyle([{ level: 2, inner: defRPr('sz="5000"') }]))],
      slides: [
        {
          layout: 0,
          name: 'Level cross',
          shapes: [
            shape({
              id: 2,
              name: 's-ph',
              ph: 'type="body" idx="1"',
              paragraphs: [0, 1, 2].map((lvl) =>
                paragraph({ pPr: pPr({ attrs: `lvl="${String(lvl)}"` }) }),
              ),
            }),
          ],
        },
      ],
    },
    probes: [0, 1, 2].map((lvl) => ({
      id: `level-cross-${String(lvl)}`,
      deck: 'level-cross',
      kind: 'level',
      slide: 1,
      shape: 's-ph',
      paragraph: lvl + 1,
      question: `lvl="${String(lvl)}" where the layout placeholder declares only a:lvl2pPr`,
      expectPt: [8, 50, 16][lvl] ?? 0,
    })),
  };
}

/**
 * Does a nearer list style shadow the levels it does not declare?
 *
 * The layout placeholder declares `a:lvl1pPr` and nothing else, and the master's
 * `p:bodyStyle` declares all three. A paragraph at `lvl="1"` therefore has a
 * nearer style that speaks - but not about the level being asked for. If the
 * levels really are independent it takes the master's second level; if a style
 * that declares *anything* shadows all nine, it takes the layout's first, and
 * every second-level bullet in the deck is the size of a first-level one.
 *
 * `level-cross` looks like this experiment and is not: there the nearer style
 * declares the level being asked for, which both readings answer the same way.
 * This deck exists because a mutation that fell back to `a:lvl1pPr` survived the
 * whole suite - the difference had never been measured.
 */
function levelShadowDeck(): TextDeck {
  const layoutLvl1 = lstStyle([
    { level: 1, attrs: 'marL="914400"', inner: defRPr(`sz="${hundredths(50)}"`) },
  ]);
  const masterBody = lvlList([
    { level: 1, attrs: 'marL="228600"', inner: defRPr(`sz="${hundredths(8)}"`) },
    { level: 2, attrs: 'marL="457200"', inner: defRPr(`sz="${hundredths(12)}"`) },
    { level: 3, attrs: 'marL="685800"', inner: defRPr(`sz="${hundredths(16)}"`) },
  ]);
  return {
    deck: 'level-shadow',
    pkg: {
      themes: [{ scheme: SCHEME_ONE, majorLatin: THEME_MAJOR, minorLatin: THEME_MINOR }],
      masters: [
        {
          theme: 0,
          clrMap: IDENTITY_CLR_MAP,
          shapes: [
            shape({ id: 2, name: 'm-title', rect: MASTER_RECT, ph: 'type="title"' }),
            shape({ id: 3, name: 'm-body', rect: MASTER_RECT, ph: 'type="body" idx="1"' }),
          ],
          txStyles: txStyles({
            title: bucketSize(BUCKET_PT.title),
            body: masterBody,
            other: bucketSize(BUCKET_PT.other),
          }),
        },
      ],
      layouts: [ladderLayout(layoutLvl1)],
      slides: [
        {
          layout: 0,
          name: 'Level shadow',
          shapes: [
            shape({
              id: 2,
              name: 's-ph',
              ph: 'type="body" idx="1"',
              paragraphs: [0, 1, 2].map((lvl) =>
                paragraph({ pPr: pPr({ attrs: `lvl="${String(lvl)}"` }) }),
              ),
            }),
          ],
        },
      ],
    },
    probes: [0, 1, 2].map((lvl) => ({
      id: `level-shadow-${String(lvl)}`,
      deck: 'level-shadow',
      kind: 'level',
      slide: 1,
      shape: 's-ph',
      paragraph: lvl + 1,
      question:
        `lvl="${String(lvl)}" where the layout placeholder declares a:lvl1pPr only ` +
        'and the master declares all three',
      expectPt: [50, 12, 16][lvl] ?? 0,
    })),
  };
}

/**
 * `a:defPPr`, the tenth element of `CT_TextListStyle` and the one nobody reads.
 *
 * It sits before `a:lvl1pPr` in every list style in the format - a shape's, a
 * master's `p:txStyles` bucket, `p:defaultTextStyle` - and every implementation
 * this project has read skips straight to the numbered levels. Either it is dead
 * markup or it is a source, and the difference is one deck.
 */
function defPPrDeck(): TextDeck {
  const defPPr = (pt: number): string =>
    '<a:defPPr>' + defRPr(`sz="${hundredths(pt)}"`) + '</a:defPPr>';
  const bothLevels = lvlList([{ level: 1, inner: defRPr('sz="3400"') }]);
  return {
    deck: 'defppr',
    pkg: {
      themes: [{ scheme: SCHEME_ONE, majorLatin: THEME_MAJOR, minorLatin: THEME_MINOR }],
      masters: [ladderMaster({ txStyles: bucketTxStyles() })],
      layouts: [ladderLayout(undefined)],
      slides: [
        {
          layout: 0,
          name: 'defPPr',
          shapes: [
            shape({
              id: 2,
              name: 's-shape-defppr',
              rect: PH_RECT,
              lstStyle: `<a:lstStyle>${defPPr(30)}</a:lstStyle>`,
              paragraphs: [paragraph({})],
            }),
            shape({
              id: 3,
              name: 's-shape-both',
              rect: PLAIN_RECT,
              lstStyle: '<a:lstStyle>' + defPPr(30) + bothLevels + '</a:lstStyle>',
              paragraphs: [paragraph({})],
            }),
            shape({
              id: 4,
              name: 's-package-defppr',
              rect: { x: 500, y: 60, w: 400, h: 120 },
              paragraphs: [paragraph({})],
            }),
          ],
        },
      ],
      defaultTextStyle: `<p:defaultTextStyle>${defPPr(26)}</p:defaultTextStyle>`,
    },
    probes: [
      {
        id: 'defppr-shape',
        deck: 'defppr',
        kind: 'defppr',
        slide: 1,
        shape: 's-shape-defppr',
        paragraph: 1,
        question: 'a shape a:lstStyle declaring only a:defPPr',
      },
      {
        id: 'defppr-shape-both',
        deck: 'defppr',
        kind: 'defppr',
        slide: 1,
        shape: 's-shape-both',
        paragraph: 1,
        question: 'a shape a:lstStyle declaring a:defPPr at 30 and a:lvl1pPr at 34',
        expectPt: 34,
      },
      {
        id: 'defppr-package',
        deck: 'defppr',
        kind: 'defppr',
        slide: 1,
        shape: 's-package-defppr',
        paragraph: 1,
        question: 'p:defaultTextStyle declaring only a:defPPr',
      },
    ],
  };
}

/**
 * What PowerPoint substitutes when the master declares no `p:txStyles` at all.
 *
 * Rungs seven, eight and nine of the ladder all reported 28 points for a body
 * placeholder, which is not a size any of them declares - so there is a tenth
 * source below every source the plan names, and it is PowerPoint's own. Nine
 * paragraph levels on a title, a body and a shape that is not a placeholder
 * reads the whole of it out in one deck rather than inferring it from one
 * number.
 */
function builtinDeck(): TextDeck {
  const nine = (): readonly string[] =>
    [0, 1, 2, 3, 4, 5, 6, 7, 8].map((lvl) =>
      paragraph({ pPr: pPr({ attrs: `lvl="${String(lvl)}"` }) }),
    );
  return {
    deck: 'builtin',
    pkg: {
      themes: [{ scheme: SCHEME_ONE, majorLatin: THEME_MAJOR, minorLatin: THEME_MINOR }],
      masters: [ladderMaster({})],
      layouts: [
        {
          master: 0,
          name: 'Builtin',
          shapes: [
            shape({ id: 2, name: 'l-title', rect: PH_RECT, ph: 'type="title" idx="0"' }),
            shape({ id: 3, name: 'l-body', rect: PLAIN_RECT, ph: 'type="body" idx="1"' }),
          ],
        },
      ],
      slides: [
        {
          layout: 0,
          name: 'Builtin',
          shapes: [
            shape({ id: 2, name: 's-title', ph: 'type="title" idx="0"', paragraphs: nine() }),
            shape({ id: 3, name: 's-body', ph: 'type="body" idx="1"', paragraphs: nine() }),
            shape({
              id: 4,
              name: 's-plain',
              rect: { x: 500, y: 60, w: 400, h: 400 },
              paragraphs: nine(),
            }),
          ],
        },
      ],
    },
    probes: [
      ...[0, 1, 2, 3, 4, 5, 6, 7, 8].map((lvl) => ({
        id: `builtin-title-${String(lvl)}`,
        deck: 'builtin',
        kind: 'builtin',
        slide: 1,
        shape: 's-title',
        paragraph: lvl + 1,
        question: `a title placeholder at lvl="${String(lvl)}" with no p:txStyles anywhere`,
      })),
      ...[0, 1, 2, 3, 4, 5, 6, 7, 8].map((lvl) => ({
        id: `builtin-body-${String(lvl)}`,
        deck: 'builtin',
        kind: 'builtin',
        slide: 1,
        shape: 's-body',
        paragraph: lvl + 1,
        question: `a body placeholder at lvl="${String(lvl)}" with no p:txStyles anywhere`,
      })),
      ...[0, 1, 2, 3, 4, 5, 6, 7, 8].map((lvl) => ({
        id: `builtin-plain-${String(lvl)}`,
        deck: 'builtin',
        kind: 'builtin',
        slide: 1,
        shape: 's-plain',
        paragraph: lvl + 1,
        question: `a shape that is not a placeholder at lvl="${String(lvl)}"`,
      })),
    ],
  };
}

/**
 * Is `p:defaultTextStyle` blocked by being a placeholder, or by having a bucket?
 *
 * The header-and-footer trio reads no `p:txStyles` bucket - measured, and
 * against the plan, which assigns them `p:otherStyle`. So a `sldNum` in a
 * package that also declares `p:defaultTextStyle` separates the two readings:
 * 20 points means the rule is "a shape with no bucket falls through to the
 * package default", 18 means it is "a placeholder never reads the package
 * default", and the two rules disagree on every slide-number box in the world.
 */
function hfDefaultDeck(): TextDeck {
  return {
    deck: 'hf-default',
    pkg: {
      themes: [{ scheme: SCHEME_ONE, majorLatin: THEME_MAJOR, minorLatin: THEME_MINOR }],
      masters: [bucketMaster()],
      layouts: [
        {
          master: 0,
          name: 'HF',
          shapes: [
            shape({ id: 2, name: 'l-sldNum', rect: PH_RECT, ph: 'type="sldNum" idx="0"' }),
            shape({ id: 3, name: 'l-body', rect: PLAIN_RECT, ph: 'type="body" idx="1"' }),
          ],
        },
      ],
      slides: [
        {
          layout: 0,
          name: 'HF',
          shapes: [
            shape({
              id: 2,
              name: 's-sldNum',
              ph: 'type="sldNum" idx="0"',
              paragraphs: [paragraph({})],
            }),
            shape({
              id: 3,
              name: 's-body',
              ph: 'type="body" idx="1"',
              paragraphs: [paragraph({})],
            }),
            shape({
              id: 4,
              name: 's-plain',
              rect: { x: 500, y: 60, w: 400, h: 120 },
              paragraphs: [paragraph({})],
            }),
          ],
        },
      ],
      defaultTextStyle: defaultTextStyle(LADDER_PT.defaultText),
    },
    probes: [
      {
        id: 'hf-default-sldNum',
        deck: 'hf-default',
        kind: 'bucket',
        slide: 1,
        shape: 's-sldNum',
        paragraph: 1,
        question: 'a sldNum placeholder in a package that declares p:defaultTextStyle',
        situation: {
          kind: 'bucket',
          slideType: 'sldNum',
          layoutType: 'sldNum',
          masterType: 'sldNum',
          txStyles: BUCKET_SIZES,
          defaultTextStyle: DEFAULT_TEXT_HUNDREDTHS,
        },
      },
      {
        id: 'hf-default-body',
        deck: 'hf-default',
        kind: 'bucket',
        slide: 1,
        shape: 's-body',
        paragraph: 1,
        question: 'a body placeholder in the same package, which does reach p:bodyStyle',
        expectPt: BUCKET_PT.body,
        situation: {
          kind: 'bucket',
          slideType: 'body',
          layoutType: 'body',
          masterType: 'body',
          txStyles: BUCKET_SIZES,
          defaultTextStyle: DEFAULT_TEXT_HUNDREDTHS,
        },
      },
      {
        id: 'hf-default-plain',
        deck: 'hf-default',
        kind: 'bucket',
        slide: 1,
        shape: 's-plain',
        paragraph: 1,
        question: 'a shape that is not a placeholder in the same package',
        expectPt: LADDER_PT.defaultText,
        situation: {
          kind: 'bucket',
          slideType: null,
          layoutType: null,
          masterType: null,
          txStyles: BUCKET_SIZES,
          defaultTextStyle: DEFAULT_TEXT_HUNDREDTHS,
        },
      },
    ],
  };
}

/**
 * Does a bucket that says nothing about a property hand on, or stop?
 *
 * `p:bodyStyle/a:lvl1pPr` here declares `marL` and no size, and the package
 * declares `p:defaultTextStyle` at 20 points. If the bucket merely contributes
 * what it has and the cascade continues, the paragraph is 20 points. If reaching
 * a bucket ends the walk, it is whatever the absolute floor turns out to be -
 * and rung seven of the ladder already showed a body placeholder ignoring a
 * `p:defaultTextStyle` that was right there, so the two readings differ on every
 * partially-specified master in the world.
 */
function partialBucketDeck(): TextDeck {
  return {
    deck: 'bucket-partial',
    pkg: {
      themes: [{ scheme: SCHEME_ONE, majorLatin: THEME_MAJOR, minorLatin: THEME_MINOR }],
      masters: [
        ladderMaster({
          txStyles: txStyles({
            title: bucketSize(BUCKET_PT.title),
            body: lvlList([{ level: 1, attrs: `marL="${String(MAR_L)}"` }]),
            other: bucketSize(BUCKET_PT.other),
          }),
        }),
      ],
      layouts: [ladderLayout(undefined)],
      slides: [
        {
          layout: 0,
          name: 'Partial',
          shapes: [
            shape({ id: 2, name: 's-ph', ph: 'type="body" idx="1"', paragraphs: [paragraph({})] }),
            shape({ id: 3, name: 's-plain', rect: PLAIN_RECT, paragraphs: [paragraph({})] }),
          ],
        },
      ],
      defaultTextStyle: defaultTextStyle(LADDER_PT.defaultText),
    },
    probes: [
      {
        id: 'bucket-partial-ph',
        deck: 'bucket-partial',
        kind: 'bucket',
        slide: 1,
        shape: 's-ph',
        paragraph: 1,
        question:
          'a body placeholder whose p:bodyStyle declares marL and no size, in a package that declares p:defaultTextStyle',
        expect: { leftIndentPt: MAR_L / 12700 },
        situation: {
          kind: 'bucket',
          slideType: 'body',
          layoutType: 'body',
          masterType: 'body',
          // The bucket exists and declares no size, which is the whole point.
          txStyles: { title: BUCKET_SIZES.title, body: 0, other: BUCKET_SIZES.other },
          defaultTextStyle: DEFAULT_TEXT_HUNDREDTHS,
        },
      },
      {
        id: 'bucket-partial-plain',
        deck: 'bucket-partial',
        kind: 'bucket',
        slide: 1,
        shape: 's-plain',
        paragraph: 1,
        question: 'the same package, from a shape with no bucket at all',
        expectPt: LADDER_PT.defaultText,
        situation: {
          kind: 'bucket',
          slideType: null,
          layoutType: null,
          masterType: null,
          txStyles: { title: BUCKET_SIZES.title, body: 0, other: BUCKET_SIZES.other },
          defaultTextStyle: DEFAULT_TEXT_HUNDREDTHS,
        },
      },
    ],
  };
}

/**
 * What, if anything, reads `p:otherStyle`.
 *
 * Nothing so far does: not a non-placeholder on a slide, not the
 * header-and-footer trio, not any of the sixteen placeholder types. ECMA calls
 * it "the text style for all other text", so the last candidates are a shape
 * that is not a placeholder on the *layout* and one on the *master* - the two
 * places whose text is drawn on the slide without being on it. If both come back
 * at the built-in size, the element is dead in this build and the finding is a
 * negative one, stated as such.
 */
function otherStyleDeck(): TextDeck {
  const plain = (id: number, name: string, rect: Rect): string =>
    shape({ id, name, rect, paragraphs: [paragraph({})] });
  return {
    deck: 'otherstyle',
    pkg: {
      themes: [{ scheme: SCHEME_ONE, majorLatin: THEME_MAJOR, minorLatin: THEME_MINOR }],
      masters: [
        {
          theme: 0,
          clrMap: IDENTITY_CLR_MAP,
          shapes: [
            shape({ id: 2, name: 'm-title', rect: MASTER_RECT, ph: 'type="title"' }),
            shape({ id: 3, name: 'm-body', rect: MASTER_RECT, ph: 'type="body" idx="1"' }),
            plain(4, 'm-plain', { x: 60, y: 400, w: 300, h: 80 }),
          ],
          txStyles: bucketTxStyles(),
        },
      ],
      layouts: [
        {
          master: 0,
          name: 'Otherstyle',
          shapes: [
            shape({ id: 2, name: 'l-body', rect: PH_RECT, ph: 'type="body" idx="1"' }),
            plain(3, 'l-plain', { x: 400, y: 400, w: 300, h: 80 }),
          ],
        },
      ],
      slides: [
        {
          layout: 0,
          name: 'Otherstyle',
          shapes: [plain(2, 's-plain', PLAIN_RECT)],
        },
      ],
    },
    probes: [
      {
        id: 'otherstyle-slide',
        deck: 'otherstyle',
        kind: 'otherstyle',
        slide: 1,
        shape: 's-plain',
        paragraph: 1,
        question: 'a shape that is not a placeholder, on the slide',
      },
      {
        id: 'otherstyle-layout',
        deck: 'otherstyle',
        kind: 'otherstyle',
        slide: 1,
        shape: 'layout:l-plain',
        paragraph: 1,
        question: 'a shape that is not a placeholder, on the layout',
      },
      {
        id: 'otherstyle-master',
        deck: 'otherstyle',
        kind: 'otherstyle',
        slide: 1,
        shape: 'master:m-plain',
        paragraph: 1,
        question: 'a shape that is not a placeholder, on the master',
      },
    ],
  };
}

/* -------------------------------------------------------------------------- */
/* deck: merge                                                                */
/* -------------------------------------------------------------------------- */

/**
 * Is the cascade per-property, or does the nearest level that says anything win?
 *
 * Five levels, one property each, none of them overlapping. Per-property gives a
 * 24-point bold italic underlined struck-through run; winner-takes-all gives an
 * 18-point run that is only struck through. There is no reading in between.
 */
function mergeDeck(): TextDeck {
  return {
    deck: 'merge',
    pkg: {
      themes: [{ scheme: SCHEME_ONE, majorLatin: THEME_MAJOR, minorLatin: THEME_MINOR }],
      masters: [
        ladderMaster({
          txStyles: bucketTxStyles(),
        }),
      ],
      layouts: [ladderLayout(lstStyle([{ level: 1, inner: defRPr('b="1"') }]))],
      slides: [
        {
          layout: 0,
          name: 'Merge',
          shapes: [
            shape({
              id: 2,
              name: 's-merge',
              ph: 'type="body" idx="1"',
              lstStyle: lstStyle([{ level: 1, inner: defRPr('i="1"') }]),
              paragraphs: [
                paragraph({
                  pPr: pPr({ defRPr: defRPr('u="sng"') }),
                  runs: [run('<a:rPr lang="en-US" strike="sngStrike"/>')],
                }),
              ],
            }),
            shape({
              id: 3,
              name: 's-nearest',
              ph: 'type="body" idx="1"',
              lstStyle: lstStyle([{ level: 1, inner: defRPr('sz="3200"') }]),
              paragraphs: [paragraph({})],
            }),
          ],
        },
      ],
    },
    probes: [
      {
        id: 'merge-per-property',
        deck: 'merge',
        kind: 'merge',
        slide: 1,
        shape: 's-merge',
        paragraph: 1,
        question: 'five levels, one property each: do they merge or does the nearest one win',
        expectPt: BUCKET_PT.body,
        expect: { bold: true, italic: true, underline: true, strike: true },
      },
      {
        id: 'merge-nearest-wins',
        deck: 'merge',
        kind: 'merge',
        slide: 1,
        shape: 's-nearest',
        paragraph: 1,
        question: 'the same property at two levels: the nearer must win',
        expectPt: 32,
      },
    ],
  };
}

/* -------------------------------------------------------------------------- */
/* deck: marL and indent                                                      */
/* -------------------------------------------------------------------------- */

/**
 * The two attributes whose schema defaults must not be applied before inheriting.
 *
 * `@marL` defaults to 347663 and `@indent` to -342900 in ECMA-376. A parser that
 * materialises those at parse time gives every paragraph a 27-point hanging
 * indent it never asked for, and - worse - blocks the inherited value, because a
 * default that has been written down is indistinguishable from a declaration.
 * The first probe reads what PowerPoint reports when nothing anywhere declares
 * either; the third asks whether the two attributes inherit independently.
 */
const MAR_L = 742950; // 58.5pt
const INDENT = -285750; // -22.5pt

function marginDeck(): TextDeck {
  return {
    deck: 'margin',
    pkg: {
      themes: [{ scheme: SCHEME_ONE, majorLatin: THEME_MAJOR, minorLatin: THEME_MINOR }],
      masters: [
        ladderMaster({
          txStyles: txStyles({
            title: lvlList([{ level: 1 }]),
            body: lvlList([{ level: 1, attrs: `marL="${String(MAR_L)}"` }]),
            other: lvlList([{ level: 1 }]),
          }),
        }),
      ],
      layouts: [
        {
          master: 0,
          name: 'Margins',
          shapes: [
            shape({
              id: 2,
              name: 'l-body',
              rect: PH_RECT,
              ph: 'type="body" idx="1"',
              lstStyle: lstStyle([{ level: 1, attrs: `indent="${String(INDENT)}"` }]),
            }),
            shape({ id: 3, name: 'l-bare', rect: PLAIN_RECT, ph: 'type="body" idx="2"' }),
          ],
        },
      ],
      slides: [
        {
          layout: 0,
          name: 'Margins',
          shapes: [
            shape({
              id: 2,
              name: 's-split',
              ph: 'type="body" idx="1"',
              paragraphs: [paragraph({})],
            }),
            shape({
              id: 3,
              name: 's-bare',
              ph: 'type="body" idx="2"',
              paragraphs: [paragraph({})],
            }),
            shape({
              id: 4,
              name: 's-plain',
              rect: { x: 500, y: 260, w: 400, h: 120 },
              paragraphs: [paragraph({})],
            }),
          ],
        },
      ],
    },
    probes: [
      {
        id: 'margin-split',
        deck: 'margin',
        kind: 'margin',
        slide: 1,
        shape: 's-split',
        paragraph: 1,
        question: 'marL from p:bodyStyle and indent from the layout placeholder: do they merge',
        expect: { leftIndentPt: MAR_L / 12700, firstLineIndentPt: INDENT / 12700 },
      },
      {
        id: 'margin-bare',
        deck: 'margin',
        kind: 'margin',
        slide: 1,
        shape: 's-bare',
        paragraph: 1,
        question: 'a body placeholder reaching p:bodyStyle, which declares marL but no indent',
        expect: { leftIndentPt: MAR_L / 12700 },
      },
      {
        id: 'margin-plain',
        deck: 'margin',
        kind: 'margin',
        slide: 1,
        shape: 's-plain',
        paragraph: 1,
        question: 'a shape that is not a placeholder, where nothing declares marL or indent',
      },
    ],
  };
}

/* -------------------------------------------------------------------------- */
/* deck: the two spellings of ST_Percentage                                   */
/* -------------------------------------------------------------------------- */

/**
 * `val="150000"` and `val="150%"` are the same number.
 *
 * 2.6 measured the second spelling accepted inside a Transitional part for a
 * colour transform. Line spacing is the place it matters most: `parseInt("150%")`
 * is 150, which is 0.15% line spacing, and the slide collapses to a line.
 */
function percentDeck(): TextDeck {
  const spaced = (lnSpc: string, spcBef: string): string =>
    paragraph({ pPr: pPr({ lnSpc, spcBef }), runs: [run('<a:rPr lang="en-US" sz="2400"/>')] });

  const numeric = spaced(
    '<a:lnSpc><a:spcPct val="150000"/></a:lnSpc>',
    '<a:spcBef><a:spcPct val="50000"/></a:spcBef>',
  );
  const strict = spaced(
    '<a:lnSpc><a:spcPct val="150%"/></a:lnSpc>',
    '<a:spcBef><a:spcPct val="50%"/></a:spcBef>',
  );
  const points = spaced(
    '<a:lnSpc><a:spcPts val="3000"/></a:lnSpc>',
    '<a:spcBef><a:spcPts val="1200"/></a:spcBef>',
  );

  return {
    deck: 'percent',
    pkg: {
      themes: [{ scheme: SCHEME_ONE, majorLatin: THEME_MAJOR, minorLatin: THEME_MINOR }],
      masters: [ladderMaster({ txStyles: bucketTxStyles() })],
      layouts: [ladderLayout(undefined)],
      slides: [
        {
          layout: 0,
          name: 'Percent',
          shapes: [
            shape({
              id: 2,
              name: 's-pct',
              rect: { x: 60, y: 60, w: 400, h: 300 },
              paragraphs: [numeric, strict, points],
            }),
          ],
        },
      ],
    },
    probes: [
      {
        id: 'percent-numeric',
        deck: 'percent',
        kind: 'percent',
        slide: 1,
        shape: 's-pct',
        paragraph: 1,
        question: 'a:spcPct val="150000" and a:spcBef val="50000"',
        expect: { spaceWithin: 1.5, lineRuleWithin: true, spaceBefore: 0.5, lineRuleBefore: true },
      },
      {
        id: 'percent-strict',
        deck: 'percent',
        kind: 'percent',
        slide: 1,
        shape: 's-pct',
        paragraph: 2,
        question: 'the Strict spelling, val="150%" and val="50%", inside a Transitional part',
        expect: { spaceWithin: 1.5, lineRuleWithin: true, spaceBefore: 0.5, lineRuleBefore: true },
      },
      {
        id: 'percent-points',
        deck: 'percent',
        kind: 'percent',
        slide: 1,
        shape: 's-pct',
        paragraph: 3,
        question: 'a:spcPts val="3000", which is hundredths of a point and not a percentage',
        expect: { spaceWithin: 30, lineRuleWithin: false, spaceBefore: 12, lineRuleBefore: false },
      },
    ],
  };
}

/* -------------------------------------------------------------------------- */
/* deck: the theme font scheme                                                */
/* -------------------------------------------------------------------------- */

/**
 * `+mj-lt` and `+mn-lt`, and what a run gets when nothing names a typeface.
 *
 * The last hop of the cascade is a theme indirection rather than a value, and it
 * is the only hop that can still be unresolved after every level has spoken: a
 * package whose master states no `a:latin` anywhere leaves PowerPoint to pick,
 * and what it picks is a measurement rather than a deduction.
 */
function fontDeck(): TextDeck {
  const latin = (face: string): string => `<a:latin typeface="${face}"/>`;
  return {
    deck: 'font',
    pkg: {
      themes: [{ scheme: SCHEME_ONE, majorLatin: THEME_MAJOR, minorLatin: THEME_MINOR }],
      masters: [
        ladderMaster({
          txStyles: txStyles({
            title: lvlList([{ level: 1, inner: defRPr('sz="4000"', latin('+mj-lt')) }]),
            body: lvlList([{ level: 1, inner: defRPr('sz="2400"', latin('+mn-lt')) }]),
            other: lvlList([{ level: 1, inner: defRPr('sz="1800"') }]),
          }),
        }),
      ],
      layouts: [
        {
          master: 0,
          name: 'Fonts',
          shapes: [
            shape({ id: 2, name: 'l-title', rect: PH_RECT, ph: 'type="title" idx="0"' }),
            shape({ id: 3, name: 'l-body', rect: PLAIN_RECT, ph: 'type="body" idx="1"' }),
          ],
        },
      ],
      slides: [
        {
          layout: 0,
          name: 'Fonts',
          shapes: [
            shape({
              id: 2,
              name: 's-title',
              ph: 'type="title" idx="0"',
              paragraphs: [paragraph({})],
            }),
            shape({
              id: 3,
              name: 's-body',
              ph: 'type="body" idx="1"',
              paragraphs: [paragraph({})],
            }),
            shape({
              id: 4,
              name: 's-explicit',
              rect: { x: 500, y: 60, w: 400, h: 120 },
              paragraphs: [
                paragraph({ runs: [run(`<a:rPr lang="en-US">${latin('Courier New')}</a:rPr>`)] }),
              ],
            }),
            shape({
              id: 5,
              name: 's-silent',
              rect: { x: 500, y: 260, w: 400, h: 120 },
              paragraphs: [paragraph({})],
            }),
          ],
        },
      ],
    },
    probes: [
      {
        id: 'font-major',
        deck: 'font',
        kind: 'font',
        slide: 1,
        shape: 's-title',
        paragraph: 1,
        question: 'p:titleStyle declares +mj-lt',
        expectFont: THEME_MAJOR,
        expectPt: 40,
      },
      {
        id: 'font-minor',
        deck: 'font',
        kind: 'font',
        slide: 1,
        shape: 's-body',
        paragraph: 1,
        question: 'p:bodyStyle declares +mn-lt',
        expectFont: THEME_MINOR,
        expectPt: 24,
      },
      {
        id: 'font-explicit',
        deck: 'font',
        kind: 'font',
        slide: 1,
        shape: 's-explicit',
        paragraph: 1,
        question: 'a run naming a typeface outright',
        expectFont: 'Courier New',
      },
      {
        id: 'font-silent',
        deck: 'font',
        kind: 'font',
        slide: 1,
        shape: 's-silent',
        paragraph: 1,
        question: 'a shape reaching p:otherStyle, which names no typeface at all',
      },
    ],
  };
}

/* -------------------------------------------------------------------------- */
/* deck: which placeholder the text comes from                                */
/* -------------------------------------------------------------------------- */

/**
 * Does text inherit along the same chain geometry does?
 *
 * 2.9 measured that chain on position: slide to layout on `@idx` alone, layout
 * to master on the folded type taking the first. If text walks a different one -
 * and every implementation this project has read matches text on `(type, idx)` -
 * then a slide `title` at idx 1 whose layout holds a `title` at idx 5 and a
 * `body` at idx 1 takes its size from the body, and the number says which.
 */
function hopDeck(): TextDeck {
  return {
    deck: 'hop',
    pkg: {
      themes: [{ scheme: SCHEME_ONE, majorLatin: THEME_MAJOR, minorLatin: THEME_MINOR }],
      masters: [
        {
          theme: 0,
          clrMap: IDENTITY_CLR_MAP,
          shapes: [
            shape({
              id: 2,
              name: 'm-title',
              rect: MASTER_RECT,
              ph: 'type="title"',
              lstStyle: sizeAt(1, 52),
            }),
            shape({
              id: 3,
              name: 'm-body',
              rect: MASTER_RECT,
              ph: 'type="body" idx="1"',
              lstStyle: sizeAt(1, 26),
            }),
          ],
          txStyles: bucketTxStyles(),
        },
      ],
      layouts: [
        {
          master: 0,
          name: 'Hops',
          shapes: [
            shape({
              id: 2,
              name: 'l-body1',
              rect: PH_RECT,
              ph: 'type="body" idx="1"',
              lstStyle: sizeAt(1, 32),
            }),
            shape({
              id: 3,
              name: 'l-title5',
              rect: PLAIN_RECT,
              ph: 'type="title" idx="5"',
              lstStyle: sizeAt(1, 48),
            }),
            shape({ id: 4, name: 'l-ctr2', rect: PH_RECT, ph: 'type="ctrTitle" idx="2"' }),
          ],
        },
      ],
      slides: [
        {
          layout: 0,
          name: 'Hops',
          shapes: [
            shape({
              id: 2,
              name: 's-title1',
              ph: 'type="title" idx="1"',
              paragraphs: [paragraph({})],
            }),
            shape({
              id: 3,
              name: 's-ctr2',
              ph: 'type="ctrTitle" idx="2"',
              paragraphs: [paragraph({})],
            }),
            shape({
              id: 4,
              name: 's-orphan9',
              rect: { x: 500, y: 400, w: 400, h: 100 },
              ph: 'type="body" idx="9"',
              paragraphs: [paragraph({})],
            }),
          ],
        },
      ],
    },
    probes: [
      {
        id: 'hop-idx-only',
        deck: 'hop',
        kind: 'hop',
        slide: 1,
        shape: 's-title1',
        paragraph: 1,
        question:
          'a slide title at idx 1: the layout body at idx 1 (32) or the title at idx 5 (48)',
        expectPt: 32,
      },
      {
        id: 'hop-type-fold',
        deck: 'hop',
        kind: 'hop',
        slide: 1,
        shape: 's-ctr2',
        paragraph: 1,
        question:
          'a layout ctrTitle declaring nothing: the master title (52) or the master body (26)',
        expectPt: 52,
      },
      {
        id: 'hop-orphan',
        deck: 'hop',
        kind: 'hop',
        slide: 1,
        shape: 's-orphan9',
        paragraph: 1,
        question: 'a slide body at idx 9 that matches nothing: which bucket does it still reach',
        expectPt: BUCKET_PT.body,
      },
    ],
  };
}

/* -------------------------------------------------------------------------- */
/* the whole experiment                                                       */
/* -------------------------------------------------------------------------- */

export function textDecks(): readonly TextDeck[] {
  const ladders: TextDeck[] = [];
  for (let from = 0; from <= SOURCES.length; from++) ladders.push(ladderDeck(from));
  return [
    ...ladders,
    ...bucketDecks(),
    crossedBucketDeck(),
    levelDeck(),
    levelCrossDeck(),
    levelShadowDeck(),
    defPPrDeck(),
    builtinDeck(),
    hfDefaultDeck(),
    partialBucketDeck(),
    otherStyleDeck(),
    mergeDeck(),
    marginDeck(),
    percentDeck(),
    fontDeck(),
    hopDeck(),
  ];
}
