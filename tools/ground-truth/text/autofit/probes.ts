/**
 * Experiment T4, step 1 - what to ask PowerPoint about autofit.
 *
 * `a:normAutofit` carries two attributes, `@fontScale` and `@lnSpcReduction`,
 * and every question worth asking about them is a question about *where those
 * numbers come from*. The plan says a "discrete 14-step ladder". That is a
 * claim, not a measurement, and this file is built so that PowerPoint either
 * confirms it in its own handwriting or refutes it.
 *
 * ## The trick that makes this measurable
 *
 * The obvious probe - put prose in a box, shrink the box, see what scale comes
 * out - cannot be scored without a line-breaking model, because the text height
 * depends on how many lines the prose wrapped into, which depends on the scale,
 * which is the unknown. Any error in our breaker would land on the ladder and
 * be indistinguishable from it.
 *
 * So the primary sweep has **one short word per paragraph**. Nothing can wrap:
 * "one" at 18pt is 28pt wide in a 300pt box, and stays narrower than the box at
 * every rung because the rungs only ever shrink. The line count is therefore
 * the paragraph count, *exactly*, independent of the scale - and PowerPoint's
 * own `Lines().Count` is recorded so that assumption is checked rather than
 * trusted. The fit test collapses to arithmetic over numbers the file states:
 *
 * ```
 * lines x 1.2 x effectiveSize(sz, fontScale) x (1 - lnSpcReduction)  <=  height
 * ```
 *
 * with `effectiveSize` the rounding rule under test. No font is consulted, no
 * break opportunity matters, and a wrong breaker cannot contaminate the answer.
 *
 * ## Reading the rungs off the box, not off the text
 *
 * The second trick is which variable to sweep. Sweeping the *content* moves the
 * required scale in steps of one whole paragraph - far too coarse to tell 92.5%
 * from 92%. Sweeping the *box height* by 1pt with the content held still moves
 * it in steps of `1 / (lines x 1.2 x sz)`, which at 20 lines of 18pt is 0.23%
 * of scale. Every height at which the emitted scale steps up is a bracket on
 * that rung's effective size, tight enough to name a quarter of a percent.
 *
 * So `ladder`, `ladder-b` and `ladder-c` hold the text still and sweep the box.
 * They differ in line count, font size and typeface precisely so that a rung
 * that is really a property of the ladder can be told apart from one that is an
 * artefact of a particular box.
 *
 * ## What the other decks are for
 *
 * | deck        | asks                                                           |
 * |-------------|----------------------------------------------------------------|
 * | `ladder*`   | the rungs, the fit predicate, and the rounding of `pt x scale`  |
 * | `stored`    | whether a stored scale is applied verbatim or recomputed        |
 * | `spacing`   | `spcBef`/`spcAft`, and whether `spcFirstLastPara` defaults false |
 * | `lnspc`     | how `lnSpcReduction` composes with `a:lnSpc` percent and points  |
 * | `inset`     | whether the fit test is against the inset-reduced height         |
 * | `wrap`      | that the same rule holds once real wrapping is involved          |
 * | `sp`        | `a:spAutoFit` - the shape grows instead of the text shrinking     |
 *
 * `stored` is the falsifier for the plan's whole view-mode design. It states
 * scales that no recomputation would ever produce for its content - 25% on text
 * that fits at 100% - and is opened without being touched. If PowerPoint
 * renders at 25%, stored values are authoritative and a viewer must apply them
 * verbatim. If it renders at 100%, the plan is wrong and view mode has to
 * recompute like everything else.
 */

/** A rectangle in points. Same shape as `sheet-pptx.ts`'s, restated so the
 *  probe table does not depend on the package builder. */
export interface ProbeRect {
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
}

/** `a:bodyPr` insets in points. PowerPoint's defaults are asymmetric. */
export interface Insets {
  readonly l: number;
  readonly t: number;
  readonly r: number;
  readonly b: number;
}

/** Zero on every side, which is what the ladder sweeps use so that the box
 *  height and the available height are the same number. */
export const NO_INSETS: Insets = { l: 0, t: 0, r: 0, b: 0 };

/** The schema defaults, in points: 0.1in left and right, 0.05in top and bottom. */
export const DEFAULT_INSETS: Insets = { l: 7.2, t: 3.6, r: 7.2, b: 3.6 };

/** `a:lnSpc`, `a:spcBef` or `a:spcAft` as the file states it. */
export interface Spacing {
  readonly kind: 'percent' | 'points';
  /** Thousandths of a percent, or hundredths of a point. */
  readonly value: number;
}

/** A stored `a:normAutofit`, in the units the attributes use. */
export interface StoredScale {
  /** Thousandths of a percent. `undefined` writes no attribute. */
  readonly fontScale?: number | undefined;
  /** Thousandths of a percent. `undefined` writes no attribute. */
  readonly lnSpcReduction?: number | undefined;
}

export type AutofitKind = 'norm' | 'sp' | 'none';

export interface Probe {
  readonly id: string;
  readonly deck: string;
  readonly rect: ProbeRect;
  readonly face: string;
  /** Nominal size in hundredths of a point, as `a:rPr/@sz` states it. */
  readonly sz: number;
  /** One string per paragraph. */
  readonly paragraphs: readonly string[];
  readonly insets: Insets;
  readonly autofit: AutofitKind;
  /** What `a:normAutofit` states before PowerPoint is asked. */
  readonly stored: StoredScale;
  /**
   * Whether the runner forces a fresh computation.
   *
   * False means the deck is opened and read without being touched, which is the
   * only way to ask what a *stored* value does.
   */
  readonly recompute: boolean;
  readonly lnSpc?: Spacing | undefined;
  readonly spcBef?: Spacing | undefined;
  readonly spcAft?: Spacing | undefined;
  /** `a:bodyPr/@spcFirstLastPara`. Absent writes no attribute. */
  readonly spcFirstLastPara?: boolean | undefined;
  /** Expected line count, where the probe is built so that it is knowable. */
  readonly expectLines?: number | undefined;
  /** Free-form note carried into the fixture so a row explains itself. */
  readonly note: string;
}

/* -------------------------------------------------------------------------- */
/* the words                                                                  */
/* -------------------------------------------------------------------------- */

/**
 * Short words, one per paragraph, that cannot wrap in any box these probes use.
 *
 * Three and four letters at 18pt Arial is at most ~38pt against a 300pt box, and
 * autofit only ever shrinks, so the margin only widens as the sweep goes on.
 * They are cycled rather than repeated so that a mis-parsed paragraph shows up
 * as a wrong word rather than as an invisible off-by-one.
 */
const SHORT_WORDS: readonly string[] = [
  'one',
  'two',
  'six',
  'ten',
  'red',
  'sky',
  'cup',
  'oak',
  'ice',
  'pen',
  'bay',
  'fog',
];

export function shortParagraphs(count: number): readonly string[] {
  const out: string[] = [];
  for (let i = 0; i < count; i += 1) {
    const word = SHORT_WORDS[i % SHORT_WORDS.length];
    if (word === undefined) throw new Error('unreachable: SHORT_WORDS is non-empty');
    out.push(word);
  }
  return out;
}

/**
 * Prose for the `wrap` deck.
 *
 * Lower-case ASCII words separated by single spaces, with no punctuation, no
 * hyphen and no character that T3 measured as a break opportunity other than
 * the space. That keeps the wrap deck a test of the *fit predicate* under
 * wrapping rather than a second line-breaking experiment.
 */
const PROSE_WORDS: readonly string[] = [
  'the',
  'quick',
  'brown',
  'fox',
  'jumps',
  'over',
  'a',
  'lazy',
  'dog',
  'while',
  'seven',
  'ships',
  'sail',
  'past',
  'the',
  'harbour',
  'wall',
  'at',
  'dawn',
  'and',
  'the',
  'keeper',
  'counts',
  'them',
  'from',
  'his',
  'tower',
  'window',
  'in',
  'the',
  'cold',
  'grey',
  'light',
];

export function prose(wordCount: number): string {
  const out: string[] = [];
  for (let i = 0; i < wordCount; i += 1) {
    const word = PROSE_WORDS[i % PROSE_WORDS.length];
    if (word === undefined) throw new Error('unreachable: PROSE_WORDS is non-empty');
    out.push(word);
  }
  return out.join(' ');
}

/* -------------------------------------------------------------------------- */
/* the sweeps                                                                 */
/* -------------------------------------------------------------------------- */

const BASE: Omit<Probe, 'id' | 'deck' | 'rect' | 'paragraphs' | 'note'> = {
  face: 'Arial',
  sz: 1800,
  insets: NO_INSETS,
  autofit: 'norm',
  stored: {},
  recompute: true,
};

/**
 * One ladder sweep: the text held still, the box height swept by `step`.
 *
 * The height range is chosen from the arithmetic rather than by eye. At the top
 * of the range the content fits at 100% and the answer must be "no scale"; at
 * the bottom it is past whatever the ladder's floor turns out to be, so the
 * floor is measured rather than assumed to exist. Everything in between is one
 * observation of "the largest rung that fits in this many points".
 */
function ladderSweep(
  deck: string,
  opts: {
    readonly lines: number;
    readonly sz: number;
    readonly face: string;
    readonly widthPt: number;
    readonly fromPt: number;
    readonly toPt: number;
    readonly stepPt: number;
    readonly note: string;
  },
): readonly Probe[] {
  const paragraphs = shortParagraphs(opts.lines);
  const out: Probe[] = [];
  let index = 0;
  for (let h = opts.fromPt; h <= opts.toPt + 1e-9; h += opts.stepPt) {
    // Quarter-point quantisation, as T3 used: dyadic, so the height is exact in
    // binary *and* a whole number of EMU (12700 = 4 x 3175).
    const height = Math.round(h * 4) / 4;
    out.push({
      ...BASE,
      id: `${deck}-h${String(index).padStart(3, '0')}`,
      deck,
      face: opts.face,
      sz: opts.sz,
      rect: { x: 0, y: 0, w: opts.widthPt, h: height },
      paragraphs,
      expectLines: opts.lines,
      note: opts.note,
    });
    index += 1;
  }
  return out;
}

/**
 * The three ladder decks.
 *
 * `ladder` is the primary: 20 lines of 18pt Arial in a box swept from far below
 * the ladder's plausible floor to comfortably above 100%. 20 lines is chosen
 * because the height resolution of a rung is `step / (lines x 1.2 x sz)` - at
 * 1pt steps that is 0.23% of scale, tight enough to name a rung to the quarter
 * percent.
 *
 * `ladder-b` and `ladder-c` change every free variable at once - line count,
 * font size, typeface - so that a rung common to all three is a property of the
 * ladder and not of a box. `ladder-c`'s 11pt is deliberately small: if the
 * effective size is quantised at all, an 11pt nominal is where the quantum is
 * the largest fraction of the size and the alternatives separate the widest.
 */
export function ladderProbes(): readonly Probe[] {
  return [
    ...ladderSweep('ladder', {
      lines: 20,
      sz: 1800,
      face: 'Arial',
      widthPt: 300,
      fromPt: 20,
      toPt: 448,
      stepPt: 1,
      note: '20 single-word paragraphs, 18pt Arial, box height swept 1pt at a time',
    }),
    ...ladderSweep('ladder-b', {
      lines: 7,
      sz: 3200,
      face: 'Arial',
      widthPt: 400,
      fromPt: 24,
      toPt: 288,
      stepPt: 1,
      note: '7 paragraphs at 32pt: a different line count and size against the same rungs',
    }),
    ...ladderSweep('ladder-c', {
      lines: 13,
      sz: 1100,
      face: 'Times New Roman',
      widthPt: 260,
      fromPt: 14,
      toPt: 186,
      stepPt: 1,
      note: '13 paragraphs at 11pt Times New Roman: small size, where a quantised effective size separates the models widest',
    }),
  ];
}

/* -------------------------------------------------------------------------- */
/* stored scales, applied without being touched                               */
/* -------------------------------------------------------------------------- */

/**
 * Scales chosen so that the candidate rounding rules disagree.
 *
 * At 18pt, `92500` is 16.65pt exactly. Rounded to a whole point that is 17,
 * to a half 16.5, to a quarter 16.75 - three different line advances, 0.3pt
 * apart, against a measurement good to about 0.01pt. Every entry here is picked
 * the same way: the products are deliberately awkward.
 */
const STORED_CASES: readonly { readonly sz: number; readonly scale: number }[] = [
  { sz: 1800, scale: 92500 },
  { sz: 1800, scale: 70000 },
  { sz: 1800, scale: 62500 },
  { sz: 1800, scale: 45000 },
  { sz: 1800, scale: 25000 },
  { sz: 1100, scale: 92500 },
  { sz: 1100, scale: 77500 },
  { sz: 1100, scale: 62500 },
  { sz: 1100, scale: 40000 },
  { sz: 2800, scale: 92500 },
  { sz: 2800, scale: 65000 },
  { sz: 2800, scale: 37500 },
  { sz: 1050, scale: 90000 },
  { sz: 1050, scale: 55000 },
  { sz: 4400, scale: 92500 },
  { sz: 4400, scale: 32500 },
];

/** Reductions swept against one size, to see how the two attributes compose. */
const REDUCTION_CASES: readonly number[] = [0, 10000, 20000, 25000, 15000];

export function storedProbes(): readonly Probe[] {
  const out: Probe[] = [];
  let index = 0;

  // Font scale alone. The box is far larger than the content needs, so a
  // recomputation - if one happened - would answer 100% and be unmistakable.
  for (const c of STORED_CASES) {
    out.push({
      ...BASE,
      id: `stored-s${String(index).padStart(3, '0')}`,
      deck: 'stored',
      sz: c.sz,
      rect: { x: 0, y: 0, w: 300, h: 380 },
      paragraphs: shortParagraphs(6),
      expectLines: 6,
      stored: { fontScale: c.scale },
      recompute: false,
      note: `stored fontScale ${String(c.scale)} at ${String(c.sz / 100)}pt in a box that fits the text at 100%`,
    });
    index += 1;
  }

  // Reduction alone, and reduction with a scale, at one size.
  for (const red of REDUCTION_CASES) {
    out.push({
      ...BASE,
      id: `stored-r${String(index).padStart(3, '0')}`,
      deck: 'stored',
      rect: { x: 0, y: 0, w: 300, h: 380 },
      paragraphs: shortParagraphs(6),
      expectLines: 6,
      stored: { lnSpcReduction: red },
      recompute: false,
      note: `stored lnSpcReduction ${String(red)} with no fontScale`,
    });
    index += 1;
    out.push({
      ...BASE,
      id: `stored-r${String(index).padStart(3, '0')}`,
      deck: 'stored',
      rect: { x: 0, y: 0, w: 300, h: 380 },
      paragraphs: shortParagraphs(6),
      expectLines: 6,
      stored: { fontScale: 62500, lnSpcReduction: red },
      recompute: false,
      note: `stored fontScale 62500 with lnSpcReduction ${String(red)}`,
    });
    index += 1;
  }

  // A scale no ladder would ever choose, on text that fits at 100%. This is the
  // falsifier for view mode: if it renders at 25% the stored value is
  // authoritative, and if it renders at 100% the plan's view mode is wrong.
  out.push({
    ...BASE,
    id: `stored-x${String(index).padStart(3, '0')}`,
    deck: 'stored',
    rect: { x: 0, y: 0, w: 300, h: 380 },
    paragraphs: shortParagraphs(2),
    expectLines: 2,
    stored: { fontScale: 25000, lnSpcReduction: 20000 },
    recompute: false,
    note: 'two words in a huge box at a stored 25% - no recomputation would ever produce this',
  });
  index += 1;

  // The mirror image: a stored 100% on content that overflows badly. A viewer
  // that recomputes would shrink it; one that applies the file verbatim leaves
  // it overflowing, which is what PowerPoint shows for a file another producer
  // wrote without running autofit.
  out.push({
    ...BASE,
    id: `stored-x${String(index).padStart(3, '0')}`,
    deck: 'stored',
    rect: { x: 0, y: 0, w: 300, h: 60 },
    paragraphs: shortParagraphs(12),
    expectLines: 12,
    stored: { fontScale: 100000 },
    recompute: false,
    note: 'twelve lines stated at 100% in a box that holds two - does opening the file shrink them',
  });

  return out;
}

/* -------------------------------------------------------------------------- */
/* paragraph spacing                                                          */
/* -------------------------------------------------------------------------- */

/**
 * `spcBef`, `spcAft`, and the flag that decides whether the outer two count.
 *
 * The plan says `spcFirstLastPara` defaults false, meaning the first
 * paragraph's `spcBef` and the last's `spcAft` are discarded. That is a claim
 * about two numbers, and it is asked here three ways at once: absent, `0` and
 * `1`, against the same content, so "absent behaves as 0" is a measurement
 * rather than a reading of the schema.
 *
 * These are recompute probes. The emitted scale is a *function of the text
 * height*, so a spacing term that is or is not included moves the rung - which
 * makes the emitted `@fontScale` a far more sensitive read-out of the height
 * model than any position measurement could be.
 */
export function spacingProbes(): readonly Probe[] {
  const out: Probe[] = [];
  let index = 0;
  const flags: readonly (boolean | undefined)[] = [undefined, false, true];
  const spacings: readonly { readonly bef?: Spacing; readonly aft?: Spacing }[] = [
    { bef: { kind: 'points', value: 1200 } },
    { aft: { kind: 'points', value: 1200 } },
    { bef: { kind: 'points', value: 1200 }, aft: { kind: 'points', value: 600 } },
    { bef: { kind: 'percent', value: 50000 } },
    { aft: { kind: 'percent', value: 50000 } },
    { bef: { kind: 'percent', value: 100000 }, aft: { kind: 'percent', value: 50000 } },
  ];

  for (const flag of flags) {
    for (const sp of spacings) {
      for (const height of [120, 168]) {
        out.push({
          ...BASE,
          id: `spacing-${String(index).padStart(3, '0')}`,
          deck: 'spacing',
          rect: { x: 0, y: 0, w: 300, h: height },
          paragraphs: shortParagraphs(5),
          expectLines: 5,
          ...(sp.bef === undefined ? {} : { spcBef: sp.bef }),
          ...(sp.aft === undefined ? {} : { spcAft: sp.aft }),
          ...(flag === undefined ? {} : { spcFirstLastPara: flag }),
          note:
            `5 paragraphs, spcFirstLastPara ${flag === undefined ? 'absent' : String(flag)}` +
            `, bef ${sp.bef === undefined ? 'none' : `${sp.bef.kind}:${String(sp.bef.value)}`}` +
            `, aft ${sp.aft === undefined ? 'none' : `${sp.aft.kind}:${String(sp.aft.value)}`}`,
        });
        index += 1;
      }
    }
  }

  // A no-autofit control for each spacing, read for its line positions rather
  // than for a scale. Without these, "the height model" is only ever observed
  // through the ladder and a spacing rule that is wrong by a constant would be
  // absorbed into a rung.
  for (const sp of spacings) {
    for (const flag of flags) {
      out.push({
        ...BASE,
        id: `spacing-ctl${String(index).padStart(3, '0')}`,
        deck: 'spacing-ctl',
        rect: { x: 0, y: 0, w: 300, h: 400 },
        paragraphs: shortParagraphs(4),
        expectLines: 4,
        autofit: 'none',
        recompute: false,
        ...(sp.bef === undefined ? {} : { spcBef: sp.bef }),
        ...(sp.aft === undefined ? {} : { spcAft: sp.aft }),
        ...(flag === undefined ? {} : { spcFirstLastPara: flag }),
        note: `control: 4 paragraphs, no autofit, line positions read directly`,
      });
      index += 1;
    }
  }

  return out;
}

/* -------------------------------------------------------------------------- */
/* line spacing against the reduction                                         */
/* -------------------------------------------------------------------------- */

/**
 * How `@lnSpcReduction` composes with `a:lnSpc`.
 *
 * T2 measured that percent line spacing multiplies the 1.2 line box and exact
 * points replaces it. The reduction is a *second* multiplier, and there are
 * three plausible places to apply it: to the 1.2 box before the percentage, to
 * the product after it, or - the reading that would surprise nobody who has
 * only read ECMA - not to exact points at all, since "reduce the line spacing"
 * has no obvious meaning when the spacing is stated as a length.
 *
 * Both halves are asked: recompute probes, where the emitted rung reports which
 * composition PowerPoint used to decide the fit, and stored probes, where the
 * line positions report which one it used to lay out.
 */
export function lnSpcProbes(): readonly Probe[] {
  const out: Probe[] = [];
  let index = 0;
  const spacings: readonly (Spacing | undefined)[] = [
    undefined,
    { kind: 'percent', value: 150000 },
    { kind: 'percent', value: 80000 },
    { kind: 'points', value: 3000 },
    { kind: 'points', value: 1500 },
  ];

  for (const ln of spacings) {
    for (const red of [0, 10000, 20000]) {
      out.push({
        ...BASE,
        id: `lnspc-${String(index).padStart(3, '0')}`,
        deck: 'lnspc-stored',
        rect: { x: 0, y: 0, w: 300, h: 380 },
        paragraphs: shortParagraphs(6),
        expectLines: 6,
        ...(ln === undefined ? {} : { lnSpc: ln }),
        stored: { fontScale: 62500, lnSpcReduction: red },
        recompute: false,
        note:
          `stored 62.5% with reduction ${String(red)}, lnSpc ` +
          `${ln === undefined ? 'absent' : `${ln.kind}:${String(ln.value)}`}`,
      });
      index += 1;
    }
    for (const height of [90, 140, 210]) {
      out.push({
        ...BASE,
        id: `lnspc-r${String(index).padStart(3, '0')}`,
        deck: 'lnspc-fit',
        rect: { x: 0, y: 0, w: 300, h: height },
        paragraphs: shortParagraphs(8),
        expectLines: 8,
        ...(ln === undefined ? {} : { lnSpc: ln }),
        note: `recompute, 8 paragraphs in ${String(height)}pt, lnSpc ${ln === undefined ? 'absent' : `${ln.kind}:${String(ln.value)}`}`,
      });
      index += 1;
    }
  }
  return out;
}

/* -------------------------------------------------------------------------- */
/* insets                                                                     */
/* -------------------------------------------------------------------------- */

/**
 * Is the fit test against the box, or against the box minus the insets?
 *
 * Zero insets everywhere else means the question never arises there, which is
 * deliberate: it keeps the ladder sweep clean. Here the insets are made large
 * enough - up to 30pt a side - that the two readings choose different rungs
 * over most of the range rather than only at a boundary.
 */
export function insetProbes(): readonly Probe[] {
  const out: Probe[] = [];
  let index = 0;
  const insets: readonly Insets[] = [
    NO_INSETS,
    DEFAULT_INSETS,
    { l: 7.2, t: 20, r: 7.2, b: 20 },
    { l: 7.2, t: 30, r: 7.2, b: 6 },
    { l: 40, t: 3.6, r: 40, b: 3.6 },
  ];
  for (const ins of insets) {
    for (const height of [120, 160, 200, 260]) {
      out.push({
        ...BASE,
        id: `inset-${String(index).padStart(3, '0')}`,
        deck: 'inset',
        rect: { x: 0, y: 0, w: 300, h: height },
        paragraphs: shortParagraphs(9),
        expectLines: 9,
        insets: ins,
        note: `9 paragraphs in ${String(height)}pt with insets t${String(ins.t)} b${String(ins.b)}`,
      });
      index += 1;
    }
  }
  return out;
}

/* -------------------------------------------------------------------------- */
/* wrapping                                                                   */
/* -------------------------------------------------------------------------- */

/**
 * The same rule, with real wrapping in the way.
 *
 * These are not used to *derive* anything. The ladder is derived where nothing
 * wraps; this deck is the check that the derived rule still predicts PowerPoint
 * when the line count is a consequence of the scale rather than a constant -
 * which is the case every real deck is in, and the case where a fit predicate
 * that is right for the wrong reason falls apart.
 */
export function wrapProbes(): readonly Probe[] {
  const out: Probe[] = [];
  let index = 0;
  for (const words of [8, 16, 24, 40, 60, 90]) {
    for (const w of [180, 260, 340]) {
      for (const h of [70, 110, 160]) {
        out.push({
          ...BASE,
          id: `wrap-${String(index).padStart(3, '0')}`,
          deck: 'wrap',
          rect: { x: 0, y: 0, w, h },
          paragraphs: [prose(words)],
          note: `${String(words)} words of prose in ${String(w)}x${String(h)}pt`,
        });
        index += 1;
      }
    }
  }
  return out;
}

/* -------------------------------------------------------------------------- */
/* the other autofit                                                          */
/* -------------------------------------------------------------------------- */

/**
 * `a:spAutoFit` - the shape grows to the text instead of the text shrinking.
 *
 * Read through the shape's own height rather than through any scale, because
 * that is where the answer is written. The question a renderer actually has is
 * what height PowerPoint would have written, so that a box whose text we edited
 * ends up the size PowerPoint would have made it - and the height it writes is
 * the text height plus the two insets, or it is not, which is the measurement.
 */
export function spProbes(): readonly Probe[] {
  const out: Probe[] = [];
  let index = 0;
  for (const lines of [1, 2, 3, 5, 8, 13]) {
    for (const sz of [1100, 1800, 3200]) {
      for (const ins of [NO_INSETS, DEFAULT_INSETS, { l: 7.2, t: 20, r: 7.2, b: 12 }]) {
        out.push({
          ...BASE,
          id: `sp-${String(index).padStart(3, '0')}`,
          deck: 'sp',
          rect: { x: 0, y: 0, w: 300, h: 100 },
          sz,
          paragraphs: shortParagraphs(lines),
          expectLines: lines,
          insets: ins,
          autofit: 'sp',
          note: `spAutoFit, ${String(lines)} paragraphs at ${String(sz / 100)}pt, insets t${String(ins.t)} b${String(ins.b)}`,
        });
        index += 1;
      }
    }
  }
  return out;
}

/* -------------------------------------------------------------------------- */

/* -------------------------------------------------------------------------- */
/* round two                                                                  */
/* -------------------------------------------------------------------------- */

/**
 * Everything round one left resting on too little evidence.
 *
 * Round one fitted the ladder, the rounding of `pt x scale`, the paragraph
 * spacing and the inset rule, and every one of those held on hundreds of
 * probes. Four things did not:
 *
 * - **The half-way rounding.** `round(sz x scale)` was confirmed on 28 stored
 *   probes, but only two of them landed exactly on a half - 18pt at 25% and
 *   28pt at 37.5% - and two samples is an anecdote. `ladder-20` puts *five*
 *   rungs on a half at once: 20pt is 18.5, 15.5, 12.5, 9.5 and 6.5 at 92.5,
 *   77.5, 62.5, 47.5 and 32.5 per cent.
 * - **A fractional nominal size.** Nothing asked what `sz="1050"` does at the
 *   rung that reduces the line spacing without scaling the font. If the size
 *   is rounded there too, 10.5pt text loses half a point for no reason a user
 *   could see.
 * - **`a:lnSpc` against the reduction.** One probe of fifteen missed, and it
 *   is the one probe where "lines x advance" and "(n-1) x advance + last line"
 *   disagree. One probe is not a finding either way, so the three `ln-*` decks
 *   sweep the box height under stated line spacing exactly the way `ladder`
 *   swept it under none. `ln-pts` is the sharp one: exact-point spacing ignores
 *   both the font scale and the reduction, so the text height cannot change,
 *   and a ladder that searches at all must produce a two-valued step function.
 * - **`a:spAutoFit`'s constant.** One face, three sizes. If 1.009765625 is real
 *   it is the same on Times New Roman and Courier New; if it is a font metric
 *   in disguise it is not. `sp-hyst` separately asks why two probes of 54 kept
 *   a height PowerPoint had every reason to shrink.
 */

function sweepDeck(
  deck: string,
  opts: {
    readonly lines: number;
    readonly sz: number;
    readonly face?: string;
    readonly widthPt?: number;
    readonly fromPt: number;
    readonly toPt: number;
    readonly stepPt?: number;
    readonly lnSpc?: Spacing;
    readonly spcBef?: Spacing;
    readonly spcAft?: Spacing;
    readonly spcFirstLastPara?: boolean;
    readonly note: string;
  },
): readonly Probe[] {
  const paragraphs = shortParagraphs(opts.lines);
  const step = opts.stepPt ?? 1;
  const out: Probe[] = [];
  let index = 0;
  for (let h = opts.fromPt; h <= opts.toPt + 1e-9; h += step) {
    out.push({
      ...BASE,
      id: `${deck}-h${String(index).padStart(3, '0')}`,
      deck,
      sz: opts.sz,
      ...(opts.face === undefined ? {} : { face: opts.face }),
      rect: { x: 0, y: 0, w: opts.widthPt ?? 300, h: Math.round(h * 4) / 4 },
      paragraphs,
      expectLines: opts.lines,
      ...(opts.lnSpc === undefined ? {} : { lnSpc: opts.lnSpc }),
      ...(opts.spcBef === undefined ? {} : { spcBef: opts.spcBef }),
      ...(opts.spcAft === undefined ? {} : { spcAft: opts.spcAft }),
      ...(opts.spcFirstLastPara === undefined ? {} : { spcFirstLastPara: opts.spcFirstLastPara }),
      note: opts.note,
    });
    index += 1;
  }
  return out;
}

export function roundTwoSweeps(): readonly Probe[] {
  return [
    ...sweepDeck('ladder-20', {
      lines: 9,
      sz: 2000,
      fromPt: 10,
      toPt: 232,
      note: '20pt, where five of the eleven font scales land exactly on a half point',
    }),
    ...sweepDeck('ladder-frac', {
      lines: 11,
      sz: 1050,
      fromPt: 8,
      toPt: 152,
      note: '10.5pt: does a fractional nominal size survive the rung that only reduces line spacing',
    }),
    ...sweepDeck('ln150', {
      lines: 8,
      sz: 1800,
      fromPt: 12,
      toPt: 276,
      lnSpc: { kind: 'percent', value: 150000 },
      note: 'lnSpc 150%: the fit predicate under percentage line spacing',
    }),
    ...sweepDeck('ln080', {
      lines: 8,
      sz: 1800,
      fromPt: 8,
      toPt: 152,
      lnSpc: { kind: 'percent', value: 80000 },
      note: 'lnSpc 80%: the same, on the other side of single spacing',
    }),
    ...sweepDeck('ln-pts', {
      lines: 8,
      sz: 1800,
      fromPt: 20,
      toPt: 210,
      lnSpc: { kind: 'points', value: 2400 },
      note: 'lnSpc 24pt exactly: a text height no rung can change, so the ladder must produce two values and no more',
    }),
    /*
     * The ladder's own completeness check.
     *
     * Under an exact point spacing the reduction changes nothing, so the only
     * thing separating one rung from the next is the effective size - and the
     * *first* rung PowerPoint reaches at each size is therefore the smallest
     * reduction that size has in the ladder. `ln-pts` swept a point at a time
     * and the whole staircase fell inside four points, so it read four of the
     * eleven sizes. A quarter of a point reads all eleven, and with them
     * whether 85% has a zero-reduction rung the way 92.5% turned out to.
     */
    ...sweepDeck('ln-pts-fine', {
      lines: 8,
      sz: 1800,
      fromPt: 185,
      toPt: 192,
      stepPt: 0.25,
      lnSpc: { kind: 'points', value: 2400 },
      note: 'the ln-pts staircase again at a quarter point, where every font scale gets its own step',
    }),
    ...sweepDeck('spc-pts', {
      lines: 5,
      sz: 1800,
      fromPt: 20,
      toPt: 200,
      stepPt: 2,
      spcBef: { kind: 'points', value: 1200 },
      spcAft: { kind: 'points', value: 600 },
      spcFirstLastPara: false,
      note: 'spcBef 12pt / spcAft 6pt, flag off: does exact paragraph spacing scale with the font',
    }),
    ...sweepDeck('spc-pts-on', {
      lines: 5,
      sz: 1800,
      fromPt: 20,
      toPt: 200,
      stepPt: 2,
      spcBef: { kind: 'points', value: 1200 },
      spcAft: { kind: 'points', value: 600 },
      spcFirstLastPara: true,
      note: 'the same with the flag on, so the two outer terms are in play',
    }),
    ...sweepDeck('spc-pct', {
      lines: 5,
      sz: 1800,
      fromPt: 20,
      toPt: 220,
      stepPt: 2,
      spcBef: { kind: 'percent', value: 100000 },
      spcAft: { kind: 'percent', value: 50000 },
      spcFirstLastPara: false,
      note: 'percentage paragraph spacing, which follows the scaled line box and so moves with the rung',
    }),
  ];
}

/**
 * `a:spAutoFit` again, across four faces.
 *
 * The constant round one measured scales the *insets* as well as the text, so
 * it cannot be a font metric - but that is an argument, and the cheap way to
 * settle it is to ask three more typefaces whose vertical metrics disagree
 * violently. Courier New and Verdana are in the set for the same reason T2 used
 * them: if any font-derived term were in this number, they would report a
 * different one.
 */
export function spRoundTwoProbes(): readonly Probe[] {
  const out: Probe[] = [];
  let index = 0;
  for (const face of ['Arial', 'Times New Roman', 'Verdana', 'Courier New']) {
    for (const sz of [1200, 1800, 2800]) {
      for (const lines of [1, 4, 9]) {
        for (const ins of [NO_INSETS, DEFAULT_INSETS]) {
          out.push({
            ...BASE,
            id: `sp2-${String(index).padStart(3, '0')}`,
            deck: 'sp2',
            face,
            sz,
            rect: { x: 0, y: 0, w: 320, h: 400 },
            paragraphs: shortParagraphs(lines),
            expectLines: lines,
            insets: ins,
            autofit: 'sp',
            note: `spAutoFit on ${face} at ${String(sz / 100)}pt, ${String(lines)} lines`,
          });
          index += 1;
        }
      }
    }
  }

  // Why two probes of 54 kept a height PowerPoint should have shrunk. The
  // content is held still - four lines of 18pt, so the target height is one
  // number - and the box is swept through it a point at a time. Whatever the
  // rule is, it is a function of the distance between the box and the target,
  // and this is the sweep that reads it off.
  for (let h = 60; h <= 120; h += 1) {
    out.push({
      ...BASE,
      id: `sp-hyst-${String(index).padStart(3, '0')}`,
      deck: 'sp-hyst',
      rect: { x: 0, y: 0, w: 320, h },
      paragraphs: shortParagraphs(4),
      expectLines: 4,
      autofit: 'sp',
      note: `spAutoFit hysteresis: 4 lines of 18pt in a ${String(h)}pt box`,
    });
    index += 1;
  }

  return out;
}

/**
 * Fractional reductions, fractional percentages, and a negative result.
 *
 * `@lnSpcReduction` is a percentage like any other and nothing says it has to
 * be one of the ladder's two values. If a fraction is honoured, then the
 * question of *where* T2's round-half-up-to-a-whole-percent applies becomes
 * answerable: subtracting 12.5 from a quantised 100 is 87.5, and quantising
 * 100.4 - 12.5 is 88, and those are two different advances.
 *
 * The last group asks what a reduction larger than the spacing does. A 15%
 * line spacing reduced by 20% is -5%, which is either clamped, or ignored, or
 * taken literally and stacks every line on the one before it.
 */
export function reductionProbes(): readonly Probe[] {
  const out: Probe[] = [];
  let index = 0;
  const push = (lnSpc: Spacing | undefined, stored: StoredScale, note: string): void => {
    out.push({
      ...BASE,
      id: `red-${String(index).padStart(3, '0')}`,
      deck: 'red',
      rect: { x: 0, y: 0, w: 300, h: 400 },
      paragraphs: shortParagraphs(6),
      expectLines: 6,
      ...(lnSpc === undefined ? {} : { lnSpc }),
      stored,
      recompute: false,
      note,
    });
    index += 1;
  };

  for (const red of [5000, 7500, 12500, 17500, 22500]) {
    push(
      undefined,
      { fontScale: 62500, lnSpcReduction: red },
      `fractional reduction ${String(red)}`,
    );
  }
  for (const pct of [106667, 100400, 149600]) {
    for (const red of [0, 12500]) {
      push(
        { kind: 'percent', value: pct },
        { fontScale: 62500, lnSpcReduction: red },
        `lnSpc ${String(pct)} with reduction ${String(red)}: where the whole-percent quantisation applies`,
      );
    }
  }
  for (const pct of [15000, 30000, 20000]) {
    push(
      { kind: 'percent', value: pct },
      { fontScale: 100000, lnSpcReduction: 20000 },
      `lnSpc ${String(pct)} reduced by 20%, which is at or below zero`,
    );
  }
  return out;
}

/* -------------------------------------------------------------------------- */
/* round three                                                                */
/* -------------------------------------------------------------------------- */

/**
 * The height of the *last* line, which is what round two turned out to be
 * about.
 *
 * Round two's `ln150`, `ln080` and `ln-pts` sweeps all refused the model that
 * fitted everything else, and they refused it in a shape: the height PowerPoint
 * used was always a little less than `lines x advance`, by an amount that grew
 * as the line spacing grew. `ln-pts` made it undeniable - eight lines at an
 * exact 24pt cannot change height whatever the font scale does, and yet the
 * emitted rung stepped up four times in four consecutive points of box height.
 * Only a term that varies with the *effective size* and not with the advance
 * can do that.
 *
 * That term is the one sub-phase 3.2 left open. `blockHeight` in the line model
 * already takes it as a parameter, with a comment saying a plausible guess
 * would be exactly the silent wrong answer this repository forbids, and 3.6 was
 * where it was going to be measured. Autofit needs it three sub-phases early,
 * because the fit test is a comparison against it.
 *
 * So this deck measures it directly and stops inferring it. Six faces whose
 * vertical metrics disagree, seven sizes, and three line spacings chosen so
 * that the last line is *not* simply the advance:
 *
 * - **150%**, where the last line is shorter than the advance.
 * - **60%**, where it is longer.
 * - **an exact 30pt**, where the advance does not follow the size at all.
 *
 * Nothing autofits. `BoundHeight` for the whole range is the block height, the
 * differences between line tops are the advance, and the last line is what is
 * left over - one subtraction, no model in between.
 */
export function lastLineProbes(): readonly Probe[] {
  const out: Probe[] = [];
  let index = 0;
  const faces = ['Arial', 'Times New Roman', 'Verdana', 'Courier New', 'Georgia', 'Tahoma'];
  const sizes = [800, 1100, 1400, 1800, 2400, 3200, 4400];
  const spacings: readonly Spacing[] = [
    { kind: 'percent', value: 150000 },
    { kind: 'percent', value: 60000 },
    { kind: 'points', value: 3000 },
  ];
  for (const face of faces) {
    for (const sz of sizes) {
      for (const lnSpc of spacings) {
        out.push({
          ...BASE,
          id: `last-${String(index).padStart(3, '0')}`,
          deck: 'last',
          face,
          sz,
          rect: { x: 0, y: 0, w: 420, h: 540 },
          paragraphs: shortParagraphs(6),
          expectLines: 6,
          autofit: 'none',
          recompute: false,
          lnSpc,
          note: `${face} ${String(sz / 100)}pt, lnSpc ${lnSpc.kind}:${String(lnSpc.value)}`,
        });
        index += 1;
      }
    }
  }
  return out;
}

/**
 * Two loose ends round two exposed and could not close.
 *
 * `frac` asks whether a fractional nominal size survives a font scale of
 * exactly 100%. `ladder-frac` says it does - 11 lines of 10.5pt fitted a 139pt
 * box, which needs 12.6pt a line and not the 13.2 that rounding to 11pt would
 * give - but that is a rung read through a fit test, and the direct reading is
 * one line advance.
 *
 * `zero` asks what a line spacing reduced to nothing does. A 15% spacing
 * reduced by 20% is -5%, and PowerPoint laid it out at 0.215pt a line rather
 * than at zero or at something negative. One size is an anecdote; five say
 * whether the floor is a length or a fraction.
 */
export function loseEndsProbes(): readonly Probe[] {
  const out: Probe[] = [];
  let index = 0;
  for (const sz of [1050, 1150, 1350, 950, 1800]) {
    for (const stored of [
      { fontScale: 100000 },
      { fontScale: 100000, lnSpcReduction: 10000 },
      { lnSpcReduction: 10000 },
    ]) {
      out.push({
        ...BASE,
        id: `frac-${String(index).padStart(3, '0')}`,
        deck: 'frac',
        sz,
        rect: { x: 0, y: 0, w: 300, h: 400 },
        paragraphs: shortParagraphs(6),
        expectLines: 6,
        stored,
        recompute: false,
        note: `${String(sz / 100)}pt at a font scale of exactly 100%`,
      });
      index += 1;
    }
  }
  for (const sz of [800, 1200, 1800, 2400, 4000]) {
    out.push({
      ...BASE,
      id: `zero-${String(index).padStart(3, '0')}`,
      deck: 'frac',
      sz,
      rect: { x: 0, y: 0, w: 300, h: 400 },
      paragraphs: shortParagraphs(6),
      expectLines: 6,
      lnSpc: { kind: 'percent', value: 15000 },
      stored: { lnSpcReduction: 20000 },
      recompute: false,
      note: `${String(sz / 100)}pt with a line spacing reduced below zero`,
    });
    index += 1;
  }
  return out;
}

/**
 * The whole rule again, end to end, on a face and a size nothing else used.
 *
 * If the last-line term is a property of the typeface then Times New Roman has
 * a different one from Arial, and a height sweep under 150% line spacing is
 * where that shows: every rung's threshold moves by the difference. This deck
 * is the check that the model built from Arial predicts a face it was not
 * built from - which is the only version of "it works" worth reporting.
 */
export function roundThreeSweeps(): readonly Probe[] {
  return [
    ...sweepDeck('ln150-tnr', {
      lines: 7,
      sz: 2400,
      face: 'Times New Roman',
      widthPt: 360,
      fromPt: 20,
      toPt: 320,
      lnSpc: { kind: 'percent', value: 150000 },
      note: 'Times New Roman 24pt at 150%: the model predicting a face it was not built from',
    }),
    ...sweepDeck('ln060-verdana', {
      lines: 9,
      sz: 1600,
      face: 'Verdana',
      widthPt: 360,
      fromPt: 12,
      toPt: 190,
      lnSpc: { kind: 'percent', value: 60000 },
      note: 'Verdana 16pt at 60%, where the last line is taller than the advance',
    }),
  ];
}

/**
 * Where the last line stops shrinking, read off a sweep instead of guessed.
 *
 * Round three fitted `lastLine = 0.75 x advance + b x size` to six faces with a
 * worst residual of four thousandths of a point, and found a floor besides: at
 * single spacing the last line measures the *advance*, which is larger. So
 * there is a threshold, and the obvious candidate is the natural 1.2 line box -
 * below it the advance wins, above it the fitted line does.
 *
 * That is right for five faces and wrong for Courier New, whose last line at
 * 24pt under a 30pt spacing measures 30 and not the 29.7 the fit predicts. A
 * 30pt advance against a 28.8pt box is *above* 1.2, so the threshold is not 1.2
 * for every face - and rather than reach for a font table and guess which of
 * ascent, descent and line gap PowerPoint added up, this sweep asks where the
 * threshold is.
 *
 * The trick is the same one the ladder used. Hold the line spacing at an exact
 * 30pt and sweep the *font size*: the advance cannot move, the natural box
 * grows a fifth of a point at a time, and the size at which the last line
 * switches from the fitted value to a flat 30 is the threshold, divided out.
 * A second sweep at 20pt spacing says whether the threshold is proportional to
 * the size, which one sweep cannot.
 */
export function floorProbes(): readonly Probe[] {
  const out: Probe[] = [];
  let index = 0;
  const faces = ['Arial', 'Times New Roman', 'Verdana', 'Courier New', 'Georgia', 'Tahoma'];
  for (const face of faces) {
    for (let sz = 1600; sz <= 2800; sz += 25) {
      out.push({
        ...BASE,
        id: `floor-${String(index).padStart(3, '0')}`,
        deck: 'floor',
        face,
        sz,
        rect: { x: 0, y: 0, w: 420, h: 400 },
        paragraphs: shortParagraphs(6),
        expectLines: 6,
        autofit: 'none',
        recompute: false,
        lnSpc: { kind: 'points', value: 3000 },
        note: `${face} at ${String(sz / 100)}pt under an exact 30pt line spacing`,
      });
      index += 1;
    }
  }
  for (const face of ['Arial', 'Courier New']) {
    for (let sz = 1100; sz <= 1900; sz += 25) {
      out.push({
        ...BASE,
        id: `floor2-${String(index).padStart(3, '0')}`,
        deck: 'floor',
        face,
        sz,
        rect: { x: 0, y: 0, w: 420, h: 300 },
        paragraphs: shortParagraphs(6),
        expectLines: 6,
        autofit: 'none',
        recompute: false,
        lnSpc: { kind: 'points', value: 2000 },
        note: `${face} at ${String(sz / 100)}pt under an exact 20pt line spacing`,
      });
      index += 1;
    }
  }
  return out;
}

export function allProbes(): readonly Probe[] {
  return [
    ...floorProbes(),
    ...ladderProbes(),
    ...storedProbes(),
    ...spacingProbes(),
    ...lnSpcProbes(),
    ...insetProbes(),
    ...wrapProbes(),
    ...spProbes(),
    ...roundTwoSweeps(),
    ...spRoundTwoProbes(),
    ...reductionProbes(),
    ...lastLineProbes(),
    ...loseEndsProbes(),
    ...roundThreeSweeps(),
  ];
}

/** The decks, in the order they are built and run. */
export function deckNames(): readonly string[] {
  const seen = new Set<string>();
  for (const p of allProbes()) seen.add(p.deck);
  return [...seen];
}
