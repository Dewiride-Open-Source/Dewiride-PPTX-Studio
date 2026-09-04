/**
 * Experiment T2 - what is a line box, and what is an advance?
 *
 * Sub-phase 3.2 rests on one claim the plan states without evidence: that
 * PowerPoint's line height is a font-independent `1.2 x font size`. That claim
 * is worth exactly as much as the measurement behind it, and there is none, so
 * this table builds the measurement.
 *
 * ## The design, and why the increment rather than the height
 *
 * A text box's `BoundHeight` is the line box plus whatever padding the first
 * and last lines carry. Reading it once tells you a sum you cannot decompose.
 * Reading it for one line and again for three gives `(H3 - H1) / 2`, which is
 * the line box on its own with every constant term cancelled - and it cancels
 * them whether or not the constant is what anyone expects. That is the whole
 * reason the `lineHeight` probes come in ones, twos and threes.
 *
 * Extra lines are made with `a:br`, not with extra paragraphs. A paragraph
 * carries `spcBef`/`spcAft`, and while both can be zeroed, a measurement that
 * needs two things zeroed to be true is a measurement of two things. `a:br` is
 * a hard line break with no paragraph spacing anywhere near it. A smaller set
 * of `paraLines` probes does use paragraphs, with the spacing explicitly
 * zeroed, so that "a line is a line however it was made" is checked rather than
 * assumed.
 *
 * ## Falsification, not confirmation
 *
 * Thirteen typefaces, chosen because their vertical metrics disagree violently
 * - Impact is a tall condensed face, Verdana a wide one with a huge x-height,
 * Courier New a monospace with a small one. If the line box were derived from
 * `hhea` or from `OS/2.usWinAscent`, as `line-height: normal` is in a browser,
 * these thirteen would report thirteen different increments. If the plan is
 * right they report one number, and it is `1.2 x size`.
 *
 * A fourteenth entry names a face that does not exist. It is not padding: it
 * measures what substitution does to the numbers, and it lets the analysis
 * prove that its "no two fonts may share an advance row" check can actually
 * fire - which is how a substituted face would otherwise slip through looking
 * like a measurement.
 */

/** Hundredths of a point, as `a:rPr/@sz`, `@spc` and `@kern` all are. */
export type Hundredths = number;

export interface LnSpc {
  readonly kind: 'pct' | 'pts';
  /** `spcPct/@val` in thousandths of a percent, or `spcPts/@val` in hundredths of a point. */
  readonly value: number;
}

export type ProbeKind =
  | 'advance'
  | 'pilcrow'
  | 'lineHeight'
  | 'paraLines'
  | 'lnSpc'
  | 'lastLine'
  | 'spc'
  | 'kern'
  | 'chars';

export interface Probe {
  readonly id: string;
  readonly kind: ProbeKind;
  readonly font: string;
  /** `a:rPr/@sz`, hundredths of a point. */
  readonly sz: Hundredths;
  readonly text: string;
  /** Rendered lines. Made with `a:br` unless `kind` is `paraLines`. */
  readonly lines: number;
  readonly lnSpc?: LnSpc | undefined;
  readonly spc?: Hundredths | undefined;
  readonly kern?: Hundredths | undefined;
  /** Ask the reader for per-character bounds. Expensive, so a handful only. */
  readonly perChar?: boolean | undefined;
}

/**
 * Thirteen installed faces plus one that is not.
 *
 * Every one of these ships with Windows 11 or with Office on this machine. That
 * is an assumption, and it is the one assumption the analysis checks first: two
 * faces reporting the same advance for the same string at the same size are two
 * names for one file, and the analysis says so by name rather than averaging
 * them into a table.
 */
export const FONTS: readonly string[] = [
  'Arial',
  'Times New Roman',
  'Courier New',
  'Georgia',
  'Verdana',
  'Tahoma',
  'Trebuchet MS',
  'Impact',
  'Comic Sans MS',
  'Consolas',
  'Calibri',
  'Cambria',
  'Segoe UI',
];

/** Named so a substitution is unmistakable in the readings. */
export const ABSENT_FONT = 'Zzz Probe Face That Is Not Installed';

/** The three faces that carry the probes where one font is enough. */
const TRIO: readonly string[] = ['Arial', 'Georgia', 'Consolas'];

/**
 * Sizes spanning the range a deck actually uses, plus one fractional one.
 *
 * 10.5pt is there because `@sz` is hundredths of a point and every
 * implementation that rounds it to a whole point somewhere still passes every
 * whole-point test.
 */
export const LINE_SIZES: readonly Hundredths[] = [800, 1050, 1200, 1800, 3200, 5400, 9600];
const ADVANCE_SIZES: readonly Hundredths[] = [1200, 1800, 3200, 5400];

/**
 * Strings chosen for what each one can refute.
 *
 * `Hamburgefonstiv` is the classic metric string. The digits ask whether
 * figures are tabular. `AVATAR Yo To Wa` is dense in kerning pairs, so it is
 * the string on which `@kern` has to show. `iiiii`/`WWWWW` are the narrow and
 * wide extremes, which is where a per-character rounding error accumulates
 * fastest. The trailing-space string asks whether a trailing space is measured,
 * which decides whether a line width is its advance or its ink.
 */
export const STRINGS: readonly { readonly key: string; readonly text: string }[] = [
  { key: 'hamburg', text: 'Hamburgefonstiv' },
  { key: 'digits', text: '0123456789' },
  { key: 'kernpairs', text: 'AVATAR Yo To Wa' },
  { key: 'narrow', text: 'iiiii' },
  { key: 'wide', text: 'WWWWW' },
  { key: 'sentence', text: 'The quick brown fox' },
  { key: 'trailing', text: 'ab   ' },
];

/** `@spc` in hundredths of a point: none, a quarter point, three points, negative. */
const SPC_VALUES: readonly Hundredths[] = [0, 25, 300, -50];

/**
 * `@kern` is a *minimum font size*, in hundredths of a point - not a flag.
 * ECMA-376 Part 1, `CT_TextCharacterProperties/@kern`: "Specifies the minimum
 * font size at which character kerning occurs for this text run." So the
 * experiment has to cross the threshold with size, and a model that reads the
 * attribute as a boolean has to come out wrong on the crossing.
 */
const KERN_VALUES: readonly Hundredths[] = [0, 100, 1200, 4000];
const KERN_SIZES: readonly Hundredths[] = [800, 1200, 2400];

/**
 * Explicit single spacing.
 *
 * Written out rather than left absent on every probe that is not asking about
 * line spacing, because "absent" and "100%" are two different files and only a
 * measurement says whether they lay out the same. The `lhdef` set below is the
 * one that leaves it out, so the difference is measured once rather than
 * assumed everywhere.
 */
export const PCT_100: LnSpc = { kind: 'pct', value: 100000 };

/** Percent line spacing, and exact-point line spacing, across the interesting range. */
const LN_SPC: readonly LnSpc[] = [
  { kind: 'pct', value: 100000 },
  { kind: 'pct', value: 150000 },
  { kind: 'pct', value: 200000 },
  { kind: 'pct', value: 80000 },
  { kind: 'pts', value: 1200 },
  { kind: 'pts', value: 3600 },
];

/**
 * Six faces for the fine line-spacing sweep, and every face for one point of it.
 *
 * The six are the metric extremes plus a monospace, which is where a
 * font-dependent term shows largest. Running all thirteen through twenty-one
 * spacings would be another eight hundred shapes for a term that only has to be
 * shown to vary; the `lastall` block below covers the remaining seven at one
 * spacing so the committed table is complete even though the sweep is not.
 */
const LAST_LINE_FONTS: readonly string[] = [
  'Arial',
  'Times New Roman',
  'Courier New',
  'Verdana',
  'Impact',
  'Consolas',
];

const LAST_LINE_SIZES: readonly Hundredths[] = [1200, 1800, 3200];

/**
 * Twenty-one spacings, placed either side of every boundary a candidate model
 * could branch on: 100% exactly, just under and just over it, and an exact-point
 * spacing both far below and far above the natural line box at each size.
 */
const LAST_LINE_SPACINGS: readonly LnSpc[] = [
  ...[
    50000, 75000, 90000, 95000, 100000, 105000, 110000, 120000, 125000, 150000, 175000, 200000,
    300000,
  ].map((value): LnSpc => ({ kind: 'pct', value })),
  ...[600, 900, 1200, 1500, 1800, 2400, 3600, 7200].map((value): LnSpc => ({ kind: 'pts', value })),
];

/**
 * Percent line spacing through the degenerate interval, at one percent a step
 * and then at a fifth of one.
 *
 * The half-percent points settled that `@val` is quantised before use - 99.5%
 * lays out as 100% and 100.5% as 101% - but a half-percent step cannot tell
 * round-half-up from ceiling, since both send 99.5 to 100 and 100.5 to 101. The
 * fifths do: round sends 100.2 and 100.4 down and 100.6 and 100.8 up, ceiling
 * sends all four up, floor sends all four down. `@val` is in thousandths of a
 * percent and producers other than PowerPoint do write values like 106667, so
 * which of the three it is decides real line counts and not only this table.
 */
const BOUNDARY_PCTS: readonly number[] = [
  96000, 98000, 99000, 99500, 100000, 100200, 100400, 100500, 100600, 100800, 101000, 102000,
  103000, 104000, 105000, 106000, 107000, 108000, 109000, 110000, 112000, 115000,
];

function id(parts: readonly (string | number)[]): string {
  return parts
    .map((p) => String(p).replace(/[^A-Za-z0-9-]+/g, '_'))
    .join('-')
    .toLowerCase();
}

export function metricProbes(): readonly Probe[] {
  const probes: Probe[] = [];

  // ---- advances -----------------------------------------------------------
  // Every installed face, every size, every string. This is the reference table
  // a browser measurer has to reproduce.
  for (const font of [...FONTS, ABSENT_FONT]) {
    for (const sz of ADVANCE_SIZES) {
      for (const s of STRINGS) {
        probes.push({
          id: id(['adv', font, sz, s.key]),
          kind: 'advance',
          font,
          sz,
          text: s.text,
          lines: 1,
          lnSpc: PCT_100,
        });
      }
    }
  }

  // ---- the paragraph mark -------------------------------------------------
  // `TextRange2.BoundWidth` is wider than the string, by an amount that is
  // constant for a given face and size and happens to equal that face's space
  // advance: the range includes the paragraph mark. Comparing raw `BoundWidth`
  // against a browser therefore reports a 5-20% disagreement that is entirely
  // an artefact of the reader, and every advance in the table has to have this
  // subtracted before it means anything.
  //
  // Two probes measure it without needing to know what the mark is. `HH` minus
  // `H` is one `H` advance with the mark cancelling; `H` minus that advance is
  // the mark. Neither step needs a font file, which is what makes it usable.
  for (const font of [...FONTS, ABSENT_FONT]) {
    for (const sz of ADVANCE_SIZES) {
      for (const text of ['H', 'HH']) {
        probes.push({
          id: id(['pil', font, sz, text.length]),
          kind: 'pilcrow',
          font,
          sz,
          text,
          lines: 1,
          lnSpc: PCT_100,
        });
      }
    }
  }

  // ---- the line box -------------------------------------------------------
  // One line and three, so the increment falls out with the first line's
  // padding cancelled. Two as well, so "the increment is constant" is a
  // measured fact rather than an interpolation between two points.
  for (const font of FONTS) {
    for (const sz of LINE_SIZES) {
      for (const lines of [1, 2, 3]) {
        probes.push({
          id: id(['lh', font, sz, lines]),
          kind: 'lineHeight',
          font,
          sz,
          text: 'Hxy',
          lines,
          lnSpc: PCT_100,
        });
      }
    }
  }

  // The control for the block above: a file that states no `a:lnSpc` at all. If
  // these disagree with the explicit-100% probes then single spacing is not the
  // default, and every other probe here is measuring against the wrong
  // baseline.
  for (const font of FONTS) {
    for (const sz of [1200, 3200]) {
      for (const lines of [1, 3]) {
        probes.push({
          id: id(['lhdef', font, sz, lines]),
          kind: 'lineHeight',
          font,
          sz,
          text: 'Hxy',
          lines,
        });
      }
    }
  }

  // A paragraph-made line, with paragraph spacing zeroed, must measure the same
  // as an `a:br`-made one. If it does not, the line box is not the only term.
  for (const font of TRIO) {
    for (const sz of [1200, 1800, 3200]) {
      for (const lines of [1, 3]) {
        probes.push({
          id: id(['para', font, sz, lines]),
          kind: 'paraLines',
          font,
          sz,
          text: 'Hxy',
          lines,
          lnSpc: PCT_100,
        });
      }
    }
  }

  // ---- line spacing -------------------------------------------------------
  // The fork the plan does not name: is `spcPct val="150000"` 1.5 x the line
  // box, or 1.5 x the font size? At a 1.2 line box those two differ by 20% and
  // nothing else in the file distinguishes them.
  for (const font of TRIO) {
    for (const sz of [1200, 1800, 3200]) {
      for (const lnSpc of LN_SPC) {
        for (const lines of [1, 3]) {
          probes.push({
            id: id(['ls', font, sz, lnSpc.kind, lnSpc.value, lines]),
            kind: 'lnSpc',
            font,
            sz,
            text: 'Hxy',
            lines,
            lnSpc,
          });
        }
      }
    }
  }

  // ---- the last line ------------------------------------------------------
  // The first pass measured six line spacings and found that every line box
  // except the last equals the advance, while the last one does not, and is
  // font-dependent besides. Six points is an anecdote: two of them fitted
  // `advance` and four fitted a quite different expression, and no rule
  // separating them can be read off six samples. This sweep is what makes it a
  // rule or refutes it - twenty-one spacings either side of every boundary a
  // candidate model could put a branch on, with two- and three-line controls so
  // that "only the last line is special" is re-checked at each one.
  for (const font of LAST_LINE_FONTS) {
    for (const sz of LAST_LINE_SIZES) {
      for (const lnSpc of LAST_LINE_SPACINGS) {
        for (const lines of font === 'Arial' ? [1, 2, 3] : [1]) {
          probes.push({
            id: id(['last', font, sz, lnSpc.kind, lnSpc.value, lines]),
            kind: 'lastLine',
            font,
            sz,
            text: 'Hxy',
            lines,
            lnSpc,
          });
        }
      }
    }
  }

  // The boundary, at one percent a step.
  //
  // The twenty-one-point sweep fitted 519 of 523 one-line probes, and all four
  // misses were Courier New just above 100%. Courier New is the face whose
  // font-dependent term is exactly 0.3 x the size, which is the one value that
  // makes the two branches of the fitted rule meet at 100% instead of stepping
  // - so the misses are all in the interval where the rule is degenerate and
  // the sweep has no samples. Four exceptions on the one font that cannot
  // discriminate is not a rule with an exception, it is a rule with an
  // unsampled interval, and this closes it.
  for (const font of LAST_LINE_FONTS) {
    for (const sz of LAST_LINE_SIZES) {
      for (const value of BOUNDARY_PCTS) {
        // Three lines as well as one, on the one face that carries the fine
        // sweep. A one-line box only reveals the quantisation through the
        // last-line height, which is itself a fitted model; three lines put two
        // baselines on the slide and measure the advance outright, so the
        // quantisation finding does not rest on the model it was found with.
        for (const lines of font === 'Arial' ? [1, 3] : [1]) {
          probes.push({
            id: id(['bound', font, sz, value, lines]),
            kind: 'lastLine',
            font,
            sz,
            text: 'Hxy',
            lines,
            lnSpc: { kind: 'pct', value },
          });
        }
      }
    }
  }

  // The same question asked of every installed face at one spacing well inside
  // the regime where the last line stops equalling the advance, so the table
  // covers all thirteen rather than the six the sweep can afford.
  for (const font of FONTS) {
    for (const sz of LINE_SIZES) {
      probes.push({
        id: id(['lastall', font, sz]),
        kind: 'lastLine',
        font,
        sz,
        text: 'Hxy',
        lines: 1,
        lnSpc: { kind: 'pct', value: 200000 },
      });
    }
  }

  // ---- letter spacing -----------------------------------------------------
  // Ten characters, so `n x spc` and `(n-1) x spc` differ by a tenth of the
  // total - far outside any rounding the reader could introduce.
  for (const font of TRIO) {
    for (const sz of [1200, 3200]) {
      for (const spc of SPC_VALUES) {
        probes.push({
          id: id(['spc', font, sz, spc]),
          kind: 'spc',
          font,
          sz,
          text: '0123456789',
          lines: 1,
          lnSpc: PCT_100,
          spc,
        });
      }
    }
  }

  // ---- kerning ------------------------------------------------------------
  for (const font of TRIO) {
    for (const sz of KERN_SIZES) {
      for (const kern of KERN_VALUES) {
        probes.push({
          id: id(['kern', font, sz, kern]),
          kind: 'kern',
          font,
          sz,
          text: 'AVATAR Yo To Wa',
          lines: 1,
          lnSpc: PCT_100,
          kern,
        });
      }
    }
  }

  // ---- kerning with no @kern at all ---------------------------------------
  // ECMA says an omitted `@kern` kerns at every size. The advance probes agree
  // at 12pt - they set no `@kern` and measure as kerned - but 12pt cannot tell
  // "omitted means always on" from "omitted, and the cascade supplied 1200",
  // because 1200 is on at 12pt either way. Below 12pt the two part company, and
  // that is the only place the question can be asked.
  //
  // Arial alone, because the readings say Arial is the only one of the three
  // faces with kern pairs for this string - which is why three were measured.
  for (const sz of [800, 1000, 1200]) {
    probes.push({
      id: id(['kerndef', 'Arial', sz]),
      kind: 'kern',
      font: 'Arial',
      sz,
      text: 'AVATAR Yo To Wa',
      lines: 1,
      lnSpc: PCT_100,
    });
  }

  // ---- per-character ------------------------------------------------------
  // The whole-string width is one number and many models fit it. The pen
  // position after each character is fifteen numbers and almost none do.
  for (const font of ['Arial', 'Georgia', 'Consolas', 'Verdana']) {
    for (const sz of [1800, 5400]) {
      for (const key of ['hamburg', 'kernpairs']) {
        const s = STRINGS.find((x) => x.key === key);
        if (s === undefined) throw new Error(`no string named ${key}`);
        probes.push({
          id: id(['ch', font, sz, key]),
          kind: 'chars',
          font,
          sz,
          text: s.text,
          lines: 1,
          lnSpc: PCT_100,
          perChar: true,
        });
      }
    }
  }

  return probes;
}
