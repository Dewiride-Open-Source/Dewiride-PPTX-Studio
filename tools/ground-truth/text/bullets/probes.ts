/**
 * Experiment T5 - the probes: bullets, fields and script runs.
 *
 * ## What the authoring step already settled
 *
 * `tools/ground-truth/text/bullets/author.ps1` walked `PpNumberedBulletStyle` and read the saved XML,
 * and PowerPoint wrote out the whole of `ST_TextAutonumberScheme` itself: 41
 * values, 0 through 40, with 41 and up refused as out of range. So the
 * vocabulary in `SCHEMES` below is not transcribed from the standard - it is
 * what the format's own author produced, in order, and the analysis asserts the
 * two agree.
 *
 * It settled four more things that shape every probe here:
 *
 * - `a:buChar/@char` keeps a private-use code point **verbatim**. Asked for
 *   U+F0A7 PowerPoint writes U+F0A7; asked for U+00A7 it writes U+00A7. It
 *   never folds one into the other, so the PUA question is not about what
 *   PowerPoint writes but about what it *draws* for each - which is `charProbes`.
 * - "use the text's colour" is expressed by writing **nothing at all**, not by
 *   `a:buClrTx`. `a:buClrTx` therefore only ever appears to cancel something
 *   inherited, which is what `cascadeProbes` is for.
 * - A six-paragraph numbered list carries **no `startAt` anywhere**. The numbers
 *   are not in the file; a reader has to count. `sequenceProbes` asks what the
 *   counting rule is, and it is the only family here whose answer cannot be
 *   guessed from a single paragraph.
 * - A field's cached `a:t` is **discarded on open**. A deck reopened four
 *   minutes later showed `12:04 PM` where the file said `12:02 PM`, so the
 *   cache is a fallback for readers that cannot compute, never the answer.
 *
 * ## The measurement
 *
 * Two instruments, and the first one is new to this directory.
 *
 * **The EMF says what was drawn.** `Slide.Export(path, "EMF")` records
 * PowerPoint's own GDI calls, and a bullet is drawn by `ExtTextOutW` with the
 * characters in it. So the rendered string comes back as a string - `"(i)"`,
 * `"a."` - along with the face, the size and the colour it was drawn with.
 * Nothing is fitted and no candidate has to be enumerated to identify it.
 *
 * **COM says where it landed.** `Paragraphs(i).BoundLeft` reports the text's
 * left edge in points, exactly, and with `marL="0" indent="0"` that is the
 * bullet's advance width and nothing else - measured in the probe run at
 * 26.625pt for `"1."` in a face whose digit is 0.5546875em and whose period is
 * 0.2773em, which is that string and no other. The two instruments fail
 * differently: the EMF cannot say where the text went and COM cannot say what
 * the characters were.
 *
 * ## Attribution
 *
 * EMF records arrive in draw order with no shape names, so every probe
 * paragraph ends with a marker run - `#12#` - and the rule is that **everything
 * strictly between two markers belongs to the probe whose marker follows it**.
 * A bullet is drawn before its own paragraph's text, so it lands in that
 * window; a script-run probe puts its subject text before the marker for the
 * same reason. `#` appears in no bullet any scheme can produce and in no probe
 * string here, which the builder asserts.
 */

import { EMU_PER_POINT } from '../../lib/sheet-pptx.ts';

/* -------------------------------------------------------------------------- */
/* the vocabulary                                                             */
/* -------------------------------------------------------------------------- */

/**
 * `ST_TextAutonumberScheme`, in `PpNumberedBulletStyle` order.
 *
 * Index is the object-model value: `SCHEMES[3]` is `arabicPeriod` because
 * `ppBulletArabicPeriod` is 3. `tools/ground-truth/text/bullets/analyse.ts` re-derives this list from
 * `author-bullets-log.json` and throws if it disagrees, so the order is a
 * measurement rather than a transcription.
 */
export const SCHEMES: readonly string[] = [
  'alphaLcPeriod',
  'alphaUcPeriod',
  'arabicParenR',
  'arabicPeriod',
  'romanLcParenBoth',
  'romanLcParenR',
  'romanLcPeriod',
  'romanUcPeriod',
  'alphaLcParenBoth',
  'alphaLcParenR',
  'alphaUcParenBoth',
  'alphaUcParenR',
  'arabicParenBoth',
  'arabicPlain',
  'romanUcParenBoth',
  'romanUcParenR',
  'ea1ChsPlain',
  'ea1ChsPeriod',
  'circleNumDbPlain',
  'circleNumWdWhitePlain',
  'circleNumWdBlackPlain',
  'ea1ChtPlain',
  'ea1ChtPeriod',
  'arabic2Minus',
  'arabic1Minus',
  'hebrew2Minus',
  'ea1JpnKorPlain',
  'ea1JpnKorPeriod',
  'arabicDbPlain',
  'arabicDbPeriod',
  'thaiAlphaPeriod',
  'thaiAlphaParenR',
  'thaiAlphaParenBoth',
  'thaiNumPeriod',
  'thaiNumParenR',
  'thaiNumParenBoth',
  'hindiAlphaPeriod',
  'hindiNumPeriod',
  'ea1JpnChsDbPeriod',
  'hindiNumParenR',
  'hindiAlpha1Period',
];

/**
 * The reserved `a:fld/@type` names PowerPoint itself produced.
 *
 * Sixteen date formats, not the thirteen `PpDateTimeFormat` documents: formats
 * 15 and 16 write `uaqdatetime1` and `uaqdatetime2`, which no enumeration this
 * project has read mentions. Plus `slidenum`, and `datetime` - which
 * PowerPoint's UI never writes but ECMA-376 reserves, so it is probed as a
 * value only another producer can put in a file.
 */
export const FIELD_TYPES: readonly string[] = [
  'slidenum',
  'datetime',
  'datetime1',
  'datetime2',
  'datetime3',
  'datetime4',
  'datetime5',
  'datetime6',
  'datetime7',
  'datetime8',
  'datetime9',
  'datetime10',
  'datetime11',
  'datetime12',
  'datetime13',
  'datetimeFigureOut',
  'uaqdatetime1',
  'uaqdatetime2',
];

/* -------------------------------------------------------------------------- */
/* probe shapes                                                               */
/* -------------------------------------------------------------------------- */

/** `a:lnSpc`/`a:spcBef`/`a:spcAft`, as the file states them. */
export interface Spacing {
  readonly kind: 'percent' | 'points';
  readonly value: number;
}

/** What a paragraph's `a:pPr` says about its bullet. */
export interface Bullet {
  /**
   * `none` writes `a:buNone`; `inherit` writes no bullet element at all, which
   * is a different question and the only way to ask what a level inherits.
   */
  readonly kind: 'none' | 'inherit' | 'char' | 'autonum' | 'blip';
  /** `a:buChar/@char`. */
  readonly char?: string | undefined;
  /** `a:buAutoNum/@type`. */
  readonly scheme?: string | undefined;
  /** `a:buAutoNum/@startAt`. Omitted when undefined, so absence is a case. */
  readonly startAt?: number | undefined;
  /** `a:buFont/@typeface`; `'tx'` writes `a:buFontTx`; absent writes neither. */
  readonly font?: string | undefined;
  /** `a:buSzPct/@val`, thousandths of a percent. */
  readonly szPct?: number | undefined;
  /** `a:buSzPts/@val`, hundredths of a point. */
  readonly szPts?: number | undefined;
  /** `a:buSzTx`. */
  readonly szTx?: boolean | undefined;
  /** `a:buClr` as a 6-hex-digit `a:srgbClr`, or a `a:schemeClr` name. */
  readonly clr?: string | undefined;
  /** `a:buClrTx`. */
  readonly clrTx?: boolean | undefined;
  /** `a:buBlip`, naming a media file the builder embeds. */
  readonly blip?: string | undefined;
}

/** One `a:p` inside a probe shape. */
export interface Para {
  /** `a:pPr/@lvl`. Omitted when 0. */
  readonly level?: number | undefined;
  /** `@marL` in points. Omitted when undefined. */
  readonly marL?: number | undefined;
  /** `@indent` in points, usually negative. */
  readonly indent?: number | undefined;
  /** `@defTabSz` in points. */
  readonly defTabSz?: number | undefined;
  readonly algn?: 'l' | 'ctr' | 'r' | 'just' | undefined;
  readonly rtl?: boolean | undefined;
  readonly bullet: Bullet;
  /** The body text, before the marker. Empty means the marker alone. */
  readonly text?: string | undefined;
  /**
   * Split `text` down the middle with an `a:br`.
   *
   * A flag rather than a sentinel string. The first version of this keyed off
   * the literal text `'aabb'`, which coupled two files through a magic value
   * and hid a stray U+0001 in the source for long enough that PowerPoint
   * refused the package rather than answering the question.
   */
  readonly hardBreak?: boolean | undefined;
  /** `a:rPr/@sz`, hundredths of a point. */
  readonly sz: number;
  /** `a:latin/@typeface`. */
  readonly face: string;
  readonly ea?: string | undefined;
  readonly cs?: string | undefined;
  readonly sym?: string | undefined;
  readonly lang?: string | undefined;
  /** `a:defRPr/@sz` on the paragraph, which `buSzPct` may or may not follow. */
  readonly defSz?: number | undefined;
  /**
   * A second run after the first, at its own size.
   *
   * The only way to ask which run a bullet takes its size from: with a single
   * run, "the first run" and "the paragraph" are the same answer.
   */
  readonly secondRun?: { readonly text: string; readonly sz: number } | undefined;
  /** A run colour, so a bullet that inherits one can be told from one that does not. */
  readonly runClr?: string | undefined;
  /** An `a:fld` in place of a run: `[type, cachedText]`. */
  readonly field?: readonly [string, string] | undefined;
}

/** One probe: one shape, one or more paragraphs. */
export interface Probe {
  readonly id: string;
  readonly deck: string;
  readonly family: string;
  readonly paragraphs: readonly Para[];
  /** Frame width in points. */
  readonly width: number;
  /** Frame height in points. */
  readonly height: number;
  /** `a:bodyPr/@wrap`. */
  readonly wrap: boolean;
  /** Read every character's position too. Expensive, so only where it is the question. */
  readonly readChars?: boolean | undefined;
  /** What this probe is asking, in one line, carried into the fixture. */
  readonly asks: string;
}

/** Points to EMU, rejecting anything that is not a whole number of them. */
export function emuPt(points: number, what: string): number {
  const emu = points * EMU_PER_POINT;
  if (!Number.isInteger(emu)) {
    throw new Error(`${what}: ${String(points)}pt is ${String(emu)} EMU, not a whole number`);
  }
  return emu;
}

/* -------------------------------------------------------------------------- */
/* the constants every probe shares                                           */
/* -------------------------------------------------------------------------- */

/**
 * 24pt, and a frame 44pt tall.
 *
 * Small enough for twenty probes a slide and large enough that the quarter-point
 * granularity `BoundLeft` reports is a thousandth of a bullet advance. T4 swept
 * box heights and needed the resolution; here every reading is a single
 * position, so the size is chosen for probes-per-slide instead.
 */
export const SIZE = 2400;

/** Arial, because it is present on every Windows machine and its metrics are in T2's fixture. */
export const FACE = 'Arial';

/** A monospace second instrument: every glyph 0.6em, so an advance *is* a character count. */
export const MONO = 'Courier New';

/**
 * The start values every scheme is asked for.
 *
 * Chosen so that each family of candidate rules is separated by at least one of
 * them: 4 and 9 separate subtractive roman numerals from additive, 26 and 27
 * separate bijective base-26 letters from the a..z-then-aa-bb form Word uses,
 * 10 and 11 separate a digit rollover from a letter one, and 3999 and 4000 ask
 * what happens past the largest roman numeral anybody writes.
 */
export const STARTS: readonly number[] = [
  1, 2, 3, 4, 5, 9, 10, 11, 20, 26, 27, 50, 100, 400, 900, 1000, 3999, 4000,
];

/* -------------------------------------------------------------------------- */
/* F1 - the 41 schemes                                                        */
/* -------------------------------------------------------------------------- */

function autonum(scheme: string, startAt: number, font: string): Bullet {
  return { kind: 'autonum', scheme, startAt, font };
}

function para(bullet: Bullet, over: Partial<Para> = {}): Para {
  return { bullet, sz: SIZE, face: FACE, marL: 0, indent: 0, ...over };
}

/**
 * Every scheme at every start value, twice: once in Arial and once in Courier
 * New.
 *
 * The second face is not redundancy. The EMF gives the string outright, so the
 * width is a check on it rather than the primary reading - and in a monospace
 * face the check is arithmetic a person can do: an advance of 14.4pt at 24pt is
 * exactly one glyph, 28.8 is two, and a model that emits `"iiii."` where
 * PowerPoint drew `"iv."` is off by two whole glyphs rather than by a fraction
 * of one. Courier New also has no Thai, Devanagari or CJK glyphs, so the pair
 * says which schemes fall back to another face and which do not.
 */
export function schemeProbes(): Probe[] {
  const out: Probe[] = [];
  for (const face of [FACE, MONO]) {
    for (const scheme of SCHEMES) {
      for (const startAt of STARTS) {
        out.push({
          id: `scheme-${face === MONO ? 'mono' : 'prop'}-${scheme}-${String(startAt)}`,
          deck: face === MONO ? 'scheme-mono' : 'scheme-prop',
          family: 'scheme',
          paragraphs: [para(autonum(scheme, startAt, face), { face })],
          width: 460,
          height: 44,
          wrap: false,
          asks: `what ${scheme} renders for ${String(startAt)} in ${face}`,
        });
      }
    }
  }
  return out;
}

/**
 * `startAt` outside `ST_TextBulletStartAtNum`, which is 1 to 32767.
 *
 * Each is alone in its own package. A value the schema forbids is a candidate
 * for the refusal that takes 800 other probes down with it, and C2's lesson was
 * that a hostile probe sharing a file with anything else measures the repair
 * rather than the question.
 */
export function hostileStartProbes(): Probe[] {
  return [0, -1, 32767, 32768, 65536].map((startAt) => ({
    id: `start-${startAt < 0 ? 'neg' : ''}${String(Math.abs(startAt))}`,
    deck: `hostile-start-${startAt < 0 ? 'neg' : ''}${String(Math.abs(startAt))}`,
    family: 'hostile-start',
    paragraphs: [para(autonum('arabicPeriod', startAt, FACE))],
    width: 460,
    height: 44,
    wrap: false,
    asks: `whether startAt=${String(startAt)} is accepted, and what it renders`,
  }));
}

/**
 * The alphabetic schemes, swept one number at a time.
 *
 * `STARTS` above is eighteen well-chosen values, and on eighteen values the
 * alphabetic schemes look like "repeat the (n mod A)th letter ceil(n/A) times" -
 * right at 27, 50, 100 and 400 and wrong at 900, 1000, 3999 and 4000, where the
 * repeat count collapses from 35 to 5 and from 154 to 4. Four anomalies is not a
 * rule and neither is the model they break; the only way to say what happens is
 * to walk n and look.
 *
 * `alphaLcPeriod` gets every value to 1300 because it is the cheapest alphabet
 * to read - 26 letters, no shaping, no fallback - and whatever the rule is, it
 * is visible there first. The other alphabets get every value to 64, which is
 * enough to read the alphabet itself and its first wrap, and then every
 * eleventh to 800.
 *
 * The alphabets are not the same size and are not the ones a list of the
 * script's letters would give: Thai turns out to use 41 of its 44 consonants,
 * skipping three, and that is only visible by walking.
 */
/**
 * A whole-slide-width frame for the count sweep.
 *
 * Not cosmetic. At n=611 the bullet is twenty-four `m`s, which is 480pt of
 * Arial at 24 - and in the right-hand column of a 960pt slide that runs off the
 * edge, taking the probe's own marker with it and silently welding two probes
 * into one reading. Six of thirteen hundred did exactly that on the first run.
 */
export const WIDE = 900;

/**
 * Eight points for the count sweep.
 *
 * At 24pt a forty-nine character Devanagari bullet is over a thousand points
 * wide, which runs past the slide edge and takes its own marker with it - and a
 * probe whose marker is missing welds itself onto the next one and reads as a
 * plausible string in the wrong alphabet. The EMF reports the characters
 * whatever the size, so the size is chosen to fit the longest string.
 */
export const COUNT_SIZE = 800;

export const ALPHABET_SCHEMES: readonly string[] = [
  'alphaLcPeriod',
  'alphaUcPeriod',
  'thaiAlphaPeriod',
  'hindiAlphaPeriod',
  'hindiAlpha1Period',
  'hebrew2Minus',
  'arabic1Minus',
  'arabic2Minus',
];

export function countProbes(): Probe[] {
  const out: Probe[] = [];
  const dense: number[] = [];
  for (let n = 1; n <= 1300; n++) dense.push(n);
  for (const n of dense) {
    out.push({
      id: `count-alphaLcPeriod-${String(n)}`,
      deck: 'count-latin',
      family: 'count',
      paragraphs: [para(autonum('alphaLcPeriod', n, FACE), { sz: COUNT_SIZE })],
      width: WIDE,
      height: 44,
      wrap: false,
      asks: `what alphaLcPeriod renders for ${String(n)}`,
    });
  }
  // A step of eleven locates a boundary to within eleven, which is enough to
  // read the alphabet size and not enough to settle an off-by-one - and Thai
  // has one: its transitions sit at 287, 328 and 369, which is 41k, while every
  // other alphabet's sit at 26k+1, 16k+1 and 22k+1. Whether that holds at k=1
  // as well as at k=7 decides between a rule and a special case, and only a
  // walk can say. So every value to 400 is probed, then the step resumes, and
  // 750..800 walks the wrap itself rather than fitting a modulus to samples on
  // either side of it.
  const coarse: number[] = [];
  for (let n = 1; n <= 400; n++) coarse.push(n);
  for (let n = 411; n <= 740; n += 11) coarse.push(n);
  for (let n = 750; n <= 800; n++) coarse.push(n);
  for (const scheme of ALPHABET_SCHEMES) {
    if (scheme === 'alphaLcPeriod') continue;
    for (const n of coarse) {
      out.push({
        id: `count-${scheme}-${String(n)}`,
        deck: 'count-other',
        family: 'count',
        paragraphs: [para(autonum(scheme, n, FACE), { sz: COUNT_SIZE })],
        width: WIDE,
        height: 44,
        wrap: false,
        asks: `what ${scheme} renders for ${String(n)}`,
      });
    }
  }
  return out;
}

/* -------------------------------------------------------------------------- */
/* F2 - which face draws the bullet                                           */
/* -------------------------------------------------------------------------- */

/**
 * `a:buFont` against `a:buFontTx` against neither, and the theme references.
 *
 * The run is always Arial, so a bullet drawn in the run's face and one drawn in
 * a stated face are different strings in the EMF's font record and different
 * advances in COM's. `+mj-lt` and `+mn-lt` are here because PowerPoint's own UI
 * writes `+mj-lt` when it converts a paragraph to a numbered list, which means
 * every deck a user makes that way depends on the theme indirection resolving.
 */
export function fontProbes(): Probe[] {
  const out: Probe[] = [];
  const fonts: readonly (string | undefined)[] = [
    'Arial',
    'Courier New',
    'Times New Roman',
    'Wingdings',
    '+mj-lt',
    '+mn-lt',
    'tx',
    undefined,
    'No Such Face At All',
  ];
  for (const font of fonts) {
    const label = font === undefined ? 'absent' : font.replace(/[^A-Za-z+-]/g, '');
    out.push({
      id: `font-num-${label}`,
      deck: 'font',
      family: 'font',
      paragraphs: [para({ kind: 'autonum', scheme: 'arabicPeriod', startAt: 8, font })],
      width: 460,
      height: 44,
      wrap: false,
      asks: `which face draws an autonumber bullet when buFont is ${label}`,
    });
    out.push({
      id: `font-char-${label}`,
      deck: 'font',
      family: 'font',
      paragraphs: [para({ kind: 'char', char: '•', font })],
      width: 460,
      height: 44,
      wrap: false,
      asks: `which face draws a character bullet when buFont is ${label}`,
    });
  }
  // The run in a face of its own, so `buFontTx` has something to point at that
  // is not the default. Without this the `tx` rows above cannot be told from
  // "PowerPoint ignored buFont", because both would draw in Arial.
  for (const runFace of ['Courier New', 'Times New Roman']) {
    out.push({
      id: `font-tx-run-${runFace.replace(/\s/g, '')}`,
      deck: 'font',
      family: 'font',
      paragraphs: [
        para(
          { kind: 'autonum', scheme: 'arabicPeriod', startAt: 8, font: 'tx' },
          { face: runFace },
        ),
      ],
      width: 460,
      height: 44,
      wrap: false,
      asks: `whether buFontTx follows a run in ${runFace}`,
    });
    // The same, with `buFont` naming a third face. The first run of probes
    // showed an autonumber bullet drawn in the run's face whatever `buFont`
    // said; this asks it where the two disagree *and* the run is not the
    // default, so "buFont was ignored" and "buFont happened to be Arial" are
    // different readings.
    out.push({
      id: `font-num-override-${runFace.replace(/\s/g, '')}`,
      deck: 'font',
      family: 'font',
      paragraphs: [
        para(
          { kind: 'autonum', scheme: 'arabicPeriod', startAt: 8, font: 'Wingdings' },
          {
            face: runFace,
          },
        ),
      ],
      width: 460,
      height: 44,
      wrap: false,
      asks: `which face draws an autonumber when buFont is Wingdings and the run is ${runFace}`,
    });
  }
  return out;
}

/* -------------------------------------------------------------------------- */
/* F3 - the numbering pass                                                    */
/* -------------------------------------------------------------------------- */

/**
 * What a reader has to count, since the file records nothing.
 *
 * Every probe is one shape of several paragraphs, and the EMF gives the number
 * each one was drawn with. The questions are all of the form "does this
 * interruption restart the sequence", and each interruption is asked on its own
 * so a disagreement names its own cause.
 *
 * `startAt` is stated on the *first* paragraph only unless a probe says
 * otherwise, because that is what PowerPoint's own files look like.
 */
export function sequenceProbes(): Probe[] {
  const out: Probe[] = [];
  const A = (startAt?: number, over: Partial<Bullet> = {}): Bullet => ({
    kind: 'autonum',
    scheme: 'arabicPeriod',
    font: FACE,
    ...(startAt === undefined ? {} : { startAt }),
    ...over,
  });

  const cases: readonly { id: string; asks: string; paras: readonly Para[] }[] = [
    {
      id: 'plain',
      asks: 'whether an unbroken run of paragraphs counts 1,2,3,4,5',
      paras: [A(), A(), A(), A(), A()].map((b) => para(b)),
    },
    {
      id: 'startat-first',
      asks: 'whether startAt on the first paragraph carries to the rest',
      paras: [A(7), A(), A(), A()].map((b) => para(b)),
    },
    {
      id: 'startat-every',
      asks: 'whether startAt repeated on every paragraph restarts every time',
      paras: [A(7), A(7), A(7), A(7)].map((b) => para(b)),
    },
    {
      id: 'startat-third',
      asks: 'whether a startAt part-way down restarts the sequence there',
      paras: [A(), A(), A(20), A()].map((b) => para(b)),
    },
    {
      id: 'buNone-middle',
      asks: 'whether an unnumbered paragraph consumes a number or is skipped',
      paras: [para(A()), para(A()), para({ kind: 'none' }), para(A()), para(A())],
    },
    {
      id: 'buChar-middle',
      asks: 'whether a character bullet interrupts the count',
      paras: [
        para(A()),
        para(A()),
        para({ kind: 'char', char: '•', font: FACE }),
        para(A()),
        para(A()),
      ],
    },
    {
      id: 'scheme-change',
      asks: 'whether changing the scheme part-way restarts the count',
      paras: [
        para(A()),
        para(A()),
        para(A(undefined, { scheme: 'alphaLcPeriod' })),
        para(A(undefined, { scheme: 'alphaLcPeriod' })),
        para(A()),
      ],
    },
    {
      id: 'level-out-and-back',
      asks: 'whether a nested level restarts, and whether the outer one resumes',
      paras: [
        para(A()),
        para(A()),
        para(A(), { level: 1, marL: 40, indent: 0 }),
        para(A(), { level: 1, marL: 40, indent: 0 }),
        para(A()),
        para(A()),
      ],
    },
    {
      id: 'level-twice',
      asks: 'whether a level restarts on every re-entry or resumes where it left off',
      paras: [
        para(A()),
        para(A(), { level: 1, marL: 40, indent: 0 }),
        para(A()),
        para(A(), { level: 1, marL: 40, indent: 0 }),
      ],
    },
    {
      id: 'level-deep',
      asks: 'whether three levels keep three independent counters',
      paras: [
        para(A()),
        para(A(), { level: 1, marL: 40, indent: 0 }),
        para(A(), { level: 2, marL: 80, indent: 0 }),
        para(A(), { level: 2, marL: 80, indent: 0 }),
        para(A(), { level: 1, marL: 40, indent: 0 }),
        para(A()),
      ],
    },
    {
      id: 'empty-para',
      asks: 'whether a paragraph with no runs still takes a number',
      paras: [para(A()), para(A(), { text: '' }), para(A())],
    },
    {
      id: 'hard-break',
      asks: 'whether a hard line break inside a paragraph takes a second number',
      paras: [para(A()), para(A(), { text: 'wwww', hardBreak: true }), para(A())],
    },
    {
      id: 'startat-zero-then',
      asks: 'whether a startAt of 1 mid-list restarts at 1',
      paras: [para(A()), para(A()), para(A(1)), para(A())],
    },
  ];

  // One package per probe, which is not caution but a finding already paid for
  // twice: C2 lost sixty-seven swatches to one out-of-range hue, C5 lost four
  // decks to one placeholder type, and the first run of this family came back
  // REPAIRED with six of fourteen shapes gone and no way to say which one did
  // it. A repaired deck measures the repair.
  for (const c of cases) {
    out.push({
      id: `seq-${c.id}`,
      deck: `sequence-${c.id}`,
      family: 'sequence',
      paragraphs: c.paras,
      width: 300,
      height: 32 * c.paras.length + 8,
      wrap: false,
      asks: c.asks,
    });
  }

  // Two shapes in one package, to confirm the counter is per text body. A
  // shared counter is a design a renderer could plausibly write and would be
  // wrong on every slide with two bulleted boxes.
  for (const half of ['a', 'b']) {
    out.push({
      id: `seq-second-body-${half}`,
      deck: 'sequence-second-body',
      family: 'sequence',
      paragraphs: [para(A()), para(A())],
      width: 300,
      height: 72,
      wrap: false,
      asks: 'whether a second text body in the same package restarts at one',
    });
  }
  return out;
}

/* -------------------------------------------------------------------------- */
/* F4 - character bullets and the private-use question                        */
/* -------------------------------------------------------------------------- */

/**
 * What PowerPoint draws for `a:buChar`, and in particular for the U+00xx
 * spelling of a symbol-font glyph.
 *
 * The plan says a `buChar char="§"` with `buFont typeface="Wingdings"` is a
 * Wingdings glyph rather than a section sign. The authoring step showed
 * PowerPoint writes whichever of the two it was given, so both spellings exist
 * in real files and the question is what each one *draws*. The EMF answers it
 * directly: the code point in the text record is the one GDI was asked for.
 */
export function charProbes(): Probe[] {
  const out: Probe[] = [];
  const cases: readonly { label: string; char: string; font?: string | undefined }[] = [
    { label: 'wd-pua-a7', char: '', font: 'Wingdings' },
    { label: 'wd-low-a7', char: '§', font: 'Wingdings' },
    { label: 'wd-pua-6c', char: '', font: 'Wingdings' },
    { label: 'wd-low-6c', char: 'l', font: 'Wingdings' },
    { label: 'wd-pua-a8', char: '', font: 'Wingdings' },
    { label: 'sym-pua-b7', char: '', font: 'Symbol' },
    { label: 'sym-low-b7', char: '·', font: 'Symbol' },
    { label: 'arial-bullet', char: '•', font: 'Arial' },
    { label: 'arial-pua-a7', char: '', font: 'Arial' },
    { label: 'arial-low-a7', char: '§', font: 'Arial' },
    { label: 'nofont-pua-a7', char: '', font: undefined },
    { label: 'nofont-bullet', char: '•', font: undefined },
    { label: 'wd-dash', char: '-', font: 'Wingdings' },
    { label: 'wd2-pua-a2', char: '', font: 'Wingdings 2' },
    { label: 'webdings-a4', char: '', font: 'Webdings' },
    { label: 'two-chars', char: 'ab', font: 'Arial' },
    { label: 'two-chars-xy', char: 'xy', font: 'Arial' },
    { label: 'two-chars-12', char: '12', font: 'Arial' },
    { label: 'three-chars', char: 'abc', font: 'Arial' },
    { label: 'astral', char: '\u{1F600}', font: 'Segoe UI Emoji' },
    { label: 'space', char: ' ', font: 'Arial' },
    { label: 'digit', char: '7', font: 'Arial' },
    { label: 'missing-glyph', char: 'ก', font: 'Courier New' },
    // Above U+00FF in a symbol face: if the rule is "low byte plus 0xF000" this
    // has no low byte to take, and what happens instead is the rule's edge.
    { label: 'wd-high-2022', char: '•', font: 'Wingdings' },
    { label: 'wd-latin-A', char: 'A', font: 'Wingdings' },
    { label: 'wd-pua-hi', char: '', font: 'Wingdings' },
    // A symbol face for a character it does have at the low code point, so
    // "mapped into the private-use area" and "left alone" differ.
    { label: 'sym-alpha', char: 'a', font: 'Symbol' },
  ];
  for (const c of cases) {
    out.push({
      id: `char-${c.label}`,
      deck: 'char',
      family: 'char',
      paragraphs: [para({ kind: 'char', char: c.char, font: c.font })],
      width: 460,
      height: 44,
      wrap: false,
      asks: `what buChar ${c.label} draws`,
    });
  }
  // An empty `@char` is legal against the schema and is alone in its package:
  // an empty required attribute is the shape of thing PowerPoint refuses.
  out.push({
    id: 'char-empty',
    deck: 'hostile-char-empty',
    family: 'hostile-char',
    paragraphs: [para({ kind: 'char', char: '', font: 'Arial' })],
    width: 460,
    height: 44,
    wrap: false,
    asks: 'whether an empty buChar/@char is accepted, and what it draws',
  });
  return out;
}

/* -------------------------------------------------------------------------- */
/* F5 - bullet size                                                           */
/* -------------------------------------------------------------------------- */

/**
 * `a:buSzPct`, `a:buSzPts` and `a:buSzTx`, and the size each is relative to.
 *
 * The last is the question a plausible implementation gets wrong. A paragraph
 * can state a size in three places - the run's `a:rPr/@sz`, the paragraph's
 * `a:defRPr/@sz`, and whatever the level inherits - and `buSzPct` is a
 * percentage of exactly one of them. Probes state two that disagree, so the
 * measured size names which.
 *
 * The EMF's `LOGFONTW.lfHeight` is the reading, in logical units proportional to
 * points, so a ratio between a probe and its control is exact even without
 * knowing the unit.
 */
export function sizeProbes(): Probe[] {
  const out: Probe[] = [];
  for (const szPct of [25000, 50000, 75000, 100000, 150000, 200000, 400000]) {
    out.push({
      id: `size-pct-${String(szPct)}`,
      deck: 'size',
      family: 'size',
      paragraphs: [para({ kind: 'char', char: '•', font: FACE, szPct })],
      width: 460,
      height: 60,
      wrap: false,
      asks: `what buSzPct ${String(szPct)} draws the bullet at`,
    });
  }
  for (const szPts of [600, 1200, 2400, 4800]) {
    out.push({
      id: `size-pts-${String(szPts)}`,
      deck: 'size',
      family: 'size',
      paragraphs: [para({ kind: 'char', char: '•', font: FACE, szPts })],
      width: 460,
      height: 60,
      wrap: false,
      asks: `what buSzPts ${String(szPts)} draws the bullet at`,
    });
  }
  out.push({
    id: 'size-tx',
    deck: 'size',
    family: 'size',
    paragraphs: [para({ kind: 'char', char: '•', font: FACE, szTx: true })],
    width: 460,
    height: 60,
    wrap: false,
    asks: 'what buSzTx draws the bullet at',
  });
  // Which size is 50% a percentage of: the run's 24pt or the paragraph's 12pt?
  out.push({
    id: 'size-pct-vs-defrpr',
    deck: 'size',
    family: 'size',
    paragraphs: [para({ kind: 'char', char: '•', font: FACE, szPct: 50000 }, { defSz: 1200 })],
    width: 460,
    height: 60,
    wrap: false,
    asks: 'whether buSzPct scales the run size or the paragraph defRPr size',
  });
  // Two runs of different sizes in one paragraph: which one does the bullet follow?
  // Two runs, 8pt then 40pt and the reverse. A bullet that follows "the
  // paragraph" or "the largest run" draws at 40 in both; one that follows the
  // first run draws at 8 in one and 40 in the other.
  out.push({
    id: 'size-two-runs-small-first',
    deck: 'size',
    family: 'size',
    paragraphs: [
      para(
        { kind: 'char', char: '•', font: FACE },
        { sz: 800, text: 'aa', secondRun: { text: 'bb', sz: 4000 } },
      ),
    ],
    width: 460,
    height: 70,
    wrap: false,
    asks: 'which run a bullet takes its size from when the first is the smaller',
  });
  out.push({
    id: 'size-two-runs-large-first',
    deck: 'size',
    family: 'size',
    paragraphs: [
      para(
        { kind: 'char', char: '•', font: FACE },
        { sz: 4000, text: 'aa', secondRun: { text: 'bb', sz: 800 } },
      ),
    ],
    width: 460,
    height: 70,
    wrap: false,
    asks: 'which run a bullet takes its size from when the first is the larger',
  });
  // The same question for the face, since the face is the one thing `buFont`
  // turned out not to control for an autonumber.
  out.push({
    id: 'size-two-runs-face',
    deck: 'size',
    family: 'size',
    paragraphs: [
      para(
        { kind: 'autonum', scheme: 'arabicPeriod', startAt: 8, font: 'Wingdings' },
        { face: MONO, text: 'aa', secondRun: { text: 'bb', sz: SIZE } },
      ),
    ],
    width: 460,
    height: 70,
    wrap: false,
    asks: 'which run an autonumber bullet takes its face from',
  });
  return out;
}

/** `buSzPct` outside `ST_TextBulletSizePercent` (25000..400000), each alone. */
export function hostileSizeProbes(): Probe[] {
  return [1000, 10000, 24999, 400001, 1000000].map((szPct) => ({
    id: `size-hostile-${String(szPct)}`,
    deck: `hostile-size-${String(szPct)}`,
    family: 'hostile-size',
    paragraphs: [para({ kind: 'char', char: '•', font: FACE, szPct })],
    width: 460,
    height: 120,
    wrap: false,
    asks: `whether buSzPct ${String(szPct)} is accepted, and what it clamps to`,
  }));
}

/* -------------------------------------------------------------------------- */
/* F6 - bullet colour                                                         */
/* -------------------------------------------------------------------------- */

/**
 * `a:buClr` and `a:buClrTx`, read out of `EMR_SETTEXTCOLOR`.
 *
 * The run is given a colour of its own in every probe, so "the bullet followed
 * the text" and "the bullet used the default" are different readings rather than
 * the same black.
 */
export function colourProbes(): Probe[] {
  const out: Probe[] = [];
  const runClr = '1F7A1F';
  const cases: readonly { label: string; bullet: Bullet }[] = [
    { label: 'srgb', bullet: { kind: 'char', char: '•', font: FACE, clr: 'C00000' } },
    { label: 'scheme', bullet: { kind: 'char', char: '•', font: FACE, clr: 'accent2' } },
    { label: 'tx', bullet: { kind: 'char', char: '•', font: FACE, clrTx: true } },
    { label: 'absent', bullet: { kind: 'char', char: '•', font: FACE } },
    // `buFont` turned out not to reach an autonumber bullet at all. Whether
    // `buClr` does is a separate question with a separate answer, and these are
    // the probes that make it one.
    {
      label: 'num-srgb',
      bullet: { kind: 'autonum', scheme: 'arabicPeriod', startAt: 8, font: FACE, clr: 'C00000' },
    },
    {
      label: 'num-scheme',
      bullet: { kind: 'autonum', scheme: 'arabicPeriod', startAt: 8, font: FACE, clr: 'accent2' },
    },
    { label: 'num-absent', bullet: autonum('arabicPeriod', 8, FACE) },
  ];
  for (const c of cases) {
    out.push({
      id: `clr-${c.label}`,
      deck: 'colour',
      family: 'colour',
      paragraphs: [para(c.bullet, { runClr })],
      width: 460,
      height: 44,
      wrap: false,
      asks: `what colour a bullet is drawn in when buClr is ${c.label}`,
    });
  }
  return out;
}

/* -------------------------------------------------------------------------- */
/* F7 - picture bullets                                                       */
/* -------------------------------------------------------------------------- */

/**
 * `a:buBlip`, which is the one bullet kind the EMF cannot answer: a picture is
 * not a text record. COM's `BoundLeft` carries the whole reading, so these
 * probes ask only about *size* - what box a picture bullet occupies - which is
 * the part a renderer has to get right for the text to start in the right place.
 *
 * Three aspect ratios, because a square source cannot tell "scaled to the line
 * box" from "scaled to the em" from "scaled to fit a square".
 */
export function blipProbes(): Probe[] {
  const out: Probe[] = [];
  // A second and third font size on the square source. Nine probes at one size
  // cannot tell "0.7 times the font size" from "16.8 points, always".
  for (const sz of [1200, 4800]) {
    out.push({
      id: `blip-size-${String(sz)}`,
      deck: 'blip',
      family: 'blip',
      paragraphs: [para({ kind: 'blip', blip: 'bullet-square' }, { sz })],
      width: 460,
      height: 70,
      wrap: false,
      asks: `what box a square picture bullet occupies at ${String(sz / 100)}pt`,
    });
  }
  for (const media of ['bullet-square', 'bullet-wide', 'bullet-tall']) {
    for (const szPct of [undefined, 50000, 200000]) {
      out.push({
        id: `blip-${media}-${szPct === undefined ? 'plain' : String(szPct)}`,
        deck: 'blip',
        family: 'blip',
        paragraphs: [para({ kind: 'blip', blip: media, szPct })],
        width: 460,
        height: 60,
        wrap: false,
        asks: `what box a ${media} picture bullet occupies at ${szPct === undefined ? 'default' : String(szPct)}`,
      });
    }
  }
  return out;
}

/* -------------------------------------------------------------------------- */
/* F8 - inheritance                                                           */
/* -------------------------------------------------------------------------- */

/**
 * Whether the bullet is one slot in the cascade or several.
 *
 * 3.1 measured that run and paragraph properties merge **per property**, and the
 * obvious extension is that `buChar`, `buFont`, `buSzPct` and `buClr` do too.
 * The obvious extension is not obviously right: `buNone`, `buChar`, `buAutoNum`
 * and `buBlip` are mutually exclusive in the schema, so *something* about the
 * bullet is a single slot, and whether the decorations travel with it or
 * separately is the question.
 *
 * The three readings that differ:
 *
 * - **per property**: a level declaring only `buFont` re-faces an inherited
 *   `buChar` and leaves everything else alone;
 * - **whole bullet**: any bullet element at a level replaces the whole bullet,
 *   so a level declaring only `buFont` loses the inherited character;
 * - **kind and decorations separate**: the kind is one slot and each decoration
 *   is its own, which is per-property with a rule for the exclusive group.
 *
 * Each probe declares one thing at the slide and something else at the master,
 * and the pair is chosen so the three readings disagree.
 */
export function cascadeProbes(): Probe[] {
  const out: Probe[] = [];
  const cases: readonly { id: string; slide: Bullet; asks: string }[] = [
    {
      id: 'font-only',
      slide: { kind: 'inherit', font: MONO },
      asks: 'whether a buFont alone re-faces an inherited buChar or loses it',
    },
    {
      id: 'size-only',
      slide: { kind: 'inherit', szPct: 50000 },
      asks: 'whether a buSzPct alone resizes an inherited buChar or loses it',
    },
    {
      id: 'clr-only',
      slide: { kind: 'inherit', clr: 'C00000' },
      asks: 'whether a buClr alone recolours an inherited buChar or loses it',
    },
    {
      id: 'none',
      slide: { kind: 'none' },
      asks: 'whether buNone suppresses an inherited buChar',
    },
    {
      id: 'autonum',
      slide: autonum('alphaLcPeriod', 3, MONO),
      asks: 'whether a buAutoNum replaces an inherited buChar',
    },
    {
      id: 'autonum-nofont',
      slide: { kind: 'autonum', scheme: 'alphaLcPeriod', startAt: 3 },
      asks: 'whether a buAutoNum with no buFont inherits the master buFont',
    },
    {
      id: 'char',
      slide: { kind: 'char', char: '●' },
      asks: 'whether a buChar with no buFont inherits the master buFont',
    },
    {
      id: 'clrtx',
      slide: { kind: 'inherit', clrTx: true },
      asks: 'whether buClrTx cancels a buClr inherited from the master',
    },
    {
      id: 'sztx',
      slide: { kind: 'inherit', szTx: true },
      asks: 'whether buSzTx cancels a buSzPct inherited from the master',
    },
    {
      id: 'fonttx',
      slide: { kind: 'inherit', font: 'tx' },
      asks: 'whether buFontTx cancels a buFont inherited from the master',
    },
    {
      id: 'inherit-all',
      slide: { kind: 'inherit' },
      asks: 'what a paragraph declaring no bullet element inherits',
    },
  ];
  // Each case twice, against two different outer levels. `cascade` puts the
  // inherited bullet in the shape's own `a:lstStyle`, which is one hop and
  // isolates the merge rule; `cascade-master` puts it in the master's
  // `p:bodyStyle` and binds the shape as a `body` placeholder, which is three
  // hops through the chain 3.1 measured. A rule that holds for one and not the
  // other would mean the bullet is not merged by the same machinery as the rest
  // of `a:pPr`, and that is worth finding out rather than assuming.
  for (const c of cases) {
    for (const deck of ['cascade', 'cascade-master']) {
      out.push({
        id: `casc-${deck === 'cascade' ? 'shape' : 'master'}-${c.id}`,
        deck,
        family: 'cascade',
        paragraphs: [para(c.slide, { runClr: '1F7A1F' })],
        width: 460,
        height: 44,
        wrap: false,
        asks: `${c.asks} (from the ${deck === 'cascade' ? 'shape lstStyle' : 'master bodyStyle'})`,
      });
    }
  }
  return out;
}

/**
 * The `a:lvl1pPr` the cascade probes inherit from, as master `p:bodyStyle`.
 *
 * One character bullet with all four decorations stated, so every probe above
 * has something distinguishable to keep or lose. The character is U+2756 - a
 * black diamond with a white X - because it exists in Wingdings and in no
 * text face, so "the master's font survived" and "the run's font was used" draw
 * different things rather than the same lozenge.
 */
export const CASCADE_MASTER_BULLET: Bullet = {
  kind: 'char',
  char: '',
  font: 'Wingdings',
  szPct: 150000,
  clr: '0070C0',
};

/* -------------------------------------------------------------------------- */
/* F9 - where the bullet goes                                                 */
/* -------------------------------------------------------------------------- */

/**
 * The indent arithmetic, which is the part of a bullet a fit test depends on.
 *
 * The authoring probe measured `text left = max(marL, bullet advance)` on four
 * numbers of one scheme, which is an anecdote. This sweeps `marL` and `indent`
 * across the regime boundary in both directions, with a bullet whose advance is
 * known from F1, and asks a second question the four numbers could not: whether
 * a *wrapped* line goes to `marL` regardless.
 */
export function indentProbes(): Probe[] {
  const out: Probe[] = [];
  // `arabicPeriod` at 888 in Courier New is five glyphs at 14.4pt = 72pt, and at
  // 8 it is two = 28.8pt. Sweeping marL past both says which term wins where.
  for (const startAt of [8, 888]) {
    for (const marL of [0, 12, 24, 28, 30, 36, 48, 72, 96]) {
      for (const indent of [0, -12, -24, -36, -72, 12]) {
        out.push({
          id: `ind-${String(startAt)}-${String(marL)}-${String(indent)}`,
          deck: 'indent',
          family: 'indent',
          paragraphs: [
            para(
              { kind: 'autonum', scheme: 'arabicPeriod', startAt, font: MONO },
              {
                face: MONO,
                marL,
                indent,
              },
            ),
          ],
          width: 460,
          height: 44,
          wrap: false,
          asks: `where the text starts at marL=${String(marL)} indent=${String(indent)}`,
        });
      }
    }
  }
  // A wrapped paragraph, so the second line's left edge is a separate reading.
  for (const marL of [0, 36, 96]) {
    for (const indent of [0, -36]) {
      out.push({
        id: `ind-wrap-${String(marL)}-${String(indent)}`,
        deck: 'indent-wrap',
        family: 'indent',
        paragraphs: [
          para(
            { kind: 'autonum', scheme: 'arabicPeriod', startAt: 8, font: MONO },
            {
              face: MONO,
              marL,
              indent,
              text: 'wwww wwww wwww wwww wwww wwww',
            },
          ),
        ],
        width: 220,
        height: 200,
        wrap: true,
        asks: `where a wrapped second line starts at marL=${String(marL)} indent=${String(indent)}`,
      });
    }
  }
  // Alignment, which decides whether the bullet is part of the line's width.
  for (const algn of ['l', 'ctr', 'r'] as const) {
    out.push({
      id: `ind-algn-${algn}`,
      deck: 'indent',
      family: 'indent',
      paragraphs: [
        para(
          { kind: 'autonum', scheme: 'arabicPeriod', startAt: 8, font: MONO },
          {
            face: MONO,
            algn,
            marL: 0,
            indent: 0,
            text: 'ww',
          },
        ),
      ],
      width: 300,
      height: 44,
      wrap: true,
      asks: `where a bullet goes when the paragraph is ${algn}-aligned`,
    });
  }
  // `defTabSz`, in case the gap after a bullet is a tab rather than an advance.
  for (const defTabSz of [36, 72, 144]) {
    out.push({
      id: `ind-tab-${String(defTabSz)}`,
      deck: 'indent',
      family: 'indent',
      paragraphs: [
        para(
          { kind: 'autonum', scheme: 'arabicPeriod', startAt: 8, font: MONO },
          {
            face: MONO,
            marL: 0,
            indent: 0,
            defTabSz,
          },
        ),
      ],
      width: 460,
      height: 44,
      wrap: false,
      asks: `whether defTabSz=${String(defTabSz)} moves the text after a bullet`,
    });
  }
  return out;
}

/* -------------------------------------------------------------------------- */
/* F10 - script runs                                                          */
/* -------------------------------------------------------------------------- */

/**
 * One character per Unicode block, and the four faces a run may name.
 *
 * The EMF makes this the cheapest family here rather than the most expensive.
 * GDI is asked for one face at a time, so a run whose characters resolve to
 * different faces comes out as *several* text records, each naming its face -
 * which means the split points and the face chosen for each side are both read
 * rather than inferred from an advance.
 *
 * Every probe declares all four of `a:latin`, `a:ea`, `a:cs` and `a:sym`, each a
 * face no other slot uses, so the record that comes back names the slot outright.
 */
export const SCRIPT_SAMPLE = [
  ['latin-basic', 'Az'],
  ['latin-1', 'Éñ'],
  ['latin-ext-a', 'Łš'],
  ['greek', 'Αβ'],
  ['cyrillic', 'Дж'],
  ['hebrew', 'את'],
  ['arabic', 'اب'],
  ['thai', 'กข'],
  ['devanagari', 'कख'],
  ['hiragana', 'あい'],
  ['katakana', 'アイ'],
  ['han', '一二'],
  ['hangul', '가나'],
  ['fullwidth', 'ＡＢ'],
  ['cjk-punct', '、。'],
  ['general-punct', '–—'],
  ['currency', '€£'],
  ['symbol-math', '∑∞'],
  ['dingbat', '✔❖'],
  ['digits', '01'],
  ['space-punct', '. '],
  ['pua', ''],
] as const;

/**
 * Faces chosen so that no two slots could be confused in an EMF font record,
 * and so that each covers the characters its slot is supposed to claim.
 */
export const SCRIPT_FACES = {
  latin: 'Georgia',
  ea: 'MS Gothic',
  cs: 'Courier New',
  sym: 'Wingdings',
} as const;

export function scriptProbes(): Probe[] {
  const out: Probe[] = [];
  for (const [label, text] of SCRIPT_SAMPLE) {
    out.push({
      id: `script-all-${label}`,
      deck: 'script',
      family: 'script',
      paragraphs: [
        para(
          { kind: 'none' },
          {
            text,
            face: SCRIPT_FACES.latin,
            ea: SCRIPT_FACES.ea,
            cs: SCRIPT_FACES.cs,
            sym: SCRIPT_FACES.sym,
            lang: 'en-US',
          },
        ),
      ],
      width: 460,
      height: 44,
      wrap: false,
      readChars: true,
      asks: `which of latin/ea/cs/sym draws ${label}`,
    });
  }
  // The same characters with one slot at a time left out, so "it used latin"
  // and "it fell back to latin because ea was absent" are different readings.
  const dropped: readonly (keyof typeof SCRIPT_FACES)[] = ['ea', 'cs', 'sym'];
  for (const drop of dropped) {
    for (const [label, text] of SCRIPT_SAMPLE) {
      const faces: Record<string, string | undefined> = { ...SCRIPT_FACES };
      faces[drop] = undefined;
      out.push({
        id: `script-no${drop}-${label}`,
        deck: 'script',
        family: 'script',
        paragraphs: [
          para(
            { kind: 'none' },
            {
              text,
              face: SCRIPT_FACES.latin,
              ea: faces['ea'],
              cs: faces['cs'],
              sym: faces['sym'],
              lang: 'en-US',
            },
          ),
        ],
        width: 460,
        height: 44,
        wrap: false,
        asks: `which face draws ${label} when ${drop} is absent`,
      });
    }
  }
  // `@lang`, which is the other thing a producer varies. If the split depends on
  // it, a deck authored in Japanese lays out differently from the same
  // characters authored in English, and a reader that ignores `@lang` is wrong.
  for (const lang of ['ja-JP', 'ar-SA', 'he-IL', 'th-TH', 'zh-CN']) {
    for (const [label, text] of SCRIPT_SAMPLE.slice(0, 14)) {
      out.push({
        id: `script-lang-${lang}-${label}`,
        deck: 'script-lang',
        family: 'script',
        paragraphs: [
          para(
            { kind: 'none' },
            {
              text,
              face: SCRIPT_FACES.latin,
              ea: SCRIPT_FACES.ea,
              cs: SCRIPT_FACES.cs,
              sym: SCRIPT_FACES.sym,
              lang,
            },
          ),
        ],
        width: 460,
        height: 44,
        wrap: false,
        asks: `whether lang=${lang} changes which face draws ${label}`,
      });
    }
  }
  return out;
}

/* -------------------------------------------------------------------------- */
/* F11 - fields                                                               */
/* -------------------------------------------------------------------------- */

/**
 * Every reserved `@type`, with a cache that is deliberately wrong.
 *
 * The cached text is `#stale#` in every probe - a string no date format can
 * produce - so a reading that comes back as the cache is unmistakable. The
 * authoring step already showed the cache is discarded for the types
 * PowerPoint's UI writes; these ask the same of `datetime`, which its UI never
 * writes, and of a type nothing reserves.
 *
 * `@lang` varies across sixteen locales because that is the whole of the plan's
 * claim for this family - `Intl.DateTimeFormat` keyed on `@lang` - and it is
 * only worth building if PowerPoint keys on it too.
 */
export const FIELD_LANGS: readonly string[] = [
  'en-US',
  'en-GB',
  'de-DE',
  'fr-FR',
  'ja-JP',
  'zh-CN',
  'ar-SA',
  'he-IL',
  'th-TH',
  'hi-IN',
  'ru-RU',
  'ko-KR',
  'es-ES',
  'pt-BR',
  'tr-TR',
  'vi-VN',
];

export function fieldProbes(): Probe[] {
  const out: Probe[] = [];
  for (const type of FIELD_TYPES) {
    out.push({
      id: `fld-${type}`,
      deck: 'field',
      family: 'field',
      paragraphs: [para({ kind: 'none' }, { field: [type, '#stale#'], lang: 'en-US', text: '' })],
      width: 460,
      height: 44,
      wrap: false,
      asks: `what ${type} renders, and whether the cache is used`,
    });
  }
  // Every reserved type in every locale, not a sample of five.
  //
  // The first run probed five types across sixteen locales and that was enough
  // to refute the plan's design - `Intl.DateTimeFormat` reproduces PowerPoint on
  // 25 of 97 - but not enough to replace it. What PowerPoint is really doing is
  // formatting with Windows' own per-locale patterns, and `datetime3` is
  // "5 September 2026" in one locale and "05/09/26" in the next, so the table
  // has to be measured cell by cell rather than derived from a semantic idea of
  // what "format 3" means.
  for (const lang of FIELD_LANGS) {
    for (const type of FIELD_TYPES) {
      if (type === 'slidenum') continue;
      out.push({
        id: `fld-lang-${lang}-${type}`,
        deck: 'field-lang',
        family: 'field',
        paragraphs: [para({ kind: 'none' }, { field: [type, '#stale#'], lang, text: '' })],
        width: 460,
        height: 44,
        wrap: false,
        asks: `what ${type} renders at lang=${lang}`,
      });
    }
  }
  // A type nothing reserves, and a field with no `@type` at all. Both are files
  // another producer can write, and the only sensible reading is the cache -
  // but that is a guess until it is measured.
  out.push({
    id: 'fld-unknown-type',
    deck: 'field',
    family: 'field',
    paragraphs: [
      para({ kind: 'none' }, { field: ['notAReservedType', 'CACHED'], lang: 'en-US', text: '' }),
    ],
    width: 460,
    height: 44,
    wrap: false,
    asks: 'whether an unreserved @type falls back to the cached text',
  });
  out.push({
    id: 'fld-no-type',
    deck: 'hostile-fld-no-type',
    family: 'hostile-field',
    paragraphs: [para({ kind: 'none' }, { field: ['', 'CACHED'], lang: 'en-US', text: '' })],
    width: 460,
    height: 44,
    wrap: false,
    asks: 'whether an a:fld with no @type is accepted',
  });
  return out;
}

/**
 * `slidenum`, on five slides, so "the cache" and "the slide's position" are
 * different numbers everywhere except by accident on slide one.
 */
export function slideNumProbes(): Probe[] {
  const out: Probe[] = [];
  for (let slide = 1; slide <= 5; slide++) {
    out.push({
      id: `fld-slidenum-${String(slide)}`,
      deck: 'field-slidenum',
      family: 'field',
      paragraphs: [
        para({ kind: 'none' }, { field: ['slidenum', '#stale#'], lang: 'en-US', text: '' }),
      ],
      width: 460,
      height: 44,
      wrap: false,
      asks: `what slidenum renders on slide ${String(slide)}`,
    });
  }
  return out;
}

/* -------------------------------------------------------------------------- */
/* everything                                                                 */
/* -------------------------------------------------------------------------- */

export function allProbes(): Probe[] {
  return [
    ...schemeProbes(),
    ...countProbes(),
    ...hostileStartProbes(),
    ...fontProbes(),
    ...sequenceProbes(),
    ...charProbes(),
    ...sizeProbes(),
    ...hostileSizeProbes(),
    ...colourProbes(),
    ...blipProbes(),
    ...cascadeProbes(),
    ...indentProbes(),
    ...scriptProbes(),
    ...fieldProbes(),
    ...slideNumProbes(),
  ];
}
