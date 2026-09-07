/**
 * Experiment T8 - what a renderer has to know to draw text where PowerPoint
 * draws it.
 *
 * Sub-phases 3.1 to 3.7 settled where a *line* goes: the cascade, the 1.2 line
 * box, breaking, autofit, bullets, the frame and the font. None of them asked
 * where the glyphs sit inside that line, which way a mirrored shape's text
 * faces, or what an underline is a fraction of. Those are the drawing
 * questions, and this is where they get asked.
 *
 * Seven questions, each with candidate readings written down before the deck
 * was built, including the one a renderer written by hand would have used.
 */

/** A rectangle in points, on the 960 x 540 slide. */
export interface Rect {
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
}

/** One `a:r`, as much of `a:rPr` as a probe states. */
export interface Run {
  readonly text: string;
  readonly typeface?: string;
  /** `a:rPr/@sz` in points; the deck writes hundredths. */
  readonly szPt?: number;
  readonly b?: boolean;
  readonly i?: boolean;
  /** `ST_TextUnderlineType`. */
  readonly u?: string;
  /** `ST_TextStrikeType`. */
  readonly strike?: string;
  /** `ST_TextCapsType`. */
  readonly cap?: string;
  /** `@baseline`, thousandths of a percent: 30000 is a third of an em up. */
  readonly baseline?: number;
  /** `@spc` in points. */
  readonly spcPt?: number;
  /** `@kern` in points; `0` is GDI for never. */
  readonly kernPt?: number;
  /** `a:solidFill` on the run, as six hex digits. */
  readonly color?: string;
}

/** One `a:p`. */
export interface Para {
  readonly runs: readonly Run[];
  /** `ST_TextAlignType`. */
  readonly algn?: string;
}

export interface Probe {
  readonly id: string;
  readonly question: string;
  readonly rect: Rect;
  /** `a:xfrm/@rot` in degrees; the deck writes sixtieths. */
  readonly rotDeg?: number;
  readonly flipH?: boolean;
  readonly flipV?: boolean;
  /** `a:bodyPr/@rot` in degrees. */
  readonly bodyRotDeg?: number;
  readonly upright?: boolean;
  readonly anchor?: string;
  readonly wrap?: 'none' | 'square';
  /** All four insets at once, in points. */
  readonly insetPt?: number;
  readonly paragraphs: readonly Para[];
  /** Export this slide as a bitmap too: the question needs ink, not records. */
  readonly bitmap?: boolean;
  readonly asks: string;
}

/** Every probe deck names one question, so a repair localises to one. */
export interface ProbePackage {
  readonly deck: string;
  readonly question: string;
  readonly probes: readonly Probe[];
}

/* -------------------------------------------------------------------------- */
/* the shared frame                                                           */
/* -------------------------------------------------------------------------- */

/**
 * Whole points, so every edge lands on a pixel boundary at 1920 x 1080.
 *
 * The slide is 960 x 540 pt, so one point is exactly two pixels there.
 */
export const FRAME: Rect = { x: 120, y: 120, w: 600, h: 240 };

/**
 * The frame the two turn questions use: small enough to survive any rotation.
 *
 * Centred on the slide with a half-diagonal of 215pt, so no angle puts a corner
 * off the 960 x 540 page - where the first run of this experiment lost four
 * probes to clipping and read them as "PowerPoint drew nothing".
 */
export const TURN_FRAME: Rect = { x: 280, y: 190, w: 400, h: 160 };

/** A string with no descender below the baseline, so ink bottom is the baseline. */
export const FLAT_TEXT = 'xxxxxxxx';

/** The six faces `MEASURED_FACE_METRICS` covers, plus the two Office defaults. */
export const FACES: readonly string[] = [
  'Arial',
  'Times New Roman',
  'Verdana',
  'Courier New',
  'Georgia',
  'Tahoma',
  'Calibri',
  'Consolas',
];

/** Sizes far enough apart that a constant and a proportion cannot both fit. */
export const SIZES: readonly number[] = [12, 18, 32, 54];

const PLAIN = { anchor: 't', wrap: 'square', insetPt: 0 } as const;

function para(text: string, run: Omit<Run, 'text'> = {}, algn?: string): Para {
  return { runs: [{ text, ...run }], ...(algn === undefined ? {} : { algn }) };
}

/* -------------------------------------------------------------------------- */
/* Q1 - does a flipped shape mirror its text?                                  */
/* -------------------------------------------------------------------------- */

/**
 * The plan asserts PowerPoint does not mirror text on a flipped shape.
 *
 * Three readings fit that sentence and they disagree about where the line goes:
 * the glyphs could be mirrored with the block, counter-mirrored in place, or the
 * whole text layer drawn as if the shape had never been flipped.
 */
export function flipProbes(): readonly Probe[] {
  const out: Probe[] = [];
  for (const rot of [0, 30, 90, 180, 270]) {
    for (const [h, v] of [
      [false, false],
      [true, false],
      [false, true],
      [true, true],
    ] as const) {
      const axes = `${h ? 'h' : ''}${v ? 'v' : ''}` || 'none';
      out.push({
        id: `flip-r${String(rot)}-${axes}`,
        question: 'flip',
        rect: TURN_FRAME,
        rotDeg: rot,
        ...(h ? { flipH: true } : {}),
        ...(v ? { flipV: true } : {}),
        ...PLAIN,
        paragraphs: [para('Rb', { typeface: 'Arial', szPt: 40 }, 'l')],
        bitmap: true,
        asks: 'whether the glyphs and the block follow a flip',
      });
    }
  }
  return out;
}

/* -------------------------------------------------------------------------- */
/* Q2 - text and rotation                                                      */
/* -------------------------------------------------------------------------- */

/** `@upright` and `a:bodyPr/@rot` against the shape's own `a:xfrm/@rot`. */
export function rotationProbes(): readonly Probe[] {
  const out: Probe[] = [];
  for (const rot of [0, 30, 90, 180, 270]) {
    for (const upright of [undefined, false, true]) {
      out.push({
        id: `rot-${String(rot)}-${upright === undefined ? 'unset' : upright ? 'up' : 'down'}`,
        question: 'rotation',
        rect: TURN_FRAME,
        rotDeg: rot,
        ...(upright === undefined ? {} : { upright }),
        ...PLAIN,
        paragraphs: [para('Rotate', { typeface: 'Arial', szPt: 32 }, 'l')],
        bitmap: true,
        asks: 'the angle the glyphs are drawn at',
      });
    }
  }
  for (const bodyRot of [30, 90, -30]) {
    out.push({
      id: `rot-body-${bodyRot < 0 ? 'neg' : ''}${String(Math.abs(bodyRot))}`,
      question: 'rotation',
      rect: TURN_FRAME,
      bodyRotDeg: bodyRot,
      ...PLAIN,
      paragraphs: [para('Rotate', { typeface: 'Arial', szPt: 32 }, 'l')],
      bitmap: true,
      asks: 'whether a:bodyPr/@rot turns the text on an unrotated shape',
    });
  }
  out.push({
    id: 'rot-both-30-45',
    question: 'rotation',
    rect: TURN_FRAME,
    rotDeg: 30,
    bodyRotDeg: 45,
    ...PLAIN,
    paragraphs: [para('Rotate', { typeface: 'Arial', szPt: 32 }, 'l')],
    bitmap: true,
    asks: 'whether the two rotations add',
  });
  return out;
}

/* -------------------------------------------------------------------------- */
/* Q3 - where the baseline sits inside the line box                            */
/* -------------------------------------------------------------------------- */

/**
 * The one number a renderer cannot guess.
 *
 * 3.2 measured the line box at 1.2 x the size for thirteen faces; nothing has
 * asked where in that box the baseline is, and every candidate scored against
 * this is what some real implementation uses.
 */
export function baselineProbes(): readonly Probe[] {
  const out: Probe[] = [];
  for (const face of FACES) {
    for (const size of SIZES) {
      out.push({
        id: `base-${face.replace(/\s+/g, '')}-${String(size)}`,
        question: 'baseline',
        rect: FRAME,
        ...PLAIN,
        paragraphs: [para(FLAT_TEXT, { typeface: face, szPt: size }, 'l')],
        bitmap: size === 32,
        asks: 'the baseline offset from the top of the line box',
      });
    }
  }
  // Two lines of different heights, which is the only thing that tells a
  // line's own box from the box of the line after it.
  out.push({
    id: 'base-two-sizes',
    question: 'baseline',
    rect: FRAME,
    ...PLAIN,
    paragraphs: [
      { runs: [{ text: FLAT_TEXT, typeface: 'Arial', szPt: 18 }] },
      { runs: [{ text: FLAT_TEXT, typeface: 'Arial', szPt: 44 }] },
    ],
    bitmap: true,
    asks: 'whether a line advance is this line box or the next one',
  });
  out.push({
    id: 'base-two-lines',
    question: 'baseline',
    rect: FRAME,
    ...PLAIN,
    paragraphs: [
      { runs: [{ text: FLAT_TEXT, typeface: 'Arial', szPt: 32 }] },
      { runs: [{ text: FLAT_TEXT, typeface: 'Arial', szPt: 32 }] },
    ],
    bitmap: true,
    asks: 'whether the second baseline is one line advance below the first',
  });
  // Three orders and three size pairs, because one mixed line cannot tell "the
  // largest run sets the box" from "the first run does".
  for (const [id, sizes] of [
    ['small-large', [18, 44]],
    ['large-small', [44, 18]],
    ['small-large-small', [12, 36, 12]],
    ['large-small-large', [30, 14, 30]],
  ] as const) {
    out.push({
      id: `base-mixed-${id}`,
      question: 'baseline',
      rect: FRAME,
      ...PLAIN,
      paragraphs: [{ runs: sizes.map((szPt) => ({ text: 'xxxx', typeface: 'Arial', szPt })) }],
      bitmap: true,
      asks: 'the line box and the baseline of a line holding several sizes',
    });
  }
  return out;
}

/* -------------------------------------------------------------------------- */
/* Q4 - where a line starts                                                    */
/* -------------------------------------------------------------------------- */

const WRAPPED = 'Alpha bravo charlie delta echo foxtrot golf hotel india juliet kilo';

/** `@algn`, and whether a trailing space counts towards the line's width. */
export function alignProbes(): readonly Probe[] {
  const out: Probe[] = [];
  for (const algn of ['l', 'ctr', 'r', 'just', 'justLow', 'dist', 'thaiDist']) {
    out.push({
      id: `align-one-${algn}`,
      question: 'align',
      rect: FRAME,
      ...PLAIN,
      paragraphs: [para('Alpha bravo', { typeface: 'Arial', szPt: 24 }, algn)],
      asks: 'where a single short line starts',
    });
    out.push({
      id: `align-wrap-${algn}`,
      question: 'align',
      rect: FRAME,
      ...PLAIN,
      paragraphs: [para(WRAPPED, { typeface: 'Arial', szPt: 24 }, algn)],
      asks: 'where each line of a wrapped paragraph starts, last line included',
    });
  }
  for (const [id, text] of [
    ['bare', 'Alpha bravo'],
    ['one-space', 'Alpha bravo '],
    ['three-spaces', 'Alpha bravo   '],
  ] as const) {
    for (const algn of ['ctr', 'r']) {
      out.push({
        id: `align-trail-${algn}-${id}`,
        question: 'align',
        rect: FRAME,
        ...PLAIN,
        paragraphs: [para(text, { typeface: 'Arial', szPt: 24 }, algn)],
        asks: 'whether trailing spaces move a centred or right-aligned line',
      });
    }
  }
  return out;
}

/* -------------------------------------------------------------------------- */
/* Q5 - underline and strikethrough                                            */
/* -------------------------------------------------------------------------- */

/** Whether GDI draws the decoration from the face, and where it lands. */
export function decorationProbes(): readonly Probe[] {
  const out: Probe[] = [];
  for (const u of ['sng', 'dbl', 'heavy', 'dotted', 'wavy', 'words', 'dash']) {
    for (const size of [18, 40, 66]) {
      out.push({
        id: `dec-u-${u}-${String(size)}`,
        question: 'decoration',
        rect: FRAME,
        ...PLAIN,
        paragraphs: [para('Alpha bravo', { typeface: 'Arial', szPt: size, u }, 'l')],
        bitmap: true,
        asks: 'who draws the underline and where it sits',
      });
    }
  }
  for (const strike of ['sngStrike', 'dblStrike']) {
    for (const size of [18, 40, 66]) {
      out.push({
        id: `dec-s-${strike}-${String(size)}`,
        question: 'decoration',
        rect: FRAME,
        ...PLAIN,
        paragraphs: [para('Alpha bravo', { typeface: 'Arial', szPt: size, strike }, 'l')],
        bitmap: true,
        asks: 'who draws the strikethrough and where it sits',
      });
    }
  }
  // 120pt, because two faces' underline offsets can differ by six thousandths
  // of an em and a 1920-pixel export resolves half a point.
  for (const face of FACES) {
    for (const kind of ['u', 's'] as const) {
      out.push({
        id: `dec-face-${kind}-${face.replace(/\s+/g, '')}`,
        question: 'decoration',
        rect: FRAME,
        ...PLAIN,
        paragraphs: [
          para(
            'Alpha',
            {
              typeface: face,
              szPt: 120,
              ...(kind === 'u' ? { u: 'sng' } : { strike: 'sngStrike' }),
            },
            'l',
          ),
        ],
        bitmap: true,
        asks: 'whether the rule offset is a property of the face',
      });
    }
  }
  for (const [id, text, face] of [
    ['one', 'Alpha ', 'Arial'],
    ['three', 'Alpha   ', 'Arial'],
    ['georgia', 'Alpha ', 'Georgia'],
  ] as const) {
    out.push({
      id: `dec-trail-${id}`,
      question: 'decoration',
      rect: FRAME,
      ...PLAIN,
      paragraphs: [para(text, { typeface: face, szPt: 40, u: 'sng' }, 'l')],
      bitmap: true,
      asks: 'whether the rule runs under a trailing space',
    });
    out.push({
      id: `dec-trail-${id}-bare`,
      question: 'decoration',
      rect: FRAME,
      ...PLAIN,
      paragraphs: [para(text.trimEnd(), { typeface: face, szPt: 40, u: 'sng' }, 'l')],
      bitmap: true,
      asks: 'the same rule with the trailing spaces taken out of the file',
    });
  }
  return out;
}

/* -------------------------------------------------------------------------- */
/* Q6 - baseline shift and capitalisation                                      */
/* -------------------------------------------------------------------------- */

/** `@baseline` and `@cap`: both change the glyph, and one may change the size. */
export function shiftProbes(): readonly Probe[] {
  const out: Probe[] = [];
  for (const baseline of [30000, 50000, -25000, 100000]) {
    out.push({
      id: `shift-${baseline < 0 ? 'down' : 'up'}${String(Math.abs(baseline))}`,
      question: 'shift',
      rect: FRAME,
      ...PLAIN,
      paragraphs: [
        {
          runs: [
            { text: 'xx', typeface: 'Arial', szPt: 32 },
            { text: 'xx', typeface: 'Arial', szPt: 32, baseline },
          ],
        },
      ],
      bitmap: true,
      asks: 'how far a shifted run moves and at what size it is drawn',
    });
  }
  for (const cap of ['none', 'small', 'all']) {
    out.push({
      id: `shift-cap-${cap}`,
      question: 'shift',
      rect: FRAME,
      ...PLAIN,
      paragraphs: [para('Hamburg fox', { typeface: 'Arial', szPt: 32, cap }, 'l')],
      bitmap: true,
      asks: 'what a:rPr/@cap draws and at what size',
    });
  }
  return out;
}

/* -------------------------------------------------------------------------- */
/* Q7 - what a run boundary does to the advances                               */
/* -------------------------------------------------------------------------- */

/**
 * A renderer emitting one span per run has to know whether that splits a pair.
 *
 * `AVA` kerns; `AV` + `A` as two runs is the same string with a boundary in the
 * middle of the pair, and the two readings differ by exactly the kern.
 */
export function runSplitProbes(): readonly Probe[] {
  const out: Probe[] = [];
  const font = { typeface: 'Arial', szPt: 40 } as const;
  for (const kernPt of [undefined, 0] as const) {
    const k = kernPt === undefined ? 'unset' : 'off';
    const kern = kernPt === undefined ? {} : { kernPt };
    out.push({
      id: `runs-whole-${k}`,
      question: 'runs',
      rect: FRAME,
      ...PLAIN,
      paragraphs: [{ runs: [{ text: 'AVAWA', ...font, ...kern }] }],
      asks: 'the advances of a kerning string in one run',
    });
    out.push({
      id: `runs-split-${k}`,
      question: 'runs',
      rect: FRAME,
      ...PLAIN,
      paragraphs: [
        {
          runs: [
            { text: 'AV', ...font, ...kern },
            { text: 'AWA', ...font, ...kern },
          ],
        },
      ],
      asks: 'the same string with a run boundary inside a kerning pair',
    });
    out.push({
      id: `runs-colored-${k}`,
      question: 'runs',
      rect: FRAME,
      ...PLAIN,
      paragraphs: [
        {
          runs: [
            { text: 'AV', ...font, color: 'C00000', ...kern },
            { text: 'AWA', ...font, color: '0070C0', ...kern },
          ],
        },
      ],
      asks: 'whether a colour change alone splits the drawing call',
    });
  }
  out.push({
    id: 'runs-spc',
    question: 'runs',
    rect: FRAME,
    ...PLAIN,
    paragraphs: [{ runs: [{ text: 'AVAWA', typeface: 'Arial', szPt: 40, spcPt: 4 }] }],
    asks: 'whether @spc is added after the last character of a line',
  });
  return out;
}

/* -------------------------------------------------------------------------- */
/* the packages                                                               */
/* -------------------------------------------------------------------------- */

/** One package per question, so a repair names one question and not sixty. */
export function allPackages(): readonly ProbePackage[] {
  return [
    { deck: 'flip', question: 'flip', probes: flipProbes() },
    { deck: 'rotation', question: 'rotation', probes: rotationProbes() },
    { deck: 'baseline', question: 'baseline', probes: baselineProbes() },
    { deck: 'align', question: 'align', probes: alignProbes() },
    { deck: 'decoration', question: 'decoration', probes: decorationProbes() },
    { deck: 'shift', question: 'shift', probes: shiftProbes() },
    { deck: 'runs', question: 'runs', probes: runSplitProbes() },
  ];
}

/** Every probe in every package, in deck order. */
export function allProbes(): readonly Probe[] {
  return allPackages().flatMap((pkg) => pkg.probes);
}
