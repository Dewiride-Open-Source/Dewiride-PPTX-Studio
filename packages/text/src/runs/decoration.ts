/**
 * The rules `a:rPr/@u` and `@strike` draw, and where they go.
 *
 * A renderer cannot delegate this to CSS. T8 measured Chromium's own
 * `text-decoration` against PowerPoint's and it agrees with none of eight faces
 * on either offset or thickness - the browser draws every face at 0.05em below
 * the baseline and 0.1em thick - and CSS has no way to place a strikethrough at
 * all. So both rules are drawn as geometry. ADR 0034.
 */

import { TextError } from '../errors.js';

import { MEASURED_FACE_RULES, type FaceRules, type RuleMetrics } from './face-rules.gen.js';

export { MEASURED_FACE_RULES, type FaceRules, type RuleMetrics };

/** `ST_TextUnderlineType`. */
export type Underline =
  | 'none'
  | 'words'
  | 'sng'
  | 'dbl'
  | 'heavy'
  | 'dotted'
  | 'dottedHeavy'
  | 'dash'
  | 'dashHeavy'
  | 'dashLong'
  | 'dashLongHeavy'
  | 'dotDash'
  | 'dotDashHeavy'
  | 'dotDotDash'
  | 'dotDotDashHeavy'
  | 'wavy'
  | 'wavyHeavy'
  | 'wavyDbl';

/** `ST_TextStrikeType`. */
export type Strike = 'noStrike' | 'sngStrike' | 'dblStrike';

/** How a rule is painted along its length. */
export type RulePattern = 'solid' | 'dotted' | 'dashed' | 'longDashed' | 'dotDash' | 'dotDotDash';

/** One drawn rule, in points relative to the baseline. */
export interface DrawnRule {
  /** The rule's top edge below the baseline; a strikethrough's is negative. */
  readonly top: number;
  readonly thickness: number;
  readonly pattern: RulePattern;
  /** Drawn as a wave rather than a straight rail. */
  readonly wavy: boolean;
}

/**
 * What to use for a typeface nobody measured.
 *
 * Arial's, exported under a name that says what it is so that reaching for it
 * is a decision at the call site rather than a default this module applies
 * behind the caller's back. The eight measured faces span 0.084 to 0.235 em, so
 * being wrong here is visible on a monospace face and invisible on most others.
 */
export const APPROXIMATE_FACE_RULES: FaceRules = {
  underline: { offset: 0.103, thickness: 0.075 },
  strike: { offset: -0.2595, thickness: 0.05 },
};

/** The rules measured for a typeface, or an error naming what is missing. */
export function faceRules(typeface: string): FaceRules {
  const rules = MEASURED_FACE_RULES[typeface];
  if (rules === undefined) {
    throw new TextError(
      'TEXT_FACE_METRICS',
      `no measured underline or strike metrics for ${typeface}; pass APPROXIMATE_FACE_RULES deliberately if that is what you want`,
      typeface,
    );
  }
  return rules;
}

/** Whether this typeface was one of the eight measured. */
export function hasFaceRules(typeface: string): boolean {
  return Object.hasOwn(MEASURED_FACE_RULES, typeface);
}

const PATTERNS: Readonly<Record<string, RulePattern>> = {
  dotted: 'dotted',
  dottedHeavy: 'dotted',
  dash: 'dashed',
  dashHeavy: 'dashed',
  dashLong: 'longDashed',
  dashLongHeavy: 'longDashed',
  dotDash: 'dotDash',
  dotDashHeavy: 'dotDash',
  dotDotDash: 'dotDotDash',
  dotDotDashHeavy: 'dotDotDash',
};

const HEAVY = new Set([
  'heavy',
  'dottedHeavy',
  'dashHeavy',
  'dashLongHeavy',
  'dotDashHeavy',
  'dotDotDashHeavy',
  'wavyHeavy',
]);

const WAVY = new Set(['wavy', 'wavyHeavy', 'wavyDbl']);

const DOUBLE = new Set(['dbl', 'wavyDbl']);

function checkSize(sizePt: number): number {
  if (!Number.isFinite(sizePt) || sizePt <= 0) {
    throw new TextError(
      'TEXT_SIZE',
      `font size ${String(sizePt)} is not a positive size`,
      String(sizePt),
    );
  }
  return sizePt;
}

/**
 * The rules an `@u` value draws, in points relative to the baseline.
 *
 * Three derived shapes, each holding at 18, 40 and 66 points: a heavy rule is
 * half again as thick and starts one plain thickness higher, and a double rule
 * is two half-thickness rails at the same two edges.
 */
export function underlineRules(
  underline: Underline | undefined,
  sizePt: number,
  rules: FaceRules,
): readonly DrawnRule[] {
  const kind = underline ?? 'none';
  if (kind === 'none') return [];
  checkSize(sizePt);
  const top = rules.underline.offset * sizePt;
  const thickness = rules.underline.thickness * sizePt;
  const pattern = PATTERNS[kind] ?? 'solid';
  const wavy = WAVY.has(kind);

  if (DOUBLE.has(kind)) {
    return [
      { top: top - thickness, thickness: thickness / 2, pattern, wavy },
      { top: top + thickness / 2, thickness: thickness / 2, pattern, wavy },
    ];
  }
  if (HEAVY.has(kind)) {
    return [{ top: top - thickness, thickness: thickness * 1.5, pattern, wavy }];
  }
  return [{ top, thickness, pattern, wavy }];
}

/**
 * The rules a `@strike` value draws.
 *
 * `dblStrike` is two rails of the plain thickness, one thickness either side of
 * where the single one goes - exact at all three measured sizes.
 */
export function strikeRules(
  strike: Strike | undefined,
  sizePt: number,
  rules: FaceRules,
): readonly DrawnRule[] {
  const kind = strike ?? 'noStrike';
  if (kind === 'noStrike') return [];
  checkSize(sizePt);
  const top = rules.strike.offset * sizePt;
  const thickness = rules.strike.thickness * sizePt;
  if (kind === 'dblStrike') {
    return [
      { top: top - thickness, thickness, pattern: 'solid', wavy: false },
      { top: top + thickness, thickness, pattern: 'solid', wavy: false },
    ];
  }
  if (kind !== 'sngStrike') {
    throw new TextError(
      'TEXT_DECORATION',
      `${String(strike)} is not an ST_TextStrikeType`,
      String(strike),
    );
  }
  return [{ top, thickness, pattern: 'solid', wavy: false }];
}

/**
 * The text a rule is drawn under, which is not always the run's text.
 *
 * Measured on three probes: the rule stops at the last glyph, so the spaces a
 * line ends with are not underlined. `u="words"` skips every space, not only
 * the trailing ones, which is why this returns a list of stretches.
 */
export function decoratedStretches(
  text: string,
  underline: Underline | undefined,
): readonly (readonly [number, number])[] {
  if (underline === undefined || underline === 'none') return [];
  if (underline === 'words') {
    const out: [number, number][] = [];
    let start: number | null = null;
    for (let i = 0; i < text.length; i += 1) {
      const isSpace = text[i] === ' ';
      if (!isSpace && start === null) start = i;
      if (isSpace && start !== null) {
        out.push([start, i]);
        start = null;
      }
    }
    if (start !== null) out.push([start, text.length]);
    return out;
  }
  const end = text.replace(/ +$/, '').length;
  return end === 0 ? [] : [[0, end]];
}
