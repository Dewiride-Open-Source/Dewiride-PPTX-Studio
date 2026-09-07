/**
 * Every typeface a deck asks for, and what this machine will actually draw it in.
 *
 * The diagnostics panel reads this, and so does anything that needs to know a
 * measurement is approximate. ADR 0033.
 */

import { TextError } from '../errors.js';
import { createFontProbe, isFontAvailable, type FontProbe } from './presence.js';
import {
  fontStack,
  POWERPOINT_LAST_RESORT,
  substituteFor,
  verifySubstitute,
  type Substitute,
} from './substitute.js';

/**
 * `present` - the machine resolved the name to something of its own.
 * `substituted` - absent, and the table names what to use instead.
 * `missing` - absent, with nothing better than what PowerPoint itself would do.
 */
export type FontStatus = 'present' | 'substituted' | 'missing';

export interface FontFinding {
  /** The typeface as the deck names it, after any `+mj-lt` was resolved. */
  readonly typeface: string;
  readonly status: FontStatus;
  /** The family the text will be laid out in. */
  readonly rendersAs: string;
  /** The whole CSS family list, which is what a renderer sets. */
  readonly stack: string;
  /** `null` when the claim could not be checked because a face is absent. */
  readonly metricCompatible: boolean | null;
  /** Whether `metricCompatible` was measured here rather than taken from the table. */
  readonly verified: boolean;
  /** How many runs asked for this typeface. */
  readonly runs: number;
}

/** Case and whitespace are not part of a typeface's identity for a report. */
function key(family: string): string {
  return family.trim().replace(/\s+/g, ' ').toLowerCase();
}

function findingFor(
  typeface: string,
  runs: number,
  probe: FontProbe,
  available: (family: string) => boolean,
): FontFinding {
  const stack = fontStack(typeface);
  if (available(typeface)) {
    return {
      typeface,
      status: 'present',
      rendersAs: typeface,
      stack,
      metricCompatible: true,
      verified: true,
      runs,
    };
  }
  const substitute: Substitute | undefined = substituteFor(typeface);
  if (substitute !== undefined && available(substitute.use)) {
    const measured = verifySubstitute(substitute, probe, available);
    return {
      typeface,
      status: 'substituted',
      rendersAs: substitute.use,
      stack,
      metricCompatible: measured ?? substitute.metricCompatible,
      verified: measured !== null,
      runs,
    };
  }
  // Nothing here has it. PowerPoint would draw Calibri, so say so rather than
  // letting the browser choose its own default, which is a different face.
  return {
    typeface,
    status: 'missing',
    rendersAs: available(POWERPOINT_LAST_RESORT) ? POWERPOINT_LAST_RESORT : '',
    stack,
    metricCompatible: false,
    verified: true,
    runs,
  };
}

/**
 * One finding per distinct typeface, ordered by run count and then by name.
 *
 * `requested` is every typeface named by a run, repeats included: the count is
 * what makes a report about a hundred-slide deck readable.
 */
export function fontReport(
  requested: Iterable<string>,
  probe: FontProbe = createFontProbe(),
): readonly FontFinding[] {
  const counts = new Map<string, { typeface: string; runs: number }>();
  for (const raw of requested) {
    const typeface = raw.trim();
    if (typeface.length === 0) {
      throw new TextError('TEXT_FONT_FAMILY', 'a run named an empty typeface', raw);
    }
    const found = counts.get(key(typeface));
    if (found === undefined) counts.set(key(typeface), { typeface, runs: 1 });
    else found.runs += 1;
  }

  const cache = new Map<string, boolean>();
  const available = (family: string): boolean => {
    const cached = cache.get(key(family));
    if (cached !== undefined) return cached;
    const answer = isFontAvailable(family, probe);
    cache.set(key(family), answer);
    return answer;
  };

  return [...counts.values()]
    .map(({ typeface, runs }) => findingFor(typeface, runs, probe, available))
    .sort((a, b) => b.runs - a.runs || a.typeface.localeCompare(b.typeface));
}

/** The findings a user needs to be told about: everything not drawn as asked. */
export function substitutedFonts(findings: readonly FontFinding[]): readonly FontFinding[] {
  return findings.filter((f) => f.status !== 'present');
}
