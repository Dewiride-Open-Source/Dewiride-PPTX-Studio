/**
 * What to render with when the deck names a typeface this machine lacks.
 *
 * The table is data, and `metricCompatible` is a claim its designer makes, not
 * a measurement - `metricsAgree` is what settles it at run time. ADR 0033.
 */

import { TextError } from '../errors.js';
import { cssFamily } from '../runs/measure.js';
import { metricsAgree, type FontProbe } from './presence.js';

/**
 * What PowerPoint draws a Latin face it does not have in.
 *
 * Measured, 22 of 22, against a name-derived reading at 17 and the theme's own
 * minor face at 20. Chromium's last resort is its `serif`, which is a different
 * face, so a stack that does not say this renders the deck in the wrong one.
 */
export const POWERPOINT_LAST_RESORT = 'Calibri';

/** Chased in order when `POWERPOINT_LAST_RESORT` is itself absent. */
export const LAST_RESORT_FAMILIES: readonly string[] = ['Calibri', 'Carlito'];

/** The one generic every stack ends on, so a stack is never open-ended. */
export const LAST_RESORT_GENERIC = 'sans-serif';

export interface Substitute {
  /** The typeface a deck names. */
  readonly missing: string;
  /** What to render with instead. */
  readonly use: string;
  /** Whether `use` was designed to the same advances as `missing`. */
  readonly metricCompatible: boolean;
  readonly note?: string;
}

/**
 * The substitutions, most specific first.
 *
 * The metric-compatible entries are the open clones sub-phase 8.8 bundles; the
 * rest are what Microsoft itself falls back to.
 */
export const SUBSTITUTES: readonly Substitute[] = [
  { missing: 'Arial', use: 'Liberation Sans', metricCompatible: true },
  { missing: 'Arial Narrow', use: 'Liberation Sans Narrow', metricCompatible: true },
  { missing: 'Helvetica', use: 'Arial', metricCompatible: true },
  { missing: 'Times New Roman', use: 'Liberation Serif', metricCompatible: true },
  { missing: 'Times', use: 'Times New Roman', metricCompatible: true },
  { missing: 'Courier New', use: 'Liberation Mono', metricCompatible: true },
  { missing: 'Courier', use: 'Courier New', metricCompatible: true },
  { missing: 'Calibri', use: 'Carlito', metricCompatible: true },
  { missing: 'Cambria', use: 'Caladea', metricCompatible: true },
  { missing: 'Georgia', use: 'Gelasio', metricCompatible: true },
  { missing: 'Segoe UI', use: 'Selawik', metricCompatible: true },
  {
    missing: 'Aptos',
    use: 'Carlito',
    metricCompatible: false,
    note: "Microsoft's own fallback is Calibri, whose open clone this is; no clone of Aptos exists.",
  },
  { missing: 'Aptos Display', use: 'Carlito', metricCompatible: false },
  { missing: 'Aptos Narrow', use: 'Carlito', metricCompatible: false },
];

/** Case and whitespace are not part of a typeface's identity for a lookup. */
function key(family: string): string {
  return family.trim().replace(/\s+/g, ' ').toLowerCase();
}

const BY_NAME = new Map<string, Substitute>(SUBSTITUTES.map((s) => [key(s.missing), s]));

export function substituteFor(family: string): Substitute | undefined {
  if (family.length === 0) {
    throw new TextError('TEXT_FONT_FAMILY', 'typeface name is empty', family);
  }
  return BY_NAME.get(key(family));
}

/**
 * The CSS family list a run is set in.
 *
 * The requested face first, so a machine that has it wins; then the substitute;
 * then what PowerPoint would have used, so a face nobody has lands where the
 * deck's author saw it rather than on the browser's own default.
 */
export function fontStack(family: string): string {
  const substitute = substituteFor(family);
  const families = [family];
  const add = (name: string): void => {
    if (!families.includes(name)) families.push(name);
  };
  if (substitute !== undefined) add(substitute.use);
  for (const name of LAST_RESORT_FAMILIES) add(name);
  return [...families.map(cssFamily), LAST_RESORT_GENERIC].join(', ');
}

/**
 * Whether a substitution's `metricCompatible` claim holds on this machine.
 *
 * `null` when either face is absent, which is the usual case and is not a
 * failure: an unverified claim must be reported as unverified, not as true.
 */
export function verifySubstitute(
  substitute: Substitute,
  probe: FontProbe,
  available: (family: string) => boolean,
): boolean | null {
  if (!available(substitute.missing) || !available(substitute.use)) return null;
  return metricsAgree(substitute.missing, substitute.use, probe);
}
