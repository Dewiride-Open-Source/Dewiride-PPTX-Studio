/**
 * `+mn-lt` and its five siblings, resolved against a theme's `a:fontScheme`.
 *
 * Measured 8 of 8 against a reading that swaps the two collections, which fits
 * 2. ADR 0033.
 */

import { themeFontRef, type Typeface } from '../text.js';
import type { FontCollection, FontScheme } from '../types.js';

/** The four slots a `CT_TextCharacterProperties` can name. */
export type ScriptSlot = 'latin' | 'ea' | 'cs' | 'sym';

/**
 * The typeface a slot asks for, with any theme reference followed.
 *
 * `undefined` means the reference resolved to a collection entry the theme
 * leaves empty, which is not the same as the theme naming nothing.
 */
export function resolveTypeface(typeface: string, scheme: FontScheme): string | undefined {
  const ref = themeFontRef(typeface);
  if (ref === null) return typeface;
  const collection: FontCollection = ref.collection === 'major' ? scheme.major : scheme.minor;
  const name = collection[ref.script];
  if (name === null || name.length === 0) return undefined;
  return name;
}

/** The same, for a whole `CT_TextFont`. */
export function resolveTypefaceOf(font: Typeface, scheme: FontScheme): string | undefined {
  return resolveTypeface(font.typeface, scheme);
}

/**
 * Every distinct typeface a set of runs asks for, theme references followed.
 *
 * This is what `fontReport` is given: a report keyed on `+mn-lt` would name a
 * token no user recognises and would collapse two themes into one row.
 */
export function requestedTypefaces(
  fonts: Iterable<Typeface | undefined>,
  scheme: FontScheme,
): readonly string[] {
  const out: string[] = [];
  for (const font of fonts) {
    if (font === undefined) continue;
    const name = resolveTypeface(font.typeface, scheme);
    if (name !== undefined) out.push(name);
  }
  return out;
}
