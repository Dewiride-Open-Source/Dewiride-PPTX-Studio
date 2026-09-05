/**
 * Which of `a:latin`, `a:ea`, `a:cs` and `a:sym` draws each character.
 *
 * ## How this was measured, and why it is a reading rather than a fit
 *
 * A run may name four typefaces and PowerPoint picks one per character. Nothing
 * in the file says which, and the obvious instrument - measure each character's
 * advance and compare it against the same character in each candidate face - is
 * a fit with a tolerance and a tie-breaking problem.
 *
 * T5 read it instead. `Slide.Export(path, "EMF")` records PowerPoint's own GDI
 * calls, and GDI is asked for one face at a time, so a run whose characters
 * resolve to different faces comes out as **several text records, each naming
 * the face it was drawn in**. Every probe declared all four slots with a
 * different face in each, so the record names the slot outright: Georgia for
 * latin, MS Gothic for ea, Courier New for cs, Wingdings for sym.
 *
 * ## The finding
 *
 * Twenty-one samples, one per Unicode block, and one rule fits all of them:
 *
 * | block                                             | slot    |
 * | ------------------------------------------------- | ------- |
 * | Latin, Greek, Cyrillic, punctuation, currency, maths | `latin` |
 * | Hebrew, Arabic, Thai, Devanagari                  | `cs`    |
 * | Kana, Han, Hangul, fullwidth forms, CJK punctuation | `ea`    |
 * | the private-use area                              | `sym`   |
 *
 * The two readings worth naming because they are what a person writes first:
 *
 * - **`a:cs` is not "everything non-Latin".** The name says complex script and
 *   that is what it means - a script with shaping - so Greek and Cyrillic go to
 *   `a:latin`. Scored 8 of 21 for the "above Latin-1 is cs" reading.
 * - **`a:sym` is used, and only for the private-use area.** Scored 20 of 21 for
 *   never using it, which is close enough to look right on a whole deck and
 *   wrong on every Wingdings run in it.
 *
 * ## What this is not
 *
 * Not font fallback. The slot says which typeface PowerPoint *asks* for; if that
 * face has no glyph, Windows substitutes another and may scale it - Courier New
 * has no Thai, and the Thai probe came back in Leelawadee UI at 1.4 times the
 * size it asked for. Substitution is 3.7's subject and 8.1's data; the slot is
 * this one's.
 */

import { TextError } from '../errors.js';

/** The four typeface slots a `CT_TextCharacterProperties` can name. */
export type ScriptSlot = 'latin' | 'ea' | 'cs' | 'sym';

/** The four slots as a run states them. An absent slot falls back to `latin`. */
export interface RunTypefaces {
  readonly latin: string;
  readonly ea?: string | undefined;
  readonly cs?: string | undefined;
  readonly sym?: string | undefined;
}

/** Inclusive code point ranges, in the order they are tested. */
const COMPLEX_SCRIPT: readonly (readonly [number, number])[] = [
  [0x0590, 0x05ff], // Hebrew
  [0x0600, 0x06ff], // Arabic
  [0x0700, 0x074f], // Syriac
  [0x0780, 0x07bf], // Thaana
  [0x0900, 0x097f], // Devanagari
  [0x0980, 0x0dff], // Bengali through Sinhala
  [0x0e00, 0x0e7f], // Thai
  [0x0e80, 0x0eff], // Lao
  [0x0f00, 0x0fff], // Tibetan
  [0x1000, 0x109f], // Myanmar
  [0x1780, 0x17ff], // Khmer
  [0xfb00, 0xfdff], // Hebrew and Arabic presentation forms
  [0xfe70, 0xfeff], // Arabic presentation forms-B
];

const EAST_ASIAN: readonly (readonly [number, number])[] = [
  [0x1100, 0x11ff], // Hangul Jamo
  [0x2e80, 0x2eff], // CJK radicals
  [0x3000, 0x303f], // CJK symbols and punctuation
  [0x3040, 0x30ff], // Hiragana and Katakana
  [0x3100, 0x312f], // Bopomofo
  [0x3130, 0x318f], // Hangul compatibility Jamo
  [0x3190, 0x319f], // Kanbun
  [0x31f0, 0x31ff], // Katakana phonetic extensions
  [0x3200, 0x33ff], // Enclosed CJK and CJK compatibility
  [0x3400, 0x4dbf], // CJK extension A
  [0x4e00, 0x9fff], // CJK unified ideographs
  [0xac00, 0xd7af], // Hangul syllables
  [0xf900, 0xfaff], // CJK compatibility ideographs
  [0xff00, 0xffef], // Halfwidth and fullwidth forms
];

/**
 * The private use area.
 *
 * Only the Basic Multilingual Plane's. The supplementary private use planes
 * exist and were not probed, so they are not claimed: a character there falls to
 * `latin` like anything else this experiment did not measure.
 */
const PRIVATE_USE: readonly [number, number] = [0xe000, 0xf8ff];

const inAny = (code: number, ranges: readonly (readonly [number, number])[]): boolean =>
  ranges.some(([from, to]) => code >= from && code <= to);

/**
 * Which slot draws `code`.
 *
 * The ranges beyond the twenty-one samples - Syriac, Khmer, Bopomofo and the
 * rest - are the rest of the block each measured sample sits in, added because a
 * table that named only the four probed scripts would send Lao to `a:latin`
 * while sending Thai to `a:cs`, which is a difference nothing measured supports.
 * They are a generalisation over the same rule and are marked as such here so
 * that nobody later mistakes them for readings.
 */
export function scriptSlotOf(code: number): ScriptSlot {
  if (!Number.isInteger(code) || code < 0 || code > 0x10ffff) {
    throw new TextError('TEXT_SCRIPT_RUN', `${String(code)} is not a code point`, String(code));
  }
  if (code >= PRIVATE_USE[0] && code <= PRIVATE_USE[1]) return 'sym';
  if (inAny(code, COMPLEX_SCRIPT)) return 'cs';
  if (inAny(code, EAST_ASIAN)) return 'ea';
  return 'latin';
}

/**
 * The typeface a character is drawn in, given what the run names.
 *
 * A slot the run does not state falls back to `a:latin`, which is measured: with
 * `a:ea` absent, the Japanese probes came back in the latin face rather than in
 * a substituted one.
 */
export function typefaceFor(code: number, faces: RunTypefaces): string {
  const slot = scriptSlotOf(code);
  if (slot === 'ea') return faces.ea ?? faces.latin;
  if (slot === 'cs') return faces.cs ?? faces.latin;
  if (slot === 'sym') return faces.sym ?? faces.latin;
  return faces.latin;
}

/** One maximal stretch of a string drawn in a single typeface. */
export interface ScriptRun {
  /** Index of the first code unit, into the original string. */
  readonly start: number;
  /** Index one past the last code unit. */
  readonly end: number;
  readonly slot: ScriptSlot;
  readonly typeface: string;
  readonly text: string;
}

/**
 * Split a run of text into the stretches PowerPoint draws in one face each.
 *
 * This is the shape a renderer needs: the EMF showed PowerPoint issuing one
 * drawing call per stretch, and a layout that measures the whole string in one
 * face measures something PowerPoint never drew.
 *
 * Surrogate pairs are kept whole - iterating a JavaScript string by code unit
 * would split an astral character down the middle and ask for two faces.
 *
 * That guard is, today, **unexercised**: 3.5's mutation sweep changed it to a
 * plain code-unit walk and nothing went red, because every range in the tables
 * above is inside the Basic Multilingual Plane and a lone surrogate falls to
 * `latin` exactly as an astral character does. It survives as an equivalent
 * mutant rather than a missing assertion, and it is kept because the first
 * supplementary-plane range anyone adds - CJK extension B is the obvious one -
 * turns it into the difference between one run and two. Nobody should read it
 * as measured behaviour until there is a range that needs it.
 */
export function splitScriptRuns(text: string, faces: RunTypefaces): readonly ScriptRun[] {
  const runs: ScriptRun[] = [];
  let start = 0;
  let slot: ScriptSlot | null = null;
  let typeface = '';
  for (let at = 0; at < text.length;) {
    const code = text.codePointAt(at);
    if (code === undefined) break;
    const width = code > 0xffff ? 2 : 1;
    const here = scriptSlotOf(code);
    const face = typefaceFor(code, faces);
    if (slot !== null && (here !== slot || face !== typeface)) {
      runs.push({ start, end: at, slot, typeface, text: text.slice(start, at) });
      start = at;
    }
    slot = here;
    typeface = face;
    at += width;
  }
  if (slot !== null) {
    runs.push({ start, end: text.length, slot, typeface, text: text.slice(start) });
  }
  return runs;
}
