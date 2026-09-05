/**
 * Turn the T5 fixture into the one table `@pptx-studio/text` cannot compute.
 *
 * ```
 * node tools/ground-truth/text/bullets/write-tables.ts
 * ```
 *
 * Everything else T5 measured is a *rule* - an arithmetic the package
 * implements and the tests re-derive from the fixture. The date patterns are
 * not: PowerPoint formats a date field with Windows' own per-locale pattern, and
 * there is no arithmetic that produces "05/09/26" for `datetime3` in German and
 * "5 September 2026" for the same type in American English. So the cells are
 * data, and data that is measured gets generated rather than typed.
 *
 * The same shape as `tools/ground-truth/paint/fills/write-tables.ts` and `tools/ground-truth/paint/lines/write-tables.ts`, and for
 * the same reason: a measured constant must not quietly become somebody's
 * memory of one.
 */

import { readFileSync, writeFileSync } from 'node:fs';

interface DatePattern {
  readonly lang: string;
  readonly type: string;
  readonly pattern: string;
  readonly rendered: string;
}

interface Fixture {
  readonly datePatterns: readonly DatePattern[];
  readonly datePatternGaps: readonly { lang: string; type: string; rendered: string }[];
}

const fixture = JSON.parse(
  readFileSync('corpus/ground-truth/bullets.json', 'utf8'),
) as unknown as Fixture;

const byLang = new Map<string, Map<string, string>>();
for (const row of fixture.datePatterns) {
  const types = byLang.get(row.lang) ?? new Map<string, string>();
  const existing = types.get(row.type);
  if (existing !== undefined && existing !== row.pattern) {
    throw new Error(`${row.lang}/${row.type} derived two patterns: ${existing} and ${row.pattern}`);
  }
  types.set(row.type, row.pattern);
  byLang.set(row.lang, types);
}

/** A string literal with every non-ASCII character escaped, as the fixture is. */
function literal(value: string): string {
  const escaped = value
    .replace(/\\/g, '\\\\')
    .replace(/'/g, "\\'")
    .replace(/[-￿]/g, (ch) => `\\u${(ch.codePointAt(0) ?? 0).toString(16).padStart(4, '0')}`);
  return `'${escaped}'`;
}

const langs = [...byLang.keys()].sort();
const gaps = fixture.datePatternGaps.length;

const body = langs
  .map((lang) => {
    const types = byLang.get(lang);
    if (types === undefined) throw new Error(`no types for ${lang}`);
    const cells = [...types.keys()]
      .sort()
      .map((type) => `    ${type}: ${literal(types.get(type) ?? '')},`)
      .join('\n');
    return `  ${literal(lang)}: {\n${cells}\n  },`;
  })
  .join('\n');

const source = `/**
 * Windows' own date patterns, per locale and per \`a:fld/@type\`.
 *
 * **Generated. Do not edit.** Regenerate with
 * \`node tools/ground-truth/text/bullets/write-tables.ts\` after re-running experiment
 * T5; \`fields.test.ts\` re-derives every cell from
 * \`corpus/ground-truth/bullets.json\` and fails if the two disagree.
 *
 * ${String(fixture.datePatterns.length)} cells across ${String(langs.length)} locales, derived from what PowerPoint
 * rendered and verified by re-rendering each pattern and comparing it to the
 * string it came from. ${String(gaps)} of the measured readings could not be turned into
 * a pattern at all - the Japanese era calendar among them - and are deliberately
 * absent rather than approximated: \`renderField\` shows the cached text for a
 * cell that is not here, which is a visibly stale date rather than a confidently
 * wrong one.
 */
export const FIELD_DATE_PATTERNS: Readonly<Record<string, Readonly<Record<string, string>>>> = {
${body}
};
`;

writeFileSync('packages/text/src/fields/field-formats.gen.ts', source);
console.log(
  `wrote packages/text/src/fields/field-formats.gen.ts: ${String(fixture.datePatterns.length)} cells, ` +
    `${String(langs.length)} locales, ${String(gaps)} gaps`,
);
