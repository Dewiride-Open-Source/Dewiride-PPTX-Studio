import { FEATURE_RULES, PART_FEATURE_RULES } from '@pptx-studio/census';
import { describe, expect, it } from 'vitest';
import { CENSUS_FEATURE_KEYS } from './census-keys.gen.ts';

/**
 * The committed key list against the census's own table.
 *
 * `pnpm corpus` runs on a bare clone with nothing built, so it cannot import
 * `@pptx-studio/census` - nothing links workspace packages into `tools/`, and
 * the resolution that lets `tools/bench/bench.test.ts` do it comes from
 * Vitest's alias table rather than from Node. So the key list is committed,
 * which means it can go stale, which means something has to notice.
 *
 * This is that something. It runs in the `tools` project, where the census
 * *is* importable, and it is the only reason the generated file is trustworthy.
 * A feature rule added to the census fails here until
 * `node tools/corpus/write-census-keys.ts` is re-run.
 *
 * Deliberately not a regex over `features.ts`. Every rule there is a one-line
 * `key: 'x'` today, and a future `key: SOME_CONST` would silently under-report
 * - which is worse than not checking, because it looks checked.
 */

describe('the committed census key list', () => {
  const actual = [...FEATURE_RULES, ...PART_FEATURE_RULES].map((rule) => rule.key).sort();

  it('matches the census feature table exactly', () => {
    expect(CENSUS_FEATURE_KEYS).toEqual(actual);
  });

  it('has no duplicates', () => {
    expect(new Set(actual).size).toBe(actual.length);
  });

  it('is sorted, so a regeneration diff shows only what changed', () => {
    expect([...CENSUS_FEATURE_KEYS]).toEqual([...CENSUS_FEATURE_KEYS].sort());
  });
});
