/**
 * The locks on `--record`, and that every declared failure can happen.
 *
 * `--record` is the only way to turn a red gate green, so these assert the
 * refusals rather than the success: a guard nothing tests is a comment. ADR
 * 0035.
 */

import { readFileSync, readdirSync } from 'node:fs';
import { join } from 'node:path';

import { describe, expect, it } from 'vitest';

import { repoPath } from '../repo/root.ts';

import { FIDELITY_ERROR_CODES, isFidelityError, type FidelityErrorCode } from './errors.ts';
import {
  assertBaselineCovers,
  assertExpectedCount,
  assertMayRecord,
  baselineGaps,
  parseRecordArgs,
  type RecordArgs,
} from './record.ts';

const RECORDING: RecordArgs = {
  record: true,
  why: 'the renderer is right',
  expect: 2,
  bootstrap: false,
};

function codeOf(run: () => void): string | false {
  try {
    run();
  } catch (error) {
    return isFidelityError(error) ? error.code : false;
  }
  return 'nothing was thrown';
}

describe('the arguments a record takes', () => {
  it('reads the three it knows', () => {
    expect(parseRecordArgs(['--record', '--why', 'because', '--expect', '4'])).toEqual({
      record: true,
      why: 'because',
      expect: 4,
      bootstrap: false,
    });
  });

  it('defaults to scoring, not recording', () => {
    expect(parseRecordArgs([])).toEqual({
      record: false,
      why: null,
      expect: null,
      bootstrap: false,
    });
  });

  it('reads the bootstrap flag CI passes', () => {
    expect(parseRecordArgs(['--record', '--bootstrap'])).toEqual({
      record: true,
      why: null,
      expect: null,
      bootstrap: true,
    });
  });

  it('refuses an argument it does not know rather than ignoring it', () => {
    expect(codeOf(() => parseRecordArgs(['--force']))).toBe('FID_RECORD_UNJUSTIFIED');
  });
});

describe('leave to rewrite a baseline', () => {
  it('is not needed by a run that only scores', () => {
    const args = parseRecordArgs([]);
    expect(() => {
      assertMayRecord(args, { inCi: true, hasBaseline: true });
    }).not.toThrow();
  });

  it('is refused in CI, where no human reads the diff', () => {
    expect(
      codeOf(() => {
        assertMayRecord(RECORDING, { inCi: true, hasBaseline: true });
      }),
    ).toBe('FID_RECORD_IN_CI');
  });

  it('is refused in CI even for a first baseline', () => {
    expect(
      codeOf(() => {
        assertMayRecord(RECORDING, { inCi: true, hasBaseline: false });
      }),
    ).toBe('FID_RECORD_IN_CI');
  });

  it('is not needed to write a first baseline, which overwrites nothing', () => {
    expect(() => {
      assertMayRecord(
        { record: true, why: null, expect: null, bootstrap: false },
        { inCi: false, hasBaseline: false },
      );
    }).not.toThrow();
  });

  it('needs a reason to overwrite one that exists', () => {
    expect(
      codeOf(() => {
        assertMayRecord({ ...RECORDING, why: null }, { inCi: false, hasBaseline: true });
      }),
    ).toBe('FID_RECORD_UNJUSTIFIED');
    expect(
      codeOf(() => {
        assertMayRecord({ ...RECORDING, why: '   ' }, { inCi: false, hasBaseline: true });
      }),
    ).toBe('FID_RECORD_UNJUSTIFIED');
  });

  it('needs a count, and a whole non-negative one', () => {
    for (const expected of [null, Number.NaN, -1, 1.5]) {
      expect(
        codeOf(() => {
          assertMayRecord({ ...RECORDING, expect: expected }, { inCi: false, hasBaseline: true });
        }),
        String(expected),
      ).toBe('FID_RECORD_UNJUSTIFIED');
    }
  });

  it('lets a runner write a platform its first baseline, and only that', () => {
    const boot = { record: true, why: null, expect: null, bootstrap: true };
    expect(() => {
      assertMayRecord(boot, { inCi: true, hasBaseline: false });
    }).not.toThrow();
    // The file it writes gates nothing until a person commits it. Rewriting one
    // that already gates could turn a red run green, which is the whole point.
    expect(
      codeOf(() => {
        assertMayRecord(boot, { inCi: true, hasBaseline: true });
      }),
    ).toBe('FID_RECORD_UNJUSTIFIED');
  });

  it('refuses a bare --record in CI even with a reason and a count', () => {
    expect(
      codeOf(() => {
        assertMayRecord(RECORDING, { inCi: true, hasBaseline: false });
      }),
    ).toBe('FID_RECORD_IN_CI');
  });

  it('refuses --bootstrap where a baseline already exists, on any machine', () => {
    expect(
      codeOf(() => {
        assertMayRecord({ ...RECORDING, bootstrap: true }, { inCi: false, hasBaseline: true });
      }),
    ).toBe('FID_RECORD_UNJUSTIFIED');
  });

  it('lets a justified rewrite through', () => {
    expect(() => {
      assertMayRecord(RECORDING, { inCi: false, hasBaseline: true });
    }).not.toThrow();
  });
});

describe('the count the author committed to', () => {
  it('passes when the rewrite touches exactly that many', () => {
    expect(() => {
      assertExpectedCount(RECORDING, 2);
    }).not.toThrow();
  });

  it('fails when it touches more, which is a second change riding along', () => {
    expect(
      codeOf(() => {
        assertExpectedCount(RECORDING, 3);
      }),
    ).toBe('FID_RECORD_UNJUSTIFIED');
  });

  it('fails when it touches fewer, which means the reason was wrong', () => {
    expect(
      codeOf(() => {
        assertExpectedCount(RECORDING, 1);
      }),
    ).toBe('FID_RECORD_UNJUSTIFIED');
  });

  it('says nothing about a run that is not recording', () => {
    expect(() => {
      assertExpectedCount({ record: false, why: null, expect: 2, bootstrap: false }, 9);
    }).not.toThrow();
  });
});

describe('what the baseline and the run disagree exists', () => {
  it('names both directions, sorted', () => {
    expect(baselineGaps(['b', 'a', 'gone'], ['b', 'a', 'new'])).toEqual({
      unrecorded: ['new'],
      vanished: ['gone'],
    });
  });

  it('is empty when they describe the same slides', () => {
    expect(baselineGaps(['a', 'b'], ['b', 'a'])).toEqual({ unrecorded: [], vanished: [] });
  });

  it('refuses to score a slide nothing gates, and names it', () => {
    expect(
      codeOf(() => {
        assertBaselineCovers({ unrecorded: ['a10-rtl-cjk-03'], vanished: [] });
      }),
    ).toBe('FID_BASELINE_INCOMPLETE');
  });

  it('says nothing about a slide the baseline has and the run lost', () => {
    // That one is a regression the gate reports, not a hole in the record.
    expect(() => {
      assertBaselineCovers({ unrecorded: [], vanished: ['b02-layouts-10'] });
    }).not.toThrow();
  });
});

/**
 * Every `.ts` that could throw one, so a new file cannot hide a dead code.
 *
 * `errors.ts` declares the codes and the tests quote them, so counting either
 * would make every code look live and the check vacuous.
 */
function throwSitesUnder(directory: string): readonly string[] {
  const out: string[] = [];
  for (const entry of readdirSync(directory, { withFileTypes: true })) {
    const full = join(directory, entry.name);
    if (entry.isDirectory()) out.push(...throwSitesUnder(full));
    else if (entry.name.endsWith('.ts') && !entry.name.endsWith('.test.ts')) {
      if (entry.name !== 'errors.ts') out.push(full);
    }
  }
  return out;
}

describe('the failures this harness declares', () => {
  it('can each actually happen', () => {
    const sources = [
      repoPath('tools/fidelity'),
      repoPath('tools/gate2'),
      repoPath('tools/ground-truth/render/fidelity'),
    ]
      .flatMap((directory) => throwSitesUnder(directory))
      .map((file) => readFileSync(file, 'utf8'))
      .join('\n');
    const unthrown = FIDELITY_ERROR_CODES.filter(
      (code: FidelityErrorCode) => !sources.includes(`'${code}'`),
    );
    expect(unthrown, 'declared and never thrown').toEqual([]);
  });

  it('lists every member of the union exactly once', () => {
    expect(new Set(FIDELITY_ERROR_CODES).size).toBe(FIDELITY_ERROR_CODES.length);
  });
});
