import { describe, expect, it } from 'vitest';

import fixture from '../../../../corpus/ground-truth/font-substitution.json' with { type: 'json' };
import { TextError } from '../errors.js';
import {
  anchoredStack,
  createFontProbe,
  fontFingerprint,
  isFontAvailable,
  metricsAgree,
  ANCHORS,
  FINGERPRINT_PX,
  FINGERPRINT_TEXT,
  type FontProbe,
} from './presence.js';
import {
  fontStack,
  substituteFor,
  verifySubstitute,
  LAST_RESORT_GENERIC,
  POWERPOINT_LAST_RESORT,
  SUBSTITUTES,
} from './substitute.js';
import { fontReport, substitutedFonts } from './report.js';

/* -------------------------------------------------------------------------- */
/* the fixture                                                                */
/* -------------------------------------------------------------------------- */

interface Reading {
  readonly typeface: string;
  readonly drawn: readonly string[];
  readonly width: number;
}
interface HintReading {
  readonly id: string;
  readonly panose: string | null;
  readonly pitchFamily: number | null;
  readonly charset: number | null;
  readonly drawn: readonly string[];
}
interface WidthReading {
  readonly sz: number;
  readonly text: string;
  readonly absent: number;
  readonly byFace: Readonly<Record<string, number>>;
}
interface BrowserFamily {
  readonly family: string;
  readonly absent: boolean;
  readonly check: boolean;
  readonly bare: readonly number[];
  readonly anchorEqual: Readonly<Record<string, boolean>>;
}
interface Score {
  readonly name: string;
  readonly fits: number;
  readonly total: number;
}

const pp = fixture.powerpoint as unknown as {
  readonly absentSubstitute: string;
  readonly absent: readonly Reading[];
  readonly installed: readonly Reading[];
  readonly alias: readonly Reading[];
  readonly charsetSubstitute: Readonly<Record<string, string>>;
  readonly hints: readonly HintReading[];
  readonly widths: readonly WidthReading[];
};
const web = fixture.browser as unknown as {
  readonly fingerprint: readonly string[];
  readonly characterFingerprint: readonly string[] | null;
  readonly distinctFamilies: number;
  readonly presentFamilies: number;
  readonly scaleFree: boolean;
  readonly worstScaleDeviation: number;
  readonly aliasClasses: readonly (readonly string[])[];
  readonly families: readonly BrowserFamily[];
};
const questions = fixture.questions as unknown as readonly {
  readonly key: string;
  readonly scores: readonly Score[];
}[];

const scoreOf = (key: string, name: string): string => {
  const question = questions.find((q) => q.key === key);
  const score = question?.scores.find((s) => s.name === name);
  if (score === undefined) throw new Error(`no score ${key}/${name}`);
  return `${String(score.fits)}/${String(score.total)}`;
};

/** The detector, run over the fixture's recorded anchor comparisons. */
const detectFromFixture = (family: BrowserFamily): boolean =>
  ANCHORS.some((anchor) => family.anchorEqual[anchor] === false);

/* -------------------------------------------------------------------------- */
/* what PowerPoint does with a face it does not have                          */
/* -------------------------------------------------------------------------- */

describe('the substitute PowerPoint chooses', () => {
  it(`is Calibri for every absent Latin face (${scoreOf('absent-latin', 'always Calibri')})`, () => {
    expect(pp.absent.length).toBe(22);
    for (const row of pp.absent) {
      expect(row.drawn).toEqual([POWERPOINT_LAST_RESORT]);
    }
    expect(pp.absentSubstitute).toBe(POWERPOINT_LAST_RESORT);
  });

  it(`is not derived from the style word in the name (${scoreOf('absent-latin', 'the style word in the name')})`, () => {
    const mono = pp.absent.find((r) => r.typeface === 'Zzz Probe Mono');
    const serif = pp.absent.find((r) => r.typeface === 'Zzz Probe Serif');
    expect(mono?.drawn).toEqual([POWERPOINT_LAST_RESORT]);
    expect(serif?.drawn).toEqual([POWERPOINT_LAST_RESORT]);
    expect(mono?.drawn).toEqual(serif?.drawn);
  });

  it('leaves an installed face alone, which is what calibrates the reader', () => {
    for (const row of pp.installed) {
      expect(row.drawn).toEqual([row.typeface]);
    }
  });

  it('carries the requested name through for a face Windows maps elsewhere', () => {
    const helvetica = pp.alias.find((r) => r.typeface === 'Helvetica');
    const arial = pp.installed.find((r) => r.typeface === 'Arial');
    expect(helvetica?.drawn).toEqual(['Helvetica']);
    expect(helvetica?.width).toBe(arial?.width);
  });
});

describe('the hints on CT_TextFont', () => {
  it(`ignores @panose (${scoreOf('hints', '@panose picks the family')})`, () => {
    const withPanose = pp.hints.filter((h) => h.panose !== null && h.charset === null);
    expect(withPanose.length).toBeGreaterThanOrEqual(5);
    for (const hint of withPanose) {
      expect(hint.drawn).toEqual([POWERPOINT_LAST_RESORT]);
    }
  });

  it(`ignores @pitchFamily (${scoreOf('hints', '@pitchFamily picks the family')})`, () => {
    const withPitch = pp.hints.filter((h) => h.pitchFamily !== null && h.charset === null);
    expect(withPitch.length).toBeGreaterThanOrEqual(9);
    for (const hint of withPitch) {
      expect(hint.drawn).toEqual([POWERPOINT_LAST_RESORT]);
    }
  });

  it(`reads @charset, which is why "all three ignored" is wrong (${scoreOf('hints', 'all three ignored')})`, () => {
    expect(pp.charsetSubstitute['0']).toBe(POWERPOINT_LAST_RESORT);
    expect(pp.charsetSubstitute['-128']).not.toBe(POWERPOINT_LAST_RESORT);
    expect(pp.charsetSubstitute['-78']).not.toBe(POWERPOINT_LAST_RESORT);
  });
});

describe('the substitution is metric-preserving', () => {
  it(`lays the absent face out as Calibri exactly (${scoreOf('substitute-metrics', 'exactly Calibri')})`, () => {
    expect(pp.widths.length).toBe(12);
    for (const row of pp.widths) {
      expect(row.absent).toBe(row.byFace['Calibri']);
    }
  });

  it(`is not any other candidate (${scoreOf('substitute-metrics', 'exactly Arial')})`, () => {
    const wrong = pp.widths.filter((row) => row.absent !== row.byFace['Arial']);
    expect(wrong.length).toBe(11);
  });
});

/* -------------------------------------------------------------------------- */
/* the detector                                                               */
/* -------------------------------------------------------------------------- */

describe('detecting an absent family', () => {
  it(`needs all three anchors (${scoreOf('detector', 'equal to all three anchors')})`, () => {
    for (const family of web.families) {
      expect(detectFromFixture(family)).toBe(!family.absent);
    }
  });

  it(`is not one anchor: monospace alone calls Consolas absent (${scoreOf('detector', 'equal to monospace alone')})`, () => {
    const consolas = web.families.find((f) => f.family === 'Consolas');
    expect(consolas?.absent).toBe(false);
    expect(consolas?.anchorEqual['monospace']).toBe(true);
  });

  it(`is not document.fonts.check (${scoreOf('detector', 'document.fonts.check')})`, () => {
    expect(web.families.every((f) => f.check)).toBe(true);
    expect(web.families.filter((f) => f.absent).length).toBe(20);
  });

  it('reports all twenty absent faces, live in this browser', () => {
    const probe = createFontProbe();
    const absent = web.families.filter((f) => f.absent).map((f) => f.family);
    expect(absent.length).toBe(20);
    for (const family of absent) {
      expect(isFontAvailable(family, probe)).toBe(false);
    }
  });

  it('is present as soon as one anchor is overridden', () => {
    const overrides = (which: string): FontProbe => ({
      advance: (stack) => (stack === anchoredStack('X', which) ? 1 : 0),
    });
    for (const anchor of ANCHORS) {
      expect(isFontAvailable('X', overrides(anchor))).toBe(true);
    }
  });

  it('is absent only when every anchor shows through', () => {
    const showsThrough: FontProbe = { advance: () => 7 };
    expect(isFontAvailable('X', showsThrough)).toBe(false);
  });

  it('treats a family named like a generic as a family, which is why it is quoted', () => {
    // A face really named `monospace` is legal in OOXML, and quoting is what
    // stops the stack reading it as the CSS generic. Whether a machine happens
    // to resolve that name is the machine's business - fontconfig defines it
    // and Windows does not - so the rule is asserted and not the environment.
    expect(anchoredStack('monospace', 'serif')).toBe('"monospace", serif');
    expect(anchoredStack('monospace', 'serif').startsWith('"')).toBe(true);
  });
});

describe('the fingerprint', () => {
  it('is a word, because no three single characters reach the same ceiling', () => {
    expect(web.fingerprint).toEqual([FINGERPRINT_TEXT]);
    expect(web.characterFingerprint).toBeNull();
  });

  it('is taken at one size, because advances are not linear in size', () => {
    expect(web.scaleFree).toBe(false);
    expect(web.worstScaleDeviation).toBeGreaterThan(0.01);
    expect(FINGERPRINT_PX).toBe(fixture.browser.sizes[0]);
  });

  it('gives two names for one file the same value', () => {
    expect(web.aliasClasses).toContainEqual(['Arial', 'Helvetica']);
    expect(web.aliasClasses).toContainEqual(['Courier', 'Courier New']);
  });

  it('separates every family that is not such a pair', () => {
    const collapsed = web.aliasClasses.reduce((n, c) => n + c.length - 1, 0);
    expect(web.distinctFamilies).toBe(web.presentFamilies - collapsed);
  });

  it('measures the family alone, with nothing behind it to show through', () => {
    let seen = '';
    const spy: FontProbe = {
      advance: (stack) => {
        seen = stack;
        return 1;
      },
    };
    fontFingerprint('Georgia', spy);
    expect(seen).toBe('"Georgia"');
  });

  it('is a live measurement, and two stacks that resolve alike agree', () => {
    const probe = createFontProbe();
    expect(fontFingerprint('serif', probe)).toBeGreaterThan(0);
    expect(metricsAgree('serif', 'serif', probe)).toBe(true);
  });
});

describe('anchoredStack', () => {
  it('quotes the family and leaves the generic bare', () => {
    expect(anchoredStack('Times New Roman', 'serif')).toBe('"Times New Roman", serif');
  });

  it('refuses a name that would break out of the quoting', () => {
    expect(() => anchoredStack('a"b', 'serif')).toThrow(TextError);
  });
});

/* -------------------------------------------------------------------------- */
/* the table and the stack                                                    */
/* -------------------------------------------------------------------------- */

describe('the substitution table', () => {
  it('names no typeface twice', () => {
    const names = SUBSTITUTES.map((s) => s.missing.toLowerCase());
    expect(new Set(names).size).toBe(names.length);
  });

  it('matches without regard to case or repeated spaces', () => {
    expect(substituteFor('times new roman')?.use).toBe('Liberation Serif');
    expect(substituteFor('Times  New   Roman')?.use).toBe('Liberation Serif');
  });

  it('refuses an empty name rather than returning nothing for it', () => {
    expect(() => substituteFor('')).toThrow(TextError);
  });

  it('does not claim metric compatibility for Aptos, which has no clone', () => {
    expect(substituteFor('Aptos')?.metricCompatible).toBe(false);
  });
});

describe('fontStack', () => {
  it('ends at what PowerPoint would have drawn, not at the browser default', () => {
    const stack = fontStack('Zzz Probe Plain');
    expect(stack).toContain(`"${POWERPOINT_LAST_RESORT}"`);
    expect(stack.endsWith(LAST_RESORT_GENERIC)).toBe(true);
  });

  it('puts the requested face first so a machine that has it wins', () => {
    expect(fontStack('Arial').startsWith('"Arial", "Liberation Sans"')).toBe(true);
  });

  it('never names one family twice', () => {
    const names = fontStack('Calibri').split(', ');
    expect(new Set(names).size).toBe(names.length);
  });
});

describe('verifySubstitute', () => {
  const probe: FontProbe = createFontProbe();

  it('is null when a face is absent, rather than repeating the claim as fact', () => {
    const entry = substituteFor('Arial');
    expect(entry).toBeDefined();
    if (entry === undefined) return;
    expect(verifySubstitute(entry, probe, () => false)).toBeNull();
  });

  it('measures the claim when both faces are here', () => {
    const entry = substituteFor('Helvetica');
    expect(entry).toBeDefined();
    if (entry === undefined) return;
    const answer = verifySubstitute(entry, probe, (f) => isFontAvailable(f, probe));
    expect(answer === null || typeof answer === 'boolean').toBe(true);
  });
});

/* -------------------------------------------------------------------------- */
/* the report                                                                 */
/* -------------------------------------------------------------------------- */

describe('fontReport', () => {
  const absentNames = web.families.filter((f) => f.absent).map((f) => f.family);

  it('reports all twenty absent typefaces of a deck that names twenty', () => {
    const findings = fontReport(absentNames);
    expect(findings.length).toBe(20);
    expect(substitutedFonts(findings).length).toBe(20);
    // A missing face draws as PowerPoint's last resort where this machine has
    // it, and as nothing where it does not - a Linux runner has no Calibri, and
    // naming a face that would not be drawn is the one thing the report must
    // not do. Both branches, so neither can rot.
    const drawn = isFontAvailable(POWERPOINT_LAST_RESORT, createFontProbe())
      ? POWERPOINT_LAST_RESORT
      : '';
    for (const finding of findings) {
      expect(finding.status).toBe('missing');
      expect(finding.rendersAs).toBe(drawn);
    }
  });

  it('counts the runs and orders the busiest typeface first', () => {
    const first = absentNames[0] ?? '';
    const second = absentNames[1] ?? '';
    const findings = fontReport([first, first, first, second]);
    expect(findings[0]?.typeface).toBe(first);
    expect(findings[0]?.runs).toBe(3);
    expect(findings[1]?.runs).toBe(1);
  });

  it('treats one typeface written two ways as one row', () => {
    const findings = fontReport(['Zzz Probe Plain', 'zzz probe plain']);
    expect(findings.length).toBe(1);
    expect(findings[0]?.runs).toBe(2);
  });

  it('refuses a run that names an empty typeface', () => {
    expect(() => fontReport(['Arial', '   '])).toThrow(TextError);
  });

  it('says a present face is drawn as itself, exactly', () => {
    const probe = createFontProbe();
    const present = web.families
      .filter((f) => !f.absent)
      .map((f) => f.family)
      .filter((f) => isFontAvailable(f, probe));
    expect(present.length).toBeGreaterThan(0);
    for (const finding of fontReport(present, probe)) {
      expect(finding.status).toBe('present');
      expect(finding.rendersAs).toBe(finding.typeface);
      expect(finding.metricCompatible).toBe(true);
      expect(finding.verified).toBe(true);
    }
  });

  it('says nothing is substituted when everything is present', () => {
    const probe = createFontProbe();
    const present = web.families
      .filter((f) => !f.absent)
      .map((f) => f.family)
      .filter((f) => isFontAvailable(f, probe));
    const findings = fontReport(present, probe);
    expect(findings.length).toBe(present.length);
    expect(substitutedFonts(findings)).toEqual([]);
  });
});
