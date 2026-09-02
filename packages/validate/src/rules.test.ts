import { describe, expect, it } from 'vitest';
import { RULES, RULE_IDS, BASELINE_RULES, ruleById } from './rules.js';
import { validatePackage } from './validate.js';
import { deck } from './testing/deck.js';

/**
 * The table itself.
 *
 * A rule that is defined and never called looks exactly like a rule that is
 * enforced, from every angle except the one that matters. These tests are the
 * angle that matters: they hold the count, the ids, the wiring and the shape of
 * each entry, so that adding a thirtieth rule is a deliberate act with a diff
 * rather than something that happens.
 */

describe('the rule table', () => {
  it('has twenty-nine rules, and the plan says twenty-nine', () => {
    expect(RULES).toHaveLength(29);
  });

  it('numbers them V001…V029 with no gaps', () => {
    expect(RULE_IDS).toEqual(
      Array.from({ length: 29 }, (_, i) => 'V' + String(i + 1).padStart(3, '0')),
    );
  });

  it('covers all seven categories, and the six the plan names', () => {
    const counted = new Map<string, number>();
    for (const rule of RULES) counted.set(rule.category, (counted.get(rule.category) ?? 0) + 1);

    // The plan's appendix names six groups. The seventh, `refused`, is the one
    // this package could not have got from any specification: five rules, each
    // the record of a package built with one change in it and declined.
    expect([...counted.keys()].sort()).toEqual([
      'ids',
      'order',
      'package',
      'preservation',
      'refused',
      'relationships',
      'required',
    ]);
    expect(counted.get('refused')).toBe(5);
    expect(counted.get('preservation')).toBe(3);
  });

  it('says how it knows, for every rule', () => {
    for (const rule of RULES) {
      expect(['schema', 'measured', 'both'], rule.id).toContain(rule.evidence);
      // Not a formatting check. A rule with no argument behind it is a rule
      // somebody deletes the first time it fires on their file, and they will
      // be right to.
      expect(rule.why.length, rule.id).toBeGreaterThan(80);
      expect(rule.title.length, rule.id).toBeGreaterThan(10);
    }
  });

  it('leaves nothing measured resting on the specification alone', () => {
    // Every `refused` rule must claim measurement. If one of them ever says
    // `schema` alone, it belongs in another category and the categories have
    // stopped meaning anything. `both` is allowed and `V024` is why: CT_SerTx
    // really is a choice of two elements in the schema, and we still only
    // learnt that a third one is a whole-package refusal by trying it.
    for (const rule of RULES.filter((r) => r.category === 'refused')) {
      expect(['measured', 'both'], rule.id).toContain(rule.evidence);
    }
  });

  it('marks exactly the preservation rules as needing a baseline', () => {
    expect(BASELINE_RULES).toEqual(['V027', 'V028', 'V029']);
    for (const rule of RULES) {
      expect('needsBaseline' in rule, rule.id).toBe(rule.category === 'preservation');
    }
  });

  it('has one warning, and it is the one PowerPoint opens', () => {
    // Everything else is a refusal or a repair. `V021` is the odd one: an
    // unmatched placeholder inherits nothing and the file still opens, so
    // blocking an export over it would stop a user saving a deck that already
    // works everywhere.
    const warnings = RULES.filter((rule) => rule.severity === 'warning').map((rule) => rule.id);
    expect(warnings).toEqual(['V021']);
  });

  it('resolves by id, and only by an id that exists', () => {
    expect(ruleById('V001')?.category).toBe('package');
    expect(ruleById('V030')).toBeUndefined();
    expect(ruleById('')).toBeUndefined();
  });

  it('runs every rule it defines', () => {
    // The wiring check. `validatePackage` reports which rules ran; if a rule
    // had a definition and no implementation the table lookup would throw, and
    // if one were quietly dropped from the dispatch table it would be missing
    // here.
    const { bytes, store } = deck();
    const report = validatePackage({ bytes, store, baseline: store });
    expect([...report.checked].sort()).toEqual([...RULE_IDS].sort());
  });

  it('refuses a rule id that does not exist rather than ignoring it', () => {
    const { bytes } = deck();
    expect(() => validatePackage({ bytes, rules: ['V099' as 'V001'] })).toThrow(/no rule V099/);
  });
});
