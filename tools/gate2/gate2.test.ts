/**
 * What Gate 2 detects, and what it refuses to call a match.
 *
 * The gate's whole value is that it fails when a feature it names has no slide,
 * so a predicate that is too generous is worse than no gate at all: it reports
 * coverage the corpus does not have. Each of these asserts the near miss as
 * well as the hit. ADR 0038.
 */

import { describe, expect, it } from 'vitest';

import { GATE2_FEATURES, MARKUP_PATTERNS, type SlideFacts } from './features.ts';

function facts(overrides: Partial<SlideFacts> = {}): SlideFacts {
  return {
    key: 'a00-probe-01',
    deck: 'a00-probe',
    slide: 1,
    matched: [],
    groupDepth: 0,
    transformedGroups: 0,
    ...overrides,
  };
}

function holds(key: string, given: SlideFacts): boolean {
  const feature = GATE2_FEATURES.find((entry) => entry.key === key);
  if (feature === undefined) throw new Error(`no gate feature called ${key}`);
  return feature.holds(given);
}

function matches(key: string, xml: string): boolean {
  const rule = MARKUP_PATTERNS.find((entry) => entry.key === key);
  if (rule === undefined) throw new Error(`no markup pattern called ${key}`);
  return new RegExp(rule.pattern).test(xml);
}

describe('the nine features Gate 2 names', () => {
  it('names exactly the nine, each with a distinct key', () => {
    expect(GATE2_FEATURES).toHaveLength(9);
    expect(new Set(GATE2_FEATURES.map((feature) => feature.key)).size).toBe(9);
  });

  it('holds none of them for a slide that carries nothing', () => {
    const bare = facts();
    for (const feature of GATE2_FEATURES) {
      expect(feature.holds(bare), feature.key).toBe(false);
    }
  });

  it('wants a nested group AND a transformed one, not either alone', () => {
    expect(holds('nestedGroups', facts({ groupDepth: 2, transformedGroups: 1 }))).toBe(true);
    // Two groups side by side, one of them rotated, is not a nest.
    expect(holds('nestedGroups', facts({ groupDepth: 1, transformedGroups: 1 }))).toBe(false);
    // A group inside a group, neither of them turned, is not what Gate 2 asks.
    expect(holds('nestedGroups', facts({ groupDepth: 2, transformedGroups: 0 }))).toBe(false);
  });

  it('counts a themed background from either p:bgRef or p:bgPr', () => {
    expect(holds('themedBackgrounds', facts({ matched: ['bgRef'] }))).toBe(true);
    expect(holds('themedBackgrounds', facts({ matched: ['bgPr'] }))).toBe(true);
    expect(holds('themedBackgrounds', facts({ matched: ['gradFill'] }))).toBe(false);
  });
});

describe('the markup patterns, at their near misses', () => {
  it('is a set of well-formed expressions with distinct keys', () => {
    expect(new Set(MARKUP_PATTERNS.map((rule) => rule.key)).size).toBe(MARKUP_PATTERNS.length);
    for (const rule of MARKUP_PATTERNS) {
      expect(() => new RegExp(rule.pattern), rule.key).not.toThrow();
    }
  });

  it('does not read a solid line as a dashed one', () => {
    expect(matches('dash', '<a:prstDash val="solid"/>')).toBe(false);
    expect(matches('dash', '<a:prstDash val="dash"/>')).toBe(true);
    expect(matches('dash', '<a:prstDash val="sysDot"/>')).toBe(true);
    expect(matches('dash', '<a:custDash><a:ds d="400" sp="200"/></a:custDash>')).toBe(true);
  });

  it('does not read an absent arrowhead as one', () => {
    expect(matches('arrowEnd', '<a:tailEnd type="none"/>')).toBe(false);
    expect(matches('arrowEnd', '<a:tailEnd type="triangle" w="med" len="med"/>')).toBe(true);
    expect(matches('arrowEnd', '<a:headEnd type="oval"/>')).toBe(true);
  });

  it('looks for p:bgRef, which is PresentationML and has no a: spelling', () => {
    expect(matches('bgRef', '<p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef>')).toBe(true);
    // The mistake this pattern exists to have already made once: `a:bgRef` is
    // not an element, and looking for one reads zero across the whole corpus.
    expect(matches('bgRef', '<a:bgRef idx="1001"/>')).toBe(false);
  });

  it('separates a source crop from a destination inset', () => {
    expect(matches('srcRect', '<a:srcRect l="25000" t="25000"/>')).toBe(true);
    expect(matches('srcRect', '<a:stretch><a:fillRect l="10000"/></a:stretch>')).toBe(false);
  });

  it('does not read a gradient stop list as a pattern fill', () => {
    expect(matches('gradFill', '<a:gradFill rotWithShape="1">')).toBe(true);
    expect(matches('pattFill', '<a:gradFill rotWithShape="1">')).toBe(false);
    expect(matches('pattFill', '<a:pattFill prst="pct5">')).toBe(true);
  });

  it('reads all three shadow spellings and not the glow', () => {
    for (const tag of ['outerShdw', 'innerShdw', 'prstShdw']) {
      expect(matches('shadow', `<a:${tag} blurRad="1"/>`), tag).toBe(true);
    }
    expect(matches('shadow', '<a:glow rad="101600"/>')).toBe(false);
    expect(matches('glow', '<a:glow rad="101600"/>')).toBe(true);
  });
});
