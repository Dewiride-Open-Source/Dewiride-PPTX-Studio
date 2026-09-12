/**
 * The words on the hundred slides, drawn deterministically from a seeded generator.
 *
 * The generator is the glibc linear congruential one, seeded once per slide from its index, so
 * every build writes the same bytes and no slide's content depends on the one before it.
 */

/** A generator of numbers in [0, 1): `x = (1103515245 x + 12345) mod 2^31`. */
export function lcg(seed: number): () => number {
  let state = (seed >>> 0) % 2147483648;
  return () => {
    state = (Math.imul(1103515245, state) + 12345) % 2147483648;
    if (state < 0) state += 2147483648;
    return state / 2147483648;
  };
}

/** The seed a slide's content is drawn from. */
export const seedOf = (index: number): number => 46 * 1000 + index;

export function pick<T>(rng: () => number, from: readonly T[]): T {
  return from[Math.floor(rng() * from.length)]!;
}

export function between(rng: () => number, low: number, high: number): number {
  return low + Math.floor(rng() * (high - low + 1));
}

export const DECK_TITLE = 'Northwind Analytics';

export const DECK_SUBTITLE = 'FY26 operating review';

export const CHAPTER_TITLES: readonly string[] = [
  'Overview',
  'Market',
  'Product',
  'Customers',
  'Operations',
  'Finance',
  'People',
  'Roadmap',
];

const SUBJECTS: readonly string[] = [
  'Revenue',
  'Gross margin',
  'Net retention',
  'Pipeline coverage',
  'Onboarding time',
  'Support volume',
  'Cloud spend',
  'Release cadence',
  'Headcount',
  'Churn',
  'Deal size',
  'Uptime',
  'Cycle time',
  'Renewal rate',
  'Backlog',
  'Unit cost',
];

const VERBS: readonly string[] = [
  'grew',
  'held at',
  'moved to',
  'reached',
  'settled at',
  'closed the quarter at',
  'improved to',
  'slipped to',
];

const OBJECTS: readonly string[] = [
  'the plan',
  'last year',
  'the forecast',
  'the target',
  'the prior quarter',
  'the industry median',
  'the board case',
  'the revised model',
];

const TAILS: readonly string[] = [
  'ahead of {o}',
  'in line with {o}',
  'against {o}',
  'a point under {o}',
  'two points over {o}',
  'on the strength of {o}',
  'before the effect of {o}',
  'with {o} still to land',
];

const NAMES: readonly string[] = [
  'A. Okafor',
  'M. Lindqvist',
  'R. Tanaka',
  'S. Mehta',
  'J. Alvarez',
  'P. Novak',
  'L. Bernard',
  'K. Osei',
];

const ROLES: readonly string[] = [
  'Head of Product',
  'VP Sales',
  'Customer Success',
  'Finance',
  'Engineering',
  'Operations',
  'People',
  'Design',
];

/** A figure with one decimal and a unit, the way a review slide quotes one. */
export function figure(rng: () => number): string {
  const kinds = ['%', 'm', 'k', 'x', 'd', 'pts'];
  const unit = pick(rng, kinds);
  const value = between(rng, 1, 980) / 10;
  return `${value.toFixed(1)}${unit}`;
}

export function sentence(rng: () => number): string {
  const tail = pick(rng, TAILS).replace('{o}', pick(rng, OBJECTS));
  return `${pick(rng, SUBJECTS)} ${pick(rng, VERBS)} ${figure(rng)}, ${tail}.`;
}

export function phrase(rng: () => number): string {
  return `${pick(rng, SUBJECTS)} ${pick(rng, VERBS)} ${figure(rng)}`;
}

export function name(rng: () => number): string {
  return pick(rng, NAMES);
}

export function role(rng: () => number): string {
  return pick(rng, ROLES);
}

/** A slide title: the chapter, and what the slide is about within it. */
export function slideTitle(chapter: number, rng: () => number): string {
  const topics = [
    'at a glance',
    'by segment',
    'quarter on quarter',
    'what changed',
    'risks',
    'next steps',
    'by region',
    'the numbers',
  ];
  return `${CHAPTER_TITLES[chapter] ?? 'Appendix'}: ${pick(rng, topics)}`;
}
