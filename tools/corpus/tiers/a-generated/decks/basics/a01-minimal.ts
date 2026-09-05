import type { ProbeDeck } from '../types.ts';

/**
 * The baseline: one slide, one title, nothing else.
 *
 * Two jobs, and the second is the reason it is deck one rather than an
 * afterthought.
 *
 * It is the smallest thing the chassis can produce, so when `cli bisect`
 * (sub-phase 1.5) reduces a deck PowerPoint repaired, this is the floor it
 * reduces towards. A minimal reproduction is only useful if the minimum is
 * known to be good.
 *
 * And it is the **subtrahend** for every other Tier A deck. The chassis
 * contributes a fixed census of its own - a master with two placeholders, one
 * layout with a third, and the two `a:gradFill` entries the theme's
 * `fillStyleLst` must have for `fillRef/@idx` to resolve. Any other deck's
 * `features` map minus this one is exactly what that deck's probe adds, which
 * is what makes the maps readable at a glance instead of arithmetic.
 */
export const a01Minimal: ProbeDeck = {
  id: 'a01-minimal',
  title: 'PPTX Studio corpus: a01 minimal',
  description:
    'One slide, one title placeholder, nothing optional. The floor for cli bisect, and the ' +
    'baseline every other Tier A deck is measured against: its census is the chassis alone - ' +
    'the master, one layout, and the two theme gradients fillRef/@idx needs.',
  features: {
    shape: 4,
    placeholder: 4,
    gradientFill: 2,
  },
  build: () => ({
    title: 'PPTX Studio corpus: a01 minimal',
    slides: [{ title: 'a01 — minimal', body: '' }],
  }),
};
