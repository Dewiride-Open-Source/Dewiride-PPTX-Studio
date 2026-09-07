/**
 * What Gate 2 claims, as nine questions a slide either answers or does not.
 *
 * The gate's wording is "a real deck rendered with gradients, pattern fills,
 * cropped image fills, dashed strokes, arrowheads, shadows, glows, nested
 * rotated/flipped groups and themed backgrounds". Written out as detections
 * over committed markup it stops being a description and starts being a claim
 * that can fail: on 2026-09-07 two of the nine had no corpus slide at all, and
 * nothing in the repository said so. ADR 0038.
 *
 * Detection is on the slide part's own bytes rather than on our model, because
 * a gate that asked our parser whether a feature was present would be asking
 * the thing under test. The regexes are element tests, not a parser - they are
 * matched against markup this repository generates or PowerPoint wrote, both of
 * which the corpus commits, and a false positive shows up as a slide whose
 * heatmap has nothing to show.
 */

/** A regular expression the page tests against one slide part, by key. */
export interface MarkupPattern {
  readonly key: string;
  /** `RegExp` source, rebuilt in the browser - so it must stay serialisable. */
  readonly pattern: string;
}

export const MARKUP_PATTERNS: readonly MarkupPattern[] = [
  { key: 'gradFill', pattern: '<a:gradFill[ />]' },
  { key: 'pattFill', pattern: '<a:pattFill[ />]' },
  // The crop rectangle on the source, which is not `a:fillRect` on the
  // destination; both are percentages and they are routinely confused.
  { key: 'srcRect', pattern: '<a:srcRect[ />]' },
  { key: 'dash', pattern: '<a:prstDash val="(?!solid")|<a:custDash[ />]' },
  { key: 'arrowEnd', pattern: '<a:(?:head|tail)End[^>]*type="(?!none")' },
  { key: 'shadow', pattern: '<a:(?:outerShdw|innerShdw|prstShdw)[ />]' },
  { key: 'glow', pattern: '<a:glow[ />]' },
  // `p:bgRef`, in PresentationML. There is no `a:bgRef`, and looking for one
  // is how this table read zero themed backgrounds across all 52 decks.
  { key: 'bgRef', pattern: '<p:bgRef[ />]' },
  { key: 'bgPr', pattern: '<p:bgPr[ />]' },
];

/** What the page measured about one slide, and what the features are read off. */
export interface SlideFacts {
  readonly key: string;
  readonly deck: string;
  readonly slide: number;
  /** Keys from `MARKUP_PATTERNS` whose expression matched this slide's part. */
  readonly matched: readonly string[];
  /** How deeply `p:grpSp` nests, counted by tokens rather than by a regex. */
  readonly groupDepth: number;
  /** `p:grpSpPr/a:xfrm` carrying a non-zero `@rot` or a flip. */
  readonly transformedGroups: number;
}

export interface GateFeature {
  readonly key: string;
  readonly label: string;
  readonly holds: (facts: SlideFacts) => boolean;
}

const has =
  (key: string) =>
  (facts: SlideFacts): boolean =>
    facts.matched.includes(key);

export const GATE2_FEATURES: readonly GateFeature[] = [
  { key: 'gradients', label: 'Gradient fills', holds: has('gradFill') },
  { key: 'patterns', label: 'Pattern fills', holds: has('pattFill') },
  { key: 'croppedImages', label: 'Cropped image fills', holds: has('srcRect') },
  { key: 'dashedStrokes', label: 'Dashed strokes', holds: has('dash') },
  { key: 'arrowheads', label: 'Arrowheads', holds: has('arrowEnd') },
  { key: 'shadows', label: 'Shadows', holds: has('shadow') },
  { key: 'glows', label: 'Glows', holds: has('glow') },
  {
    // Both at once, on one slide: a deck full of flat groups and a separate
    // deck full of rotated ones would satisfy neither half of what Gate 2 says.
    key: 'nestedGroups',
    label: 'Nested rotated or mirrored groups',
    holds: (facts) => facts.groupDepth >= 2 && facts.transformedGroups > 0,
  },
  {
    key: 'themedBackgrounds',
    label: 'Themed backgrounds',
    holds: (facts) => has('bgRef')(facts) || has('bgPr')(facts),
  },
];
