/**
 * Which kind of slide sits at each index of the hundred, as a fixed schedule.
 *
 * Eight chapters of twelve, each opened by a section header, with the content kinds rotating one
 * step later in each chapter so no two chapters read the same. A schedule rather than a draw, so a
 * reader can say which slide is which without running anything.
 */

export const SLIDE_COUNT = 100;

export const CHAPTERS = 8;

/** The first slide of each chapter: 2, 14, 26, ... 86. */
export const SECTION_HEADERS: readonly number[] = Array.from(
  { length: CHAPTERS },
  (_, chapter) => 2 + chapter * 12,
);

export type Kind =
  | 'cover'
  | 'agenda'
  | 'section'
  | 'bullets'
  | 'twoColumn'
  | 'pictureCaption'
  | 'shapeGrid'
  | 'quote'
  | 'barDiagram'
  | 'closing';

/** The content kinds, in the order they rotate; bullets twice, as in any real deck. */
export const ROTATION: readonly Kind[] = [
  'bullets',
  'twoColumn',
  'pictureCaption',
  'shapeGrid',
  'quote',
  'barDiagram',
  'bullets',
];

/** The chapter a slide belongs to, counting the cover and agenda as chapter 0's. */
export function chapterOf(index: number): number {
  return Math.max(0, Math.floor((index - 2) / 12));
}

export function kindAt(index: number): Kind {
  if (index < 0 || index >= SLIDE_COUNT || !Number.isInteger(index)) {
    throw new RangeError(`slide index ${String(index)} is not one of the hundred`);
  }
  if (index === 0) return 'cover';
  if (index === 1) return 'agenda';
  if (index === SLIDE_COUNT - 1) return 'closing';
  if (SECTION_HEADERS.includes(index)) return 'section';
  return ROTATION[(index + chapterOf(index)) % ROTATION.length]!;
}

/** How many slides of each kind the schedule yields. */
export function kindCounts(): Readonly<Record<Kind, number>> {
  const counts: Record<Kind, number> = {
    cover: 0,
    agenda: 0,
    section: 0,
    bullets: 0,
    twoColumn: 0,
    pictureCaption: 0,
    shapeGrid: 0,
    quote: 0,
    barDiagram: 0,
    closing: 0,
  };
  for (let index = 0; index < SLIDE_COUNT; index++) counts[kindAt(index)] += 1;
  return counts;
}
