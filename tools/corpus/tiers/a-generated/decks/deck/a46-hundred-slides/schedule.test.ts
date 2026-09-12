import { describe, expect, it } from 'vitest';

import { a46HundredSlides } from './deck.ts';
import {
  CHAPTERS,
  ROTATION,
  SECTION_HEADERS,
  SLIDE_COUNT,
  kindAt,
  kindCounts,
} from './schedule.ts';

describe('the schedule', () => {
  it('opens with a cover and an agenda, closes with a closing, and heads eight chapters', () => {
    expect(kindAt(0)).toBe('cover');
    expect(kindAt(1)).toBe('agenda');
    expect(kindAt(SLIDE_COUNT - 1)).toBe('closing');
    expect(SECTION_HEADERS).toEqual([2, 14, 26, 38, 50, 62, 74, 86]);
    for (const at of SECTION_HEADERS) expect(kindAt(at)).toBe('section');
    expect(SECTION_HEADERS).toHaveLength(CHAPTERS);
  });

  it('rotates the content kinds one step later in each chapter', () => {
    // Chapter 0 starts the rotation at its head; chapter 1 one step on.
    expect(kindAt(3)).toBe(ROTATION[(3 + 0) % 7]);
    expect(kindAt(15)).toBe(ROTATION[(15 + 1) % 7]);
    expect(kindAt(15)).not.toBe(kindAt(3));
  });

  it('yields the counts the features literal is built from', () => {
    expect(kindCounts()).toEqual({
      cover: 1,
      agenda: 1,
      section: 8,
      bullets: 25,
      twoColumn: 13,
      pictureCaption: 12,
      shapeGrid: 13,
      quote: 13,
      barDiagram: 13,
      closing: 1,
    });
  });

  it('refuses an index outside the hundred', () => {
    expect(() => kindAt(100)).toThrow(RangeError);
    expect(() => kindAt(-1)).toThrow(RangeError);
  });
});

describe('the features literal, re-derived from the schedule', () => {
  const counts = kindCounts();
  /** Kinds bound to a layout with a title placeholder. */
  const titled =
    counts.agenda +
    counts.bullets +
    counts.twoColumn +
    counts.pictureCaption +
    counts.shapeGrid +
    counts.barDiagram;
  /** The master's two placeholders and the three on the layouts. */
  const chassisShapes = 5;
  /** What each kind draws as `p:sp`, its title placeholder aside. */
  const spPerKind = {
    cover: 4,
    agenda: 2,
    section: 2,
    bullets: 2,
    twoColumn: 3,
    pictureCaption: 3,
    shapeGrid: 15,
    quote: 3,
    barDiagram: 26,
    closing: 2,
  } as const;
  const sum = (per: Readonly<Record<string, number>>): number =>
    Object.entries(per).reduce(
      (n, [kind, each]) => n + each * counts[kind as keyof typeof counts],
      0,
    );

  /** Gradient and pattern fills on a shape grid follow `(k + nth) % 4` over eleven cells. */
  const gridFills = (residues: readonly number[]): number => {
    let total = 0;
    for (let nth = 0; nth < counts.shapeGrid; nth++) {
      for (let k = 0; k < 11; k++) if (residues.includes((k + nth) % 4)) total += 1;
    }
    return total;
  };

  it('adds up', () => {
    const f = a46HundredSlides.features as Readonly<Record<string, number>> & {
      connector: number;
      picture: number;
    };
    expect(f['shape']).toBe(chassisShapes + titled + sum(spPerKind));
    // The body placeholders on the agenda and the bullets slides carry no geometry.
    expect(f['presetGeom']).toBe(
      sum(spPerKind) - counts.agenda - counts.bullets + f['connector'] + f['picture'],
    );
    expect(f['placeholder']).toBe(chassisShapes + titled + counts.agenda + counts.bullets);
    expect(f['field']).toBe(titled + counts.quote);
    expect(f['connector']).toBe(counts.shapeGrid + counts.quote + counts.barDiagram);
    expect(f['picture']).toBe(counts.pictureCaption);
    expect(f['group']).toBe(counts.shapeGrid);
    expect(f['shadow']).toBe(counts.shapeGrid);
    expect(f['glow']).toBe(counts.shapeGrid);
    // The theme's two, the section layout's own, the three chapters that state one, every bar,
    // and the grids.
    const themeGradients = 2;
    expect(f['gradientFill']).toBe(
      themeGradients + 1 + 3 + 8 * counts.barDiagram + gridFills([1, 2]),
    );
    expect(f['patternFill']).toBe(gridFills([3]));
  });
});
