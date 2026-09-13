/**
 * Scores rolled up by deck: the grain a systematic defect shows at, since a hundred-slide deck
 * fills any worst-ten list on its own (ADR 0054).
 */

import { agreementBp, type SlideScore } from './score.ts';

export interface DeckScore {
  readonly deck: string;
  readonly slides: number;
  /** Cells that agree over every slide of the deck, as basis points. */
  readonly meanBp: number;
  /** The slide furthest from PowerPoint, and how far. */
  readonly worstKey: string;
  readonly worstBp: number;
}

/** Every deck's agreement over its own cells, worst deck first. */
export function byDeck(
  slides: readonly { readonly key: string; readonly deck: string; readonly score: SlideScore }[],
): readonly DeckScore[] {
  const decks = new Map<
    string,
    { sumD: number; cells: number; count: number; worst: [string, number] }
  >();
  for (const slide of slides) {
    const row = decks.get(slide.deck) ?? { sumD: 0, cells: 0, count: 0, worst: [slide.key, 10000] };
    row.sumD += slide.score.sumD;
    row.cells += slide.score.cells;
    row.count += 1;
    if (slide.score.meanBp < row.worst[1]) row.worst = [slide.key, slide.score.meanBp];
    decks.set(slide.deck, row);
  }
  return [...decks]
    .map(([deck, row]) => ({
      deck,
      slides: row.count,
      meanBp: agreementBp(row.sumD, row.cells),
      worstKey: row.worst[0],
      worstBp: row.worst[1],
    }))
    .sort((a, b) => a.meanBp - b.meanBp || (a.deck < b.deck ? -1 : 1));
}
