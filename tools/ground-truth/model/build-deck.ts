/**
 * Experiment C5, step 1 - write the probe decks.
 *
 * ```
 * node tools/ground-truth/model/build-deck.ts <out-dir>
 * ```
 *
 * One `.pptx` per deck named in `sheets.ts`, plus `sheet-inputs.json` carrying
 * every probe's question, every candidate box and the two colour schemes, so the
 * analysis never has to re-derive what it was asking.
 *
 * There is no read window here and no bitmap geometry, which is the difference
 * from C3 and C4. An inheritance is answered by a *position*, and PowerPoint
 * reports the position of a placeholder that has none of its own directly,
 * through `Shape.Left`/`Top`/`Width`/`Height`. The bitmaps are exported anyway,
 * for a human to look at and for the two probes whose answer really is a colour
 * on the slide rather than a number in the object model.
 */

import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import {
  BG_FILL_STYLES,
  buildSheetPackage,
  EMU_PER_POINT,
  FILL_STYLES,
  LINE_STYLE_WIDTHS,
  SCHEME_ONE,
  SCHEME_TWO,
  SLIDE_HEIGHT_PT,
  SLIDE_WIDTH_PT,
} from '../lib/sheet-pptx.ts';
import { ALL_BOXES, sheetDecks } from './probes.ts';

const outDir = process.argv[2];
if (outDir === undefined)
  throw new Error('usage: tools/ground-truth/model/build-deck.ts <out-dir>');
mkdirSync(outDir, { recursive: true });

const decks = sheetDecks();

const seen = new Set<string>();
for (const deck of decks) {
  for (const probe of deck.probes) {
    // A duplicate id would silently overwrite a finding, so it is a build error
    // rather than something the analysis has to notice.
    if (seen.has(probe.id)) throw new Error(`duplicate probe id: ${probe.id}`);
    seen.add(probe.id);
  }
}

interface DeckEntry {
  readonly deck: string;
  readonly file: string;
  readonly hostile: boolean;
  readonly slides: number;
  readonly masters: number;
  readonly layouts: number;
}

const entries: DeckEntry[] = [];
for (const deck of decks) {
  const file = `sheet-${deck.deck}.pptx`;
  writeFileSync(join(outDir, file), buildSheetPackage(deck.pkg));
  entries.push({
    deck: deck.deck,
    file,
    hostile: deck.hostile,
    slides: deck.pkg.slides.length,
    masters: deck.pkg.masters.length,
    layouts: deck.pkg.layouts.length,
  });
}

writeFileSync(
  join(outDir, 'sheet-inputs.json'),
  JSON.stringify(
    {
      slidePoints: { w: SLIDE_WIDTH_PT, h: SLIDE_HEIGHT_PT },
      emuPerPoint: EMU_PER_POINT,
      schemes: { one: SCHEME_ONE, two: SCHEME_TWO },
      styleMatrix: {
        fillStyleLst: FILL_STYLES,
        bgFillStyleLst: BG_FILL_STYLES,
        lnStyleWidths: LINE_STYLE_WIDTHS,
      },
      boxes: ALL_BOXES,
      decks: entries,
      probes: decks.flatMap((d) =>
        d.probes.map((p) => ({
          id: p.id,
          deck: p.deck,
          kind: p.kind,
          slide: p.slide,
          question: p.question,
          expectBox: p.expectBox ?? null,
          expectFill: p.expectFill ?? null,
        })),
      ),
    },
    null,
    2,
  ),
);

const probes = decks.flatMap((d) => d.probes);
console.log(`${String(probes.length)} probes across ${String(entries.length)} decks`);
const kinds = new Map<string, number>();
for (const p of probes) kinds.set(p.kind, (kinds.get(p.kind) ?? 0) + 1);
for (const [kind, n] of [...kinds].sort((a, b) => a[0].localeCompare(b[0]))) {
  console.log(`  ${kind.padEnd(12)} ${String(n)}`);
}
console.log(`  ${String(entries.filter((e) => e.hostile).length)} hostile package(s)`);
