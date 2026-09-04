/**
 * Experiment C6, step 1 - write the probe decks.
 *
 * ```
 * node tools/ground-truth/build-transform-deck.ts <out-dir>
 * ```
 *
 * One `.pptx` per deck named in `transforms.ts`, plus `transform-inputs.json`
 * carrying every probe's whole transform chain - the group rectangles, their
 * child coordinate systems, and the leaf's own `a:xfrm`. The analysis computes
 * each candidate model from that rather than comparing against a number typed
 * into the catalogue, so a probe cannot quietly ask one question and check
 * another.
 */

import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

import { buildPptx } from './pptx.ts';
import {
  EMU_PER_POINT,
  EXPORT_HEIGHT,
  EXPORT_WIDTH,
  SLIDE_HEIGHT_PT,
  SLIDE_WIDTH_PT,
  transformDecks,
} from './transforms.ts';

const outDir = process.argv[2];
if (outDir === undefined) throw new Error('usage: build-transform-deck.ts <out-dir>');
mkdirSync(outDir, { recursive: true });

const decks = transformDecks();

const seen = new Set<string>();
for (const deck of decks) {
  for (const probe of deck.probes) {
    if (seen.has(probe.id)) throw new Error(`duplicate probe id: ${probe.id}`);
    seen.add(probe.id);
    for (const sample of probe.samples ?? []) {
      if (seen.has(sample.id)) throw new Error(`duplicate sample id: ${sample.id}`);
      seen.add(sample.id);
      if (sample.x < 0 || sample.y < 0 || sample.x > SLIDE_WIDTH_PT || sample.y > SLIDE_HEIGHT_PT) {
        throw new Error(`sample ${sample.id} is off the slide`);
      }
    }
  }
}

interface DeckEntry {
  readonly deck: string;
  readonly file: string;
  readonly hostile: boolean;
  readonly slides: number;
}

const entries: DeckEntry[] = [];
for (const deck of decks) {
  const file = `transform-${deck.deck}.pptx`;
  writeFileSync(join(outDir, file), deck.bytes?.() ?? buildPptx({ slides: deck.slides }));
  entries.push({
    deck: deck.deck,
    file,
    hostile: deck.hostile,
    slides: deck.bytes === undefined ? deck.slides.length : deck.probes.length,
  });
}

writeFileSync(
  join(outDir, 'transform-inputs.json'),
  JSON.stringify(
    {
      slidePoints: { w: SLIDE_WIDTH_PT, h: SLIDE_HEIGHT_PT },
      emuPerPoint: EMU_PER_POINT,
      exportPixels: { w: EXPORT_WIDTH, h: EXPORT_HEIGHT },
      decks: entries,
      probes: decks.flatMap((d) => d.probes),
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
  console.log(`  ${kind.padEnd(8)} ${String(n)}`);
}
console.log(`  ${String(probes.flatMap((p) => p.samples ?? []).length)} bitmap sample(s)`);
console.log(`  ${String(entries.filter((e) => e.hostile).length)} hostile package(s)`);
