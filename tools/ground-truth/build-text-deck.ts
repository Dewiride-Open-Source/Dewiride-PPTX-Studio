/**
 * Experiment T1, step 1 - write the probe decks.
 *
 * ```
 * node tools/ground-truth/build-text-deck.ts <out-dir>
 * ```
 *
 * One `.pptx` per deck named in `text-cascade.ts`, plus `text-inputs.json`
 * carrying every probe's question and the size each level of the cascade
 * declares, so `analyse-text.ts` never has to re-derive what was asked.
 */

import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

import { buildSheetPackage, EMU_PER_POINT } from './sheet-pptx.ts';
import {
  BUCKET_PT,
  LADDER_PT,
  SOURCES,
  textDecks,
  THEME_MAJOR,
  THEME_MINOR,
} from './text-cascade.ts';

const outDir = process.argv[2];
if (outDir === undefined) throw new Error('usage: build-text-deck.ts <out-dir>');
mkdirSync(outDir, { recursive: true });

const decks = textDecks();

// A duplicate id would silently overwrite a finding, so it is a build error
// rather than something the analysis has to notice.
const seen = new Set<string>();
for (const deck of decks) {
  for (const probe of deck.probes) {
    if (seen.has(probe.id)) throw new Error(`duplicate probe id: ${probe.id}`);
    seen.add(probe.id);
  }
}

interface DeckEntry {
  readonly deck: string;
  readonly file: string;
  readonly slides: number;
}

const entries: DeckEntry[] = [];
for (const deck of decks) {
  const file = `text-${deck.deck}.pptx`;
  writeFileSync(join(outDir, file), buildSheetPackage(deck.pkg));
  entries.push({ deck: deck.deck, file, slides: deck.pkg.slides.length });
}

writeFileSync(
  join(outDir, 'text-inputs.json'),
  JSON.stringify(
    {
      emuPerPoint: EMU_PER_POINT,
      sources: SOURCES,
      ladderPt: LADDER_PT,
      bucketPt: BUCKET_PT,
      themeFonts: { major: THEME_MAJOR, minor: THEME_MINOR },
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
  console.log(`  ${kind.padEnd(12)} ${String(n)}`);
}
