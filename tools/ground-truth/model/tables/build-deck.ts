/**
 * Experiment C7, step 1 - write the probe packages.
 *
 * ```
 * node tools/ground-truth/model/tables/build-deck.ts <out-dir>
 * ```
 *
 * One `.pptx` per probe, so a refusal or a repair names its cause, plus `table-inputs.json`
 * carrying every probe's question and the exact `a:tbl` markup it was asked about.
 */

import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

import { buildPptx } from '../../lib/pptx.ts';
import { allProbes, EMU_PER_POINT, FRAME, frameXml, tableXml } from './probes.ts';

const outDir = process.argv[2];
if (outDir === undefined)
  throw new Error('usage: tools/ground-truth/model/tables/build-deck.ts <out-dir>');
mkdirSync(outDir, { recursive: true });

const probes = allProbes();
const decks = probes.map((probe) => {
  const file = `table-${probe.id}.pptx`;
  writeFileSync(join(outDir, file), buildPptx({ slides: [frameXml(probe)] }));
  return {
    deck: probe.id,
    file,
    family: probe.family,
    question: probe.question,
    hostile: probe.hostile === true,
    cols: probe.cols,
    rows: probe.rows.map((row) => ({ h: row.h, cells: row.cells })),
    frame: probe.frame ?? null,
    markup: tableXml(probe),
  };
});

writeFileSync(
  join(outDir, 'table-inputs.json'),
  JSON.stringify({ emuPerPoint: EMU_PER_POINT, frame: FRAME, decks }, null, 2),
);

console.log(`${String(decks.length)} probes, one package each`);
const families = new Map<string, number>();
for (const p of probes) families.set(p.family, (families.get(p.family) ?? 0) + 1);
for (const [family, n] of [...families].sort((a, b) => a[0].localeCompare(b[0]))) {
  console.log(`  ${family.padEnd(10)} ${String(n)}`);
}
console.log(`  ${String(probes.filter((p) => p.hostile === true).length)} hostile`);
