/**
 * Experiment B, the decisive control - put our EOT inside a package PowerPoint
 * itself wrote.
 *
 * ```
 * node tools/ground-truth/graft-font.ts <powerpoint-deck.pptx> <probe.eot> <typeface> <out.pptx>
 * ```
 *
 * The variant matrix showed PowerPoint rendering none of our uncompressed EOTs.
 * Two explanations survive that: the EOT is unacceptable, or the *deck* we built
 * around it is. Nothing measured so far separates them - `Font.Embedded` returns
 * msoTrue either way, and retagging a borrowed compressed EOT does not settle it
 * because PowerPoint may check the name inside the payload.
 *
 * So: take a deck PowerPoint wrote, with PowerPoint's own content types,
 * relationships, presentation part and slide, and change exactly three things -
 * the bytes of one `.fntdata` part, the `@typeface` on the `p:embeddedFont`
 * entry that points at it, and the `a:latin/@typeface` of the runs that used it.
 * Every other byte of the package is Microsoft's.
 *
 * If the probe renders, our EOT is fine and our deck was at fault. If it does
 * not, the EOT is at fault, and it is at fault inside a package that is not ours
 * to blame.
 */

import { readFileSync, writeFileSync } from 'node:fs';
import { readZip, writeZip, type ZipEntry } from './zip.ts';
import { readEot } from './eot.ts';

const [, , deckPath, eotPath, typeface, outPath] = process.argv;
if (
  deckPath === undefined ||
  eotPath === undefined ||
  typeface === undefined ||
  outPath === undefined
) {
  throw new Error('usage: graft-font.ts <deck.pptx> <probe.eot> <typeface> <out.pptx>');
}

const entries = readZip(new Uint8Array(readFileSync(deckPath)));
const eot = new Uint8Array(readFileSync(eotPath));
const decoder = new TextDecoder();
const encoder = new TextEncoder();

const text = (name: string): string => {
  const found = entries.find((e) => e.name === name);
  if (found === undefined) throw new Error(`no ${name} in ${deckPath}`);
  return decoder.decode(found.bytes);
};

// The first embedded font, and the part its regular slot points at.
const presentation = text('ppt/presentation.xml');
const block = /<p:embeddedFont>[\s\S]*?<\/p:embeddedFont>/.exec(presentation);
if (block === null) throw new Error('the deck embeds no fonts');
const original = /<p:font\b[^>]*\btypeface="([^"]*)"/.exec(block[0])?.[1];
const regularId = /<p:regular\b[^>]*\br:id="([^"]*)"/.exec(block[0])?.[1];
if (original === undefined || regularId === undefined) {
  throw new Error('could not read the first embedded font entry');
}

const rels = text('ppt/_rels/presentation.xml.rels');
const target = new RegExp(`<Relationship[^>]*\\bId="${regularId}"[^>]*\\bTarget="([^"]*)"`).exec(
  rels,
)?.[1];
if (target === undefined) throw new Error(`relationship ${regularId} not found`);
const partName = `ppt/${target.replace(/^\.\//, '')}`;

const bare = eot[0] === 0x00 && eot[1] === 0x01 && eot[2] === 0x00 && eot[3] === 0x00;
const header = bare ? undefined : readEot(eot);
console.log(`grafting ${eotPath} over ${partName}`);
console.log(`  replacing typeface ${JSON.stringify(original)} with ${JSON.stringify(typeface)}`);
console.log(
  header === undefined
    ? '  payload is a BARE SFNT, not an EOT - testing whether PowerPoint accepts one'
    : `  our EOT: version=0x${header.version.toString(16).padStart(8, '0')} ` +
        `flags=0x${header.flags.toString(16).padStart(8, '0')} family=${JSON.stringify(header.familyName)}`,
);

let renamedRuns = 0;
let replacedPart = false;

const out: ZipEntry[] = entries.map((entry) => {
  if (entry.name === partName) {
    replacedPart = true;
    return { name: entry.name, bytes: eot };
  }
  if (entry.name === 'ppt/presentation.xml') {
    // Only the one p:font element, and only its @typeface.
    const patched = presentation.replace(block[0], (b) =>
      b.replace(`typeface="${original}"`, `typeface="${typeface}"`),
    );
    return { name: entry.name, bytes: encoder.encode(patched) };
  }
  if (/^ppt\/(slides|slideLayouts|slideMasters)\/[^/]+\.xml$/.test(entry.name)) {
    const xml = decoder.decode(entry.bytes);
    const patched = xml.replace(new RegExp(`typeface="${original}"`, 'g'), () => {
      renamedRuns += 1;
      return `typeface="${typeface}"`;
    });
    if (patched === xml) return entry;
    return { name: entry.name, bytes: encoder.encode(patched) };
  }
  return entry;
});

if (!replacedPart) throw new Error(`part ${partName} not found`);
console.log(`  renamed ${String(renamedRuns)} typeface references in slides/layouts/masters`);
if (renamedRuns === 0) {
  console.log('  WARNING: no run used that typeface, so the deck will not exercise the font');
}

writeFileSync(outPath, writeZip(out));
console.log(`  wrote ${outPath}`);
