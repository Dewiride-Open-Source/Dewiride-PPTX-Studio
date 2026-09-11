/**
 * The one assertion this example exists to make: the packages, installed from
 * the registry rather than linked from the workspace, render a real deck.
 *
 * Run by CI after a lockfile-free `npm install`, so a publish that ships a
 * broken exports map or an unresolvable dependency fails here.
 */

import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';

import { renderDeck } from '@pptx-studio/cli';
import { censusPackage } from '@pptx-studio/census';
import { PartStore } from '@pptx-studio/opc';
import { loadDocument } from '@pptx-studio/model';
import { validatePackage } from '@pptx-studio/validate';
import { openPackage, exportPackage } from '@pptx-studio/writer';

const DECKS = join(import.meta.dirname, 'public', 'decks');

let failures = 0;
const check = (name, ok, detail) => {
  console.log(`${ok ? 'ok  ' : 'FAIL'} ${name}${detail === undefined ? '' : ` - ${detail}`}`);
  if (!ok) failures += 1;
};

for (const file of readdirSync(DECKS)
  .filter((n) => n.endsWith('.pptx'))
  .sort()) {
  const bytes = new Uint8Array(readFileSync(join(DECKS, file)));

  const store = PartStore.open(bytes);
  const document = loadDocument(store);
  check(`${file}: opens`, document.slides.length > 0, `${document.slides.length} slide(s)`);

  const census = censusPackage(bytes);
  check(`${file}: census`, census.parts.length > 0, `${census.parts.length} parts`);

  const report = validatePackage({ store });
  const fatal = report.findings.filter((f) => f.severity === 'fatal');
  check(`${file}: validates`, fatal.length === 0, `${report.findings.length} finding(s)`);

  const rendered = renderDeck(bytes, { width: 1280 });
  // Every slide is well formed, and the deck as a whole draws something. Not a
  // size floor per slide: a deck of the eleven built-in layouts contains the
  // Blank one, and an empty `<svg>` is the right answer for it.
  const wellFormed = rendered.slides.every(
    (s) => s.svg.startsWith('<svg') && s.svg.endsWith('</svg>'),
  );
  const widest = rendered.slides.reduce((most, s) => Math.max(most, s.svg.length), 0);
  check(
    `${file}: renders`,
    wellFormed && widest > 500 && rendered.slides.length === document.slides.length,
    `${rendered.slides.length} svg, ${rendered.fonts.length} typeface(s), ` +
      `${rendered.fonts.filter((f) => f.substituted).length} substituted`,
  );

  const opened = openPackage(bytes);
  const result = exportPackage(opened);
  check(`${file}: round trips`, result.bytes.byteLength > 0, `${result.streamed} streamed`);
}

console.log(failures === 0 ? '\nall green' : `\n${failures} failure(s)`);
process.exitCode = failures === 0 ? 0 : 1;
