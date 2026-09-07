/**
 * Experiment T7, step 1 - write the probe decks.
 *
 * ```
 * node tools/ground-truth/fonts/substitution/build-deck.ts <work-dir>
 * ```
 *
 * One slide per probe, because an EMF names no shapes and the face is read from
 * the EMF. One package per theme, because a package has one theme.
 */

import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

import { allPackages, type Face, type Probe, type Slot } from './probes.ts';
import { buildPptx, escapeText, textBox, EMU_PER_INCH } from '../../lib/pptx.ts';

const arg = process.argv[2];
if (arg === undefined) throw new Error('usage: build-deck.ts <out-dir>');
const outDir: string = arg;
mkdirSync(outDir, { recursive: true });

const BOX_X = EMU_PER_INCH;
const BOX_Y = EMU_PER_INCH;
const BOX_W = EMU_PER_INCH * 10;
const BOX_H = EMU_PER_INCH * 2;

function attr(name: string, value: string | number | undefined): string {
  if (value === undefined) return '';
  return ` ${name}="${String(value)}"`;
}

/** `CT_TextFont`, with only the hints the probe states. */
function faceXml(tag: string, f: Face): string {
  return (
    `<a:${tag} typeface="${f.typeface}"` +
    attr('panose', f.panose) +
    attr('pitchFamily', f.pitchFamily) +
    attr('charset', f.charset) +
    '/>'
  );
}

const TAG: Readonly<Record<Slot, string>> = {
  latin: 'a:latin',
  ea: 'a:ea',
  cs: 'a:cs',
  sym: 'a:sym',
};

/** `a:rPr` children are a sequence: latin, ea, cs, sym, in that order. */
function runXml(probe: Probe): string {
  const slots: Slot[] = ['latin', 'ea', 'cs', 'sym'];
  const faces = slots
    .map((slot) => {
      const f = probe.faces[slot];
      return f === undefined ? '' : faceXml(TAG[slot].slice(2), f);
    })
    .join('');
  return (
    `<a:r><a:rPr lang="en-US" sz="${String(probe.sz)}" dirty="0">${faces}</a:rPr>` +
    `<a:t>${escapeText(probe.text)}</a:t></a:r>`
  );
}

/** Reject anything XML 1.0 cannot carry; PowerPoint refuses the package silently. */
function assertProbeText(text: string, what: string): void {
  for (const char of text) {
    const code = char.codePointAt(0) ?? 0;
    const legal =
      code === 0x09 || code === 0x0a || code === 0x0d || (code >= 0x20 && code !== 0x7f);
    if (!legal) {
      throw new Error(
        `${what} holds U+${code.toString(16).toUpperCase().padStart(4, '0')}, which XML 1.0 forbids`,
      );
    }
  }
}

interface DeckEntry {
  readonly key: string;
  readonly file: string;
  readonly themeFonts: Record<string, string>;
  readonly probes: readonly {
    readonly id: string;
    readonly kind: string;
    readonly slide: number;
    readonly sz: number;
    readonly text: string;
    readonly faces: Readonly<Partial<Record<Slot, Face>>>;
    readonly control?: string | undefined;
  }[];
}

const decks: DeckEntry[] = [];

for (const pack of allPackages()) {
  const slides = pack.probes.map((probe) => {
    assertProbeText(probe.text, probe.id);
    return textBox(2, probe.id, BOX_X, BOX_Y, BOX_W, BOX_H, runXml(probe));
  });
  const bytes = buildPptx({
    slides,
    fontScheme: {
      majorLatin: pack.themeFonts.majorLatin,
      majorEa: '',
      majorCs: '',
      minorLatin: pack.themeFonts.minorLatin,
      minorEa: pack.themeFonts.minorEa,
      minorCs: pack.themeFonts.minorCs,
    },
  });
  const file = `t7-${pack.key}.pptx`;
  writeFileSync(join(outDir, file), bytes);
  decks.push({
    key: pack.key,
    file,
    themeFonts: { ...pack.themeFonts },
    probes: pack.probes.map((probe, i) => ({
      id: probe.id,
      kind: probe.kind,
      slide: i + 1,
      sz: probe.sz,
      text: probe.text,
      faces: probe.faces,
      control: probe.control,
    })),
  });
  console.log(`${file.padEnd(24)} ${String(pack.probes.length).padStart(3)} probe(s)`);
}

const inputs = {
  experiment: 'T7',
  subPhase: '3.7',
  question: 'which typeface a run is actually drawn in, and how absence is detected',
  decks,
};
writeFileSync(join(outDir, 'substitution-inputs.json'), `${JSON.stringify(inputs, null, 2)}\n`);

const total = decks.reduce((n, d) => n + d.probes.length, 0);
console.log(
  `\n${String(decks.length)} package(s), ${String(total)} probe(s)\n` +
    `wrote ${join(outDir, 'substitution-inputs.json')}`,
);
