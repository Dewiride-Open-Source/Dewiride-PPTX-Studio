/**
 * Experiment B - build a deck that embeds the probe font several times over,
 * each copy wrapped in a differently shaped EOT, and ask PowerPoint which ones
 * it will render.
 *
 * ```
 * node tools/ground-truth/build-font-deck.ts <out-dir>
 * ```
 *
 * The first pass of this experiment used one variant - an uncompressed EOT
 * version 1 - and PowerPoint silently substituted. Silently is the operative
 * word: `Font.Embedded` reported msoTrue, the deck opened without a repair
 * prompt, and the slide showed perfectly good letters. Only the negative
 * control gave it away.
 *
 * So this is a matrix rather than a single test, because "PowerPoint rejected
 * it" is not a finding - "PowerPoint rejected *this and not that*" is. The
 * variables are the ones where our EOT differed from one PowerPoint wrote:
 * the version, whether FamilyName and StyleName carry a terminating NUL inside
 * their counted length, and the charset byte.
 *
 * Each variant gets its own font with its own family name, so PowerPoint cannot
 * satisfy one line from another line's data, and its own row of the same eight
 * characters. A row that renders as a staircase is a variant PowerPoint
 * accepted; a row that renders as letters is one it substituted.
 */

import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { buildProbeFont } from './build-font.ts';
import { readEot, writeEot, EOT_VERSION_1, EOT_VERSION_2_2 } from './eot.ts';
import { readZip } from './zip.ts';
import { buildPptx, escapeText, textBox, SLIDE_WIDTH, type EmbeddedFont } from './pptx.ts';

const outDir = process.argv[2];
if (outDir === undefined) throw new Error('usage: build-font-deck.ts <out-dir>');
mkdirSync(outDir, { recursive: true });

interface Variant {
  readonly family: string;
  readonly label: string;
  readonly version: number;
  readonly nul: boolean;
  readonly charset: number;
}

const VARIANTS: Variant[] = [
  {
    family: 'ProbeAlpha',
    label: 'v1   no-NUL  charset=1',
    version: EOT_VERSION_1,
    nul: false,
    charset: 1,
  },
  {
    family: 'ProbeBravo',
    label: 'v1   NUL     charset=0',
    version: EOT_VERSION_1,
    nul: true,
    charset: 0,
  },
  {
    family: 'ProbeCharlie',
    label: 'v2.2 NUL     charset=0  - PowerPoint shape',
    version: EOT_VERSION_2_2,
    nul: true,
    charset: 0,
  },
  {
    family: 'ProbeDelta',
    label: 'v2.2 no-NUL  charset=0',
    version: EOT_VERSION_2_2,
    nul: false,
    charset: 0,
  },
  {
    family: 'ProbeEcho',
    label: 'v1   no-NUL  charset=0',
    version: EOT_VERSION_1,
    nul: false,
    charset: 0,
  },
];

const ABSENT = 'ProbeAbsent';

const fonts: EmbeddedFont[] = [];
for (const variant of VARIANTS) {
  const font = buildProbeFont(variant.family);
  const eot = writeEot(font.bytes, {
    familyName: font.familyName,
    styleName: font.styleName,
    versionName: font.versionName,
    fullName: font.fullName,
    panose: font.panose,
    charset: variant.charset,
    italic: false,
    weight: font.weight,
    fsType: font.fsType,
    unicodeRange: font.unicodeRange,
    codePageRange: font.codePageRange,
    checkSumAdjustment: font.checkSumAdjustment,
    version: variant.version,
    nulTerminateNames: variant.nul,
  });
  writeFileSync(join(outDir, `${variant.family}.eot`), eot);
  writeFileSync(join(outDir, `${variant.family}.ttf`), font.bytes);
  fonts.push({
    typeface: font.familyName,
    panose: [...font.panose].map((b) => b.toString(16).padStart(2, '0').toUpperCase()).join(''),
    pitchFamily: (2 << 4) | 2,
    charset: 0,
    slots: { regular: eot },
  });
  console.log(
    `${variant.family.padEnd(14)} ${variant.label.padEnd(44)} eot=${String(eot.length)} bytes`,
  );
}

/* -------------------------------------------------------------------------- */
/* the plumbing control                                                       */
/* -------------------------------------------------------------------------- */

// If a PowerPoint-written deck is supplied, lift one of its `.fntdata` parts -
// a real, MicroType-Express-compressed EOT - and embed it here under a family
// name that exists nowhere on this machine, patched into the EOT header so the
// header and the `@typeface` agree.
//
// This is the control for the one thing the variant matrix cannot rule out on
// its own: that the deck we build is wrong rather than the EOT we build. If a
// compressed EOT renders from this deck and none of ours do, the deck is fine
// and the payload is the difference. A row using the original family name is
// included beside it, since that font *is* installed and so fixes what its
// letterforms look like.
interface Borrowed {
  readonly typeface: string;
  readonly original: string;
  readonly font: EmbeddedFont;
}

function borrowFromPowerPoint(deckPath: string): Borrowed | undefined {
  const entries = readZip(new Uint8Array(readFileSync(deckPath)));
  const part = entries.find((e) => e.name === 'ppt/fonts/font1.fntdata');
  if (part === undefined) return undefined;

  const header = readEot(part.bytes);
  const original = header.familyName.replace(/\0+$/, '');
  // Same length, so every size field in the header stays correct: swap the
  // first letter for one that starts no font on Windows.
  const retagged = 'Z' + original.slice(1);
  const bytes = new Uint8Array(part.bytes);

  const utf16 = (text: string): number[] => {
    const out: number[] = [];
    for (const ch of text) out.push(ch.charCodeAt(0) & 0xff, ch.charCodeAt(0) >> 8);
    return out;
  };
  const needle = utf16(original);
  let patched = 0;
  // Only inside the header - never in the payload, which is compressed anyway.
  for (let at = 0x52; at + needle.length <= header.eotSize - header.fontDataSize; at++) {
    if (needle.every((b, i) => bytes[at + i] === b)) {
      bytes[at] = 'Z'.charCodeAt(0);
      patched += 1;
      at += needle.length - 1;
    }
  }
  console.log(
    `\nborrowed ${part.name} from ${deckPath}: ${JSON.stringify(original)} -> ` +
      `${JSON.stringify(retagged)} (${String(patched)} header occurrences patched), ` +
      `${String(part.bytes.length)} bytes, flags=0x${header.flags.toString(16)}`,
  );

  return {
    typeface: retagged,
    original,
    font: {
      typeface: retagged,
      panose: [...header.panose].map((b) => b.toString(16).padStart(2, '0').toUpperCase()).join(''),
      pitchFamily: (2 << 4) | 2,
      charset: 0,
      slots: { regular: bytes },
    },
  };
}

const borrowFrom = process.argv[3];
const borrowed = borrowFrom === undefined ? undefined : borrowFromPowerPoint(borrowFrom);
if (borrowed !== undefined) fonts.push(borrowed.font);

const SAMPLE = 'ABCDEFGH';

function run(text: string, typeface: string, size: number): string {
  return (
    `<a:r><a:rPr lang="en-US" sz="${String(size * 100)}" dirty="0">` +
    '<a:solidFill><a:srgbClr val="000000"/></a:solidFill>' +
    `<a:latin typeface="${typeface}"/><a:cs typeface="${typeface}"/>` +
    `</a:rPr><a:t>${escapeText(text)}</a:t></a:r>`
  );
}

const MARGIN = 400000;
const WIDTH = SLIDE_WIDTH - 2 * MARGIN;
const ROW_H = 720000;
const LABEL_W = 4600000;

// One row per variant, plus the negative control (never embedded) and a
// positive control (Arial) that proves the slide rendered at all.
const rows: { name: string; typeface: string; label: string }[] = [
  ...VARIANTS.map((v) => ({ name: v.family, typeface: v.family, label: v.label })),
  ...(borrowed === undefined
    ? []
    : [
        {
          name: 'borrowed',
          typeface: borrowed.typeface,
          label: 'PowerPoint EOT, retagged ' + borrowed.typeface,
        },
        {
          name: 'installed',
          typeface: borrowed.original,
          label: borrowed.original + ' (installed) - reference shapes',
        },
      ]),
  { name: 'absent', typeface: ABSENT, label: 'not embedded  - must fall back' },
  { name: 'control', typeface: 'Arial', label: 'Arial  - must be letters' },
];

let id = 2;
const body = rows
  .map((row, i) => {
    const y = 250000 + i * ROW_H;
    const label = textBox(
      id++,
      `label-${row.name}`,
      MARGIN,
      y + 180000,
      LABEL_W,
      300000,
      run(row.label, 'Arial', 11),
    );
    const sample = textBox(
      id++,
      row.name,
      MARGIN + LABEL_W,
      y,
      WIDTH - LABEL_W,
      ROW_H,
      run(SAMPLE, row.typeface, 24),
    );
    return label + sample;
  })
  .join('');

const pptx = buildPptx({ slides: [body], fonts });
writeFileSync(join(outDir, 'probe-deck.pptx'), pptx);
console.log(`\nprobe-deck.pptx  ${String(pptx.length)} bytes, ${String(rows.length)} rows`);
