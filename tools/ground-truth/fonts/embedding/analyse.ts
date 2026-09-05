/**
 * Experiment B, step 3 - did PowerPoint actually render our font?
 *
 * ```
 * node tools/ground-truth/fonts/embedding/analyse.ts <dir>
 * ```
 *
 * Reads the bitmap PowerPoint exported and measures the ink. The probe row must
 * be a strictly increasing staircase of eight bars; the absent row must not be.
 * "PowerPoint opened it without complaining" is not a pass and neither is
 * `Font.Embedded` reporting msoTrue - a substituted font would produce plausible
 * letters and a green-looking report.
 */

import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { readBmp } from '../../lib/bmp.ts';

interface Shape {
  slide: number;
  name: string;
  left: number;
  top: number;
  width: number;
  height: number;
  text: string;
  fontName: string;
}
interface Readback {
  slideWidth: number;
  slideHeight: number;
  exportWidth: number;
  exportHeight: number;
  fonts: { name: string; embeddable: number; embedded: number }[];
  shapes: Shape[];
}

const dir = process.argv[2];
if (dir === undefined)
  throw new Error('usage: tools/ground-truth/fonts/embedding/analyse.ts <dir>');

// PowerShell writes a byte-order mark; `JSON.parse` will not accept one.
// Stripped by code point, not by a literal BOM in a regex - that literal is
// invisible in an editor and does not survive a whitespace tidy-up.
const readback = readFileSync(join(dir, 'com-readback.json'), 'utf8');
const com = JSON.parse(
  readback.charCodeAt(0) === 0xfeff ? readback.slice(1) : readback,
) as Readback;
const bmp = readBmp(new Uint8Array(readFileSync(join(dir, 'slide1.bmp'))));

const scaleX = bmp.width / com.slideWidth;
const scaleY = bmp.height / com.slideHeight;
const dark = (x: number, y: number): boolean => {
  const [r, g, b] = bmp.pixel(x, y);
  return r + g + b < 384;
};

interface Profile {
  readonly name: string;
  readonly font: string;
  readonly inkColumns: number;
  readonly cells: number[];
  readonly staircase: boolean;
}

function profile(shape: Shape, cellCount: number): Profile {
  // A generous band around the shape: autofit may have moved or resized it, and
  // glyphs can overhang.
  const x0 = Math.max(0, Math.floor(shape.left * scaleX) - 4);
  const x1 = Math.min(bmp.width - 1, Math.ceil((shape.left + shape.width) * scaleX) + 4);
  const y0 = Math.max(0, Math.floor(shape.top * scaleY) - 8);
  const y1 = Math.min(bmp.height - 1, Math.ceil((shape.top + shape.height) * scaleY) + 8);

  // Ink extent, so the cells are measured over the text and not the box.
  let inkLeft = -1;
  let inkRight = -1;
  let inkColumns = 0;
  for (let x = x0; x <= x1; x++) {
    let any = false;
    for (let y = y0; y <= y1 && !any; y++) if (dark(x, y)) any = true;
    if (any) {
      if (inkLeft < 0) inkLeft = x;
      inkRight = x;
      inkColumns += 1;
    }
  }
  if (inkLeft < 0) {
    return { name: shape.name, font: shape.fontName, inkColumns: 0, cells: [], staircase: false };
  }

  const cellWidth = (inkRight - inkLeft + 1) / cellCount;
  const cells: number[] = [];
  for (let c = 0; c < cellCount; c++) {
    let top = -1;
    let bottom = -1;
    const cx0 = Math.round(inkLeft + c * cellWidth);
    const cx1 = Math.round(inkLeft + (c + 1) * cellWidth) - 1;
    for (let y = y0; y <= y1; y++) {
      let any = false;
      for (let x = cx0; x <= cx1 && !any; x++) if (dark(x, y)) any = true;
      if (any) {
        if (top < 0) top = y;
        bottom = y;
      }
    }
    cells.push(top < 0 ? 0 : bottom - top + 1);
  }
  const staircase = cells.every((h, i) => h > 0 && (i === 0 || h > cells[i - 1]!));
  return { name: shape.name, font: shape.fontName, inkColumns, cells, staircase };
}

// Every row that is not a label. The deck names each sample shape after the
// variant it exercises, so the report reads as the matrix it is.
const rows = com.shapes
  .filter((s) => !s.name.startsWith('label-') && s.text !== '')
  .map((s) => profile(s, 8));

console.log('='.repeat(78));
console.log('EXPERIMENT B - does PowerPoint render an EOT that we wrote?');
console.log('='.repeat(78));
for (const f of com.fonts) {
  console.log(
    `  Font "${f.name}": embeddable=${f.embeddable === -1 ? 'msoTrue' : 'msoFalse'} ` +
      `embedded=${f.embedded === -1 ? 'msoTrue' : 'msoFalse'}`,
  );
}
console.log(
  '\n  Note: `embedded=msoTrue` means a p:embeddedFont entry matched the typeface.\n' +
    '  It does NOT mean the font data loaded. Every rejected variant reports msoTrue.\n',
);

for (const row of rows) {
  console.log(
    `  ${row.name.padEnd(14)} font=${row.font.padEnd(16)} ink=${String(row.inkColumns).padStart(4)} ` +
      `bars=[${row.cells.join(' ')}]${row.staircase ? '   <- STAIRCASE: rendered' : ''}`,
  );
}

// The negative control: a typeface that is neither installed nor embedded, so
// PowerPoint must substitute. Any row whose ink profile equals this one was
// substituted too, whatever the object model claims.
const absent = rows.find((r) => r.name === 'absent');
const control = rows.find((r) => r.name === 'control');
// `installed` is a reference row in a typeface this machine has - it renders
// from the installed font whatever the embedding does, so it is not a variant
// and a staircase is not what it should produce.
const REFERENCE = new Set(['absent', 'control', 'installed']);
const variants = rows.filter((r) => !REFERENCE.has(r.name));

console.log('');
if (absent === undefined || control === undefined) {
  console.log('  no `absent` / `control` row found - this deck cannot decide anything');
  process.exitCode = 1;
} else {
  const fingerprint = (r: Profile): string => r.cells.join(',');
  let rendered = 0;
  for (const v of variants) {
    const same = fingerprint(v) === fingerprint(absent);
    const verdict = v.staircase
      ? 'RENDERED'
      : same
        ? 'substituted (matches the absent row)'
        : 'substituted';
    if (v.staircase) rendered += 1;
    console.log(`  ${v.staircase ? 'PASS' : 'FAIL'}  ${v.name.padEnd(14)} ${verdict}`);
  }
  console.log(
    `  ${!absent.staircase ? 'PASS' : 'FAIL'}  ${'absent'.padEnd(14)} fell back, as it must`,
  );
  console.log(`  ${!control.staircase ? 'PASS' : 'FAIL'}  ${'control'.padEnd(14)} Arial is Arial`);
  console.log(
    `\n${String(rendered)} of ${String(variants.length)} variants rendered. ` +
      (rendered === 0
        ? 'PowerPoint rendered NONE of the payloads we can construct.'
        : 'At least one payload shape is accepted - see which.'),
  );
  if (rendered === 0) process.exitCode = 1;
}
