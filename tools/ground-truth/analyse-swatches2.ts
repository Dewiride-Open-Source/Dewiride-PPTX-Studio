/**
 * Experiment C2, step 3 - join the readbacks and answer each question.
 *
 * ```
 * node tools/ground-truth/analyse-swatches2.ts <work-dir> [--fixture <path>]
 * ```
 *
 * Reads `swatch2-inputs.json`, `com-readback2.json` and the exported bitmaps.
 * Every block prints the hypothesis it is testing and what a wrong answer would
 * have looked like, because a table of hex values with no claim attached is not
 * evidence of anything.
 */

import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { hex, readBmp } from './bmp.ts';
import { CLR_SCHEME } from './pptx.ts';
import { CROSSED_CLR_MAP, type Base2 } from './swatches2.ts';

interface InputSwatch {
  id: string;
  deck: string;
  slide: number;
  index: number;
  group: string;
  description: string;
  base: Base2;
  transforms: { name: string; val?: number }[];
  xml: string;
}
interface Inputs {
  clrScheme: Record<string, string>;
  decks: { deck: string; file: string; slides: number; swatches: number }[];
  swatches: InputSwatch[];
}
interface ComShape {
  id: string;
  slide: number;
  comRgb: number | null;
  transparency: number | null;
  left: number;
  top: number;
  width: number;
  height: number;
}
interface ComDeck {
  deck: string;
  file: string;
  opened: boolean;
  repaired: boolean | null;
  error: string | null;
  slides: number;
  shapes: ComShape[];
  bitmaps: string[];
}
interface ComReadback {
  slideWidth: number;
  slideHeight: number;
  exportWidth: number;
  exportHeight: number;
  decks: ComDeck[];
}

const dir = process.argv[2];
if (dir === undefined) throw new Error('usage: analyse-swatches2.ts <work-dir> [--fixture <path>]');
const fixtureAt = process.argv.indexOf('--fixture');
const fixturePath = fixtureAt > 0 ? process.argv[fixtureAt + 1] : undefined;

const readJson = <T>(path: string): T => {
  const text = readFileSync(path, 'utf8');
  return JSON.parse(text.charCodeAt(0) === 0xfeff ? text.slice(1) : text) as T;
};

const inputs = readJson<Inputs>(join(dir, 'swatch2-inputs.json'));
const com = readJson<ComReadback>(join(dir, 'com-readback2.json'));

/** COLORREF is 0x00BBGGRR, not RGB. */
function colorRef(value: number): [number, number, number] {
  return [value & 0xff, (value >> 8) & 0xff, (value >> 16) & 0xff];
}

interface Measured {
  input: InputSwatch;
  comHex: string | null;
  bmpHex: string;
  transparency: number | null;
  repaired: boolean;
}

const measured: Measured[] = [];
const deckState = new Map<string, ComDeck>();

for (const deck of com.decks) {
  deckState.set(deck.deck, deck);
  if (!deck.opened) continue;
  const byId = new Map(deck.shapes.map((s) => [s.id, s]));
  const bitmaps = new Map<number, ReturnType<typeof readBmp>>();
  for (let s = 1; s <= deck.slides; s++) {
    bitmaps.set(
      s,
      readBmp(new Uint8Array(readFileSync(join(dir, `${deck.deck}-slide${String(s)}.bmp`)))),
    );
  }
  for (const input of inputs.swatches.filter((x) => x.deck === deck.deck)) {
    const shape = byId.get(input.id);
    if (shape === undefined) continue;
    const bmp = bitmaps.get(shape.slide)!;
    const scaleX = bmp.width / com.slideWidth;
    const scaleY = bmp.height / com.slideHeight;
    const px = Math.min(bmp.width - 1, Math.round((shape.left + shape.width / 2) * scaleX));
    const py = Math.min(bmp.height - 1, Math.round((shape.top + shape.height / 2) * scaleY));
    measured.push({
      input,
      comHex: shape.comRgb === null ? null : hex(colorRef(shape.comRgb)),
      bmpHex: hex(bmp.pixel(px, py)),
      transparency: shape.transparency,
      repaired: deck.repaired === true,
    });
  }
}

const group = (name: string): Measured[] => measured.filter((m) => m.input.group === name);
const rule = (holds: boolean): string => (holds ? 'HOLDS' : 'FAILS');

const line = (n = 78): string => '='.repeat(n);

/* -------------------------------------------------------------------------- */

console.log(line());
console.log('0. DID POWERPOINT ACCEPT EACH PACKAGE?');
console.log(line());
console.log('  A repaired deck is not a measurement - PowerPoint may have rewritten');
console.log('  anything in it. One question per package, so a refusal names its cause.');
for (const deck of com.decks) {
  const state = !deck.opened ? 'REFUSED ' : deck.repaired ? 'REPAIRED' : 'ok      ';
  console.log(
    `  ${state} ${deck.deck.padEnd(18)} ${String(deck.shapes.length).padStart(3)} shape(s)` +
      (deck.error === null ? '' : `   ${deck.error.split('(')[0]!.trim()}`),
  );
}

console.log('\n' + line());
console.log('1. DO THE OBJECT MODEL AND THE PAINTED PIXELS AGREE?');
console.log(line());
{
  const opaque = measured.filter((m) => m.comHex !== null && (m.transparency ?? 0) === 0);
  const disagree = opaque.filter((m) => m.comHex !== m.bmpHex);
  console.log(
    `  ${String(opaque.length - disagree.length)}/${String(opaque.length)} opaque swatches agree exactly`,
  );
  for (const m of disagree.slice(0, 40)) {
    console.log(`    ${m.input.description.padEnd(46)} com=${m.comHex!} bmp=${m.bmpHex}`);
  }
  if (disagree.length > 40) console.log(`    ... and ${String(disagree.length - 40)} more`);
}

/* -------------------------------------------------------------------------- */

const fromLin = (c: number): number =>
  c <= 0.0031308 ? c * 12.92 : 1.055 * Math.pow(c, 1 / 2.4) - 0.055;
const to8 = (c: number): number => Math.max(0, Math.min(255, Math.ceil(c * 255 - 0.5)));
const asHex = (rgb: readonly [number, number, number]): string =>
  rgb
    .map((c) => to8(c).toString(16).padStart(2, '0'))
    .join('')
    .toUpperCase();

console.log('\n' + line());
console.log('2. a:scrgbClr - sRGB PERCENTAGES, OR LINEAR LIGHT?');
console.log(line());
console.log('  "scRGB" names a linear-light space. If the attributes are read that way,');
console.log('  50% is BCBCBC; if they are plain sRGB percentages, 50% is 808080.');
{
  let srgbExact = 0;
  let linearExact = 0;
  let n = 0;
  for (const m of [...group('scrgb-grey'), ...group('scrgb'), ...group('scrgb-range')]) {
    if (m.input.base.kind !== 'scrgb') continue;
    const raw: [number, number, number] = [
      m.input.base.r / 100000,
      m.input.base.g / 100000,
      m.input.base.b / 100000,
    ];
    const clamped = raw.map((c) => Math.max(0, Math.min(1, c))) as [number, number, number];
    const asSrgb = asHex(clamped);
    const asLinear = asHex(clamped.map(fromLin) as [number, number, number]);
    if (asSrgb === m.bmpHex) srgbExact += 1;
    if (asLinear === m.bmpHex) linearExact += 1;
    n += 1;
    console.log(
      `  ${m.input.description.padEnd(34)} actual=${m.bmpHex}  sRGB%=${asSrgb}  linear=${asLinear}`,
    );
  }
  console.log(`  sRGB percentages exact ${String(srgbExact)}/${String(n)}`);
  console.log(`  linear light     exact ${String(linearExact)}/${String(n)}`);
}

console.log('\n' + line());
console.log('3. a:hslClr - IS IT THE SAME BI-HEXCONE THE lum/sat TRANSFORMS USE?');
console.log(line());
{
  const hslToRgb = (h: number, s: number, l: number): [number, number, number] => {
    if (s === 0) return [l, l, l];
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    const p = 2 * l - q;
    const f = (t: number): number => {
      let x = t;
      if (x < 0) x += 1;
      if (x > 1) x -= 1;
      if (x < 1 / 6) return p + (q - p) * 6 * x;
      if (x < 1 / 2) return q;
      if (x < 2 / 3) return p + (q - p) * (2 / 3 - x) * 6;
      return p;
    };
    const k = h / 360;
    return [f(k + 1 / 3), f(k), f(k - 1 / 3)];
  };
  let exact = 0;
  let n = 0;
  for (const m of [
    ...group('hsl-hue'),
    ...group('hsl-sat'),
    ...group('hsl-lum'),
    ...group('hsl-range'),
  ]) {
    if (m.input.base.kind !== 'hsl') continue;
    const predicted = asHex(
      hslToRgb(m.input.base.hue / 60000, m.input.base.sat / 100000, m.input.base.lum / 100000),
    );
    const ok = predicted === m.bmpHex;
    if (ok) exact += 1;
    n += 1;
    console.log(
      `  ${ok ? 'ok  ' : 'DIFF'} ${m.input.description.padEnd(34)} actual=${m.bmpHex} predicted=${predicted}` +
        (m.repaired ? '   (deck was REPAIRED)' : ''),
    );
  }
  console.log(`  ${String(exact)}/${String(n)} exact`);
}

console.log('\n' + line());
console.log('4. a:prstClr - THE TABLE, AND WHETHER IT IS THE CSS ONE');
console.log(line());
{
  const rows = [...group('prst'), ...group('prst-grey')];
  console.log(`  ${String(rows.length)} names, all accepted without repair.`);
  for (const m of rows) {
    if (m.input.base.kind !== 'prst') continue;
    console.log(`    ${m.input.base.value.padEnd(20)} ${m.bmpHex}`);
  }
}

console.log('\n' + line());
console.log('5. a:sysClr - THE MACHINE VALUES, AND WHETHER @lastClr WINS');
console.log(line());
for (const m of group('sys')) {
  if (m.input.base.kind !== 'sys') continue;
  console.log(`    ${m.input.base.value.padEnd(24)} ${m.bmpHex}`);
}
console.log('  with a @lastClr that no system theme would produce:');
for (const m of group('sys-lastclr')) {
  if (m.input.base.kind !== 'sys') continue;
  const wins = m.bmpHex === m.input.base.lastClr;
  console.log(
    `    ${rule(wins).padEnd(6)} ${m.input.description.padEnd(28)} actual=${m.bmpHex} lastClr=${String(m.input.base.lastClr)}`,
  );
}

console.log('\n' + line());
console.log('6. ALPHA');
console.log(line());
console.log('  Transparency comes from the object model. A pixel cannot answer this:');
console.log('  a translucent swatch composites with the swatch behind it.');
for (const m of [...group('alpha'), ...group('alpha-chain'), ...group('alpha-mixed')]) {
  console.log(
    `    ${m.input.description.padEnd(56)} rgb=${m.bmpHex} transparency=${m.transparency === null ? 'null' : m.transparency.toFixed(4)}`,
  );
}

console.log('\n' + line());
console.log('7. THE CLAMP BOUNDARY');
console.log(line());
console.log('  satMod 300% overflows saturation; satMod 33.333% brings it back. If');
console.log('  nothing clamps in between, row 2 returns to the base. Row 3 interposes a');
console.log('  gamma round trip and row 4 an alpha, to see whether either forces one.');
for (const m of group('clamp')) {
  console.log(`    ${m.input.description.padEnd(64)} ${m.bmpHex}`);
}

console.log('\n' + line());
console.log('8. THE ABSOLUTE CHANNEL SETTERS, AND THE EDGES');
console.log(line());
for (const m of [...group('channel-set'), ...group('edge'), ...group('identity')]) {
  console.log(`    ${m.input.description.padEnd(46)} ${m.bmpHex}`);
}

console.log('\n' + line());
console.log('9. A CROSSED clrMap');
console.log(line());
console.log('  bg1->dk1, tx1->lt1, bg2->dk2, tx2->lt2, accent1->accent2, accent2->accent1.');
console.log('  dk1/lt1/dk2/lt2 name theme slots and must IGNORE the map.');
{
  const slot = (name: string): string => CLR_SCHEME[name as keyof typeof CLR_SCHEME] ?? '??????';
  for (const m of group('clrmap')) {
    if (m.input.base.kind !== 'scheme') continue;
    const name = m.input.base.value;
    const bypass = slot(name);
    const mapped = slot(CROSSED_CLR_MAP[name as keyof typeof CROSSED_CLR_MAP] ?? name);
    const verdict =
      m.bmpHex === bypass && m.bmpHex === mapped
        ? 'both agree'
        : m.bmpHex === bypass
          ? 'BYPASSES the map'
          : m.bmpHex === mapped
            ? 'FOLLOWS the map'
            : 'NEITHER';
    console.log(
      `    ${name.padEnd(10)} actual=${m.bmpHex}  bypass=${bypass} mapped=${mapped}   ${verdict}`,
    );
  }
}

console.log('\n' + line());
console.log('10. ST_Percentage WRITTEN THE STRICT WAY, IN A TRANSITIONAL PART');
console.log(line());
for (const m of [...group('percent-literal'), ...group('percent-control')]) {
  console.log(`    ${m.input.description.padEnd(46)} ${m.bmpHex}`);
}

/* -------------------------------------------------------------------------- */

if (fixturePath !== undefined) {
  const fixture = {
    $comment:
      'Ground truth for the DrawingML colour bases, alpha, the clrMap and the percentage ' +
      'grammar, measured in Microsoft PowerPoint. The companion to color-transforms.json, ' +
      'which covers the transforms. See docs/adr/0021-colour.md. `expected` is the RGB ' +
      'PowerPoint painted, sampled from its own bitmap export at the centre of each swatch; ' +
      '`transparency` is Shape.Fill.Transparency from the object model.',
    source: {
      application: 'Microsoft PowerPoint',
      generatedBy: 'tools/ground-truth/build-swatch-deck2.ts + analyse-swatches2.ts',
    },
    clrScheme: inputs.clrScheme,
    crossedClrMap: CROSSED_CLR_MAP,
    decks: com.decks.map((d) => ({
      deck: d.deck,
      opened: d.opened,
      repaired: d.repaired,
      error: d.error,
    })),
    swatches: measured.map((m) => ({
      id: m.input.id,
      group: m.input.group,
      description: m.input.description,
      base: m.input.base,
      transforms: m.input.transforms,
      xml: m.input.xml,
      expected: m.bmpHex,
      objectModel: m.comHex,
      transparency: m.transparency,
      fromRepairedDeck: m.repaired,
    })),
  };
  writeFileSync(fixturePath, JSON.stringify(fixture, null, 2) + '\n');
  console.log(`\nwrote ${fixturePath}  (${String(measured.length)} swatches)`);
}
