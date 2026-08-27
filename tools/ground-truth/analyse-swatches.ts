/**
 * Experiment C, step 3 - join the readbacks, decide which colour model
 * PowerPoint implements, and write the fixture.
 *
 * ```
 * node tools/ground-truth/analyse-swatches.ts <work-dir> [--fixture <path>]
 * ```
 *
 * Reads `swatch-inputs.json`, `com-readback.json` and `slideN.bmp` from the work
 * directory. Reports:
 *
 *   1. whether the object model and the painted pixels agree - if they do not,
 *      the pixels win and the object model cannot be used as an oracle;
 *   2. for every candidate model, the worst and mean error against PowerPoint,
 *      per transform group;
 *   3. the fixture: input markup in, expected RGB out.
 */

import { readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { hex, readBmp } from './bmp.ts';
import {
  BLEND_MODELS,
  HSL_MODELS,
  parseHex,
  to8,
  toHex,
  type Channel,
  type HslOp,
} from './color-models.ts';

interface InputSwatch {
  id: string;
  slide: number;
  index: number;
  group: string;
  description: string;
  base: { kind: 'srgb' | 'scheme'; value: string };
  transforms: { name: string; val?: number }[];
  xml: string;
}
interface Inputs {
  clrScheme: Record<string, string>;
  grid: { cols: number; rows: number; perSlide: number };
  swatches: InputSwatch[];
}
interface ComShape {
  id: string;
  slide: number;
  comRgb: number;
  left: number;
  top: number;
  width: number;
  height: number;
  transparency: number;
}
interface ComReadback {
  slideCount: number;
  slideWidth: number;
  slideHeight: number;
  shapes: ComShape[];
}

const dir = process.argv[2];
if (dir === undefined) throw new Error('usage: analyse-swatches.ts <work-dir> [--fixture <path>]');
const fixtureAt = process.argv.indexOf('--fixture');
const fixturePath = fixtureAt > 0 ? process.argv[fixtureAt + 1] : undefined;

// PowerShell 5.1's `Set-Content -Encoding utf8` writes a byte-order mark, and
// `JSON.parse` will not accept one. Stripped by code point rather than by a
// literal in a regex, because a literal BOM in source is invisible and does not
// survive the first person who tidies the whitespace in this file.
const readJson = <T>(path: string): T => {
  const text = readFileSync(path, 'utf8');
  return JSON.parse(text.charCodeAt(0) === 0xfeff ? text.slice(1) : text) as T;
};

const inputs = readJson<Inputs>(join(dir, 'swatch-inputs.json'));
const com = readJson<ComReadback>(join(dir, 'com-readback.json'));

/** COLORREF is 0x00BBGGRR, not RGB. Getting this backwards is silently plausible. */
function colorRef(value: number): [number, number, number] {
  return [value & 0xff, (value >> 8) & 0xff, (value >> 16) & 0xff];
}

const bitmaps = new Map<number, ReturnType<typeof readBmp>>();
for (let s = 1; s <= com.slideCount; s++) {
  bitmaps.set(s, readBmp(new Uint8Array(readFileSync(join(dir, `slide${String(s)}.bmp`)))));
}

const byId = new Map(com.shapes.map((s) => [s.id, s]));

interface Measured {
  input: InputSwatch;
  comHex: string;
  bmpHex: string;
  transparency: number;
}

const measured: Measured[] = [];
const missing: string[] = [];

for (const input of inputs.swatches) {
  const shape = byId.get(input.id);
  if (shape === undefined) {
    missing.push(input.id);
    continue;
  }
  const bmp = bitmaps.get(shape.slide)!;
  // Points to pixels: the export was 1920x1080 over a 960x540pt slide.
  const scaleX = bmp.width / com.slideWidth;
  const scaleY = bmp.height / com.slideHeight;
  const px = Math.min(bmp.width - 1, Math.round((shape.left + shape.width / 2) * scaleX));
  const py = Math.min(bmp.height - 1, Math.round((shape.top + shape.height / 2) * scaleY));
  measured.push({
    input,
    comHex: hex(colorRef(shape.comRgb)),
    bmpHex: hex(bmp.pixel(px, py)),
    transparency: shape.transparency,
  });
}

console.log('='.repeat(78));
console.log('1. DO THE OBJECT MODEL AND THE PAINTED PIXELS AGREE?');
console.log('='.repeat(78));
if (missing.length > 0) console.log(`MISSING SHAPES: ${missing.join(' ')}`);
const disagree = measured.filter((m) => m.comHex !== m.bmpHex);
console.log(
  `${String(measured.length - disagree.length)}/${String(measured.length)} agree exactly`,
);
for (const m of disagree.slice(0, 30)) {
  console.log(`  ${m.input.id} ${m.input.description.padEnd(44)} com=${m.comHex} bmp=${m.bmpHex}`);
}
if (disagree.length > 30) console.log(`  ... and ${String(disagree.length - 30)} more`);
console.log(
  '\nThe painted pixels are the ground truth: they are what a user sees, and what a\n' +
    'renderer has to match. Where the two differ the fixture records the bitmap.',
);

/* -------------------------------------------------------------------------- */

const truth = (m: Measured): [Channel, Channel, Channel] => {
  const value = m.bmpHex;
  return parseHex(value);
};

function baseRgb(input: InputSwatch): [Channel, Channel, Channel] {
  if (input.base.kind === 'srgb') return parseHex(input.base.value);
  // The clrMap in our master is the identity one, so bg1 -> lt1 and tx1 -> dk1.
  const map: Record<string, string> = {
    bg1: 'lt1',
    tx1: 'dk1',
    bg2: 'lt2',
    tx2: 'dk2',
  };
  const slot = map[input.base.value] ?? input.base.value;
  const value = inputs.clrScheme[slot];
  if (value === undefined) throw new Error(`no scheme colour ${slot}`);
  return parseHex(value);
}

const delta = (
  a: readonly [number, number, number],
  b: readonly [number, number, number],
): number => Math.max(Math.abs(a[0] - b[0]), Math.abs(a[1] - b[1]), Math.abs(a[2] - b[2]));

const as8 = (rgb: readonly [Channel, Channel, Channel]): [number, number, number] => [
  to8(rgb[0]),
  to8(rgb[1]),
  to8(rgb[2]),
];

console.log('\n' + '='.repeat(78));
console.log('2. WHICH MODEL IS POWERPOINT USING?');
console.log('='.repeat(78));

interface Score {
  model: string;
  note: string;
  worst: number;
  mean: number;
  exact: number;
  n: number;
  worstCase: string;
}

function score(
  label: string,
  rows: Measured[],
  models: readonly { id: string; note: string }[],
  predict: (modelId: string, m: Measured) => [Channel, Channel, Channel] | undefined,
): void {
  if (rows.length === 0) return;
  const scores: Score[] = [];
  for (const model of models) {
    let worst = 0;
    let sum = 0;
    let exact = 0;
    let n = 0;
    let worstCase = '';
    for (const m of rows) {
      const predicted = predict(model.id, m);
      if (predicted === undefined) continue;
      const d = delta(as8(predicted), as8(truth(m)));
      if (d > worst) {
        worst = d;
        worstCase = `${m.input.description} -> ${toHex(predicted)} vs ${m.bmpHex}`;
      }
      if (d === 0) exact += 1;
      sum += d;
      n += 1;
    }
    if (n > 0) {
      scores.push({ model: model.id, note: model.note, worst, mean: sum / n, exact, n, worstCase });
    }
  }
  scores.sort((a, b) => a.worst - b.worst || a.mean - b.mean);
  console.log(`\n--- ${label}  (${String(rows.length)} swatches) ---`);
  for (const s of scores) {
    console.log(
      `  ${s.model.padEnd(16)} worst=${String(s.worst).padStart(3)}/255  ` +
        `mean=${s.mean.toFixed(2).padStart(6)}  exact=${String(s.exact)}/${String(s.n)}   ${s.note}`,
    );
  }
  const best = scores[0];
  if (best === undefined || best.worst === 0) return;

  // Every case the winning model does not get exactly right. A model that is
  // "close" is a model we have not understood yet, and the residual is where
  // the remaining rule is hiding.
  console.log(`  residuals for ${best.model}:`);
  for (const m of rows) {
    const predicted = predict(best.model, m);
    if (predicted === undefined) continue;
    const d = delta(as8(predicted), as8(truth(m)));
    if (d === 0) continue;
    console.log(
      `    d=${String(d).padStart(3)}  ${m.input.description.padEnd(40)} ` +
        `predicted=${toHex(predicted)} actual=${m.bmpHex}`,
    );
  }
}

const group = (...names: string[]): Measured[] =>
  measured.filter((m) => names.includes(m.input.group));

// tint and shade
score('tint', group('tint-sweep', 'tint-fine', 'tint-of-black'), BLEND_MODELS, (id, m) => {
  const model = BLEND_MODELS.find((x) => x.id === id)!;
  const t = m.input.transforms[0];
  if (t?.name !== 'tint' || t.val === undefined) return undefined;
  const p = t.val / 100000;
  const rgb = baseRgb(m.input);
  return [model.tint(rgb[0], p), model.tint(rgb[1], p), model.tint(rgb[2], p)];
});

score('shade', group('shade-sweep', 'shade-fine', 'shade-of-white'), BLEND_MODELS, (id, m) => {
  const model = BLEND_MODELS.find((x) => x.id === id)!;
  const t = m.input.transforms[0];
  if (t?.name !== 'shade' || t.val === undefined) return undefined;
  const p = t.val / 100000;
  const rgb = baseRgb(m.input);
  return [model.shade(rgb[0], p), model.shade(rgb[1], p), model.shade(rgb[2], p)];
});

// HSL transforms
const HSL_NAMES = new Set([
  'lumMod',
  'lumOff',
  'lum',
  'satMod',
  'satOff',
  'sat',
  'hueMod',
  'hueOff',
  'hue',
]);
const hslRows = measured.filter(
  (m) => m.input.transforms.length > 0 && m.input.transforms.every((t) => HSL_NAMES.has(t.name)),
);
score('HSL transforms (lumMod / lumOff / sat / hue)', hslRows, HSL_MODELS, (id, m) => {
  const model = HSL_MODELS.find((x) => x.id === id)!;
  const ops: HslOp[] = m.input.transforms.map((t) => {
    const isHue = t.name.startsWith('hue');
    const p = isHue && t.name !== 'hueMod' ? (t.val ?? 0) / 60000 : (t.val ?? 0) / 100000;
    return { kind: t.name, p } as HslOp;
  });
  return model.apply(baseRgb(m.input), ops);
});

console.log('\n' + '='.repeat(78));
console.log('3. ORDER DEPENDENCE');
console.log('='.repeat(78));
for (const m of group('order')) {
  console.log(`  ${m.input.description.padEnd(52)} -> ${m.bmpHex}`);
}

console.log('\n' + '='.repeat(78));
console.log('4. THEME RESOLUTION');
console.log('='.repeat(78));
for (const m of group('scheme')) {
  const expected = baseRgb(m.input);
  const ok = toHex(expected) === m.bmpHex ? 'ok  ' : 'DIFF';
  console.log(
    `  ${ok} ${m.input.description.padEnd(24)} expected=${toHex(expected)} actual=${m.bmpHex}`,
  );
}

console.log('\n' + '='.repeat(78));
console.log('5. IDENTITY CONTROL - no transform at all');
console.log('='.repeat(78));
for (const m of group('identity')) {
  const ok = m.input.base.value === m.bmpHex ? 'ok  ' : 'DIFF';
  console.log(`  ${ok} ${m.input.description.padEnd(24)} actual=${m.bmpHex} com=${m.comHex}`);
}

console.log('\n' + '='.repeat(78));
console.log('6. THE REST');
console.log('='.repeat(78));
for (const m of group('flag', 'channelMod', 'channelOff')) {
  console.log(`  ${m.input.description.padEnd(30)} -> ${m.bmpHex}`);
}

console.log('\n' + '='.repeat(78));
console.log('7. DERIVED RULES - each one a hypothesis, checked against every swatch');
console.log('='.repeat(78));

const toLin = (c: Channel): Channel =>
  c <= 0.04045 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);
const fromLin = (c: Channel): Channel =>
  c <= 0.0031308 ? c * 12.92 : 1.055 * Math.pow(c, 1 / 2.4) - 0.055;
const clamp01 = (c: number): number => Math.min(1, Math.max(0, c));

interface Rule {
  readonly name: string;
  readonly claim: string;
  matches(m: Measured): boolean;
  predict(m: Measured): [Channel, Channel, Channel];
}

const perChannel =
  (f: (c: Channel) => Channel) =>
  (m: Measured): [Channel, Channel, Channel] => {
    const [r, g, b] = baseRgb(m.input);
    return [f(r), f(g), f(b)];
  };

const only = (name: string) => (m: Measured) =>
  m.input.transforms.length === 1 && m.input.transforms[0]!.name === name;

const RULES: Rule[] = [
  {
    name: 'inv',
    claim: 'a:inv complements each channel in LINEARISED sRGB, not in sRGB',
    matches: only('inv'),
    predict: perChannel((c) => fromLin(clamp01(1 - toLin(c)))),
  },
  {
    name: 'gamma',
    claim: 'a:gamma is the sRGB de-linearisation (encode a linear value to sRGB)',
    matches: only('gamma'),
    predict: perChannel((c) => fromLin(c)),
  },
  {
    name: 'invGamma',
    claim: 'a:invGamma is the sRGB linearisation, the exact inverse of a:gamma',
    matches: only('invGamma'),
    predict: perChannel((c) => toLin(c)),
  },
  {
    name: 'gray',
    claim: 'a:gray is the Rec.709 luma of the sRGB values, with NO linearisation',
    matches: only('gray'),
    predict: (m) => {
      const [r, g, b] = baseRgb(m.input);
      const y = 0.2126 * r + 0.7152 * g + 0.0722 * b;
      return [y, y, y];
    },
  },
  {
    name: 'channelMod',
    claim: 'a:redMod/greenMod/blueMod scale the channel in linearised sRGB',
    matches: (m) => m.input.group === 'channelMod',
    predict: (m) => {
      const rgb = [...baseRgb(m.input)] as [Channel, Channel, Channel];
      const t = m.input.transforms[0]!;
      const i = t.name.startsWith('red') ? 0 : t.name.startsWith('green') ? 1 : 2;
      rgb[i] = fromLin(clamp01(toLin(rgb[i]) * ((t.val ?? 0) / 100000)));
      return rgb;
    },
  },
  {
    name: 'channelOff',
    claim: 'a:redOff/greenOff/blueOff add to the channel in linearised sRGB',
    matches: (m) => m.input.group === 'channelOff',
    predict: (m) => {
      const rgb = [...baseRgb(m.input)] as [Channel, Channel, Channel];
      const t = m.input.transforms[0]!;
      const i = t.name.startsWith('red') ? 0 : t.name.startsWith('green') ? 1 : 2;
      rgb[i] = fromLin(clamp01(toLin(rgb[i]) + (t.val ?? 0) / 100000));
      return rgb;
    },
  },
];

for (const rule of RULES) {
  const rows = measured.filter((m) => rule.matches(m));
  let worst = 0;
  const bad: string[] = [];
  for (const m of rows) {
    const d = delta(as8(rule.predict(m)), as8(truth(m)));
    if (d > worst) worst = d;
    if (d !== 0)
      bad.push(`${m.input.description} predicted=${toHex(rule.predict(m))} actual=${m.bmpHex}`);
  }
  console.log(
    `  ${worst === 0 ? 'HOLDS ' : 'FAILS '} ${rule.name.padEnd(12)} ` +
      `${String(rows.length - bad.length)}/${String(rows.length)} exact, worst=${String(worst)}   ${rule.claim}`,
  );
  for (const line of bad) console.log(`      ${line}`);
}

console.log('\n' + '='.repeat(78));
console.log('8. HOW DOES POWERPOINT GET FROM A REAL NUMBER TO A BYTE?');
console.log('='.repeat(78));

// The residuals that survive a correct model are all exactly x.5, which is the
// one place rounding modes disagree. Worth settling, because it is the
// difference between a fixture that passes and one that is off by one.
{
  const modes: [string, (n: number) => number][] = [
    ['round', (n) => Math.round(n)], // half away from zero - JavaScript's default
    ['half-down', (n) => Math.ceil(n - 0.5)],
    [
      'half-even',
      (n) => {
        const f = Math.floor(n);
        if (n - f !== 0.5) return Math.round(n);
        return f % 2 === 0 ? f : f + 1;
      },
    ],
    ['floor', (n) => Math.floor(n)],
    ['ceil', (n) => Math.ceil(n)],
    ['trunc', (n) => Math.trunc(n)],
  ];
  const blend = BLEND_MODELS.find((x) => x.id === 'linear-srgb')!;
  const hsl = HSL_MODELS.find((x) => x.id === 'hsl-srgb')!;

  const cases: { predicted: [Channel, Channel, Channel]; actual: string }[] = [];
  for (const m of measured) {
    const ts = m.input.transforms;
    let predicted: [Channel, Channel, Channel] | undefined;
    if (ts.length === 1 && (ts[0]!.name === 'tint' || ts[0]!.name === 'shade')) {
      const p = (ts[0]!.val ?? 0) / 100000;
      const rgb = baseRgb(m.input);
      const f = ts[0]!.name === 'tint' ? blend.tint : blend.shade;
      predicted = [f(rgb[0], p), f(rgb[1], p), f(rgb[2], p)];
    } else if (ts.length > 0 && ts.every((t) => HSL_NAMES.has(t.name))) {
      predicted = hsl.apply(
        baseRgb(m.input),
        ts.map((t) => {
          const isHue = t.name.startsWith('hue');
          const p = isHue && t.name !== 'hueMod' ? (t.val ?? 0) / 60000 : (t.val ?? 0) / 100000;
          return { kind: t.name, p } as HslOp;
        }),
      );
    }
    if (predicted !== undefined) cases.push({ predicted, actual: m.bmpHex });
  }

  for (const [name, fn] of modes) {
    let exact = 0;
    let worst = 0;
    for (const c of cases) {
      const bytes = c.predicted.map((x) => Math.max(0, Math.min(255, fn(x * 255)))) as [
        number,
        number,
        number,
      ];
      const actual = [
        parseInt(c.actual.slice(0, 2), 16),
        parseInt(c.actual.slice(2, 4), 16),
        parseInt(c.actual.slice(4, 6), 16),
      ] as [number, number, number];
      const d = delta(bytes, actual);
      if (d === 0) exact += 1;
      if (d > worst) worst = d;
    }
    console.log(
      `  ${name.padEnd(6)} exact=${String(exact)}/${String(cases.length)}  worst=${String(worst)}`,
    );
  }
}

if (fixturePath !== undefined) {
  const fixture = {
    $comment:
      'Ground truth for DrawingML colour transforms, measured in Microsoft PowerPoint. ' +
      'See docs/adr/0007-ground-truth.md. `expected` is the RGB PowerPoint painted, ' +
      'sampled from its own bitmap export at the centre of each swatch.',
    source: {
      application: 'Microsoft PowerPoint',
      generatedBy: 'tools/ground-truth/build-swatch-deck.ts + analyse-swatches.ts',
    },
    clrScheme: inputs.clrScheme,
    swatches: measured.map((m) => ({
      id: m.input.id,
      group: m.input.group,
      description: m.input.description,
      base: m.input.base,
      transforms: m.input.transforms,
      xml: m.input.xml,
      expected: m.bmpHex,
      objectModel: m.comHex,
    })),
  };
  writeFileSync(fixturePath, JSON.stringify(fixture, null, 2) + '\n');
  console.log(`\nwrote ${fixturePath}  (${String(measured.length)} swatches)`);
}
